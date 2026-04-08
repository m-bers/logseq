(ns frontend.sync.onenote
  "OneNote sync orchestrator.
   Manages bidirectional sync between local MemoryFs (LightningFS/IndexedDB)
   and a OneNote notebook via ZIP attachments on pages.

   Architecture:
   - Local state: LightningFS at /onenote-{name}/
   - Base snapshot: LightningFS at /onenote-{name}-base/ (last synced state for merge)
   - Remote state: ZIP attachment on a OneNote page
   - Merge: three-way merge per file using diff_merge.cljs"
  (:require [clojure.string :as string]
            [frontend.auth.msal :as msal]
            [frontend.fs.diff-merge :as diff-merge]
            [frontend.fs.memory-fs :as memory-fs]
            [frontend.fs.onenote :as onenote]
            [frontend.fs.watcher-handler :as watcher-handler]
            [frontend.fs.zip-graph :as zip-graph]
            [lambdaisland.glogi :as log]
            [logseq.common.path :as path]
            [promesa.core :as p]))

;; ---- State ----

(defonce ^:private sync-state
  (atom {:status :idle        ;; :idle :syncing :error
         :section-id nil      ;; OneNote section ID
         :notebook-id nil     ;; OneNote notebook ID
         :site-id nil         ;; Graph API site ID (SharePoint)
         :graph-name nil      ;; Page title in OneNote
         :local-dir nil       ;; memory:///onenote-{name}
         :last-sync nil       ;; js/Date of last successful sync
         :dirty? false}))     ;; true if local has unsaved changes

(defonce ^:private pulling? (atom false))

;; ---- Config persistence ----

(def ^:private config-key "logseq-onenote-sync-config")

(defn save-config!
  "Save sync config to localStorage for reconnect."
  [config]
  (.setItem js/localStorage config-key (js/JSON.stringify (clj->js config))))

(defn load-config
  "Load saved sync config from localStorage."
  []
  (when-let [json (.getItem js/localStorage config-key)]
    (js->clj (js/JSON.parse json) :keywordize-keys true)))

;; ---- Base snapshot management ----

(defn- base-dir
  "Returns the absolute LightningFS path for the base snapshot."
  [local-dir]
  (str (path/url-to-path local-dir) "-base"))

(defn- <save-base-snapshot!
  "Write all files from file-map to the base snapshot directory."
  [local-dir file-map]
  (let [bdir (base-dir local-dir)]
    ;; Clear existing base
    (-> (js/window.workerThread.rimraf bdir)
        (p/catch (fn [_] nil)))
    (p/let [_ (js/window.pfs.mkdir bdir)]
      (p/all
       (map (fn [[rel-path content]]
              (let [full-path (path/path-join bdir rel-path)
                    parent (path/parent full-path)]
                (p/let [_ (zip-graph/<mkdir-recursive! parent)]
                  (.writeFile js/window.pfs full-path content))))
            file-map)))))

(defn- <read-base-snapshot
  "Read all files from the base snapshot. Returns {path -> content} or empty map."
  [local-dir]
  (let [bdir (base-dir local-dir)]
    (-> (zip-graph/<read-local-files (str "memory://" bdir))
        (p/catch (fn [_] {})))))

;; ---- Three-way merge ----

(defn- hash-content
  "Simple hash for content comparison. Returns nil for nil content."
  [content]
  (when content
    (str (count content) "-" (.toString (js/Uint32Array. #js [(reduce (fn [h c] (+ (bit-shift-left h 5) (- h) (.charCodeAt c 0))) 0 content)]) "36"))))

(defn <merge-file-maps
  "Three-way merge of file maps.
   base-files: {path content} from last sync
   local-files: {path content} from LightningFS
   remote-files: {path content} from OneNote ZIP
   Returns {:merged {path content} :conflicts [{:path :reason}]}."
  [base-files local-files remote-files]
  (let [all-paths (set (concat (keys base-files) (keys local-files) (keys remote-files)))
        results (atom {:merged {} :conflicts []})]
    (doseq [fpath all-paths]
      (let [base-content (get base-files fpath)
            local-content (get local-files fpath)
            remote-content (get remote-files fpath)
            base-hash (hash-content base-content)
            local-hash (hash-content local-content)
            remote-hash (hash-content remote-content)
            local-changed? (not= local-hash base-hash)
            remote-changed? (not= remote-hash base-hash)]
        (cond
          ;; No changes
          (and (not local-changed?) (not remote-changed?))
          (when local-content
            (swap! results assoc-in [:merged fpath] local-content))

          ;; Only remote changed
          (and (not local-changed?) remote-changed?)
          (when remote-content
            (swap! results assoc-in [:merged fpath] remote-content))

          ;; Only local changed
          (and local-changed? (not remote-changed?))
          (when local-content
            (swap! results assoc-in [:merged fpath] local-content))

          ;; Both changed to same content
          (= local-hash remote-hash)
          (when local-content
            (swap! results assoc-in [:merged fpath] local-content))

          ;; Both changed differently — merge .md files, prefer local for others
          :else
          (if (string/ends-with? fpath ".md")
            (let [format (if (string/ends-with? fpath ".org") :org :markdown)
                  merged (diff-merge/three-way-merge
                          (or base-content "")
                          (or remote-content "")
                          (or local-content "")
                          format)]
              (swap! results assoc-in [:merged fpath] merged))
            (do
              (when local-content
                (swap! results assoc-in [:merged fpath] local-content))
              (swap! results update :conflicts conj
                     {:path fpath :reason "both-changed-non-mergeable"}))))))
    @results))

;; ---- Watcher notification ----

(defn- notify-file-change!
  "Notify Logseq's watcher handler about a file change so the datascript DB is updated."
  [local-dir local-path content change-type]
  (let [local-root (path/url-to-path local-dir)
        rel-path (subs local-path (inc (count local-root)))]
    (watcher-handler/handle-changed!
     change-type
     {:dir local-dir
      :path rel-path
      :content content
      :stat {:mtime (js/Date.now)}})))

;; ---- Sync operations ----

(defn pull!
  "Download ZIP from OneNote, three-way merge with local, update LightningFS + datascript.
   Returns {:pulled count :conflicts [...]}."
  []
  (when (msal/logged-in?)
    (reset! pulling? true)
    (-> (p/let [token (msal/get-token)
                {:keys [section-id site-id graph-name local-dir]} @sync-state
                ;; Step 1-3: Find page and download ZIP (3 API calls)
                page (onenote/find-page token section-id graph-name site-id)]
          (if-not page
            (do (log/info :onenote-sync/no-remote-page {:graph-name graph-name})
                {:pulled 0 :conflicts []})
            (p/let [zip-url (onenote/get-page-zip-url token (:id page) site-id)
                    _ (when-not zip-url
                        (throw (ex-info "No ZIP attachment on page" {:page-id (:id page)})))
                    zip-data (onenote/download-zip token zip-url)
                    ;; Step 4: Unpack remote
                    remote-files (zip-graph/<unpack-zip zip-data)
                    ;; Step 5-6: Read base and local
                    base-files (<read-base-snapshot local-dir)
                    local-files (zip-graph/<read-local-files local-dir)
                    ;; Step 7: Three-way merge
                    {:keys [merged conflicts]} (<merge-file-maps base-files local-files remote-files)
                    ;; Step 8: Write merged files to local FS
                    local-root (path/url-to-path local-dir)
                    _ (p/all
                       (map (fn [[rel-path content]]
                              (let [old-content (get local-files rel-path)
                                    full-path (path/path-join local-root rel-path)]
                                (when (not= content old-content)
                                  (let [parent (path/parent full-path)]
                                    (p/let [_ (zip-graph/<mkdir-recursive! parent)
                                            _ (.writeFile js/window.pfs full-path content)]
                                      ;; Step 9: Notify watcher
                                      (notify-file-change! local-dir full-path content "change"))))))
                            merged))
                    ;; Step 10: Save base snapshot
                    _ (<save-base-snapshot! local-dir merged)]
              (log/info :onenote-sync/pull-complete {:files (count merged) :conflicts (count conflicts)})
              {:pulled (count merged) :conflicts conflicts})))
        (p/finally (fn [_] (reset! pulling? false))))))

(defn push!
  "Pack local graph to ZIP and upload to OneNote.
   Returns Promise."
  []
  (when (msal/logged-in?)
    (p/let [token (msal/get-token)
            {:keys [section-id site-id graph-name local-dir]} @sync-state
            ;; Pack local files to ZIP
            zip-data (zip-graph/<pack-graph local-dir)
            ;; Upload (delete + wait + create = 2-3 API calls)
            _ (onenote/replace-page-zip token section-id graph-name zip-data site-id)
            ;; Save current local as base snapshot
            local-files (zip-graph/<read-local-files local-dir)
            _ (<save-base-snapshot! local-dir local-files)]
      (swap! sync-state assoc :dirty? false)
      (log/info :onenote-sync/push-complete {:files (count local-files)}))))

(defn sync!
  "Full sync: pull remote changes (with merge), then push local state."
  []
  (when (and (msal/logged-in?) js/navigator.onLine)
    (swap! sync-state assoc :status :syncing)
    (-> (p/let [{:keys [conflicts]} (pull!)
                _ (push!)]
          (swap! sync-state assoc
                 :status :idle
                 :last-sync (js/Date.))
          (when (seq conflicts)
            (log/warn :onenote-sync/conflicts {:conflicts conflicts}))
          (log/info :onenote-sync/complete {}))
        (p/catch (fn [error]
                   (swap! sync-state assoc :status :error)
                   (log/error :onenote-sync/error {:error error})
                   (throw error))))))

;; ---- Initial sync (first connection) ----

(defn initial-pull!
  "First-time sync: download ZIP from OneNote, unpack to local FS + base snapshot.
   No merge needed (no local state yet).
   Returns file count."
  [section-id graph-name site-id local-dir]
  (p/let [token (msal/get-token)
          page (onenote/find-page token section-id graph-name site-id)]
    (if-not page
      ;; No remote page yet — empty graph, just create base dir
      (do (log/info :onenote-sync/no-remote {:graph-name graph-name})
          0)
      (p/let [zip-url (onenote/get-page-zip-url token (:id page) site-id)
              _ (when-not zip-url
                  (throw (ex-info "No ZIP attachment on page" {:page-id (:id page)})))
              zip-data (onenote/download-zip token zip-url)
              ;; Unpack to local FS
              local-root (path/url-to-path local-dir)
              file-map (zip-graph/<unpack-zip-to-fs zip-data local-root)
              ;; Save as base snapshot
              _ (<save-base-snapshot! local-dir file-map)]
        (log/info :onenote-sync/initial-pull {:files (count file-map)})
        (count file-map)))))

;; ---- Lifecycle ----

(defn start!
  "Initialize sync state and write hook (manual sync only).
   config: {:section-id :notebook-id :site-id :graph-name :local-dir}"
  [{:keys [section-id notebook-id site-id graph-name local-dir]}]
  (swap! sync-state assoc
         :section-id section-id
         :notebook-id notebook-id
         :site-id site-id
         :graph-name graph-name
         :local-dir local-dir)

  ;; Register write hook to track dirty state
  (let [local-root (path/url-to-path local-dir)]
    (reset! memory-fs/on-write-hook
            (fn [fpath]
              (when (and (not @pulling?)
                         (string/starts-with? fpath local-root))
                (swap! sync-state assoc :dirty? true)))))

  ;; Save config for reconnect
  (save-config! {:section-id section-id
                  :notebook-id notebook-id
                  :site-id site-id
                  :graph-name graph-name})

  (log/info :onenote-sync/started {:graph-name graph-name}))

(defn stop!
  "Stop the sync system."
  []
  (reset! memory-fs/on-write-hook nil)
  (swap! sync-state assoc :status :idle)
  (log/info :onenote-sync/stopped {}))

(defn initialized?
  "Returns true if sync state has been configured."
  []
  (some? (:section-id @sync-state)))

(defn get-sync-status
  "Returns current sync status for UI display."
  []
  (select-keys @sync-state [:status :last-sync :dirty?]))
