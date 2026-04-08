(ns frontend.sync.onenote
  "OneNote sync orchestrator.
   Syncs a local folder graph to/from a OneNote notebook via ZIP attachments.

   Architecture:
   - Local: real folder via File System Access API (NFS handles)
   - Remote: ZIP attachment on a OneNote page
   - Base: .logseq-onenote-base/ subfolder for three-way merge ancestry
   - Merge: per-file three-way merge using diff_merge.cljs"
  (:require [clojure.string :as string]
            [frontend.auth.msal :as msal]
            [frontend.config :as config]
            [frontend.fs :as fs]
            [frontend.fs.diff-merge :as diff-merge]
            [frontend.fs.onenote :as onenote]
            [frontend.fs.watcher-handler :as watcher-handler]
            [frontend.fs.zip-graph :as zip-graph]
            [frontend.state :as state]
            [lambdaisland.glogi :as log]
            [promesa.core :as p]))

;; ---- State ----

(defonce ^:private sync-state
  (atom {:status :idle        ;; :idle :syncing :error
         :section-id nil      ;; OneNote section ID
         :notebook-id nil     ;; OneNote notebook ID
         :site-id nil         ;; Graph API site ID (SharePoint)
         :graph-name nil      ;; Page title in OneNote (= folder name)
         :repo nil            ;; Logseq repo URL (e.g. "logseq_local_my-notes")
         :repo-dir nil        ;; Local folder name (e.g. "my-notes")
         :last-sync nil}))    ;; js/Date of last successful sync

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
;; The base snapshot tracks what was last synced so we can three-way merge.
;; It's stored as a simple JSON map in localStorage (not files).
;; Key: "logseq-onenote-base-{repo-dir}"

(defn- base-key [repo-dir]
  (str "logseq-onenote-base-" repo-dir))

(defn- save-base-snapshot!
  "Save the base snapshot (file map) to localStorage."
  [repo-dir file-map]
  ;; Store as JSON. For large graphs this could be big, but localStorage has 5-10MB.
  ;; For very large graphs, consider IndexedDB.
  (let [json (js/JSON.stringify (clj->js file-map))]
    (.setItem js/localStorage (base-key repo-dir) json)))

(defn- load-base-snapshot
  "Load the base snapshot from localStorage. Returns {path content} or empty map."
  [repo-dir]
  (if-let [json (.getItem js/localStorage (base-key repo-dir))]
    (js->clj (js/JSON.parse json))
    {}))

;; ---- Three-way merge ----

(defn- hash-content
  "Simple hash for content comparison."
  [content]
  (when content
    (str (count content) "-"
         (.toString
          (js/Uint32Array.
           #js [(reduce (fn [h c] (+ (bit-shift-left h 5) (- h) (.charCodeAt c 0)))
                        0 content)])
          "36"))))

(defn <merge-file-maps
  "Three-way merge of file maps.
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

;; ---- Sync operations ----

(defn pull!
  "Download ZIP from OneNote, three-way merge with local, write changes.
   Returns {:pulled count :conflicts [...]}."
  []
  (when (msal/logged-in?)
    (p/let [token (msal/get-token)
            {:keys [section-id site-id graph-name repo repo-dir]} @sync-state
            ;; Find page
            page (onenote/find-page token section-id graph-name site-id)]
      (if-not page
        (do (log/info :onenote-sync/no-remote-page {:graph-name graph-name})
            {:pulled 0 :conflicts []})
        (p/let [;; Download ZIP (3 API calls total)
                zip-url (onenote/get-page-zip-url token (:id page) site-id)
                _ (when-not zip-url
                    (throw (ex-info "No ZIP attachment on page" {:page-id (:id page)})))
                zip-data (onenote/download-zip token zip-url)
                ;; Unpack remote
                remote-files (zip-graph/<unpack-zip zip-data)
                ;; Read base and local
                base-files (load-base-snapshot repo-dir)
                local-files (zip-graph/<read-local-files repo-dir)
                ;; Three-way merge
                {:keys [merged conflicts]} (<merge-file-maps base-files local-files remote-files)
                ;; Write changed files to local folder
                _ (p/all
                   (keep (fn [[rel-path content]]
                           (let [old-content (get local-files rel-path)]
                             (when (not= content old-content)
                               (p/let [_ (zip-graph/<write-file-to-local repo repo-dir rel-path content)]
                                 ;; Notify watcher so datascript DB updates
                                 (watcher-handler/handle-changed!
                                  "change"
                                  {:dir repo-dir
                                   :path rel-path
                                   :content content
                                   :stat {:mtime (js/Date.now)}})))))
                         merged))
                ;; Save base snapshot
                _ (save-base-snapshot! repo-dir merged)]
          (log/info :onenote-sync/pull-complete {:files (count merged) :conflicts (count conflicts)})
          {:pulled (count merged) :conflicts conflicts})))))

(defn push!
  "Pack local graph to ZIP and upload to OneNote."
  []
  (when (msal/logged-in?)
    (p/let [token (msal/get-token)
            {:keys [section-id site-id graph-name repo-dir]} @sync-state
            ;; Pack local files to ZIP
            zip-data (zip-graph/<pack-graph repo-dir)
            ;; Upload (delete + wait + create = 2-3 API calls)
            _ (onenote/replace-page-zip token section-id graph-name zip-data site-id)
            ;; Save current local as base snapshot
            local-files (zip-graph/<read-local-files repo-dir)
            _ (save-base-snapshot! repo-dir local-files)]
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

;; ---- Lifecycle ----

(defn start!
  "Initialize sync state from config."
  [{:keys [section-id notebook-id site-id graph-name]}]
  (let [repo (state/get-current-repo)
        repo-dir (config/get-repo-dir repo)]
    (swap! sync-state assoc
           :section-id section-id
           :notebook-id notebook-id
           :site-id site-id
           :graph-name graph-name
           :repo repo
           :repo-dir repo-dir)
    ;; Save config for reconnect
    (save-config! {:section-id section-id
                    :notebook-id notebook-id
                    :site-id site-id
                    :graph-name graph-name})
    (log/info :onenote-sync/started {:graph-name graph-name :repo-dir repo-dir})))

(defn stop! []
  (swap! sync-state assoc :status :idle)
  (log/info :onenote-sync/stopped {}))

(defn initialized? []
  (some? (:section-id @sync-state)))
