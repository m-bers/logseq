(ns frontend.sync.onedrive
  "OneDrive sync orchestrator.
   Manages bidirectional sync between local MemoryFs (LightningFS/IndexedDB)
   and OneDrive via Microsoft Graph API.

   Architecture:
   - All reads/writes go to local MemoryFs (unchanged Logseq behavior)
   - This layer runs in the background, pushing/pulling changes
   - Uses Graph API delta endpoint for efficient incremental sync
   - Queues writes when offline, flushes when back online"
  (:require [clojure.string :as string]
            [frontend.auth.msal :as msal]
            [frontend.fs.memory-fs :as memory-fs]
            [frontend.fs.onedrive :as graph]
            [frontend.fs.watcher-handler :as watcher-handler]
            [lambdaisland.glogi :as log]
            [logseq.common.path :as path]
            [promesa.core :as p]))

;; ---- State ----

(defonce ^:private sync-state
  (atom {:status :idle           ;; :idle :syncing :error :offline
         :delta-link nil         ;; Graph API delta link for incremental sync
         :dirty-files #{}        ;; Set of local paths that need pushing
         :last-sync nil          ;; js/Date of last successful sync
         :sync-interval-id nil   ;; setInterval ID
         :onedrive-folder nil    ;; e.g. "Notes"
         :local-dir nil}))       ;; e.g. "memory:///onedrive-notes"

(def ^:private sync-interval-ms 30000) ;; 30 seconds
(defonce ^:private pulling? (atom false)) ;; true during pull to suppress write-hook

(def ^:private delta-link-key "logseq-onedrive-delta-link")

;; ---- IndexedDB persistence for delta link ----

(defn- save-delta-link! [link]
  (when link
    (js/localStorage.setItem delta-link-key link)))

(defn- load-delta-link []
  (js/localStorage.getItem delta-link-key))

;; ---- Local FS helpers (via LightningFS / window.pfs) ----

(defn- local-read [local-path]
  (-> (.readFile js/window.pfs local-path #js {:encoding "utf8"})
      (p/then (fn [content] (.toString content)))))

(defn- <mkdir-recursive!
  "Create directory and all parents, ignoring already-exists errors."
  [dir]
  (when (and dir (not= dir "/") (not= dir ""))
    (-> (.stat js/window.pfs dir)
        (p/then (fn [_] nil)) ;; already exists
        (p/catch (fn [_]
                   ;; doesn't exist, create parent first
                   (p/let [_ (<mkdir-recursive! (path/parent dir))]
                     (-> (.mkdir js/window.pfs dir)
                         (p/catch (fn [_] nil)))))))))

(defn- local-write [local-path content]
  (p/let [parent (path/parent local-path)
          _ (<mkdir-recursive! parent)]
    (.writeFile js/window.pfs local-path content)))

(defn- local-unlink [local-path]
  (-> (.unlink js/window.pfs local-path)
      (p/catch (fn [_] nil))))

(defn- local-stat [local-path]
  (-> (.stat js/window.pfs local-path)
      (p/catch (fn [_] nil))))

(defn- local-readdir-recursive
  "Read all files under dir recursively. Returns seq of relative paths."
  [dir]
  (p/let [entries (-> (.readdir js/window.pfs dir)
                      (p/then (fn [r] (js->clj r))))]
    (p/let [results
            (p/all
             (map (fn [entry]
                    (let [full (path/path-join dir entry)]
                      (p/let [stat (.stat js/window.pfs full)]
                        (if (= (.-type stat) "file")
                          (p/resolved [full])
                          (local-readdir-recursive full)))))
                  entries))]
      (apply concat results))))

;; ---- Path mapping ----

(defn- remote->local
  "Convert OneDrive path (e.g. 'Notes/pages/foo.md') to local path."
  [remote-path]
  (let [{:keys [local-dir onedrive-folder]} @sync-state
        relative (subs remote-path (inc (count onedrive-folder)))]
    (path/path-join (path/url-to-path local-dir) relative)))

(defn- local->remote
  "Convert local path to OneDrive path."
  [local-path]
  (let [{:keys [local-dir onedrive-folder]} @sync-state
        local-root (path/url-to-path local-dir)
        relative (subs local-path (inc (count local-root)))]
    (str onedrive-folder "/" relative)))

;; ---- Pull (remote → local) ----

(defn- notify-file-change!
  "Notify Logseq's watcher handler about a file change so the datascript DB is updated."
  [local-path content change-type]
  (let [{:keys [local-dir]} @sync-state
        ;; local-path is absolute like /onedrive-Notes/pages/foo.md
        ;; relative path should be like pages/foo.md
        local-root (path/url-to-path local-dir)
        rel-path (subs local-path (inc (count local-root)))]
    ;; handle-changed! compares dir to (config/get-local-dir repo) which is just
    ;; the dir name without memory:// prefix. Pass local-dir so get-fs routes
    ;; correctly, but the watcher handler's (= dir repo-dir) check uses
    ;; get-local-dir which returns the name without prefix.
    ;; Use "add" type which doesn't check (= dir repo-dir).
    (watcher-handler/handle-changed!
     (if (= change-type "change") "add" change-type)
     {:dir local-dir
      :path rel-path
      :content content
      :stat {:mtime (js/Date.now)}})))

(defn- skip-remote-path?
  "Skip backup, recycle, and non-content files from delta sync."
  [remote-path]
  (or (string/includes? remote-path "/logseq/bak/")
      (string/includes? remote-path "/logseq/.recycle/")
      (string/includes? remote-path "/.git/")))

(defn- apply-remote-change
  "Apply a single remote change to local fs and update datascript DB."
  [token change]
  (let [remote-path (get-in change [:parentReference :path])
        name (:name change)
        deleted? (some? (:deleted change))
        is-file? (nil? (:folder change))]
    (when (and name is-file?)
      (let [;; parentReference.path looks like /drive/root:/Notes/pages
            parent-suffix (second (re-find #"/drive/root:/(.*)" (or remote-path "")))
            full-remote (when parent-suffix (str parent-suffix "/" name))
            local-path (when full-remote (remote->local full-remote))]
        (when (and local-path (not (skip-remote-path? (or full-remote ""))))
          (if deleted?
            (p/do!
             (log/info :onedrive-sync/delete-local {:path local-path})
             (local-unlink local-path)
             (notify-file-change! local-path nil "unlink"))
            (p/let [content (graph/read-file token full-remote)]
              (log/info :onedrive-sync/pull-file {:remote full-remote :local local-path})
              (local-write local-path content)
              (notify-file-change! local-path content "change"))))))))

(defn pull!
  "Pull remote changes from OneDrive to local and update datascript DB."
  []
  (when (msal/logged-in?)
    (reset! pulling? true)
    (-> (p/let [token (msal/get-token)
                {:keys [onedrive-folder]} @sync-state
                delta-link (or (:delta-link @sync-state) (load-delta-link))
                {:keys [changes delta-link]} (graph/get-delta token onedrive-folder delta-link)]
          (log/info :onedrive-sync/pull {:changes (count changes)})
          (p/let [_ (p/all (map (partial apply-remote-change token) changes))]
            (swap! sync-state assoc :delta-link delta-link)
            (save-delta-link! delta-link)
            (count changes)))
        (p/finally (fn [_] (reset! pulling? false))))))

;; ---- Push (local → remote) ----

(defn mark-dirty!
  "Mark a local file as needing to be pushed to OneDrive."
  [local-path]
  (swap! sync-state update :dirty-files conj local-path))

(defn- push-file!
  "Push a single local file to OneDrive."
  [token local-path]
  (p/let [stat (local-stat local-path)]
    (if stat
      (p/let [content (local-read local-path)
              remote-path (local->remote local-path)]
        (log/info :onedrive-sync/push-file {:local local-path :remote remote-path})
        (graph/write-file token remote-path content))
      ;; File was deleted locally — delete remotely too
      (p/let [remote-path (local->remote local-path)]
        (log/info :onedrive-sync/push-delete {:remote remote-path})
        (-> (graph/delete-file token remote-path)
            (p/catch (fn [e]
                       (log/warn :onedrive-sync/delete-remote-failed {:error e}))))))))

(defn push!
  "Push all dirty local files to OneDrive."
  []
  (when (and (msal/logged-in?) (seq (:dirty-files @sync-state)))
    (p/let [token (msal/get-token)
            dirty (vec (:dirty-files @sync-state))
            _ (swap! sync-state assoc :dirty-files #{})]
      (log/info :onedrive-sync/push {:files (count dirty)})
      (p/all (map (partial push-file! token) dirty)))))

;; ---- Full sync cycle ----

(defn sync!
  "Run a full sync cycle: push dirty files, then pull remote changes."
  []
  (when (and (msal/logged-in?) js/navigator.onLine)
    (swap! sync-state assoc :status :syncing)
    (log/info :onedrive-sync/start {})
    (-> (p/let [_ (push!)
                n-pulled (pull!)]
          (swap! sync-state assoc
                 :status :idle
                 :last-sync (js/Date.))
          (log/info :onedrive-sync/complete {:pulled n-pulled}))
        (p/catch (fn [error]
                   (swap! sync-state assoc :status :error)
                   (log/error :onedrive-sync/error {:error error}))))))

;; ---- Initial sync (first connection) ----

(defn initial-pull!
  "Pull all files from OneDrive to local for the first time.
   Returns the list of file maps."
  [onedrive-folder local-dir]
  (p/let [token (msal/get-token)
          files (graph/list-files-recursive token onedrive-folder)]
    (log/info :onedrive-sync/initial-pull {:files (count files)})
    (p/let [_ (p/all
               (map (fn [{:keys [path name]}]
                      (when-not (:folder? path)
                        (p/let [content (graph/read-file token path)
                                local-path (str (path/url-to-path local-dir)
                                                "/"
                                                (subs path (inc (count onedrive-folder))))]
                          (local-write local-path content))))
                    files))]
      files)))

;; ---- Lifecycle ----

(defn start!
  "Start the OneDrive sync system.
   onedrive-folder: remote folder name, e.g. 'Notes'
   local-dir: local memory:// path, e.g. 'memory:///onedrive-notes'"
  [onedrive-folder local-dir]
  (swap! sync-state assoc
         :onedrive-folder onedrive-folder
         :local-dir local-dir
         :delta-link (load-delta-link))

  ;; Listen for online/offline
  (js/window.addEventListener "online"
    (fn [_] (log/info :onedrive-sync/online {}) (sync!)))
  (js/window.addEventListener "offline"
    (fn [_]
      (log/info :onedrive-sync/offline {})
      (swap! sync-state assoc :status :offline)))

  ;; Register write hook to track dirty files for push
  ;; Skip during pulls to avoid push-back loops
  (let [local-root (path/url-to-path local-dir)]
    (reset! memory-fs/on-write-hook
            (fn [fpath]
              (when (and (not @pulling?)
                         (string/starts-with? fpath local-root))
                (mark-dirty! fpath)))))

  ;; Periodic sync
  (let [interval-id (js/setInterval sync! sync-interval-ms)]
    (swap! sync-state assoc :sync-interval-id interval-id))

  ;; Initial sync
  (sync!)

  (log/info :onedrive-sync/started {:folder onedrive-folder :local local-dir}))

(defn stop!
  "Stop the sync system."
  []
  (when-let [id (:sync-interval-id @sync-state)]
    (js/clearInterval id))
  (reset! memory-fs/on-write-hook nil)
  (swap! sync-state assoc
         :sync-interval-id nil
         :status :idle)
  (log/info :onedrive-sync/stopped {}))

(defn get-sync-status
  "Returns current sync status for UI display."
  []
  (select-keys @sync-state [:status :last-sync :dirty-files]))
