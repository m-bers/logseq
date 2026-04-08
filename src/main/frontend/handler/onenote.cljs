(ns frontend.handler.onenote
  "Handler for OneNote graph operations:
   - Login via MSAL
   - Pull graph ZIP from OneNote to LightningFS
   - Open as a file-based Logseq graph
   - Manual sync (push/pull with three-way merge)"
  (:require [clojure.string :as string]
            [frontend.auth.msal :as msal]
            [frontend.config :as config]
            [frontend.db :as db]
            [frontend.fs.onenote :as onenote]
            [frontend.handler.notification :as notification]
            [frontend.handler.web.nfs :as nfs-handler]
            [frontend.state :as state]
            [frontend.sync.onenote :as onenote-sync]
            [lambdaisland.glogi :as log]
            [logseq.common.path :as path]
            [promesa.core :as p]))

(defn <init-msal!
  "Initialize MSAL. Call once at app startup if client ID is configured."
  []
  (let [client-id config/MSAL-CLIENT-ID]
    (if (seq client-id)
      (-> (msal/init! client-id config/msal-redirect-uri)
          (p/catch (fn [e]
                     (log/error :onenote/msal-init-failed {:error e}))))
      (log/warn :onenote/no-client-id "Set MSAL-CLIENT-ID to enable OneNote sync"))))

(defn <login!
  "Interactive OneNote login."
  []
  (if (msal/initialized?)
    (-> (msal/login!)
        (p/then (fn [account]
                  (notification/show! (str "Signed in as " (:name (msal/get-account))) :success)
                  account))
        (p/catch (fn [e]
                   (notification/show! "OneNote sign-in failed" :error)
                   (log/error :onenote/login-failed {:error e}))))
    (do
      (notification/show! "MSAL not initialized." :warning)
      (p/rejected (ex-info "MSAL not initialized" {})))))

(defn <logout!
  "Logout from OneNote."
  []
  (onenote-sync/stop!)
  (-> (msal/logout!)
      (p/then (fn [_]
                (notification/show! "Signed out of OneNote" :success)))
      (p/catch (fn [e]
                 (log/error :onenote/logout-failed {:error e})))))

;; ---- Reading graph from LightningFS ----

(defn- <readdir-recursive
  "Recursively read all file paths from LightningFS under dir."
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
                          (<readdir-recursive full)))))
                  entries))]
      (vec (apply concat results)))))

(defn- <read-dir-as-graph
  "Read all files from LightningFS and return in the format
   that ls-dir-files-with-handler! expects from fs/open-dir."
  [local-dir]
  (let [root (path/url-to-path local-dir)
        root-name (subs root 1)]
    (p/let [all-paths (-> (<readdir-recursive root)
                          (p/catch (fn [_] [])))
            file-objs (p/all
                       (map (fn [fpath]
                              (p/let [content (-> (.readFile js/window.pfs fpath #js {:encoding "utf8"})
                                                  (p/then (fn [c] (.toString c))))
                                      stat (.stat js/window.pfs fpath)]
                                (let [rel-path (subs fpath (inc (count root)))
                                      fname (last (string/split fpath #"/"))]
                                  {:name    fname
                                   :path    rel-path
                                   :mtime   (.-mtimeMs stat)
                                   :size    (or (.-size stat) (count content))
                                   :type    "file"
                                   :content content})))
                            all-paths))]
      (log/info :onenote/read-dir {:root root-name :files (count file-objs)})
      {:path root-name
       :files (vec file-objs)})))

;; ---- Connect flow ----

(defn <connect-onenote-graph!
  "Full flow: login, resolve notebook from URL, pull ZIP, open as graph.
   notebook-url: a OneNote URL or OneIntraNote URL to identify the notebook."
  [notebook-url]
  (-> (p/let [;; Step 1: Ensure logged in
              _ (when-not (msal/logged-in?)
                  (<login!))
              token (msal/get-token)
              ;; Step 2: Resolve notebook from URL
              _ (notification/show! "Finding notebook..." :info)
              notebook (onenote/resolve-notebook-from-url token notebook-url)
              _ (when-not notebook
                  (throw (ex-info "Could not find notebook from URL" {:url notebook-url})))
              {:keys [name id site-id]} notebook
              ;; Step 3: Find or create "Logseq" section
              section (onenote/find-or-create-section token id config/ONENOTE-SECTION-NAME site-id)
              section-id (:id section)
              graph-name name
              local-dir (str "memory:///onenote-" (string/replace graph-name #"[^a-zA-Z0-9_-]" "_"))
              ;; Step 4: Close current graph
              _ (let [current-repo (state/get-current-repo)]
                  (when current-repo
                    (db/remove-conn! current-repo)
                    (state/set-current-repo! nil)))
              ;; Step 5: Ensure local directory exists
              local-root (path/url-to-path local-dir)
              _ (-> (js/window.pfs.mkdir local-root)
                    (p/catch (fn [_] nil)))
              ;; Step 6: Pull ZIP from OneNote
              _ (notification/show! (str "Syncing from OneNote/" graph-name "...") :info)
              _ (state/set-state! :onenote/syncing? true)
              file-count (onenote-sync/initial-pull! section-id graph-name site-id local-dir)
              _ (state/set-state! :onenote/syncing? false)
              _ (notification/show! (str "Pulled " file-count " files from OneNote") :success)
              ;; Step 6: Open as a file graph
              _ (nfs-handler/ls-dir-files-with-handler!
                 nil
                 {:dir-result-fn (fn [] (<read-dir-as-graph local-dir))})]
        ;; Step 7: Initialize sync state
        (onenote-sync/start! {:section-id section-id
                               :notebook-id id
                               :site-id site-id
                               :graph-name graph-name
                               :local-dir local-dir})
        (log/info :onenote/connected {:notebook name :files file-count}))
      (p/catch (fn [e]
                 (state/set-state! :onenote/syncing? false)
                 (notification/show! (str "Failed to connect OneNote: " (str e)) :error)
                 (log/error :onenote/connect-failed {:error e})))))

(defn <reconnect-onenote-graph!
  "Reconnect using saved config (no URL needed)."
  []
  (when-let [{:keys [section-id notebook-id site-id graph-name]} (onenote-sync/load-config)]
    (let [local-dir (str "memory:///onenote-" (string/replace graph-name #"[^a-zA-Z0-9_-]" "_"))]
      (-> (p/let [;; Ensure logged in
                  _ (when-not (msal/logged-in?)
                      (<login!))
                  ;; Close current graph
                  _ (let [current-repo (state/get-current-repo)]
                      (when current-repo
                        (db/remove-conn! current-repo)
                        (state/set-current-repo! nil)))
                  ;; Ensure local directory exists
                  local-root (path/url-to-path local-dir)
                  _ (-> (js/window.pfs.mkdir local-root)
                        (p/catch (fn [_] nil)))
                  ;; Pull
                  _ (notification/show! (str "Syncing from OneNote/" graph-name "...") :info)
                  _ (state/set-state! :onenote/syncing? true)
                  file-count (onenote-sync/initial-pull! section-id graph-name site-id local-dir)
                  _ (state/set-state! :onenote/syncing? false)
                  _ (notification/show! (str "Pulled " file-count " files from OneNote") :success)
                  ;; Open graph
                  _ (nfs-handler/ls-dir-files-with-handler!
                     nil
                     {:dir-result-fn (fn [] (<read-dir-as-graph local-dir))})]
            (onenote-sync/start! {:section-id section-id
                                   :notebook-id notebook-id
                                   :site-id site-id
                                   :graph-name graph-name
                                   :local-dir local-dir})
            (log/info :onenote/reconnected {:notebook graph-name :files file-count}))
          (p/catch (fn [e]
                     (state/set-state! :onenote/syncing? false)
                     (notification/show! (str "Failed to reconnect OneNote: " (str e)) :error)
                     (log/error :onenote/reconnect-failed {:error e})))))))

(defn <sync-onenote!
  "Manual sync: pull remote changes (with merge), then push local state.
   If sync state isn't initialized, loads saved config first."
  []
  ;; Ensure sync state is initialized from saved config
  (when (and (nil? (:section-id @onenote-sync/sync-state))
             (onenote-sync/load-config))
    (let [config (onenote-sync/load-config)
          local-dir (str "memory:///onenote-" (string/replace (:graph-name config) #"[^a-zA-Z0-9_-]" "_"))]
      (onenote-sync/start! (assoc config :local-dir local-dir))))
  (-> (p/let [_ (notification/show! "Syncing with OneNote..." :info)
              _ (state/set-state! :onenote/syncing? true)
              _ (onenote-sync/sync!)
              _ (state/set-state! :onenote/syncing? false)]
        (notification/show! "OneNote sync complete" :success))
      (p/catch (fn [e]
                 (state/set-state! :onenote/syncing? false)
                 (notification/show! (str "OneNote sync failed: " (str e)) :error)
                 (log/error :onenote/sync-failed {:error e})))))
