(ns frontend.handler.onedrive
  "Handler for OneDrive graph operations:
   - Login via MSAL
   - Pull files from OneDrive to LightningFS
   - Open as a file-based Logseq graph"
  (:require [clojure.string :as string]
            [frontend.auth.msal :as msal]
            [frontend.config :as config]
            [frontend.db :as db]
            [frontend.handler.notification :as notification]
            [frontend.handler.web.nfs :as nfs-handler]
            [frontend.state :as state]
            [frontend.sync.onedrive :as onedrive-sync]
            [lambdaisland.glogi :as log]
            [logseq.common.path :as path]
            [promesa.core :as p]))

(def ^:private default-onedrive-folder "Notes")

(defn <init-msal!
  "Initialize MSAL. Call once at app startup if client ID is configured."
  []
  (let [client-id config/MSAL-CLIENT-ID]
    (if (seq client-id)
      (-> (msal/init! client-id config/msal-redirect-uri)
          (p/catch (fn [e]
                     (log/error :onedrive/msal-init-failed {:error e}))))
      (log/warn :onedrive/no-client-id "Set MSAL-CLIENT-ID to enable OneDrive sync"))))

(defn <login!
  "Interactive OneDrive login."
  []
  (if (msal/initialized?)
    (-> (msal/login!)
        (p/then (fn [account]
                  (notification/show! (str "Signed in as " (:name (msal/get-account))) :success)
                  account))
        (p/catch (fn [e]
                   (notification/show! "OneDrive sign-in failed" :error)
                   (log/error :onedrive/login-failed {:error e}))))
    (do
      (notification/show! "MSAL not initialized. Set your Azure AD client ID." :warning)
      (p/rejected (ex-info "MSAL not initialized" {})))))

(defn <logout!
  "Logout from OneDrive."
  []
  (onedrive-sync/stop!)
  (-> (msal/logout!)
      (p/then (fn [_]
                (notification/show! "Signed out of OneDrive" :success)))
      (p/catch (fn [e]
                 (log/error :onedrive/logout-failed {:error e})))))

(defn- <readdir-recursive
  "Recursively read all file paths from LightningFS under dir.
   Returns a vector of absolute paths (strings)."
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

(defn- <read-onedrive-dir-as-graph
  "Read all files from LightningFS and return in the format
   that ls-dir-files-with-handler! expects from fs/open-dir:
   {:path \"onedrive-Notes\" :files [{:name :path :mtime :size :type :content} ...]}
   File paths must be RELATIVE to the root (e.g. \"pages/foo.md\" not \"onedrive-Notes/pages/foo.md\")
   to match how fs/open-dir normalizes paths."
  [local-dir]
  (let [root (path/url-to-path local-dir)   ;; "/onedrive-Notes"
        root-name (subs root 1)]            ;; "onedrive-Notes"
    (p/let [all-paths (<readdir-recursive root)
            file-objs (p/all
                       (map (fn [fpath]
                              (p/let [content (-> (.readFile js/window.pfs fpath #js {:encoding "utf8"})
                                                  (p/then (fn [c] (.toString c))))
                                      stat (.stat js/window.pfs fpath)]
                                (let [rel-path (subs fpath (inc (count root)))  ;; "pages/foo.md"
                                      fname (last (string/split fpath #"/"))]   ;; "foo.md"
                                  {:name    fname
                                   :path    rel-path  ;; "pages/foo.md" (relative to root)
                                   :mtime   (.-mtimeMs stat)
                                   :size    (or (.-size stat) (count content))
                                   :type    "file"
                                   :content content})))
                            all-paths))]
      (log/info :onedrive/read-dir {:root root-name :files (count file-objs)})
      {:path root-name
       :files (vec file-objs)})))

(defn <connect-onedrive-graph!
  "Full flow: login (if needed), pull files from OneDrive, open as graph."
  [& {:keys [onedrive-folder]
      :or {onedrive-folder default-onedrive-folder}}]
  (let [local-dir (str "memory:///onedrive-" onedrive-folder)]
    (-> (p/let [;; Step 1: Ensure logged in
                _ (when-not (msal/logged-in?)
                    (<login!))
                ;; Step 2: Close the current graph to avoid collisions
                _ (let [current-repo (state/get-current-repo)]
                    (when current-repo
                      (db/remove-conn! current-repo)
                      (state/set-current-repo! nil)))
                ;; Step 3: Pull all files from OneDrive to LightningFS
                _ (notification/show! (str "Syncing from OneDrive/" onedrive-folder "...") :info)
                _ (state/set-state! :onedrive/syncing? true)
                files (onedrive-sync/initial-pull! onedrive-folder local-dir)
                _ (state/set-state! :onedrive/syncing? false)
                _ (notification/show! (str "Pulled " (count files) " files from OneDrive") :success)
                ;; Step 4: Open as a file graph
                ;; Provide dir-result-fn that reads from LightningFS in the format
                ;; that matches what fs/open-dir returns for NFS graphs
                _ (nfs-handler/ls-dir-files-with-handler!
                   nil
                   {:dir-result-fn (fn [] (<read-onedrive-dir-as-graph local-dir))})]
          ;; Step 5: Initialize sync state (manual sync only, no auto-sync)
          (onedrive-sync/start! onedrive-folder local-dir)
          (log/info :onedrive/connected {:folder onedrive-folder :files (count files)}))
        (p/catch (fn [e]
                   (state/set-state! :onedrive/syncing? false)
                   (notification/show! (str "Failed to connect OneDrive: " (str e)) :error)
                   (log/error :onedrive/connect-failed {:error e}))))))

(defn <sync-onedrive!
  "Manual sync: push dirty files then pull remote changes."
  []
  (-> (p/let [_ (notification/show! "Syncing with OneDrive..." :info)
              _ (state/set-state! :onedrive/syncing? true)
              _ (onedrive-sync/sync!)
              _ (state/set-state! :onedrive/syncing? false)]
        (notification/show! "OneDrive sync complete" :success))
      (p/catch (fn [e]
                 (state/set-state! :onedrive/syncing? false)
                 (notification/show! (str "OneDrive sync failed: " (str e)) :error)
                 (log/error :onedrive/sync-failed {:error e})))))
