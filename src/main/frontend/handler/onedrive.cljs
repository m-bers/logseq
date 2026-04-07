(ns frontend.handler.onedrive
  "Handler for OneDrive graph operations:
   - Login via MSAL
   - Pull files from OneDrive to LightningFS
   - Open as a file-based Logseq graph"
  (:require [frontend.auth.msal :as msal]
            [frontend.config :as config]
            [frontend.fs.onedrive :as graph-api]
            [frontend.handler.notification :as notification]
            [frontend.handler.repo :as repo-handler]
            [frontend.handler.web.nfs :as nfs-handler]
            [frontend.state :as state]
            [frontend.sync.onedrive :as onedrive-sync]
            [lambdaisland.glogi :as log]
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

(defn <connect-onedrive-graph!
  "Full flow: login (if needed), pull files from OneDrive, open as graph.
   onedrive-folder: e.g. 'Notes' (the folder in OneDrive root)"
  [& {:keys [onedrive-folder]
      :or {onedrive-folder default-onedrive-folder}}]
  (let [local-dir (str "memory:///onedrive-" onedrive-folder)
        graph-name (str "OneDrive: " onedrive-folder)]
    (-> (p/let [;; Step 1: Ensure logged in
                _ (when-not (msal/logged-in?)
                    (<login!))
                ;; Step 2: Pull all files
                _ (notification/show! (str "Syncing from OneDrive/" onedrive-folder "...") :info)
                _ (state/set-state! :onedrive/syncing? true)
                files (onedrive-sync/initial-pull! onedrive-folder local-dir)
                _ (state/set-state! :onedrive/syncing? false)
                _ (notification/show! (str "Pulled " (count files) " files from OneDrive") :success)
                ;; Step 3: Open the memory:// directory as a file graph
                _ (nfs-handler/ls-dir-files-with-handler!
                   nil
                   {:dir local-dir})]
          ;; Step 4: Start background sync
          (onedrive-sync/start! onedrive-folder local-dir)
          (log/info :onedrive/connected {:folder onedrive-folder :files (count files)}))
        (p/catch (fn [e]
                   (state/set-state! :onedrive/syncing? false)
                   (notification/show! (str "Failed to connect OneDrive: " (str e)) :error)
                   (log/error :onedrive/connect-failed {:error e}))))))
