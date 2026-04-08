(ns frontend.handler.onenote
  "Handler for OneNote graph operations:
   - Login via MSAL
   - Connect current graph to a OneNote notebook
   - Manual sync (push/pull with three-way merge)"
  (:require [frontend.auth.msal :as msal]
            [frontend.config :as config]
            [frontend.fs.onenote :as onenote]
            [frontend.handler.notification :as notification]
            [frontend.state :as state]
            [frontend.sync.onenote :as onenote-sync]
            [lambdaisland.glogi :as log]
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

;; ---- Connect flow ----

(defn <connect-onenote-graph!
  "Connect the current graph to a OneNote notebook.
   Resolves notebook from URL, sets up sync config, and pushes current graph.
   notebook-url: a OneNote URL or OneIntraNote URL to identify the notebook."
  [notebook-url]
  (let [current-repo (state/get-current-repo)
        local-dir (when current-repo
                    (config/get-repo-dir current-repo))]
    (when-not local-dir
      (notification/show! "No graph loaded to sync" :warning))
    (when local-dir
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
                  graph-name name]
            ;; Step 4: Initialize sync state pointing to current graph
            (onenote-sync/start! {:section-id section-id
                                   :notebook-id id
                                   :site-id site-id
                                   :graph-name graph-name
                                   :local-dir local-dir})
            ;; Step 5: Push current graph to OneNote
            (p/let [_ (notification/show! (str "Pushing to OneNote/" graph-name "...") :info)
                    _ (state/set-state! :onenote/syncing? true)
                    _ (onenote-sync/push!)
                    _ (state/set-state! :onenote/syncing? false)]
              (notification/show! "Connected to OneNote and synced" :success)
              (log/info :onenote/connected {:notebook name})))
          (p/catch (fn [e]
                     (state/set-state! :onenote/syncing? false)
                     (notification/show! (str "Failed to connect OneNote: " (str e)) :error)
                     (log/error :onenote/connect-failed {:error e})))))))

(defn <sync-onenote!
  "Manual sync: pull remote changes (with merge), then push local state.
   If sync state isn't initialized, loads saved config first."
  []
  ;; Ensure sync state is initialized from saved config
  (when-let [config (and (not (onenote-sync/initialized?))
                         (onenote-sync/load-config))]
    (let [current-repo (state/get-current-repo)
          local-dir (when current-repo (config/get-repo-dir current-repo))]
      (when local-dir
        (onenote-sync/start! (assoc config :local-dir local-dir)))))
  (-> (p/let [_ (notification/show! "Syncing with OneNote..." :info)
              _ (state/set-state! :onenote/syncing? true)
              _ (onenote-sync/sync!)
              _ (state/set-state! :onenote/syncing? false)]
        (notification/show! "OneNote sync complete" :success))
      (p/catch (fn [e]
                 (state/set-state! :onenote/syncing? false)
                 (notification/show! (str "OneNote sync failed: " (str e)) :error)
                 (log/error :onenote/sync-failed {:error e})))))
