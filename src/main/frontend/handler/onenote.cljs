(ns frontend.handler.onenote
  "Handler for OneNote graph sync:
   - Connect current local folder graph to a OneNote notebook
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
        (p/then (fn [_account]
                  (notification/show! (str "Signed in as " (:name (msal/get-account))) :success)))
        (p/catch (fn [e]
                   (notification/show! "OneNote sign-in failed" :error)
                   (log/error :onenote/login-failed {:error e}))))
    (do
      (notification/show! "MSAL not initialized." :warning)
      (p/rejected (ex-info "MSAL not initialized" {})))))

;; ---- Connect flow ----

(defn <connect-onenote-graph!
  "Connect the current local folder graph to a OneNote notebook.
   1. Login if needed
   2. Resolve notebook from pasted URL
   3. Find/create 'Logseq' section
   4. If page matching folder name exists → sync (pull+merge+push)
   5. If no matching page → push local graph as new page
   6. Save config for future syncs"
  [notebook-url]
  (let [repo (state/get-current-repo)
        repo-dir (when repo (config/get-repo-dir repo))
        is-local? (and repo (config/local-db? repo) (not (config/demo-graph? repo)))]
    (cond
      (not is-local?)
      (do (notification/show! "Open a local folder graph first, then connect to OneNote." :warning)
          (p/resolved nil))

      (not (seq notebook-url))
      (do (notification/show! "No notebook URL provided." :warning)
          (p/resolved nil))

      :else
      (-> (p/let [;; Step 1: Login
                  _ (when-not (msal/logged-in?) (<login!))
                  token (msal/get-token)
                  ;; Step 2: Resolve notebook from URL
                  _ (notification/show! "Finding notebook..." :info)
                  notebook (onenote/resolve-notebook-from-url token notebook-url)
                  _ (when-not notebook
                      (throw (ex-info "Could not find notebook from URL" {:url notebook-url})))
                  {:keys [name id site-id]} notebook
                  ;; Step 3: Find/create section
                  section (onenote/find-or-create-section token id config/ONENOTE-SECTION-NAME site-id)
                  section-id (:id section)
                  ;; Use the local folder name as the page title in OneNote
                  graph-name repo-dir]
            ;; Step 4: Initialize sync state
            (onenote-sync/start! {:section-id section-id
                                   :notebook-id id
                                   :site-id site-id
                                   :graph-name graph-name})
            ;; Step 5: Check if matching page exists
            (p/let [existing-page (onenote/find-page token section-id graph-name site-id)]
              (if existing-page
                ;; Page exists → full sync (pull+merge+push)
                (p/let [_ (notification/show! (str "Found existing graph '" graph-name "', syncing...") :info)
                        _ (state/set-state! :onenote/syncing? true)
                        _ (onenote-sync/sync!)
                        _ (state/set-state! :onenote/syncing? false)]
                  (notification/show! "Connected and synced with OneNote" :success))
                ;; No page → push local as new
                (p/let [_ (notification/show! (str "Pushing '" graph-name "' to OneNote...") :info)
                        _ (state/set-state! :onenote/syncing? true)
                        _ (onenote-sync/push!)
                        _ (state/set-state! :onenote/syncing? false)]
                  (notification/show! "Connected to OneNote and uploaded graph" :success))))
            (log/info :onenote/connected {:notebook name :graph graph-name}))
          (p/catch (fn [e]
                     (state/set-state! :onenote/syncing? false)
                     (notification/show! (str "Failed to connect OneNote: " (str e)) :error)
                     (log/error :onenote/connect-failed {:error e})))))))

;; ---- Sync flow ----

(defn <sync-onenote!
  "Manual sync: pull remote changes (with merge), then push local state."
  []
  ;; Ensure sync state is initialized from saved config
  (when (and (not (onenote-sync/initialized?))
             (onenote-sync/load-config))
    (onenote-sync/start! (onenote-sync/load-config)))
  (if-not (onenote-sync/initialized?)
    (do (notification/show! "Not connected to OneNote. Use 'Connect OneNote' first." :warning)
        (p/resolved nil))
    (-> (p/let [_ (notification/show! "Syncing with OneNote..." :info)
                _ (state/set-state! :onenote/syncing? true)
                _ (onenote-sync/sync!)
                _ (state/set-state! :onenote/syncing? false)]
          (notification/show! "OneNote sync complete" :success))
        (p/catch (fn [e]
                   (state/set-state! :onenote/syncing? false)
                   (notification/show! (str "OneNote sync failed: " (str e)) :error)
                   (log/error :onenote/sync-failed {:error e}))))))
