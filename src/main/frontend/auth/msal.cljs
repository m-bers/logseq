(ns frontend.auth.msal
  "ClojureScript wrapper for @azure/msal-browser.
   Provides OneDrive OAuth2 authentication via Microsoft Entra ID."
  (:require ["@azure/msal-browser" :as msal-browser]
            [promesa.core :as p]
            [frontend.state :as state]
            [lambdaisland.glogi :as log]))

(defonce ^:private msal-instance (atom nil))
(defonce ^:private current-account (atom nil))
(defonce ^:private auth-redirect-url (atom nil))

(def ^:private graph-scopes #js ["Files.ReadWrite" "User.Read"])

(def ^:private default-authority "https://login.microsoftonline.com/common")

(defn- derive-auth-redirect-uri
  "Derive the auth-redirect.html URL relative to the current page base."
  [base-uri]
  (str base-uri "auth-redirect.html"))

(defn- create-msal-config
  [client-id redirect-uri]
  #js {:auth #js {:clientId client-id
                   :authority default-authority
                   :redirectUri redirect-uri}
       :cache #js {:cacheLocation "localStorage"
                   :storeAuthStateInCookie false}})

(defn initialized?
  "Returns true if MSAL has been initialized."
  []
  (some? @msal-instance))

(defn init!
  "Initialize the MSAL PublicClientApplication.
   client-id: Azure AD app registration client ID
   redirect-uri: OAuth redirect base URI (e.g. https://joshcha.mbe.rs/logseq/)"
  [client-id redirect-uri]
  (let [popup-uri (derive-auth-redirect-uri redirect-uri)
        config (create-msal-config client-id popup-uri)
        pca (new (.-PublicClientApplication msal-browser) config)]
    (reset! auth-redirect-url popup-uri)
    (p/let [_ (.initialize pca)
            ;; Handle redirect response if returning from a redirect login
            response (.handleRedirectPromise pca)
            _ (when response
                (let [account (.-account response)]
                  (reset! current-account account)
                  (.setActiveAccount pca account)
                  (log/info :msal/redirect-login-success {:name (.-name account)})
                  ;; Clean up the hash fragment so Logseq's router isn't confused
                  (when (and js/window.location.hash
                             (.includes js/window.location.hash "code="))
                    (set! js/window.location.hash ""))))
            accounts (.getAllAccounts pca)]
      (when (and (nil? @current-account) (pos? (.-length accounts)))
        (reset! current-account (aget accounts 0))
        (.setActiveAccount pca (aget accounts 0)))
      (reset! msal-instance pca)
      (log/info :msal/initialized {:accounts (.-length accounts)
                                    :redirect-uri popup-uri})
      pca)))

(defn login!
  "Trigger interactive login via popup. Returns the account on success."
  []
  (when-let [pca @msal-instance]
    (p/let [response (.loginPopup pca #js {:scopes graph-scopes
                                            :redirectUri @auth-redirect-url})]
      (let [account (.-account response)]
        (reset! current-account account)
        (.setActiveAccount pca account)
        (state/pub-event! [:onedrive/logged-in {:name (.-name account)
                                                 :username (.-username account)}])
        (log/info :msal/login-success {:name (.-name account)})
        account))))

(defn logout!
  "Logout the current account."
  []
  (when-let [pca @msal-instance]
    (p/let [_ (.logoutPopup pca #js {:account @current-account
                                      :postLogoutRedirectUri @auth-redirect-url})]
      (reset! current-account nil)
      (state/pub-event! [:onedrive/logged-out])
      (log/info :msal/logout {}))))

(defn get-token
  "Acquire an access token silently, falling back to interactive if needed.
   Returns the access token string."
  []
  (when-let [pca @msal-instance]
    (let [request #js {:scopes graph-scopes
                       :account @current-account}]
      (-> (.acquireTokenSilent pca request)
          (p/then (fn [response] (.-accessToken response)))
          (p/catch (fn [_error]
                     (log/warn :msal/silent-token-failed "falling back to interactive")
                     (p/then (.acquireTokenPopup pca #js {:scopes graph-scopes
                                                          :redirectUri @auth-redirect-url})
                             (fn [response] (.-accessToken response)))))))))

(defn logged-in?
  "Returns true if there's a cached account."
  []
  (some? @current-account))

(defn get-account
  "Returns the current account map or nil."
  []
  (when-let [acct @current-account]
    {:name (.-name acct)
     :username (.-username acct)}))
