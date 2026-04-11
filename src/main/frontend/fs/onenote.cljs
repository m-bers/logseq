(ns frontend.fs.onenote
  "Microsoft Graph API client for OneNote operations.
   Stores/retrieves Logseq graphs as ZIP attachments on OneNote pages."
  (:require [clojure.string :as string]
            [lambdaisland.glogi :as log]
            [promesa.core :as p]))

(def ^:private graph-base "https://graph.microsoft.com/v1.0")

;; ---- Helpers ----

(defn onenote-base
  "Returns the OneNote API base URL for the given site.
   nil site-id → /me/onenote (personal)
   string site-id → /sites/{id}/onenote (SharePoint)"
  [site-id]
  (if (seq site-id)
    (str graph-base "/sites/" site-id "/onenote")
    (str graph-base "/me/onenote")))

(defn- graph-fetch
  "Authenticated GET/POST/DELETE to Microsoft Graph API.
   Returns parsed JSON as ClojureScript map, or raw response for special types."
  [token url & {:keys [method body content-type response-type]
                :or {method "GET"}}]
  (let [headers (cond-> {"Authorization" (str "Bearer " token)}
                  content-type (assoc "Content-Type" content-type))
        opts (cond-> {:method method
                      :headers headers}
               body (assoc :body body))]
    (p/let [resp (js/fetch url (clj->js opts))]
      (if (.-ok resp)
        (case response-type
          :text (.text resp)
          :arraybuffer (.arrayBuffer resp)
          ;; default: json
          (p/let [text (.text resp)]
            (when (seq text)
              (-> (js/JSON.parse text)
                  (js->clj :keywordize-keys true)))))
        (p/let [text (.text resp)]
          (log/error :onenote/api-error {:status (.-status resp) :url url :body text})
          (throw (ex-info "Graph API request failed"
                          {:status (.-status resp)
                           :url url
                           :body text})))))))

;; ---- Section operations ----

(defn find-section
  "Find a section by name in a notebook. Returns {:id :displayName} or nil."
  [token notebook-id section-name site-id]
  (p/let [base (onenote-base site-id)
          url (str base "/notebooks/" notebook-id "/sections")
          data (graph-fetch token url)]
    (some (fn [s] (when (= (:displayName s) section-name)
                    {:id (:id s) :displayName (:displayName s)}))
          (:value data))))

(defn create-section
  "Create a new section in a notebook. Returns {:id :displayName}."
  [token notebook-id section-name site-id]
  (p/let [base (onenote-base site-id)
          url (str base "/notebooks/" notebook-id "/sections")
          data (graph-fetch token url
                            :method "POST"
                            :body (js/JSON.stringify (clj->js {:displayName section-name}))
                            :content-type "application/json")]
    {:id (:id data) :displayName (:displayName data)}))

(defn find-or-create-section
  "Find a section by name, or create it if missing."
  [token notebook-id section-name site-id]
  (p/let [existing (find-section token notebook-id section-name site-id)]
    (or existing
        (create-section token notebook-id section-name site-id))))

;; ---- Page operations ----

(defn list-pages
  "List pages in a section. Returns [{:title :id :lastModifiedDateTime} ...]."
  [token section-id site-id]
  (p/let [base (onenote-base site-id)
          url (str base "/sections/" section-id "/pages?$orderby=title")
          data (graph-fetch token url)]
    (mapv (fn [p] {:title (:title p)
                   :id (:id p)
                   :lastModifiedDateTime (:lastModifiedDateTime p)})
          (:value data))))

(defn find-page
  "Find a page by title in a section. Returns {:title :id} or nil."
  [token section-id page-title site-id]
  (p/let [pages (list-pages token section-id site-id)]
    (some (fn [p] (when (= (:title p) page-title) p))
          pages)))

(defn get-page-zip-url
  "Get the ZIP attachment URL from a OneNote page.
   Fetches page HTML, parses the <object data-attachment$='.zip'> tag,
   and returns the download URL."
  [token page-id site-id]
  (p/let [base (onenote-base site-id)
          url (str base "/pages/" page-id "/content")
          html (graph-fetch token url :response-type :text)
          doc (.parseFromString (js/DOMParser.) html "text/html")
          obj (.querySelector doc "object[data-attachment$=\".zip\"]")
          zip-url (when obj (.getAttribute obj "data"))]
    ;; OneNote embeds siteCollections in resource URLs for site notebooks,
    ;; but Graph API only recognizes sites. Fix the path.
    (when zip-url
      (string/replace zip-url "/siteCollections/" "/sites/"))))

(defn download-zip
  "Download a ZIP file as ArrayBuffer."
  [token zip-url]
  (graph-fetch token zip-url :response-type :arraybuffer))

(defn delete-page
  "Delete a OneNote page."
  [token page-id site-id]
  (let [base (onenote-base site-id)
        url (str base "/pages/" page-id)]
    (p/let [resp (js/fetch url (clj->js {:method "DELETE"
                                          :headers {"Authorization" (str "Bearer " token)}}))]
      (when-not (.-ok resp)
        (throw (ex-info "Delete page failed" {:status (.-status resp) :page-id page-id}))))))

(defn upload-page-with-zip
  "Create a OneNote page with a ZIP attachment using multipart upload.
   page-title: the page title (used as graph name)
   zip-arraybuffer: the ZIP file as ArrayBuffer"
  [token section-id page-title zip-arraybuffer site-id]
  (let [base (onenote-base site-id)
        url (str base "/sections/" section-id "/pages")
        boundary (str "LogseqSync" (js/Date.now))
        page-html (str "<!DOCTYPE html>\n<html><head><title>" page-title "</title></head><body>\n"
                       "<p data-graph-type=\"logseq\" data-updated=\"" (.toISOString (js/Date.)) "\">" page-title "</p>\n"
                       "<object data-attachment=\"graph.zip\" data=\"name:graph.zip\" type=\"application/zip\" />\n"
                       "</body></html>")
        text-parts (str "--" boundary "\r\n"
                        "Content-Disposition: form-data; name=\"Presentation\"\r\n"
                        "Content-Type: text/html\r\n\r\n"
                        page-html "\r\n"
                        "--" boundary "\r\n"
                        "Content-Disposition: form-data; name=\"graph.zip\"\r\n"
                        "Content-Type: application/zip\r\n\r\n")
        closing (str "\r\n--" boundary "--")
        encoder (js/TextEncoder.)
        text-bytes (.encode encoder text-parts)
        close-bytes (.encode encoder closing)
        zip-bytes (js/Uint8Array. zip-arraybuffer)
        body (js/Uint8Array. (+ (.-length text-bytes) (.-length zip-bytes) (.-length close-bytes)))]
    ;; Assemble: text-parts + zip-binary + closing
    (.set body text-bytes 0)
    (.set body zip-bytes (.-length text-bytes))
    (.set body close-bytes (+ (.-length text-bytes) (.-length zip-bytes)))
    (p/let [resp (js/fetch url (clj->js {:method "POST"
                                          :headers {"Authorization" (str "Bearer " token)
                                                    "Content-Type" (str "multipart/form-data; boundary=" boundary)}
                                          :body body}))
            data (.json resp)]
      (let [result (js->clj data :keywordize-keys true)]
        (when (:error result)
          (throw (ex-info "Upload page failed" {:error (:error result)})))
        result))))

(defn replace-page-zip
  "Replace the ZIP on an existing page (delete + wait + upload).
   If page doesn't exist, just uploads a new one."
  [token section-id page-title zip-arraybuffer site-id]
  (p/let [existing (find-page token section-id page-title site-id)
          _ (when existing
              (log/info :onenote/replacing-page {:title page-title :id (:id existing)})
              (delete-page token (:id existing) site-id)
              ;; Wait for OneNote eventual consistency
              (js/Promise. (fn [resolve] (js/setTimeout resolve 2000))))]
    (upload-page-with-zip token section-id page-title zip-arraybuffer site-id)))

;; ---- URL parsing & notebook discovery ----

(defn- parse-sourcedoc-url
  "Extract notebook GUID from a OneNote URL with sourcedoc parameter.
   Format: ...?sourcedoc={GUID}...
   Returns {:notebook-guid :page-name} or nil."
  [url]
  (let [decoded (js/decodeURIComponent url)
        guid-match (re-find #"(?i)sourcedoc=\{?([a-f0-9-]+)\}?" decoded)
        page-match (re-find #"target\([^|]+\|[^|]*?([^|/]+)\|" decoded)]
    (when guid-match
      {:notebook-guid (string/lower-case (second guid-match))
       :page-name (when page-match (second page-match))})))

(defn- parse-oneintranote-url
  "Parse a OneIntraNote-style URL to extract site-id and notebook-id directly.
   Format: .../s/{siteId}/nb/{notebookId}/...
   Returns {:site-id :notebook-id} or nil."
  [url]
  (let [decoded (js/decodeURIComponent url)
        match (re-find #"/s/([^/]+)/nb/([^/]+)" decoded)]
    (when match
      {:site-id (second match)
       :notebook-id (nth match 2)})))

(defn- parse-sharepoint-site-path
  "Extract the SharePoint host and site path from a SharePoint URL.
   Supports:
   - https://tenant.sharepoint.com/sites/SiteName/...
   - https://tenant.sharepoint.com/:o:/s/SiteName/...
   Returns {:host :site-path} or nil."
  [url]
  (let [decoded (js/decodeURIComponent url)
        ;; Match /sites/Name or /:o:/s/Name patterns
        match (or (re-find #"([\w-]+\.sharepoint\.com)/sites/([^/?#]+)" decoded)
                  (re-find #"([\w-]+\.sharepoint\.com)/:[a-z]+:/s/([^/?#]+)" decoded))]
    (when match
      {:host (second match)
       :site-path (str "/sites/" (nth match 2))})))

(defn- resolve-site-id
  "Resolve a SharePoint site URL to a Graph API site ID.
   Uses GET /sites/{host}:{path} — requires Sites.Read.All scope.
   Falls back with a helpful error if permission is denied."
  [token host site-path]
  (p/let [url (str graph-base "/sites/" host ":" site-path)
          resp (js/fetch url (clj->js {:headers {"Authorization" (str "Bearer " token)}}))]
    (if (.-ok resp)
      (p/let [data (.json resp)]
        (:id (js->clj data :keywordize-keys true)))
      (throw (ex-info (str "Cannot resolve SharePoint site. Try pasting a OneIntraNote URL instead "
                           "(format: .../s/{siteId}/nb/{notebookId}/...)")
                       {:status (.-status resp) :host host :site-path site-path})))))

(defn- list-site-notebooks
  "List all notebooks on a SharePoint site."
  [token site-id]
  (p/let [base (onenote-base site-id)
          url (str base "/notebooks?$select=displayName,id")
          data (graph-fetch token url)]
    (mapv (fn [nb] {:name (:displayName nb)
                     :id (:id nb)
                     :site-id site-id})
          (:value data))))

(defn resolve-notebook-from-url
  "Given a pasted URL, find the notebook. Supports:
   1. OneIntraNote URLs: /s/{siteId}/nb/{notebookId}/... (direct)
   2. sourcedoc URLs: ...?sourcedoc={GUID}... (search personal notebooks)
   3. SharePoint URLs: tenant.sharepoint.com/sites/Name/... (resolve site, list notebooks)
   Returns a promise of {:name :id :site-id} or nil."
  [token url]
  ;; 1. Try OneIntraNote URL format (has site-id + notebook-id directly)
  (if-let [{:keys [site-id notebook-id]} (parse-oneintranote-url url)]
    (p/let [base (onenote-base site-id)
            nb-url (str base "/notebooks/" notebook-id "?$select=displayName,id")
            data (graph-fetch token nb-url)]
      {:name (:displayName data)
       :id (:id data)
       :site-id site-id})
    ;; 2. Try SharePoint URL (resolve site → list notebooks → pick first or only)
    (if-let [{:keys [host site-path]} (parse-sharepoint-site-path url)]
      (p/let [site-id (resolve-site-id token host site-path)
              notebooks (list-site-notebooks token site-id)]
        ;; If sourcedoc GUID is in the URL, match by GUID; otherwise take first notebook
        (let [{:keys [notebook-guid]} (parse-sourcedoc-url url)]
          (or (when notebook-guid
                (some (fn [nb]
                        (when (string/includes? (string/lower-case (:id nb))
                                                notebook-guid)
                          nb))
                      notebooks))
              (first notebooks))))
      ;; 3. Fall back to sourcedoc GUID search in personal notebooks
      (when-let [{:keys [notebook-guid]} (parse-sourcedoc-url url)]
        (p/let [personal-url (str graph-base "/me/onenote/notebooks?$select=displayName,id&$top=100")
                personal-data (graph-fetch token personal-url)
                notebooks (mapv (fn [nb] {:name (:displayName nb) :id (:id nb) :site-id nil})
                                (:value personal-data))]
          (some (fn [nb]
                  (when (string/includes? (string/lower-case (:id nb)) notebook-guid)
                    nb))
                notebooks))))))
