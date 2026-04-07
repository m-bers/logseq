(ns frontend.fs.onedrive
  "Microsoft Graph API client for OneDrive file operations.
   All functions take an access token and return promises."
  (:require [promesa.core :as p]
            [cljs-bean.core :as bean]
            [clojure.string :as string]
            [lambdaisland.glogi :as log]))

(def ^:private graph-base "https://graph.microsoft.com/v1.0")

(defn- graph-fetch
  "Make an authenticated request to Microsoft Graph API.
   Returns parsed JSON as ClojureScript map."
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
          :blob (.blob resp)
          :arraybuffer (.arrayBuffer resp)
          ;; default: json
          (p/let [text (.text resp)]
            (when (seq text)
              (-> (js/JSON.parse text)
                  (js->clj :keywordize-keys true)))))
        (p/let [text (.text resp)]
          (log/error :onedrive/api-error {:status (.-status resp) :url url :body text})
          (throw (ex-info "Graph API request failed"
                          {:status (.-status resp)
                           :url url
                           :body text})))))))

(defn- encode-path
  "Encode a path for use in Graph API URLs.
   Handles special characters but preserves /."
  [path]
  (-> path
      (string/replace #"[#%]" (fn [c] (js/encodeURIComponent c)))))

(defn- drive-item-url
  "Build a Graph API URL for a drive item by path.
   folder-path: e.g. 'Notes' or 'Notes/pages'"
  [path & [suffix]]
  (let [encoded (encode-path path)]
    (str graph-base "/me/drive/root:/" encoded (when suffix (str ":/" suffix)))))

;; ---- Public API ----

(defn list-children
  "List files and folders in a directory.
   Returns a seq of maps with :name, :id, :size, :lastModifiedDateTime, :folder?, :path"
  [token folder-path]
  (p/let [url (drive-item-url folder-path "children?$top=1000")
          data (graph-fetch token url)]
    (->> (:value data)
         (map (fn [item]
                {:name (:name item)
                 :id (:id item)
                 :size (:size item)
                 :mtime (:lastModifiedDateTime item)
                 :folder? (boolean (:folder item))
                 :path (str folder-path "/" (:name item))})))))

(defn list-files-recursive
  "Recursively list all files under a directory.
   Returns a flat seq of file maps (no folders)."
  [token folder-path]
  (p/let [children (list-children token folder-path)]
    (p/let [results
            (p/all
             (map (fn [child]
                    (if (:folder? child)
                      (list-files-recursive token (:path child))
                      (p/resolved [child])))
                  children))]
      (apply concat results))))

(defn read-file
  "Read a file's content as text.
   file-path: e.g. 'Notes/pages/my-page.md'"
  [token file-path]
  (let [url (drive-item-url file-path "content")]
    (graph-fetch token url :response-type :text)))

(defn read-file-raw
  "Read a file's content as ArrayBuffer (for binary files)."
  [token file-path]
  (let [url (drive-item-url file-path "content")]
    (graph-fetch token url :response-type :arraybuffer)))

(defn write-file
  "Write content to a file (creates or overwrites).
   content: string or ArrayBuffer"
  [token file-path content]
  (let [url (drive-item-url file-path "content")
        content-type (if (string? content)
                       "text/plain"
                       "application/octet-stream")]
    (graph-fetch token url
                 :method "PUT"
                 :body content
                 :content-type content-type)))

(defn delete-file
  "Delete a file or folder."
  [token file-path]
  (let [url (drive-item-url file-path)]
    (p/let [resp (js/fetch url (clj->js {:method "DELETE"
                                          :headers {"Authorization" (str "Bearer " token)}}))]
      (when-not (.-ok resp)
        (throw (ex-info "Delete failed" {:status (.-status resp) :path file-path}))))))

(defn mkdir
  "Create a folder. parent-path is the parent, folder-name is the new folder."
  [token parent-path folder-name]
  (let [url (drive-item-url parent-path "children")]
    (graph-fetch token url
                 :method "POST"
                 :body (js/JSON.stringify
                        (clj->js {:name folder-name
                                  :folder {}
                                  "@microsoft.graph.conflictBehavior" "fail"}))
                 :content-type "application/json")))

(defn get-delta
  "Get changes since last sync using the delta API.
   delta-url: nil for first sync, or the @odata.deltaLink from previous call.
   folder-path: e.g. 'Notes'
   Returns {:changes [...] :delta-link \"...\"}"
  [token folder-path delta-url]
  (p/let [url (or delta-url
               (str graph-base "/me/drive/root:/" (encode-path folder-path) ":/delta"))
          result (p/loop [url url
                          all-changes []]
                   (p/let [data (graph-fetch token url)
                           changes (into all-changes (:value data))
                           next-link (get data (keyword "@odata.nextLink"))
                           delta-link (get data (keyword "@odata.deltaLink"))]
                     (if next-link
                       (p/recur next-link changes)
                       {:changes changes
                        :delta-link delta-link})))]
    result))

(defn get-item-metadata
  "Get metadata for a single item."
  [token file-path]
  (graph-fetch token (drive-item-url file-path)))
