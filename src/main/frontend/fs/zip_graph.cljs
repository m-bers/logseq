(ns frontend.fs.zip-graph
  "ZIP pack/unpack utilities for Logseq graphs.
   Bridges between the local filesystem (NFS/File System Access API) and JSZip archives."
  (:require ["jszip" :as JSZip]
            [clojure.string :as string]
            [frontend.fs :as fs]
            [frontend.fs.nfs :as nfs]
            [logseq.common.path :as path]
            [promesa.core :as p]))

;; ---- Path filtering ----

(defn skip-path?
  "Returns true for paths that should be excluded from ZIP sync."
  [relative-path]
  (or (string/starts-with? relative-path "logseq/bak/")
      (string/starts-with? relative-path "logseq/.recycle/")
      (string/starts-with? relative-path ".logseq-onenote-base/")
      (string/starts-with? relative-path ".git/")
      (string/starts-with? relative-path "node_modules/")
      (string/starts-with? relative-path "logseq/version-files/")))

;; ---- Pack (local folder → ZIP) ----

(defn <pack-graph
  "Read all files from a local folder graph and pack into a ZIP.
   repo-dir: the graph directory name (e.g. 'my-notes')
   Returns Promise<ArrayBuffer>."
  [repo-dir]
  (let [zip (JSZip.)]
    (p/let [{:keys [files]} (fs/get-files repo-dir)]
      (p/let [_ (p/all
                 (map (fn [{:keys [path content]}]
                        (when (and (not (skip-path? path))
                                   (some? content))
                          (.file zip path content)))
                      files))]
        (.generateAsync zip #js {:type "arraybuffer"})))))

;; ---- Unpack (ZIP → map) ----

(defn <unpack-zip
  "Unpack a ZIP ArrayBuffer into a map of {relative-path -> content-string}.
   Skips directories and filtered paths."
  [arraybuffer]
  (p/let [zip (.loadAsync JSZip arraybuffer)
          entries (js/Object.entries (.-files zip))
          results (p/all
                   (map (fn [entry]
                          (let [name (aget entry 0)
                                file (aget entry 1)]
                            (when (and (not (.-dir file))
                                       (not (skip-path? name)))
                              (p/let [content (.async file "string")]
                                [name content]))))
                        entries))]
    (into {} (remove nil? results))))

;; ---- Read local files to map ----

(defn <read-local-files
  "Read all files from a local folder graph into a map.
   repo-dir: the graph directory name
   Returns {relative-path -> content-string}."
  [repo-dir]
  (p/let [{:keys [files]} (fs/get-files repo-dir)]
    (into {} (keep (fn [{:keys [path content]}]
                     (when (and (not (skip-path? path))
                                (some? content))
                       [path content]))
                   files))))

;; ---- Write files to local folder ----

(defn <write-file-to-local
  "Write a single file to the local folder graph via NFS handles.
   repo-dir: the graph directory name
   rel-path: relative path within the graph (e.g. 'pages/foo.md')
   content: file content string"
  [repo repo-dir rel-path content]
  (fs/write-plain-text-file! repo repo-dir rel-path content
                              {:skip-compare? true}))
