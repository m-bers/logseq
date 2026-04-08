(ns frontend.fs.zip-graph
  "ZIP pack/unpack utilities for Logseq graphs.
   Bridges between LightningFS (in-browser filesystem) and JSZip archives."
  (:require ["jszip" :as JSZip]
            [clojure.string :as string]
            [logseq.common.path :as path]
            [promesa.core :as p]))

;; ---- Path filtering ----

(defn skip-path?
  "Returns true for paths that should be excluded from ZIP sync."
  [relative-path]
  (or (string/starts-with? relative-path "logseq/bak/")
      (string/starts-with? relative-path "logseq/.recycle/")
      (string/starts-with? relative-path ".git/")
      (string/starts-with? relative-path "node_modules/")
      (string/starts-with? relative-path "logseq/version-files/")))

;; ---- LightningFS helpers ----

(defn- <readdir-recursive
  "Read all file paths from LightningFS under dir recursively.
   Returns a vector of absolute paths."
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

(defn- <mkdir-recursive!
  "Create a directory and all parent directories."
  [dir]
  (-> (js/window.pfs.stat dir)
      (p/then (fn [_] nil))
      (p/catch
       (fn [_]
         (p/let [parent (path/parent dir)
                 _ (when (and parent (not= parent dir))
                     (<mkdir-recursive! parent))]
           (-> (js/window.pfs.mkdir dir)
               (p/catch (fn [_] nil))))))))

;; ---- Pack (LightningFS → ZIP) ----

(defn <pack-graph
  "Read all files from a LightningFS directory and pack into a ZIP.
   local-dir: memory:// URL, e.g. 'memory:///onenote-Notes'
   Returns Promise<ArrayBuffer>."
  [local-dir]
  (let [root (path/url-to-path local-dir)
        zip (JSZip.)]
    (p/let [all-paths (<readdir-recursive root)
            _ (p/all
               (map (fn [fpath]
                      (let [rel-path (subs fpath (inc (count root)))]
                        (when-not (skip-path? rel-path)
                          (p/let [content (.readFile js/window.pfs fpath)]
                            (.file zip rel-path content)))))
                    all-paths))]
      (.generateAsync zip #js {:type "arraybuffer"}))))

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

;; ---- Unpack (ZIP → LightningFS) ----

(defn <unpack-zip-to-fs
  "Unpack a ZIP ArrayBuffer directly into a LightningFS directory.
   target-dir: absolute LightningFS path, e.g. '/onenote-Notes'
   Returns the file map {relative-path -> content}."
  [arraybuffer target-dir]
  (p/let [file-map (<unpack-zip arraybuffer)
          _ (p/all
             (map (fn [[rel-path content]]
                    (let [full-path (path/path-join target-dir rel-path)
                          parent (path/parent full-path)]
                      (p/let [_ (<mkdir-recursive! parent)]
                        (.writeFile js/window.pfs full-path content))))
                  file-map))]
    file-map))

;; ---- Read local files to map ----

(defn <read-local-files
  "Read all files from a LightningFS directory into a map.
   local-dir: memory:// URL, e.g. 'memory:///onenote-Notes'
   Returns {relative-path -> content-string}."
  [local-dir]
  (let [root (path/url-to-path local-dir)]
    (p/let [all-paths (<readdir-recursive root)
            results (p/all
                     (map (fn [fpath]
                            (let [rel-path (subs fpath (inc (count root)))]
                              (when-not (skip-path? rel-path)
                                (p/let [content (-> (.readFile js/window.pfs fpath #js {:encoding "utf8"})
                                                    (p/then (fn [c] (.toString c))))]
                                  [rel-path content]))))
                          all-paths))]
      (into {} (remove nil? results)))))
