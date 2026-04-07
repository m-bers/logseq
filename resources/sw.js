const CACHE_NAME = 'logseq-pwa-v1';

// Derive base path from service worker location
// e.g. /logseq/sw.js → /logseq/
const SW_PATH = self.location.pathname;
const BASE = SW_PATH.substring(0, SW_PATH.lastIndexOf('/') + 1);

const PRECACHE_PATHS = [
  '',
  'index.html',
  'css/style.css',
  'img/logo.png',
  'img/logo-192x192.png',
  'img/logo-512x512.png',
  'js/magic_portal.js',
  'js/worker.js',
  'js/main.js',
  'js/ui.js',
  'js/react.production.min.js',
  'js/react-dom.production.min.js',
  'js/highlight.min.js',
  'js/interact.min.js',
  'js/marked.umd.js',
  'js/eventemitter3.umd.min.js',
  'js/html2canvas.min.js',
  'js/lsplugin.core.js',
  'js/prop-types.min.js',
  'js/tabler-icons-react.min.js',
  'js/tabler.ext.js',
  'js/code-editor.js',
];

const PRECACHE_URLS = PRECACHE_PATHS.map((p) => BASE + p);

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(PRECACHE_URLS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) =>
      Promise.all(
        cacheNames
          .filter((name) => name !== CACHE_NAME)
          .map((name) => caches.delete(name))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Don't cache Graph API or external requests
  if (url.origin !== self.location.origin) return;

  // NetworkFirst for the app shell (index.html)
  if (url.pathname === BASE || url.pathname === BASE + 'index.html') {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // StaleWhileRevalidate for JS/CSS
  if (event.request.destination === 'script' || event.request.destination === 'style') {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        const fetched = fetch(event.request).then((response) => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          return response;
        });
        return cached || fetched;
      })
    );
    return;
  }

  // CacheFirst for images
  if (event.request.destination === 'image') {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        return cached || fetch(event.request).then((response) => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          return response;
        });
      })
    );
    return;
  }

  // Default: try network, fall back to cache
  event.respondWith(
    fetch(event.request)
      .catch(() => caches.match(event.request))
  );
});
