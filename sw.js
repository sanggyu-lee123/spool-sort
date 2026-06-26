const CACHE_NAME = 'spool-sort-v1';
const ASSETS = [
  './',
  './index.html',
  './app.js',
  './convert-engine.js',
  './xlsx-print-patch.js',
  './manifest.json',
  './vendor/xlsx.full.min.js',
  './vendor/jszip.min.js',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// network-first: 항상 최신 버전을 우선 받아오고, 네트워크가 없을 때만 캐시로 대체
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
