// Offline-first service worker for the app shell.
// Bump CACHE when any app file changes so installed copies update.
const CACHE = 'giving-envelopes-v3';
const SHELL = [
  './',
  'index.html',
  'styles.css',
  'app.js',
  'xlsx-reader.js',
  'manifest.webmanifest',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request).then((hit) => {
      if (hit) return hit;
      return fetch(e.request).catch(() => (e.request.mode === 'navigate' ? caches.match('index.html') : Response.error()));
    })
  );
});
