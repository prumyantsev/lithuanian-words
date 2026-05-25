const CACHE = 'lt-words-v2';
const ASSETS = [
  './index.html',
  './app.js',
  './style.css',
  './manifest.json',
  './images.json',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Network-first strategy for images: try network, fall back to cache
  if (url.pathname.startsWith('/images/') || url.pathname.includes('/images/')) {
    e.respondWith(
      fetch(e.request)
        .then(r => {
          if (r.ok) {
            // Cache successful image responses for offline use
            caches.open(CACHE).then(c => c.put(e.request, r.clone()));
          }
          return r;
        })
        .catch(() => caches.match(e.request))
    );
  } else {
    // Cache-first for other assets (faster loading)
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
  }
});
