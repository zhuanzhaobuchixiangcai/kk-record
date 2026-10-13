const CACHE = 'kk-growth-record-v1';
self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(['/', '/index.html', '/manifest.webmanifest', '/icon.svg'])));
  self.skipWaiting();
});
self.addEventListener('activate', event => event.waitUntil(self.clients.claim()));
self.addEventListener('fetch', event => {
  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request)));
});
