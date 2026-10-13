const CACHE = 'kk-growth-record-v1';
self.addEventListener('install', event => {
  const base = self.registration.scope;
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll([
    base,
    `${base}index.html`,
    `${base}manifest.webmanifest`,
    `${base}icon.svg`
  ])));
  self.skipWaiting();
});
self.addEventListener('activate', event => event.waitUntil(self.clients.claim()));
self.addEventListener('fetch', event => {
  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request)));
});
