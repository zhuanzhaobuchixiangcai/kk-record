const CACHE = 'kk-growth-record-v2';
self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const base = self.registration.scope;
    const response = await fetch(`${base}index.html`, { cache: 'no-store' });
    if (!response.ok) throw new Error('无法加载应用入口');
    const html = await response.text();
    const assets = [...html.matchAll(/(?:src|href)="([^\"]+\.(?:js|css))"/g)]
      .map(match => new URL(match[1], base).href);
    const cache = await caches.open(CACHE);
    await cache.addAll([base, `${base}index.html`, `${base}manifest.webmanifest`, `${base}icon.svg`, ...assets]);
  })());
  self.skipWaiting();
});
self.addEventListener('activate', event => event.waitUntil((async () => {
  const keys = await caches.keys();
  await Promise.all(keys.filter(key => key.startsWith('kk-growth-record-') && key !== CACHE).map(key => caches.delete(key)));
  await self.clients.claim();
})()));
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET' || !event.request.url.startsWith(self.registration.scope)) return;
  if (event.request.mode === 'navigate') {
    event.respondWith(fetch(event.request).catch(() => caches.match(event.request)));
    return;
  }
  event.respondWith(caches.match(event.request, { ignoreVary: true }).then(cached => cached || fetch(event.request)));
});
