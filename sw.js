const CACHE_NAME = 'ginjian-pwa-v1';
const OFFLINE_URL = new URL('offline.html', self.registration.scope).href;
const ASSETS = ['offline.html', 'manifest.webmanifest', 'pwa.js',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png',
  'icons/apple-touch-icon.png', 'icons/favicon-32.png'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(
    ASSETS.map(path => new URL(path, self.registration.scope).href)
  )));
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(key => key.startsWith('ginjian-pwa-') && key !== CACHE_NAME)
      .map(key => caches.delete(key)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const request = event.request;
  const url = new URL(request.url);
  // Do not cache external calendars, APIs, images or other sites on this origin.
  if (request.method !== 'GET' || url.origin !== self.location.origin ||
      !url.href.startsWith(self.registration.scope)) return;
  if (request.mode === 'navigate') {
    event.respondWith(fetch(request).catch(() => caches.match(OFFLINE_URL)));
    return;
  }
  if (ASSETS.some(path => new URL(path, self.registration.scope).href === url.href)) {
    event.respondWith(caches.match(request).then(cached => cached || fetch(request)));
  }
});
