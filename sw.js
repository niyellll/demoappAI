const CACHE = 'rute-gpp-v1';
const APP_FILES = ['/', '/manifest.webmanifest', '/icon-192.png', '/icon-512.png'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(
    APP_FILES.map(path => new Request(path, { cache: 'reload' }))
  )));
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter(name => name.startsWith('rute-gpp-') && name !== CACHE)
      .map(name => caches.delete(name)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const { request } = event;
  const url = new URL(request.url);
  if (request.method !== 'GET' || url.origin !== self.location.origin) return;
  if (request.mode === 'navigate' && ['/', '/index.html'].includes(url.pathname)) {
    event.respondWith((async () => {
      try {
        const response = await fetch(request);
        if (response.ok) {
          const cache = await caches.open(CACHE);
          await cache.put('/', response.clone());
        }
        return response;
      } catch {
        return (await caches.match('/')) ?? Response.error();
      }
    })());
  } else if (APP_FILES.includes(url.pathname)) {
    event.respondWith(fetch(request).catch(async () =>
      (await caches.match(url.pathname)) ?? Response.error()
    ));
  }
});
