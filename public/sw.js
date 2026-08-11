// SAHI London Catalogue — Service Worker
const CACHE = 'sahi-catalogue-v1';

// Assets to pre-cache on install
const PRE_CACHE = [
  '/catalogue.html?sales',
  '/catalogue.html',
  '/manifest.json',
  '/images/pwa/icon-192.png',
  '/images/pwa/icon-512.png'
];

// Install: pre-cache static assets
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(PRE_CACHE).catch(() => {}))
  );
  self.skipWaiting();
});

// Activate: clean old caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch: cache-first for images, network-first for HTML
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Skip non-GET
  if (event.request.method !== 'GET') return;

  // Skip Shopify CDN (images are on Shopify, not our origin)
  if (url.hostname.includes('cdn.shopify.com')) {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        const fetchPromise = fetch(event.request).then((response) => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE).then((cache) => cache.put(event.request, clone));
          }
          return response;
        });
        return cached || fetchPromise;
      })
    );
    return;
  }

  // For our own origin: network-first for HTML, cache-first for static
  if (url.origin === self.location.origin) {
    const isHtml = event.request.headers.get('accept')?.includes('text/html') || url.pathname.endsWith('.html');

    if (isHtml) {
      event.respondWith(
        fetch(event.request).then((response) => {
          const clone = response.clone();
          caches.open(CACHE).then((cache) => cache.put(event.request, clone));
          return response;
        }).catch(() => caches.match(event.request))
      );
    } else {
      event.respondWith(
        caches.match(event.request).then((cached) => cached || fetch(event.request))
      );
    }
  }
});
