// Minimal Service Worker enabling PWA standalone installation & full mobile app view
const CACHE_NAME = 'whos-that-saint-pwa-v1';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(clients.claim());
});

self.addEventListener('fetch', (event) => {
  // Let network requests pass directly to ensure fresh cloud sync
  event.respondWith(fetch(event.request).catch(() => caches.match(event.request)));
});
