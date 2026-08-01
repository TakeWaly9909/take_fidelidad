// Service worker de la PWA Take Away. Sube la versión del cache cuando cambien los archivos estáticos.
const CACHE_NAME = 'take-away-v1';
const ASSETS = [
  'index.html',
  'registro.html',
  'tarjeta.html',
  'admin.html',
  'config.js',
  'manifest.json',
  'manifest-admin.json',
  'img/logo.png',
  'img/sello-lleno.png',
  'img/sello-vacio.png',
  'img/icon-192.png',
  'img/icon-512.png',
  'img/icon-192-maskable.png',
  'img/icon-512-maskable.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // La API de Google Apps Script (script.google.com) nunca se cachea: siempre va a la red.
  if (url.origin !== self.location.origin) return;

  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request).catch(() =>
        caches.match(event.request).then((cached) => cached || caches.match('index.html'))
      )
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
