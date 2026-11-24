/* Service worker: makes the dashboard installable and usable offline.
   Bump VERSION whenever app files change so users get the update. */
const VERSION = 'bulk-v19';
const PRECACHE = [
  './',
  './index.html',
  './manifest.webmanifest',
  './css/styles.css',
  './js/app.js',
  './js/data.js',
  './js/upload.js',
  './js/pwa.js',
  './js/vendor/chart.umd.js',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png',
];

self.addEventListener('install', event => {
  // cache: 'reload' bypasses the browser's HTTP cache, so a new version never re-saves old files.
  event.waitUntil(caches.open(VERSION).then(c => c.addAll(PRECACHE.map(u => new Request(u, { cache: 'reload' })))));
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('message', event => {
  if (event.data === 'skipWaiting') self.skipWaiting();
});

// Stale-while-revalidate: answer from cache instantly, refresh the cache in the background.
async function staleWhileRevalidate(request) {
  const cache = await caches.open(VERSION);
  const cached = await cache.match(request, { ignoreSearch: true });
  const sameOrigin = new URL(request.url).origin === location.origin;
  // 'no-cache' makes the browser check with the server (cheap 304 if unchanged) instead of using a stale HTTP-cache copy.
  const network = fetch(sameOrigin ? new Request(request, { cache: 'no-cache' }) : request).then(res => {
    if (res && (res.ok || res.type === 'opaque')) cache.put(request, res.clone());
    return res;
  }).catch(() => null);
  return cached || (await network) || Response.error();
}

async function networkFirst(request) {
  const cache = await caches.open(VERSION);
  try {
    const res = await Promise.race([fetch(new Request(request, { cache: 'no-cache' })), new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 4000))]);
    if (res && res.ok) cache.put(request, res.clone());
    return res;
  } catch {
    return (await cache.match(request, { ignoreSearch: true })) || Response.error();
  }
}

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Page navigations: network first so new versions show up, cached app shell when offline.
  if (req.mode === 'navigate') {
    event.respondWith(fetch(req, { cache: 'no-cache' }).catch(async () => (await caches.match('./index.html')) || Response.error()));
    return;
  }
  // App files: network first so the page, scripts and styles always come from the same version.
  // Falls back to the saved copy when offline or when the network takes longer than 4 s.
  if (url.origin === location.origin) { event.respondWith(networkFirst(req)); return; }
  // Google Fonts rarely change: answer from cache, refresh in the background.
  if (/fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)) event.respondWith(staleWhileRevalidate(req));
});
