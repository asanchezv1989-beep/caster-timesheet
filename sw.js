// Red primero, caché de respaldo: funciona sin internet y toma actualizaciones solo.
const CACHE = 'caster-ts-v3';
const SHELL = ['./', 'index.html', 'styles.css', 'manifest.webmanifest', 'vendor/jszip.min.js',
  'js/app.js', 'js/calc.js', 'js/xlsx.js', 'js/db.js', 'js/i18n.js', 'js/rules.js', 'js/graph.js', 'vendor/msal-browser.min.js', 'js/colref.js', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/mark.png'];
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(
    fetch(e.request).then((r) => {
      const copy = r.clone();
      caches.open(CACHE).then((c) => c.put(e.request, copy));
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match('index.html')))
  );
});
