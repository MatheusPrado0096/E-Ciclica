// E-Cíclica • service worker: rede primeiro, cache como reserva (funciona offline)
const C = 'eciclica-v8';
self.addEventListener('install', (e) => { self.skipWaiting(); e.waitUntil(caches.open(C).then((c) => c.addAll(['./', 'index.html', 'config.js', 'assets/css/style.css', 'assets/js/app.js', 'assets/js/xlsx.js', 'assets/js/store-local.js', 'assets/js/store-firebase.js', 'assets/vendor/jszip.min.js', 'assets/img/mrs-mark.png', 'assets/img/mrs-logo.png']).catch(() => {}))); });
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return; // Firebase e CDN passam direto
  e.respondWith(fetch(e.request).then((r) => { const cp = r.clone(); caches.open(C).then((c) => c.put(e.request, cp)); return r; })
    .catch(() => caches.match(e.request, { ignoreSearch: true })));
});
