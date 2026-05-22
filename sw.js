// Service Worker — Jantar dos Casais
const CACHE = 'jantar-v1';
const ASSETS = [
  './index.html',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,600;1,400;1,600&family=Inter:wght@300;400;500;600&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js',
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => {
      // Adiciona assets principais; ignora falhas de recursos externos
      return Promise.allSettled(ASSETS.map(url => c.add(url).catch(() => {})));
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Requisições ao Google Apps Script sempre vão direto à rede
  if (e.request.url.includes('script.google.com')) {
    e.respondWith(fetch(e.request).catch(() => new Response('{"ok":false,"error":"offline"}', {headers:{'Content-Type':'application/json'}})));
    return;
  }
  // Demais recursos: cache-first
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (res && res.status === 200 && res.type === 'basic') {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => cached || new Response('Offline', {status: 503}));
    })
  );
});
