// Raporlama Portali - Service Worker
// Blazor Server icin minimal SW: statik dosyalari cache'ler, SignalR'a karismaz.
// Cevrimdisi durumda basit bir bilgi sayfasi gosterir.

const CACHE_NAME = 'raporlama-static-v1';
const STATIC_ASSETS = [
  '/css/site.css',
  '/js/site.js',
  '/images/icon-192.png',
  '/images/icon-512.png',
  '/images/favicon-32.png',
  '/manifest.json'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);

  // SignalR, framework, API ve giris/auth isteklerine dokunma - dogrudan agdan al
  if (
    url.pathname.startsWith('/_blazor') ||
    url.pathname.startsWith('/_framework') ||
    url.pathname.startsWith('/api/') ||
    url.pathname.startsWith('/Giris') ||
    url.pathname.startsWith('/giris') ||
    req.method !== 'GET'
  ) {
    return; // varsayilan ag davranisi
  }

  // Statik dosyalar: cache-first
  if (
    url.pathname.startsWith('/css/') ||
    url.pathname.startsWith('/js/') ||
    url.pathname.startsWith('/images/') ||
    url.pathname.startsWith('/_content/') ||
    url.pathname === '/manifest.json'
  ) {
    event.respondWith(
      caches.match(req).then(cached => {
        if (cached) return cached;
        return fetch(req).then(resp => {
          if (resp.ok) {
            const clone = resp.clone();
            caches.open(CACHE_NAME).then(c => c.put(req, clone));
          }
          return resp;
        });
      })
    );
    return;
  }

  // Navigasyon istekleri (HTML): network-first, offline ise basit mesaj
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).catch(() => new Response(
        '<!DOCTYPE html><html lang="tr"><head><meta charset="utf-8">' +
        '<meta name="viewport" content="width=device-width, initial-scale=1">' +
        '<title>Cevrimdisi - Raporlama Portali</title>' +
        '<style>body{font-family:Roboto,sans-serif;background:#1B5E20;color:#fff;' +
        'display:flex;align-items:center;justify-content:center;min-height:100vh;' +
        'margin:0;padding:24px;text-align:center}' +
        '.box{max-width:400px}' +
        'h1{margin:0 0 16px;font-size:24px}' +
        'p{margin:0 0 24px;line-height:1.5;opacity:.9}' +
        'button{background:#fff;color:#1B5E20;border:none;padding:12px 24px;' +
        'border-radius:8px;font-weight:600;cursor:pointer;font-size:16px}</style>' +
        '</head><body><div class="box">' +
        '<h1>Baglanti yok</h1>' +
        '<p>Sunucuya ulasilamiyor. VPN baglantini ve ag baglantini kontrol et, sonra tekrar dene.</p>' +
        '<button onclick="location.reload()">Tekrar Dene</button>' +
        '</div></body></html>',
        { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
      ))
    );
  }
});
