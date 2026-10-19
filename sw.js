const CACHE_NAME = 'insentif-v3';
const APP_SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './vendor/xlsx.full.min.js', './vendor/html2canvas.min.js'];

self.addEventListener('install', event=>{
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.addAll(APP_SHELL)));
});

self.addEventListener('activate', event=>{
  event.waitUntil(
    caches.keys()
      .then(keys=>Promise.all(keys.filter(k=>k !== CACHE_NAME).map(k=>caches.delete(k))))
      .then(()=>self.clients.claim())
  );
});

// Network-first: selalu ambil versi terbaru dari server, cache hanya cadangan saat offline.
self.addEventListener('fetch', event=>{
  const req = event.request;
  if(req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  event.respondWith(
    fetch(req)
      .then(res=>{
        const copy = res.clone();
        caches.open(CACHE_NAME).then(cache=>cache.put(req, copy)).catch(()=>{});
        return res;
      })
      .catch(()=>caches.match(req).then(hit=>hit || caches.match('./index.html')))
  );
});
