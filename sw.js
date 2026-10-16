const CACHE_NAME = 'insentif-v1';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json'
];

self.addEventListener('install', event=>{
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache=>{
      return cache.addAll(APP_SHELL);
    })
  );
});

self.addEventListener('activate', event=>{
  event.waitUntil(
    caches.keys().then(keys=>{
      return Promise.all(keys.filter(k=>k !== CACHE_NAME).map(k=>caches.delete(k)));
    })
  );
});

self.addEventListener('fetch', event=>{
  event.respondWith(
    caches.match(event.request).then(response=>{
      return response || fetch(event.request);
    })
  );
});
