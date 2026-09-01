const CACHE="contabilidad-taxi-cloud-v0.4";
const ASSETS=[
  "./","./index.html","./css/estilos.css","./js/app.js","./js/config.js",
  "./js/auth.js","./js/onedrive.js","./js/calculos.js","./js/storage.js","./js/cloud-data.js",
  "./manifest.webmanifest","./icons/icon-192.png","./icons/icon-512.png","./icons/favicon.png"
];
self.addEventListener("install",event=>{
  event.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS)));
  self.skipWaiting();
});
self.addEventListener("activate",event=>{
  event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener("fetch",event=>{
  if(event.request.method!=="GET") return;
  event.respondWith(
    fetch(event.request).then(r=>{
      const copy=r.clone(); caches.open(CACHE).then(c=>c.put(event.request,copy)); return r;
    }).catch(()=>caches.match(event.request).then(r=>r||caches.match("./index.html")))
  );
});
