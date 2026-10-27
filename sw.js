const V='tanaoroshi-v7';
const SHELL=['./','index.html','style.css','app.js','manifest.webmanifest','icon.svg','qr.png','icon-192.png','icon-512.png','vendor/zxing.min.js','vendor/exceljs.min.js','vendor/jszip.min.js','vendor/qrcode.js'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==V).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{ if(e.request.method!=='GET'||new URL(e.request.url).origin!==location.origin||new URL(e.request.url).pathname.endsWith('list.enc')) return;
  e.respondWith(caches.open(V).then(async c=>{ const hit=await c.match(e.request); const net=fetch(e.request).then(r=>{ if(r.ok) c.put(e.request,r.clone()); return r; }).catch(()=>hit); return hit||net; })); });
