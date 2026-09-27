// 画面の静的資産だけを保存し、注文・名札・API応答はキャッシュしない。
const CACHE='beaufes-booth-shell-0.5.0';
const FILES=['./','./index.html','./app.js','./storage.js','./export.js','./style.css','./jsQR.js','./exceljs.min.js'];
self.addEventListener('install',event=>event.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES)).then(()=>self.skipWaiting())));
self.addEventListener('activate',event=>event.waitUntil((async()=>{for(const key of await caches.keys())if(key.startsWith('beaufes-booth-shell-')&&key!==CACHE)await caches.delete(key);await self.clients.claim();})()));
self.addEventListener('fetch',event=>{const u=new URL(event.request.url);if(event.request.method!=='GET'||u.origin!==self.location.origin||!FILES.some(f=>new URL(f,self.location.href).pathname===u.pathname))return;event.respondWith(caches.open(CACHE).then(async cache=>(await cache.match(event.request,{ignoreSearch:true}))||fetch(event.request)));});
