// 画面だけを保存。注文・名札・API応答・認証情報はキャッシュしない。
const CACHE='beaufes-booth-shell-0.12.0';
const FILES=['./','./index.html','./app.js','./storage.js','./export.js','./style.css','./jsQR.js','./exceljs.min.js'];
self.addEventListener('install',event=>event.waitUntil((async()=>{
 const cache=await caches.open(CACHE);
 // ブラウザーのHTTPキャッシュに残った旧版を、新しい保存領域へ持ち込まない。
 await cache.addAll(FILES.map(path=>new Request(new URL(path,self.location.href),{cache:'no-store'})));
 await self.skipWaiting();
})()));
self.addEventListener('activate',event=>event.waitUntil((async()=>{
 for(const key of await caches.keys())if(key.startsWith('beaufes-booth-shell-')&&key!==CACHE)await caches.delete(key);
 await self.clients.claim();
})()));
self.addEventListener('message',event=>{
 if(event.data==='SHELL_VERSION')event.ports[0]?.postMessage('v'+CACHE.replace('beaufes-booth-shell-',''));
});
self.addEventListener('fetch',event=>{
 const url=new URL(event.request.url);
 if(event.request.method!=='GET'||url.origin!==self.location.origin)return;
 const file=FILES.find(path=>new URL(path,self.location.href).pathname===url.pathname);
 if(!file)return;
 event.respondWith((async()=>{
  const cache=await caches.open(CACHE);const key=new URL(file,self.location.href).href;
  // オンラインでは毎回最新版を優先し、通信断やサーバー障害時だけ保存済み画面へ戻す。
  let timer;
  const network=(async()=>{
   const response=await fetch(new Request(event.request,{cache:'no-store'}));
   if(!response.ok)throw new Error('SHELL_UNAVAILABLE');
   await cache.put(key,response.clone());return response;
  })();
  try{
   return await Promise.race([network,new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('SHELL_TIMEOUT')),4000);})]);
  }catch(error){
   const saved=await cache.match(key);if(saved)return saved;throw error;
  }finally{clearTimeout(timer);}
 })());
});
