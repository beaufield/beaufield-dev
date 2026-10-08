// 注文データや認証は保持し、注文アプリの静的画面だけを再取得する。
const join=new URL(location.href).hash.match(/^#join=([-_A-Za-z0-9]{43,128})$/)?.[1];
if(location.hash)history.replaceState(null,'',location.pathname+location.search);
const status=document.getElementById('update-status');const retry=document.getElementById('retry');
async function fresh(url){
 const controller=new AbortController();const timer=setTimeout(()=>controller.abort(),15000);
 try{return await fetch(url,{cache:'no-store',signal:controller.signal});}finally{clearTimeout(timer);}
}
async function refresh(){
 retry.hidden=true;status.textContent='最新版を準備しています…';
 try{
  // 更新先へ接続できることを先に確認。通信断なら既存のオフライン画面を維持する。
  const stamp=Date.now();
  const reachable=await fresh(new URL('refresh.html?check='+stamp,location.href));
  if(!reachable.ok)throw new Error('OFFLINE');
  const scope=new URL('./',location.href).href;
  if('serviceWorker' in navigator){
   for(const registration of await navigator.serviceWorker.getRegistrations()){
    if(registration.scope===scope)await registration.unregister();
   }
  }
  // 他アプリのキャッシュ・IndexedDB・localStorage・sessionStorageには触れない。
  if('caches' in window){
   for(const key of await caches.keys())if(key.startsWith('beaufes-booth-shell-'))await caches.delete(key);
  }
  const response=await fresh(new URL('app.js?check='+stamp,location.href));
  if(!response.ok)throw new Error('UNAVAILABLE');
  const code=await response.text();const version=code.match(/const APP_VERSION=['"]([^'"]+)['"]/);if(!version)throw new Error('INVALID_APP');
  document.getElementById('update-version').textContent=version[1];
  status.textContent='最新版を取得しました。注文画面を開きます。';
  // 新しいURLのHTMLとscriptを使い、ブラウザー側の古いHTTPキャッシュも避ける。
  const target=new URL('index.html',location.href);target.searchParams.set('updated',stamp);
  if(new URLSearchParams(location.search).get('office')==='1')target.searchParams.set('office','1');
  if(join)target.hash='join='+join;
  location.replace(target.href);
 }catch{
  status.textContent='更新できませんでした。インターネット接続を確認し、もう一度更新してください。保存済みの注文は残っています。';retry.hidden=false;
 }
}
retry.addEventListener('click',refresh);void refresh();
