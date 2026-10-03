// 注文DBや接続情報は触らず、注文画面の保存資産だけを更新する。
const status=document.getElementById('update-status');const retry=document.getElementById('retry');
function activated(worker){return new Promise((resolve,reject)=>{
 if(!worker)return reject(new Error('NO_WORKER'));
 if(worker.state==='activated')return resolve();
 const timer=setTimeout(()=>finish(new Error('UPDATE_TIMEOUT')),30000);
 function finish(error){clearTimeout(timer);worker.removeEventListener('statechange',changed);error?reject(error):resolve();}
 function changed(){if(worker.state==='activated')finish();else if(worker.state==='redundant')finish(new Error('UPDATE_FAILED'));}
 worker.addEventListener('statechange',changed);
});}
async function refresh(){
 retry.hidden=true;status.textContent='最新版を準備しています…';
 try{
  if(!('serviceWorker' in navigator))throw new Error('UNSUPPORTED');
  // 旧画面が更新処理に進めなくても、この専用ページから新しいworkerを起動する。
  const registration=await navigator.serviceWorker.register('./sw.js?refresh='+Date.now(),{updateViaCache:'none'});
  const worker=registration.installing||registration.waiting||registration.active;await activated(worker);
  // clients.claim完了を待つ。保存済み注文や他アプリのキャッシュは削除しない。
  const limit=Date.now()+10000;
  while(navigator.serviceWorker.controller?.scriptURL!==worker.scriptURL){
   if(Date.now()>limit)throw new Error('CONTROL_TIMEOUT');
   await new Promise(resolve=>setTimeout(resolve,100));
  }
  const version=await new Promise((resolve,reject)=>{
   const channel=new MessageChannel();const timer=setTimeout(()=>reject(new Error('VERSION_TIMEOUT')),3000);
   channel.port1.onmessage=event=>{clearTimeout(timer);channel.port1.close();resolve(event.data);};
   worker.postMessage('SHELL_VERSION',[channel.port2]);
  });
  document.getElementById('update-version').textContent=version;
  status.textContent='更新しました。注文画面を開きます。';
  const target=new URL('./',location.href);
  if(new URLSearchParams(location.search).get('office')==='1')target.searchParams.set('office','1');
  target.searchParams.set('updated',Date.now());location.replace(target.href);
 }catch{
  status.textContent='更新できませんでした。インターネット接続を確認し、もう一度更新してください。保存済みの注文は残っています。';retry.hidden=false;
 }
}
retry.addEventListener('click',refresh);void refresh();
