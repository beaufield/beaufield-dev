/* 検証環境限定。実セッションの権限で架空データだけを書き込む。秘密・実データは表示しない。 */
(()=>{
const area=document.createElement('section');area.style='margin:16px;padding:16px;background:#e8f1ff';
area.innerHTML='<b>検証専用：保存・復元・切戻し</b><p>本番への接続・通知送信なし。架空データだけで試験します。</p><button id="dr-smoke">保存混在を1回確認</button><button id="dr-load">保存混在を3分確認</button><button id="dr-fence">保存停止・旧世代を確認</button><button id="dr-legacy">旧方式への切戻しを確認</button><button id="dr-restored">復元後の通常保存を確認</button><pre id="dr-proof"></pre>';
document.body.append(area);const proof=area.querySelector('#dr-proof');const c=window.SERIAL_TEST_CONFIG;
const tokens={},pending={};let epoch=null,lastPayload=null,lastId=null;
function assertTest(){if(c.url!=='https://qxzdypqbqduwuwmwnbgk.supabase.co'||!location.pathname.includes('/serial-apps-test/'))throw Error('TEST_ONLY');}
async function token(app){if(tokens[app]?.exp*1000>Date.now()+12000)return tokens[app].token;if(pending[app])return pending[app];pending[app]=(async()=>{const s=JSON.parse(localStorage.getItem('bf_session'));const r=await fetch(c.issuer,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'serialIssueDbToken',app,session_token:s?.token})});const d=await r.json();if(!d.success)throw Error('AUTH_'+app+'_'+d.error);tokens[app]=d;return d.token;})();try{return await pending[app];}finally{delete pending[app];}}
async function rpc(app,name,args={},schema='public'){
 const jwt=await token(app),start=performance.now();const ac=new AbortController(),timer=setTimeout(()=>ac.abort(),15000);
 try{const r=await fetch(c.url+'/rest/v1/rpc/'+name,{method:'POST',signal:ac.signal,headers:{apikey:c.publicKey,Authorization:'Bearer '+jwt,'Content-Type':'application/json','Content-Profile':schema,'Accept-Profile':schema},body:JSON.stringify(args)});const d=await r.json();return {status:r.status,ms:Math.round(performance.now()-start),data:d,success:r.ok&&d?.success!==false};}finally{clearTimeout(timer);}
}
async function serial(kind,serials,id=crypto.randomUUID()){
 const p={productCode:'DRTEST-20261107',serials,date:new Date().toLocaleDateString('sv-SE',{timeZone:'Asia/Tokyo'}).replaceAll('-','/'),customer:'検証専用',method:'単独',writer_epoch:epoch};
 lastPayload=p;lastId=id;return rpc('serial-apps','sr_mutate',{p_id:id,p_kind:kind,p_payload:p});
}
async function run(waves){
 assertTest();const runId=crypto.randomUUID().slice(0,8),start=Date.now(),results=[],orders=[],reservations=[];let snapshot=(await rpc('yoyaku-kanri','yk_get_snapshot',{},'yoyaku_api')).data.data;
 const ctl=await rpc('serial-apps','sr_control');epoch=ctl.data.epoch;if(ctl.data.writer!=='db')throw Error('WRITER_NOT_DB');
 const date=new Date().toLocaleDateString('sv-SE',{timeZone:'Asia/Tokyo'});
 for(let wave=0;wave<waves;wave++){
  const cycle=Math.floor(wave/3),kind=wave%3===1?'return':'ship',tasks=[];
  for(let i=0;i<11;i++){const names=Array.from({length:20},(_,j)=>'DR-'+runId+'-'+cycle+'-'+i+'-'+j);tasks.push(serial(kind,names).then(r=>({app:'serial',wave,kind,...r})));}
  for(let i=0;i<3;i++){
   const requestId='DR-'+crypto.randomUUID(),code='DR-'+runId+'-'+wave+'-'+i;
   tasks.push(rpc('order-app','save_order',{p:{requestId,date,supplierCode:'DRTEST',supplierName:'検証専用架空メーカー',staff:'検証専用',outputType:'',items:[{code,name:'検証専用架空商品',qty:1,unit:'個',memo:'DR-'+runId,isHandwritten:true}]}}).then(r=>{if(r.success)orders.push({requestId,orderNo:r.data.orderNo});return {app:'order',wave,...r};}));
   tasks.push(rpc('yoyaku-kanri','yk_mutate',{p_request_id:crypto.randomUUID(),p_epoch:snapshot.writer_epoch,p_action:'saveReservation',p_versions:{},p_payload:{product_id:'SRDR20261107',salon_name:'検証専用-'+runId+'-'+wave+'-'+i,quantity:1,staff_id:snapshot.user_info.user_id,delivery_method:'未定',notes:'DR-'+runId}},'yoyaku_api').then(r=>{if(r.success)reservations.push(r.data.data.reservation_no);return {app:'reservation',wave,...r};}));
  }
  tasks.push(rpc('serial-apps','sr_page',{p_filter:{maker:'検証専用'},p_csv:false}).then(r=>({app:'serial-read',wave,...r})));
  tasks.push(rpc('serial-apps','sr_page',{p_filter:{maker:'検証専用',yearMonth:date.slice(0,7)},p_csv:true}).then(r=>({app:'serial-csv',wave,...r})));
  tasks.push(rpc('order-app','get_orders',{p:{}}).then(r=>({app:'order-read',wave,...r})));
  tasks.push(rpc('yoyaku-kanri','yk_get_snapshot',{},'yoyaku_api').then(r=>({app:'reservation-read',wave,...r})));
  const settled=await Promise.allSettled(tasks);for(const x of settled){if(x.status==='rejected')results.push({wave,success:false,error:x.reason?.name});else{const r=x.value;results.push({app:r.app,wave,kind:r.kind,status:r.status,ms:r.ms,success:r.success,changed:r.data.changed,error:r.success?undefined:(r.data.error||r.data.message),rows:r.data.rows?.length});}}
  proof.textContent=JSON.stringify({running:true,runId,completed_waves:wave+1,waves,requests:results.length,failed:results.filter(x=>!x.success).length},null,2);
  if(results.some(x=>!x.success))break;
  if(wave+1<waves)await new Promise(resolve=>setTimeout(resolve,15000));
 }
 // 同一UUIDの同時再送と、異なるUUIDでの同じシリアル登録を分ける。
 const sameId=crypto.randomUUID(),sameName='DR-'+runId+'-SAME';const same=await Promise.all(Array.from({length:11},()=>serial('ship',[sameName],sameId)));
 const diffName='DR-'+runId+'-DUP';const diff=await Promise.all(Array.from({length:11},()=>serial('ship',[diffName])));
 const lostId=crypto.randomUUID(),lostName='DR-'+runId+'-ACK';await serial('ship',[lostName],lostId);const lookup=await rpc('serial-apps','sr_operation',{p_id:lostId});const replay=await serial('ship',[lostName],lostId);
 const summary={runId,waves,started_at:new Date(start).toISOString(),finished_at:new Date().toISOString(),duration_ms:Date.now()-start,results,orders,reservations,same_uuid:{calls:11,all_success:same.every(x=>x.success),same_result:same.every(x=>JSON.stringify(x.data)===JSON.stringify(same[0].data))},different_uuid:{calls:11,all_success:diff.every(x=>x.success),changed:diff.reduce((n,x)=>n+(x.data.changed||0),0),skipped:diff.reduce((n,x)=>n+(x.data.skipped||0),0)},discarded_ack:{lookup_success:lookup.success&&lookup.data?.changed===1,replay_same:replay.success&&JSON.stringify(lookup.data)===JSON.stringify(replay.data)}};
 sessionStorage.setItem('serial-drill-last',JSON.stringify({epoch,lastPayload,lastId,summary}));proof.textContent=JSON.stringify(summary,null,2);
}
async function fence(){assertTest();const saved=JSON.parse(sessionStorage.getItem('serial-drill-last')||'null');if(!saved)throw Error('NO_PRIOR_RUN');const r=await rpc('serial-apps','sr_mutate',{p_id:crypto.randomUUID(),p_kind:'ship',p_payload:{...saved.lastPayload,serials:['DR-FENCE-'+crypto.randomUUID().slice(0,8)]}});proof.textContent=JSON.stringify({scope:'old epoch write after closure/reopen',status:r.status,success:r.success,error:r.data.message||r.data.error},null,2);}
async function legacy(){assertTest();const s=JSON.parse(localStorage.getItem('bf_session'));const r=await fetch(c.issuer,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'serialRecoveryValidation',mode:'cutback',session_token:s?.token})});proof.textContent=JSON.stringify(await r.json(),null,2);}
async function freshRestore(){
 assertTest();const master=await serialAdapterCall({action:'getProductMaster'});if(!master.success||master.data.length!==58)throw Error('RESTORED_MASTER_COUNT');
 const product=master.data[0],name='DR-RESTORE-PROBE',date=new Date().toLocaleDateString('sv-SE',{timeZone:'Asia/Tokyo'}).replaceAll('-','/');
 const base={productCode:product.code,productName:product.name,jan:product.jan,maker:product.maker,series:product.series,size:product.size,serials:[name],customer:'検証専用_復元後',method:'単独'};
 const first=await serialAdapterCall({...base,action:'registerShipping',shipDate:date});
 const returned=await serialAdapterCall({...base,action:'registerReturn',kind:'返品',returnDate:date});
 const reship=await serialAdapterCall({...base,action:'registerShipping',shipDate:date});
 const cancelled=await serialAdapterCall({...base,action:'registerReturn',kind:'取消',returnDate:date});
 const read=await serialAdapterCall({action:'search',productCode:product.code,serialNo:name,statuses:[]});
 proof.textContent=JSON.stringify({success:first.registered===1&&returned.updated===1&&reship.registered===1&&cancelled.updated===1&&read.data.length===2,products:master.data.length,registered:first.registered,returned:returned.updated,reshipped:reship.registered,cancelled:cancelled.updated,readback_rows:read.data.length,active:read.data.filter(x=>x.status==='出荷中').length,scope:'normal serial adapter after external restore; synthetic serial only'},null,2);
}
for(const [id,fn] of [['#dr-smoke',()=>run(1)],['#dr-load',()=>run(12)],['#dr-fence',fence],['#dr-legacy',legacy],['#dr-restored',freshRestore]])area.querySelector(id).onclick=async()=>{for(const b of area.querySelectorAll('button'))b.disabled=true;proof.textContent='処理中です。再送せずお待ちください。';try{await fn();}catch(e){proof.textContent=JSON.stringify({success:false,error:e.message});}finally{for(const b of area.querySelectorAll('button'))b.disabled=false;}};
})();
