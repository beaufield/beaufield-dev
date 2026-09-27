import {get,set,records,update,mergeReceipt,sanitizeLegacyQr} from './storage.js';
import {buildWorkbook} from './export.js';
const APP_VERSION='v0.5.0';
const API_URL=location.hostname==='127.0.0.1'||location.hostname==='localhost'?'./api':'https://ovblkjxlbmnehgkpmyrd.supabase.co/functions/v1/booth-api';
const OFFICE_URL=API_URL==='./api'?API_URL:API_URL.replace('/booth-api','/office-api');
const officeMode=new URLSearchParams(location.search).get('office')==='1';
const localDemo=API_URL==='./api';
const PORTAL_URL='https://beaufield.github.io/beaufield-dev/';
function portalSession(){try{const s=JSON.parse(localStorage.getItem('bf_session')||'null');return s&&s.expires>Date.now()&&typeof s.token==='string'?s:null;}catch{return null;}}
const $=id=>document.getElementById(id);
const yen=n=>`${Number(n||0).toLocaleString('ja-JP')}円`;
let info, draft, pumping=false, camera, selectedSubject;
const deviceId=()=>{let id=localStorage.getItem('beaufes-booth-device');if(!id){id=crypto.randomUUID();localStorage.setItem('beaufes-booth-device',id);}return id;};
const scopeKey=()=>JSON.stringify(info.scope);
const message=text=>{$('message').textContent=text;};
const errorText=e=>({EVENT_CLOSED:'受付は締められています。未送信注文は事務に連絡してください。',QR_NOT_FOUND:'この名札を照合できません。受付情報の同期を確認してください。',SESSION_INVALID:'接続が期限切れです。再接続してから再送してください。',PORTAL_SESSION_INVALID:'社内ポータルへログインしてください。',OFFICE_ROLE_REQUIRED:'事務画面の管理者権限がありません。',PORTAL_UNAVAILABLE:'社内ポータルに接続できません。時間をおいて再試行してください。',INVALID_BUNDLE:'同期ファイルが不正です。3種類のデータを作り直してください。',STALE_MASTER_GENERATION:'得意先マスターより古い同期ファイルです。作り直してください。',OLD_GENERATION:'申込データより古い同期ファイルです。作り直してください。',IDEMPOTENCY_CONFLICT:'同じ注文番号で内容が異なります。再登録せず事務に連絡してください。',SET_QUANTITY_MISMATCH:'購入数とサービス数が条件に合っていません。',SCOPE_MISMATCH:'この端末の所属が変わっています。元の接続先で再送してください。'}[e.message]||`処理を完了できませんでした：${e.message}`);
function on(id,event,fn){$(id).addEventListener(event,async e=>{try{await fn(e);}catch(error){message(errorText(error));}});}
async function api(action,body={}){const token=officeMode&&!localDemo?portalSession()?.token:sessionStorage.getItem('booth-token');if(!token)throw Error(officeMode?'PORTAL_SESSION_INVALID':'SESSION_INVALID');const controller=new AbortController();const timer=setTimeout(()=>controller.abort(),action==='sync_bundle'?60000:12000);try{const response=await fetch(officeMode?OFFICE_URL:API_URL,{method:'POST',headers:{'Content-Type':'application/json',Authorization:`Bearer ${token}`},body:JSON.stringify({action,body}),signal:controller.signal,cache:'no-store'});const data=await response.json();if(!response.ok)throw Error(data.error||'SERVER_ERROR');return data;}finally{clearTimeout(timer);}}
function tokenFromQr(value){try{const u=new URL(value);if(!['https:','http:'].includes(u.protocol)||!u.pathname.endsWith('/pass.html'))throw Error();const values=u.searchParams.getAll('t');if(values.length!==1||!values[0]||values[0].length>256)throw Error();return values[0];}catch{throw Error('名札QRのURLを読み取ってください。');}}
async function digest(text){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(text)))].map(x=>x.toString(16).padStart(2,'0')).join('');}
async function enroll(key){const response=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'enroll',key,device_id:deviceId()}),cache:'no-store'});const data=await response.json();if(!response.ok)throw Error(data.error||'ENROLLMENT_INVALID');sessionStorage.setItem('booth-token',data.token);const catalog=await api('catalog');if(!catalog.catalog)throw Error('このメーカーの商品が未登録です。');await activate(catalog);message('メーカー用QRで接続しました。');await pump();}
function element(tag,text,className){const e=document.createElement(tag);if(text!==undefined)e.textContent=text;if(className)e.className=className;return e;}
async function fresh(){
 draft={id:crypto.randomUUID(),scope:info.scope,backend:location.origin,version:info.catalog.version,stage:'draft',form:{qr_hash:'',resolved:false,delivery:'later',offers:{}},createdAt:new Date().toISOString(),pending:null,cancelRequested:false};
 await update(draft.id,()=>draft);await set('draft:'+scopeKey(),draft.id);renderForm();
}
function renderForm(){
 $('draft-id').textContent=`注文番号 ${draft.id}`;$('qr').value='';$('delivery').value=draft.form.delivery;$('customer').textContent=draft.form.customer||'';
 $('offers').replaceChildren();
 for(const offer of info.catalog.offers){
  const section=element('div',undefined,'offer');section.append(element('h3',`${offer.label}　${offer.paid_per_set}＋${offer.gift_per_set}`));
  const row=element('label','セット数 ');const count=element('input');count.type='number';count.min='0';count.max='99';count.step='1';count.value=draft.form.offers[offer.offer_id]?.sets||0;count.dataset.offer=offer.offer_id;count.dataset.kind='sets';row.append(count);section.append(row);
  if(offer.product_codes.length>1){
   const head=element('div',undefined,'pick');head.append(element('small','商品'),element('small','購入数'),element('small','サービス数'));section.append(head);
   for(const code of offer.product_codes){const product=info.catalog.products.find(p=>p.product_code===code);const choice=element('div',undefined,'pick');choice.append(element('span',`${product.product_name}　${yen(product.event_sale_price_yen)}`));for(const kind of ['paid','gift']){const input=element('input');input.type='number';input.min='0';input.max='297';input.step='1';input.value=draft.form.offers[offer.offer_id]?.[kind]?.[code]||0;input.dataset.offer=offer.offer_id;input.dataset.kind=kind;input.dataset.code=code;input.setAttribute('aria-label',`${product.product_name} ${kind==='paid'?'購入':'サービス'}数`);choice.append(input);}section.append(choice);}
  }else{const p=info.catalog.products.find(p=>p.product_code===offer.product_codes[0]);section.append(element('p',`商品コード ${p.product_code} ／ 単価 ${yen(p.event_sale_price_yen)} ／ サービスは同じ商品です。`,'muted'));}
  $('offers').append(section);
 }
 $('offers').querySelectorAll('input').forEach(input=>input.addEventListener('change',()=>saveForm().catch(e=>message(errorText(e)))));
 const locked=draft.stage!=='draft';for(const x of [$('qr'),$('delivery'),$('submit'),$('scan'),$('resolve'),$('demo'),...$('offers').querySelectorAll('input')])x.disabled=locked;
 calculate();
}
async function readForm(){const offers={};for(const i of $('offers').querySelectorAll('input')){const o=offers[i.dataset.offer]||={sets:0,paid:{},gift:{}};if(i.dataset.kind==='sets')o.sets=Number(i.value);else o[i.dataset.kind][i.dataset.code]=Number(i.value);}const {qr:oldRawQr,...oldForm}=draft.form;const raw=$('qr').value.trim();$('qr').value='';const qr_hash=raw?await digest(tokenFromQr(raw)):oldForm.qr_hash;return {...oldForm,qr_hash,customer:raw?'':oldForm.customer,resolved:raw?false:oldForm.resolved,delivery:$('delivery').value,offers};}
async function saveForm(){if(draft.stage!=='draft')return;const form=await readForm();draft=await update(draft.id,r=>({...r,form}));$('customer').textContent=form.customer||'';calculate();}
function bundles(check=false){return info.catalog.offers.flatMap(o=>{const selected=draft.form.offers[o.offer_id];if(!selected?.sets){if(check&&selected&&[...Object.values(selected.paid),...Object.values(selected.gift)].some(x=>x!==0))throw Error(`${o.label}のセット数を指定してください。`);return [];}
 const n=selected.sets;if(check&&(!Number.isInteger(n)||n<1||n>99))throw Error('セット数は1〜99の整数で指定してください。');
 const value={offer_id:o.offer_id,sets:n};for(const k of ['paid','gift']){value[k]=o.product_codes.length===1?[{code:o.product_codes[0],qty:n*o[`${k}_per_set`]}]:Object.entries(selected[k]).filter(([,qty])=>qty!==0).map(([code,qty])=>({code,qty}));if(check&&(value[k].some(x=>!Number.isInteger(x.qty)||x.qty<1)||value[k].reduce((s,x)=>s+x.qty,0)!==n*o[`${k}_per_set`]))throw Error(`${o.label}：${k==='paid'?'購入':'サービス'}数は${n*o[`${k}_per_set`]}点にしてください。`);}return [value];});}
function calculate(){const amount=bundles().reduce((sum,b)=>sum+b.paid.reduce((s,x)=>s+x.qty*info.catalog.products.find(p=>p.product_code===x.code).event_sale_price_yen,0),0);$('total').textContent=yen(amount);}
async function activate(value){info=value;$('login').hidden=true;if(!officeMode)await set('maker-catalog',info);$('maker').textContent=info.catalog?.manufacturer_name||'メーカー未設定';$('workspace').hidden=info.role==='admin';$('admin').hidden=info.role!=='admin';$('ledger').hidden=info.role!=='admin';$('identity').hidden=info.role!=='admin';if(info.role==='admin')return;const id=await get('draft:'+scopeKey());draft=(await records()).find(x=>x.id===id);if(!draft)await fresh();else renderForm();await renderOrders();}
async function renderOrders(){
 if(!info)return;$('orders').replaceChildren();
 for(const r of (await records()).filter(r=>r.backend===location.origin&&JSON.stringify(r.scope)===scopeKey()&&(r.stage!=='draft'||r.id!==draft?.id)).sort((a,b)=>b.createdAt.localeCompare(a.createdAt))){
  const node=element('div',undefined,'order');
  const state=r.stage==='draft'?'下書き':r.pending==='cancel'?'取消待ち':r.pending?'未送信・結果確認中':r.receipt?.state==='cancelled'?'取消済み':'登録済み';
  node.append(element('strong',state),element('p',r.form.customer||'名札から照合'),element('p',`${r.receipt?.total_yen==null?'金額未確定':yen(r.receipt.total_yen)} ／ ${r.id}`,'muted'));
  if(r.error)node.append(element('p',r.error,'error'));
  if(r.stage==='draft'){
   const resume=element('button','この下書きの続きを入力','secondary');resume.onclick=async()=>{try{await saveForm();draft=(await records()).find(x=>x.id===r.id);await set('draft:'+scopeKey(),draft.id);renderForm();await renderOrders();}catch(e){message(errorText(e));}};node.append(resume);
  }else if(r.receipt?.state!=='cancelled'){
   const cancel=element('button','この注文を取り消す','secondary');cancel.disabled=r.cancelRequested;cancel.onclick=async()=>{if(!confirm('この注文を取り消しますか？'))return;try{await update(r.id,current=>({...current,cancelRequested:true,pending:'cancel'}));await renderOrders();await pump();}catch(e){message(errorText(e));}};node.append(cancel);
  }
  $('orders').append(node);
 }
}
async function pumpInner(){if(!info||!sessionStorage.getItem('booth-token')||!navigator.onLine)return;for(const entry of await records()){if(entry.backend!==location.origin||JSON.stringify(entry.scope)!==scopeKey()||entry.stage==='draft')continue;const latest=(await records()).find(x=>x.id===entry.id);try{const response=await api(latest.pending==='cancel'?'cancel':latest.pending==='submit'?'submit':'status',latest.pending==='submit'?latest.body:{request_id:latest.id});await update(latest.id,current=>({...mergeReceipt(current,response),error:null}));}catch(e){await update(latest.id,current=>({...current,error:errorText(e)}));}}await renderOrders();}
async function pump(){if(pumping)return;pumping=true;try{if(navigator.locks)await navigator.locks.request('beaufes-booth-send',pumpInner);else await pumpInner();}finally{pumping=false;}}
on('connect','click',async()=>{if(!officeMode||localDemo)sessionStorage.setItem('booth-token',$('token').value);const value=await api('catalog');if(!value.catalog)throw Error('このメーカーの商品が未登録です。');await activate(value);$('token').value='';message(officeMode?'事務画面に接続しました。':'接続できました。注文の選択を始められます。');if(!officeMode)await pump();});
on('delivery','change',saveForm);
on('resolve','click',async()=>{await saveForm();if(!draft.form.qr_hash)throw Error('名札QRを読み取ってください。');const customer=await api('resolve',{token_hash:draft.form.qr_hash});const label=`${customer.salon} ／ ${customer.name}${customer.customer_linked?'':'（得意先コードは事務確認）'}`;draft=await update(draft.id,r=>({...r,form:{...r.form,customer:label,resolved:true}}));$('customer').textContent=label;});
on('demo','click',async()=>{if(draft.stage!=='draft')throw Error('「次のお客様」から新しい注文を開いてください。');$('qr').value='https://example.invalid/pass.html?t=demo-ticket';await saveForm();$('resolve').click();});
on('new','click',async()=>{if(draft.stage==='draft'&&bundles().length&&!confirm('現在の未登録の選択を残して、新しい注文を開きますか？'))return;await saveForm();await fresh();await renderOrders();message('新しい注文を開きました。');});
on('submit','click',async()=>{if(draft.stage!=='draft')return;await saveForm();if(!draft.form.resolved||!draft.form.qr_hash)throw Error('先に名札QRでお客様を確認してください。');const b=bundles(true);if(!b.length)throw Error('商品を1セット以上選択してください。');const body={request_id:draft.id,scope:info.scope,catalog_version:info.catalog.version,token_hash:draft.form.qr_hash,delivery:draft.form.delivery,bundles:b};
 // 下書きIDを維持したまま注文と送信待ちを同じトランザクションで保存する。
 draft=await update(draft.id,r=>{if(r.stage!=='draft')return r;return {...r,body,stage:'queued',pending:'submit'};});renderForm();message('この端末に注文を保存しました。送信結果を確認します。');await renderOrders();await pump();});
on('retry','click',pump);
on('scan','click',async()=>{if(camera){camera.getTracks().forEach(t=>t.stop());camera=null;$('video').hidden=true;return;}camera=await navigator.mediaDevices.getUserMedia({video:{facingMode:'environment'}});$('video').srcObject=camera;$('video').hidden=false;await $('video').play();const canvas=$('canvas'),ctx=canvas.getContext('2d',{willReadFrequently:true});const step=()=>{if(!camera)return;if($('video').readyState>=2){canvas.width=$('video').videoWidth;canvas.height=$('video').videoHeight;ctx.drawImage($('video'),0,0);const frame=ctx.getImageData(0,0,canvas.width,canvas.height);const code=window.jsQR(frame.data,frame.width,frame.height);if(code){try{tokenFromQr(code.data);$('qr').value=code.data;camera.getTracks().forEach(t=>t.stop());camera=null;$('video').hidden=true;saveForm().then(()=>$('resolve').click()).catch(e=>message(errorText(e)));return;}catch{}}}requestAnimationFrame(step);};requestAnimationFrame(step);});
on('close','click',async()=>{if(!confirm('試験受付を締めます。未送信の注文がないか確認しましたか？'))return;await api('close');message('受付を締めました。未接続端末の注文は別途確認してください。');});
on('admin-refresh','click',refreshAdmin);
async function refreshAdmin(){const data=await api('admin_orders');$('admin-orders').replaceChildren();for(const order of data.rows){const node=element('div',undefined,'order');node.append(element('strong',`${order.customer.salon||'未対応'} ／ ${order.customer.name||''} ／ ${yen(order.total_yen)}`),element('p',`得意先 ${order.customer.customer_code||'要確認'} ／ ${order.state==='cancelled'?'取消':order.posting.state} ／ ${order.request_id}`,'muted'));
 const actionButton=(label,state)=>{const button=element('button',label,'secondary');button.onclick=async()=>{try{let slip_number;if(state==='entered'){slip_number=prompt('販売管理に入力した伝票番号を入力してください。');if(!slip_number)return;}await api('posting',{request_id:order.request_id,booth:order.booth,revision:order.posting.revision,state,slip_number});await refreshAdmin();}catch(e){message(errorText(e));}};node.append(button);};
 if(order.state==='active'&&order.customer.customer_code&&order.posting.state==='ready')actionButton('この注文の入力を担当する','claimed');
 if(order.posting.state==='claimed'&&order.posting.owner===info.scope.device){actionButton('入力済みにする','entered');actionButton('入力結果が不明','unknown');}
 $('admin-orders').append(node);}}
async function reviewSubject(row,code,state){
 const label=state==='confirmed'?code+' に確定':'保留';
 if(!confirm(row.salon+' ／ '+row.name+' を「'+label+'」にしますか？'))return;
 await api('review',{subject_id:row.id,customer_code:code,state,revision:row.decision?.revision||0});
 await renderIdentity();await refreshAdmin();message('得意先コードの確認状態を更新しました。');
}
async function renderIdentity(){
 const data=await api('list');$('identity-list').replaceChildren();
 if(!data.rows.length){$('identity-list').append(element('p','未確定・再確認のお客様はいません。'));return;}
 for(const row of data.rows){
  const node=element('div',undefined,'order');
  node.append(element('strong',row.salon+' ／ '+row.name),element('p','申込番号 '+row.id+' ／ '+(row.decision?.state==='needs_review'?'以前の確定と新しい情報が競合':'得意先コード未確定'),'muted'));
  for(const candidate of row.candidates){
   const button=element('button',candidate.code+'　'+candidate.name+'（'+candidate.evidence.join('・')+'）','secondary');
   button.onclick=()=>reviewSubject(row,candidate.code,'confirmed').catch(e=>message(errorText(e)));node.append(button);
  }
  const search=element('button','このお客様のコードを検索','secondary');search.onclick=()=>{selectedSubject=row;$('identity-selected').textContent=row.salon+' ／ '+row.name+' の得意先を検索中';$('identity-query').focus();};node.append(search);
  const hold=element('button','保留にする','secondary');hold.onclick=()=>reviewSubject(row,null,'hold').catch(e=>message(errorText(e)));node.append(hold);
  $('identity-list').append(node);
 }
}
on('identity-refresh','click',renderIdentity);
// 事務担当者が生成済みの完全世代だけを一括反映する。失敗時はDB全体をロールバックする。
on('identity-sync','click',async()=>{
 const file=$('identity-file').files[0];if(!file)throw Error('同期ファイル bundle.json を選んでください。');
 if(file.size>480000)throw Error('同期ファイルが大きすぎます。');
 const body=JSON.parse(await file.text());
 if(!body||Object.keys(body).sort().join(',')!=='candidates,identity,master'||body.master?.count!==body.master?.rows?.length||body.identity?.subject_count!==body.identity?.subjects?.length||body.identity?.link_count!==body.identity?.links?.length||body.candidates?.count!==body.candidates?.rows?.length)throw Error('INVALID_BUNDLE');
 if(!confirm(`得意先${body.master.count}件、申込${body.identity.subject_count}件、名札${body.identity.link_count}件、候補${body.candidates.count}件を一括同期しますか？`))return;
 const result=await api('sync_bundle',body);$('identity-file').value='';
 $('sync-status').textContent=`同期完了：得意先${result.master_count}件・申込${result.subject_count}件・名札${result.link_count}件・候補${result.candidate_count}件`;
 await renderIdentity();await refreshAdmin();
});
on('identity-search','click',async()=>{
 if(!selectedSubject)throw Error('先にお客様を選択してください。');
 const query=$('identity-query').value.trim();const data=await api('search_master',{query});
 $('identity-results').replaceChildren();
 for(const candidate of data.rows){
  const button=element('button',candidate.code+' ／ '+candidate.name,'secondary');
  button.onclick=()=>reviewSubject(selectedSubject,candidate.code,'confirmed').catch(e=>message(errorText(e)));
  $('identity-results').append(button);
 }
 if(!data.rows.length)$('identity-results').append(element('p','該当する得意先はありません。'));
});
on('export','click',async()=>{const batch_id=crypto.randomUUID();const data=await api('export',{batch_id});await set('last-export',data);const workbook=await buildWorkbook(window.ExcelJS,data);const blob=new Blob([await workbook.xlsx.writeBuffer()],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});const url=URL.createObjectURL(blob);const a=element('a');a.href=url;a.download=`ビューフェス注文_${batch_id}.xlsx`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);$('export-status').textContent=`${data.rows.length}件をExcelに出力しました。「入力候補」と「要確認」を分け、コードの先頭0を保持しています。`;});
async function start(){
 $('version').textContent=APP_VERSION;if(!localDemo){$('demo').hidden=true;if(!officeMode){$('token-label').hidden=true;$('connect').hidden=true;$('login-title').textContent='メーカー用QRを読み取ってください';}}
 const network=()=>{$('network').textContent=navigator.onLine?'オンライン':'オフライン';};network();window.addEventListener('online',()=>{network();pump().catch(e=>message(errorText(e)));});window.addEventListener('offline',network);
 await sanitizeLegacyQr();
 const join=new URL(location.href).hash.match(/^#join=([-_A-Za-z0-9]{43,128})$/)?.[1];
 if(location.hash)history.replaceState(null,'',location.pathname+location.search);
 if(officeMode){$('login-title').textContent='事務画面';$('token-label').hidden=!localDemo;$('connect').textContent='社内ポータルのログインで接続';$('demo').hidden=true;if(!localDemo&&!portalSession()){const link=element('a','社内ポータルでログインする');link.href=PORTAL_URL;$('login').append(link);message('社内ポータルへログインしてから、この画面に戻ってください。');}else if(!localDemo){const value=await api('catalog');await activate(value);message('事務画面に接続しました。');}}else{const saved=await get('maker-catalog');if(saved){await activate(saved);message('端末に保存した商品情報を読み込みました。再送にはメーカーQRで接続してください。');}if(join){await enroll(join);}}
 if('serviceWorker'in navigator){await navigator.serviceWorker.register('./sw.js');await navigator.serviceWorker.ready;$('prepared').textContent='アプリ本体のオフライン準備ができました。商品情報は接続時に保存されます。';}
}
start().catch(e=>message(errorText(e)));
