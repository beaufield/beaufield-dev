import {get,set,records,update,mergeReceipt,sanitizeLegacyQr} from './storage.js';
import {buildWorkbook} from './export.js';
const APP_VERSION='v0.10.0';
const API_URL=location.hostname==='127.0.0.1'||location.hostname==='localhost'?'./api':'https://ovblkjxlbmnehgkpmyrd.supabase.co/functions/v1/booth-api';
const OFFICE_URL=API_URL==='./api'?API_URL:API_URL.replace('/booth-api','/office-api');
const officeMode=new URLSearchParams(location.search).get('office')==='1';
const localDemo=API_URL==='./api';
const PORTAL_URL='https://beaufield.github.io/beaufield-dev/';
function portalSession(){try{const s=JSON.parse(localStorage.getItem('bf_session')||'null');return s&&s.expires>Date.now()&&typeof s.token==='string'?s:null;}catch{return null;}}
const $=id=>document.getElementById(id);
const yen=n=>`${Number(n||0).toLocaleString('ja-JP')}円`;
let info, draft, pumping=false, camera, selectedSubject, pickerState, scanId=0, modalReturnFocus;
let makerToken, customerRequest=0, searchRequest=0, historyRequest=0, sharedRows=[], historyCursor=null, historyScope, historyFetched, connectionVerified=false;
const historyDetails=new Map();
const deviceId=()=>{let id=localStorage.getItem('beaufes-booth-device');if(!id){id=crypto.randomUUID();localStorage.setItem('beaufes-booth-device',id);}return id;};
const scopeKey=()=>JSON.stringify(info.scope);
const message=text=>{$('message').textContent=text;};
const errorText=e=>({EVENT_CLOSED:'受付は締められています。未送信注文は事務に連絡してください。',QR_NOT_FOUND:'この名札を照合できません。受付情報の同期を確認してください。',SESSION_INVALID:'メーカー接続が無効です。メーカーQRを読み直してください。未送信注文は端末に残っています。',SUBJECT_NOT_FOUND:'このお客様の最新情報を確認できません。もう一度検索するか受付に確認してください。',INVALID_CUSTOMER_CODE:'得意先コードを確認してください。',ORDER_NOT_FOUND:'この注文を確認できません。履歴を更新してください。',PORTAL_SESSION_INVALID:'社内ポータルへログインしてください。',OFFICE_ROLE_REQUIRED:'事務画面の管理者権限がありません。',PORTAL_UNAVAILABLE:'社内ポータルに接続できません。時間をおいて再試行してください。',INVALID_BUNDLE:'同期ファイルが不正です。3種類のデータを作り直してください。',STALE_MASTER_GENERATION:'得意先マスターより古い同期ファイルです。作り直してください。',OLD_GENERATION:'申込データより古い同期ファイルです。作り直してください。',IDEMPOTENCY_CONFLICT:'同じ注文番号で内容が異なります。再登録せず事務に連絡してください。',SET_QUANTITY_MISMATCH:'購入数とサービス数が条件に合っていません。',SCOPE_MISMATCH:'この端末の所属が変わっています。元の接続先で再送してください。'}[e.message]||`処理を完了できませんでした：${e.message}`);
function on(id,event,fn){$(id).addEventListener(event,async e=>{try{await fn(e);}catch(error){const detail=errorText(error);message(detail);if(id==='submit'){const status=$('submit-status');status.hidden=false;status.className='error';status.textContent=detail;}}});}
async function api(action,body={}){const token=officeMode&&!localDemo?portalSession()?.token:(officeMode||localDemo?sessionStorage.getItem('booth-token'):makerToken);if(!token)throw Error(officeMode?'PORTAL_SESSION_INVALID':'SESSION_INVALID');const controller=new AbortController();const timer=setTimeout(()=>controller.abort(),action==='sync_bundle'?60000:12000);try{const response=await fetch(officeMode?OFFICE_URL:API_URL,{method:'POST',headers:{'Content-Type':'application/json',Authorization:`Bearer ${token}`},body:JSON.stringify({action,body}),signal:controller.signal,cache:'no-store'});const data=await response.json();if(!response.ok){if(data.error==='SESSION_INVALID'&&!officeMode&&!localDemo&&makerToken===token){makerToken=null;connectionVerified=false;sessionStorage.removeItem('booth-token');await set('maker-connection',null);showConnection();}throw Error(data.error||'SERVER_ERROR');}return data;}finally{clearTimeout(timer);}}
function tokenFromQr(value){try{const u=new URL(value);if(!['https:','http:'].includes(u.protocol)||!u.pathname.endsWith('/pass.html'))throw Error();const values=u.searchParams.getAll('t');if(values.length!==1||!values[0]||values[0].length>256)throw Error();return values[0];}catch{throw Error('名札QRのURLを読み取ってください。');}}
async function digest(text){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(text)))].map(x=>x.toString(16).padStart(2,'0')).join('');}
async function enroll(key){const response=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'enroll',key,device_id:deviceId()}),cache:'no-store'});const data=await response.json();if(!response.ok)throw Error(data.error||'ENROLLMENT_INVALID');makerToken=data.token;sessionStorage.setItem('booth-token',data.token);const catalog=await api('catalog');if(!catalog.catalog)throw Error('このメーカーの商品が未登録です。');if(catalog.scope.device!==deviceId())throw Error('SCOPE_MISMATCH');connectionVerified=true;await set('maker-connection',{token:data.token,device:deviceId()});await activate(catalog);message('メーカー用QRで接続しました。');await pump();}
function element(tag,text,className){const e=document.createElement(tag);if(text!==undefined)e.textContent=text;if(className)e.className=className;return e;}
async function fresh(){
 resetCustomerSearch();
 draft={id:crypto.randomUUID(),scope:info.scope,backend:location.origin,version:info.catalog.version,stage:'draft',form:{qr_hash:'',resolved:false,delivery:'',delivery_confirmed:false,ui_version:2,offers:{}},createdAt:new Date().toISOString(),pending:null,cancelRequested:false};
 await update(draft.id,()=>draft);await set('draft:'+scopeKey(),draft.id);renderForm();
}
function numberInput(value,limit,description){
 const input=element('input');input.type='number';input.inputMode='numeric';input.min='0';input.max=String(limit);input.step='1';input.value=value>0?String(value):'';input.placeholder='0';input.setAttribute('aria-label',description);
 // 入力開始時に既存値を選択し、先頭の0が残る誤入力を防ぐ。
 input.addEventListener('focus',()=>input.select());
 input.addEventListener('blur',()=>{if(input.value!==''){const n=Number(input.value);if(Number.isInteger(n)&&n>=0)input.value=n?String(n):'';}});
 return input;
}
function selectedOffer(offer){return draft.form.offers[offer.offer_id];}
function addSelectedProduct(group,offer,kind,code,quantity){
 const product=info.catalog.products.find(p=>p.product_code===code);if(!product)return;
 const row=element('div',undefined,'selected-product');const name=element('div',product.product_name,'product-name');
 name.append(element('small',kind==='paid'?yen(product.event_sale_price_yen):'サービス品・0円'));
 const input=numberInput(quantity,297,`${product.product_name} ${kind==='paid'?'購入':'サービス'}数`);
 input.dataset.offer=offer.offer_id;input.dataset.kind=kind;input.dataset.code=code;
 input.addEventListener('input',()=>updateQuantityGuidance(offer,group.closest('.offer-card')));
 input.addEventListener('change',()=>saveForm().catch(e=>message(errorText(e))));
 row.append(name,input);group.append(row);
}
function renderOffer(offer){
 const selected=selectedOffer(offer);const card=element('div',undefined,'offer-card');card.dataset.offerCard=offer.offer_id;
 const head=element('div',undefined,'offer-head');const title=element('div');title.append(element('h3',offer.label),element('p',`${offer.paid_per_set}点購入 ＋ ${offer.gift_per_set}点サービス`,'muted'));
 const remove=element('button','条件を外す','secondary small-action');remove.type='button';remove.onclick=()=>removeOffer(offer.offer_id).catch(e=>message(errorText(e)));head.append(title,remove);card.append(head);
 const countRow=element('div',undefined,'set-row');countRow.append(element('span','セット数','set-label'));
 const count=numberInput(selected.sets,99,`${offer.label}のセット数`);count.dataset.offer=offer.offer_id;count.dataset.kind='sets';
 countRow.append(count);
 const chips=element('div',undefined,'quick-counts');
 const syncChips=()=>chips.querySelectorAll('button').forEach(button=>button.setAttribute('aria-pressed',String(Number(count.value)>0&&Number(button.textContent)===Number(count.value))));
 count.addEventListener('input',()=>{syncChips();updateQuantityGuidance(offer,card);});
 count.addEventListener('change',()=>{syncChips();saveForm().catch(e=>message(errorText(e)));});
 for(const n of [1,2,3,4,5,6,10]){const chip=element('button',String(n),'secondary');chip.type='button';chip.setAttribute('aria-label',`${offer.label}を${n}セット`);chip.setAttribute('aria-pressed',String(Number(count.value)===n));chip.onclick=async()=>{count.value=String(n);syncChips();await saveForm();};chips.append(chip);}
 countRow.append(chips);card.append(countRow);
 if(offer.product_codes.length===1){
  const p=info.catalog.products.find(item=>item.product_code===offer.product_codes[0]);
  if(p)card.append(element('p',`${p.product_name} ／ 税抜単価 ${yen(p.event_sale_price_yen)}。サービス品は同じ商品です。`,'muted'));
 }else{
  for(const kind of ['paid','gift']){
   const group=element('div',undefined,'product-group');group.append(element('h4',kind==='paid'?'購入する商品':'サービス品（0円）'));
   const progress=element('p',undefined,'quantity-progress');progress.dataset.progressKind=kind;progress.setAttribute('role','status');group.append(progress);
   const choices=selected[kind]||{};
   // 複数商品を選べる条件は最初から全商品を並べ、未入力は0点として扱う。
   for(const code of offer.product_codes)addSelectedProduct(group,offer,kind,code,choices[code]||0);
   card.append(group);
  }
  updateQuantityGuidance(offer,card);
 }
 return card;
}
function renderForm(){
 $('draft-id').textContent=`注文番号 ${draft.id}`;$('customer').textContent=draft.form.customer||'';showConnection();
 for(const radio of document.querySelectorAll('input[name=delivery]'))radio.checked=(draft.stage!=='draft'||!!draft.form.delivery_confirmed)&&radio.value===draft.form.delivery;
 $('offers').replaceChildren();
 for(const offer of info.catalog.offers)if(selectedOffer(offer))$('offers').append(renderOffer(offer));
 if(!$('offers').children.length)$('offers').append(element('p','まだ条件を選んでいません。「特売条件を追加」から始めてください。','empty-state'));
 const locked=draft.stage!=='draft';for(const x of [$('submit'),$('scan'),$('demo'),$('customer-query'),$('customer-search'),...document.querySelectorAll('input[name=customer-search-mode]'),...$('customer-results').querySelectorAll('button'),$('add-offer'),...document.querySelectorAll('input[name=delivery]'),...$('offers').querySelectorAll('input,button')])x.disabled=locked;
 const status=$('submit-status');status.hidden=locked===false;status.className='';
 if(locked){if(draft.receipt?.state==='active'&&!draft.pending){$('submit').textContent='✓ 登録済み';status.className='success';status.textContent='登録が完了しました。次のお客様の入力欄へ切り替えています。';}
  else if(draft.receipt?.state==='cancelled'){ $('submit').textContent='取消済み';status.className='error';status.textContent='この注文は取消済みです。';}
  else{$('submit').textContent='送信・確認中…';status.className='pending';status.textContent=draft.error?`未送信・結果確認中：${draft.error}「未送信を再送・状態確認」を押してください。`:'注文を端末に保存しました。登録結果を確認中です。';}}
 else $('submit').textContent='この内容で登録する';
 calculate();
}
async function readForm(){const offers={};for(const i of $('offers').querySelectorAll('input[data-offer]')){const o=offers[i.dataset.offer]||={sets:0,paid:{},gift:{}};if(i.dataset.kind==='sets')o.sets=i.value===''?0:Number(i.value);else o[i.dataset.kind][i.dataset.code]=i.value===''?0:Number(i.value);}const {qr:oldRawQr,...oldForm}=draft.form;const delivery=document.querySelector('input[name=delivery]:checked')?.value||'';return {...oldForm,delivery,delivery_confirmed:!!delivery,ui_version:2,offers};}
function updateQuantityGuidance(offer,card){
 if(!card)return;
 const setsInput=card.querySelector('input[data-kind=sets]');const sets=Number(setsInput?.value||0);
 for(const kind of ['paid','gift']){
  const progress=card.querySelector(`[data-progress-kind=${kind}]`);if(!progress)continue;
  const inputs=[...card.querySelectorAll(`input[data-kind=${kind}]`)];
  const quantities=inputs.map(input=>Number(input.value||0));
  const valid=inputs.every((input,i)=>input.value===''||(Number.isInteger(quantities[i])&&quantities[i]>=0));
  const total=quantities.reduce((sum,n)=>sum+n,0);const perSet=offer[`${kind}_per_set`];const label=kind==='paid'?'購入':'サービス';
  progress.className='quantity-progress';
  if(!valid||!Number.isInteger(sets)||sets<0||sets>99){progress.classList.add('quantity-error');progress.textContent='整数の数量を入力してください。';continue;}
  if(!sets){progress.classList.add(total?'quantity-error':'quantity-neutral');progress.textContent=total?`${label} ${total}点入力済み。先にセット数を選んでください。`:'セット数を選ぶと必要な数量が表示されます。';continue;}
  const required=sets*perSet;const difference=required-total;
  const base=`${label}：${perSet}点 × ${sets}セット ＝ 必要${required}点 ／ 入力済み${total}点`;
  if(difference>0){progress.classList.add('quantity-error');progress.textContent=`${base}　あと${difference}点`;}
  else if(difference<0){progress.classList.add('quantity-error');progress.textContent=`${base}　${-difference}点多いです`+(kind==='paid'&&perSet===3&&total%3!==0?'（購入合計は3の倍数にしてください）':'');}
  else{progress.classList.add('quantity-complete');progress.textContent=`${base}　✓ 数量一致`;}
 }
}
function updateSelectionSummaries(){for(const offer of info.catalog.offers){if(!draft.form.offers[offer.offer_id]||offer.product_codes.length===1)continue;const card=Array.from($('offers').querySelectorAll('.offer-card')).find(node=>node.dataset.offerCard===offer.offer_id);updateQuantityGuidance(offer,card);}}
async function saveForm(){if(draft.stage!=='draft')return;const id=draft.id;const form=await readForm();const current=await update(id,r=>r.stage==='draft'?{...r,form:{...r.form,offers:form.offers,delivery:form.delivery,delivery_confirmed:form.delivery_confirmed,ui_version:2}}:r);if(draft.id!==id)return;draft=current;$('customer').textContent=draft.form.customer||'';calculate();updateSelectionSummaries();}
async function removeOffer(id){await saveForm();const current=draft.form.offers[id];if(current&&(current.sets>0||Object.values(current.paid||{}).some(Boolean)||Object.values(current.gift||{}).some(Boolean))&&!confirm('選択した条件と数量を外しますか？'))return;draft=await update(draft.id,r=>{const offers={...r.form.offers};delete offers[id];return {...r,form:{...r.form,offers}};});renderForm();}
function openPicker(state){if(draft.stage!=='draft')return;pickerState=state;modalReturnFocus=document.activeElement;$('picker-title').textContent='特売条件を追加';$('picker-search').value='';$('picker-modal').hidden=false;document.body.classList.add('modal-open');renderPicker();$('picker-search').focus();}
function closePicker(){$('picker-modal').hidden=true;pickerState=null;document.body.classList.remove('modal-open');modalReturnFocus?.focus();}
function renderPicker(){
 const state=pickerState;if(!state)return;const query=$('picker-search').value.trim().toLocaleLowerCase('ja-JP');let choices;
 choices=info.catalog.offers.filter(o=>!draft.form.offers[o.offer_id]).map(o=>({key:o.offer_id,name:o.label,sub:`${o.paid_per_set}点購入 ＋ ${o.gift_per_set}点サービス`}));
 const matches=choices.filter(c=>c.name.toLocaleLowerCase('ja-JP').includes(query));$('picker-results').replaceChildren();
 for(const choice of matches.slice(0,20)){const button=element('button',choice.name,'picker-result');button.type='button';button.append(element('small',choice.sub));button.onclick=()=>choosePicker(choice.key).catch(e=>message(errorText(e)));$('picker-results').append(button);}
 $('picker-count').textContent=matches.length>20?`${matches.length}件中20件を表示。名前で絞り込んでください。`:`${matches.length}件`;
 if(!matches.length)$('picker-results').append(element('p','選べる候補がありません。','muted'));
}
async function choosePicker(key){if(!pickerState)return;await saveForm();draft=await update(draft.id,r=>{const offers={...r.form.offers};offers[key]={sets:0,paid:{},gift:{}};return {...r,form:{...r.form,offers}};});closePicker();renderForm();const target=Array.from($('offers').querySelectorAll('.offer-card')).find(node=>node.dataset.offerCard===key)?.querySelector('input[data-kind=sets]');target?.focus();}
function bundles(check=false){return info.catalog.offers.flatMap(o=>{const selected=draft.form.offers[o.offer_id];if(!selected?.sets){if(check&&selected&&[...Object.values(selected.paid),...Object.values(selected.gift)].some(x=>x!==0))throw Error(`${o.label}のセット数を指定してください。`);return [];}
 const n=selected.sets;if(check&&(!Number.isInteger(n)||n<1||n>99))throw Error('セット数は1〜99の整数で指定してください。');
 const value={offer_id:o.offer_id,sets:n};for(const k of ['paid','gift']){value[k]=o.product_codes.length===1?[{code:o.product_codes[0],qty:n*o[`${k}_per_set`]}]:Object.entries(selected[k]).filter(([,qty])=>qty!==0).map(([code,qty])=>({code,qty}));if(check&&(value[k].some(x=>!Number.isInteger(x.qty)||x.qty<1)||value[k].reduce((s,x)=>s+x.qty,0)!==n*o[`${k}_per_set`]))throw Error(`${o.label}：${k==='paid'?'購入':'サービス'}数は${n*o[`${k}_per_set`]}点にしてください。`);}return [value];});}
function calculate(){const amount=bundles().reduce((sum,b)=>sum+b.paid.reduce((s,x)=>s+x.qty*info.catalog.products.find(p=>p.product_code===x.code).event_sale_price_yen,0),0);$('total').textContent=yen(amount);}
async function activate(value){info=value;$('login').hidden=true;if(!officeMode)await set('maker-catalog',info);$('maker').textContent=info.catalog?.manufacturer_name||'メーカー未設定';$('maker-banner').hidden=info.role==='admin';$('workspace').hidden=info.role==='admin';$('admin').hidden=info.role!=='admin';$('ledger').hidden=info.role!=='admin';$('identity').hidden=info.role!=='admin';if(info.role==='admin')return;const id=await get('draft:'+scopeKey());draft=(await records()).find(x=>x.id===id);if(!draft)await fresh();else{if(draft.stage==='draft'&&draft.form.ui_version!==2){draft=await update(draft.id,r=>{const offers={};for(const [key,item] of Object.entries(r.form.offers||{})){const paid=Object.fromEntries(Object.entries(item.paid||{}).filter(([,qty])=>qty>0));const gift=Object.fromEntries(Object.entries(item.gift||{}).filter(([,qty])=>qty>0));if(item.sets>0||Object.keys(paid).length||Object.keys(gift).length)offers[key]={sets:item.sets||0,paid,gift};}return {...r,form:{...r.form,offers,delivery:'',delivery_confirmed:false,ui_version:2}};});}if(draft.stage!=='draft'&&draft.receipt?.state==='active'&&!draft.pending)await fresh();else renderForm();}await renderOrders();}
function orderDetails(record){
 // 登録確定後はサーバーの確定明細を使用し、通信待ちは端末に保存した注文内容と明示する。
 let lines=Array.isArray(record.receipt?.lines)?record.receipt.lines:null;
 const confirmed=!!lines;
 if(!lines&&record.body?.catalog_version===info.catalog.version){
  lines=(record.body.bundles||[]).flatMap(bundle=>['paid','gift'].flatMap(kind=>(bundle[kind]||[]).map(item=>{
   const product=info.catalog.products.find(p=>p.product_code===item.code);
   return product?{kind,product_name:product.product_name,qty:item.qty,unit_price_yen:kind==='gift'?0:product.event_sale_price_yen}:null;
  }).filter(Boolean)));
 }
 const details=element('div',undefined,'order-details');
 const paid=(lines||[]).filter(line=>line.kind==='paid').reduce((sum,line)=>sum+Number(line.qty||0),0);
 const gift=(lines||[]).filter(line=>line.kind==='gift').reduce((sum,line)=>sum+Number(line.qty||0),0);
 details.append(element('h4',lines?.length?`商品明細（購入${paid}点・サービス${gift}点）`:'商品明細'));
 if(!lines?.length){details.append(element('p','商品明細を取得できません。事務担当者に確認してください。','error'));return details;}
 if(!confirmed)details.append(element('p','端末に保存した送信前の内容です。登録確定後に再確認してください。','muted'));
 for(const kind of ['paid','gift']){
  const list=lines.filter(line=>line.kind===kind);
  if(!list.length)continue;
  details.append(element('h4',kind==='paid'?'購入商品':'サービス品（0円）'));
  for(const line of list){const qty=Number(line.qty||0);const price=Number(line.unit_price_yen||0);details.append(element('p',`${line.product_name}　${qty}点 × ${yen(price)}${kind==='paid'?` ＝ ${yen(qty*price)}`:''}`,'order-line'));}
 }
 return details;
}
async function renderOrders(){
 if(!info)return;
 // 再送や取消で一覧が更新されても、確認中の行は開いた状態を維持する。
 const opened=new Set(Array.from($('orders').querySelectorAll('details[open]')).map(node=>node.dataset.orderId));
 $('orders').replaceChildren();
 const list=(await records()).filter(r=>r.backend===location.origin&&JSON.stringify(r.scope)===scopeKey()&&(r.stage==='draft'?r.id!==draft?.id&&(r.form.resolved||Object.values(r.form.offers||{}).some(o=>o.sets>0)):!!r.pending||!r.receipt||!!r.error)).sort((a,b)=>b.createdAt.localeCompare(a.createdAt));
 $('order-count').textContent=`${list.length}件`;
 if(!list.length)$('orders').append(element('p','未送信・下書きはありません。','muted'));
 for(const r of list){
  const node=element('details',undefined,'order');node.dataset.orderId=r.id;node.open=opened.has(r.id);
  const state=r.stage==='draft'?'下書き':r.pending==='cancel'?'取消待ち':r.pending?'未送信':r.receipt?.state==='cancelled'?'取消済み':'登録済み';
  const statusClass=r.pending?'order-pending':r.receipt?.state==='cancelled'?'order-cancelled':r.stage==='draft'?'order-draft':'order-complete';
  const summary=element('summary',undefined,'order-summary');
  summary.append(element('span',state+(r.error?'・要確認':''),`order-state ${statusClass}`));
  const date=new Date(r.createdAt);const time=element('time',Number.isNaN(date.getTime())?'':date.toLocaleString('ja-JP',{month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit'}),'order-time');if(!Number.isNaN(date.getTime()))time.dateTime=r.createdAt;
  summary.append(time,element('strong',r.receipt?.total_yen==null?'金額未確定':yen(r.receipt.total_yen),'order-amount'));
  const customer=element('span',r.form.customer||'名札から照合','order-customer');summary.append(customer);node.append(summary);
  const body=element('div',undefined,'order-body');
  body.append(element('p',r.form.customer||'名札から照合','order-customer-full'));
  const delivery=r.body?.delivery||r.form.delivery;body.append(element('p',`受け渡し：${delivery==='later'?'後日納品':delivery==='takeaway'?'当日持ち帰り':'未選択'}`,'order-delivery'));
  body.append(element('p',`注文番号 ${r.id}`,'muted order-reference'));
  if(r.stage!=='draft')body.append(orderDetails(r));
  if(r.error)body.append(element('p',r.error,'error'));
  if(r.stage==='draft'){
   const resume=element('button','この下書きの続きを入力','secondary');resume.onclick=async()=>{try{await saveForm();draft=(await records()).find(x=>x.id===r.id);await set('draft:'+scopeKey(),draft.id);resetCustomerSearch();renderForm();await renderOrders();$('customer-section').scrollIntoView({block:'start'});}catch(e){message(errorText(e));}};body.append(resume);
  }else if(r.receipt?.state!=='cancelled'){
   const cancel=element('button','この注文を取り消す','secondary');cancel.disabled=r.cancelRequested;cancel.onclick=async()=>{if(!confirm('この注文を取り消しますか？'))return;try{await update(r.id,current=>({...current,cancelRequested:true,pending:'cancel'}));await renderOrders();await pump();}catch(e){message(errorText(e));}};body.append(cancel);
  }
  node.append(body);$('orders').append(node);
 }
}
async function pumpInner(){if(!info||!(localDemo?sessionStorage.getItem('booth-token'):makerToken)||!navigator.onLine)return;for(const entry of await records()){if(entry.backend!==location.origin||JSON.stringify(entry.scope)!==scopeKey()||entry.stage==='draft')continue;const latest=(await records()).find(x=>x.id===entry.id);if(!latest.pending)continue;try{const response=await api(latest.pending==='cancel'?'cancel':latest.pending==='submit'?'submit':'status',latest.pending==='submit'?latest.body:{request_id:latest.id});await update(latest.id,current=>({...mergeReceipt(current,response),error:null}));}catch(e){await update(latest.id,current=>({...current,error:errorText(e)}));}}if(draft&&draft.stage!=='draft'){const current=(await records()).find(r=>r.id===draft.id);if(current){draft=current;if(current.receipt?.state==='active'&&!current.pending){
   // サーバー登録が確定した時だけ新しい入力欄を開く。前の注文は端末の注文一覧に残す。
   await fresh();message('登録完了。次のお客様の名札QRを読み取ってください。前の注文は下の一覧から確認・取消できます。');$('message').scrollIntoView({block:'start'});$('scan').focus({preventScroll:true});
  }else renderForm();}}await renderOrders();await refreshHistory();}
async function pump(){if(pumping)return;pumping=true;try{if(navigator.locks)await navigator.locks.request('beaufes-booth-send',pumpInner);else await pumpInner();}finally{pumping=false;}}

function showConnection(){
 if(officeMode)return;
 const connected=localDemo?!!sessionStorage.getItem('booth-token'):!!makerToken;
 $('connection-status').textContent=!navigator.onLine?'オフライン：未送信注文はこの端末に保存されます。検索・共有履歴には通信が必要です。':!connected?'メーカーQRを読み直して接続してください。':connectionVerified?'メーカー接続済み':'メーカー接続を確認しています。';
}
function resetCustomerSearch(){
 customerRequest++;searchRequest++;
 $('customer-query').value='';$('customer-results').replaceChildren();$('customer-search-status').textContent='';$('customer-search-panel').open=false;$('customer-recent').textContent='';
}
function openCustomerSearch(){$('customer-search-panel').open=true;$('customer-query').focus();}
function searchMode(){return document.querySelector('input[name=customer-search-mode]:checked').value;}
function searchChanged(){
 customerRequest++;searchRequest++;$('customer-results').replaceChildren();$('customer-search-status').textContent='';
 const code=searchMode()==='code';$('customer-query-label').textContent=code?'名札の得意先コード':'お客様名・サロン名';$('customer-query').placeholder=code?'名札に印字されたお客様番号':'2文字以上で検索';$('customer-query').inputMode=code?'text':'search';
}
function customerLabel(customer){return `${customer.salon} ／ ${customer.name}${customer.customer_code?`（得意先 ${customer.customer_code}）`:customer.customer_linked?'':'（得意先コードは事務確認）'}`;}
async function chooseCustomer(customer,selection,id,serial){
 if(draft.id!==id||draft.stage!=='draft'||customerRequest!==serial)return;
 const label=customerLabel(customer);
 const current=await update(id,r=>draft.id===id&&r.stage==='draft'&&customerRequest===serial?{...r,form:{...r.form,...selection,customer:label,customer_id:customer.id,resolved:true}}:r);
 if(draft.id!==id||customerRequest!==serial)return;
 draft=current;$('customer').textContent=label;$('customer-search-panel').open=false;
 message('お客様を選びました。名札のお名前と照合してから商品を登録してください。');
 await customerRecent(customer.id,id,serial);
}
async function customerRecent(subjectId,id,serial){
 $('customer-recent').textContent='このお客様の直近の注文を確認しています…';
 try{const data=await api('manufacturer_history',{subject_id:subjectId});if(draft.id!==id||customerRequest!==serial)return;
  const active=data.rows.filter(row=>row.state==='active');
  $('customer-recent').textContent=active.length?`このお客様は、このメーカーで最近${active.length}件${data.next_cursor?'以上':''}の注文があります。重複していないか、下の「メーカーの注文履歴」で確認してください。`:'このお客様の直近の登録済み注文はありません。';
 }catch(e){if(draft.id===id&&customerRequest===serial)$('customer-recent').textContent='直近の注文を確認できませんでした。通信状態と注文の重複を確認してください。';}
}
async function readBadge(value){
 if(draft.stage!=='draft')throw Error('次のお客様から新しい注文を開いてください。');
 const serial=++customerRequest;searchRequest++;const id=draft.id;
 await saveForm();const qr_hash=await digest(tokenFromQr(value));if(draft.id!==id||customerRequest!==serial)return;
 const current=await update(id,r=>draft.id===id&&customerRequest===serial?{...r,form:{...r.form,qr_hash,subject_id:null,customer_id:null,customer:'',resolved:false}}:r);
 if(draft.id!==id||customerRequest!==serial)return;draft=current;$('customer').textContent='';$('customer-results').replaceChildren();$('customer-recent').textContent='';
 try{const customer=await api('resolve',{token_hash:qr_hash});await chooseCustomer(customer,{qr_hash,subject_id:null},id,serial);}
 catch(e){if(draft.id!==id||customerRequest!==serial)return;openCustomerSearch();throw e;}
}
async function searchCustomers(event){
 event.preventDefault();if(draft.stage!=='draft')return;customerRequest++;
 const query=$('customer-query').value.normalize('NFKC').trim(),mode=searchMode();const id=draft.id,serial=++searchRequest;
 $('customer-results').replaceChildren();
 if(!query||(mode==='name'&&query.length<2)){$('customer-search-status').textContent=mode==='name'?'名前・サロン名を2文字以上入力してください。':'名札の得意先コードを入力してください。';return;}
 $('customer-search-status').textContent='検索しています…';
 try{const data=await api('search_customers',{query,mode});if(draft.id!==id||draft.stage!=='draft'||searchRequest!==serial)return;
  $('customer-search-status').textContent=data.rows.length?`${data.rows.length}件${data.more?'以上（20件まで表示）':''}。名札のお名前と照合して選んでください。`:'該当する申込者が見つかりません。コード・名前を確認し、見つからない場合は受付へ確認してください。';
  for(const row of data.rows){const card=element('div',undefined,'customer-result');card.append(element('strong',row.salon+' ／ '+row.name),element('p',`申込番号 ${row.id} ／ 得意先 ${row.customer_code||'未確定'}`));
   const button=element('button','このお客様を選ぶ');button.type='button';button.onclick=async()=>{
    if(draft.id!==id||draft.stage!=='draft'||searchRequest!==serial)return;
    const choice=++customerRequest;button.disabled=true;
    try{await saveForm();if(draft.id!==id||customerRequest!==choice||searchRequest!==serial)return;const cleared=await update(id,r=>draft.id===id&&customerRequest===choice?{...r,form:{...r.form,subject_id:null,qr_hash:'',customer_id:null,customer:'',resolved:false}}:r);if(draft.id!==id||customerRequest!==choice||searchRequest!==serial)return;draft=cleared;$('customer').textContent='';$('customer-recent').textContent='';const customer=await api('select_customer',{subject_id:row.id});if(searchRequest!==serial)return;await chooseCustomer(customer,{subject_id:customer.id,qr_hash:''},id,choice);}
    catch(e){if(draft.id===id&&customerRequest===choice&&searchRequest===serial)message(errorText(e));}
    finally{button.disabled=draft.stage!=='draft';}
   };card.append(button);$('customer-results').append(card);
  }
 }catch(e){if(draft.id===id&&searchRequest===serial){$('customer-search-status').textContent=errorText(e);}}
}
on('customer-query','input',searchChanged);
document.querySelectorAll('input[name=customer-search-mode]').forEach(radio=>radio.addEventListener('change',searchChanged));
on('customer-search-form','submit',searchCustomers);
function historySummary(row){
 const summary=element('summary',undefined,'order-summary');summary.append(element('span',row.state==='cancelled'?'取消済み':'登録済み',`order-state ${row.state==='cancelled'?'order-cancelled':'order-complete'}`));
 const date=new Date(row.received_at);const time=element('time',date.toLocaleString('ja-JP',{month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit'}),'order-time');time.dateTime=row.received_at;
 summary.append(time,element('strong',yen(row.total_yen),'order-amount'),element('span',`${row.customer?.salon||'要確認'} ／ ${row.customer?.name||''}`,'order-customer'));return summary;
}
async function cancelLocalOrder(record){
 if(!confirm('この注文を取り消しますか？'))return;
 await update(record.id,current=>({...current,cancelRequested:true,pending:'cancel'}));await renderOrders();await pump();
}
async function fillSharedDetail(row,body,scope){
 try{
  const key=row.request_id;let data=historyDetails.get(key);
  if(!data||data.state_seq!==row.state_seq){body.textContent='明細を読み込んでいます…';data=await api('manufacturer_order',{request_id:key});if(scope!==scopeKey())return;historyDetails.set(key,data);}
  if(scope!==scopeKey())return;body.replaceChildren();
  body.append(element('p',customerLabel(data.customer),'order-customer-full'),element('p',`受け渡し：${data.delivery==='later'?'後日納品':'当日持ち帰り'}`,'order-delivery'),element('p',`注文番号 ${key}`,'muted order-reference'));
  body.append(orderDetails({receipt:data}));
  if(data.state==='cancelled'){body.append(element('p','この注文は取消済みです。','muted'));return;}
  const local=(await records()).find(r=>r.id===key&&r.backend===location.origin&&JSON.stringify(r.scope)===scope);
  if(data.mine&&local){const cancel=element('button',local.cancelRequested?'取消を確認中':'この注文を取り消す','secondary');cancel.disabled=!!local.cancelRequested;cancel.onclick=()=>cancelLocalOrder(local).catch(e=>message(errorText(e)));body.append(cancel);}
  else body.append(element('p',data.mine?'この端末に元の注文データがないため閲覧のみです。':'他の端末の注文です（閲覧のみ）。','muted'));
 }catch(e){if(scope===scopeKey()){body.replaceChildren(element('p',errorText(e),'error'));const retry=element('button','明細を再取得','secondary');retry.onclick=()=>fillSharedDetail(row,body,scope);body.append(retry);}}
}
function renderSharedHistory(){
 const opened=new Set([...$('shared-orders').querySelectorAll('details[open]')].map(node=>node.dataset.orderId));$('shared-orders').replaceChildren();
 const scope=scopeKey();
 for(const row of sharedRows){const node=element('details',undefined,'order');node.dataset.orderId=row.request_id;node.append(historySummary(row));const body=element('div',undefined,'order-body');node.append(body);let loaded=false;
  node.addEventListener('toggle',()=>{if(node.open&&!loaded){loaded=true;fillSharedDetail(row,body,scope);}});node.open=opened.has(row.request_id);$('shared-orders').append(node);
  if(node.open){loaded=true;fillSharedDetail(row,body,scope);}
 }
 if(!sharedRows.length)$('shared-orders').append(element('p','登録済みの注文はありません。','muted'));
 $('history-more').hidden=!historyCursor;
}
async function refreshHistory(more=false){
 if(!info||info.role!=='manufacturer')return;
 const scope=scopeKey();
 if(historyScope!==scope){historyRequest++;historyScope=scope;sharedRows=[];historyCursor=null;historyFetched=null;historyDetails.clear();$('shared-orders').replaceChildren();$('history-more').hidden=true;}
 if(!(localDemo?sessionStorage.getItem('booth-token'):makerToken)||!navigator.onLine){$('history-status').textContent=historyFetched?`前回取得 ${historyFetched}。通信・メーカー接続を確認して更新してください。`:'履歴の取得には通信とメーカーQRでの接続が必要です。';return;}
 const serial=++historyRequest;$('history-status').classList.remove('history-error');$('history-status').textContent='履歴を取得しています…';$('history-refresh').disabled=true;$('history-more').disabled=true;
 try{const data=await api('manufacturer_history',more&&historyCursor?{cursor:historyCursor}:{});if(serial!==historyRequest||scope!==scopeKey())return;
  sharedRows=more?[...new Map([...sharedRows,...data.rows].map(row=>[row.request_id,row])).values()]:data.rows;historyCursor=data.next_cursor;historyFetched=new Date(data.fetched_at).toLocaleString('ja-JP');
  renderSharedHistory();$('history-status').textContent=`取得 ${historyFetched} ／ ${sharedRows.length}件表示${historyCursor?'（続きあり）':''}`;
 }catch(e){if(serial===historyRequest&&scope===scopeKey()){$('history-status').classList.add('history-error');$('history-status').textContent=`取得できませんでした。${historyFetched?`前回取得 ${historyFetched}の表示です。`:''}${errorText(e)}`;}}
 finally{if(serial===historyRequest){$('history-refresh').disabled=false;$('history-more').disabled=false;}}
}
on('history-refresh','click',()=>refreshHistory());on('history-more','click',()=>refreshHistory(true));

on('connect','click',async()=>{if(!officeMode||localDemo){makerToken=$('token').value;sessionStorage.setItem('booth-token',makerToken);}const value=await api('catalog');if(!value.catalog)throw Error('このメーカーの商品が未登録です。');connectionVerified=true;await activate(value);$('token').value='';message(officeMode?'事務画面に接続しました。':'接続できました。注文の選択を始められます。');if(!officeMode)await pump();});
document.querySelectorAll('input[name=delivery]').forEach(radio=>radio.addEventListener('change',()=>saveForm().catch(e=>message(errorText(e)))));
on('add-offer','click',()=>openPicker({mode:'offer'}));
on('picker-search','input',renderPicker);
on('picker-close','click',closePicker);
on('demo','click',()=>readBadge('https://example.invalid/pass.html?t=demo-ticket'));
on('new','click',async()=>{if(draft.stage==='draft'&&bundles().length&&!confirm('現在の未登録の選択を残して、新しい注文を開きますか？'))return;closeCamera();await saveForm();await fresh();await renderOrders();message('新しい注文を開きました。');});
on('submit','click',async()=>{if(draft.stage!=='draft')return;await saveForm();if(!draft.form.resolved||!(draft.form.qr_hash||draft.form.subject_id))throw Error('先に名札QRまたはコード・名前検索でお客様を選んでください。');const b=bundles(true);if(!b.length)throw Error('商品を1セット以上選択してください。');if(!draft.form.delivery_confirmed||!['later','takeaway'].includes(draft.form.delivery))throw Error('受け渡し方法を選んでください。');const body={request_id:draft.id,scope:info.scope,catalog_version:info.catalog.version,...(draft.form.subject_id?{subject_id:draft.form.subject_id}:{token_hash:draft.form.qr_hash}),delivery:draft.form.delivery,bundles:b};
 // 下書きIDを維持したまま注文と送信待ちを同じトランザクションで保存する。
 draft=await update(draft.id,r=>{if(r.stage!=='draft')return r;return {...r,body,stage:'queued',pending:'submit'};});renderForm();message('この端末に注文を保存しました。送信結果を確認します。');await renderOrders();await pump();});
on('retry','click',pump);
function closeCamera(){scanId++;if(camera){camera.getTracks().forEach(t=>t.stop());camera=null;}$('video').pause();$('video').srcObject=null;$('camera-modal').hidden=true;document.body.classList.remove('modal-open');modalReturnFocus?.focus();}
on('camera-close','click',closeCamera);
on('scan','click',async()=>{
 if(camera){closeCamera();return;}
 if(!navigator.mediaDevices?.getUserMedia){openCustomerSearch();throw Error('この端末ではカメラを使えません。得意先コード・名前で検索してください。');}
 modalReturnFocus=document.activeElement;$('camera-modal').hidden=false;document.body.classList.add('modal-open');$('camera-close').focus();const id=++scanId;
 try{const stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:'environment'}});if(id!==scanId){stream.getTracks().forEach(t=>t.stop());return;}camera=stream;$('video').srcObject=stream;await $('video').play();}
 catch(e){closeCamera();openCustomerSearch();throw Error('カメラを開けませんでした。許可を確認するか、得意先コード・名前で検索してください。');}
 const canvas=$('canvas'),ctx=canvas.getContext('2d',{willReadFrequently:true});
 const step=()=>{if(!camera||id!==scanId)return;if($('video').readyState>=2){canvas.width=$('video').videoWidth;canvas.height=$('video').videoHeight;ctx.drawImage($('video'),0,0);const frame=ctx.getImageData(0,0,canvas.width,canvas.height);const code=window.jsQR(frame.data,frame.width,frame.height);if(code){try{tokenFromQr(code.data);closeCamera();readBadge(code.data).catch(e=>message(errorText(e)));return;}catch{}}}requestAnimationFrame(step);};requestAnimationFrame(step);
});
document.addEventListener('visibilitychange',()=>{if(document.hidden&&camera)closeCamera();else if(!document.hidden)refreshHistory().catch(e=>message(errorText(e)));});
document.addEventListener('keydown',e=>{if(e.key==='Escape'){if(!$('camera-modal').hidden)closeCamera();else if(!$('picker-modal').hidden)closePicker();}});
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
 // 固定ヘッダーの実寸に合わせ、スクロール先がメーカー名の下へ隠れないようにする。
 const header=$('app-header');const fitHeader=()=>document.documentElement.style.setProperty('--header-height',`${header.getBoundingClientRect().height+12}px`);new ResizeObserver(fitHeader).observe(header);fitHeader();
 $('version').textContent=APP_VERSION;if(!localDemo){$('demo').hidden=true;if(!officeMode){$('token-label').hidden=true;$('connect').hidden=true;$('login-title').textContent='メーカー用QRを読み取ってください';}}
 const network=()=>{$('network').textContent=navigator.onLine?'オンライン':'オフライン';};network();window.addEventListener('online',()=>{network();showConnection();pump().catch(e=>message(errorText(e)));});window.addEventListener('offline',()=>{network();showConnection();});
 await sanitizeLegacyQr();
 const join=new URL(location.href).hash.match(/^#join=([-_A-Za-z0-9]{43,128})$/)?.[1];
 if(location.hash)history.replaceState(null,'',location.pathname+location.search);
 if(officeMode){$('login-title').textContent='事務画面';$('token-label').hidden=!localDemo;$('connect').textContent='社内ポータルのログインで接続';$('demo').hidden=true;if(!localDemo&&!portalSession()){const link=element('a','社内ポータルでログインする');link.href=PORTAL_URL;$('login').append(link);message('社内ポータルへログインしてから、この画面に戻ってください。');}else if(!localDemo){const value=await api('catalog');await activate(value);message('事務画面に接続しました。');}}else{const remembered=await get('maker-connection');makerToken=sessionStorage.getItem('booth-token')||(remembered?.device===deviceId()?remembered.token:null);if(makerToken)sessionStorage.setItem('booth-token',makerToken);const saved=await get('maker-catalog');if(saved){await activate(saved);message('端末の商品情報を表示しています。接続を確認します。');}if(join){await enroll(join);}else if(makerToken){try{const value=await api('catalog');if(!localDemo&&value.scope.device!==deviceId())throw Error('SCOPE_MISMATCH');connectionVerified=true;await activate(value);if(!localDemo)await set('maker-connection',{token:makerToken,device:deviceId()});message('メーカー接続を復元しました。');await pump();}catch(e){showConnection();message(errorText(e));}}else if(saved){showConnection();message('メーカーQRを読み直すと、お客様検索・送信・共有履歴が使えます。未送信注文は端末に残っています。');}}
 if('serviceWorker'in navigator){await navigator.serviceWorker.register('./sw.js');await navigator.serviceWorker.ready;$('prepared').textContent='アプリ本体のオフライン準備ができました。商品情報は接続時に保存されます。';}
}
start().catch(e=>message(errorText(e)));
