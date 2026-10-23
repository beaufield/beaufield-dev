import {get,set,records,update as storageUpdate,mergeReceipt,sanitizeLegacyQr,clearMakerConnection,storageMode} from './storage.js';
import {buildWorkbook,displayProductCode,emptyHandling,packRemarks,unpackRemarks,handlingWarnings,deliveryLabel,billingLabel} from './export.js';
const APP_VERSION='v0.21.0';
const API_URL=location.hostname==='127.0.0.1'||location.hostname==='localhost'?'./api':'https://ovblkjxlbmnehgkpmyrd.supabase.co/functions/v1/booth-api';
const OFFICE_URL=API_URL==='./api'?API_URL:API_URL.replace('/booth-api','/office-api');
const officeMode=new URLSearchParams(location.search).get('office')==='1';
const localDemo=API_URL==='./api';
const PORTAL_URL='https://beaufield.github.io/beaufield-dev/';
function portalSession(){try{const s=JSON.parse(localStorage.getItem('bf_session')||'null');return s&&s.expires>Date.now()&&typeof s.token==='string'?s:null;}catch{return null;}}
const $=id=>document.getElementById(id);
const yen=n=>`${Number(n||0).toLocaleString('ja-JP')}円`;
let submitting=false;
let info, draft, pumping=false, camera, selectedSubject, pickerState, scanId=0, modalReturnFocus;
let makerToken, customerRequest=0, searchRequest=0, historyRequest=0, sharedRows=[], historyCursor=null, historyScope, historyFetched, connectionVerified=false;
const historyDetails=new Map();
// 同じ下書きを別タブが変更したら、古い入力を保存・登録せず最新表示へ戻す。
async function update(id,change){
 const expected=id===draft?.id?(draft.edit_revision||0):null;
 try{return await storageUpdate(id,current=>{
  if(current?.stage==='draft'&&expected!==null&&(current.edit_revision||0)!==expected)throw Error('DRAFT_CHANGED');
  return change(current);
 });}catch(error){
  if(error.message==='DRAFT_CHANGED'&&draft?.id===id){const latest=(await records()).find(r=>r.id===id);if(latest){draft=latest;resetCustomerSearch();renderForm();}}
  throw error;
 }
}
// 保存禁止時は、この画面の間だけ端末番号と接続を保持する。
const temporary=new Map();
const session={getItem:key=>{try{return sessionStorage.getItem(key)||temporary.get(key)||null;}catch{return temporary.get(key)||null;}},setItem:(key,value)=>{temporary.set(key,value);try{sessionStorage.setItem(key,value);}catch{}},removeItem:key=>{temporary.delete(key);try{sessionStorage.removeItem(key);}catch{}}};
let memoryDevice;
const deviceId=()=>{if(memoryDevice)return memoryDevice;try{memoryDevice=localStorage.getItem('beaufes-booth-device');if(!memoryDevice){memoryDevice=crypto.randomUUID();localStorage.setItem('beaufes-booth-device',memoryDevice);}}catch{memoryDevice=crypto.randomUUID();}return memoryDevice;};
const scopeKey=()=>JSON.stringify(info.scope);
const message=text=>{$('message').textContent=text;};
// 接続処理の仕組みではなく、利用者が次にできる操作を案内する。
const readyMessage=value=>value.event_state==='closed'?'注文の受付は終了しました。履歴を確認できます。':'注文を入力できます。お客様を選んでください。';
const errorText=e=>({PAPER_ALREADY_REGISTERED:'同じ紙からの注文が登録済みです。メーカーの注文履歴でお客様と明細を確認し、番号を変えて再登録しないでください。',INVALID_PAPER_NUMBER:'紙受付Noは半角英数字・ハイフン32文字以内で入力してください。',INVALID_HANDLING:'受付情報を確認してください。',RESERVED_REMARKS:'受付情報は補助欄へ入力してください。',RESERVATION_DELIVERY_REQUIRED:'予約商品は後日納品です。お渡し済みの商品とは注文を分けてください。',ORDER_CHANGED_RECHECK_REQUIRED:'注文または台帳が変わりました。最新の控えと販売管理を照合し直してください。',NOT_UNRESOLVED:'この注文の確認待ちは既に解決しています。台帳を更新してください。',RESOLUTION_CONFLICT:'この照合操作は別の内容または担当者で記録されています。操作番号を変えて再送せず、事務に確認してください。',OFFICE_RESOLUTION_STORAGE_REQUIRED:'照合操作を端末に保存できません。ブラウザーの保存を許可した端末で操作してください。',SEARCH_INCOMPLETE:'販売管理の照合が終わっていません。確認待ちを維持します。',PRIOR_CORRECTION_NOT_RESOLVED:'古い誤入力・重複の解消と訂正記録を確認してください。',DRAFT_CHANGED:'別のタブで入力が変更されました。最新の氏名・商品・数量を確認してから操作してください。',PURCHASE_CONTENT_CONFLICT:'このお客様の登録確認中の注文と内容が違います。再登録せず、メーカー履歴を確認してください。',STALE_PURCHASE:'別の画面で購入が進んでいます。履歴を更新して確認してください。',ORDER_INTENT_EXPIRED:'確認待ちの注文番号の期限が切れました。再登録せず事務へ確認してください。',CUSTOMER_CHANGED:'申込情報が変わりました。再登録せず受付へ確認してください。',INVALID_REMARKS:'備考は1,000文字以内の文章で入力してください。',OFFER_PERSON_LIMIT:'この企画はお一人様1セットです。登録済みの注文を確認してください。',OFFER_EXCLUSIVE:'この2つの企画は併用できません。どちらかを選んでください。',FIXED_COMPONENT_REQUIRED:'固定セットの構成を確認してください。',LINKED_SKU_REQUIRED:'同じ種類のサービス品と数量を確認してください。',EVENT_CLOSED:'受付は締められています。未送信注文は事務に連絡してください。',QR_NOT_FOUND:'この名札を照合できません。受付情報の同期を確認してください。',SESSION_INVALID:'メーカー接続が無効です。メーカーQRを読み直してください。未送信注文は端末に残っています。',SUBJECT_NOT_FOUND:'このお客様の最新情報を確認できません。もう一度検索するか受付に確認してください。',INVALID_CUSTOMER_CODE:'得意先コードを確認してください。',ORDER_NOT_FOUND:'この注文を確認できません。履歴を更新してください。',PORTAL_SESSION_INVALID:'社内ポータルへログインしてください。',OFFICE_ROLE_REQUIRED:'事務画面の管理者権限がありません。',PORTAL_UNAVAILABLE:'社内ポータルに接続できません。時間をおいて再試行してください。',INVALID_BUNDLE:'同期ファイルが不正です。3種類のデータを作り直してください。',STALE_MASTER_GENERATION:'得意先マスターより古い同期ファイルです。作り直してください。',OLD_GENERATION:'申込データより古い同期ファイルです。作り直してください。',IDEMPOTENCY_CONFLICT:'同じ注文番号で内容が異なります。再登録せず事務に連絡してください。',SET_QUANTITY_MISMATCH:'購入数とサービス数が条件に合っていません。',SCOPE_MISMATCH:'この端末の所属が変わっています。元の接続先で再送してください。'}[e.message]||`処理を完了できませんでした：${e.message}`);
function on(id,event,fn){$(id).addEventListener(event,async e=>{try{await fn(e);}catch(error){const detail=errorText(error);message(detail);if(id==='submit'){const status=$('submit-status');status.hidden=false;status.className='error';status.textContent=detail;}}});}
async function requestApi(action,body,token){
 const controller=new AbortController();const timer=setTimeout(()=>controller.abort(),action==='sync_bundle'?60000:12000);
 try{const response=await fetch(officeMode?OFFICE_URL:API_URL,{method:'POST',headers:{'Content-Type':'application/json',Authorization:`Bearer ${token}`},body:JSON.stringify({action,body}),signal:controller.signal,cache:'no-store'});return {ok:response.ok,data:await response.json()};}finally{clearTimeout(timer);}
}
async function recoverMakerConnection(failedToken){
 // 再接続済みの同じ端末・同じメーカーに限り、サーバーで有効性を確認して復元する。
 const saved=await get('maker-connection');if(!saved?.token||saved.device!==deviceId()||saved.token===failedToken)return false;
 const checked=await requestApi('catalog',{},saved.token);if(!checked.ok||checked.data.role!=='manufacturer'||!checked.data.catalog)return false;
 if(checked.data.scope?.device!==deviceId()||info&&['event','booth','device','epoch'].some(key=>checked.data.scope[key]!==info.scope[key]))throw Error('SCOPE_MISMATCH');
 makerToken=saved.token;session.setItem('booth-token',makerToken);connectionVerified=true;showConnection();return true;
}
async function invalidateMakerConnection(token){
 if(makerToken===token){makerToken=null;connectionVerified=false;if(session.getItem('booth-token')===token)session.removeItem('booth-token');}
 await clearMakerConnection(token);showConnection();
}
async function api(action,body={}){
 const manufacturer=!officeMode&&!localDemo;
 let token=officeMode&&!localDemo?portalSession()?.token:(officeMode||localDemo?session.getItem('booth-token'):makerToken);
 if(!token&&manufacturer&&await recoverMakerConnection(null))token=makerToken;
 if(!token)throw Error(officeMode?'PORTAL_SESSION_INVALID':'SESSION_INVALID');
 let result=await requestApi(action,body,token);
 // DBが認証拒否した場合だけ1回再試行する。通信不明時の注文再送は従来の冪等処理に任せる。
 if(!result.ok&&result.data.error==='SESSION_INVALID'&&manufacturer&&await recoverMakerConnection(token)){token=makerToken;result=await requestApi(action,body,token);}
 if(!result.ok){if(result.data.error==='SESSION_INVALID'&&manufacturer)await invalidateMakerConnection(token);throw Error(result.data.error||'SERVER_ERROR');}
 return result.data;
}
function tokenFromQr(value){try{const u=new URL(value);if(!['https:','http:'].includes(u.protocol)||!u.pathname.endsWith('/pass.html'))throw Error();const values=u.searchParams.getAll('t');if(values.length!==1||!values[0]||values[0].length>256)throw Error();return values[0];}catch{throw Error('名札QRのURLを読み取ってください。');}}
async function digest(text){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(text)))].map(x=>x.toString(16).padStart(2,'0')).join('');}
async function enroll(key){const response=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'enroll',key,device_id:deviceId()}),cache:'no-store'});const data=await response.json();if(!response.ok)throw Error(data.error||'ENROLLMENT_INVALID');makerToken=data.token;session.setItem('booth-token',data.token);const catalog=await api('catalog');if(!catalog.catalog)throw Error('このメーカーの商品が未登録です。');if(catalog.scope.device!==deviceId())throw Error('SCOPE_MISMATCH');connectionVerified=true;await set('maker-connection',{token:data.token,device:deviceId()});await activate(catalog);message(readyMessage(catalog));await pump();}
function element(tag,text,className){const e=document.createElement(tag);if(text!==undefined)e.textContent=text;if(className)e.className=className;return e;}
// 指定メーカーは新規入力を後日納品で開始し、選択済み下書き・送信待ちは保持する。
const defaultDelivery=scope=>['essence','napla','schwarzkopf','actland'].includes(scope?.booth)?'later':'';
const selectedDelivery=(record,scope)=>record.stage==='draft'&&!record.form.delivery_confirmed?defaultDelivery(scope):record.form.delivery;
async function fresh(){
 resetCustomerSearch();
 const delivery=defaultDelivery(info.scope);
 draft={id:crypto.randomUUID(),scope:info.scope,backend:location.origin,version:info.catalog.version,stage:'draft',form:{qr_hash:'',resolved:false,delivery,delivery_confirmed:!!delivery,ui_version:2,remarks:'',handling:emptyHandling(),offers:{}},createdAt:new Date().toISOString(),pending:null,cancelRequested:false};
 await update(draft.id,()=>draft);await set('draft:'+scopeKey(),draft.id);renderForm();
}
function numberInput(value,limit,description){
 const input=element('input');input.type='number';input.inputMode='numeric';input.min='0';input.max=String(limit);input.step='1';input.value=value>0?String(value):'';input.placeholder='0';input.setAttribute('aria-label',description);
 // 入力開始時に既存値を選択し、先頭の0が残る誤入力を防ぐ。
 input.addEventListener('focus',()=>input.select());
 input.addEventListener('blur',()=>{if(input.value!==''){const n=Number(input.value);if(Number.isInteger(n)&&n>=0)input.value=n?String(n):'';}});
 return input;
}
// 企画ごとの売価を優先し、サービス品の必要数と固定構成を共通化する。
const offerPrice=(offer,code)=>offer?.paid_unit_prices?.[code]??info.catalog.products.find(p=>p.product_code===code)?.event_sale_price_yen??0;
// この2企画は色番別の総数だけを入力する。正式名と購入/添付の保存形式は保持する。
const totalColor=offer=>offer.quantity_mode==='total_color'&&['schwarzkopf_bond1','schwarzkopf_terra'].includes(offer.offer_id);
function colorName(offer,name){
 const text=String(name||'').normalize('NFKC');
 if(offer.offer_id==='schwarzkopf_bond1')return text.replace(/^ファイバープレックス\s*ボンドカラー\s*80g\s*/i,'');
 if(offer.offer_id==='schwarzkopf_terra')return text.match(/^テラ\s+([A-Z0-9.-]+)/i)?.[1]||text;
 return makerText(name);
}
function colorCounts(selected){
 if(selected.total)return selected.total;
 const total={};for(const kind of ['paid','gift'])for(const [code,q] of Object.entries(selected[kind]||{}))total[code]=(total[code]||0)+q;
 return total;
}
function colorSelection(offer,selected,check){
 const totals=colorCounts(selected),n=selected.sets||0,entries=Object.entries(totals);
 if(check&&entries.some(([code,q])=>!offer.product_codes.includes(code)||!Number.isInteger(q)||q<0||q>10000))throw Error(`${makerText(offer.label)}：数量は0以上の整数で入力してください。`);
 const sum=entries.reduce((s,[,q])=>s+q,0),required=(offer.paid_per_set+offer.gift_per_set)*n;
 if(check&&sum!==required)throw Error(`${makerText(offer.label)}：合計本数を${required}本にしてください（入力済み${sum}本）。`);
 // どの色を添付にするかは入力不要。同一単価を確認した企画だけで自動配分する。
 if(new Set(offer.product_codes.map(code=>offerPrice(offer,code))).size!==1)throw Error('カラーの単価を確認できません。運営スタッフへ連絡してください。');
 let remaining=offer.paid_per_set*n;const result={...selected,paid:{},gift:{}};
 for(const code of offer.product_codes){const q=totals[code]||0;if(!(q>0))continue;const paid=Math.min(q,Math.max(remaining,0));if(paid)result.paid[code]=paid;if(q-paid)result.gift[code]=q-paid;remaining-=paid;}
 return result;
}
const requiredQty=(offer,kind,n)=>kind==='gift'&&offer.gift_rule==='threshold_gift'?Math.floor(n/offer.threshold_paid)*offer.threshold_gift:n*offer[`${kind}_per_set`];
// カタログに価格条件の説明があれば、その説明を優先する。
const offerSummary=offer=>offer.summary?offer.summary:totalColor(offer)?'3本分の料金で4本。色番ごとの合計本数を入力':offer.gift_rule==='price_only'?'特売単価で購入できます':offer.gift_rule==='threshold_gift'?`${offer.threshold_paid}箱ごとに${offer.threshold_gift}包サービス`:`${offer.paid_per_set}点購入 ＋ ${offer.gift_per_set}点サービス`;
const countLimit=offer=>offer.max_sets_per_person||99;
const countLabel=offer=>offer.count_label||(offer.gift_rule==='price_only'?(quantityByProduct(offer)?'合計購入数':'購入数量'):'セット数');
// 選べる商品が複数ある通常販売だけ、商品明細から合計を求める。
const quantityByProduct=offer=>offer.gift_rule==='price_only'&&offer.paid_per_set===1&&offer.gift_per_set===0&&offer.product_codes.length>1;
function selectedCount(offer,selected){
 if(!quantityByProduct(offer)){
  // v0.18.0の単品明細を、数量ボタンの値へ引き継ぐ。
  if(offer.gift_rule==='price_only'&&offer.product_codes.length===1&&Object.keys(selected.paid||{}).length)return Object.values(selected.paid).reduce((sum,q)=>sum+q,0);
  return selected.sets;
 }
 const quantities=Object.values(selected.paid||{});
 // 旧画面の単品下書きはセット数だけを保存しているため、その数量を引き継ぐ。
 return quantities.length?quantities.reduce((sum,q)=>sum+q,0):offer.product_codes.length===1?selected.sets||0:0;
}
function automaticKind(offer,kind){return !!offer[`${kind}_components`]||kind==='gift'&&['same_sku_groups','linked_sku_groups','threshold_gift'].includes(offer.gift_rule);}
// 3＋1は同じ企画・税抜単価ごとに数える。種類は購入・サービスとも自由に選べる。
function priceGroups(offer,paid,gift){
 const groups=new Map();
 for(const [kind,items] of [['paid',paid],['gift',gift]])for(const item of items){
  const price=offerPrice(offer,item.code);let group=groups.get(price);
  if(!group){group={price,paid:0,gift:0};groups.set(price,group);}
  group[kind]+=item.qty;
 }
 return [...groups.values()].sort((a,b)=>a.price-b.price);
}
const validPriceGroup=(offer,group)=>group.paid%offer.paid_per_set===0&&group.gift===group.paid/offer.paid_per_set*offer.gift_per_set;
// 表示順だけを単価順にし、カタログと同価格内の順番は変えない。
function displayCodes(offer,kind){
 const codes=offer[`${kind}_codes`]||offer.product_codes;
 return offer.gift_rule==='same_price_groups'?[...codes].sort((a,b)=>offerPrice(offer,a)-offerPrice(offer,b)):codes;
}
// 色に加えて残数を文章で示す。入力途中の値は消さず、修正できる状態を保つ。
function priceGiftGuidance(offer,paid,gift,invalid=false){
 if(invalid)return {state:'review',text:'数量は0以上の整数で入力してください。'};
 const available=Math.floor(paid/offer.paid_per_set)*offer.gift_per_set,remaining=available-gift,remainder=paid%offer.paid_per_set;
 const pending=remainder?` 購入はあと${offer.paid_per_set-remainder}点で次の${offer.gift_per_set}点サービスです。`:'';
 if(remaining<0)return {state:'review',text:`サービスが${-remaining}点多いです。この価格で選べるのは${available}点です。${pending}`};
 if(remaining>0)return {state:'available',text:`ここからあと${remaining}点選べます（購入${paid}点・サービス${gift}／${available}点）。${pending}`};
 if(available)return {state:remainder?'waiting':'complete',text:`${remainder?'選択済み':'✓ 選択完了'}：サービス${gift}／${available}点。${pending}`};
 return {state:'waiting',text:`この価格の商品をあと${offer.paid_per_set-paid}点購入すると、ここから${offer.gift_per_set}点選べます。`};
}
function updatePriceGiftGuidance(offer,card){
 for(const section of card.querySelectorAll('[data-price-kind=gift]')){
  const price=Number(section.dataset.price),items=kind=>[...card.querySelectorAll(`input[data-kind=${kind}]`)].filter(input=>offerPrice(offer,input.dataset.code)===price).map(input=>Number(input.value||0));
  const paid=items('paid'),gift=items('gift'),invalid=[...paid,...gift].some(q=>!Number.isInteger(q)||q<0||q>10000);
  const guide=priceGiftGuidance(offer,paid.reduce((a,b)=>a+b,0),gift.reduce((a,b)=>a+b,0),invalid);
  section.dataset.giftState=guide.state;section.querySelector('.price-guide').textContent=guide.text;
 }
}
function selectedOffer(offer){return draft.form.offers[offer.offer_id];}
// 内部の仮商品IDと照合情報は保持し、メーカー向けの表示だけ簡潔にする。
function makerText(value){
 const text=String(value||'');if(officeMode)return text;
 return text.replace(/仮ID・予約/g,'予約')
  .replace(/仮ID ACT26-[A-Z]+ で受付し、販売管理入力前に正式商品コードを照合してください。/g,'')
  .replace(/（販売管理コード要照合）/g,'')
  .replace(/[^。！？\n]*(?:マスター未登録|商品に統合|商品マスターの正式表記)[^。！？\n]*[。！？]?/g,'')
  .replace(/[^。！？\n]*通常売上単価×6÷7[^。！？\n]*[。！？]?/g,'')
  .replace(/本人指定により/g,'')
  .replace(/マスター売上単価|通常売上単価/g,'通常価格').trim();
}
function addSelectedProduct(group,offer,kind,code,quantity){
 const product=info.catalog.products.find(p=>p.product_code===code);if(!product)return;
 const label=kind==='total'?colorName(offer,product.product_name):makerText(product.product_name);
 const row=element('div',undefined,'selected-product');const name=element('div',label,'product-name');name.title=product.product_name;
 if(kind!=='total')name.append(element('small',`商品コード ${displayProductCode(code)} ／ ${kind==='paid'?yen(offerPrice(offer,code)):'サービス品・0円'}${kind==='gift'&&offer.gift_rule==='same_price_groups'?` ／ 同価格グループ ${yen(offerPrice(offer,code))}`:''}`));
 const input=numberInput(quantity,10000,`${label} ${kind==='total'?'合計本数':kind==='paid'?'購入数':'サービス数'}`);
 input.dataset.offer=offer.offer_id;input.dataset.kind=kind;input.dataset.code=code;
 if(automaticKind(offer,kind)){input.readOnly=true;input.classList.add('gift-auto');input.setAttribute('aria-label',`${makerText(product.product_name)} ${kind==='paid'?'購入':'サービス'}数（自動計算）`);}
 input.addEventListener('input',()=>{updateQuantityGuidance(offer,group.closest('.offer-card'));if(quantityByProduct(offer)||kind==='total'||offer.customer_kind_optional)saveForm().catch(e=>message(errorText(e)));});
 input.addEventListener('change',()=>{if(!quantityByProduct(offer)&&kind!=='total'&&!offer.customer_kind_optional)saveForm().catch(e=>message(errorText(e)));});
 row.append(name,input);group.append(row);
}
function renderOffer(offer){
 const selected=selectedOffer(offer);const card=element('div',undefined,'offer-card');card.dataset.offerCard=offer.offer_id;
 const head=element('div',undefined,'offer-head');const title=element('div');title.append(element('h3',makerText(offer.label)),element('p',offerSummary(offer),'muted'));
 const remove=element('button','条件を外す','secondary small-action');remove.type='button';remove.onclick=()=>removeOffer(offer.offer_id).catch(e=>message(errorText(e)));head.append(title,remove);card.append(head);
 const description=makerText(offer.description);if(description)card.append(element('p',description,'muted'));
 if(offer.customer_kind){const label=element('label',undefined,'eligibility');const checkbox=element('input');checkbox.type='checkbox';checkbox.dataset.eligibility=offer.offer_id;checkbox.checked=selected.eligibility===offer.customer_kind;checkbox.addEventListener('change',()=>saveForm().catch(e=>message(errorText(e))));label.append(checkbox,element('span',`${offer.brand==='nei'?'NEi':'hyumy'}の${offer.customer_kind==='new'?'新規':'既存取扱'}サロンであることを確認しました${offer.customer_kind_optional?'（任意・未チェックでも注文できます）':''}`));card.append(label);}
 if(!quantityByProduct(offer)){
 const countRow=element('div',undefined,'set-row');countRow.append(element('span',countLabel(offer),'set-label'));
 const count=numberInput(selectedCount(offer,selected),countLimit(offer),`${makerText(offer.label)}のセット数`);count.dataset.offer=offer.offer_id;count.dataset.kind='sets';
 countRow.append(count);
 const chips=element('div',undefined,'quick-counts');
 const syncChips=()=>chips.querySelectorAll('button').forEach(button=>button.setAttribute('aria-pressed',String(Number(count.value)>0&&Number(button.textContent)===Number(count.value))));
 count.addEventListener('input',()=>{syncChips();updateQuantityGuidance(offer,card);});
 count.addEventListener('change',()=>{syncChips();saveForm().catch(e=>message(errorText(e)));});
 for(const n of [1,2,3,4,5,6,8,10,16].filter(n=>n<=countLimit(offer))){const chip=element('button',String(n),'secondary');chip.type='button';chip.setAttribute('aria-label',`${makerText(offer.label)}を${n}セット`);chip.setAttribute('aria-pressed',String(Number(count.value)===n));chip.onclick=async()=>{count.value=String(n);syncChips();await saveForm();};chips.append(chip);}
 countRow.append(chips);card.append(countRow);
 }
 if(totalColor(offer)){
  const group=element('div',undefined,'product-group');group.append(element('h4','お渡しする商品の合計本数'));
  const progress=element('p',undefined,'quantity-progress');progress.dataset.progressKind='total';progress.setAttribute('role','status');group.append(progress);
  const grid=element('div',undefined,'color-grid'),choices=colorCounts(selected);
  for(const code of offer.product_codes)addSelectedProduct(grid,offer,'total',code,choices[code]||0);
  group.append(grid);card.append(group);updateQuantityGuidance(offer,card);
 }else if(offer.product_codes.length===1&&!quantityByProduct(offer)){
  const p=info.catalog.products.find(item=>item.product_code===offer.product_codes[0]);
  if(p)card.append(element('p',`${makerText(p.product_name)} ／ ${!officeMode&&p.temporary_product_id?'':`商品コード ${displayProductCode(p.product_code)} ／ `}税抜単価 ${yen(offerPrice(offer,p.product_code))}。${offer.gift_per_set>0?'サービス品は同じ商品です。':''}`,'muted'));
 }else{
  for(const kind of ['paid','gift']){
   if(kind==='gift'&&(quantityByProduct(offer)||!(offer.gift_codes||offer.product_codes).length))continue;
   const group=element('div',undefined,'product-group');group.append(element('h4',kind==='paid'?'購入する商品':'サービス品（0円）'));
   const progress=element('p',undefined,'quantity-progress');progress.dataset.progressKind=kind;progress.setAttribute('role','status');group.append(progress);
   const choices=selected[kind]||{};let target=group,lastPrice;
   if(kind==='gift'&&offer.gift_rule==='same_price_groups')group.append(element('p','購入数に応じて、選べる価格帯を青く表示します。青い欄からサービス品を選んでください。','muted'));
   // 複数商品を選べる条件は最初から全商品を並べ、未入力は0点として扱う。
   for(const code of displayCodes(offer,kind)){
    const price=offerPrice(offer,code);
    if(offer.gift_rule==='same_price_groups'&&price!==lastPrice){
     target=element('div',undefined,'price-section');target.dataset.priceKind=kind;target.dataset.price=String(price);
     target.append(element('h5',kind==='paid'?`税抜単価 ${yen(price)}`:`購入単価 ${yen(price)}に対応するサービス品`));
     if(kind==='gift'){const guide=element('p',undefined,'price-guide');guide.id=`gift-guide-${offer.offer_id}-${price}`;guide.setAttribute('role','status');target.append(guide);}
     group.append(target);lastPrice=price;
    }
    const quantity=quantityByProduct(offer)&&offer.product_codes.length===1&&!(code in choices)?selectedCount(offer,selected):choices[code]||0;
    addSelectedProduct(target,offer,kind,code,quantity);
    if(kind==='gift'&&offer.gift_rule==='same_price_groups')target.lastElementChild.querySelector('input').setAttribute('aria-describedby',`gift-guide-${offer.offer_id}-${price}`);
   }
   card.append(group);
  }
  updateQuantityGuidance(offer,card);
 }
 return card;
}
// 備考は注文単位で保存し、次のお客様には引き継がない。
function showRemarks(parent,text,{includeBilling=true}={}){
 if(!text)return;const h=unpackRemarks(text);
 if(h.invalid){parent.append(element('p',`備考：${text}`,'order-remarks'));return;}
 if(h.encoded){
  if(includeBilling&&h.billing!=='salon')parent.append(element('p',`請求先：${billingLabel(h.billing)}`));
  if(h.contact==='review')parent.append(element('p','配送・連絡：受付で確認'));
  if(h.paper)parent.append(element('p',`以前の紙照合番号：${h.paper}`,'muted'));
 }
 if(h.remarks)parent.append(element('p',`備考：${h.remarks}`,'order-remarks'));
}
function readHandling(){return {billing:$('billing').value||'review',contact:$('contact-check').checked?'review':'registered',paper:$('paper-number').value};}
function remarksCount(){let count;try{count=[...packRemarks($('remarks').value,readHandling())].length;}catch{count='要確認';}$('remarks-count').textContent=`${count} / 1,000文字（受付情報を含む）`;}
function renderForm(){
 $('remarks').value=draft.form.remarks||'';const handling={...emptyHandling(),...draft.form.handling};$('billing').value=['salon','personal'].includes(handling.billing)?handling.billing:'';$('billing-help').hidden=['salon','personal'].includes(handling.billing)||draft.stage!=='draft';$('contact-check').checked=handling.contact==='review';$('paper-number').value=handling.paper;$('legacy-paper-field').hidden=!handling.paper;$('handling-panel').open=handling.billing!=='salon'||handling.contact!=='registered'||!!handling.paper;remarksCount();$('another-purchase').hidden=!draft.form.saved_request_id;
 $('draft-id').textContent=`注文番号 ${draft.id}`;$('customer').textContent=draft.form.customer||'';showConnection();
 for(const radio of document.querySelectorAll('input[name=delivery]'))radio.checked=radio.value===selectedDelivery(draft,info.scope);
 $('offers').replaceChildren();
 for(const offer of info.catalog.offers)if(selectedOffer(offer))$('offers').append(renderOffer(offer));
 if(!$('offers').children.length)$('offers').append(element('p','まだ条件を選んでいません。「特売条件を追加」から始めてください。','empty-state'));
 const locked=submitting||draft.stage!=='draft';for(const x of [$('remarks'),$('billing'),$('contact-check'),$('paper-number'),$('submit'),$('scan'),$('demo'),$('customer-query'),$('customer-search'),...document.querySelectorAll('input[name=customer-search-mode]'),...$('customer-results').querySelectorAll('button'),$('add-offer'),...document.querySelectorAll('input[name=delivery]'),...$('offers').querySelectorAll('input,button')])x.disabled=locked;
 $('new').disabled=submitting;$('another-purchase').disabled=submitting;const status=$('submit-status');status.hidden=locked===false;status.className='';
 if(locked){if(submitting){$('submit').textContent='注文番号を確認中…';status.className='pending';status.textContent='登録済みの注文がないか確認しています。';}else if(draft.receipt?.state==='active'&&!draft.pending){$('submit').textContent='✓ 登録済み';status.className='success';status.textContent='登録が完了しました。次のお客様の入力欄へ切り替えています。';}
  else if(draft.receipt?.state==='cancelled'){ $('submit').textContent='取消済み';status.className='error';status.textContent='この注文は取消済みです。';}
  else{$('submit').textContent='送信・確認中…';status.className='pending';status.textContent=draft.error?`未送信・結果確認中：${draft.error}「未送信を再送・状態確認」を押してください。`:'注文を端末に保存しました。登録結果を確認中です。';}}
 else $('submit').textContent='この内容で登録する';
 calculate();
}
async function readForm(){updateSelectionSummaries();const offers={};for(const i of $('offers').querySelectorAll('input[data-offer]')){const o=offers[i.dataset.offer]||={sets:0,paid:{},gift:{}};if(i.dataset.kind==='sets')o.sets=i.value===''?0:Number(i.value);else (o[i.dataset.kind]||={})[i.dataset.code]=i.value===''?0:Number(i.value);}for(const offer of info.catalog.offers){const selected=offers[offer.offer_id];if(selected&&quantityByProduct(offer))selected.sets=selectedCount(offer,selected);}for(const c of $('offers').querySelectorAll('input[data-eligibility]'))if(offers[c.dataset.eligibility])offers[c.dataset.eligibility].eligibility=c.checked?info.catalog.offers.find(o=>o.offer_id===c.dataset.eligibility).customer_kind:null;const {qr:oldRawQr,...oldForm}=draft.form;const delivery=document.querySelector('input[name=delivery]:checked')?.value||'';return {...oldForm,delivery,delivery_confirmed:!!delivery,ui_version:2,remarks:$('remarks').value,handling:readHandling(),offers};}
function updateQuantityGuidance(offer,card){
 if(!card)return;
 if(totalColor(offer)){
  const progress=card.querySelector('[data-progress-kind=total]'),sets=Number(card.querySelector('input[data-kind=sets]')?.value||0);
  const quantities=[...card.querySelectorAll('input[data-kind=total]')].map(x=>Number(x.value||0)),total=quantities.reduce((a,b)=>a+b,0),required=sets*(offer.paid_per_set+offer.gift_per_set);
  const invalid=quantities.some(q=>!Number.isInteger(q)||q<0||q>10000)||!Number.isInteger(sets)||sets<0||sets>countLimit(offer);
  progress.className='quantity-progress';progress.classList.add(invalid||sets&&total!==required?'quantity-error':sets?'quantity-complete':'quantity-neutral');
  progress.textContent=invalid?'数量は0以上の整数、セット数は1〜99で入力してください。':!sets?'セット数を選んでください。':`合計：必要${required}本 ／ 入力済み${total}本`+(total===required?'　✓ 数量一致':total<required?`　あと${required-total}本`:`　${total-required}本多いです`);return;
 }
 if(quantityByProduct(offer)){
  const progress=card.querySelector('[data-progress-kind=paid]');if(!progress)return;
  const quantities=[...card.querySelectorAll('input[data-kind=paid]')].map(input=>Number(input.value||0));
  const total=quantities.reduce((sum,q)=>sum+q,0);
  const invalid=quantities.some(q=>!Number.isInteger(q)||q<0||q>10000);
  progress.className='quantity-progress';
  progress.classList.add(invalid||total>countLimit(offer)?'quantity-error':total?'quantity-complete':'quantity-neutral');
  progress.textContent=invalid?'数量は0以上の整数で入力してください。':total>countLimit(offer)?`合計購入数は${countLimit(offer)}点以内で入力してください。`:total?`購入：合計${total}点`:'購入する商品の数量を入力してください。';
  return;
 }
 const sets=Number(card.querySelector('input[data-kind=sets]')?.value||0);
 for(const kind of ['paid','gift']){
  for(const input of card.querySelectorAll(`input[data-kind=${kind}]`)){
   let value;
   if(offer[`${kind}_components`])value=(offer[`${kind}_components`][input.dataset.code]||0)*sets;
   else if(kind==='gift'&&offer.gift_rule==='threshold_gift')value=requiredQty(offer,kind,sets);
   else if(kind==='gift'&&['same_sku_groups','linked_sku_groups'].includes(offer.gift_rule)){
    value=0;
    for(const paid of card.querySelectorAll('input[data-kind=paid]')){
     const target=offer.gift_rule==='linked_sku_groups'?offer.gift_map[paid.dataset.code]:paid.dataset.code;
     const q=Number(paid.value||0);
     if(target===input.dataset.code&&Number.isInteger(q)&&q>0&&q%offer.paid_per_set===0)value+=q/offer.paid_per_set*offer.gift_per_set;
    }
   }
   if(value!==undefined)input.value=Number.isInteger(value)&&value>0?String(value):'';
  }
  const progress=card.querySelector(`[data-progress-kind=${kind}]`);if(!progress)continue;
  const quantities=[...card.querySelectorAll(`input[data-kind=${kind}]`)].map(x=>Number(x.value||0));
  progress.className='quantity-progress';
  if(quantities.some(q=>!Number.isInteger(q)||q<0)||!Number.isInteger(sets)||sets<0||sets>countLimit(offer)){progress.classList.add('quantity-error');progress.textContent=`数量は整数、${countLabel(offer)}は1〜${countLimit(offer)}で入力してください。`;continue;}
  const total=quantities.reduce((a,b)=>a+b,0),required=requiredQty(offer,kind,sets),difference=required-total;
  const base=`${kind==='paid'?'購入':'サービス'}：必要${required}点 ／ 入力済み${total}点`;
  progress.classList.add(!sets?'quantity-neutral':difference?'quantity-error':'quantity-complete');
  progress.textContent=!sets?`${countLabel(offer)}を選んでください。`:base+(difference>0?`　あと${difference}点`:difference<0?`　${-difference}点多いです`:'　✓ 数量一致');
 }
 if(offer.gift_rule==='same_price_groups')updatePriceGiftGuidance(offer,card);
 if(offer.gift_rule==='same_price_groups'&&sets>0){
  const items=kind=>[...card.querySelectorAll(`input[data-kind=${kind}]`)].map(x=>({code:x.dataset.code,qty:Number(x.value||0)})).filter(x=>x.qty!==0);
  const groups=priceGroups(offer,items('paid'),items('gift'));
  const progress=card.querySelector('[data-progress-kind=gift]');
  const invalid=groups.some(g=>!Number.isInteger(g.paid)||!Number.isInteger(g.gift)||g.paid<0||g.gift<0||!validPriceGroup(offer,g));
  if(progress&&groups.length){progress.textContent+=' ／ '+groups.map(g=>`${yen(g.price)}：購入${g.paid}・サービス${g.gift}`).join(' ／ ');if(invalid){progress.className='quantity-progress quantity-error';progress.textContent=progress.textContent.replace('　✓ 数量一致','')+'。同じ価格ごとに購入3点＋サービス1点にしてください。';}}
 }
}
function updateSelectionSummaries(){for(const offer of info.catalog.offers){if(!draft.form.offers[offer.offer_id]||offer.product_codes.length===1&&!quantityByProduct(offer))continue;const card=Array.from($('offers').querySelectorAll('.offer-card')).find(node=>node.dataset.offerCard===offer.offer_id);updateQuantityGuidance(offer,card);}}
// 入力イベントを順に保存し、同じタブの連続入力同士を競合と扱わない。
let formSaving=Promise.resolve();
function saveForm(){const task=formSaving.then(saveFormNow,saveFormNow);formSaving=task.catch(()=>{});return task;}
async function saveFormNow(){if(draft.stage!=='draft')return;const id=draft.id;const form=await readForm();const current=await update(id,r=>r.stage==='draft'?{...r,form:{...r.form,offers:form.offers,delivery:form.delivery,delivery_confirmed:form.delivery_confirmed,remarks:form.remarks,handling:form.handling,ui_version:2}}:r);if(draft.id!==id)return;draft=current;$('customer').textContent=draft.form.customer||'';calculate();updateSelectionSummaries();}
async function removeOffer(id){await saveForm();const current=draft.form.offers[id];if(current&&(current.sets>0||Object.values(current.paid||{}).some(Boolean)||Object.values(current.gift||{}).some(Boolean)||Object.values(current.total||{}).some(Boolean))&&!confirm('選択した条件と数量を外しますか？'))return;draft=await update(draft.id,r=>{const offers={...r.form.offers};delete offers[id];return {...r,form:{...r.form,offers}};});renderForm();}
function openPicker(state){if(draft.stage!=='draft')return;pickerState=state;modalReturnFocus=document.activeElement;$('picker-title').textContent='特売条件を追加';$('picker-search').value='';$('picker-modal').hidden=false;document.body.classList.add('modal-open');renderPicker();
 // 入力欄へ自動フォーカスせず、スマホのキーボードを開かない。検索は本人のタップ時だけ。
 $('picker-close').focus({preventScroll:true});}
function closePicker(){$('picker-modal').hidden=true;pickerState=null;document.body.classList.remove('modal-open');modalReturnFocus?.focus();}
function renderPicker(){
 const state=pickerState;if(!state)return;const query=$('picker-search').value.trim().toLocaleLowerCase('ja-JP');let choices;
 const selectedGroups=info.catalog.offers.filter(o=>draft.form.offers[o.offer_id]&&o.exclusive_group).map(o=>o.exclusive_group);
 choices=info.catalog.offers.filter(o=>!draft.form.offers[o.offer_id]&&!selectedGroups.includes(o.exclusive_group)).map(o=>({key:o.offer_id,name:makerText(o.label),sub:offerSummary(o)}));
 const matches=choices.filter(c=>c.name.toLocaleLowerCase('ja-JP').includes(query));$('picker-results').replaceChildren();
 for(const choice of matches.slice(0,20)){const button=element('button',choice.name,'picker-result');button.type='button';button.append(element('small',choice.sub));button.onclick=()=>choosePicker(choice.key).catch(e=>message(errorText(e)));$('picker-results').append(button);}
 $('picker-count').textContent=matches.length>20?`${matches.length}件中20件を表示。名前で絞り込んでください。`:`${matches.length}件`;
 if(!matches.length)$('picker-results').append(element('p','選べる候補がありません。','muted'));
}
async function choosePicker(key){if(!pickerState)return;await saveForm();draft=await update(draft.id,r=>{const offers={...r.form.offers};offers[key]={sets:0,paid:{},gift:{}};return {...r,form:{...r.form,offers}};});closePicker();renderForm();const target=Array.from($('offers').querySelectorAll('.offer-card')).find(node=>node.dataset.offerCard===key)?.querySelector('input[data-kind=sets],input[data-kind=paid]');target?.focus();}
function bundles(check=false){
 const groups=new Set();
 return info.catalog.offers.flatMap(o=>{
  const stored=draft.form.offers[o.offer_id];
  if(!stored)return [];
  const selected=totalColor(o)?colorSelection(o,stored,check):stored;
  if(check&&quantityByProduct(o)&&Object.values(selected.paid||{}).some(q=>!Number.isInteger(q)||q<0||q>10000))throw Error(`${makerText(o.label)}：数量は0以上の整数で入力してください。`);
  const n=selectedCount(o,selected);
  if(!n){if(check&&selected&&[...Object.values(selected.paid),...Object.values(selected.gift)].some(x=>x!==0))throw Error(`${makerText(o.label)}の${countLabel(o)}を指定してください。`);return [];}
  if(check&&(!Number.isInteger(n)||n<1||n>countLimit(o)))throw Error(`${makerText(o.label)}：${countLabel(o)}は1〜${countLimit(o)}の整数で指定してください。`);
  if(check&&o.ready===false)throw Error('この企画の条件は確認中です。');
  if(check&&o.customer_kind&&!o.customer_kind_optional&&selected.eligibility!==o.customer_kind)throw Error(`${makerText(o.label)}：新規・既存の確認にチェックしてください。`);
  if(check&&o.exclusive_group){if(groups.has(o.exclusive_group))throw Error('この2つの企画は併用できません。どちらかを選んでください。');groups.add(o.exclusive_group);}
  const value={offer_id:o.offer_id,sets:n,...(o.customer_kind&&(!o.customer_kind_optional||selected.eligibility)?{customer_kind:selected.eligibility}:{})};
  for(const kind of ['paid','gift']){
   const required=requiredQty(o,kind,n);
   value[kind]=o.product_codes.length===1&&!quantityByProduct(o)?(required?[{code:o.product_codes[0],qty:required}]:[]):Object.entries(kind==='paid'&&quantityByProduct(o)&&o.product_codes.length===1&&!Object.keys(selected.paid||{}).length?{[o.product_codes[0]]:n}:selected[kind]||{}).filter(([,q])=>q!==0).map(([code,qty])=>({code,qty}));
   if(check&&(value[kind].some(x=>!Number.isInteger(x.qty)||x.qty<1)||value[kind].reduce((a,x)=>a+x.qty,0)!==required))throw Error(`${makerText(o.label)}：${kind==='paid'?'購入':'サービス'}数は${required}点にしてください。`);
   if(check&&o[`${kind}_components`]&&Object.entries(o[`${kind}_components`]).some(([code,q])=>value[kind].find(x=>x.code===code)?.qty!==q*n))throw Error('固定セットの構成を確認してください。');
  }
  if(check&&['same_sku_groups','linked_sku_groups'].includes(o.gift_rule)){
   const expected={};for(const x of value.paid){if(x.qty%o.paid_per_set!==0)throw Error(`同じ商品・容量ごとに購入数を${o.paid_per_set}の倍数にしてください。`);const code=o.gift_rule==='linked_sku_groups'?o.gift_map[x.code]:x.code;expected[code]=(expected[code]||0)+x.qty/o.paid_per_set*o.gift_per_set;}
   if(value.gift.some(x=>x.qty!==expected[x.code])||Object.keys(expected).length!==value.gift.length)throw Error('同じ種類のサービス品の数量を確認してください。');
  }
  if(check&&o.gift_rule==='same_price_groups'&&priceGroups(o,value.paid,value.gift).some(g=>!validPriceGroup(o,g)))throw Error(`${makerText(o.label)}：同じ価格ごとに購入3点＋サービス1点にしてください。`);
  if(check&&o.gift_rule==='pair_oil'&&o.paid_groups.some(g=>value.paid.filter(x=>g.includes(x.code)).reduce((a,x)=>a+x.qty,0)!==n))throw Error('シャンプーとトリートメントを各セット数と同じ本数にしてください。');
  return [value];
 });
}
function calculate(){const amount=bundles().reduce((sum,b)=>sum+b.paid.reduce((a,x)=>a+x.qty*offerPrice(info.catalog.offers.find(o=>o.offer_id===b.offer_id),x.code),0),0);$('total').textContent=yen(amount);}
async function activate(value){info=value;$('login').hidden=true;if(!officeMode)await set('maker-catalog',info);$('maker').textContent=info.catalog?.manufacturer_name||'メーカー未設定';$('maker-banner').hidden=info.role==='admin';$('workspace').hidden=info.role==='admin';$('admin').hidden=info.role!=='admin';$('ledger').hidden=info.role!=='admin';$('identity').hidden=info.role!=='admin';if(info.role==='admin')return;const id=await get('draft:'+scopeKey());draft=(await records()).find(x=>x.id===id);if(!draft)await fresh();else{if(draft.stage==='draft'&&draft.form.ui_version!==2){draft=await update(draft.id,r=>{const offers={};for(const [key,item] of Object.entries(r.form.offers||{})){const paid=Object.fromEntries(Object.entries(item.paid||{}).filter(([,qty])=>qty>0));const gift=Object.fromEntries(Object.entries(item.gift||{}).filter(([,qty])=>qty>0));if(item.sets>0||Object.keys(paid).length||Object.keys(gift).length)offers[key]={sets:item.sets||0,paid,gift};}return {...r,form:{...r.form,offers,delivery:'',delivery_confirmed:false,ui_version:2}};});}if(draft.stage!=='draft'&&draft.receipt?.state==='active'&&!draft.pending)await fresh();else renderForm();}await renderOrders();}
function orderDetails(record){
 // 登録確定後はサーバーの確定明細を使用し、通信待ちは端末に保存した注文内容と明示する。
 let lines=Array.isArray(record.receipt?.lines)?record.receipt.lines:null;
 const confirmed=!!lines;
 if(!lines&&record.body?.catalog_version===info.catalog.version){
  lines=(record.body.bundles||[]).flatMap(bundle=>['paid','gift'].flatMap(kind=>(bundle[kind]||[]).map(item=>{
   const product=info.catalog.products.find(p=>p.product_code===item.code);
   return product?{kind,offer_id:bundle.offer_id,offer_version:info.catalog.offers.find(o=>o.offer_id===bundle.offer_id)?.offer_version,product_code:item.code,product_name:product.product_name,qty:item.qty,unit_price_yen:kind==='gift'?0:offerPrice(info.catalog.offers.find(o=>o.offer_id===bundle.offer_id),item.code)}:null;
  }).filter(Boolean)));
 }
 const details=element('div',undefined,'order-details');
 const paid=(lines||[]).filter(line=>line.kind==='paid').reduce((sum,line)=>sum+Number(line.qty||0),0);
 const gift=(lines||[]).filter(line=>line.kind==='gift').reduce((sum,line)=>sum+Number(line.qty||0),0);
 const colorOffers=officeMode?[]:(info.catalog?.offers||[]).filter(totalColor);
 const combined=(lines||[]).filter(line=>line.offer_version>=2&&colorOffers.some(o=>o.offer_id===line.offer_id));
 details.append(element('h4',lines?.length?combined.length?`商品明細（合計${paid+gift}点）`:`商品明細（購入${paid}点・サービス${gift}点）`:'商品明細'));
 if(!lines?.length){details.append(element('p','商品明細を取得できません。事務担当者に確認してください。','error'));return details;}
 if(!confirmed)details.append(element('p','端末に保存した送信前の内容です。登録確定後に再確認してください。','muted'));
 for(const offer of colorOffers){
  const selected=combined.filter(line=>line.offer_id===offer.offer_id);if(!selected.length)continue;
  const quantities=new Map();let amount=0;
  for(const line of selected){const key=line.product_code;const item=quantities.get(key)||{name:line.product_name,qty:0};item.qty+=Number(line.qty);quantities.set(key,item);amount+=Number(line.qty)*Number(line.unit_price_yen||0);}
  details.append(element('h4',makerText(offer.label)));
  for(const item of quantities.values())details.append(element('p',`${colorName(offer,item.name)}　${item.qty}本`,'order-line'));
  details.append(element('p',`合計${selected.reduce((sum,line)=>sum+Number(line.qty),0)}本 ／ 税抜${yen(amount)}`));
 }
 for(const kind of ['paid','gift']){
  const list=lines.filter(line=>line.kind===kind&&!combined.includes(line));
  if(!list.length)continue;
  details.append(element('h4',kind==='paid'?'購入商品':'サービス品（0円）'));
  for(const line of list){const qty=Number(line.qty||0);const price=Number(line.unit_price_yen||0);details.append(element('p',`${makerText(line.product_name)}　${qty}点 × ${yen(price)}${kind==='paid'?` ＝ ${yen(qty*price)}`:''}`,'order-line'));}
 }
 return details;
}
async function renderOrders(){
 if(!info)return;
 // 再送や取消で一覧が更新されても、確認中の行は開いた状態を維持する。
 const opened=new Set(Array.from($('orders').querySelectorAll('details[open]')).map(node=>node.dataset.orderId));
 $('orders').replaceChildren();
 const list=(await records()).filter(r=>r.backend===location.origin&&JSON.stringify(r.scope)===scopeKey()&&r.stage!=='superseded'&&(r.stage==='draft'?r.id!==draft?.id&&(r.form.resolved||Object.values(r.form.offers||{}).some(o=>o.sets>0)):!!r.pending||!r.receipt||!!r.error)).sort((a,b)=>b.createdAt.localeCompare(a.createdAt));
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
  const delivery=r.body?.delivery||r.form.delivery;body.append(element('p',`受け渡し：${delivery==='later'?'後日納品':delivery==='takeaway'?'当日お渡し済み':'未選択'}`,'order-delivery'));
  body.append(element('p',`注文番号 ${r.id}`,'muted order-reference'));
  showRemarks(body,r.body?.remarks??packRemarks(r.form.remarks||'',r.form.handling));
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
async function pumpInner(){if(!info||!(localDemo?session.getItem('booth-token'):makerToken)||!navigator.onLine)return;for(const entry of await records()){if(entry.backend!==location.origin||JSON.stringify(entry.scope)!==scopeKey()||['draft','superseded'].includes(entry.stage))continue;const latest=(await records()).find(x=>x.id===entry.id);if(!latest.pending)continue;try{let response;
 if(latest.serverIntent&&latest.pending==='submit'){
  // 通信結果が不明でも、先に同じ注文番号を照会する。新しい番号は作らない。
  response=await api('status',{recover_request:latest.id});
  if(response.state==='prepared')response=await api('submit',{operation:'commit',request_id:latest.id});
  if(!['active','cancelled'].includes(response.state))throw Error('ORDER_INTENT_NOT_FOUND');
 }else response=await api(latest.pending==='cancel'?'cancel':latest.pending==='submit'?'submit':'status',latest.pending==='submit'?latest.body:{request_id:latest.id});await update(latest.id,current=>({...mergeReceipt(current,response),error:null}));}catch(e){await update(latest.id,current=>({...current,error:errorText(e)}));}}if(draft&&draft.stage!=='draft'){const current=(await records()).find(r=>r.id===draft.id);if(current){draft=current;if(current.receipt?.state==='active'&&!current.pending){
   // サーバー登録が確定した時だけ新しい入力欄を開く。前の注文は端末の注文一覧に残す。
   if(current.serverIntent)showReceipt(current.receipt);await fresh();message('登録完了。次のお客様の名札QRを読み取ってください。前の注文は下の一覧から確認・取消できます。');$('message').scrollIntoView({block:'start'});$('scan').focus({preventScroll:true});
  }else renderForm();}}await renderOrders();await refreshHistory();}
async function pump(){if(pumping)return;pumping=true;try{if(navigator.locks)await navigator.locks.request('beaufes-booth-send',pumpInner);else await pumpInner();}finally{pumping=false;}}

function showConnection(){
 if(officeMode)return;
 const connected=localDemo?!!session.getItem('booth-token'):!!makerToken;
 const ready=navigator.onLine&&connected&&connectionVerified;$('connection-status').hidden=ready&&storageMode()==='persistent';if(storageMode()==='memory'){$('connection-status').textContent='このブラウザーには下書きを保存できません。オンラインで登録してください。再読み込み後はメーカーQRで接続し、同じお客様を選ぶと登録結果を確認できます。';return;}
 $('connection-status').textContent=ready?'':!navigator.onLine?'インターネットにつながっていません。未送信注文はこの端末に保存されます。検索・履歴の取得には通信が必要です。':!connected?'メーカー用QRを読み取ってから、お客様を選んでください。':'注文画面を準備しています…';
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
function customerLabel(customer){return `来場者 ${customer.id||'要確認'} ／ ${customer.salon} ／ ${customer.name}${customer.customer_code?`（得意先 ${customer.customer_code}）`:customer.customer_linked?'':'（得意先コードは事務確認）'}`;}
async function chooseCustomer(customer,selection,id,serial){
 if(draft.id!==id||draft.stage!=='draft'||customerRequest!==serial)return;
 const label=customerLabel(customer);
 const current=await update(id,r=>draft.id===id&&r.stage==='draft'&&customerRequest===serial?{...r,form:{...r.form,...selection,customer:label,customer_id:customer.id,resolved:true,previous_request_id:null,saved_request_id:null}}:r);
 if(draft.id!==id||customerRequest!==serial)return;
 draft=current;$('customer').textContent=label;$('customer-search-panel').open=false;
 message('お客様を選びました。名札のお名前と照合してから商品を登録してください。');
 await customerRecent(customer.id,id,serial);
 if(info.server_order_recovery&&navigator.onLine){
  const recovered=await api('status',{recover_subject:customer.id});
  if(draft.id!==id||customerRequest!==serial)return;
  if(recovered.already_saved){draft=await update(id,r=>({...r,form:{...r.form,saved_request_id:recovered.request_id}}));renderForm();message('このお客様の注文は登録済みです。下の履歴で控えを確認できます。追加購入するときだけ「別の購入を開始する」を押してください。');}
 }

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
  body.append(element('p',customerLabel(data.customer),'order-customer-full'),element('p',`受け渡し：${data.delivery==='later'?'後日納品':'当日お渡し済み'}`,'order-delivery'),element('p',`注文番号 ${key}`,'muted order-reference'));
  showRemarks(body,data.remarks);body.append(orderDetails({receipt:data}));const receiptButton=element('button','購入控えを表示','secondary');receiptButton.onclick=()=>showReceipt(data.purchase_receipt||data);body.append(receiptButton);
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
 if(!(localDemo?session.getItem('booth-token'):makerToken)||!navigator.onLine){$('history-status').textContent=historyFetched?`前回取得 ${historyFetched}。通信・メーカー接続を確認して更新してください。`:'履歴の取得には通信とメーカーQRでの接続が必要です。';return;}
 const serial=++historyRequest;$('history-status').classList.remove('history-error');$('history-status').textContent='履歴を取得しています…';$('history-refresh').disabled=true;$('history-more').disabled=true;
 try{const data=await api('manufacturer_history',more&&historyCursor?{cursor:historyCursor}:{});if(serial!==historyRequest||scope!==scopeKey())return;
  sharedRows=more?[...new Map([...sharedRows,...data.rows].map(row=>[row.request_id,row])).values()]:data.rows;historyCursor=data.next_cursor;historyFetched=new Date(data.fetched_at).toLocaleString('ja-JP');
  renderSharedHistory();$('history-status').textContent=`取得 ${historyFetched} ／ ${sharedRows.length}件表示${historyCursor?'（続きあり）':''}`;
 }catch(e){if(serial===historyRequest&&scope===scopeKey()){$('history-status').classList.add('history-error');$('history-status').textContent=`取得できませんでした。${historyFetched?`前回取得 ${historyFetched}の表示です。`:''}${errorText(e)}`;}}
 finally{if(serial===historyRequest){$('history-refresh').disabled=false;$('history-more').disabled=false;}}
}
on('history-refresh','click',()=>refreshHistory());on('history-more','click',()=>refreshHistory(true));

on('connect','click',async()=>{if(!officeMode||localDemo){makerToken=$('token').value;session.setItem('booth-token',makerToken);}const value=await api('catalog');if(!value.catalog)throw Error('このメーカーの商品が未登録です。');connectionVerified=true;await activate(value);$('token').value='';message(officeMode?'事務画面に接続しました。':readyMessage(value));if(!officeMode)await pump();});
document.querySelectorAll('input[name=delivery]').forEach(radio=>radio.addEventListener('change',()=>saveForm().catch(e=>message(errorText(e)))));
on('add-offer','click',()=>openPicker({mode:'offer'}));
on('picker-search','input',renderPicker);
on('picker-close','click',closePicker);
on('demo','click',()=>readBadge('https://example.invalid/pass.html?t=demo-ticket'));
on('new','click',async()=>{if(draft.stage==='draft'&&bundles().length&&!confirm('現在の未登録の選択を残して、新しい注文を開きますか？'))return;closeCamera();await saveForm();await fresh();await renderOrders();message('新しい注文を開きました。');});
for(const id of ['billing','contact-check','paper-number'])on(id,id==='paper-number'?'input':'change',async()=>{if(id==='billing')$('billing-help').hidden=['salon','personal'].includes($('billing').value);remarksCount();await saveForm();});
on('remarks','input',async()=>{remarksCount();await saveForm();});
on('submit','click',async()=>{
 if(submitting||draft.stage!=='draft')return;
 submitting=true;
 try{
  await saveForm();
  if(!['salon','personal'].includes(draft.form.handling.billing))throw Error('請求先をサロンまたは個人から選んでください。');
  const encodedRemarks=packRemarks(draft.form.remarks||'',draft.form.handling);
  if(unpackRemarks(encodedRemarks).encoded&&!info.order_handling_v1)throw Error('受付情報の更新を確認できません。オンラインで画面を開き直してください。');
  if(!draft.form.resolved||!(draft.form.qr_hash||draft.form.subject_id))throw Error('先に名札QRまたはコード・名前検索でお客様を選んでください。');
  const b=bundles(true);if(!b.length)throw Error('商品を1セット以上選択してください。');
  if(!draft.form.delivery_confirmed||!['later','takeaway'].includes(draft.form.delivery))throw Error('受け渡し方法を選んでください。');
  if(draft.form.delivery==='takeaway'&&b.some(x=>[...(x.paid||[]),...(x.gift||[])].some(l=>/^(?:ACT26-|012693$)/.test(l.code))))throw Error('RESERVATION_DELIVERY_REQUIRED');
  const body={request_id:draft.id,scope:info.scope,catalog_version:info.catalog.version,...(draft.form.subject_id?{subject_id:draft.form.subject_id}:{token_hash:draft.form.qr_hash}),delivery:draft.form.delivery,bundles:b,...(encodedRemarks.trim()?{remarks:encodedRemarks}:{})};
  renderForm();
  if(info.server_order_recovery){
   if(!navigator.onLine)throw Error('通信が戻ってから登録してください。下書きから同じ注文を確認します。');
   const {request_id,...order}=body;
   const result=await api('submit',{operation:'prepare',order,...(draft.form.previous_request_id?{previous_request_id:draft.form.previous_request_id}:{})});
   if(result.already_saved){
    draft=await update(draft.id,r=>({...r,form:{...r.form,saved_request_id:result.request_id}}));
    showReceipt(result);message('このお客様の注文は登録済みです。控えを確認してください。追加購入するときだけ「別の購入を開始する」を押してください。');
    return;
   }
   if(result.state!=='prepared')throw Error('SERVER_ERROR');
   const source=draft;
   draft=await update(result.request_id,r=>r||({...source,id:result.request_id,body:{...order,request_id:result.request_id},serverIntent:true,stage:'queued',pending:'submit'}));
   await set('draft:'+scopeKey(),draft.id);
   await update(source.id,r=>({...r,stage:'superseded',pending:null}));
  }else{
   if(storageMode()==='memory')throw Error('保存領域が使えません。注文復旧対応の接続先を確認してください。');
   draft=await update(draft.id,r=>r.stage!=='draft'?r:{...r,body,stage:'queued',pending:'submit'});
  }
  message('同じ注文番号で登録結果を確認します。');await renderOrders();await pump();
 }finally{submitting=false;renderForm();}
});
let additionalPurchase;
on('another-purchase','click',()=>{
 if(submitting||!draft.form.saved_request_id)return;
 additionalPurchase={id:draft.id,revision:draft.edit_revision||0};
 $('additional-customer').textContent=draft.form.customer;
 $('additional-modal').hidden=false;$('additional-cancel').focus();
});
on('additional-cancel','click',()=>{$('additional-modal').hidden=true;additionalPurchase=null;});
on('additional-start','click',async()=>{
 if(submitting||!additionalPurchase||draft.id!==additionalPurchase.id||(draft.edit_revision||0)!==additionalPurchase.revision)throw Error('DRAFT_CHANGED');
 draft=await update(draft.id,r=>({...r,form:{...r.form,previous_request_id:r.form.saved_request_id,saved_request_id:null,offers:{},remarks:'',handling:emptyHandling(),delivery:'',delivery_confirmed:false}}));
 $('additional-modal').hidden=true;additionalPurchase=null;
 renderForm();message('同じお客様の別の購入を開始しました。商品と受け渡し方法を選んでください。');
});
function showReceipt(receipt){
 const target=$('receipt-content');target.replaceChildren();
 target.append(element('h2','購入控え'),element('p',receipt.state==='cancelled'?'取消済み':'登録済み','receipt-state'),element('p',`注文番号 ${receipt.request_id}`));
 const customer=receipt.customer||{};target.append(element('p',`${customer.salon||''} ／ ${customer.name||''}`),element('p',`申込番号 ${customer.id||''}`));
 const handling=unpackRemarks(receipt.remarks||'');target.append(element('p',`名札の得意先 ${customer.customer_code||'事務確認'} ／ 請求先：${billingLabel(handling.billing)}`));
 if(receipt.received_at)target.append(element('p',new Date(receipt.received_at).toLocaleString('ja-JP')));
 target.append(element('p',`受け渡し：${receipt.delivery==='later'?'後日納品':receipt.delivery==='takeaway'?'当日お渡し済み':'要確認'}`),orderDetails({receipt}),element('strong',`税抜合計 ${yen(receipt.total_yen)}`));
 showRemarks(target,receipt.remarks,{includeBilling:false});target.append(element('p','サービス品は0円です。メール送付を保証する画面ではありません。必要な場合はこの控えを撮影し、注文番号をお控えください。'));
 $('receipt-modal').hidden=false;$('receipt-close').focus();
}
on('receipt-close','click',()=>{$('receipt-modal').hidden=true;});
on('receipt-print','click',()=>window.print());
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
document.addEventListener('keydown',e=>{if(e.key==='Escape'){if(!$('additional-modal').hidden){$('additional-modal').hidden=true;additionalPurchase=null;}else if(!$('receipt-modal').hidden)$('receipt-modal').hidden=true;else if(!$('camera-modal').hidden)closeCamera();else if(!$('picker-modal').hidden)closePicker();}});
on('close','click',async()=>{if(!confirm('試験受付を締めます。未送信の注文がないか確認しましたか？'))return;await api('close');message('受付を締めました。未接続端末の注文は別途確認してください。');});
on('admin-refresh','click',refreshAdmin);
// 事務の照合は人が販売管理を確認した後に行う。アプリから売上入力・取消・送信は行わない。
const postingLabel=state=>({ready:'入力候補',claimed:'担当者が入力中',entered:'入力済み',unknown:'入力結果が不明',correction_required:'訂正必要',resolved_cancelled:'取消・売上なしの照合完了'}[state]||state);
function labeledField(parent,title,type='text',value=''){
 const label=element('label',title);const input=element(type==='textarea'?'textarea':'input');if(type!=='textarea')input.type=type;input.value=value;input.maxLength=type==='textarea'?2000:100;label.append(input);parent.append(label);return input;
}
function checkedField(parent,title){const label=element('label');const input=element('input');input.type='checkbox';label.append(input,document.createTextNode(' '+title));parent.append(label);return input;}
async function renderResolution(node,order){
 const context=order.resolution_context;if(!context)return;
 if(context.history?.length){const history=element('details');history.append(element('summary',`照合履歴 ${context.history.length}件`));for(const item of context.history){history.append(element('p',`${item.checked_at} ／ ${item.actor} ／ ${postingLabel(item.before_posting.state)} → ${postingLabel(item.after_posting.state)} ／ 操作番号 ${item.resolution_id}`),element('p',item.evidence.search_description));const sale=item.evidence.sale;if(sale)history.append(element('p',`伝票 ${sale.slip_number} ／ 売上日 ${sale.sales_date}`));if(item.evidence.correction_reference)history.append(element('p',`訂正の記録 ${item.evidence.correction_reference}`));}node.append(history);}
 const key='office-resolution:'+JSON.stringify([info.scope.event,order.booth,order.request_id]);
 let pending=await get(key);
 const unresolved=['unknown','correction_required'].includes(order.posting.state);
 if(!unresolved&&!pending)return;
 const panel=element('details');panel.append(element('summary',pending?'前の照合操作の結果を確認':'販売管理で照合して解決する'));node.append(panel);
 const status=element('p','販売管理で得意先・購入品・サービス品・数量・金額と、他担当／未保存の入力を確認してください。');panel.append(status);
 const recovery=element('div');panel.append(recovery);
 const form=element('form');panel.append(form);
 const choice=element('select');const choices=order.state==='cancelled'?[['no_sale','売上が存在しないことを確認した'],['old_sale_voided','元の売上の訂正・取消と、現在の売上なしを確認した']]:[['matching_sale','現在の内容に一致する伝票が1件ある'],['no_sale','売上が存在しないことを確認した'],['old_sale_voided','元の売上の訂正・取消と、現在の売上なしを確認した']];
 const label=element('label','照合結果');for(const [value,title] of choices){const option=element('option',title);option.value=value;choice.append(option);}label.append(choice);form.append(label);
 const search=labeledField(form,'検索した期間・得意先・担当範囲と、他担当者／未保存の確認内容','textarea');search.required=true;
 const complete=checkedField(form,'販売管理を開き、検索範囲と他担当者を含めた確認を終えた');complete.required=true;
 const saleBox=element('div');form.append(saleBox);
 const slip=labeledField(saleBox,'確認した伝票番号');const date=labeledField(saleBox,'確認した売上日','date');
 const matched=checkedField(saleBox,'得意先コード・購入品と0円サービス品・数量・税抜単価・合計が、この購入控えと一致した');
 const noSaleBox=element('div');form.append(noSaleBox);const unsaved=checkedField(noSaleBox,'該当売上も、他担当者や未保存の入力も残っていない');
 const oldBox=element('div');form.append(oldBox);const original=labeledField(oldBox,'元の伝票番号','text',order.posting.slip_number||'');
 const correctionBox=element('div');form.append(correctionBox);const reference=labeledField(correctionBox,'訂正・取消または過去入力解消の記録番号');
 const corrected=checkedField(correctionBox,'元の誤入力・重複が解消し、訂正後の記録を確認した');
 const ruledBox=element('div');form.append(ruledBox);const ruled=checkedField(ruledBox,'過去の伝票が存在する可能性も確認し、売上がないと判断した');
 const outcome=element('p');form.append(outcome);
 const approved=checkedField(form,'表示した注文番号と反映後の状態を確認し、この照合結果を記録する');approved.required=true;
 const send=element('button','照合結果を確認して台帳へ反映');send.type='submit';form.append(send);
 const updateChoice=()=>{const matching=choice.value==='matching_sale',voided=choice.value==='old_sale_voided',correction=order.posting.state==='correction_required';saleBox.hidden=!matching;noSaleBox.hidden=matching;oldBox.hidden=!voided;correctionBox.hidden=!(voided||matching&&correction);ruledBox.hidden=choice.value!=='no_sale'||!(correction||order.posting.slip_number);slip.required=matching;date.required=matching;matched.required=matching;unsaved.required=!matching;original.required=voided;reference.required=voided||matching&&correction;corrected.required=reference.required;ruled.required=!ruledBox.hidden;outcome.textContent='注文 '+order.request_id+' を '+(matching?'入力済み':order.state==='cancelled'?'取消・売上なしの照合完了（再入力不可）':'再入力候補')+' として記録します。';approved.checked=false;};choice.onchange=updateChoice;updateChoice();
 const finish=async()=>{await set(key,null);pending=null;message('照合結果を台帳へ記録しました。購入控えと注文内容は保持されています。');await refreshAdmin();};
 const verify=async()=>{
  const result=await api('posting',{operation:'resolution_status',resolution_id:pending.resolution_id,booth:pending.booth,request_id:pending.request_id});
  if(result.state==='applied'){await finish();return true;}
  if(result.state!=='not_found')throw Error('RESOLUTION_RESULT_UNKNOWN');
  status.textContent='この操作番号の反映記録はありません。同じ内容で再試行するか、最新台帳を取得して照合をやり直してください。';return false;
 };
 const recovering=()=>{
  form.hidden=!!pending||!unresolved;recovery.replaceChildren();if(!pending)return;
  panel.open=true;status.textContent='前の照合操作の結果を確認します。別の操作として登録し直さないでください。操作番号 '+pending.resolution_id;
  const action=(title,fn)=>{const button=element('button',title,'secondary');button.type='button';button.onclick=async()=>{button.disabled=true;try{await fn();}catch(e){status.textContent=errorText(e)+'。前の操作番号を保持しています。';}finally{button.disabled=false;}};recovery.append(button);};
  action('反映結果を確認',verify);
  action('同じ操作番号・同じ内容で再試行',async()=>{if(await verify())return;if(pending.snapshot_hash!==context.snapshot_hash)throw Error('ORDER_CHANGED_RECHECK_REQUIRED');await api('posting',pending);await finish();});
  action('未反映を確認して照合をやり直す',async()=>{if(await verify())return;if(!confirm('未反映を確認しました。最新の購入控えと販売管理をもう一度照合しますか？'))return;await set(key,null);pending=null;await refreshAdmin();});
 };recovering();
 form.onsubmit=async event=>{
  event.preventDefault();if(pending)return;send.disabled=true;
  try{
   if(storageMode()!=='persistent')throw Error('OFFICE_RESOLUTION_STORAGE_REQUIRED');
   const result=choice.value;const target=result==='matching_sale'?'入力済み':order.state==='cancelled'?'取消・売上なしの照合完了（再入力不可）':'再入力候補';
   if(!approved.checked)throw Error('照合結果と反映後の状態を確認してください。');
   const evidence={checked_by:info.scope.device,checked_at:new Date().toISOString(),search_description:search.value.trim(),sales_system_available:true,search_complete:complete.checked,unsaved_work_checked:unsaved.checked,result,match_count:result==='matching_sale'?1:0};
   if(result==='matching_sale'){if(!matched.checked)throw Error('LINES_MISMATCH');evidence.sale={slip_number:slip.value.trim(),sales_date:date.value,customer_code:order.customer.customer_code,lines:order.lines,total_yen:order.total_yen};}
   if(!correctionBox.hidden){evidence.prior_sales_resolved=corrected.checked;evidence.correction_verified=corrected.checked;evidence.correction_reference=reference.value.trim();}
   if(result==='old_sale_voided')evidence.original_slip_number=original.value.trim();
   if(!ruledBox.hidden)evidence.prior_sale_ruled_out=ruled.checked;
   const request={operation:'resolve',resolution_id:crypto.randomUUID(),request_id:order.request_id,booth:order.booth,revision:order.posting.revision,state_seq:order.state_seq,snapshot_hash:context.snapshot_hash,evidence};
   await set(key,request);if(storageMode()!=='persistent')throw Error('OFFICE_RESOLUTION_STORAGE_REQUIRED');pending=request;recovering();
   await api('posting',pending);await finish();
  }catch(e){status.textContent=errorText(e)+(pending?'。同じ操作番号で反映結果を確認してください。':'');}finally{send.disabled=false;}
 };
}

async function refreshAdmin(){const data=await api('admin_orders');$('admin-orders').replaceChildren();for(const order of data.rows){const node=element('div',undefined,'order');node.append(element('strong',`${order.customer.salon||'未対応'} ／ ${order.customer.name||''} ／ ${yen(order.total_yen)}`),element('p',`得意先 ${order.customer.customer_code||'要確認'} ／ ${(order.state==='cancelled'?'取消 ／ ':'')+postingLabel(order.posting.state)} ／ ${order.request_id}`,'muted'));
 const actionButton=(label,state)=>{const button=element('button',label,'secondary');button.onclick=async()=>{try{let slip_number,sales_date;if(state==='claimed'&&review.length&&!confirm(review.join('\n')+'\n正式コード・請求先・配送先の照合が完了しましたか？ 未確認ならキャンセルしてください。'))return;if(state==='entered'){slip_number=prompt('販売管理に入力した伝票番号を入力してください。');if(!slip_number)return;sales_date=prompt('販売管理に記録した売上日を YYYY-MM-DD で入力してください。');if(!sales_date)return;if(!/^\d{4}-\d{2}-\d{2}$/.test(sales_date))throw Error('売上日は YYYY-MM-DD で入力してください。');}await api('posting',{request_id:order.request_id,booth:order.booth,revision:order.posting.revision,state,slip_number,sales_date});await refreshAdmin();}catch(e){message(errorText(e));}};node.append(button);};
 const review=handlingWarnings(order);if(review.length)node.append(element('p',review.join('／'),'error'));
 if(order.state==='active'&&order.customer.customer_code&&order.posting.state==='ready')actionButton('この注文の入力を担当する','claimed');
 if(order.posting.state==='claimed'&&order.posting.owner===info.scope.device){actionButton('入力済みにする','entered');actionButton('入力結果が不明','unknown');}
 showRemarks(node,order.remarks);const receiptButton=element('button','購入控えを表示','secondary');receiptButton.onclick=()=>showReceipt(order.purchase_receipt||order);node.append(receiptButton);await renderResolution(node,order);$('admin-orders').append(node);}}
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
 // カタログ更新履歴は注文入力画面へ表示しない。旧HTMLでも安全に取り除く。
 $('campaign-notice')?.remove();
 // 固定ヘッダーの実寸に合わせ、スクロール先がメーカー名の下へ隠れないようにする。
 const header=$('app-header');const fitHeader=()=>document.documentElement.style.setProperty('--header-height',`${header.getBoundingClientRect().height+12}px`);new ResizeObserver(fitHeader).observe(header);fitHeader();
 // 更新中に旧HTMLが残っていても、注文画面の起動を止めない。
 const makerHelp=$('maker-help');if(makerHelp)makerHelp.hidden=officeMode;
 $('version').textContent=APP_VERSION;if(!localDemo){$('demo').hidden=true;if(!officeMode){$('token-label').hidden=true;$('connect').hidden=true;$('login-title').textContent='メーカー用QRを読み取ってください';}}
 const network=()=>{$('network').textContent=navigator.onLine?'オンライン':'オフライン';};network();window.addEventListener('online',()=>{network();showConnection();pump().catch(e=>message(errorText(e)));});window.addEventListener('offline',()=>{network();showConnection();});
 await sanitizeLegacyQr();
 const join=new URL(location.href).hash.match(/^#join=([-_A-Za-z0-9]{43,128})$/)?.[1];
 if(location.hash)history.replaceState(null,'',location.pathname+location.search);
 if(officeMode){$('login-title').textContent='事務画面';$('token-label').hidden=!localDemo;$('connect').textContent='社内ポータルのログインで接続';$('demo').hidden=true;if(!localDemo&&!portalSession()){const link=element('a','社内ポータルでログインする');link.href=PORTAL_URL;$('login').append(link);message('社内ポータルへログインしてから、この画面に戻ってください。');}else if(!localDemo){const value=await api('catalog');await activate(value);message('事務画面に接続しました。');}}else{const remembered=await get('maker-connection');makerToken=localDemo?(session.getItem('booth-token')||(remembered?.device===deviceId()?remembered.token:null)):(remembered?.device===deviceId()?remembered.token:remembered===undefined?session.getItem('booth-token'):null);if(makerToken)session.setItem('booth-token',makerToken);else if(!localDemo)session.removeItem('booth-token');const saved=await get('maker-catalog');if(saved){await activate(saved);message('注文画面を準備しています…');}if(join){await enroll(join);}else if(makerToken){try{const value=await api('catalog');if(!localDemo&&value.scope.device!==deviceId())throw Error('SCOPE_MISMATCH');connectionVerified=true;await activate(value);if(!localDemo)await set('maker-connection',{token:makerToken,device:deviceId()});message(readyMessage(value));await pump();}catch(e){showConnection();message(errorText(e));}}else if(saved){showConnection();message('メーカーQRを読み直すと、お客様検索・送信・共有履歴が使えます。未送信注文は端末に残っています。');}}
}
async function prepareShell(){
 if(!('serviceWorker' in navigator))return;
 try{
  const registration=await navigator.serviceWorker.register('./sw.js',{updateViaCache:'none'});
  await registration.update();await navigator.serviceWorker.ready;
  $('prepared').textContent='アプリ本体のオフライン準備ができました。商品情報は接続時に保存されます。';
 }catch{
  // オフラインでも保存済み画面と注文を利用できる。更新の失敗で注文処理を止めない。
 }
}
void prepareShell();
start().catch(e=>message(errorText(e)));
