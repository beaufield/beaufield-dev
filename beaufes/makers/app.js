// 業務データは社員認証後にだけ取得する。接続キーや価格をブラウザーの永続領域へ保存しない。
const APP_VERSION='v0.2.0';
const API_URL=document.querySelector('meta[name=directory-api]').content;
const ORDER_URL='https://beaufield.github.io/beaufield-dev/beaufes/order/';
const $=id=>document.getElementById(id);
const create=(tag,text)=>{const el=document.createElement(tag);if(text!==undefined)el.textContent=text;return el;};
const yen=value=>value==null?'単価確認待ち':`${Number(value).toLocaleString('ja-JP')}円`;
let makers=[],selected='',detailsSequence=0,qrSequence=0,qrBusy=false,currentQr=null;
const issued=new Map();
function session(){try{const s=JSON.parse(localStorage.getItem('bf_session')||'null');return s&&typeof s.token==='string'&&s.expires>Date.now()?s:null;}catch{return null;}}
function closeQr(){qrSequence++;currentQr=null;$('order-url').value='';$('open-order').removeAttribute('href');$('qr-box').replaceChildren();if($('qr-modal').open)$('qr-modal').close();}
function lock(){detailsSequence++;issued.clear();closeQr();makers=[];selected='';$('maker-select').replaceChildren(create('option','一覧から選んでください'));$('makers').replaceChildren();$('offers').replaceChildren();$('maker-name').textContent='';$('workspace').hidden=true;$('gate').hidden=false;}
const errors={PORTAL_SESSION_INVALID:'社員用ポータルでログインし直してください。',BEAUFES_ROLE_REQUIRED:'ビューフェスの利用権限がありません。管理者に確認してください。',PORTAL_UNAVAILABLE:'社員ログインを確認できません。時間をおいて再試行してください。',ORDER_APP_NOT_READY:'このメーカーの注文アプリは準備中です。',EVENT_CLOSED:'注文受付は終了しています。',ENROLLMENT_CONFLICT:'このQRを発行できません。一覧を更新して再試行してください。',DIRECTORY_UNAVAILABLE:'メーカー情報を取得できません。時間をおいて再試行してください。'};
function report(e){$('status').textContent=errors[e.message]||'処理を完了できませんでした。一覧を更新して再試行してください。';}
async function api(action,body={}){
 const s=session();if(!s){lock();throw Error('PORTAL_SESSION_INVALID');}
 const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),15000);
 try{const response=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'application/json',Authorization:`Bearer ${s.token}`},body:JSON.stringify({action,body}),cache:'no-store',signal:controller.signal});const data=await response.json();
  if(!session()||session().token!==s.token){lock();throw Error('PORTAL_SESSION_INVALID');}
  if(!response.ok){if([401,403].includes(response.status))lock();throw Error(data.error||'DIRECTORY_UNAVAILABLE');}
  return data;
 }finally{clearTimeout(timer);}
}
function orderState(){const maker=makers.find(m=>m.id===selected);$('direct-order').disabled=$('order-qr').disabled=qrBusy||!maker?.order_ready;for(const button of document.querySelectorAll('[data-open-maker]'))button.disabled=qrBusy||!makers.find(m=>m.id===button.dataset.openMaker)?.order_ready;$('order-note').textContent=!maker?'メーカーを選んで注文アプリへ進めます。メーカーへの案内用QRも発行できます。':maker.order_ready?'「注文アプリを開く」で直接入力画面へ進めます。メーカーへの案内にはQRを使えます。':'注文アプリ準備中です。キャンペーン内容は確認できます。';}
async function load(){
 const seq=++detailsSequence;$('status').textContent='メーカー一覧を読み込んでいます…';
 const result=await api('list');if(seq!==detailsSequence)return;makers=result.makers;selected='';
 $('maker-select').replaceChildren();const placeholder=create('option','一覧から選んでください');placeholder.value='';$('maker-select').append(placeholder);$('makers').replaceChildren();$('makers').hidden=false;$('details').hidden=true;
 for(const maker of makers){const option=create('option',maker.name);option.value=maker.id;$('maker-select').append(option);const card=create('section');card.className='maker-card';card.append(create('h2',maker.name),create('p',`${maker.offer_count}企画`));const note=create('p',maker.order_ready?'注文アプリ接続可':'注文アプリ準備中');note.className='muted';card.append(note);const direct=create('button','注文アプリを開く');direct.dataset.openMaker=maker.id;direct.disabled=!maker.order_ready;direct.onclick=()=>openOrder(maker.id).catch(report);card.append(direct);const button=create('button','キャンペーンを確認');button.className='secondary';button.onclick=()=>choose(maker.id).catch(report);card.append(button);$('makers').append(card);}
 $('gate').hidden=true;$('workspace').hidden=false;$('status').textContent=`${makers.length}社のキャンペーンを確認できます。`;orderState();
}
async function choose(id){
 closeQr();selected=id;$('maker-select').value=id;orderState();const seq=++detailsSequence;
 $('offers').replaceChildren();$('details').hidden=true;if(!id){$('makers').hidden=false;return;}
 const detail=await api('details',{maker_id:id});if(seq!==detailsSequence||id!==selected)return;
 $('makers').hidden=true;$('maker-name').textContent=detail.name;$('details').hidden=false;
 if(detail.notice){const notice=create('p',detail.notice);notice.className='questions';$('offers').append(notice);}
 for(const offer of detail.offers){const block=create('article');block.className='offer';block.append(create('h3',offer.title),create('p',offer.condition));if(offer.questions?.length){const questions=create('ul');questions.className='questions';for(const q of offer.questions)questions.append(create('li',q));block.append(questions);}
  const table=create('table'),head=create('thead'),hr=create('tr');for(const title of ['商品コード','商品名・容量','種別','税抜単価'])hr.append(create('th',title));head.append(hr);table.append(head);const tbody=create('tbody');
  for(const product of offer.products){const row=create('tr');for(const [label,text] of [['商品コード',product.code||'商品確認待ち'],['商品名・容量',product.name],['種別',product.role],['税抜単価',offer.print_mode==='set'&&product.role==='購入'?'セット価格に含む':yen(product.event_price)]]){const td=create('td',text);td.dataset.label=label;if(label==='税抜単価')td.className='price';row.append(td);}tbody.append(row);}table.append(tbody);block.append(table);$('offers').append(block);
 }
 $('status').textContent=`${detail.name}のキャンペーン内容です。`;
}
function randomKey(){return btoa(String.fromCharCode(...crypto.getRandomValues(new Uint8Array(48)))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');}
// 直接リンクもQRも同じメーカー専用接続を使用する。通信再試行は同じキーで行う。
async function prepareOrder(maker){
 const entry=issued.get(maker.id)||{key:randomKey()};issued.set(maker.id,entry);
 const result=await api('issue_order_qr',{maker_id:maker.id,key:entry.key});
 if(result.registered!==true||result.maker_id!==maker.id)throw Error('DIRECTORY_UNAVAILABLE');
 return ORDER_URL+'#join='+entry.key;
}
async function openOrder(id=selected){
 const maker=makers.find(m=>m.id===id);if(!maker?.order_ready||qrBusy)return;
 closeQr();selected=id;$('maker-select').value=id;const seq=++qrSequence;qrBusy=true;orderState();$('status').textContent='注文アプリへ接続しています…';
 try{const url=await prepareOrder(maker);if(seq===qrSequence&&selected===maker.id)location.assign(url);}finally{qrBusy=false;orderState();}
}
async function showQr(){
 const maker=makers.find(m=>m.id===selected);if(!maker?.order_ready||qrBusy)return;
 const seq=++qrSequence;qrBusy=true;orderState();$('status').textContent='注文URL・QRを準備しています…';
 try{
  const url=await prepareOrder(maker);
  if(seq!==qrSequence||selected!==maker.id)return;
  const qr=qrcode(0,'M');qr.addData(url);qr.make();
  $('qr-box').innerHTML=qr.createSvgTag({cellSize:5,margin:20,scalable:true});$('qr-box').querySelector('svg').setAttribute('aria-label',`${maker.name}の注文用QR`);
  currentQr={url,name:maker.name,qr};$('qr-name').textContent=`${maker.name}｜注文用QR`;$('order-url').value=url;$('open-order').href=url;$('qr-status').textContent='';$('qr-modal').showModal();$('status').textContent='メーカー担当者へ注文用QRをご案内できます。';
 }finally{qrBusy=false;orderState();}
}
// QRのセルを直接描き、保存時の画像変換を不要にする。
async function saveQr(){
 if(!currentQr)return;const snapshot=currentQr,qr=snapshot.qr,n=qr.getModuleCount(),scale=16,quiet=4;
 const canvas=document.createElement('canvas');canvas.width=canvas.height=(n+quiet*2)*scale;
 const ctx=canvas.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,canvas.width,canvas.height);ctx.fillStyle='black';
 for(let y=0;y<n;y++)for(let x=0;x<n;x++)if(qr.isDark(y,x))ctx.fillRect((x+quiet)*scale,(y+quiet)*scale,scale,scale);
 const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));if(!blob||currentQr!==snapshot)return;
 const png=URL.createObjectURL(blob),a=create('a');a.href=png;a.download=`BEAUFes_${snapshot.name}_注文QR.png`;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(png),1000);
}
$('version').textContent=APP_VERSION;
$('direct-order').onclick=()=>openOrder().catch(report);
$('refresh').onclick=()=>{closeQr();load().catch(report);};$('maker-select').onchange=e=>choose(e.target.value).catch(report);$('back').onclick=()=>choose('').catch(report);$('order-qr').onclick=()=>showQr().catch(report);$('qr-close').onclick=closeQr;$('qr-modal').addEventListener('cancel',closeQr);
$('copy-url').onclick=async()=>{if(!currentQr)return;try{await navigator.clipboard.writeText(currentQr.url);$('qr-status').textContent='注文URLをコピーしました。';}catch{$('order-url').focus();$('order-url').select();$('qr-status').textContent='URLを選択しました。コピー操作でお渡しください。';}};
$('save-qr').onclick=()=>saveQr().catch(report);
window.addEventListener('storage',e=>{if(e.key==='bf_session'){lock();$('status').textContent='社員ログインが変更されました。一覧を更新してください。';}});
document.addEventListener('visibilitychange',()=>{if(!document.hidden&&!session())lock();});
setInterval(()=>{if(!session())lock();},10000);
if(session())load().catch(report);
