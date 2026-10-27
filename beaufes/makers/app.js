// 業務データは社員認証後にだけ取得する。接続キーや価格をブラウザーの永続領域へ保存しない。
const APP_VERSION='v0.8.1';
// 表示用だけ0埋めを除き、検索・送信・既存注文の識別子は維持する。
const displayProductCode=value=>/^AD26-ATFIT-/.test(String(value??''))?'未登録':String(value??'').replace(/^0+(?=\d+$)/,'');
const API_URL=document.querySelector('meta[name=directory-api]').content;
const ORDER_URL='https://beaufield.github.io/beaufield-dev/beaufes/order/refresh.html';
const $=id=>document.getElementById(id);
const create=(tag,text)=>{const el=document.createElement(tag);if(text!==undefined)el.textContent=text;return el;};
const yen=value=>value==null?'単価確認待ち':`${Number(value).toLocaleString('ja-JP')}円`;
let makers=[],selected='',detailsSequence=0,qrSequence=0,qrBusy=false,currentQr=null;
const issued=new Map();
function session(){try{const s=JSON.parse(localStorage.getItem('bf_session')||'null');return s&&typeof s.token==='string'&&s.expires>Date.now()?s:null;}catch{return null;}}
function closeQr(){qrSequence++;currentQr=null;$('order-url').value='';$('open-order').removeAttribute('href');$('qr-box').replaceChildren();if($('qr-modal').open)$('qr-modal').close();}
function lock(){detailsSequence++;$('excel-all').disabled=true;$('excel-maker').disabled=true;$('excel-status').textContent='';$('pdf-customers').disabled=true;$('pdf-maker').disabled=true;$('pdf-all').disabled=true;$('pdf-status').textContent='';$('order-overview').replaceChildren();$('order-overview').hidden=true;issued.clear();closeQr();makers=[];selected='';$('maker-select').replaceChildren(create('option','一覧から選んでください'));$('makers').replaceChildren();$('offers').replaceChildren();$('maker-name').textContent='';$('workspace').hidden=true;$('gate').hidden=false;}
const errors={EXPORT_TOO_LARGE:'注文件数が多いため、メーカーを選んで出力してください。',REPORT_DATA_MISMATCH:'取得した注文の集計が一致しません。再入力せず、時間をおいて注文データを再取得してください。',PDF_FONT_UNSUPPORTED:'PDFで表示できない文字が含まれています。注文内容を変えず、事務へご連絡ください。',PDF_DEPENDENCY_UNAVAILABLE:'PDF作成用のファイルを取得できません。オンラインで再試行してください。',PORTAL_SESSION_INVALID:'社員用ポータルでログインし直してください。',BEAUFES_ROLE_REQUIRED:'ビューフェスの利用権限がありません。管理者に確認してください。',PORTAL_UNAVAILABLE:'社員ログインを確認できません。時間をおいて再試行してください。',ORDER_APP_NOT_READY:'このメーカーの注文アプリは準備中です。',EVENT_CLOSED:'注文受付は終了しています。',ENROLLMENT_CONFLICT:'このQRを発行できません。一覧を更新して再試行してください。',DIRECTORY_UNAVAILABLE:'メーカー情報を取得できません。時間をおいて再試行してください。'};
function report(e){staleOverview();$('status').textContent=errors[e.message]||'処理を完了できませんでした。一覧を更新して再試行してください。';}
async function api(action,body={}){
 const s=session();if(!s){lock();throw Error('PORTAL_SESSION_INVALID');}
 const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),15000);
 try{const response=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'application/json',Authorization:`Bearer ${s.token}`},body:JSON.stringify({action,body}),cache:'no-store',signal:controller.signal});const data=await response.json();
  if(!session()||session().token!==s.token){lock();throw Error('PORTAL_SESSION_INVALID');}
  if(!response.ok){if([401,403].includes(response.status))lock();throw Error(data.error||'DIRECTORY_UNAVAILABLE');}
  return data;
 }finally{clearTimeout(timer);}
}
function orderState(){const maker=makers.find(m=>m.id===selected),pdfLocked=pdfBusy||!session()||$("workspace").hidden;$('excel-all').disabled=excelBusy||!session()||$("workspace").hidden;$('excel-maker').disabled=excelBusy||!session()||$("workspace").hidden||!maker;$('pdf-all').disabled=pdfLocked;$('pdf-customers').disabled=pdfLocked;$('pdf-maker').disabled=pdfLocked||!maker;for(const button of document.querySelectorAll('[data-pdf-maker]'))button.disabled=pdfLocked;$('direct-order').disabled=$('order-qr').disabled=qrBusy||!maker?.order_ready;for(const button of document.querySelectorAll('[data-open-maker]'))button.disabled=qrBusy||!makers.find(m=>m.id===button.dataset.openMaker)?.order_ready;$('order-note').textContent=!maker?'メーカーを選んで注文アプリへ進めます。メーカーへの案内用QRも発行できます。':maker.order_ready?'「注文アプリを開く」で直接入力画面へ進めます。メーカーへの案内にはQRを使えます。':maker.id==='bjc'?'BJCはお客様が用紙へ記入し、受付へ提出します。予約セットは営業へ個別にご相談ください。':'注文アプリ準備中です。キャンペーン内容は確認できます。';}
// 社員だけが取得できる全体/メーカー別の集計。端末の注文や表示件数は合算しない。
function summaryBlock(summary,title){
 const box=create('div');box.className='order-totals';
 if(!summary||!['total_yen','customer_count','order_count'].every(k=>Number.isSafeInteger(summary[k])&&summary[k]>=0)){
  box.append(create('p','集計を取得できませんでした。一覧を更新してください。'));return box;
 }
 const heading=create('p',title);heading.className='summary-title';box.append(heading);
 const values=create('div');values.className='summary-values';
 for(const [label,value]of [['注文金額（税抜）',yen(summary.total_yen)],['注文人数',`${summary.customer_count.toLocaleString('ja-JP')}人`]]){
  const item=create('div');item.append(create('span',label),create('strong',value));values.append(item);
 }
 const note=create('p',`登録済み${summary.order_count.toLocaleString('ja-JP')}件・取消済みを除く`);note.className='summary-note';box.append(values,note);const delivery=create('p','注文の外部通知は停止中です。注文明細は社員用画面からPDF・Excelで出力できます。');delivery.className='summary-note';box.append(delivery);return box;
}
function staleOverview(){
 const box=$('order-overview');if(box.hidden)return;box.classList.add('summary-stale');
 const note=box.querySelector('.overview-note');if(note&&!note.textContent.startsWith('前回取得の集計です。'))note.textContent='前回取得の集計です。'+note.textContent;
}
async function load(){
 const seq=++detailsSequence;staleOverview();$('status').textContent='メーカー一覧を読み込んでいます…';
 // 紙で受け付けるBJCは、社員用のアプリ一覧には表示しない。
 const result=await api('list');if(seq!==detailsSequence)return;makers=result.makers.filter(m=>m.id!=='bjc');selected='';
 const overview=$('order-overview');overview.replaceChildren(summaryBlock(result.order_summary,result.summary_phase==='practice'?'全メーカーの練習注文の集計':'全メーカーの注文集計'));overview.hidden=false;overview.classList.remove('summary-stale');
 const note=create('p','全体の注文人数は、複数メーカーで購入した方も1人です。同じ方の追加注文も重複して数えません。');note.className='overview-note';overview.append(note);
 if(result.summary_fetched_at){const at=create('p','集計取得 '+new Date(result.summary_fetched_at).toLocaleString('ja-JP'));at.className='summary-note';overview.append(at);}
 $('maker-select').replaceChildren();const placeholder=create('option','一覧から選んでください');placeholder.value='';$('maker-select').append(placeholder);$('makers').replaceChildren();$('makers').hidden=false;$('details').hidden=true;
 for(const maker of makers){const option=create('option',maker.name);option.value=maker.id;$('maker-select').append(option);const card=create('section');card.className='maker-card';card.append(create('h2',maker.name),create('p',`${maker.offer_count}企画`));const note=create('p',maker.order_ready?'注文アプリ接続可':maker.id==='bjc'?'お客様が紙で記入':'注文アプリ準備中');note.className='muted';card.append(note);card.append(summaryBlock(maker.order_summary,'このメーカーの注文集計'));const direct=create('button','注文アプリを開く');direct.dataset.openMaker=maker.id;direct.disabled=!maker.order_ready;direct.onclick=()=>openOrder(maker.id).catch(report);card.append(direct);const button=create('button','キャンペーンを確認');button.className='secondary';button.onclick=()=>choose(maker.id).catch(report);card.append(button);const pdf=create('button','このメーカーの注文をPDF出力');pdf.className='secondary';pdf.dataset.pdfMaker=maker.id;pdf.onclick=()=>exportPdf(maker.id).catch(report);card.append(pdf);$('makers').append(card);}
 $('gate').hidden=true;$('workspace').hidden=false;$('status').textContent=`${makers.length}社のキャンペーンを確認できます。`;orderState();
}
async function choose(id){
 closeQr();selected=id;$('maker-select').value=id;orderState();const seq=++detailsSequence;
 $('offers').replaceChildren();$('details').hidden=true;if(!id){$('makers').hidden=false;return;}
 const detail=await api('details',{maker_id:id});if(seq!==detailsSequence||id!==selected)return;
 $('makers').hidden=true;$('maker-name').textContent=detail.name;$('details').hidden=false;
 $('offers').append(summaryBlock(makers.find(m=>m.id===id)?.order_summary,'このメーカーの注文集計（一覧取得時点）'));
 if(detail.notice){const notice=create('p',detail.notice);notice.className='questions';$('offers').append(notice);}
 for(const offer of detail.offers){const block=create('article');block.className='offer';block.append(create('h3',offer.title),create('p',offer.condition));if(offer.questions?.length){const questions=create('ul');questions.className='questions';for(const q of offer.questions)questions.append(create('li',q));block.append(questions);}
  const table=create('table'),head=create('thead'),hr=create('tr');for(const title of ['商品コード','商品名・容量','種別','税抜単価'])hr.append(create('th',title));head.append(hr);table.append(head);const tbody=create('tbody');
  for(const product of offer.products){const row=create('tr');for(const [label,text] of [['商品コード',product.code?displayProductCode(product.code):'商品確認待ち'],['商品名・容量',product.name],['種別',product.role],['税抜単価',offer.print_mode==='set'&&product.role==='購入'?'セット価格に含む':yen(product.event_price)]]){const td=create('td',text);td.dataset.label=label;if(label==='税抜単価')td.className='price';row.append(td);}tbody.append(row);}table.append(tbody);block.append(table);$('offers').append(block);
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
// PDFは依存ファイルも同じ公開元から読み、個票を第三者の変換サービスへ送らない。
let pdfBusy=false,pdfLibraries;
function loadScript(src){return new Promise((resolve,reject)=>{const script=document.createElement('script');script.src=src;script.onload=resolve;script.onerror=()=>{script.remove();reject(Error('PDF_DEPENDENCY_UNAVAILABLE'));};document.head.append(script);});}
async function dependencies(){
 if(!pdfLibraries)pdfLibraries=(async()=>{await Promise.all([window.PDFLib?null:loadScript('pdf-lib.min.js'),window.fontkit?null:loadScript('fontkit.umd.min.js')]);const response=await fetch('NotoSansJP-Regular.ttf',{cache:'force-cache'});if(!response.ok)throw Error('PDF_DEPENDENCY_UNAVAILABLE');const canvas=document.createElement("canvas");canvas.width=canvas.height=64;const context=canvas.getContext("2d");context.font='52px "Segoe UI Emoji","Apple Color Emoji","Noto Color Emoji",sans-serif';context.textAlign="center";context.textBaseline="middle";context.fillText("🌹",32,32);return {pdfLib:window.PDFLib,fontkit:window.fontkit,fontBytes:await response.arrayBuffer(),glyphImages:{"🌹":canvas.toDataURL("image/png")}};})().catch(e=>{pdfLibraries=null;throw e;});
 return pdfLibraries;
}
async function exportPdf(id='',view='orders'){
 if(pdfBusy)return;const auth=session();if(!auth){lock();throw Error('PORTAL_SESSION_INVALID');}
 const maker=id?makers.find(m=>m.id===id):null;if(id&&!maker)throw Error('DIRECTORY_UNAVAILABLE');
 pdfBusy=true;orderState();$('pdf-status').textContent='最新の注文を取得してPDFを作成しています…';
 try{
  const [data,libs,module]=await Promise.all([api('export_orders',id?{maker_id:id}:{}),dependencies(),import('./order-pdf.mjs?build=20261027-customer-boundaries')]);
  if(session()?.token!==auth.token){lock();throw Error('PORTAL_SESSION_INVALID');}
  const result=await (view==='customers'?module.buildCustomerPdf:module.buildOrderPdf)(data,{...libs,makerId:id,title:maker?.name||'全メーカー'});
  if(session()?.token!==auth.token){lock();throw Error('PORTAL_SESSION_INVALID');}
  const url=URL.createObjectURL(new Blob([result.bytes],{type:'application/pdf'})),link=create('a');
  link.href=url;link.download=`BEAUFes_${data.phase==='practice'?'練習_':''}${view==='customers'?'お客様別_営業担当順':id||'全メーカー'}_注文_${new Date().toISOString().replace(/[:.]/g,'-')}.pdf`;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
  $('pdf-status').textContent=`${view==='customers'?'お客様別（営業担当順）':maker?.name||'全メーカー'}：${result.orderCount}件・${result.pages}ページのPDFを保存しました（${view==='customers'?result.customerCount+'人のお客様別':'末尾に商品別集計'}・税抜 ${yen(result.totalYen)}）。`;
 }catch(error){$('pdf-status').textContent=errors[error.message]||'PDFを作成できませんでした。オンラインで再試行してください。';throw error;}
 finally{pdfBusy=false;orderState();}
}
// Excelも端末内で作成する。取得APIは既存PDFと同じ読取専用操作を使う。
let excelBusy=false,excelLibrary;
async function exportExcel(id=''){
 if(excelBusy)return;const auth=session();if(!auth){lock();throw Error('PORTAL_SESSION_INVALID');}
 const maker=id?makers.find(m=>m.id===id):null;if(id&&!maker)throw Error('DIRECTORY_UNAVAILABLE');
 excelBusy=true;orderState();$('excel-status').textContent='最新の注文からExcelを作成しています…';
 try{
  if(!excelLibrary)excelLibrary=(window.ExcelJS?Promise.resolve():loadScript('../order/exceljs.min.js')).catch(error=>{excelLibrary=null;throw error;});
  const [data,,module]=await Promise.all([api('export_orders',id?{maker_id:id}:{}),excelLibrary,import('./order-excel.mjs?build=20261027-customer-view')]);
  if(session()?.token!==auth.token){lock();throw Error('PORTAL_SESSION_INVALID');}
  const result=await module.buildOrderExcel(window.ExcelJS,data,{makerId:id,title:maker?.name||'全メーカー'});
  if(session()?.token!==auth.token){lock();throw Error('PORTAL_SESSION_INVALID');}
  const url=URL.createObjectURL(new Blob([result.bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})),link=create('a');
  link.href=url;link.download=`BEAUFes_${data.phase==='practice'?'練習_':''}${id||'全メーカー'}_注文集計_${new Date().toISOString().replace(/[:.]/g,'-')}.xlsx`;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
  $('excel-status').textContent=`${maker?.name||'全メーカー'}：注文${result.orderCount}件・明細${result.detailCount}行・商品${result.productCount}件・お客様${result.customerCount}人のExcelを保存しました（税抜 ${yen(result.totalYen)}）。`;
 }catch(error){$('excel-status').textContent=errors[error.message]||'Excelを作成できませんでした。オンラインで再試行してください。';throw error;}
 finally{excelBusy=false;orderState();}
}
$('version').textContent=APP_VERSION;
$('excel-all').onclick=()=>exportExcel().catch(report);$('excel-maker').onclick=()=>exportExcel(selected).catch(report);
$('pdf-all').onclick=()=>exportPdf().catch(report);$('pdf-customers').onclick=()=>exportPdf('','customers').catch(report);$('pdf-maker').onclick=()=>exportPdf(selected).catch(report);
$('direct-order').onclick=()=>openOrder().catch(report);
$('refresh').onclick=()=>{closeQr();load().catch(report);};$('maker-select').onchange=e=>choose(e.target.value).catch(report);$('back').onclick=()=>choose('').catch(report);$('order-qr').onclick=()=>showQr().catch(report);$('qr-close').onclick=closeQr;$('qr-modal').addEventListener('cancel',closeQr);
$('copy-url').onclick=async()=>{if(!currentQr)return;try{await navigator.clipboard.writeText(currentQr.url);$('qr-status').textContent='注文URLをコピーしました。';}catch{$('order-url').focus();$('order-url').select();$('qr-status').textContent='URLを選択しました。コピー操作でお渡しください。';}};
$('save-qr').onclick=()=>saveQr().catch(report);
window.addEventListener('storage',e=>{if(e.key==='bf_session'){lock();$('status').textContent='社員ログインが変更されました。一覧を更新してください。';}});
document.addEventListener('visibilitychange',()=>{if(!document.hidden&&!session())lock();});
setInterval(()=>{if(!session())lock();},10000);
if(session())load().catch(report);
