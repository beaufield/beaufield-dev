// 社員認証で取得した注文だけを端末内でPDF化する。宛先・注文を外部サービスへ送らない。
import {unpackRemarks,billingLabel,deliveryLabel,handlingWarnings,displayProductCode} from '../order/export.js?build=20261025-atfit';
export function prepareReport(data,makerId=''){
 if(!data||!Array.isArray(data.rows)||!['practice','live'].includes(data.phase))throw Error('REPORT_DATA_MISMATCH');
 const seen=new Set(),rows=data.rows.filter(r=>r.state==='active');let total=0;
 for(const r of rows){
  const key=r.booth+':'+r.request_id;
  if(!r.request_id||seen.has(key)||makerId&&r.booth!==makerId||!Array.isArray(r.lines)||!r.customer||!Number.isSafeInteger(r.total_yen)||r.total_yen<0)throw Error('REPORT_DATA_MISMATCH');
  seen.add(key);let subtotal=0;
  for(const line of r.lines){if(!['paid','gift'].includes(line.kind)||!Number.isSafeInteger(line.qty)||line.qty<1||!Number.isSafeInteger(line.unit_price_yen)||line.unit_price_yen<0||line.kind==='gift'&&line.unit_price_yen!==0)throw Error('REPORT_DATA_MISMATCH');subtotal+=line.qty*line.unit_price_yen;}
  if(subtotal!==r.total_yen||!Number.isSafeInteger(subtotal))throw Error('REPORT_DATA_MISMATCH');total+=subtotal;
 }
 if(!Number.isSafeInteger(total)||data.order_count!==rows.length||data.total_yen!==total)throw Error('REPORT_DATA_MISMATCH');
 // 全メーカー・個社とも注文日時の新しい順。同時刻は注文番号で安定させ、日時不明は末尾に置く。
 const time=value=>{const n=Date.parse(value);return Number.isFinite(n)?n:-Infinity;};
 rows.sort((a,b)=>time(b.received_at)-time(a.received_at)||a.request_id.localeCompare(b.request_id));
 return {...data,rows,total_yen:total,customer_count:new Set(rows.map(r=>r.customer.id).filter(Boolean)).size};
}
const yen=value=>Number(value).toLocaleString('ja-JP');
const when=value=>{const d=new Date(value);return Number.isFinite(d.getTime())?d.toLocaleString('ja-JP',{timeZone:'Asia/Tokyo',hour12:false}):'時刻未確認';};
export function orderWarnings(order){
 const h=unpackRemarks(order.remarks||'');
 return [...new Set([...handlingWarnings(order),!order.customer.customer_code?'得意先コードを確認':null,h.invalid?'受付情報の形式を確認':null].filter(Boolean))];
}
// 商品別集計は確定済みの明細を合算する。現在のカタログで価格を再計算しない。
export function prepareProductSummary(input,makerId=''){
 const data=prepareReport(input,makerId),makers=new Map();
 for(const order of data.rows){
  if(!makers.has(order.booth))makers.set(order.booth,{id:order.booth,name:order.manufacturer||order.booth,orders:0,totalYen:0,products:new Map()});
  const maker=makers.get(order.booth);maker.orders++;maker.totalYen+=order.total_yen;
  for(const line of order.lines){
   const code=String(line.product_code??''),name=String(line.product_name??''),unit=String(line.unit||'点');
   if(!code||!name)throw Error('REPORT_DATA_MISMATCH');
   const key=JSON.stringify([code,name,unit]);
   if(!maker.products.has(key))maker.products.set(key,{code,name,unit,paidQty:0,giftQty:0,totalQty:0,totalYen:0});
   const product=maker.products.get(key);product[line.kind==='gift'?'giftQty':'paidQty']+=line.qty;product.totalQty+=line.qty;product.totalYen+=line.qty*line.unit_price_yen;
  }
 }
 const result=[...makers.values()].sort((a,b)=>a.name.localeCompare(b.name,'ja')||a.id.localeCompare(b.id)).map(m=>({...m,products:[...m.products.values()].sort((a,b)=>a.code.localeCompare(b.code)||a.name.localeCompare(b.name,'ja')||a.unit.localeCompare(b.unit,'ja'))}));
 for(const maker of result)if(maker.products.reduce((n,p)=>n+p.totalYen,0)!==maker.totalYen)throw Error('REPORT_DATA_MISMATCH');
 if(result.reduce((n,m)=>n+m.totalYen,0)!==data.total_yen)throw Error('REPORT_DATA_MISMATCH');
 return {makers:result,productCount:result.reduce((n,m)=>n+m.products.length,0),totalYen:data.total_yen};
}
export async function buildOrderPdf(input,{pdfLib,fontkit,fontBytes,glyphImages={},makerId='',title='全メーカー'}={}){
 const data=prepareReport(input,makerId);
 const summary=prepareProductSummary(data,makerId);
 if(!pdfLib?.PDFDocument||!fontkit||!fontBytes)throw Error('PDF_DEPENDENCY_UNAVAILABLE');
 const {PDFDocument,rgb}=pdfLib,doc=await PDFDocument.create();doc.registerFontkit(fontkit);
 // 日本語フォントの一部だけの埋込みで文字が欠けるビューアーを避け、全字形を埋め込む。
 const font=await doc.embedFont(fontBytes,{subset:false});
 // 日本語フォントにない絵文字も、元の表記を削らず端末内で描いた画像として埋め込む。
 const inlineImages=new Map();for(const [character,png] of Object.entries(glyphImages))inlineImages.set(character,await doc.embedPng(png));
 const glyphs=new Set(font.getCharacterSet());
 const clean=value=>{const s=String(value??'').replace(/[\u0000-\u0008\u000b-\u001f\u007f]/g,' ').normalize('NFC');if([...s].some(c=>c!=='\n'&&!glyphs.has(c.codePointAt(0))&&!inlineImages.has(c)))throw Error('PDF_FONT_UNSUPPORTED');return s;};
 const W=595.28,H=841.89,M=36,WIDTH=W-M*2,BOTTOM=45;
 const navy=rgb(.09,.25,.35),gray=rgb(.35,.4,.44),blue=rgb(.91,.95,.98),light=rgb(.96,.97,.98),warn=rgb(.58,.25,.07);
 let page,y,context='';const pages=[];
 const parts=text=>{const result=[];let chunk='';for(const ch of text){if(inlineImages.has(ch)){if(chunk)result.push({text:chunk});chunk='';result.push({image:inlineImages.get(ch)});}else chunk+=ch;}if(chunk)result.push({text:chunk});return result;};
 const textWidth=(text,size)=>parts(text).reduce((n,p)=>n+(p.image?size:font.widthOfTextAtSize(p.text,size)),0);
 const wrap=(value,width,size)=>{const lines=[];for(const para of clean(value).split('\n')){let line='';for(const ch of para){if(line&&textWidth(line+ch,size)>width){lines.push(line);line='';}line+=ch;}lines.push(line);}return lines;};
 function draw(text,x,at,size=9,color=navy){let left=x;for(const part of parts(clean(text))){if(part.image){page.drawImage(part.image,{x:left,y:at-size*.15,width:size,height:size});left+=size;}else{page.drawText(part.text,{x:left,y:at,size,font,color});left+=font.widthOfTextAtSize(part.text,size);}}}
 function newPage(){page=doc.addPage([W,H]);pages.push(page);draw('BEAU Fes 2026 注文データ',M,H-42,16);draw(`${data.phase==='practice'?'【練習データ・売上入力しない】':'【本番データ】'} ${title}`,M,H-62,10,data.phase==='practice'?warn:navy);draw(`取得 ${when(data.fetched_at)} ／ 金額は税抜・取消済みを除く`,M,H-78,8,gray);y=H-100;if(context){for(const line of wrap(context+'（続き）',WIDTH,9)){draw(line,M,y,9);y-=13;}y-=6;}}
 function ensure(height){if(y-height<BOTTOM)newPage();}
 function paragraph(text,{size=9,color=navy,gap=4}={}){for(const line of wrap(text,WIDTH,size)){ensure(size+7);draw(line,M,y,size,color);y-=size+5;}y-=gap;}
 function rule(){ensure(10);page.drawLine({start:{x:M,y:y+3},end:{x:W-M,y:y+3},thickness:.6,color:rgb(.75,.8,.83)});y-=10;}
 function tableHeader(){ensure(28);page.drawRectangle({x:M,y:y-19,width:WIDTH,height:24,color:blue});let x=M;const cols=[52,210,40,32,79,WIDTH-413];for(const [i,label]of ['商品コード','商品名・企画','区分','数量','税抜単価','税抜金額'].entries()){draw(label,x+4,y-10,8);x+=cols[i];}y-=27;return cols;}
 function lineRow(line,cols,index){
  const pending=/^NAPLA26-NOTE-/.test(line.product_code),name=pending?String(line.product_name||'').replace('（6本・明細はお客様発注待ち）','（6本・5＋1）'):String(line.product_name||''),cells=[pending?'-':displayProductCode(line.product_code),name,line.kind==='gift'?'サービス':'購入',String(line.qty),yen(line.unit_price_yen),yen(line.qty*line.unit_price_yen)];
  const lines=cells.map((s,i)=>wrap(s,cols[i]-8,8.5)),height=Math.max(...lines.map(a=>a.length))*12+10;
  if(y-height<BOTTOM){newPage();tableHeader();}
  if(height>H-180)throw Error('PDF_ROW_TOO_LONG');
  if(index%2===0)page.drawRectangle({x:M,y:y-height+7,width:WIDTH,height,color:light});
  let x=M;for(let i=0;i<cells.length;i++){let at=y-5;for(const text of lines[i]){const right=i>=3?Math.max(4,cols[i]-5-font.widthOfTextAtSize(text,8.5)):4;draw(text,x+right,at,8.5,line.kind==='gift'?rgb(.15,.36,.26):navy);at-=12;}x+=cols[i];}y-=height;
 }
 newPage();paragraph(`${data.order_count}件 ／ ${data.customer_count}人 ／ 税抜合計 ${yen(data.total_yen)}円`,{size:12});
 const counts=new Map();for(const r of data.rows){const name=r.manufacturer||r.booth;const c=counts.get(name)||{n:0,total:0};c.n++;c.total+=r.total_yen;counts.set(name,c);}
 for(const [name,c] of counts)paragraph(`${name}：${c.n}件 ／ ${yen(c.total)}円`,{size:9});
 if(!data.rows.length)paragraph('対象の有効注文はありません。',{size:12});
 rule();
 for(const r of data.rows){
  const h=unpackRemarks(r.remarks||''),warnings=orderWarnings(r),customer=r.customer;context='';
  ensure(210);context=`${r.manufacturer||r.booth} ／ 注文 ${r.request_id}`;
  paragraph(r.manufacturer||r.booth,{size:13});paragraph(`注文番号：${r.request_id}`,{size:8});
  paragraph(`注文日時：${when(r.received_at)} ／ 来場者番号：${customer.id||'未確認'}`,{size:8});
  paragraph(`得意先コード：${h.billing==='salon'?(customer.customer_code||'要確認'):'事務確認（個人請求・未確定）'} ／ サロン：${customer.salon||'未登録'}`,{size:10});
  paragraph(`お客様：${customer.name||'未登録'} ／ 請求先：${billingLabel(h.billing)} ／ 受渡：${deliveryLabel(r.delivery)}`,{size:9});
  for(const text of warnings)paragraph('要確認：'+text,{size:9,color:warn});
  const cols=tableHeader();r.lines.forEach((line,i)=>lineRow(line,cols,i));
  paragraph(`注文合計（税抜）：${yen(r.total_yen)}円`,{size:10});
  if(h.remarks)paragraph('備考：'+h.remarks,{size:9});
  if(h.contact!=='registered')paragraph('配送先・一部お渡し等は備考と受付情報を確認してください。',{size:8,color:warn});
  context='';rule();
 }
 // 注文単位の明細とは別のページ・色・見出しで、最後に商品合算の一覧を付ける。
 const summaryStartPage=pages.length+1,green=rgb(.10,.32,.25),paleGreen=rgb(.91,.96,.92),summaryStripe=rgb(.96,.98,.96);
 const summaryCols=[58,213,28,42,48,42,WIDTH-431];
 function summaryPage(){
  page=doc.addPage([W,H]);pages.push(page);context='';
  page.drawRectangle({x:M,y:H-82,width:WIDTH,height:46,color:paleGreen});
  page.drawRectangle({x:M,y:H-82,width:5,height:46,color:green});
  draw('商品別集計',M+14,H-54,18,green);draw('全注文を商品ごとに合算した一覧です',M+14,H-71,8,gray);
  draw(`${data.phase==='practice'?'【練習データ・売上入力しない】':'【本番データ】'} ${title}`,M,H-101,10,data.phase==='practice'?warn:green);
  draw(`取得 ${when(data.fetched_at)} ／ 取消済みを除く・金額は税抜`,M,H-116,8,gray);
  draw(`有効注文 ${data.order_count}件 ／ 商品 ${summary.productCount}種類 ／ 税抜合計 ${yen(summary.totalYen)}円`,M,H-136,10,green);
  draw('購入とサービスを区別。数量は本・個・セットなど、登録された単位のままです。',M,H-152,8,gray);y=H-173;
 }
 function summaryHeader(){
  page.drawRectangle({x:M,y:y-27,width:WIDTH,height:34,color:green});let x=M;
  for(const [i,label]of ['商品コード','商品名','単位','購入数量','サービス\n数量','合計数量','税抜金額'].entries()){
   let at=y-7;for(const text of wrap(label,summaryCols[i]-8,8)){draw(text,x+4,at,8,rgb(1,1,1));at-=11;}x+=summaryCols[i];
  }y-=35;
 }
 function summaryMaker(maker,continued=false){
  for(const text of wrap(maker.name+(continued?'（続き）':''),WIDTH,12)){draw(text,M,y,12,green);y-=16;}y-=2;
  draw(`${maker.orders}注文 ／ ${maker.products.length}商品 ／ メーカー税抜合計 ${yen(maker.totalYen)}円`,M,y,8,gray);y-=16;summaryHeader();
 }
 function productLayout(product){
  const pending=/^NAPLA26-NOTE-/.test(product.code),name=pending?product.name.replace('（6本・明細はお客様発注待ち）','（6本・5＋1）'):product.name;
  const cells=[pending?'-':displayProductCode(product.code),name,product.unit,String(product.paidQty),String(product.giftQty),String(product.totalQty),yen(product.totalYen)];
  const rows=cells.map((value,i)=>wrap(value,summaryCols[i]-8,8.5));return {cells,rows,height:Math.max(...rows.map(r=>r.length))*12+10};
 }
 summaryPage();
 if(!summary.makers.length){draw('対象の有効注文はありません。',M,y,11,green);y-=28;}
 for(let mi=0;mi<summary.makers.length;mi++){
  const maker=summary.makers[mi];
  const first=productLayout(maker.products[0]),headingHeight=wrap(maker.name,WIDTH,12).length*16+53;
  const firstReserve=first.height+(maker.products.length===1?34:0)+(maker.products.length===1&&mi===summary.makers.length-1?46:0);
  if(y-headingHeight-firstReserve<BOTTOM)summaryPage();summaryMaker(maker);
  for(let index=0;index<maker.products.length;index++){
   const {cells,rows,height}=productLayout(maker.products[index]);
   const last=index===maker.products.length-1,reserve=height+(last?34:0)+(last&&mi===summary.makers.length-1?46:0);
   if(reserve+wrap(maker.name+'（続き）',WIDTH,12).length*16+53>H-173-BOTTOM)throw Error('PDF_ROW_TOO_LONG');
   if(y-reserve<BOTTOM){summaryPage();summaryMaker(maker,true);}
   if(index%2===0)page.drawRectangle({x:M,y:y-height+7,width:WIDTH,height,color:summaryStripe});
   let x=M;
   for(let i=0;i<cells.length;i++){
    let at=y-5;for(const text of rows[i]){const dx=i>=3?Math.max(4,summaryCols[i]-5-font.widthOfTextAtSize(text,8.5)):4;draw(text,x+dx,at,8.5,i===4?rgb(.23,.43,.32):green);at-=12;}x+=summaryCols[i];
   }y-=height;
  }
  page.drawRectangle({x:M,y:y-19,width:WIDTH,height:25,color:paleGreen});
  draw('メーカー合計（税抜）',M+8,y-11,9,green);const amount=yen(maker.totalYen)+'円';draw(amount,W-M-8-font.widthOfTextAtSize(amount,10),y-11,10,green);y-=39;
 }
 if(y-38<BOTTOM)summaryPage();
 page.drawRectangle({x:M,y:y-26,width:WIDTH,height:33,color:green});draw('商品別集計 合計（税抜）',M+10,y-14,11,rgb(1,1,1));
 const grandTotal=yen(summary.totalYen)+'円';draw(grandTotal,W-M-10-font.widthOfTextAtSize(grandTotal,12),y-14,12,rgb(1,1,1));
 pages.forEach((p,i)=>{p.drawText(`${i+1} / ${pages.length}`,{x:W-M-60,y:22,size:8,font,color:gray});p.drawText(i+1>=summaryStartPage?'商品別集計 / BEAU Fes 注文データ':data.phase==='practice'?'練習データ / BEAU Fes 注文データ':'BEAU Fes 注文データ / 出力時点の注文',{x:M,y:22,size:8,font,color:i+1>=summaryStartPage?green:gray});});
 doc.setTitle('BEAU Fes 注文データ');doc.setCreator('BEAU Fes 注文PDF出力');doc.setLanguage('ja-JP');
 return {bytes:await doc.save(),pages:pages.length,orderCount:data.order_count,totalYen:data.total_yen,summaryStartPage,summaryPageCount:pages.length-summaryStartPage+1,productCount:summary.productCount};
}
