// 社員認証で取得した注文だけを端末内でPDF化する。宛先・注文を外部サービスへ送らない。
import {unpackRemarks,billingLabel,deliveryLabel,handlingWarnings,displayProductCode} from '../order/export.js';
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
 rows.sort((a,b)=>String(a.manufacturer||a.booth).localeCompare(String(b.manufacturer||b.booth),'ja')||String(a.customer.customer_code||'~').localeCompare(String(b.customer.customer_code||'~'))||String(a.received_at).localeCompare(String(b.received_at))||a.request_id.localeCompare(b.request_id));
 return {...data,rows,total_yen:total,customer_count:new Set(rows.map(r=>r.customer.id).filter(Boolean)).size};
}
const yen=value=>Number(value).toLocaleString('ja-JP');
const when=value=>{const d=new Date(value);return Number.isFinite(d.getTime())?d.toLocaleString('ja-JP',{timeZone:'Asia/Tokyo',hour12:false}):'時刻未確認';};
const states={ready:'未入力',claimed:'入力中・再入力禁止',entered:'入力済み・再入力禁止',unknown:'入力結果不明・再入力禁止',correction_required:'訂正確認',resolved_cancelled:'取消照合済み'};
export function orderWarnings(order){
 const h=unpackRemarks(order.remarks||'');
 return [...new Set([...handlingWarnings(order),!order.customer.customer_code?'得意先コードを確認':null,h.invalid?'受付情報の形式を確認':null,order.posting?.state!=='ready'?`台帳：${states[order.posting?.state]||'要確認'}。重複入力しない`:null].filter(Boolean))];
}
export async function buildOrderPdf(input,{pdfLib,fontkit,fontBytes,makerId='',title='全メーカー'}={}){
 const data=prepareReport(input,makerId);
 if(!pdfLib?.PDFDocument||!fontkit||!fontBytes)throw Error('PDF_DEPENDENCY_UNAVAILABLE');
 const {PDFDocument,rgb}=pdfLib,doc=await PDFDocument.create();doc.registerFontkit(fontkit);
 // 日本語フォントの一部だけの埋込みで文字が欠けるビューアーを避け、全字形を埋め込む。
 const font=await doc.embedFont(fontBytes,{subset:false});
 const glyphs=new Set(font.getCharacterSet());
 const clean=value=>{const s=String(value??'').replace(/[\u0000-\u0008\u000b-\u001f\u007f]/g,' ').normalize('NFC');if([...s].some(c=>c!=='\n'&&!glyphs.has(c.codePointAt(0))))throw Error('PDF_FONT_UNSUPPORTED');return s;};
 const W=595.28,H=841.89,M=36,WIDTH=W-M*2,BOTTOM=45;
 const navy=rgb(.09,.25,.35),gray=rgb(.35,.4,.44),blue=rgb(.91,.95,.98),light=rgb(.96,.97,.98),warn=rgb(.58,.25,.07);
 let page,y,context='';const pages=[];
 const wrap=(value,width,size)=>{const lines=[];for(const para of clean(value).split('\n')){let line='';for(const ch of para){if(line&&font.widthOfTextAtSize(line+ch,size)>width){lines.push(line);line='';}line+=ch;}lines.push(line);}return lines;};
 function draw(text,x,at,size=9,color=navy){page.drawText(clean(text),{x,y:at,size,font,color});}
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
 paragraph('社員向け・販売管理入力の確認用です。PDFの保存では売上入力済みになりません。売上日は入力時に確認してください。',{size:8,color:gray});
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
  paragraph(`売上日：${r.posting?.sales_date||'入力時に確認'} ／ 台帳：${states[r.posting?.state]||'要確認'} ／ 伝票番号：${r.posting?.slip_number||'未登録'}`,{size:8,color:gray});
  for(const text of warnings)paragraph('要確認：'+text,{size:9,color:warn});
  const cols=tableHeader();r.lines.forEach((line,i)=>lineRow(line,cols,i));
  paragraph(`注文合計（税抜）：${yen(r.total_yen)}円`,{size:10});
  if(h.remarks)paragraph('備考：'+h.remarks,{size:9});
  if(h.contact!=='registered')paragraph('配送先・一部お渡し等は備考と受付情報を確認してください。',{size:8,color:warn});
  context='';rule();
 }
 pages.forEach((p,i)=>{p.drawText(`${i+1} / ${pages.length}`,{x:W-M-60,y:22,size:8,font,color:gray});p.drawText(data.phase==='practice'?'練習データ / BEAU Fes 社員用':'BEAU Fes 社員用 / 出力時点の注文',{x:M,y:22,size:8,font,color:gray});});
 doc.setTitle('BEAU Fes 注文データ');doc.setCreator('BEAU Fes 社員用PDF出力');doc.setLanguage('ja-JP');
 return {bytes:await doc.save(),pages:pages.length,orderCount:data.order_count,totalYen:data.total_yen};
}
