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
 // 全メーカー・個社とも注文日時の古い順。同時刻は注文番号で安定させ、日時不明は末尾に置く。
 const time=value=>{const n=Date.parse(value);return Number.isFinite(n)?n:Infinity;};
 rows.sort((a,b)=>time(a.received_at)-time(b.received_at)||a.request_id.localeCompare(b.request_id));
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
// 受渡区分を先に分け、同じ商品の購入・サービス・金額を区分の間で相殺しない。
export function prepareDeliveryProductSummary(input){
 if(input?.maker_id)throw Error('REPORT_DATA_MISMATCH');
 const data=prepareReport(input),buckets=[{key:'later',label:'後日納品',rows:[]},{key:'takeaway',label:'当日お渡し済み',rows:[]}];
 const pending={key:'unknown',label:'受渡方法未確認',rows:[]};
 for(const row of data.rows)(buckets.find(g=>g.key===row.delivery)||pending).rows.push(row);
 if(pending.rows.length)buckets.push(pending);
 const groups=buckets.map(group=>{
  const part={...data,rows:group.rows,order_count:group.rows.length,total_yen:group.rows.reduce((n,r)=>n+r.total_yen,0),customer_count:new Set(group.rows.map(r=>r.customer.id).filter(Boolean)).size};
  return {key:group.key,label:group.label,data:part,summary:prepareProductSummary(part)};
 });
 if(groups.reduce((n,g)=>n+g.data.order_count,0)!==data.order_count||groups.reduce((n,g)=>n+g.summary.totalYen,0)!==data.total_yen)throw Error('REPORT_DATA_MISMATCH');
 return {data,groups,totalYen:data.total_yen,productCount:groups.reduce((n,g)=>n+g.summary.productCount,0)};
}
// 同姓同名・共通の得意先番号は統合せず、来場者番号でお客様をまとめる。
export function prepareCustomerReport(input,makerId=''){
 const data=prepareReport(input,makerId),people=new Map();
 const stamp=value=>{const n=Date.parse(value);return Number.isFinite(n)?n:-Infinity;};
 for(const order of data.rows){
  const key=order.customer.id?'id:'+order.customer.id:'order:'+order.booth+':'+order.request_id;
  if(!people.has(key))people.set(key,{key,customer:order.customer,latest:-Infinity,rows:[],totalYen:0});
  const person=people.get(key);person.rows.push(order);person.totalYen+=order.total_yen;
  if(stamp(order.received_at)>=person.latest){person.customer=order.customer;person.latest=stamp(order.received_at);}
 }
 const name=value=>String(value||'').trim(),cmp=(a,b)=>name(a).localeCompare(name(b),'ja');
 const configured=[...new Set((Array.isArray(data.tantou_list)?data.tantou_list:[]).map(name).filter(Boolean))];
 const ranks=new Map(configured.map((v,i)=>[v,i])),rank=v=>v?(ranks.get(v)??configured.length):configured.length+1;
 const customers=[...people.values()].sort((a,b)=>{
  const left=name(a.customer.tantou),right=name(b.customer.tantou);
  return rank(left)-rank(right)||cmp(left,right)||cmp(a.customer.salon,b.customer.salon)||cmp(a.customer.name,b.customer.name)||cmp(a.key,b.key);
 });
 const groups=[];
 for(const person of customers){const owner=name(person.customer.tantou);let group=groups.at(-1);if(!group||group.name!==owner){group={name:owner,label:owner||'未割当',customers:[],orderCount:0,totalYen:0};groups.push(group);}group.customers.push(person);group.orderCount+=person.rows.length;group.totalYen+=person.totalYen;}
 if(customers.reduce((n,p)=>n+p.totalYen,0)!==data.total_yen)throw Error('REPORT_DATA_MISMATCH');
 return {data,customers,groups,customerCount:customers.length};
}
export async function buildCustomerPdf(input,options={}){return buildOrderPdf(input,{...options,view:'customers'});}
export async function buildDeliveryProductPdf(input,options={}){return buildOrderPdf(input,{...options,makerId:'',title:'全メーカー',view:'delivery-products'});}
export async function buildOrderPdf(input,{pdfLib,fontkit,fontBytes,glyphImages={},makerId='',title='全メーカー',view='orders'}={}){
 if(!['orders','customers','delivery-products'].includes(view))throw Error('REPORT_DATA_MISMATCH');
 const data=prepareReport(input,makerId),customerReport=view==='customers'?prepareCustomerReport(data,makerId):null;
 const summary=prepareProductSummary(data,makerId),deliverySummary=view==='delivery-products'?prepareDeliveryProductSummary(data):null;
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
 let page,y,context='',activeCustomer=null;const pages=[],pageCustomers=new Map(),pageDeliveries=new Map();
 const parts=text=>{const result=[];let chunk='';for(const ch of text){if(inlineImages.has(ch)){if(chunk)result.push({text:chunk});chunk='';result.push({image:inlineImages.get(ch)});}else chunk+=ch;}if(chunk)result.push({text:chunk});return result;};
 const textWidth=(text,size)=>parts(text).reduce((n,p)=>n+(p.image?size:font.widthOfTextAtSize(p.text,size)),0);
 const wrap=(value,width,size)=>{const lines=[];for(const para of clean(value).split('\n')){let line='';for(const ch of para){if(line&&textWidth(line+ch,size)>width){lines.push(line);line='';}line+=ch;}lines.push(line);}return lines;};
 function draw(text,x,at,size=9,color=navy){let left=x;for(const part of parts(clean(text))){if(part.image){page.drawImage(part.image,{x:left,y:at-size*.15,width:size,height:size});left+=size;}else{page.drawText(part.text,{x:left,y:at,size,font,color});left+=font.widthOfTextAtSize(part.text,size);}}}
 // 濃紺の大きな見出しを、そのお客様の全ページへ繰り返す。
 function customerHeader(continued){
  const section=activeCustomer,customer=section.person.customer,white=rgb(1,1,1),top=y+8;
  const names=wrap(customer.name||'お名前未登録',WIDTH-32,21),salons=wrap('サロン：'+(customer.salon||'未登録'),WIDTH-32,12);
  const height=48+names.length*27+salons.length*18;
  if(height>H-330)throw Error('PDF_ROW_TOO_LONG');
  page.drawRectangle({x:M-12,y:BOTTOM+12,width:4,height:top-BOTTOM-12,color:navy});
  page.drawRectangle({x:M,y:top-height,width:WIDTH,height,color:navy});
  draw(`お客様 ${section.number} / ${customerReport.customerCount}${continued?'（続き）':''}`,M+16,top-20,10,white);
  const rep='営業担当：'+section.group.label;draw(rep,W-M-16-textWidth(rep,9),top-20,9,white);
  let at=top-48;for(const name of names){draw(name,M+16,at,21,white);at-=27;}
  for(const salon of salons){draw(salon,M+16,at,12,white);at-=18;}
  y=top-height-20;
  paragraph(`来場者番号：${customer.id||'未確認'} ／ 得意先コード：${customer.customer_code||'要確認'}`,{size:9,color:gray});
  if(!continued)paragraph(`注文 ${section.person.rows.length}件 ／ 税抜合計 ${yen(section.person.totalYen)}円`,{size:10});
  y-=4;
 }
 function newPage(){
  page=doc.addPage([W,H]);pages.push(page);
  draw(view==='customers'?'お客様別 注文一覧（営業担当順）':'BEAU Fes 2026 注文データ',M,H-42,16);
  draw(`${data.phase==='practice'?'【練習データ・売上入力しない】':'【本番データ】'} ${title}`,M,H-62,10,data.phase==='practice'?warn:navy);
  draw(`取得 ${when(data.fetched_at)} ／ 金額は税抜・取消済みを除く`,M,H-78,8,gray);y=H-100;
  if(view==='customers'&&activeCustomer){const continued=activeCustomer.pages.length>0;activeCustomer.pages.push(page);pageCustomers.set(page,activeCustomer);customerHeader(continued);}
  if(context){for(const line of wrap(context+'（続き）',WIDTH,9)){draw(line,M,y,9);y-=13;}y-=6;}
 }
 function ensure(height){if(y-height<BOTTOM)newPage();}
 function paragraph(text,{size=9,color=navy,gap=4}={}){for(const line of wrap(text,WIDTH,size)){ensure(size+7);draw(line,M,y,size,color);y-=size+5;}y-=gap;}
 function rule(){ensure(10);page.drawLine({start:{x:M,y:y+3},end:{x:W-M,y:y+3},thickness:.6,color:rgb(.75,.8,.83)});y-=10;}
 function tableHeader(){ensure(28);page.drawRectangle({x:M,y:y-19,width:WIDTH,height:24,color:blue});let x=M;const cols=[52,210,40,32,79,WIDTH-413];for(const [i,label]of ['商品コード','商品名・企画','区分','数量','税抜単価','税抜金額'].entries()){draw(label,x+4,y-10,8);x+=cols[i];}y-=27;return cols;}
 function lineRow(line,cols,index,reserve=0){
  const pending=/^NAPLA26-NOTE-/.test(line.product_code),name=pending?String(line.product_name||'').replace('（6本・明細はお客様発注待ち）','（6本・5＋1）'):String(line.product_name||''),cells=[pending?'-':displayProductCode(line.product_code),name,line.kind==='gift'?'サービス':'購入',String(line.qty),yen(line.unit_price_yen),yen(line.qty*line.unit_price_yen)];
  const lines=cells.map((s,i)=>wrap(s,cols[i]-8,8.5)),height=Math.max(...lines.map(a=>a.length))*12+10;
  if(y-height-reserve<BOTTOM){newPage();tableHeader();}
  if(height>H-180)throw Error('PDF_ROW_TOO_LONG');
  if(index%2===0)page.drawRectangle({x:M,y:y-height+7,width:WIDTH,height,color:light});
  let x=M;for(let i=0;i<cells.length;i++){let at=y-5;for(const text of lines[i]){const right=i>=3?Math.max(4,cols[i]-5-font.widthOfTextAtSize(text,8.5)):4;draw(text,x+right,at,8.5,line.kind==='gift'?rgb(.15,.36,.26):navy);at-=12;}x+=cols[i];}y-=height;
 }
 function orderBody(r,{showCustomer=true,lastForCustomer=false}={}){
  const h=unpackRemarks(r.remarks||''),warnings=orderWarnings(r),customer=r.customer;
  ensure(showCustomer?225:lastForCustomer?220:155);context=`${r.manufacturer||r.booth} ／ 注文 ${r.request_id}`;
  if(!showCustomer)context=`営業担当：${String(customer.tantou||'').trim()||'未割当'} ／ ${customer.salon||''} ${customer.name||''} ／ 注文 ${r.request_id}`;
  paragraph(r.manufacturer||r.booth,{size:13});paragraph(`注文番号：${r.request_id}`,{size:8});
  paragraph(`注文日時：${when(r.received_at)} ／ 来場者番号：${customer.id||'未確認'}`,{size:8});
  if(showCustomer){
   paragraph(`得意先コード：${h.billing==='salon'?(customer.customer_code||'要確認'):'事務確認（個人請求・未確定）'} ／ サロン：${customer.salon||'未登録'}`,{size:10});
   paragraph(`お客様：${customer.name||'未登録'} ／ 営業担当：${String(customer.tantou||'').trim()||'未割当'}`,{size:9});
  }
  paragraph(`請求先：${billingLabel(h.billing)} ／ 受渡：${deliveryLabel(r.delivery)}`,{size:9});
  for(const text of warnings)paragraph('要確認：'+text,{size:9,color:warn});
  // お客様合計だけが次のページへ孤立しないよう、最後の明細に後続の表示領域も確保する。
  const closing=lastForCustomer?48+activeCustomer.endHeight+(h.remarks?wrap('備考：'+h.remarks,WIDTH,9).length*14+4:0)+(h.contact!=='registered'?wrap('配送先・一部お渡し等は備考と受付情報を確認してください。',WIDTH,8).length*13+4:0):0;
  const cols=tableHeader();r.lines.forEach((line,i)=>lineRow(line,cols,i,i===r.lines.length-1?closing:0));
  paragraph(`注文合計（税抜）：${yen(r.total_yen)}円`,{size:10});
  if(h.remarks)paragraph('備考：'+h.remarks,{size:9});
  if(h.contact!=='registered')paragraph('配送先・一部お渡し等は備考と受付情報を確認してください。',{size:8,color:warn});
  context='';rule();
 }
 if(view!=='delivery-products'){
 newPage();paragraph(`${data.order_count}件 ／ ${view==='customers'?customerReport.customerCount:data.customer_count}人 ／ 税抜合計 ${yen(data.total_yen)}円`,{size:12});
 if(view==='customers'){
  paragraph('営業担当ごとに、お客様の全メーカーの注文をまとめています。お客様の中は注文日時の古い順です。',{size:9,color:gray});
  paragraph('営業担当は出力時点のお客様情報。担当の表示順は社員画面と同じ設定順、未割当は最後です。',{size:8,color:gray});
  for(const group of customerReport.groups)paragraph(`${group.label}：お客様 ${group.customers.length}人 ／ ${group.orderCount}注文 ／ ${yen(group.totalYen)}円`,{size:10});
  if(!customerReport.customers.length)paragraph('対象の有効注文はありません。',{size:12});
  paragraph('お客様ごとに改ページ。濃紺の見出しから「ご注文はここまで」の合計欄までが、そのお客様の注文です。',{size:9,color:gray});
  let customerNumber=0;
  for(const group of customerReport.groups)for(const person of group.customers){
   const customer=person.customer,number=String(++customerNumber).padStart(2,'0');
   const endLines=wrap(`${customer.name||'お名前未登録'} 様のご注文はここまで`,WIDTH-28,11),endHeight=40+endLines.length*15;
   activeCustomer={person,group,number,endLines,endHeight,pages:[]};context='';newPage();
   person.rows.forEach((r,i)=>orderBody(r,{showCustomer:false,lastForCustomer:i===person.rows.length-1}));
   context='';ensure(endHeight+14);const top=y+7;
   page.drawRectangle({x:M,y:top-endHeight,width:WIDTH,height:endHeight,color:navy});
   let at=top-19;for(const label of endLines){draw(label,M+14,at,11,rgb(1,1,1));at-=15;}
   const amount=yen(person.totalYen)+'円';draw('お客様合計（税抜）',M+14,top-endHeight+13,10,rgb(1,1,1));draw(amount,W-M-14-textWidth(amount,13),top-endHeight+13,13,rgb(1,1,1));
   // 終了枠の下で縦線も止め、空白部分まで注文が続くように見せない。
   const below=top-endHeight-(BOTTOM+12);if(below>0)page.drawRectangle({x:M-12,y:BOTTOM+12,width:4,height:below,color:rgb(1,1,1)});
   y=top-endHeight-14;activeCustomer=null;
  }

 }else{
  const counts=new Map();for(const r of data.rows){const name=r.manufacturer||r.booth;const c=counts.get(name)||{n:0,total:0};c.n++;c.total+=r.total_yen;counts.set(name,c);}
  for(const [name,c] of counts)paragraph(`${name}：${c.n}件 ／ ${yen(c.total)}円`,{size:9});
  if(!data.rows.length)paragraph('対象の有効注文はありません。',{size:12});rule();for(const r of data.rows)orderBody(r);
 }
 }
 let summaryStartPage=null;
 if(view==='orders'||view==='delivery-products'){
 const segments=deliverySummary?.groups||[{key:'all',label:'全注文',data,summary}];
 for(const segment of segments){
 const summary=segment.summary,sectionData=segment.data,split=view==='delivery-products',later=split&&segment.key==='later',unknown=split&&segment.key==='unknown';
 // 注文単位の明細とは別のページ・色・見出しで、最後に商品合算の一覧を付ける。
 if(summaryStartPage===null)summaryStartPage=pages.length+1;
 const green=later?rgb(.12,.29,.54):unknown?rgb(.55,.32,.04):rgb(.10,.32,.25),paleGreen=later?rgb(.91,.95,.99):unknown?rgb(1,.96,.88):rgb(.91,.96,.92),summaryStripe=later?rgb(.96,.98,1):rgb(.96,.98,.96);
 const summaryTop=H-(split?188:173);
 const summaryCols=[58,213,28,42,48,42,WIDTH-431];
 function summaryPage(){
  page=doc.addPage([W,H]);pages.push(page);context='';
  if(split)pageDeliveries.set(page,segment.label);
  page.drawRectangle({x:M,y:H-82,width:WIDTH,height:46,color:paleGreen});
  page.drawRectangle({x:M,y:H-82,width:5,height:46,color:green});
  draw(split?`商品別集計 / ${segment.label}`:'商品別集計',M+14,H-54,18,green);draw(split?'この受渡区分の注文を、メーカー・商品ごとに合算した一覧です':'全注文を商品ごとに合算した一覧です',M+14,H-71,8,gray);
  draw(`${data.phase==='practice'?'【練習データ・売上入力しない】':'【本番データ】'} ${title}`,M,H-101,10,data.phase==='practice'?warn:green);
  draw(`取得 ${when(data.fetched_at)} ／ 取消済みを除く・金額は税抜`,M,H-116,8,gray);
  draw(`有効注文 ${sectionData.order_count}件 ／ 商品 ${summary.productCount}種類 ／ 税抜合計 ${yen(summary.totalYen)}円`,M,H-136,10,green);
  draw('購入とサービスを区別。数量は本・個・セットなど、登録された単位のままです。',M,H-152,8,gray);
  if(split)draw(`全メーカー・全受渡区分の合計（税抜） ${yen(data.total_yen)}円`,M,H-167,8,gray);y=summaryTop;
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
   if(reserve+wrap(maker.name+'（続き）',WIDTH,12).length*16+53>summaryTop-BOTTOM)throw Error('PDF_ROW_TOO_LONG');
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
 page.drawRectangle({x:M,y:y-26,width:WIDTH,height:33,color:green});draw(split?'受渡区分合計（税抜）':'商品別集計 合計（税抜）',M+10,y-14,11,rgb(1,1,1));
 const grandTotal=yen(summary.totalYen)+'円';draw(grandTotal,W-M-10-font.widthOfTextAtSize(grandTotal,12),y-14,12,rgb(1,1,1));
 }
 }
 pages.forEach((p,i)=>{
  p.drawText(`${i+1} / ${pages.length}`,{x:W-M-60,y:22,size:8,font,color:gray});
  const section=pageCustomers.get(p);let label;
  if(section){
   const suffix=` ｜ このお客様の ${section.pages.indexOf(p)+1} / ${section.pages.length} ページ`,prefix=`お客様 ${section.number} ｜ `;
   let name=String(section.person.customer.name||'お名前未登録');while(name&&textWidth(prefix+name+suffix,8)>WIDTH-82)name=[...name].slice(0,-1).join('');
   label=prefix+name+suffix;
  }else label=pageDeliveries.has(p)?`商品別集計 / ${pageDeliveries.get(p)} / BEAU Fes`:summaryStartPage&&i+1>=summaryStartPage?'商品別集計 / BEAU Fes 注文データ':view==='customers'?'お客様別 / 営業担当順 / BEAU Fes':data.phase==='practice'?'練習データ / BEAU Fes 注文データ':'BEAU Fes 注文データ / 出力時点の注文';
  // 絵文字を含むお名前も通常の本文と同じ画像併用描画でフッターへ表示する。
  const previous=page;page=p;draw(label,M,22,8,gray);page=previous;
 });

 doc.setTitle(view==='delivery-products'?'BEAU Fes 全メーカー商品別集計（受渡別）':view==='customers'?'BEAU Fes お客様別注文（営業担当順）':'BEAU Fes 注文データ');doc.setCreator('BEAU Fes 注文PDF出力');doc.setLanguage('ja-JP');
 return {bytes:await doc.save(),pages:pages.length,orderCount:data.order_count,totalYen:data.total_yen,summaryStartPage,summaryPageCount:summaryStartPage?pages.length-summaryStartPage+1:0,productCount:deliverySummary?.productCount??summary.productCount,customerCount:customerReport?.customerCount??data.customer_count,view,...(deliverySummary?{deliveryGroups:deliverySummary.groups.map(g=>({key:g.key,label:g.label,orderCount:g.data.order_count,totalYen:g.summary.totalYen,productCount:g.summary.productCount}))}:{})};
}
