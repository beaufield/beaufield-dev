// 日本語コメント: 認証後に読んだ有効注文を明細・商品・お客様の3つの見方で端末内に集計する。
import {prepareReport} from './order-pdf.mjs?build=20261025-atfit';
import {unpackRemarks,billingLabel,deliveryLabel,handlingWarnings} from '../order/export.js?build=20261025-atfit';

export const detailHeaders=['メーカー','注文番号','注文日時（日本時間）','来場者番号','得意先コード','サロン名','お客様名','商品コード','商品名','単位','区分','数量','税抜単価（円）','税抜金額（円）','受渡方法','請求先','紙受付No','備考','確認事項'];
export const summaryHeaders=['メーカー','商品コード','商品名','単位','購入数量','サービス数量','合計数量','税抜金額（円）','注文件数','購入最小単価（円）','購入最大単価（円）','確認事項'];
export const detailWidths=[24,40,24,19,17,28,22,19,58,12,14,12,19,19,20,23,18,52,60];
export const summaryWidths=[24,19,58,12,14,16,14,20,14,22,22,60];
export const HEADER_ROW=6,DATA_ROW=7;
const literal=s=>String(s??'').startsWith('=')?"'"+s:s;
const excelDate=value=>{const n=Date.parse(value);if(!Number.isFinite(n))throw Error('REPORT_DATA_MISMATCH');return n/86400000+25569+9/24;};

export function prepareExcelData(input,makerId=''){
 const data=prepareReport(input,makerId),details=[],groups=new Map();
 for(const order of data.rows){
  const h=unpackRemarks(order.remarks||''),maker=order.manufacturer||order.booth;
  const notes=[...handlingWarnings(order),!order.customer.customer_code?'得意先コードを確認':'',h.invalid?'受付情報の形式を確認':''].filter(Boolean).join('／');
  for(const line of order.lines){
   // 商品・得意先コードは識別子として元の文字列と先頭ゼロを保持する。
   const code=String(line.product_code??''),name=String(line.product_name??''),unit=String(line.unit??'');
   if(!code||!name||!unit)throw Error('REPORT_DATA_MISMATCH');
   const amount=line.qty*line.unit_price_yen;
   details.push([maker,order.request_id,excelDate(order.received_at),order.customer.id||'',order.customer.customer_code||'',order.customer.salon||'',order.customer.name||'',code,name,unit,line.kind==='gift'?'サービス':'購入',line.qty,line.unit_price_yen,amount,deliveryLabel(order.delivery),billingLabel(h.billing),h.paper,h.remarks,notes].map(literal));
   const key=JSON.stringify([order.booth,code,name,unit]);
   if(!groups.has(key))groups.set(key,{maker,code,name,unit,paid:0,gift:0,amount:0,orders:new Set(),prices:[],notes:new Set()});
   const group=groups.get(key);group[line.kind==='gift'?'gift':'paid']+=line.qty;group.amount+=amount;group.orders.add(order.request_id);
   if(line.kind==='paid')group.prices.push(line.unit_price_yen);
   for(const warning of handlingWarnings({lines:[line]}))group.notes.add(warning);
  }
 }
 const summary=[...groups.values()].sort((a,b)=>a.maker.localeCompare(b.maker,'ja')||a.code.localeCompare(b.code)||a.name.localeCompare(b.name,'ja')||a.unit.localeCompare(b.unit,'ja')).map(g=>[g.maker,g.code,g.name,g.unit,g.paid,g.gift,g.paid+g.gift,g.amount,g.orders.size,g.prices.length?Math.min(...g.prices):null,g.prices.length?Math.max(...g.prices):null,[...g.notes].join('／')].map(literal));
 if(details.reduce((n,r)=>n+r[13],0)!==data.total_yen||summary.reduce((n,r)=>n+r[7],0)!==data.total_yen)throw Error('REPORT_DATA_MISMATCH');
 return {data,details,summary,customerMatrix:prepareCustomerMatrix(data)};
}

export const excelColumn=number=>{let s='';for(let n=number;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const textUnits=value=>[...String(value)].reduce((n,c)=>n+(c.charCodeAt(0)>255?2:1),0);
export function prepareCustomerMatrix(data){
 const manufacturers=[...new Map(data.rows.map(r=>[r.booth,{id:r.booth,name:r.manufacturer||r.booth}])).values()].sort((a,b)=>a.name.localeCompare(b.name,'ja')||a.id.localeCompare(b.id));
 // 明細の表示名で計算を照合するため、異なるメーカーの同名を黙って合算しない。
 if(new Set(manufacturers.map(m=>m.name)).size!==manufacturers.length)throw Error('REPORT_DATA_MISMATCH');
 const people=new Map();
 for(const order of data.rows){
  const id=order.customer.id||'',key=id?'id:'+id:'order:'+order.booth+':'+order.request_id;
  if(!people.has(key))people.set(key,{id,requestId:order.request_id,key,customer:order.customer,latest:'',orders:new Set(),makers:new Map(),total:0});
  const person=people.get(key);person.orders.add(order.booth+':'+order.request_id);person.total+=order.total_yen;
  if(String(order.received_at)>=person.latest){person.customer=order.customer;person.latest=String(order.received_at);}
  if(!person.makers.has(order.booth))person.makers.set(order.booth,{amount:0,items:new Map()});
  const cell=person.makers.get(order.booth);cell.amount+=order.total_yen;
  for(const line of order.lines){
   const itemKey=JSON.stringify([line.product_code,line.product_name,line.unit,line.kind]);
   if(!cell.items.has(itemKey))cell.items.set(itemKey,{...line,qty:0});cell.items.get(itemKey).qty+=line.qty;
  }
 }
 const customers=[...people.values()].sort((a,b)=>Number(!a.customer.customer_code)-Number(!b.customer.customer_code)||String(a.customer.customer_code||'').localeCompare(String(b.customer.customer_code||''))||String(a.customer.salon||'').localeCompare(String(b.customer.salon||''),'ja')||String(a.customer.name||'').localeCompare(String(b.customer.name||''),'ja')||a.key.localeCompare(b.key));
 const headers=['来場者番号','得意先コード','サロン名','お客様名','注文件数','購入メーカー数','税抜合計（円）',...manufacturers.flatMap(m=>[m.name+'｜購入内容',m.name+'｜税抜金額（円）'])];
 const widths=[21,18,28,22,14,18,20,...manufacturers.flatMap(()=>[58,22])];
 const rows=[],rowHeights=[];
 for(const person of customers){
  const row=[person.id||'番号未登録（注文ごと）',person.customer.customer_code||'',person.customer.salon||'',person.customer.name||'',person.orders.size,person.makers.size,person.total];let wrappedLines=1;
  for(const maker of manufacturers){
   const cell=person.makers.get(maker.id);
   const contents=cell?[...cell.items.values()].sort((a,b)=>(a.kind==='gift')-(b.kind==='gift')||String(a.product_code).localeCompare(String(b.product_code))||String(a.product_name).localeCompare(String(b.product_name),'ja')).map(item=>`${item.product_name} [${item.product_code}] ×${item.qty}${item.unit}${item.kind==='gift'?'（サービス）':''}`).join('\n'):'';
   if(contents.length>32767)throw Error('EXPORT_TOO_LARGE');
   wrappedLines=Math.max(wrappedLines,contents.split('\n').reduce((n,line)=>n+Math.max(1,Math.ceil(textUnits(line)/55)),0));
   row.push(contents,cell?.amount||0);
  }
  // Excelの行高上限を超える内容を、切れたまま保存しない。
  const height=Math.max(36,wrappedLines*14+8);if(height>409)throw Error('EXPORT_TOO_LARGE');
  rows.push(row.map(literal));rowHeights.push(height);
 }
 if(rows.reduce((n,r)=>n+r[6],0)!==data.total_yen)throw Error('REPORT_DATA_MISMATCH');
 return {manufacturers,customers,headers,widths,rows,rowHeights};
}

export function customerAmountFormula(row,last,maker,customer){
 const source="'注文明細'!",matchColumn=customer.id?'D':'B',matchValue=customer.id?`$A${row}`:`"${customer.requestId.replace(/"/g,'""')}"`;
 return `=SUMPRODUCT((${source}$${matchColumn}$${DATA_ROW}:$${matchColumn}$${last}=${matchValue})*(${source}$A$${DATA_ROW}:$A$${last}="${maker.name.replace(/"/g,'""')}")*${source}$N$${DATA_ROW}:$N$${last})`;
}

// 数量と金額は明細を参照する式。購入数とサービス数を分け、同じメーカー・コード・商品名・単位で集計する。
export function summaryFormulas(row,last){
 const prefix="'注文明細'!",range=column=>`${prefix}$${column}$${DATA_ROW}:$${column}$${last}`;
 // SUMPRODUCTの文字列比較で、数字だけの商品コードも先頭ゼロを含めて照合する。
 const criteria=[['A','A'],['H','B'],['I','C'],['J','D']].map(([source,target])=>`(${range(source)}=${target}${row})`).join('*');
 return [`=SUMPRODUCT(${criteria}*(${range('K')}="購入")*${range('L')})`,`=SUMPRODUCT(${criteria}*(${range('K')}="サービス")*${range('L')})`,`=E${row}+F${row}`,`=SUMPRODUCT(${criteria}*${range('N')})`];
}

export async function buildOrderExcel(ExcelJS,input,{makerId='',title='全メーカー'}={}){
 const model=prepareExcelData(input,makerId),workbook=new ExcelJS.Workbook();workbook.creator='BEAU Fes';workbook.calcProperties.fullCalcOnLoad=true;
 const last=Math.max(DATA_ROW,DATA_ROW+model.details.length-1);
 for(const [index,name,headers,widths,rows] of [[0,'注文明細',detailHeaders,detailWidths,model.details],[1,'商品別集計',summaryHeaders,summaryWidths,model.summary]]){
  const sheet=workbook.addWorksheet(name);sheet.views=[{state:'frozen',ySplit:HEADER_ROW,showGridLines:false}];
  sheet.columns=widths.map(width=>({width}));sheet.getCell('A2').value=`BEAU Fes 2026 ${title} ${name}`;sheet.getCell('A2').font={name:'Arial',size:14,bold:true};
  sheet.getCell('A3').value=`${input.phase==='practice'?'練習注文':'本番・有効注文のみ'}／取得 ${new Date(input.fetched_at).toLocaleString('ja-JP',{timeZone:'Asia/Tokyo',hour12:false})}／金額は税抜`;
  sheet.getCell('A4').value='税抜合計（円）';sheet.getCell('B4').value={formula:`SUM(${index?'H':'N'}${DATA_ROW}:${index?'H':'N'}${index?Math.max(DATA_ROW,DATA_ROW+rows.length-1):last})`,result:input.total_yen};sheet.getCell('B4').numFmt='#,##0';
  sheet.getCell('C4').value=`注文 ${input.order_count}件／お客様 ${model.data.customer_count}人／${index?rows.length+'商品':'明細 '+rows.length+'行'}`;
  sheet.getRow(HEADER_ROW).values=headers;
  for(let i=0;i<rows.length;i++){
   const row=sheet.getRow(DATA_ROW+i);row.values=rows[i];row.font={name:'Arial',size:10};row.alignment={vertical:'middle'};row.height=30;
   if(index){const formulas=summaryFormulas(DATA_ROW+i,last);for(let j=0;j<4;j++)row.getCell(5+j).value={formula:formulas[j].slice(1),result:rows[i][4+j]};}
   else row.getCell(14).value={formula:`L${DATA_ROW+i}*M${DATA_ROW+i}`,result:rows[i][13]};
  }
  const end=Math.max(HEADER_ROW,DATA_ROW+rows.length-1);sheet.autoFilter={from:{row:HEADER_ROW,column:1},to:{row:end,column:headers.length}};
  const header=sheet.getRow(HEADER_ROW);header.font={name:'Arial',size:10,bold:true,color:{argb:'FFFFFFFF'}};header.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF244860'}};header.alignment={horizontal:'center',vertical:'middle'};header.height=30;
  const numeric=index?[5,6,7,8,9,10,11]:[12,13,14];for(const col of numeric)sheet.getColumn(col).numFmt='#,##0';
  for(const col of index?[2]:[2,4,5,8])sheet.getColumn(col).numFmt='@';
  if(!index)sheet.getColumn(3).numFmt='yyyy-mm-dd hh:mm:ss';
  for(const col of index?[3,12]:[9,18,19])sheet.getColumn(col).alignment={vertical:'middle',wrapText:true};
  if(!rows.length)sheet.getCell(`A${DATA_ROW}`).value='対象の有効注文はありません。';
 }
 const matrix=model.customerMatrix,sheet=workbook.addWorksheet('お客様×メーカー');
 sheet.views=[{state:'frozen',xSplit:4,ySplit:HEADER_ROW,showGridLines:false}];sheet.columns=matrix.widths.map(width=>({width}));
 sheet.getCell('A2').value=`BEAU Fes 2026 ${title} お客様別の購入内容`;sheet.getCell('A2').font={name:'Arial',size:14,bold:true};
 sheet.getCell('A3').value='来場者番号ごとに追加注文を合算。メーカーごとに購入内容と税抜金額を表示。サービス品は別表示。氏名・コードは最新注文時点。';
 sheet.getCell('A4').value='税抜合計（円）';sheet.getCell('B4').value={formula:`SUM(G${DATA_ROW}:G${Math.max(DATA_ROW,DATA_ROW+matrix.rows.length-1)})`,result:input.total_yen};sheet.getCell('B4').numFmt='#,##0';
 sheet.getCell('C4').value=`お客様 ${matrix.rows.length}人／メーカー ${matrix.manufacturers.length}社／注文 ${input.order_count}件`;
 sheet.getRow(HEADER_ROW).values=matrix.headers;
 const header=sheet.getRow(HEADER_ROW);header.height=44;header.font={name:'Arial',size:10,bold:true,color:{argb:'FFFFFFFF'}};header.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF244860'}};header.alignment={horizontal:'center',vertical:'middle',wrapText:true};
 for(let i=0;i<matrix.rows.length;i++){
  const number=DATA_ROW+i,row=sheet.getRow(number);row.values=matrix.rows[i];row.height=matrix.rowHeights[i];row.font={name:'Arial',size:10};row.alignment={vertical:'top',wrapText:true};
  const amounts=[];
  for(let m=0;m<matrix.manufacturers.length;m++){const column=9+m*2;amounts.push(excelColumn(column)+number);row.getCell(column).value={formula:customerAmountFormula(number,last,matrix.manufacturers[m],matrix.customers[i]).slice(1),result:matrix.rows[i][column-1]};row.getCell(column).numFmt='#,##0';}
  row.getCell(7).value={formula:amounts.length?`SUM(${amounts.join(',')})`:'0',result:matrix.rows[i][6]};row.getCell(7).numFmt='#,##0';
 }
 for(const column of [1,2])sheet.getColumn(column).numFmt='@';for(const column of [5,6,7])sheet.getColumn(column).numFmt='#,##0';
 sheet.autoFilter={from:{row:HEADER_ROW,column:1},to:{row:Math.max(HEADER_ROW,DATA_ROW+matrix.rows.length-1),column:matrix.headers.length}};
 if(!matrix.rows.length)sheet.getCell(`A${DATA_ROW}`).value='対象の有効注文はありません。';
 return {bytes:await workbook.xlsx.writeBuffer(),orderCount:input.order_count,totalYen:input.total_yen,detailCount:model.details.length,productCount:model.summary.length,customerCount:matrix.rows.length};
}
