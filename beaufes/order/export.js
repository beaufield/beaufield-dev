// アプリ組込みのExcel出力。識別コードは文字列セルにし、先頭0を維持する。
export async function buildWorkbook(ExcelJS, data) {
 const workbook=new ExcelJS.Workbook();workbook.creator='BEAU Fes';
 const headings=['出力番号','注文番号','メーカー','得意先コード','サロン名','お客様名','商品コード','商品名','区分','数量','税抜単価','税抜金額','受渡方法','売上日','確認事項'];
 const sheets=['入力候補','要確認'].map(name=>{const sheet=workbook.addWorksheet(name);sheet.addRow(headings);sheet.views=[{state:'frozen',ySplit:1}];sheet.columns.forEach((col,i)=>{col.width=[38,38,18,19,24,20,14,46,19,10,14,14,18,14,44][i];});sheet.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};sheet.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF214D42'}};sheet.getRow(1).height=28;sheet.autoFilter={from:'A1',to:'O1'};return sheet;});
 for(const order of data.rows){
  const notes=[order.state==='cancelled'?'取消・入力対象外':'',!order.customer.customer_code?'得意先コード要確認':'',order.posting.state!=='ready'?`台帳:${order.posting.state}・再入力禁止`:''].filter(Boolean).join('／');
  const sheet=sheets[notes?1:0];
  for(const line of order.lines||[]){const row=sheet.addRow([data.batch_id,order.request_id,order.booth,String(order.customer.customer_code||''),order.customer.salon,order.customer.name,String(line.product_code),line.product_name,line.kind==='gift'?'サービス（0円）':'購入',line.qty,line.unit_price_yen,line.qty*line.unit_price_yen,order.delivery==='later'?'後日納品':'当日持ち帰り',order.posting.sales_date||'',notes]);
   row.alignment={vertical:'top',wrapText:true};row.height=42;for(const index of [4,7])row.getCell(index).numFmt='@';for(const index of [10,11,12])row.getCell(index).numFmt='#,##0';if(line.kind==='gift')row.getCell(9).font={color:{argb:'FF246753'},bold:true};
  }
 }
 return workbook;
}
