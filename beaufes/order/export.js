// 例外時だけ使う受付情報。既存の備考・購入控え・出力経路で同じ内容を保持する。
const START='【受付情報】\n', END='\n【連絡事項】\n';
export const emptyHandling=()=>({billing:'salon',contact:'registered',paper:''});
export function packRemarks(remarks='',value={}){
 const h={...emptyHandling(),...value};
 const paper=String(h.paper||'').normalize('NFKC').trim().toUpperCase();
 if(!['salon','personal','review'].includes(h.billing)||!['registered','review'].includes(h.contact))throw Error('INVALID_HANDLING');
 if(paper&&!/^[A-Z0-9][A-Z0-9-]{0,31}$/.test(paper))throw Error('INVALID_PAPER_NUMBER');
 if(String(remarks).includes('【受付情報】')||String(remarks).includes('【連絡事項】'))throw Error('RESERVED_REMARKS');
 const encoded=h.billing==='salon'&&h.contact==='registered'&&!paper?String(remarks):START+
  `請求先：${{salon:'サロン',personal:'個人（事務確認）',review:'事務確認'}[h.billing]}\n配送先・連絡先：${h.contact==='registered'?'登録情報':'事務確認'}\n紙受付No：${paper||'なし'}`+END+String(remarks);
 if([...encoded].length>1000)throw Error('INVALID_REMARKS');
 return encoded;
}
export function unpackRemarks(text=''){
 const raw=String(text||'');const end=raw.indexOf(END);
 if(!raw.startsWith(START)||end<0)return {...emptyHandling(),remarks:raw,encoded:false};
 const header=raw.slice(START.length,end);
 const m=/^請求先：(サロン|個人（事務確認）|事務確認)\n配送先・連絡先：(登録情報|事務確認)\n紙受付No：(なし|[A-Z0-9][A-Z0-9-]{0,31})$/.exec(header);
 if(!m)return {...emptyHandling(),billing:'review',contact:'review',remarks:raw,invalid:true,encoded:true};
 return {billing:{'サロン':'salon','個人（事務確認）':'personal','事務確認':'review'}[m[1]],contact:m[2]==='登録情報'?'registered':'review',paper:m[3]==='なし'?'':m[3],remarks:raw.slice(end+END.length),encoded:true};
}
export function handlingWarnings(order){
 const h=unpackRemarks(order.remarks||'');const warnings=[];
 if(h.billing!=='salon')warnings.push('請求先・請求用得意先コードを事務確認');
 if(h.contact!=='registered')warnings.push('配送・連絡を事務確認（備考の先渡し数も確認）');
 if((order.lines||[]).some(l=>/^ACT26-/.test(l.product_code)))warnings.push('コフレ仮ID：正式商品コード照合まで販売管理入力保留');
 return warnings;
}
export const deliveryLabel=value=>value==='later'?'後日納品':value==='takeaway'?'当日お渡し済み':'未選択';

// 商品コードは表示時だけ0埋めを外す。得意先コードや注文の内部キーは変更しない。
export const displayProductCode=value=>String(value??'').replace(/^0+(?=\d+$)/,'');
export async function buildWorkbook(ExcelJS, data) {
 const workbook=new ExcelJS.Workbook();workbook.creator='BEAU Fes';
 const headings=['出力番号','注文番号','メーカー','得意先コード','サロン名','お客様名','商品コード','商品名','区分','数量','税抜単価','税抜金額','受渡方法','売上日','確認事項','備考','来場者番号','請求先','紙受付No'];
 const sheets=['入力候補','要確認'].map(name=>{const sheet=workbook.addWorksheet(name);sheet.addRow(headings);sheet.views=[{state:'frozen',ySplit:1}];sheet.columns.forEach((col,i)=>{col.width=[38,38,18,19,24,20,14,46,19,10,14,14,18,14,44,48,22,24,25][i];});sheet.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};sheet.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF214D42'}};sheet.getRow(1).height=28;sheet.autoFilter={from:'A1',to:'S1'};return sheet;});
 for(const order of data.rows){
  const handling=unpackRemarks(order.remarks||'');
  const notes=[...handlingWarnings(order),order.state==='cancelled'?'取消・入力対象外':'',!order.customer.customer_code?'得意先コード要確認':'',order.posting.state!=='ready'?`台帳:${({claimed:'入力中',entered:'入力済み',unknown:'入力結果が不明',correction_required:'訂正必要',resolved_cancelled:'取消・売上なし照合完了'}[order.posting.state]||order.posting.state)}・再入力禁止`:''].filter(Boolean).join('／');
  const sheet=sheets[notes?1:0];
  for(const line of order.lines||[]){const row=sheet.addRow([data.batch_id,order.request_id,order.booth,handling.billing==='salon'?String(order.customer.customer_code||''):'',order.customer.salon,order.customer.name,displayProductCode(line.product_code),line.product_name,line.kind==='gift'?'サービス（0円）':'購入',line.qty,line.unit_price_yen,line.qty*line.unit_price_yen,order.delivery==='later'?'後日納品':'当日お渡し済み',order.posting.sales_date||'',notes,order.remarks||'',order.customer.id||'',{salon:'サロン',personal:'個人（事務確認）',review:'事務確認'}[handling.billing],handling.paper]);
   row.alignment={vertical:'top',wrapText:true};row.height=42;for(const index of [4,7])row.getCell(index).numFmt='@';for(const index of [10,11,12])row.getCell(index).numFmt='#,##0';if(line.kind==='gift')row.getCell(9).font={color:{argb:'FF246753'},bold:true};
  }
 }
 return workbook;
}
