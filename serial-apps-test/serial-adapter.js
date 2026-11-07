/* 現行画面から検証DBへの接続。カメラ部分・読取エンジンは現行のまま。 */
let serialDb=null,serialCursor=null,serialFilter=null;
function serialSession(){try{return JSON.parse(localStorage.getItem('bf_session'));}catch{return null;}}
function serialDatabase(){if(!serialDb){const c=window.SERIAL_TEST_CONFIG;serialDb=new SerialClient({...c,session:serialSession,issueToken:async token=>{const r=await fetch(c.issuer,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'serialIssueDbToken',session_token:token})});return r.json();}});}return serialDb;}
function serialPendingKey(){return 'sa_test_v31_pending:'+serialSession()?.user_id;}
async function serialAdapterCall(p,meta){
 if(window.serialRecoveryUseLegacy&&window.serialLegacyAdapterCall)return window.serialLegacyAdapterCall(p);
 const db=serialDatabase();
 if(p.action==='getProductMaster')return {success:true,data:await db.rpc('products',{})};
 if(p.action==='search'){
  serialFilter={};for(const k of ['dateFrom','dateTo','statuses','maker','series','size','productName','productCode'])if(p[k])serialFilter[k]=p[k];if(p.serialNo)serialFilter.serial=p.serialNo;
  serialCursor=await db.page(serialFilter);return {success:true,data:serialCursor.rows,total:serialCursor.rows.length};
 }
 if(!['registerShipping','registerReturn'].includes(p.action))throw Error('TEST_ACTION_NOT_ALLOWED');
 const kind=p.action==='registerShipping'?'ship':p.kind==='取消'?'cancel':'return';
 const control=await db.rpc('control',{});if(control.writer!=='db')throw Error('切替中のため保存できません');
 const payload={writer_epoch:control.epoch,productCode:p.productCode,serials:p.serials,date:p.shipDate||p.returnDate,method:p.method,customer:p.customer};
 const clean=JSON.parse(JSON.stringify(payload)),body=JSON.stringify({kind,payload:clean});
 let pending=JSON.parse(sessionStorage.getItem(serialPendingKey())||'null');
 if(pending&&pending.body!==body)throw Error('未確認の操作があります。「未確認操作を確認」を押してください');
 if(pending&&meta)meta.retried=true;
 if(!pending){pending={id:crypto.randomUUID(),body};sessionStorage.setItem(serialPendingKey(),JSON.stringify(pending));}
 const r=await db.mutate(kind,clean,pending.id);sessionStorage.removeItem(serialPendingKey());
 return {success:true,registered:r.changed,updated:r.changed,skipped:r.skipped,skippedSerials:r.skippedSerials||[]};
}
async function serialResolvePending(){try{const pending=JSON.parse(sessionStorage.getItem(serialPendingKey())||'null');if(!pending){showToast('未確認操作はありません','success');return;}const r=await serialDatabase().rpc('operation',{p_id:pending.id});sessionStorage.removeItem(serialPendingKey());showToast(r?'前の操作は登録済みです':'前の操作は未登録です。入力を確認して再登録できます','success');}catch(e){showToast(e.message,'error');}}
async function serialNextPage(){if(!serialCursor||serialCursor.done){showToast('次のページはありません','warn');return;}showLoading(true);try{serialCursor=await serialDatabase().page(serialFilter,serialCursor);renderSearchResults(serialCursor.rows,serialCursor.rows.length);}catch(e){showToast(e.message,'error');}finally{showLoading(false);}}
function serialCsvText(rows){const fields=['shipDate','productCode','productName','jan','serial','status','returnDate','cancelDate','reason','customer'],escape=v=>{const s=String(v??'');return /[",\r\n]/.test(s)?'"'+s.replaceAll('"','""')+'"':s;};return '\ufeff'+[['出荷日','商品コード','商品名','JANコード','シリアルNo','状態','返品日','取消日','取消理由','得意先'],...rows.map(r=>fields.map(k=>r[k]||''))].map(row=>row.map(escape).join(',')).join('\r\n');}
async function serialCsvDownload(){const maker=prompt('メーカー名（商品マスターと同じ表記）');if(!maker)return;const month=prompt('対象月（YYYY-MM）');if(!/^\d{4}-\d{2}$/.test(month||''))return;showLoading(true);const rows=[];try{await serialDatabase().csv({maker,yearMonth:month},part=>rows.push(...part));const url=URL.createObjectURL(new Blob([serialCsvText(rows)],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download='出荷シリアルNo管理'+month.slice(0,4)+'年'+Number(month.slice(5))+'月.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);showToast(rows.length+'件のCSVを作成しました','success');}catch(e){showToast('CSVを破棄しました: '+e.message,'error');}finally{rows.length=0;showLoading(false);}}
async function serialFullRun(mode){
 // 管理者・検証環境・取込状態の照合をサーバー側で行い、完了済みは再取込しない。
 const proof=document.getElementById('serial-proof');proof.textContent='処理中です。画面を閉じないでください。';showLoading(true);
 try{const c=window.SERIAL_TEST_CONFIG;if(c.url!=='https://qxzdypqbqduwuwmwnbgk.supabase.co')throw Error('TEST_CONFIG_REQUIRED');const r=await fetch(c.issuer,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'serialFullValidation',mode,session_token:serialSession()?.token})});const result=await r.json();proof.textContent=JSON.stringify(result,null,2);if(result.success&&mode==='import'){localStorage.removeItem('sa_test_v31_product_master_cache');serialDb=null;await loadProductMaster();}}catch(e){proof.textContent=e.message;}finally{showLoading(false);}
}
