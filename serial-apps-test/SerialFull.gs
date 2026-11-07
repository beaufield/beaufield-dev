// 検証環境専用。原票は読み取りだけ。取込先の固定照合後に処理する。
var SX_CACHE_=null;
function sxStable_(v){if(Array.isArray(v))return v.map(sxStable_);if(v&&typeof v==='object'){const o={};Object.keys(v).sort().forEach(k=>o[k]=sxStable_(v[k]));return o;}return v;}
function sxHash_(v){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,typeof v==='string'?v:JSON.stringify(sxStable_(v)),Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}
function sxRpc_(name,args,jwt){
 if(_PROPS.getProperty('ENV')!=='test'||DB_URL!=='https://qxzdypqbqduwuwmwnbgk.supabase.co')throw Error('TEST_ENV_REQUIRED');
 const headers={apikey:DB_SECRET_KEY,'Content-Profile':'public','Accept-Profile':'public'};if(jwt)headers.Authorization='Bearer '+jwt;
 const r=UrlFetchApp.fetch(DB_URL+'/rest/v1/rpc/'+name,{method:'post',contentType:'application/json',headers:headers,payload:JSON.stringify(args),muteHttpExceptions:true});
 if(r.getResponseCode()!==200){let code='UNKNOWN';try{const e=JSON.parse(r.getContentText());if(/^[A-Z0-9]+$/.test(e.code||''))code=e.code;}catch(_){}throw Error('RPC_'+name+'_'+r.getResponseCode()+'_'+code);}
 return JSON.parse(r.getContentText());
}
function sxAdmin_(p){
 const r=JSON.parse(srIssueDbToken_(Object.assign({},p,{app:'serial-apps'})).getContent());
 if(!r.success)throw Error(r.error);
 const claims=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(r.token.split('.')[1])).getDataAsString());
 const users=SpreadsheetApp.openById(SX_AUTH_ID_).getSheetByName('users').getDataRange().getValues().slice(1);
 const u=users.find(x=>String(x[0])===claims.user_id);
 if(claims.app_role!=='admin'&&!(u&&(u[5]===true||u[5]==='TRUE')))throw Error('ADMIN_REQUIRED');
 return r.token;
}
function sxSnapshot_(){
 const ss=SpreadsheetApp.openById(SX_SOURCE_ID_);
 const products=ss.getSheetByName('ProductMaster').getDataRange().getValues(),shipping=ss.getSheetByName('SerialShipping').getDataRange().getValues();
 if(products[0][0]!=='商品コード'||shipping[0][0]!=='ID'||shipping[0].length!==17)throw Error('SOURCE_SCHEMA_CHANGED');
 return {products:products,shipping:shipping};
}
function sxRaw_(row){return row.map(v=>v instanceof Date?{type:'date',iso:v.toISOString()}:v);}
function sxSourceAdapter_(){return {getSheetByName:name=>({getDataRange:()=>({getValues:()=>name==='ProductMaster'?SX_CACHE_.products:SX_CACHE_.shipping})})};}
function sxData_(row,kind){
 if(kind==='products'){const keys=['code','name','jan','maker','series','size','serialType'];return Object.fromEntries(keys.map((k,i)=>[k,String(row[i]==null?'':row[i])]));}
 const keys=['id','shipDate','productCode','productName','jan','serial','status','returnDate','cancelDate','reason','method','customer','createdAt','registeredBy','maker','series','size'];
 return Object.fromEntries(keys.map((k,i)=>[k,[1,7,8].includes(i)?sx_formatDate(row[i]):String(row[i]==null?'':row[i])]));
}
function sxImport_(snapshot){
 const raw={products:snapshot.products.map(sxRaw_),shipping:snapshot.shipping.map(sxRaw_)};
 const sourceHash=sxHash_(raw),info=sxRpc_('sr_import',{p:{phase:'info'}});
 if(info.phase!=='idle'&&info.source_hash!==sourceHash)throw Error('EXISTING_IMPORT_SOURCE_MISMATCH');
 if(info.phase==='staging')throw Error('IMPORT_IN_PROGRESS_OR_NEEDS_RECOVERY');
 if(info.phase==='verified')return {success:true,already_imported:true,products:raw.products.length-1,shipping:raw.shipping.length-1,source_hash:sourceHash};
 const batch=info.phase==='published'?info.batch:Utilities.getUuid();
 if(info.phase==='idle'){
 sxRpc_('sr_import',{p:{phase:'begin',batch:batch,source_hash:sourceHash}});
 for(const kind of ['products','shipping']){
  const list=snapshot[kind].slice(1).map((row,i)=>({n:i+1,data:sxData_(row,kind),raw:sxRaw_(row)}));
  for(let i=0;i<list.length;i+=500)sxRpc_('sr_import',{p:{phase:'stage',batch:batch,kind:kind,rows:list.slice(i,i+500)}});
 }
 sxRpc_('sr_import',{p:{phase:'publish',batch:batch,products:raw.products.length-1,shipping:raw.shipping.length-1}});
 }
 for(const kind of ['products','shipping']){
  let after=0;const expected=raw[kind].slice(1);
  while(after<expected.length){const check=sxRpc_('sr_import',{p:{phase:'verify',batch:batch,kind:kind,after:after}});const chunk=expected.slice(after,after+500);if(sxHash_(chunk)!==sxHash_(check.rows))throw Error('RAW_READBACK_MISMATCH_'+kind);after+=chunk.length;}
 }
 if(sxHash_({products:sxSnapshot_().products.map(sxRaw_),shipping:sxSnapshot_().shipping.map(sxRaw_)})!==sourceHash)throw Error('SOURCE_CHANGED_DURING_IMPORT');
 sxRpc_('sr_import',{p:{phase:'finish',batch:batch}});
 return {success:true,products:raw.products.length-1,shipping:raw.shipping.length-1,source_hash:sourceHash,all_raw_rows_verified:true};
}
function sxCsvText_(rows){const fields=['shipDate','productCode','productName','jan','serial','status','returnDate','cancelDate','reason','customer'];return '\ufeff'+[['出荷日','商品コード','商品名','JANコード','シリアルNo','状態','返品日','取消日','取消理由','得意先'],...rows.map(r=>fields.map(k=>r[k]||''))].map(row=>row.map(sx_csvEscape).join(',')).join('\r\n');}
function sxCompare_(snapshot,jwt,p){
 const info=sxRpc_('sr_import',{p:{phase:'info'}});
 const hash=sxHash_({products:snapshot.products.map(sxRaw_),shipping:snapshot.shipping.map(sxRaw_)});
 if(info.source_hash!==hash||info.phase!=='verified')throw Error('SOURCE_NOT_SAME_AS_IMPORT');
 SX_CACHE_=snapshot;
 const makers=[...new Set(snapshot.products.slice(1).map(r=>String(r[3])).filter(Boolean))].sort();
 const months=[...new Set(snapshot.shipping.slice(1).map(r=>sx_shipDateYmd(r[1]).slice(0,7)).filter(Boolean))].sort();
 const cases=[];
 for(const maker of makers)for(const month of months){
  const exp=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(jwt.split('.')[1])).getDataAsString()).exp;
  if(exp*1000<Date.now()+15000)jwt=sxAdmin_(p);
  const expected=sxLegacyExport_({maker:maker,yearMonth:month}).getContent();
  let cursor={after:0},data=[],sizes=[];
  for(;;){const page=sxRpc_('sr_page',{p_filter:{maker:maker,yearMonth:month},p_after:cursor.after,p_highwater:cursor.highwater||null,p_revision:cursor.revision==null?null:cursor.revision,p_csv:true},jwt);if(page.rows.length>500)throw Error('CSV_PAGE_LIMIT');sizes.push(page.rows.length);data.push(...page.rows);if(page.done)break;cursor=page;}
  const actual=sxCsvText_(data),equal=actual===expected;
  cases.push({maker:maker,month:month,rows:data.length,pages:sizes,equal:equal,legacy_sha256:sxHash_(expected),db_sha256:sxHash_(actual)});
 }
 return {success:cases.every(x=>x.equal),cases:cases,source_hash:hash};
}
function sxHandle_(p){const lock=LockService.getScriptLock();let locked=false;try{const jwt=sxAdmin_(p);if(p.mode==='import'){locked=lock.tryLock(1000);if(!locked)throw Error('IMPORT_ALREADY_RUNNING');}const snapshot=sxSnapshot_();return jsonResponse(p.mode==='import'?sxImport_(snapshot):sxCompare_(snapshot,jwt,p));}catch(e){return jsonResponse({success:false,error:String(e.message)});}finally{if(locked)lock.releaseLock();}}
