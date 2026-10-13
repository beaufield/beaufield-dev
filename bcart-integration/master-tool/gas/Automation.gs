// Bモール新商品検知専用。旧12680経路・価格更新・登録・公開・通知には触れない。
const BMALL_AUTO_ACTIONS_ = Object.freeze(['bmallAutoStatus','bmallAutoSnapshot','bmallAutoReferences','bmallAutoPrepare','bmallAutoProposal','bmallAutoApprove','bmallAutoSave','bmallAutoResult']);
const BMALL_AUTO_JOURNAL_ = 'Bmall_新商品確認記録';
function autoFail_(s) { throw new Error('AUTO_'+s); }
function autoExact_(v, keys) { if(!v||Array.isArray(v)||typeof v!=='object'||Object.keys(v).sort().join('|')!==keys.slice().sort().join('|')) autoFail_('INVALID_FIELDS'); }
function autoId_(s) { if(typeof s!=='string'||!/^P-[0-9a-f-]{36}$/.test(s)) autoFail_('INVALID_PROPOSAL'); return s; }
function autoOwner_() { if(typeof BMALL_AUTO_OWNER_==='undefined'||typeof BMALL_AUTO_OWNER_!=='string') autoFail_('OWNER_NOT_CONFIGURED'); return BMALL_AUTO_OWNER_; }
function autoMac_(value) {
  const key=PropertiesService.getScriptProperties().getProperty('BMALL_BRIDGE_API_KEY');if(typeof key!=='string'||key.length<32)autoFail_('KEY_NOT_CONFIGURED');
  return Utilities.computeHmacSha256Signature(bridgeStable_(value),key,Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}
function autoJournal_(create) {
  const ss=getDataSpreadsheet_(); let sh=ss.getSheetByName(BMALL_AUTO_JOURNAL_);
  if(!sh&&create){ sh=ss.insertSheet(BMALL_AUTO_JOURNAL_);sh.getRange(1,1,1,3).setValues([['proposal_id','record_json','updated_at']]);SpreadsheetApp.flush(); }
  if(sh){const h=sh.getRange(1,1,1,3).getValues()[0];if(bridgeStable_(h)!==bridgeStable_(['proposal_id','record_json','updated_at']))autoFail_('JOURNAL_SCHEMA');}
  return sh;
}
function autoGet_(id) {
  autoId_(id);const sh=autoJournal_(false);if(!sh)return null;
  const a=sh.getDataRange().getValues().slice(1).map((v,i)=>({v:v,row:i+2})).filter(x=>String(x.v[0])===id);
  if(a.length>1)autoFail_('JOURNAL_DUPLICATE');if(!a.length)return null;
  if(sh.getRange(a[0].row,1,1,3).getFormulas()[0].some(Boolean))autoFail_('JOURNAL_FORMULA');
  let p;try{p=JSON.parse(String(a[0].v[1]));}catch(e){autoFail_('JOURNAL_CORRUPT');}
  const unsigned=Object.assign({},p);delete unsigned.recordMac;
  if(p.id!==id||p.version!==1||bridgeHash_(p.payload)!==p.payloadHash||!bridgeKeyMatches_(p.recordMac,autoMac_(unsigned)))autoFail_('JOURNAL_CORRUPT');return {p:p,row:a[0].row};
}
function autoPut_(p) {
  autoId_(p.id);if(bridgeHash_(p.payload)!==p.payloadHash)autoFail_('JOURNAL_CORRUPT');
  const unsigned=Object.assign({},p);delete unsigned.recordMac;p.recordMac=autoMac_(unsigned);
  const raw=JSON.stringify(p);if(raw.length>30000)autoFail_('PROPOSAL_TOO_LARGE');
  const sh=autoJournal_(true),found=autoGet_(p.id),row=found?found.row:sh.getLastRow()+1;
  sh.getRange(row,1,1,3).setNumberFormat('@');sh.getRange(row,1,1,3).setValues([[p.id,raw,new Date().toISOString()]]);SpreadsheetApp.flush();
  if(bridgeStable_(autoGet_(p.id).p)!==bridgeStable_(p))autoFail_('JOURNAL_DURABILITY_UNKNOWN');
}
function autoPublic_(p) {
  return {ok:true,proposalId:p.id,state:p.state,payload:p.payload,codes:p.codes,payloadHash:p.payloadHash,masterHash:p.masterHash,
    createdAt:p.createdAt,expiresAt:p.expiresAt,approvedAt:p.approvedAt||null,approvalExpiresAt:p.approvalExpiresAt||null,
    reviewLabels:p.reviewLabels||null,draftId:p.draftId||null,readbackVerified:p.state==='VERIFIED',notificationEnabled:false};
}
function autoMaster_() {
  if(!CSV_FOLDER_ID)autoFail_('MASTER_NOT_CONFIGURED');const it=DriveApp.getFolderById(CSV_FOLDER_ID).getFilesByName(CSV_FILENAME);
  if(!it.hasNext())autoFail_('MASTER_MISSING');const f=it.next();if(it.hasNext())autoFail_('MASTER_AMBIGUOUS');
  if(f.getMimeType()==='application/vnd.google-apps.spreadsheet')autoFail_('MASTER_WRONG_MIME');
  const updatedAt=f.getLastUpdated().getTime(),now=Date.now();if(updatedAt>now||now-updatedAt>86400000)autoFail_('MASTER_STALE');
  const blob=f.getBlob(),bytes=blob.getBytes(),size=f.getSize();if(bytes.length!==size||size<100)autoFail_('MASTER_INCOMPLETE');
  const hash=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,bytes).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
  const csv=Utilities.parseCsv(blob.getDataAsString('Shift_JIS'));if(csv.length<2)autoFail_('MASTER_INCOMPLETE');
  const h=csv[0].map(x=>String(x).trim().replace(/^\uFEFF/,''));if(new Set(h).size!==h.length)autoFail_('MASTER_SCHEMA');
  const price=h.includes('定価１')?'定価１':'定価1';
  const req=['コード','商品名','仕入先CD','仕入先名','JANCD','売上単価',price,'仕入単価','単位名','最終売上日','廃番'];
  if(req.some(x=>!h.includes(x)))autoFail_('MASTER_SCHEMA');const codes=new Set(),rows=[];
  csv.slice(1).forEach(v=>{if(v.every(x=>String(x).trim()===''))return;if(v.length!==h.length)autoFail_('MASTER_INCOMPLETE');
    const r={};h.forEach((k,i)=>r[k]=String(v[i]).trim());const code=bridgeNormalizeCode_(r['コード']);
    if(!/^\d{1,16}$/.test(code)||codes.has(code)||!r['商品名'])autoFail_('MASTER_CODE_INVALID');codes.add(code);
    const releaseHeader=['発売日','発売年月日'].find(x=>h.includes(x));
    rows.push({code:code,name:r['商品名'],jan:bridgeNormalizeJan_(r['JANCD']),supplierCd:bridgeNormalizeCode_(r['仕入先CD']),supplierName:r['仕入先名'],
      unitPrice:autoMasterNumber_(r['売上単価']),jodai:autoMasterNumber_(r[price]),shiire:autoMasterNumber_(r['仕入単価']),unit:r['単位名'],
      lastSaleDate:bridgeMasterDate_(r['最終売上日']),releaseDate:releaseHeader?bridgeMasterDate_(r[releaseHeader]):'',discontinued:!['','0','FALSE'].includes(r['廃番'].toUpperCase())});
  });
  if(rows.length<1000)autoFail_('MASTER_UNEXPECTED_SMALL_REVIEW_REQUIRED');
  if(f.getLastUpdated().getTime()!==updatedAt||f.getSize()!==size)autoFail_('MASTER_CHANGED');
  return {complete:true,masterHash:hash,generation:hash+':'+updatedAt,updatedAt:updatedAt,sizeBytes:size,rowCount:rows.length,rows:rows};
}
function autoMasterNumber_(raw) {
  const s=String(raw).trim();if(s==='')return null;
  if(!/^(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d{1,4})?$/.test(s))autoFail_('MASTER_NUMBER_INVALID');
  const n=Number(s.replace(/,/g,''));if(!Number.isFinite(n)||n<0||n>1000000000)autoFail_('MASTER_NUMBER_INVALID');return n;
}
function autoAudit_(sheets) {
  const reads={};['/products','/product_sets','/categories','/product_features'].forEach(path=>reads[path]=bridgeAuditCatalog_(path));
  const products=reads['/products'].rows,sets=reads['/product_sets'].rows,ids=new Set(products.map(x=>String(x.id)));
  if(products.some(x=>typeof x.name!=='string'||!['表示','非表示'].includes(x.flag)||![x.product_no,x.main_no,x.jan_code].every(bridgeScalarKey_))||
    sets.some(x=>!ids.has(String(x.product_id))||typeof x.name!=='string'||!['表示','非表示'].includes(x.set_flag)||![x.product_no,x.jan_code].every(bridgeScalarKey_)))autoFail_('CATALOG_COVERAGE');
  const identity=(ps,ss)=>({products:ps.map(x=>[x.id,x.product_no,x.main_no,x.jan_code,x.name,x.flag,x.category_id,x.feature_id1,x.feature_id2,x.feature_id3]),sets:ss.map(x=>[x.id,x.product_id,x.product_no,x.jan_code,x.name,x.set_flag])});
  const p2=bridgeAuditCatalog_('/products'),s2=bridgeAuditCatalog_('/product_sets');if(bridgeHash_(identity(products,sets))!==bridgeHash_(identity(p2.rows,s2.rows)))autoFail_('CATALOG_CHANGED');
  const index={code:Object.create(null),jan:Object.create(null),name:Object.create(null)},parents=new Map(),counts={},orphans=[];
  const add=(kind,id,c,j,n)=>{[[index.code,bridgeNormalizeCode_(c)],[index.jan,bridgeNormalizeJan_(j)],[index.name,bridgeNormalizeName_(n)]].forEach(a=>{if(a[1]){if(!a[0][a[1]])a[0][a[1]]=[];a[0][a[1]].push({kind:kind,id:String(id)});}});};
  products.forEach(x=>{add('product',x.id,x.product_no,x.jan_code,x.name);if(x.main_no&&bridgeNormalizeCode_(x.main_no)!==bridgeNormalizeCode_(x.product_no))add('product',x.id,x.main_no,'','');});
  sets.forEach(x=>add('productSet',x.id,x.product_no,x.jan_code,x.name));
  sheets.parents.values.slice(1).forEach((r,i)=>{if(r.every(v=>v===''))return;const id=String(r[0]||'');if(!id||parents.has(id))orphans.push('PARENT_ID');parents.set(id,r);add('draft',id,'','',r[4]);});
  sheets.children.values.slice(1).forEach(r=>{if(r.every(v=>v===''))return;const id=String(r[0]||'');if(!id||!parents.has(id))orphans.push('MISSING_PARENT');counts[id]=(counts[id]||0)+1;add('draftSet',id,r[1],r[3],r[2]);});
  parents.forEach((r,id)=>{if(!counts[id]&&!bridgeVerifiedArchive_(r,sheets))orphans.push('MISSING_SETS');});
  if(orphans.length)autoFail_('ORPHAN_ROWS_REQUIRE_REVIEW');
  if(bridgeHash_(sheets)!==bridgeHash_(bridgeReadSheets_()))autoFail_('SHEETS_CHANGED');
  return {index:index,products:products,sets:sets,categories:reads['/categories'].rows.map(x=>({id:Number(x.id),name:String(x.name)})),features:reads['/product_features'].rows.map(x=>({id:Number(x.id),name:String(x.name)})),evidence:Object.keys(reads).map(k=>reads[k].evidence).concat([p2.evidence,s2.evidence]),complete:true,includesHidden:true};
}
function autoMatches_(audit,r,names) {
  const all=[...(audit.index.code[r.code]||[]),...(r.jan?audit.index.jan[r.jan]||[]:[])];
  [r.name].concat(names||[]).forEach(n=>all.push(...(audit.index.name[bridgeNormalizeName_(n)]||[])));return all;
}
function autoSnapshot_() {
  const m=autoMaster_(),s=bridgeReadSheets_(),a=autoAudit_(s),again=autoMaster_();if(m.generation!==again.generation)autoFail_('MASTER_CHANGED');
  return {ok:true,complete:true,includesHidden:true,masterHash:m.masterHash,generation:m.generation,updatedAt:m.updatedAt,sizeBytes:m.sizeBytes,rowCount:m.rowCount,
    rows:m.rows.map(r=>Object.assign({},r,{matches:autoMatches_(a,r)})),categories:a.categories,features:a.features,catalogEvidence:a.evidence,notificationEnabled:false};
}
function autoValidate_(d,m,a) {
  autoExact_(d,'draftType targetProductId productName categoryId featureId1 featureId2 featureId3 description confidence reasoning refUrls supplierCd supplierName sets jodaiType taxTypeId viewGroupRestricted productTag'.split(' '));
  if(d.draftType!=='new_product'||d.targetProductId!==null)autoFail_('NEW_DRAFT_ONLY');
  if(!bridgeSafeText_(d.productName,200)||!bridgeSafeText_(d.description,2000)||Array.from(d.description).length>150||!bridgeSafeText_(d.reasoning,2000)||!['高','中','低'].includes(d.confidence))autoFail_('INVALID_TEXT');
  if(!['メーカー希望小売価格','非表示','オープン'].includes(d.jodaiType)||![1,2,3].includes(d.taxTypeId)||typeof d.viewGroupRestricted!=='boolean')autoFail_('INVALID_ATTRIBUTES');
  if(!normalizeDraftProductTag_(d.productTag).ok||!Array.isArray(d.refUrls)||d.refUrls.length>5||d.refUrls.some(u=>typeof u!=='string'||u.length>1000||!/^https:\/\/[^\s]+$/.test(u)))autoFail_('INVALID_REFERENCE');
  if(!Number.isSafeInteger(d.categoryId)||!a.categories.some(x=>x.id===d.categoryId))autoFail_('CATEGORY_UNVERIFIED');
  ['featureId1','featureId2','featureId3'].forEach(k=>{if(d[k]!==null&&(!Number.isSafeInteger(d[k])||!a.features.some(x=>x.id===d[k])))autoFail_('FEATURE_UNVERIFIED');});
  if(!Array.isArray(d.sets)||!d.sets.length||d.sets.length>30)autoFail_('INVALID_SETS');const seen=new Set();
  d.sets.forEach(s=>{autoExact_(s,'code setName jan unitPrice jodai shiire unit lastSaleDate'.split(' '));if(typeof s.code!=='string'||seen.has(s.code))autoFail_('DUPLICATE_SET_CODE');seen.add(s.code);
    const r=m.rows.find(x=>x.code===s.code);if(!r||r.discontinued)autoFail_('MASTER_TARGET_INVALID');
    if([r.unitPrice,r.jodai,r.shiire].some(x=>x===null))autoFail_('MASTER_PRICE_UNKNOWN');
    if(d.supplierCd!==r.supplierCd||d.supplierName!==r.supplierName||s.jan!==r.jan||s.unitPrice!==r.unitPrice||s.jodai!==r.jodai||s.shiire!==r.shiire||s.unit!==r.unit||s.lastSaleDate!==r.lastSaleDate)autoFail_('MASTER_PAYLOAD_MISMATCH');
    if(!bridgeSafeText_(s.setName,200)||!bridgeSafeText_(s.unit,30))autoFail_('INVALID_SET_TEXT');
    if(autoMatches_(a,r,[d.productName,s.setName]).length)autoFail_('POSSIBLE_DUPLICATE');
  });
  if(d.sets.length===1&&d.productName!==d.sets[0].setName)autoFail_('SINGLE_SET_NAME');return Array.from(seen).sort();
}
function autoEnvelopeValid_(e,signature,key) {
  autoExact_(e,'version proposalId ownerId codes payloadHash masterHash approvedAt expiresAt nonce'.split(' '));
  if(e.version!==1||e.ownerId!==autoOwner_()||typeof e.nonce!=='string'||!/^[0-9a-f-]{36}$/.test(e.nonce)||!Number.isSafeInteger(e.approvedAt)||Date.now()-e.approvedAt<0||Date.now()-e.approvedAt>120000||!Number.isSafeInteger(e.expiresAt)||e.expiresAt<=Date.now()||e.expiresAt-e.approvedAt>7200000)autoFail_('APPROVAL_INVALID');
  const expected=Utilities.computeHmacSha256Signature(bridgeStable_(e),key,Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
  if(!bridgeKeyMatches_(signature,expected))autoFail_('APPROVAL_SIGNATURE_INVALID');
}
function autoParentRow_(p,id) {
  const d=p.payload;return [id,'下書き','new_product','',d.productName,d.categoryId,d.featureId1===null?'':d.featureId1,d.featureId2===null?'':d.featureId2,d.featureId3===null?'':d.featureId3,d.description,d.confidence,d.reasoning,d.refUrls.join('\n'),d.supplierCd,d.supplierName,new Date(p.attemptedAt).toISOString(),'','',d.jodaiType,d.taxTypeId,d.viewGroupRestricted?'TRUE':'FALSE',normalizeDraftProductTag_(d.productTag).tag];
}
function autoSetRows_(p,id) {return p.payload.sets.map(s=>[id,s.code,s.setName,s.jan,s.unitPrice,s.jodai,s.shiire,s.unit,s.lastSaleDate]);}
function autoVerify_(before,after,p) {
  if(bridgeBaselineHash_(after,p.draftId)!==bridgeBaselineHash_(before,null))autoFail_('EXISTING_ROWS_CHANGED');
  const ps=after.parents.values.slice(1).filter(r=>String(r[0])===p.draftId),ss=after.children.values.slice(1).filter(r=>String(r[0])===p.draftId);
  const canon=r=>r.map((x,i)=>i===20?(parseViewGroupRestricted(x)?'TRUE':'FALSE'):String(x==null?'':x));
  if(ps.length!==1||bridgeStable_(canon(ps[0]))!==bridgeStable_(canon(autoParentRow_(p,p.draftId)))||bridgeStable_(ss.map(canon))!==bridgeStable_(autoSetRows_(p,p.draftId).map(canon)))autoFail_('READBACK_MISMATCH');
  [after.parents,after.children].forEach(t=>t.values.forEach((r,i)=>{if(String(r[0])===p.draftId&&t.formulas[i].some(Boolean))autoFail_('TARGET_FORMULA');}));
}
function autoSave_(p) {
  if(p.state==='VERIFIED')return autoPublic_(p);if(['IN_FLIGHT','UNKNOWN'].includes(p.state))autoFail_('OUTCOME_UNKNOWN_NO_RESEND');
  if(p.state!=='APPROVED'||p.ownerId!==autoOwner_()||p.approvalExpiresAt<=Date.now()||p.expiresAt<=Date.now())autoFail_('APPROVAL_REQUIRED');
  const proof=p.approvalEnvelope;
  if(!proof||proof.proposalId!==p.id||proof.ownerId!==autoOwner_()||bridgeStable_(proof.codes)!==bridgeStable_(p.codes)||proof.payloadHash!==p.payloadHash||proof.masterHash!==p.masterHash||proof.expiresAt!==p.approvalExpiresAt||!bridgeKeyMatches_(p.approvalSignature,autoMac_(proof)))autoFail_('APPROVAL_PROOF_INVALID');
  const m=autoMaster_();if(m.masterHash!==p.masterHash){p.state='INVALID';autoPut_(p);autoFail_('MASTER_CHANGED_APPROVAL_INVALID');}
  const before=bridgeReadSheets_(),a=autoAudit_(before);if(bridgeStable_(autoValidate_(p.payload,m,a))!==bridgeStable_(p.codes))autoFail_('PAYLOAD_CHANGED');
  if(autoMaster_().generation!==m.generation)autoFail_('MASTER_CHANGED');
  const today=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd'),prefix='D'+today+'-';let seq=0;
  const claim=id=>{if(typeof id==='string'&&id.startsWith(prefix)&&/^D\d{8}-\d{3}$/.test(id))seq=Math.max(seq,Number(id.slice(-3)));};
  before.parents.values.slice(1).forEach(r=>claim(String(r[0])));
  const journal=autoJournal_(false);if(journal)journal.getDataRange().getValues().slice(1).forEach(r=>{let x;try{x=JSON.parse(String(r[1]));}catch(e){autoFail_('JOURNAL_CORRUPT');}claim(x.draftId);});
  const old=bridgeReadJournal_();if(old)claim(old.draftId);if(seq>=999)autoFail_('DAILY_ID_EXHAUSTED');
  p.draftId=prefix+('00'+(seq+1)).slice(-3);p.state='IN_FLIGHT';p.attemptedAt=Date.now();
  // 独立journalの確保・耐久再読が完了するまで商品行を追記しない。
  autoPut_(p);
  try {
    const ss=getDataSpreadsheet_(),parents=ss.getSheetByName(SHEET_DRAFT),sets=ss.getSheetByName(SHEET_DRAFT_SETS);
    if(draftProductTagColumn_(parents,false)!==22)autoFail_('DRAFT_TAG_SCHEMA');
    const pr=parents.getLastRow()+1,sr=sets.getLastRow()+1,setRows=autoSetRows_(p,p.draftId);
    parents.getRange(pr,1,1,22).setNumberFormat('@');parents.getRange(pr,1,1,22).setValues([autoParentRow_(p,p.draftId)]);
    sets.getRange(sr,1,setRows.length,9).setNumberFormat('@');sets.getRange(sr,1,setRows.length,9).setValues(setRows);
    SpreadsheetApp.flush();autoVerify_(before,bridgeReadSheets_(),p);p.state='VERIFIED';autoPut_(p);return autoPublic_(p);
  }catch(e){p.state='UNKNOWN';try{autoPut_(p);}catch(ignored){}autoFail_('OUTCOME_UNKNOWN_NO_RESEND');}
}
function handleBmallAuto_(params) {
  let lock;try{
    if(!BMALL_AUTO_ACTIONS_.includes(params.action))autoFail_('ACTION_NOT_ALLOWED');const key=PropertiesService.getScriptProperties().getProperty('BMALL_BRIDGE_API_KEY');
    if(!bridgeKeyMatches_(params.apiKey,key))autoFail_('UNAUTHORIZED');autoOwner_();
    if(params.action==='bmallAutoStatus'){autoExact_(params,['action','apiKey']);return {ok:true,version:1,gasVersion:VERSION,approval:'owner_oidc_review_post',notificationEnabled:false};}
    if(params.action==='bmallAutoSnapshot'){autoExact_(params,['action','apiKey']);return autoSnapshot_();}
    if(params.action==='bmallAutoReferences'){
      autoExact_(params,['action','apiKey','supplierCd','maxProducts']);if(typeof params.supplierCd!=='string'||!/^\d{1,16}$/.test(params.supplierCd)||!Number.isSafeInteger(params.maxProducts)||params.maxProducts<1||params.maxProducts>80)autoFail_('INVALID_REFERENCE_REQUEST');
      const m=autoMaster_(),a=autoAudit_(bridgeReadSheets_()),codes=new Set(m.rows.filter(x=>x.supplierCd===params.supplierCd).map(x=>x.code));
      const matchingSets=a.sets.filter(x=>codes.has(bridgeNormalizeCode_(x.product_no))),productIds=new Set(matchingSets.map(x=>String(x.product_id)));
      if(autoMaster_().generation!==m.generation)autoFail_('MASTER_CHANGED');
      return {ok:true,masterHash:m.masterHash,complete:true,includesHidden:true,categories:a.categories,features:a.features,
        examples:a.products.filter(x=>productIds.has(String(x.id))).slice(0,params.maxProducts).map(x=>({productId:x.id,productName:x.name,categoryId:x.category_id,featureIds:[x.feature_id1,x.feature_id2,x.feature_id3],sets:matchingSets.filter(s=>String(s.product_id)===String(x.id)).map(s=>({setId:s.id,productNo:s.product_no,setName:s.name}))})),totalMatchingProducts:productIds.size};
    }
    lock=LockService.getScriptLock();if(!lock.tryLock(1000))autoFail_('BUSY_NO_WRITE');
    if(params.action==='bmallAutoPrepare'){
      autoExact_(params,['action','apiKey','proposalId','draft']);autoId_(params.proposalId);if(autoGet_(params.proposalId))autoFail_('PROPOSAL_EXISTS');
      const m=autoMaster_(),a=autoAudit_(bridgeReadSheets_()),codes=autoValidate_(params.draft,m,a);if(autoMaster_().generation!==m.generation)autoFail_('MASTER_CHANGED');
      const d=params.draft,p={version:1,id:params.proposalId,payload:d,payloadHash:bridgeHash_(d),masterHash:m.masterHash,codes:codes,state:'PROPOSED',createdAt:Date.now(),expiresAt:Date.now()+86400000,
        reviewLabels:{categoryName:(a.categories.find(x=>x.id===d.categoryId)||{}).name||'ID '+d.categoryId,features:[d.featureId1,d.featureId2,d.featureId3].map(id=>id===null?'未指定':(a.features.find(x=>x.id===id)||{}).name||'ID '+id)}};autoPut_(p);return autoPublic_(p);
    }
    if(params.action==='bmallAutoApprove'){
      autoExact_(params,['action','apiKey','envelope','signature']);autoEnvelopeValid_(params.envelope,params.signature,key);const e=params.envelope,g=autoGet_(e.proposalId);if(!g)autoFail_('PROPOSAL_MISSING');const p=g.p;
      if(p.state!=='PROPOSED'||p.expiresAt<=Date.now()||bridgeStable_(p.codes)!==bridgeStable_(e.codes)||p.payloadHash!==e.payloadHash||p.masterHash!==e.masterHash||autoMaster_().masterHash!==p.masterHash)autoFail_('APPROVAL_STALE');
      p.ownerId=e.ownerId;p.approvedAt=e.approvedAt;p.approvalExpiresAt=e.expiresAt;p.approvalNonce=e.nonce;p.approvalEnvelope=e;p.approvalSignature=params.signature;p.state='APPROVED';autoPut_(p);return autoPublic_(p);
    }
    autoExact_(params,['action','apiKey','proposalId']);const got=autoGet_(params.proposalId);if(!got){if(params.action==='bmallAutoResult')return {ok:true,state:'NOT_FOUND',attempted:false};autoFail_('PROPOSAL_MISSING');}
    if(params.action==='bmallAutoSave')return autoSave_(got.p);
    if(params.action==='bmallAutoProposal'||params.action==='bmallAutoResult')return autoPublic_(got.p);autoFail_('ACTION_NOT_ALLOWED');
  }catch(e){return {ok:false,error:e&&/^AUTO_[A-Z0-9_]+$/.test(e.message)?e.message:e&&/^BRIDGE_[A-Z0-9_]+$/.test(e.message)?'AUTO_'+e.message:'AUTO_INTERNAL_FAILURE'};}finally{if(lock)lock.releaseLock();}
}
