// 日本語コメント: 当日登録の段階と操作IDを保持する。応答不明・再読込でも作成済みの申込を作り直さない。
export const VERSION='0.2.0';
export const BUSINESS_TYPES=Object.freeze(['美容室','理容室','エステサロン','ネイルサロン','アイサロン','その他']);
const phases=new Set(['creating','registered','assigning','assigned','enqueuing','done','cancelling']);
const copy=x=>x===null?null:JSON.parse(JSON.stringify(x));
const fail=code=>{throw Error(code);};
export const normalized=value=>String(value||'').normalize('NFKC').replace(/\s+/g,'').toLowerCase();
export function possibleMatches(roster,fields){return (roster?.roster||[]).filter(r=>normalized(r.staff_name)===normalized(fields.staff_name)).slice(0,12);}
export function createWalkin({storage,call,enqueue,withLock,uuid=()=>crypto.randomUUID(),now=()=>new Date().toISOString(),key='walkin-journal-v1'}={}){
 if(!storage||!call||!enqueue||!withLock)fail('WALKIN_NOT_CONFIGURED');
 const read=()=>{
  let raw;try{raw=storage.getItem(key);}catch{fail('WALKIN_STORAGE_FAILED');}
  if(raw===null)return null;
  let s;try{s=JSON.parse(raw);}catch{fail('WALKIN_RECOVERY_INVALID');}
  if(!s||s.version!==1||!phases.has(s.phase)||typeof s.owner!=='string'||!s.owner||!s.fields||typeof s.fields.staff_name!=='string'||typeof s.fields.salon_name!=='string'||typeof s.fields.business_type!=='string')fail('WALKIN_RECOVERY_INVALID');
  const phase=s.phase==='cancelling'?s.cancelFrom:s.phase;
  if(s.phase==='cancelling'&&(!['creating','registered','assigning','assigned'].includes(phase)||typeof s.cancelId!=='string'))fail('WALKIN_RECOVERY_INVALID');
  if(phase==='creating'&&typeof s.createId!=='string'||phase!=='creating'&&typeof s.appId!=='string')fail('WALKIN_RECOVERY_INVALID');
  if(['assigning','assigned','enqueuing','done'].includes(phase)&&(!/^P-\d{2}$/.test(s.spareNo)||typeof s.assignId!=='string'))fail('WALKIN_RECOVERY_INVALID');
  if(['enqueuing','done'].includes(s.phase)&&(!s.checkinId||!s.checkedAt))fail('WALKIN_RECOVERY_INVALID');
  return s;
 };
 const save=s=>{try{storage.setItem(key,JSON.stringify(s));if(storage.getItem(key)!==JSON.stringify(s))fail('WALKIN_STORAGE_FAILED');}catch{fail('WALKIN_STORAGE_FAILED');}return copy(s);};
 const actor=async owner=>{const who=await call('staffSession',{});if(!who?.user_id)fail('SESSION_INVALID');if(owner&&owner!==who.user_id)fail('WALKIN_OWNER_CHANGED');return who.user_id;};
 const locked=run=>withLock(async()=>run(read()));
 async function create(s,ownerChecked=false){
  if(!ownerChecked)await actor(s.owner);
  const r=await call('adminCreateApplication',{...s.fields,request_id:s.createId});
  if(!r?.app_id)fail('WALKIN_RESPONSE_INVALID');
  return save({...s,phase:'registered',appId:r.app_id});
 }
 return {
  state:()=>copy(read()),
  async register(fields,{reviewed=false}={}){return locked(async s=>{
   if(s){
    if(fields&&['staff_name','salon_name','business_type'].some(k=>String(fields[k]||'').trim()!==s.fields[k]))fail('WALKIN_IN_PROGRESS');
    if(s.phase==='cancelling')fail('WALKIN_CANCELLING');if(s.phase==='creating')return create(s);return s;
   }
   const clean=Object.fromEntries(['staff_name','salon_name','business_type'].map(k=>[k,String(fields?.[k]||'').trim()]));
   if(!clean.staff_name||clean.staff_name.length>60||!clean.salon_name||clean.salon_name.length>100||!BUSINESS_TYPES.includes(clean.business_type))fail('INVALID_APPLICANT');
   // 確認画面で候補を見た場合は名簿を二重取得しない。認可は各APIで継続する。
   const owner=await actor();
   const roster=reviewed?null:await call('scanRoster',{});
   if(!reviewed&&possibleMatches(roster,clean).length)fail('WALKIN_EXISTING_REVIEW_REQUIRED');
   // QRやメール・電話、認証情報は作業台帳へ保存しない。
   s=save({version:1,phase:'creating',owner,fields:clean,createId:uuid()});return create(s,true);
  });},
  async useExisting(appId){return locked(async s=>{
   if(s)fail('WALKIN_IN_PROGRESS');const owner=await actor();const roster=await call('scanRoster',{});const entry=roster?.roster?.find(r=>r.app_id===appId);
   if(!entry)fail('APP_NOT_FOUND');return save({version:1,phase:'registered',owner,appId,fields:{staff_name:entry.staff_name??'',salon_name:entry.salon_name??'',business_type:entry.business_type??''}});
  });},
  async spares(){const s=read();if(!s?.appId)fail('WALKIN_NOT_REGISTERED');await actor(s.owner);return (await call('listSpareBadges',{})).spares||[];},
  async assign(spareNo){return locked(async s=>{
   if(!s)fail('WALKIN_NOT_REGISTERED');if(['assigned','enqueuing','done'].includes(s.phase))return s;
   if(!['registered','assigning'].includes(s.phase))fail('WALKIN_NOT_REGISTERED');await actor(s.owner);
   if(s.phase==='registered'){
    if(!/^P-\d{2}$/.test(spareNo||''))fail('SPARE_NOT_FOUND');
    // 最新の札の競合・使用済み判定は割当RPCが同じトランザクションで行う。
    s=save({...s,phase:'assigning',spareNo,assignId:uuid()});
   }
   try{
    const r=await call('assignSpare',{app_id:s.appId,spare_no:s.spareNo,request_id:s.assignId});
    if(r?.app_id!==s.appId||r?.spare_no!==s.spareNo)fail('WALKIN_RESPONSE_INVALID');
    return save({...s,phase:'assigned'});
   }catch(e){
    // 札の競合が確定した場合だけ再選択へ戻す。通信失敗では同じ操作を再送する。
    if(['SPARE_ALREADY_ASSIGNED','SPARE_RETIRED','SPARE_NOT_FOUND'].includes(e.message))save({...s,phase:'registered',spareNo:undefined,assignId:undefined});
    throw e;
   }
  });},
  async checkin(){return locked(async s=>{
   if(!s||!['assigned','enqueuing','done'].includes(s.phase))fail('WALKIN_NOT_ASSIGNED');if(s.phase==='done')return s;await actor(s.owner);
   const roster=await call('scanRoster',{});const entry=roster?.roster?.find(r=>r.app_id===s.appId);const link=roster?.spares?.find(r=>r.spare_no===s.spareNo&&r.app_id===s.appId);
   if(!entry||!link)fail('WALKIN_ROSTER_NOT_READY');
   if(s.phase==='assigned')s=save({...s,phase:'enqueuing',checkinId:uuid(),checkedAt:now()});
   const receipt=await enqueue({entry,entry_code:s.spareNo,checkin_id:s.checkinId,checked_in_at:s.checkedAt,roster});
   if(receipt?.checkin_id!==s.checkinId)fail('WALKIN_QUEUE_FAILED');
   return save({...s,phase:'done'});
  });},
  async cancel(){return locked(async s=>{
   if(!s)return null;
   if(!['creating','registered','assigning','assigned','cancelling'].includes(s.phase))fail('WALKIN_CANCEL_AFTER_CHECKIN');
   await actor(s.owner);
   if(s.phase!=='cancelling')s=save({...s,phase:'cancelling',cancelFrom:s.phase,cancelId:uuid()});
   let result;
   try{result=await call('cancelWalkin',{app_id:s.appId||'',create_request_id:s.createId||'',assign_request_id:s.assignId||'',spare_no:s.spareNo||'',request_id:s.cancelId});}
   catch(e){
    // 確定した拒否だけ元の段階へ戻す。応答不明は取消IDを維持して再確認する。
    if(['WALKIN_CANCEL_IN_USE','WALKIN_CANCEL_NOT_OWNED','WALKIN_CANCEL_REVIEW_REQUIRED','FORBIDDEN','EVENT_WRITE_CLOSED'].includes(e.message))save({...s,phase:s.cancelFrom,cancelFrom:undefined,cancelId:undefined});
    throw e;
   }
   if(result?.cancelled!==true)fail('WALKIN_RESPONSE_INVALID');
   try{storage.removeItem(key);if(storage.getItem(key)!==null)fail('WALKIN_STORAGE_FAILED');}catch{fail('WALKIN_STORAGE_FAILED');}
   return result;
  });},
  async next(){return locked(async s=>{if(s&&s.phase!=='done')fail('WALKIN_IN_PROGRESS');try{storage.removeItem(key);}catch{fail('WALKIN_STORAGE_FAILED');}if(storage.getItem(key)!==null)fail('WALKIN_STORAGE_FAILED');return null;});},
  hasUnfinished(){const s=read();return !!s&&s.phase!=='done';}
 };
}
