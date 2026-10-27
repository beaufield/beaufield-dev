import {EVENT_DATE,japanTime} from './attendance-core.mjs?b=20261027a';
// 日本語コメント: 時刻不明は既存の自由記述時刻で保持し、実際の来場時刻を捏造しない。
export const UNKNOWN_TIME=EVENT_DATE+' 時刻不明';
const prefix='attendance-manual-v1:';
export function receiptTime(value){
 if(!value)return UNKNOWN_TIME;
 if(!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(value)||!japanTime(value)||value.slice(0,10)!==EVENT_DATE)throw Error('INVALID_ARRIVAL_TIME');
 return value.replace('T',' ')+':00';
}
export function pendingCheckin(storage,appId){
 const raw=storage.getItem(prefix+appId);if(!raw)return null;
 let row;try{row=JSON.parse(raw);}catch{throw Error('PENDING_INVALID');}
 if(!row||row.entry_code!==appId||!/^manual-[A-Za-z0-9_-]{8,93}$/.test(row.checkin_id)||row.device_id!=='attendance-manual'||row.note!=='社員用来場状況から手動受付（本人通知なし）'||typeof row.checked_in_at!=='string'||!(row.checked_in_at===UNKNOWN_TIME||receiptTime(row.checked_in_at.slice(0,16).replace(' ','T'))===row.checked_in_at))throw Error('PENDING_INVALID');
 return row;
}
export function prepareCheckin(storage,appId,value,uuid=()=>crypto.randomUUID()){
 const old=pendingCheckin(storage,appId);if(old)return old;
 const row={checkin_id:'manual-'+uuid(),entry_code:appId,checked_in_at:receiptTime(value),device_id:'attendance-manual',note:'社員用来場状況から手動受付（本人通知なし）'};
 // 通信前にIDと本文を永続化する。再読込・応答喪失時も同じIDだけを再送する。
 try{storage.setItem(prefix+appId,JSON.stringify(row));
  if(storage.getItem(prefix+appId)!==JSON.stringify(row))throw Error('PENDING_STORAGE_FAILED');
 }catch{throw Error('PENDING_STORAGE_FAILED');}
 return row;
}
export function clearCheckin(storage,appId){storage.removeItem(prefix+appId);}
export function validateCheckinResult(json,row){
 const results=json?.data?.results;
 return json?.success===true&&Array.isArray(results)&&results.length===1&&results[0].checkin_id===row.checkin_id&&results[0].app_id===row.entry_code&&results[0].resolve_status==='ok'&&['recorded','duplicate'].includes(results[0].status);
}
