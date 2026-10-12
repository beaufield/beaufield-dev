// 保存領域を使えない端末でも、オンラインでサーバーの注文番号を確認する。
const memory={settings:new Map(),records:new Map()};
let connection,volatile=false;
const copy=value=>value===undefined?undefined:structuredClone(value);
export const storageMode=()=>volatile?'memory':'persistent';
export function mergeReceipt(record,receipt){
 if((receipt.state_seq||0)>=(record.receipt?.state_seq||0))record.receipt=receipt;
 if(record.receipt?.state==='cancelled'){record.pending=null;record.cancelRequested=true;}
 else if(record.cancelRequested)record.pending='cancel';
 else if(record.receipt?.state==='active')record.pending=null;
 return record;
}
export async function database(){
 if(volatile)return null;if(connection)return connection;
 try{connection=await new Promise((resolve,reject)=>{const request=indexedDB.open('beaufes-booth-trial-v1',1);request.onupgradeneeded=()=>{request.result.createObjectStore('records',{keyPath:'id'});request.result.createObjectStore('settings');};request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);request.onblocked=()=>reject(Error('STORAGE_BLOCKED'));});return connection;}
 catch{volatile=true;return null;}
}
async function transact(store,write,work){
 const db=await database();if(!db)return work(null);
 try{return await new Promise((resolve,reject)=>{const transaction=db.transaction(store,write?'readwrite':'readonly');let value,workError;work(transaction.objectStore(store),v=>{value=v;},error=>{workError=error;transaction.abort();});transaction.oncomplete=()=>resolve(value);transaction.onerror=()=>reject(transaction.error);transaction.onabort=()=>reject(workError?{applicationError:workError}:transaction.error||Error('STORAGE_ABORTED'));});}
 catch(error){if(error?.applicationError)throw error.applicationError;volatile=true;connection=null;return work(null);}
}
export async function get(key){return transact('settings',false,(store,done)=>{if(!store)return copy(memory.settings.get(key));const request=store.get(key);request.onsuccess=()=>{memory.settings.set(key,copy(request.result));done(request.result);};});}
export async function set(key,value){return transact('settings',true,(store,done)=>{memory.settings.set(key,copy(value));if(!store)return;store.put(value,key);done();});}
export async function records(){return transact('records',false,(store,done)=>{if(!store)return [...memory.records.values()].map(copy);const request=store.getAll();request.onsuccess=()=>{for(const row of request.result)memory.records.set(row.id,copy(row));done(request.result);};});}
export async function update(id,change){
 // 人物・数量など入力内容の世代を保存し、古いタブの入力で置き換えない。
 const apply=current=>{const before=JSON.stringify(current?.form),revision=current?.edit_revision||0;const value=change(current);if(value&&before!==JSON.stringify(value.form))value.edit_revision=revision+1;return value;};
 const value=await transact('records',true,(store,done,fail)=>{if(!store)return apply(copy(memory.records.get(id)));const request=store.get(id);request.onsuccess=()=>{try{const value=apply(request.result);store.put(value);done(value);}catch(error){fail(error);}};});
 memory.records.set(id,copy(value));return value;
}
export async function clearMakerConnection(token){return transact('settings',true,(store,done)=>{if(!store){if(memory.settings.get('maker-connection')?.token===token)memory.settings.set('maker-connection',null);return;}const request=store.get('maker-connection');request.onsuccess=()=>{if(request.result?.token===token){store.put(null,'maker-connection');memory.settings.set('maker-connection',null);}done();};});}
export async function sanitizeLegacyQr(){for(const row of await records())if(row.form&&Object.hasOwn(row.form,'qr'))await update(row.id,current=>{const {qr,...form}=current.form;return {...current,form:{...form,qr_hash:current.body?.token_hash||form.qr_hash||'',resolved:false}};});}
