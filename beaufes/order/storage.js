export function mergeReceipt(record, receipt) {
  // 遅い成功応答で取消要求や新しい状態を巻き戻さない。
  if((receipt.state_seq||0)>=(record.receipt?.state_seq||0)) record.receipt=receipt;
  if(record.receipt?.state==='cancelled'){record.pending=null;record.cancelRequested=true;}
  else if(record.cancelRequested) record.pending='cancel';
  else if(record.receipt?.state==='active') record.pending=null;
  return record;
}
let connection;
export async function database(){if(connection)return connection;connection=await new Promise((resolve,reject)=>{const r=indexedDB.open('beaufes-booth-trial-v1',1);r.onupgradeneeded=()=>{r.result.createObjectStore('records',{keyPath:'id'});r.result.createObjectStore('settings');};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});return connection;}
export async function get(key){const db=await database();return new Promise((resolve,reject)=>{const r=db.transaction('settings').objectStore('settings').get(key);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});}
export async function set(key,value){const db=await database();return new Promise((resolve,reject)=>{const t=db.transaction('settings','readwrite');t.objectStore('settings').put(value,key);t.oncomplete=()=>resolve();t.onerror=()=>reject(t.error);});}
export async function clearMakerConnection(token){
 // 古いタブの失敗応答で、別タブが保存した新しい接続を消さない。
 const db=await database();return new Promise((resolve,reject)=>{const t=db.transaction('settings','readwrite');const store=t.objectStore('settings');const request=store.get('maker-connection');request.onsuccess=()=>{if(request.result?.token===token)store.put(null,'maker-connection');};t.oncomplete=()=>resolve();t.onerror=()=>reject(t.error);});
}
export async function records(){const db=await database();return new Promise((resolve,reject)=>{const r=db.transaction('records').objectStore('records').getAll();r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});}
export async function update(id, change){const db=await database();return new Promise((resolve,reject)=>{const t=db.transaction(['records','settings'],'readwrite');const store=t.objectStore('records');const request=store.get(id);let value;request.onsuccess=()=>{try{value=change(request.result);store.put(value);}catch(error){t.abort();reject(error);}};t.oncomplete=()=>resolve(value);t.onerror=()=>reject(t.error);});}
export async function sanitizeLegacyQr(){
 // 旧試作で端末に残した名札URLを削除する。注文本文のハッシュは維持する。
 for(const row of await records())if(row.form&&Object.hasOwn(row.form,'qr')){
  await update(row.id,current=>{const {qr,...form}=current.form;return {...current,form:{...form,qr_hash:current.body?.token_hash||form.qr_hash||'',resolved:false}};});
 }
}
