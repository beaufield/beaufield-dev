/* シリアル検証用通信。通行証はメモリーのみ。秘密キー・本番URLは受け付けない。 */
(function(root){'use strict';
class SerialClient {
 constructor({url,publicKey,issueToken,session,fetchImpl}) {
  if(url!=='https://qxzdypqbqduwuwmwnbgk.supabase.co'||!/^sb_publishable_/.test(publicKey))throw Error('TEST_CONFIG_REQUIRED');
  this.url=url;this.key=publicKey;this.issue=issueToken;this.session=session;this.fetch=fetchImpl||root.fetch.bind(root);this.jwt=null;this.pending=null;
 }
 async rpc(name,args){
  const s=this.session();if(!s?.token||!s.user_id||!(s.expires>Date.now()))throw Error('SESSION_INVALID');
  if(!this.jwt||this.jwt.session!==s.token||this.jwt.exp*1000<=Date.now()+5000){const r=await this.issue(s.token);if(!r?.success)throw Error(r?.error||'AUTH_UNAVAILABLE');this.jwt={...r,session:s.token};}
  const abort=new AbortController(),timer=setTimeout(()=>abort.abort(),20000);
  try{
   const r=await this.fetch(this.url+'/rest/v1/rpc/sr_'+name,{method:'POST',signal:abort.signal,headers:{apikey:this.key,Authorization:'Bearer '+this.jwt.token,'Content-Type':'application/json','Content-Profile':'public','Accept-Profile':'public'},body:JSON.stringify(args)});
   const data=await r.json();if(this.session()?.token!==s.token)throw Error('SESSION_CHANGED');if(!r.ok){if(r.status===401)this.jwt=null;throw Error(data?.message||'DB_HTTP_'+r.status);}return data;
  }finally{clearTimeout(timer);}
 }
 page(filter={},cursor={}){return this.rpc('page',{p_filter:filter,p_after:cursor.after||0,p_highwater:cursor.highwater??null,p_revision:cursor.revision??null,p_csv:false});}
 async csv(filter,write){
  // 途中更新・通信失敗時には呼出側の一時出力を破棄する。完了まで公開しない。
  let cursor={after:0},count=0;
  for(;;){const p=await this.rpc('page',{p_filter:filter,p_after:cursor.after,p_highwater:cursor.highwater??null,p_revision:cursor.revision??null,p_csv:true});if(!Array.isArray(p.rows)||p.rows.length>500)throw Error('INVALID_PAGE');await write(p.rows);count+=p.rows.length;if(p.done)return count;if(p.after<=cursor.after)throw Error('CURSOR_STALLED');cursor=p;}
 }
 async mutate(kind,payload,requestId){
  if(!requestId)throw Error('REQUEST_ID_REQUIRED');
  // 結果不明のときも同じUUID・内容で再照会できる。新しいUUIDを勝手に作らない。
  return this.rpc('mutate',{p_id:requestId,p_kind:kind,p_payload:payload});
 }
}
if(typeof module!=='undefined')module.exports={SerialClient};else root.SerialClient=SerialClient;
})(typeof window!=='undefined'?window:globalThis);
