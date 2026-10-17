// 日本語コメント: 本人方針により共通ポータルを継続。本人ID・アプリ権限は毎要求サーバーで照合する。
export const VERSION='0.2.0';
const permissions=Object.freeze(['applications.read','applications.write','booking.read','booking.write','checkins.read','checkins.write']);
const tokenValid=token=>typeof token==='string'&&/^[A-Za-z0-9_-]{32,256}$/.test(token);
export function createPortalStaffVerifier({portalUrl,fetchImpl=fetch}={}){
 if(!portalUrl)throw Error('AUTH_NOT_CONFIGURED');
 const url=new URL(portalUrl);
 if(url.protocol!=='https:'||url.hostname!=='script.google.com'||!/^\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url.pathname)||url.search||url.hash||url.username||url.password)throw Error('INVALID_PORTAL_URL');
 return async token=>{
  if(!tokenValid(token))return null;
  let identity;
  // 認証の読取だけを最大2回。拒否応答は再試行せず、成功結果を別要求へキャッシュしない。
  for(let attempt=0;attempt<2;attempt++){
   try{
    const response=await fetchImpl(url,{method:'POST',redirect:'follow',signal:AbortSignal.timeout(15000),headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'validateSession',token,app_name:'beaufes'})});
    if(!response.ok)throw Error('AUTH_PROVIDER_UNAVAILABLE');
    identity=await response.json();
    if(!identity||typeof identity!=='object'||Array.isArray(identity)||typeof identity.ok!=='boolean')throw Error('AUTH_PROVIDER_UNAVAILABLE');
    break;
   }catch{
    if(attempt===1)throw Error('AUTH_PROVIDER_UNAVAILABLE');
    await new Promise(resolve=>setTimeout(resolve,250));
   }
  }
  if(identity.ok!==true)return null;
  if(typeof identity.user_id!=='string'||!/^[A-Za-z0-9_-]{1,64}$/.test(identity.user_id)||!['user','admin'].includes(identity.role))return null;
  // 現行受付GASもbeaufes user/adminに社員操作を認めている。全社is_adminから昇格させない。
  return {user_id:identity.user_id,name:typeof identity.name==='string'?identity.name:identity.user_id,permissions:[...permissions],auth_source:'portal',role:identity.role};
 };
}

// ブラウザー側は既存bf_sessionを読み、検証後の表示用identityだけをメモリー保持する。元セッションは変更しない。
export function createPortalStaffProvider({storage,origin,portalUrl,endpoint,fetchImpl=fetch,navigate,now=()=>Date.now()}={}){
 if(!storage||!origin||!portalUrl||!endpoint)throw Error('PORTAL_NOT_CONFIGURED');
 const portal=new URL(portalUrl),api=new URL(endpoint);
 if(portal.protocol!=='https:'||portal.username||portal.password||portal.search||portal.hash||portal.pathname!=='/beaufield-dev/')throw Error('INVALID_PORTAL_URL');
 if(api.protocol!=='https:'||api.username||api.password||api.search||api.hash)throw Error('INVALID_BACKEND_URL');
 let verified=null;
 return {getSession:async()=>{
  // 別originのlocalhostから、ポータルのbf_sessionを手入力・コピーさせない。
  if(origin!==portal.origin)return null;
  let session;try{session=JSON.parse(storage.getItem('bf_session')||'null');}catch{return null;}
  if(!session||!tokenValid(session.token)||!Number.isSafeInteger(session.expires)||session.expires<=now())return null;
  if(verified?.token!==session.token){
   let response;try{response=await fetchImpl(api,{method:'POST',redirect:'error',cache:'no-store',signal:AbortSignal.timeout(45000),headers:{'Content-Type':'application/json',Authorization:'Bearer '+session.token},body:JSON.stringify({action:'staffSession',data:{}})});}catch{throw Error('AUTH_PROVIDER_UNAVAILABLE');}
   let result;try{result=await response.json();}catch{throw Error('AUTH_PROVIDER_UNAVAILABLE');}
   if(response.status===401||response.status===403)return null;
   if(!response.ok||result?.success!==true)throw Error('AUTH_PROVIDER_UNAVAILABLE');
   const identity=result.data;
   if(!identity||identity.auth_source!=='portal'||typeof identity.user_id!=='string'||!Array.isArray(identity.permissions))throw Error('AUTH_PROVIDER_UNAVAILABLE');
   verified={token:session.token,identity};
  }
  return {user_id:verified.identity.user_id,name:verified.identity.name,access_token:session.token,expires_at:session.expires};
 },signIn:async()=>{if(typeof navigate!=='function')throw Error('PORTAL_NOT_CONFIGURED');navigate(portal.href);}};
}
