// 日本語コメント: 隔離画面専用。未設定時は外部へ保存せず停止する。本番の名簿・未送信・セッションを読まない。
(function(){
 'use strict';
 const config=window.BF_EVENT_CANDIDATE_CONFIG||{};
 const endpoint=typeof config.endpoint==='string'&&config.environment==='isolated'&&typeof config.eventId==='string'&&config.eventId.trim()&&typeof config.backendGeneration==='string'&&config.backendGeneration.trim()?config.endpoint:'';
 const target=new URL(endpoint||'./__backend_not_configured__',location.href).href;
 const nativeFetch=window.fetch.bind(window);
 let staffProvider=window.BF_EVENT_STAFF_PROVIDER;
 const banner=document.createElement('div');
 banner.style.cssText='position:sticky;top:0;z-index:20000;padding:9px 15px;background:#fff0c9;color:#684213;font:13px/1.6 sans-serif;text-align:center';
 banner.textContent=endpoint?(config.fixture===true?'架空データの画面操作試験 ／ 実社員認証・メール送信なし':'検証用の新受付 ／ 架空データのみ。当日登録・札割当・入場の保存可。一般メール停止・本番未切替'):'未公開の隔離候補 ／ 保存先は未設定です';
 const mount=()=>document.body.prepend(banner);
 if(document.body)mount();else document.addEventListener('DOMContentLoaded',mount,{once:true});
 const namespace='bf-event-candidate::'+encodeURIComponent([target,config.eventId||'unset',config.backendGeneration||'unset'].join('|'))+'::';
 const storage={getItem:key=>localStorage.getItem(namespace+key),setItem:(key,value)=>localStorage.setItem(namespace+key,value),removeItem:key=>localStorage.removeItem(namespace+key)};
 const ready=Promise.all([import('./client.mjs'),import('./portal-auth.mjs')]).then(async([{createClient},{createPortalStaffProvider}])=>{
  if(!staffProvider&&config.authSource==='portal'&&endpoint&&config.portalUrl){
   staffProvider=createPortalStaffProvider({storage:localStorage,origin:location.origin,portalUrl:config.portalUrl,endpoint:target,fetchImpl:nativeFetch,navigate:url=>location.assign(url)});
  }
  let session=null;
  try{session=await staffProvider?.getSession?.();}catch{banner.textContent='ポータルのログイン確認に失敗しました。通信状態を確認して再読込してください。';}
  const key='bf_event_candidate_session';
  try{
   if(session?.user_id&&session?.access_token&&session.expires_at>Date.now())storage.setItem(key,JSON.stringify({user_id:session.user_id,name:session.name||session.user_id,expires:session.expires_at,token:'server-auth-required'}));
   else storage.removeItem(key);
  }catch{/* 保存不可は元画面のガードで表示する。 */}
  const call=createClient({endpoint:endpoint?target:'',fetchImpl:nativeFetch,getStaffToken:async()=>{
   const current=await staffProvider?.getSession?.();return current?.expires_at>Date.now()?current.access_token:null;
  }});
  window.BFEventCandidate.call=call;
  window.fetch=async(url,options={})=>{
   const actual=new URL(typeof url==='string'?url:url.url,location.href);
   if(actual.href.split('?')[0]!==target.split('?')[0])return new Response(JSON.stringify({success:false,error:'CANDIDATE_EXTERNAL_ROUTE_BLOCKED'}),{status:503});
   let input;
   if((options.method||'GET').toUpperCase()==='GET')input={action:actual.searchParams.get('action'),data:actual.searchParams.get('data')||'{}'};
   else if(options.body instanceof URLSearchParams)input=Object.fromEntries(options.body);
   else input=JSON.parse(options.body||'{}');
   const data=typeof input.data==='string'?JSON.parse(input.data):input.data;
   return call(input.action,data,options);
  };
 });
 // 検証LIFFには別の登録先が必要。既存本番IDへ黙って戻さない。
 const lineLiffId=typeof config.lineLiffId==='string'&&/^2010404613-[A-Za-z0-9]+$/.test(config.lineLiffId)?config.lineLiffId:'';
 window.BFEventCandidate={endpoint:target,ready,storage,lineEnabled:config.enableLine===true&&!!lineLiffId,lineLiffId,signIn:async()=>{
  if(typeof staffProvider?.signIn!=='function')throw Error('社員認証はまだ設定されていません。');
  return staffProvider.signIn();
 }};
})();
