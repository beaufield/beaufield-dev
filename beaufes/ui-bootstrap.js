// 日本語コメント: 隔離画面専用。未設定時は外部へ保存せず停止する。本番の名簿・未送信・セッションを読まない。
(function(){
 'use strict';
 const config=window.BF_EVENT_CANDIDATE_CONFIG||{};
 const endpoint=typeof config.endpoint==='string'&&config.environment==='production-candidate'&&config.releaseEnabled===true&&config.eventId==='beaufes2026'&&config.endpoint==='https://ovblkjxlbmnehgkpmyrd.supabase.co/functions/v1/beaufes-event'&&typeof config.eventId==='string'&&config.eventId.trim()&&typeof config.backendGeneration==='string'&&config.backendGeneration.trim()?config.endpoint:'';
 const target=new URL(endpoint||'./__backend_not_configured__',location.href).href;
 const nativeFetch=window.fetch.bind(window);
 let staffProvider=window.BF_EVENT_STAFF_PROVIDER;
 const banner=document.createElement('div');
 banner.style.cssText='position:sticky;top:0;z-index:20000;padding:9px 15px;background:#fff0c9;color:#684213;font:13px/1.6 sans-serif;text-align:center';
 banner.textContent=endpoint?(config.fixture===true?'架空データの画面操作試験 ／ 実社員認証・メール送信なし':'ビューフェス2026'):'ただいまシステムを切り替えています';
 const mount=()=>document.body.prepend(banner);
 if(document.body)mount();else document.addEventListener('DOMContentLoaded',mount,{once:true});
 const namespace='bf-event-production::'+encodeURIComponent([target,config.eventId||'unset',config.backendGeneration||'unset'].join('|'))+'::';
 const storage={getItem:key=>localStorage.getItem(namespace+key),setItem:(key,value)=>localStorage.setItem(namespace+key,value),removeItem:key=>localStorage.removeItem(namespace+key)};
// 日本語コメント: 公開済みの新画面を、サーバー側の最終開放まで保守表示で待機させる。
let releaseLayer=null,releaseOpen=false;
const showReleaseState=state=>{
 if(state==='ready'){releaseOpen=true;releaseLayer?.remove();banner.textContent='ビューフェス2026';return;}
 if(!releaseLayer){
  releaseLayer=document.createElement('div');releaseLayer.setAttribute('role','status');releaseLayer.setAttribute('aria-live','polite');
  releaseLayer.style.cssText='position:fixed;inset:0;z-index:30000;background:#f4f8f8;color:#173b49;display:flex;align-items:center;justify-content:center;padding:28px;font:16px/1.9 sans-serif;text-align:center';
  const box=document.createElement('div'),heading=document.createElement('h1'),message=document.createElement('p'),retry=document.createElement('button');
  heading.style.cssText='font-size:26px;line-height:1.5';heading.textContent='ただいまシステムを切り替えています';
  message.textContent='しばらくお待ちください。この画面は自動で確認します。';message.dataset.releaseMessage='true';
  retry.type='button';retry.textContent='もう一度確認する';retry.style.cssText='margin-top:18px;padding:12px 22px;font:inherit;border:0;border-radius:8px;background:#087b78;color:white';
  retry.addEventListener('click',()=>location.reload());retry.hidden=true;retry.dataset.releaseRetry='true';
  box.append(heading,message,retry);releaseLayer.append(box);
  const mountRelease=()=>{if(!releaseOpen)document.body.append(releaseLayer);};
  if(document.body)mountRelease();else document.addEventListener('DOMContentLoaded',mountRelease,{once:true});
 }
 if(state==='paused'){
  releaseLayer.querySelector('[data-release-message]').textContent='まだ完了を確認できませんでした。時間をおいて、もう一度確認してください。';
  releaseLayer.querySelector('[data-release-retry]').hidden=false;
 }
};

 const releaseGate=endpoint?import('./production-release-gate.mjs?b=383c6f228909').then(({waitForProductionRelease})=>waitForProductionRelease({endpoint:target,fetchImpl:nativeFetch,onState:showReleaseState})):Promise.resolve(true);
 const ready=releaseGate.then(open=>{if(!open)throw Error('切替完了を確認できません。もう一度開いてください。');return Promise.all([import('./client.mjs'),import('./portal-auth.mjs')]);}).then(async([{createClient},{createPortalStaffProvider}])=>{
  let portalProvider=false;
  if(!staffProvider&&config.authSource==='portal'&&endpoint&&config.portalUrl){
   staffProvider=createPortalStaffProvider({storage:localStorage,origin:location.origin,portalUrl:config.portalUrl,endpoint:target,fetchImpl:nativeFetch,navigate:url=>location.assign(url)});
   portalProvider=true;
  }
  // 日本語コメント: ポータル確認の一時障害では、期限内の端末ログインを消さない。権限は毎要求サーバーが確認する。
  const portalSession=()=>{
   if(!portalProvider||location.origin!==new URL(config.portalUrl).origin)return null;
   let saved;try{saved=JSON.parse(localStorage.getItem('bf_session')||'null');}catch{return null;}
   if(!saved||typeof saved.token!=='string'||!/^[A-Za-z0-9_-]{32,256}$/.test(saved.token)||!Number.isSafeInteger(saved.expires)||saved.expires<=Date.now()||typeof saved.user_id!=='string'||!saved.user_id)return null;
   return {user_id:saved.user_id,name:typeof saved.name==='string'&&saved.name?saved.name:saved.user_id,access_token:saved.token,expires_at:saved.expires};
  };
  const key='bf_event_candidate_session';
  // 受付画面はこの写しで期限を表示する。別タブで再ログインした後も、再読込なしで新しい期限へ更新する。
  const remember=current=>{try{storage.setItem(key,JSON.stringify({user_id:current.user_id,name:current.name||current.user_id,expires:current.expires_at,token:'server-auth-required'}));}catch{/* 保存不可は元画面のガードで表示する。 */}};
  const normalBanner=banner.textContent;
  let session=null;
  try{session=await staffProvider?.getSession?.();}
  catch{
   session=portalSession();
   banner.textContent=session?'ログインの確認が一時的にできません。受付は続けられます（記録は端末に残り、つながり次第送ります）。':'ポータルのログイン確認に失敗しました。通信状態を確認して再読込してください。';
  }
  try{
   if(session?.user_id&&session?.access_token&&session.expires_at>Date.now())remember(session);
   else storage.removeItem(key);
  }catch{/* 保存不可は元画面のガードで表示する。 */}
  const call=createClient({endpoint:endpoint?target:'',fetchImpl:nativeFetch,getStaffToken:async()=>{
   let current;
   try{current=await staffProvider?.getSession?.();}
   catch(error){
    // 確認だけが一時的に失敗した場合も、期限内の元セッションで送り、可否はサーバーの判定に任せる。
    const saved=portalSession();if(!saved)throw error;return saved.access_token;
   }
   if(!(current?.expires_at>Date.now()))return null;
   remember(current);if(banner.textContent!==normalBanner)banner.textContent=normalBanner;
   return current.access_token;
  }});
  window.BFEventCandidate.call=call;
  window.fetch=async(url,options={})=>{
   const actual=new URL(typeof url==='string'?url:url.url,location.href);
   // LINE公式SDK自身のログイン通信だけは元のfetchへ渡す。GASやメッセージ送信APIは許可しない。
   const lineApiPaths=['/oauth2/v2.1/token','/oauth2/v2.1/certs','/oauth2/v2.1/verify','/v2/profile'];
   const lineAccessPaths=['/.well-known/openid-configuration','/liff/v1/authorize'];
   const lineAppPath='/liff/v2/apps/'+lineLiffId;
   const lineSdkRequest=lineLiffId&&(actual.origin==='https://api.line.me'&&(lineApiPaths.includes(actual.pathname)||actual.pathname===lineAppPath||actual.pathname.startsWith(lineAppPath+'/'))||actual.origin==='https://access.line.me'&&lineAccessPaths.includes(actual.pathname));
   if(lineSdkRequest)return nativeFetch(url,options);
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
