// 日本語コメント: 既存画面の要求に新しい版数・操作ID・社員Bearerを付ける。旧認証へフォールバックしない。
export const VERSION='0.3.2';
const mutations=new Set(['apply','applyLiff','updateApplication','setTantou','adminCreateApplication','assignSpare','releaseSpare','resendPass','adminResendPass']);
const staffActions=new Set(['staffSession','adminBootstrap','listApplications','applicationCustomers','listBadges','listSpareBadges','listReservations','staffGetReservations','staffChangeReservation','setTantou','adminCreateApplication','assignSpare','releaseSpare','adminResendPass','scanRoster','scanCheckedIn','scanCheckin']);
export function createClient({endpoint,fetchImpl=fetch,getStaffToken,uuid=()=>crypto.randomUUID()}={}){
 const applicationVersions=new Map(),lineProfiles=new Map(),pending=new Map();
 return async(action,data={},options={})=>{
  if(!endpoint)return new Response(JSON.stringify({success:false,error:'BACKEND_NOT_CONFIGURED'}),{status:503,headers:{'Content-Type':'application/json'}});
  if(!data||Array.isArray(data)||typeof data!=='object')throw Error('INVALID_REQUEST');
  // 同じ本文の通信再試行では操作IDと読取版数を維持し、応答喪失後の二重保存を防ぐ。
  const source={...data};delete source.session_token;
  const key=action+':'+JSON.stringify(source);
  let body=pending.get(key);
  if(!body){
   body={...source};
   if(action==='updateApplication'&&!Object.hasOwn(body,'application_revision')){
    const revision=applicationVersions.get(body.ticket_token);
    if(!revision)throw Error('APPLICATION_REFRESH_REQUIRED');body.application_revision=revision;
   }
   if(action==='applyLiff'){
    const profile=lineProfiles.get(body.id_token);if(!profile)throw Error('LINE_PROFILE_CHANGED');
    if(!body.profile_revision)body.profile_revision=profile.profile_revision;
    if(profile.existing_application)body.application_revision=profile.existing_application.application_revision;
   }
   if(mutations.has(action)&&!body.request_id)body.request_id=uuid();
   if(mutations.has(action)){
    if(pending.size>=256)throw Error('PENDING_OPERATIONS_REQUIRE_REVIEW');pending.set(key,body);
   }
  }
  const headers={'Content-Type':'application/json'};
  if(staffActions.has(action)){
   if(typeof getStaffToken!=='function')throw Error('AUTH_NOT_CONFIGURED');
   const token=await getStaffToken();if(!token)throw Error('SESSION_INVALID');headers.Authorization='Bearer '+token;
  }
  const response=await fetchImpl(endpoint,{method:'POST',headers,body:JSON.stringify({action,data:body}),signal:options.signal,cache:'no-store',redirect:'error'});
  let json;try{json=await response.clone().json();}catch{return response;}
  if(json?.success){
   pending.delete(key);
   if(action==='getPass'&&Number.isSafeInteger(json.data?.application_revision))applicationVersions.set(body.ticket_token,json.data.application_revision);
   if(action==='updateApplication'&&Number.isSafeInteger(json.data?.revision))applicationVersions.set(body.ticket_token,json.data.revision);
   if(action==='liffPrefill')lineProfiles.set(body.id_token,json.data);
   if(action==='applyLiff')lineProfiles.delete(body.id_token);
  }else if(response.status<500&&json?.error&&json.error!=='AUTH_TRANSIENT'&&json.error!=='BACKEND_FAILED'&&json.error!=='BACKEND_NOT_CONFIGURED'){
   pending.delete(key);
   if(['APPLICATION_CONFLICT','APPLICATION_REFRESH_REQUIRED'].includes(json.error))applicationVersions.delete(body.ticket_token);
  }
  // 顧客の操作を変えず、上限時だけ固定の日本語案内にする。APIコードは別項目へ保持する。
  if(json?.success===false&&json.error==='PUBLIC_RATE_LIMIT')return new Response(JSON.stringify({...json,code:json.error,error:'お申し込み・入場パス再送の回数が上限に達しました。10分ほど時間をおいて、もう一度お試しください。'}),{status:200,headers:{'Content-Type':'application/json'}});
  // 既存GAS画面は業務エラーをHTTP200内のsuccess:falseで扱う。通信障害はそのまま残す。
  if(!json.success&&response.status>=400&&response.status<500)return new Response(JSON.stringify(json),{status:200,headers:{'Content-Type':'application/json'}});
  return response;
 };
}
