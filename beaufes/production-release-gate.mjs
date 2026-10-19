// 日本語コメント: 公開完了とサーバー切替を分ける待機。権限・保存可否の検証は毎要求サーバー側で行う。
export const VERSION='0.1.0';
const ENDPOINT='https://ovblkjxlbmnehgkpmyrd.supabase.co/functions/v1/beaufes-event';
export async function waitForProductionRelease({endpoint,fetchImpl=fetch,waitImpl=ms=>new Promise(resolve=>setTimeout(resolve,ms)),onState=()=>{},maxAttempts=48,intervalMs=15000}={}){
 if(endpoint!==ENDPOINT||!Number.isInteger(maxAttempts)||maxAttempts<1||maxAttempts>48||!Number.isInteger(intervalMs)||intervalMs<0||intervalMs>30000)throw Error('EXACT_RELEASE_GATE_REQUIRED');
 onState('waiting');
 for(let attempt=0;attempt<maxAttempts;attempt++){
  const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),5000);
  try{
   // healthは公開状態の固定項目だけ。社員/LINEトークン、QR、申込情報は送らない。
   const r=await fetchImpl(ENDPOINT+'?health=1',{method:'GET',redirect:'error',credentials:'omit',cache:'no-store',headers:{Accept:'application/json'},signal:controller.signal});
   const b=r.status===200?await r.json():null;
   if(b?.ok===true&&b.event_id==='beaufes2026'&&b.namespace==='bf_event_'&&b.release_enabled===true&&b.configuration_ready===true&&b.access_mode==='full'&&b.public_source_mode==='shared-event'){
    onState('ready');return true;
   }
  }catch{/* 通信や未設定を、旧GASへの保存再試行へ変換しない。 */}
  finally{clearTimeout(timer);}
  if(attempt+1<maxAttempts)await waitImpl(intervalMs);
 }
 onState('paused');return false;
}
