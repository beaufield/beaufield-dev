/* 実ログインの通行証で、検証DBだけの読取を同時実行する。JWTは出力しない。 */
(function(){
const button=document.createElement('button');button.textContent='接続・権限・同時利用を検証';document.body.append(button);const output=document.createElement('pre');output.id='diagnostics';document.body.append(output);
button.onclick=async()=>{button.disabled=true;const results=[];const c=window.SERIAL_TEST_CONFIG;
try{
 const s=JSON.parse(localStorage.getItem('bf_session'));if(!s?.token)throw Error('SESSION_INVALID');
 async function issue(app,token=s.token){const r=await fetch(c.issuer,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'serialIssueDbToken',app,session_token:token})});const d=await r.json();if(!d.success)throw Error(app+': '+d.error);return d.token;}
 async function rpc(label,name,args,jwt,schema='public'){const start=performance.now();let r;try{r=await fetch(c.url+'/rest/v1/rpc/'+name,{method:'POST',headers:{apikey:c.publicKey,...(jwt?{Authorization:'Bearer '+jwt}:{}),'Content-Type':'application/json','Content-Profile':schema,'Accept-Profile':schema},body:JSON.stringify(args)});const d=await r.json();const item={label,status:r.status,ms:Math.round(performance.now()-start),success:r.ok&&d.success!==false,rows:Array.isArray(d.rows)?d.rows.length:undefined,error:r.ok?(d.error||undefined):d.message};results.push(item);return item;}catch(e){results.push({label,error:e.message});return null;}}
 const serial=await issue('serial-apps'),order=await issue('order-app'),reservation=await issue('yoyaku-kanri');
 await rpc('anonymous','sr_products',{},null);await rpc('invalid_signature','sr_products',{},'invalid');await rpc('wrong_app','sr_products',{},order);
 const begin=performance.now();await Promise.all([...Array.from({length:11},(_,i)=>rpc('serial-'+i,'sr_page',{},serial)),...Array.from({length:3},(_,i)=>rpc('order-'+i,'get_orders',{p:{}},order)),...Array.from({length:3},(_,i)=>rpc('reservation-'+i,'yk_get_snapshot',{},reservation,'yoyaku_api'))]);
 output.textContent=JSON.stringify({scope:'isolated test project; real portal session; read-only RPCs',elapsed_ms:Math.round(performance.now()-begin),results},null,2);
}catch(e){output.textContent=JSON.stringify({error:e.message,results},null,2);}finally{button.disabled=false;}
};
})();
