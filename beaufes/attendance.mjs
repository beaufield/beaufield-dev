import {EVENT_DATE,snapshot,selectBase,totals,histogram,selectList,minuteLabel,sessionClock} from './attendance-core.mjs?b=20261027b';
import {UNKNOWN_TIME,pendingCheckin,prepareCheckin,clearCheckin,validateCheckinResult} from './attendance-checkin.mjs?b=20261027a';
// 日本語コメント: 既存の社員認証と受付APIを共用する。申込・予約・パス再送APIは呼ばない。
const APP_VERSION='1.2.0';
const $=id=>document.getElementById(id);
let data=null,status='arrived',minute=null,inFlight=false,timer=null,lastSuccess=0,denied=false;
let manualAppId=null,manualBusy=false;
const fullTime=epoch=>new Intl.DateTimeFormat('ja-JP',{timeZone:'Asia/Tokyo',month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false}).format(epoch);
function textNode(tag,text,className){const n=document.createElement(tag);n.textContent=text;if(className)n.className=className;return n;}
function criteria(){return {query:$('query').value,tantou:$('tantou').value,booking:$('booking').value,session:$('session').value};}
function setOptions(id,pairs){
 const select=$(id),old=select.value;select.replaceChildren();
 for(const [value,label] of pairs){const o=document.createElement('option');o.value=value;o.textContent=label;select.append(o);}
 if(pairs.some(p=>p[0]===old))select.value=old;
}
function sessionLabel(s){
 const start=sessionClock(s.starts_at),end=sessionClock(s.ends_at);
 const time=start?start+(end?'〜'+end:''):'';
 return [time,s.title||s.slot||s.session_id,s.title&&s.slot&&s.slot!==s.title?s.slot:''].filter(Boolean).join(' ');
}
function syncOptions(){
 const names=[...new Set([...data.tantou,...data.rows.map(a=>a.tantou)].filter(Boolean))];
 setOptions('tantou',[['','すべての担当'],...names.map(t=>[t,t]),['__empty__','担当未割当']]);
 const used=new Set(data.rows.flatMap(a=>a.reserved.map(s=>s.session_id)));
 setOptions('session',[['','すべての予約枠'],...data.sessions.filter(s=>used.has(s.session_id)).sort((a,b)=>String(a.starts_at).localeCompare(String(b.starts_at))||a.session_id.localeCompare(b.session_id)).map(s=>[s.session_id,sessionLabel(s)])]);
}
function render(){
 if(!data)return;
 const filter=criteria(),base=selectBase(data.rows,filter),counts=totals(base);
 $('arrivedCount').textContent=counts.arrived;$('waitingCount').textContent=counts.waiting;$('totalCount').textContent=counts.total;
 const percent=counts.total?Math.round(counts.arrived/counts.total*100):0;
 $('receptionProgress').value=percent;$('progressText').textContent='受付率 '+percent+'%';
 const filtered=filter.query.trim()||filter.tantou||filter.booking!=='all'||filter.session;
 $('scopeLabel').textContent=filtered?'絞り込み中 '+counts.total+'人 ／ 全'+data.rows.length+'人':'全申込者 '+data.rows.length+'人';
 document.querySelectorAll('[data-status]').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.status===status)));
 renderChart(base);renderList(base);
}
function renderChart(base){
 const date=$('chartDate').value,interval=Number($('interval').value);
 if(!/^\d{4}-\d{2}-\d{2}$/.test(date))return;
 const chart=histogram(base,date,interval),root=$('chart');root.replaceChildren();
 root.style.minWidth=chart.bins.length>12?Math.max(820,chart.bins.length*24)+'px':'100%';
 for(const [index,bin] of chart.bins.entries()){
  const range=minuteLabel(bin.minute)+'〜'+minuteLabel(bin.minute+interval);
  const button=textNode('button','','bar'+(bin.minute===minute?' active':''));button.type='button';
  button.setAttribute('aria-label',`${date} ${range} ${bin.count}人・この時間帯の終了までの累計${bin.cumulative}人。押すと来場者を表示`);button.setAttribute('aria-pressed',String(bin.minute===minute));button.title=range+'：'+bin.count+'人 ／ 累計'+bin.cumulative+'人';
  const fill=textNode('span','','barFill'+(bin.count?'':' zero'));fill.style.height=(bin.count/chart.max*145)+'px';
  if(bin.count)fill.append(textNode('span',String(bin.count),'barNumber'));
  button.append(fill);
  if(interval===60||index%3===0)button.append(textNode('span',minuteLabel(bin.minute),'barLabel'));
  button.addEventListener('click',()=>{minute=minute===bin.minute?null:bin.minute;status='arrived';render();$('listPanel').scrollIntoView({behavior:'smooth',block:'start'});});root.append(button);
 }
 renderCumulative(chart);
 $('chartSummary').textContent=`${date.replaceAll('-','/')} の来場・累計 ${chart.count}人 · 棒を押すとその時間帯の来場者を表示`;
 $('chartCaption').textContent=`${interval}分ごと・最初の受付時刻で集計（日本時間）。線は各時間帯の終了までの累計`;
 const notices=[];
 if(date!==EVENT_DATE)notices.push('開催日（10/26）以外の受付を表示しています。');
 if(chart.otherDate)notices.push(`別日の受付 ${chart.otherDate}人はグラフの対象外です。上の受付済み人数には含みます。`);
 if(chart.unknown)notices.push(`受付時刻を確認できない ${chart.unknown}人はグラフの対象外です。上の受付済み人数と一覧には含みます。`);
 if(!chart.count)notices.push('この日の受付記録はまだありません。');
 $('chartNotice').textContent=notices.join(' ');$('clearTime').hidden=minute===null;
}
function renderCumulative(chart){
 const max=Math.max(1,chart.count);
 function axis(id,limit){
  const root=$(id);root.replaceChildren(textNode('span','人','axisTitle'));
  for(const value of [...new Set([0,Math.ceil(limit/2),limit])]){
   const label=textNode('span',String(value),'axisTick');
   label.style.top=(65+(1-value/limit)*145)+'px';root.append(label);
  }
 }
 // 日本語コメント: 棒の高さ・クリックは維持。累計は右軸の独立目盛で重ね、左右の尺度を明記する。
 axis('barAxis',chart.max);axis('cumulativeAxis',max);
 const ns='http://www.w3.org/2000/svg',svg=document.createElementNS(ns,'svg');
 svg.classList.add('cumulativeOverlay');svg.setAttribute('viewBox','0 0 1000 145');svg.setAttribute('preserveAspectRatio','none');svg.setAttribute('aria-hidden','true');
 const points=chart.bins.map((bin,index)=>({x:(index+.5)/chart.bins.length*1000,y:145-bin.cumulative/max*145}));
 const line=document.createElementNS(ns,'polyline');line.setAttribute('points',points.map(p=>p.x+','+p.y).join(' '));line.classList.add('cumulativeLine');svg.append(line);
 for(const [index,point] of points.entries()){
  const dot=document.createElementNS(ns,'circle');dot.setAttribute('cx',point.x);dot.setAttribute('cy',point.y);dot.setAttribute('r','3');dot.dataset.cumulative=chart.bins[index].cumulative;dot.classList.add('cumulativeDot');svg.append(dot);
 }
 $('chart').append(svg);
}
function renderList(base){
 const interval=Number($('interval').value),date=$('chartDate').value;
 const list=selectList(base,{status,date,minute,interval,sort:$('sort').value});
 const title=status==='arrived'?'来場者一覧（受付済み）':status==='waiting'?'未受付の方の一覧':'申込者一覧（全員）';
 $('listTitle').textContent=title;
 $('listSummary').textContent=`表示 ${list.length}人`+(minute!==null?' · '+date.replaceAll('-','/')+' '+minuteLabel(minute)+'〜'+minuteLabel(minute+interval):'');
 const root=$('rows');root.replaceChildren();$('empty').hidden=list.length!==0;
 for(const a of list){
  const row=document.createElement('tr'),check=document.createElement('td');
  check.append(textNode('span',a.arrived?'受付済み':'未受付','status'+(a.arrived?'':' wait')));
  if(a.arrived)check.append(textNode('span',a.time?a.time.date.slice(5).replace('-','/')+' '+a.time.label:a.checked_in_at===UNKNOWN_TIME?'来場時刻不明（手動）':'時刻確認中','receiptTime'));
  row.append(check,textNode('td',a.salon_name||'—'),textNode('td',a.staff_name||'—'),textNode('td',a.tantou||'未割当'));
  const booking=document.createElement('td');
  if(a.reserved.length)for(const s of a.reserved)booking.append(textNode('span',sessionLabel(s),'reservation'));
  else booking.textContent='予約なし';
  row.append(booking,textNode('td',a.app_id));
  const action=document.createElement('td');
  if(!a.arrived){const button=textNode('button','受付済みにする','manualButton');button.type='button';button.disabled=manualBusy;button.addEventListener('click',()=>openManual(a.app_id));action.append(button);}
  row.append(action);
  [...row.children].forEach((cell,index)=>{cell.dataset.label=['受付状態・時刻','サロン名','お名前','担当','予約内容','受付番号','手動受付'][index];});root.append(row);
 }
}
function setStatus(value){status=value;minute=null;render();$('listPanel').scrollIntoView({behavior:'smooth',block:'start'});}
function showError(message){$('error').textContent=message;$('error').hidden=false;$('liveDot').className='liveDot stale';$('liveLabel').textContent=lastSuccess?'更新できていません（前回の情報）':'取得できていません';}
function requireLogin(){
 if($('manualDialog').open)$('manualDialog').close();manualAppId=null;
 denied=true;data=null;$('rows').replaceChildren();$('dashboard').hidden=true;$('gate').hidden=false;$('error').hidden=true;
 $('liveDot').className='liveDot stale';$('liveLabel').textContent='社員ログインを確認してください';$('lastUpdated').textContent='';
}
function schedule(){clearTimeout(timer);if(!denied&&!manualBusy&&$('autoRefresh').checked&&!document.hidden)timer=setTimeout(()=>refresh(),30000);}
async function refresh(){
 if(inFlight)return false;
 inFlight=true;clearTimeout(timer);$('refresh').disabled=true;
 const controller=new AbortController(),timeout=setTimeout(()=>controller.abort(),45000);
 try{
  await window.BFEventCandidate.ready;
  const response=await window.BFEventCandidate.call('adminBootstrap',{include_applications:true},{signal:controller.signal});
  const json=await response.json();
  if(response.status===401||response.status===403||['SESSION_INVALID','SESSION_EXPIRED','FORBIDDEN','AUTH_REQUIRED','TOKEN_EXPIRED'].includes(json?.error)){requireLogin();return;}
  if(!response.ok||json?.success!==true)throw Error(json?.error||'READ_FAILED');
  const next=snapshot(json.data);
  data=next;lastSuccess=Date.now();denied=false;syncOptions();render();
  $('gate').hidden=true;$('dashboard').hidden=false;$('error').hidden=true;$('liveDot').className='liveDot on';
  $('liveLabel').textContent=$('autoRefresh').checked?'自動更新中':'手動更新';
  $('lastUpdated').textContent='最終更新 '+fullTime(lastSuccess)+' · サーバーに届いた受付を表示';
  return true;
 }catch(error){
  if(['SESSION_INVALID','SESSION_EXPIRED','FORBIDDEN','AUTH_REQUIRED'].includes(error?.message)){requireLogin();return;}
  showError('最新の状況を取得できませんでした。'+(lastSuccess?'表示は '+fullTime(lastSuccess)+' 時点です。':'')+'通信状態を確認して「最新の状況を取得」を押してください。');
  return false;
 }finally{clearTimeout(timeout);inFlight=false;$('refresh').disabled=false;schedule();}
}
function manualMessage(message){$('manualNotice').textContent=message;$('manualNotice').hidden=false;}
function openManual(appId){
 if(manualBusy||denied)return;
 const a=data?.rows.find(a=>a.app_id===appId);if(!a||a.arrived)return;
 manualAppId=appId;$('manualError').hidden=true;$('manualNotice').hidden=true;
 $('manualPerson').textContent=[a.salon_name,a.staff_name,appId].filter(Boolean).join('\n');
 $('manualTime').value='';$('manualTime').disabled=false;$('manualSave').disabled=false;$('manualSave').textContent='受付済みにする（通知なし）';
 try{
  const pending=pendingCheckin(window.BFEventCandidate.storage,appId);
  if(pending){$('manualTime').value=pending.checked_in_at===UNKNOWN_TIME?'':pending.checked_in_at.slice(0,16).replace(' ','T');$('manualTime').disabled=true;$('manualSave').textContent='結果を確認して再試行（通知なし）';}
 }catch{$('manualError').textContent='前回の記録を確認できないため保存を停止しました。画面を再読み込みして確認してください。';$('manualError').hidden=false;$('manualSave').disabled=true;}
 $('manualDialog').showModal();
}
async function saveManual(event){
 event.preventDefault();if(manualBusy||!manualAppId)return;
 const appId=manualAppId;let acknowledged=false;
 manualBusy=true;clearTimeout(timer);$('manualSave').disabled=true;$('manualCancel').disabled=true;$('manualTime').disabled=true;$('manualError').hidden=true;
 function finish(message){if($('manualDialog').open)$('manualDialog').close();manualAppId=null;manualMessage(message);}
 try{
  // 最新の取消・他端末受付を確認できない間は新しい記録を送らない。
  if(!await refresh())throw Error('READ_REQUIRED');
  const a=data.rows.find(a=>a.app_id===appId);
  if(!a)throw Error('APPLICATION_CHANGED');
  if(a.arrived){try{clearCheckin(window.BFEventCandidate.storage,appId);}catch{}finish(a.staff_name+'さんはすでに受付済みです。人数は重複して増やしません。');return;}
  const row=prepareCheckin(window.BFEventCandidate.storage,appId,$('manualTime').value);
  const response=await window.BFEventCandidate.call('scanCheckin',{checkins:[row]},{signal:AbortSignal.timeout(45000)});
  const json=await response.json();
  if(['SESSION_INVALID','SESSION_EXPIRED','FORBIDDEN','AUTH_REQUIRED','TOKEN_EXPIRED'].includes(json?.error)){requireLogin();manualMessage('受付の保存を確認できませんでした。社員ログイン後、同じ方の状態を確認してください。');return;}
  if(!response.ok||!validateCheckinResult(json,row))throw Error('SAVE_UNCONFIRMED');
  acknowledged=true;
  try{clearCheckin(window.BFEventCandidate.storage,appId);}catch{}
  const read=await refresh();
  if(read&&data?.rows.some(a=>a.app_id===appId&&a.arrived))finish(a.staff_name+'さんを受付済みにしました。本人への通知は送信していません。');
  else finish('受付記録は保存されました。人数の更新を確認できないため「最新の状況を取得」を押してください。本人への通知は送信していません。');
 }catch(error){
  // 応答が失われた場合も、読取で受付済みと確認できれば再送しない。
  if(!acknowledged&&!['READ_REQUIRED','APPLICATION_CHANGED','INVALID_ARRIVAL_TIME','PENDING_INVALID','PENDING_STORAGE_FAILED'].includes(error?.message)){
   const read=await refresh();
   if(read&&data?.rows.some(a=>a.app_id===appId&&a.arrived)){try{clearCheckin(window.BFEventCandidate.storage,appId);}catch{}finish('受付済みへの反映を確認しました。本人への通知は送信していません。');return;}
  }
  if(denied)return;
  const messages={READ_REQUIRED:'最新の状態を確認できませんでした。通信状態を確認してから再試行してください。',APPLICATION_CHANGED:'この方の申込状態が変わりました。最新の一覧を確認してください。',INVALID_ARRIVAL_TIME:'来場日時は開催日（2026/10/26）の正しい日時を入力するか、空欄にしてください。',PENDING_INVALID:'前回の記録を確認できないため保存を停止しました。',PENDING_STORAGE_FAILED:'再試行用の記録を端末に保存できないため受付を送信していません。'};
  $('manualError').textContent=messages[error?.message]||'受付の保存を確認できませんでした。最新の一覧と通信状態を確認してください。再試行は同じ受付記録を使い、人数の重複を防ぎます。';$('manualError').hidden=false;
 }finally{
  manualBusy=false;$('manualSave').disabled=false;$('manualCancel').disabled=false;
  let pending=true;try{pending=!!pendingCheckin(window.BFEventCandidate.storage,appId);}catch{$('manualSave').disabled=true;}
  $('manualTime').disabled=pending;
  $('manualSave').textContent=pending?'結果を確認して再試行（通知なし）':'受付済みにする（通知なし）';
  render();schedule();
 }
}
$('manualForm').addEventListener('submit',saveManual);
$('manualCancel').addEventListener('click',()=>{if(!manualBusy){$('manualDialog').close();manualAppId=null;}});
$('manualDialog').addEventListener('cancel',event=>{if(manualBusy)event.preventDefault();else manualAppId=null;});
$('version').textContent='v'+APP_VERSION;
document.querySelectorAll('[data-status]').forEach(b=>b.addEventListener('click',()=>setStatus(b.dataset.status)));
for(const id of ['query','tantou','booking','session'])$(id).addEventListener(id==='query'?'input':'change',()=>{minute=null;render();});
for(const id of ['chartDate','interval'])$(id).addEventListener('change',()=>{if(!$('chartDate').value)$('chartDate').value=EVENT_DATE;minute=null;render();});
$('sort').addEventListener('change',render);$('clearTime').addEventListener('click',()=>{minute=null;render();});
$('reset').addEventListener('click',()=>{$('query').value='';$('tantou').value='';$('booking').value='all';$('session').value='';minute=null;status='arrived';render();});
$('refresh').addEventListener('click',refresh);$('autoRefresh').addEventListener('change',()=>{if(lastSuccess&&!denied)$('liveLabel').textContent=$('autoRefresh').checked?'自動更新中':'手動更新';schedule();});
document.addEventListener('visibilitychange',()=>{if(document.hidden){clearTimeout(timer);return;}if(!denied&&$('autoRefresh').checked)void refresh();});
window.addEventListener('offline',()=>{showError('オフラインです。'+(lastSuccess?'表示は '+fullTime(lastSuccess)+' 時点です。':'')+'つながると自動更新を再開します。');});
window.addEventListener('online',()=>{if(!denied&&$('autoRefresh').checked)void refresh();});
void refresh();
