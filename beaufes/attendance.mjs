import {EVENT_DATE,snapshot,selectBase,totals,histogram,selectList,minuteLabel,sessionClock} from './attendance-core.mjs?b=20261025b';
// 日本語コメント: バージョンの定義はここだけ。既存の社員認証・読取APIを共用し、保存APIは呼ばない。
const APP_VERSION='1.0.0';
const $=id=>document.getElementById(id);
let data=null,status='arrived',minute=null,inFlight=false,timer=null,lastSuccess=0,denied=false;
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
  button.setAttribute('aria-label',`${date} ${range} ${bin.count}人。押すと来場者を表示`);button.setAttribute('aria-pressed',String(bin.minute===minute));button.title=range+'：'+bin.count+'人';
  const fill=textNode('span','','barFill'+(bin.count?'':' zero'));fill.style.height=(bin.count/chart.max*145)+'px';
  if(bin.count)fill.append(textNode('span',String(bin.count),'barNumber'));
  button.append(fill);
  if(interval===60||index%3===0)button.append(textNode('span',minuteLabel(bin.minute),'barLabel'));
  button.addEventListener('click',()=>{minute=minute===bin.minute?null:bin.minute;status='arrived';render();$('listPanel').scrollIntoView({behavior:'smooth',block:'start'});});root.append(button);
 }
 $('chartSummary').textContent=`${date.replaceAll('-','/')} の来場 ${chart.count}人 · 棒を押すとその時間帯の来場者を表示`;
 $('chartCaption').textContent=`${interval}分ごと・最初の受付時刻で集計（日本時間）`;
 const notices=[];
 if(date!==EVENT_DATE)notices.push('開催日（10/26）以外の受付を表示しています。');
 if(chart.otherDate)notices.push(`別日の受付 ${chart.otherDate}人はグラフの対象外です。上の受付済み人数には含みます。`);
 if(chart.unknown)notices.push(`受付時刻を確認できない ${chart.unknown}人はグラフの対象外です。上の受付済み人数と一覧には含みます。`);
 if(!chart.count)notices.push('この日の受付記録はまだありません。');
 $('chartNotice').textContent=notices.join(' ');$('clearTime').hidden=minute===null;
}
function renderList(base){
 const interval=Number($('interval').value),date=$('chartDate').value;
 const list=selectList(base,{status,date,minute,interval,sort:$('sort').value});
 const title=status==='arrived'?'来場者一覧（受付済み）':status==='waiting'?'まだ来ていない人（未受付）':'申込者一覧（全員）';
 $('listTitle').textContent=title;
 $('listSummary').textContent=`表示 ${list.length}人`+(minute!==null?' · '+date.replaceAll('-','/')+' '+minuteLabel(minute)+'〜'+minuteLabel(minute+interval):'');
 const root=$('rows');root.replaceChildren();$('empty').hidden=list.length!==0;
 for(const a of list){
  const row=document.createElement('tr'),check=document.createElement('td');
  check.append(textNode('span',a.arrived?'受付済み':'未受付','status'+(a.arrived?'':' wait')));
  if(a.arrived)check.append(textNode('span',a.time?a.time.date.slice(5).replace('-','/')+' '+a.time.label:'時刻確認中','receiptTime'));
  row.append(check,textNode('td',a.salon_name||'—'),textNode('td',a.staff_name||'—'),textNode('td',a.tantou||'未割当'));
  const booking=document.createElement('td');
  if(a.reserved.length)for(const s of a.reserved)booking.append(textNode('span',sessionLabel(s),'reservation'));
  else booking.textContent='予約なし';
  row.append(booking,textNode('td',a.app_id));
  [...row.children].forEach((cell,index)=>{cell.dataset.label=['受付状態・時刻','サロン名','お名前','担当','予約内容','受付番号'][index];});root.append(row);
 }
}
function setStatus(value){status=value;minute=null;render();$('listPanel').scrollIntoView({behavior:'smooth',block:'start'});}
function showError(message){$('error').textContent=message;$('error').hidden=false;$('liveDot').className='liveDot stale';$('liveLabel').textContent=lastSuccess?'更新できていません（前回の情報）':'取得できていません';}
function requireLogin(){
 denied=true;data=null;$('rows').replaceChildren();$('dashboard').hidden=true;$('gate').hidden=false;$('error').hidden=true;
 $('liveDot').className='liveDot stale';$('liveLabel').textContent='社員ログインを確認してください';$('lastUpdated').textContent='';
}
function schedule(){clearTimeout(timer);if(!denied&&$('autoRefresh').checked&&!document.hidden)timer=setTimeout(()=>refresh(),30000);}
async function refresh(){
 if(inFlight)return;
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
 }catch(error){
  if(['SESSION_INVALID','SESSION_EXPIRED','FORBIDDEN','AUTH_REQUIRED'].includes(error?.message)){requireLogin();return;}
  showError('最新の状況を取得できませんでした。'+(lastSuccess?'表示は '+fullTime(lastSuccess)+' 時点です。':'')+'通信状態を確認して「最新の状況を取得」を押してください。');
 }finally{clearTimeout(timeout);inFlight=false;$('refresh').disabled=false;schedule();}
}
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
