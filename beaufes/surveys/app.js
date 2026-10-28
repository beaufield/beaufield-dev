// 日本語コメント: 回答は認証中のメモリー内だけに保持し、両タブを先読みする。
const APP_VERSION='v0.4.0';
const $=id=>document.getElementById(id),el=(tag,text)=>{const n=document.createElement(tag);if(text!==undefined)n.textContent=text;return n;};
const endpoint=document.querySelector('meta[name=survey-api]').content,portal=document.querySelector('meta[name=portal-url]').content;
const names={maker:'メーカー',customer:'お客様'},date=value=>value?new Intl.DateTimeFormat('ja-JP',{timeZone:'Asia/Tokyo',dateStyle:'short',timeStyle:'short'}).format(new Date(value)):'まだ同期されていません';
const freshTabs=()=>Object.fromEntries(Object.keys(names).map(k=>[k,{items:[],state:'idle',error:null,shown:20,received:0}]));
let generation=0,viewGeneration=0,currentToken=null,admin=false,authorized=false,userRoleSeen=false,activeKind='maker',tabs=freshTabs(),summaryData={surveys:[]},summaryError=null,loading=false,filter={},details=new Map(),currentDetail=null,pendingDetail=null,listPosition=0,listFocusId=null;
$('version').textContent=APP_VERSION;$('portal').href=$('signin').href=portal;
const session=()=>{try{const s=JSON.parse(localStorage.getItem('bf_session')||'null');return s&&typeof s.token==='string'&&Number.isSafeInteger(s.expires)&&s.expires>Date.now()?s:null;}catch{return null;}};
// 日本語コメント: 初回HTMLから読込中を表示し、確認前のログイン案内と背面操作を防ぐ。
function busy(show,title='アンケートを読み込んでいます',note='メーカー・お客様の回答を準備しています。',cancel=false){
 $('busy').hidden=!show;$('app-header').inert=$('app-main').inert=show;document.body.classList.toggle('is-busy',show);$('app-main').setAttribute('aria-busy',String(show));
 $('busy-title').textContent=title;$('busy-note').textContent=note;$('busy-cancel').hidden=!cancel;
 if(show)$('busy').focus({preventScroll:true});
}
function feedback(message){$('feedback').textContent=message;$('feedback').hidden=!message;}
function lock(message='社員用ポータルでログインしてください。'){
 generation++;viewGeneration++;currentToken=null;admin=authorized=false;loading=false;tabs=freshTabs();details.clear();currentDetail=pendingDetail=null;summaryData={surveys:[]};summaryError=null;
 for(const id of ['responses','detail','summary','health'])$(id).replaceChildren();
 $('count').textContent='';$('workspace').hidden=true;$('gate').hidden=false;$('detail').hidden=$('admin').hidden=$('load-error').hidden=true;$('list-view').hidden=false;$('status').textContent=message;feedback('');busy(false);
}
function cacheSession(){const s=session();if(!s||s.token!==currentToken){lock('ログイン状態を再確認してください。');return false;}return true;}
const errors={PORTAL_SESSION_INVALID:'ログインの有効期限が切れています。ポータルでログインし直してください。',BEAUFES_ROLE_REQUIRED:'ビューフェスの閲覧権限がありません。管理者へ確認してください。',PORTAL_UNAVAILABLE:'社員認証を確認できません。時間をおいて再試行してください。',RESPONSE_NOT_FOUND:'回答が見つからないか、閲覧が停止されています。',SURVEY_UNAVAILABLE:'回答を取得できません。時間をおいて更新してください。',INVALID_FILTER:'検索条件を確認してください。',INVALID_RESPONSE_ID:'回答リンクを確認してください。'};
const errorText=e=>errors[e.message]||'通信を完了できませんでした。接続を確認して再試行してください。';
function report(e){if(e.message!=='STALE_VIEW'){$('status').textContent=errorText(e);feedback(errorText(e));}}
function applyRole(data){authorized=true;if(data.viewer.role==='user')userRoleSeen=true;admin=!userRoleSeen&&data.viewer.role==='admin';$('admin').hidden=!admin;$('workspace').hidden=false;$('gate').hidden=true;}
async function api(action,body,seq){
 const s=session();if(!s||s.token!==currentToken){lock();throw Error('STALE_VIEW');}const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),45000);
 try{const r=await fetch(endpoint,{method:'POST',headers:{'Content-Type':'application/json',Authorization:'Bearer '+s.token},body:JSON.stringify({action,body}),cache:'no-store',redirect:'error',signal:controller.signal});
  const d=await r.json(),active=session();if(!active||active.token!==s.token){lock();throw Error('STALE_VIEW');}if(seq!==generation)throw Error('STALE_VIEW');
  if([401,403].includes(r.status)||d.error==='PORTAL_UNAVAILABLE'){lock(errors[d.error]||'閲覧権限を確認できません。');throw Error('STALE_VIEW');}if(!r.ok)throw Error(d.error||'SURVEY_UNAVAILABLE');return d;
 }finally{clearTimeout(timer);}
}
function selectKind(kind){if(!names[kind])throw Error('INVALID_RESPONSE_ID');activeKind=kind;for(const key of Object.keys(names)){$('tab-'+key).setAttribute('aria-selected',String(key===kind));$('tab-'+key).tabIndex=key===kind?0:-1;}$('survey-panel').setAttribute('aria-labelledby','tab-'+kind);}
function clearDetailLink(){if(location.hash)history.replaceState(null,'',location.pathname+location.search);}
function hideDetail(){viewGeneration++;pendingDetail=currentDetail=null;$('detail').replaceChildren();$('detail').hidden=true;$('list-view').hidden=false;}
function dateFilter(value,end=false){if(!value)return '';const t=Date.parse(value+'T00:00:00+09:00');if(!Number.isFinite(t))throw Error('INVALID_FILTER');return new Date(t+(end?86400000:0)).toISOString();}
function collect(){const f={q:$('query').value.trim(),score:$('score').value,from:dateFilter($('from').value),until:dateFilter($('until').value,true)};if(f.from&&f.until&&f.from>=f.until)throw Error('INVALID_FILTER');return f;}
// 日本語コメント: 色だけに頼らず、点数・件数・割合も凡例と読み上げで示す。
function ratingChart(s){
 const colors={5:'#21666e',4:'#59a9cc',3:'#e4ba4d',2:'#e18c4c',1:'#bd5770'},counts=Object.fromEntries([5,4,3,2,1].map(n=>[n,Math.max(0,Number(s.distribution[n])||0)])),total=Object.values(counts).reduce((a,b)=>a+b,0);
 const chart=el('div');chart.className='rating-chart';const pie=el('div');pie.className='rating-pie';pie.setAttribute('role','img');const legend=el('ul');legend.className='rating-legend';let start=0;const stops=[],description=[];
 for(let n=5;n>=1;n--){const count=counts[n],percent=total?count/total*100:0,end=start+percent;if(count)stops.push(`${colors[n]} ${start}% ${end}%`);start=end;
  const item=el('li'),swatch=el('span');swatch.className='rating-swatch';swatch.style.background=colors[n];swatch.setAttribute('aria-hidden','true');item.append(swatch,el('span',n+'点'),el('strong',count+'件'),el('span',`${Math.round(percent*10)/10}%`));legend.append(item);description.push(`${n}点 ${count}件`);
 }
 pie.style.background=total?'conic-gradient('+stops.join(',')+')':'#dce5e8';pie.setAttribute('aria-label',total?'満足度の内訳：'+description.join('、'):'満足度の有効回答はありません');chart.append(pie,legend);return chart;
}
function showSummary(){
 $('summary').replaceChildren();for(const s of summaryData.surveys.filter(s=>s.key===activeKind)){const box=el('article');box.className='summary-card';const layout=el('div');layout.className='summary-layout';const stats=el('div');stats.append(el('h2',names[s.key]+'アンケート'),el('strong',s.total+'件'),el('p',s.scored?`平均 ${s.average}／5（有効回答${s.scored}件）`:'満足度の有効回答はありません'));
  const unscored=Math.max(0,s.total-s.scored);if(unscored){const note=el('p',`満足度未回答・該当なし：${unscored}件（円グラフの対象外）`);note.className='meta';stats.append(note);}layout.append(stats,ratingChart(s));box.append(layout);
  const freshness=el('p','最終同期：'+date(s.last_successful_sync_at));freshness.className='meta';if(!s.last_successful_sync_at||Date.now()-Date.parse(s.last_successful_sync_at)>900000)freshness.classList.add('warning');box.append(freshness);$('summary').append(box);
 }
}
function showCards(items){
 $('responses').replaceChildren();for(const r of items.filter(r=>r.survey_key===activeKind)){const card=el('button');card.type='button';card.className='card';card.dataset.id=r.id;
  const label=el('span',names[r.survey_key]);label.className='label';card.append(label,el('span',`満足度 ${r.satisfaction??'未回答・該当なし'}`));card.append(el('h3',[r.company,r.salon,r.person].filter(Boolean).join(' ／ ')||'記名なし'));
  const m=el('p',`${date(r.submitted_at)} ／ ${r.survey_key==='maker'?'M':'C'}-${String(r.number).padStart(6,'0')}`);m.className='meta';card.append(m);if(r.excerpt)card.append(el('p',r.excerpt));if(r.source_missing)card.append(el('p','原本で確認できない回答です'));card.onclick=()=>openResponse(r.id);$('responses').append(card);
 }
}
function renderTab(){
 if(!cacheSession())return;const tab=tabs[activeKind];showSummary();showCards(tab.state==='ready'?tab.items.slice(0,tab.shown):[]);$('more').hidden=tab.state!=='ready'||tab.shown>=tab.items.length;
 $('count').textContent=tab.state==='ready'?(tab.items.length?'表示中：'+Math.min(tab.shown,tab.items.length)+'件':'条件に合う回答はありません。'):'';if(currentDetail)return;
 if(tab.state==='error')$('status').textContent=errorText(tab.error);else if(tab.state!=='ready')$('status').textContent=names[activeKind]+'の回答を読み込んでいます…';else if(loading)$('status').textContent='回答を準備しています…';else if(summaryError)$('status').textContent='一覧を更新しました。集計の取得に失敗しました。「検索・更新」で再試行してください。';else if(!summaryData.surveys.length){$('status').textContent='アンケート回答は公開準備中です。';$('count').textContent='準備が整うと、ここに回答が表示されます。';}else $('status').textContent='回答を更新しました。';
}
function loadProgress(){if(loading)$('busy-note').textContent=`メーカー ${tabs.maker.received}件・お客様 ${tabs.customer.received}件を取得しました。両方の回答を準備しています。`;}
async function fetchKind(kind,seq){
 const rows=[],ids=new Set(),cursors=new Set();let cursor=null;do{
  const data=await api('list',{...filter,kind,cursor},seq);if(seq!==generation||!cacheSession())throw Error('STALE_VIEW');applyRole(data);for(const r of data.items)if(r.survey_key===kind&&!ids.has(r.id)){ids.add(r.id);rows.push(r);}cursor=data.next;tabs[kind].received=rows.length;loadProgress();
  if(rows.length>10000)throw Error('SURVEY_UNAVAILABLE');if(cursor){const key=JSON.stringify(cursor);if(cursors.has(key))throw Error('SURVEY_UNAVAILABLE');cursors.add(key);}
 }while(cursor);tabs[kind]={items:rows,state:'ready',error:null,shown:20,received:rows.length};renderTab();
}
async function load(){
 if(!endpoint){$('candidate').hidden=false;lock('本番接続は準備中です。');return;}const s=session();if(!s){lock();return;}filter=collect();const seq=++generation;
 currentToken=s.token;userRoleSeen=false;admin=authorized=false;tabs=freshTabs();details.clear();summaryData={surveys:[]};summaryError=null;loading=true;hideDetail();feedback('');$('gate').hidden=$('load-error').hidden=true;busy(true);for(const kind of Object.keys(names))tabs[kind].state='loading';renderTab();
 const tasks=Object.keys(names).map(kind=>fetchKind(kind,seq).catch(e=>{if(seq!==generation)return;tabs[kind]={items:[],state:'error',error:e,shown:20,received:0};renderTab();}));
 tasks.push(api('summary',{},seq).then(data=>{if(seq!==generation||!cacheSession())return;applyRole(data);summaryData=data;renderTab();}).catch(e=>{if(seq===generation)summaryError=e;}));
 const linked=location.hash.startsWith('#response=')?location.hash.slice(10):null;
 if(linked)tasks.push(api('detail',{id:linked},seq).then(r=>{if(seq===generation&&cacheSession())details.set(linked,r);}).catch(()=>{}));
 await Promise.allSettled(tasks);if(seq!==generation)return;loading=false;renderTab();
 if(!authorized){$('workspace').hidden=true;$('load-error').hidden=false;$('status').textContent='回答を読み込めませんでした。再読み込みをお試しください。';busy(false);return;}
 if(location.hash.startsWith('#response='))await detail(location.hash.slice(10));else busy(false);
}
function runLoad(){load().catch(e=>{if(e.message!=='STALE_VIEW'){busy(false);report(e);}});}
function openResponse(id){
 if(!cacheSession())return;if(!currentDetail){listPosition=window.scrollY;listFocusId=id;}
 const previousHash=currentDetail?'#response='+currentDetail.id:'';history.pushState(null,'','#response='+id);detail(id,previousHash);
}
// 日本語コメント: 通信中は旧画面を保持し、取得後に画面を一度だけ切り替える。
async function detail(id,previousHash=currentDetail?'#response='+currentDetail.id:''){
 if(!cacheSession())return;const view=++viewGeneration,seq=generation,focus=document.activeElement;pendingDetail={view,seq,previousHash,focus};feedback('');busy(true,'回答を開いています','内容を読み込んでから詳細画面へ移動します。',true);
 try{const r=details.get(id)||await api('detail',{id},seq);if(view!==viewGeneration||seq!==generation)return;if(!cacheSession())return;
  details.set(id,r);selectKind(r.survey_key);renderTab();currentDetail=r;showDetail(r);pendingDetail=null;busy(false);window.scrollTo(0,0);$('detail-title').focus({preventScroll:true});
 }catch(e){if(view!==viewGeneration||seq!==generation)return;pendingDetail=null;history.replaceState(null,'',location.pathname+location.search+previousHash);busy(false);report(e);focus?.focus({preventScroll:true});}
}
function cancelDetail(){if(!pendingDetail)return;const p=pendingDetail;pendingDetail=null;viewGeneration++;history.replaceState(null,'',location.pathname+location.search+p.previousHash);busy(false);p.focus?.focus({preventScroll:true});}
function returnList(restore=true){
 if(!cacheSession())return;clearDetailLink();hideDetail();feedback('');renderTab();busy(false);window.scrollTo(0,restore?listPosition:0);
 const card=[...$('responses').children].find(n=>n.dataset.id===listFocusId);if(restore)card?.focus({preventScroll:true});
}
function detailNavigation(r){
 const nav=el('nav');nav.className='detail-nav';nav.setAttribute('aria-label','回答間の移動');const close=el('button','一覧に戻る');close.className='secondary';close.onclick=()=>returnList();
 const rows=tabs[r.survey_key].items,index=rows.findIndex(x=>x.id===r.id),prev=el('button','前に戻る'),next=el('button','次に進む');prev.disabled=index<=0;next.disabled=index<0||index>=rows.length-1;prev.onclick=()=>openResponse(rows[index-1].id);next.onclick=()=>openResponse(rows[index+1].id);
 const position=el('span',index<0?'この回答は一覧の検索条件外です':`${index+1} / ${rows.length}件`);position.className='detail-position';nav.append(close,prev,next,position);return nav;
}
function showDetail(r){
 $('detail').replaceChildren();const heading=el('h2',names[r.survey_key]+'アンケートの回答');heading.id='detail-title';heading.tabIndex=-1;$('detail').append(detailNavigation(r),heading,el('p',date(r.submitted_at)));
 if(r.person)$('detail').append(el('p','記名内容（原文）：'+r.person));if(r.company)$('detail').append(el('p','会社・ブランド：'+r.company));if(r.salon)$('detail').append(el('p','サロン：'+r.salon));if(r.updated_at!==r.first_seen_at)$('detail').append(el('p','原本の変更を反映しています。'));
 if(r.source_missing){const p=el('p','原本で確認できない回答です。管理者が確認してください。');p.className='warning';$('detail').append(p);}for(const i of r.items){const box=el('article');box.className='answer';box.append(el('h3',i.title));if(i.rows?.length){for(let n=0;n<i.rows.length;n++)box.append(el('p',i.rows[n]+'：'+answer(i.answer?.[n])));}else box.append(el('p',answer(i.answer)));$('detail').append(box);}
 $('detail').append(detailNavigation(r));$('list-view').hidden=true;$('detail').hidden=false;$('status').textContent='回答の詳細です。';
}
function answer(v){if(v===null||v===undefined||v===''||(Array.isArray(v)&&!v.length))return '未回答';return Array.isArray(v)?v.map(answer).join(' ／ '):String(v);}
async function health(){if(!admin||!cacheSession())return;const seq=generation,d=await api('health',{},seq);if(seq!==generation||!cacheSession())return;$('health').replaceChildren();for(const s of d.surveys)$('health').append(el('p',`${names[s.key]}：最終同期 ${date(s.last_successful_sync_at)} ／ 通知 ${s.notify_enabled?'ON':'OFF'} ／ バックアップ ${date(s.last_backup_at)}`));for(const [key,label] of [['pending','通知待ち'],['retry_wait','再試行待ち'],['sent','API受付成功'],['failed','失敗'],['unknown','結果不明']])$('health').append(el('p',`${label}：${d.delivery[key]||0}件`));}
$('filters').onsubmit=e=>{e.preventDefault();clearDetailLink();runLoad();};$('more').onclick=()=>{if(cacheSession()){tabs[activeKind].shown+=20;renderTab();}};$('retry').onclick=$('load-retry').onclick=runLoad;$('health-refresh').onclick=()=>health().catch(report);$('busy-cancel').onclick=cancelDetail;
for(const key of Object.keys(names)){const tab=$('tab-'+key);tab.onclick=()=>{if(key===activeKind||!cacheSession())return;clearDetailLink();hideDetail();feedback('');selectKind(key);renderTab();window.scrollTo(0,0);};tab.onkeydown=e=>{if(['ArrowLeft','ArrowRight','Home','End'].includes(e.key)){e.preventDefault();const next=e.key==='Home'?'maker':e.key==='End'?'customer':key==='maker'?'customer':'maker';$('tab-'+next).focus();$('tab-'+next).click();}};}
addEventListener('hashchange',()=>{if(!currentToken||loading)return;if(location.hash.startsWith('#response=')){if(!currentDetail){listPosition=window.scrollY;listFocusId=location.hash.slice(10);}detail(location.hash.slice(10));}else returnList();});
addEventListener('keydown',e=>{if(e.key==='Escape'&&pendingDetail){e.preventDefault();cancelDetail();}});
setInterval(()=>{if(currentToken&&(!session()||session().token!==currentToken))lock('ログイン状態が変わりました。再確認してください。');},1000);
addEventListener('pageshow',e=>{if(e.persisted){lock('ログイン状態を確認しています…');runLoad();}else if(currentToken&&!session())lock();});
runLoad();
