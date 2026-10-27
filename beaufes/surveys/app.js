// 日本語コメント: 回答はメモリー内だけに保持。共有セッションを編集せずサーバーで毎回検証する。
const APP_VERSION='v0.1.4';
const $=id=>document.getElementById(id),el=(tag,text)=>{const n=document.createElement(tag);if(text!==undefined)n.textContent=text;return n;};
const endpoint=document.querySelector('meta[name=survey-api]').content;
const portal=document.querySelector('meta[name=portal-url]').content;
const names={maker:'メーカー',customer:'お客様'};
const date=value=>value?new Intl.DateTimeFormat('ja-JP',{timeZone:'Asia/Tokyo',dateStyle:'short',timeStyle:'short'}).format(new Date(value)):'まだ同期されていません';
let generation=0,currentToken=null,cursor=null,shown=0,filter={},admin=false;
$('version').textContent=APP_VERSION;$('portal').href=$('signin').href=portal;
const session=()=>{try{const s=JSON.parse(localStorage.getItem('bf_session')||'null');return s&&typeof s.token==='string'&&Number.isSafeInteger(s.expires)&&s.expires>Date.now()?s:null;}catch{return null;}};
function lock(message='社員用ポータルでログインしてください。'){
 generation++;currentToken=null;cursor=null;shown=0;admin=false;
 for(const id of ['responses','detail','summary','health'])$(id).replaceChildren();
 $('count').textContent='';$('workspace').hidden=true;$('gate').hidden=false;$('detail').hidden=$('admin').hidden=true;$('status').textContent=message;
}
const errors={PORTAL_SESSION_INVALID:'ログインの有効期限が切れています。ポータルでログインし直してください。',BEAUFES_ROLE_REQUIRED:'ビューフェスの閲覧権限がありません。管理者へ確認してください。',PORTAL_UNAVAILABLE:'社員認証を確認できません。時間をおいて再試行してください。',RESPONSE_NOT_FOUND:'回答が見つからないか、閲覧が停止されています。',SURVEY_UNAVAILABLE:'回答を取得できません。時間をおいて更新してください。',INVALID_FILTER:'検索条件を確認してください。',INVALID_RESPONSE_ID:'回答リンクを確認してください。'};
function report(e){if(e.message==='STALE_VIEW')return;$('status').textContent=errors[e.message]||'通信を完了できませんでした。接続を確認して再試行してください。';}
async function api(action,body,seq){
 const s=session();if(!s){lock();throw Error('STALE_VIEW');}
 const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),45000);
 try{
  const r=await fetch(endpoint,{method:'POST',headers:{'Content-Type':'application/json',Authorization:'Bearer '+s.token},body:JSON.stringify({action,body}),cache:'no-store',redirect:'error',signal:controller.signal});
  const d=await r.json();const active=session();
  if(!active||active.token!==s.token){lock();throw Error('STALE_VIEW');}if(seq!==generation)throw Error('STALE_VIEW');
  if([401,403].includes(r.status)||d.error==='PORTAL_UNAVAILABLE'){lock(errors[d.error]||'閲覧権限を確認できません。');throw Error('STALE_VIEW');}
  if(!r.ok)throw Error(d.error||'SURVEY_UNAVAILABLE');return d;
 }finally{clearTimeout(timer);}
}
function dateFilter(value,end=false){if(!value)return '';const t=Date.parse(value+'T00:00:00+09:00');if(!Number.isFinite(t))throw Error('INVALID_FILTER');return new Date(t+(end?86400000:0)).toISOString();}
function collect(){const f={kind:$('kind').value,q:$('query').value.trim(),score:$('score').value,from:dateFilter($('from').value),until:dateFilter($('until').value,true)};if(f.from&&f.until&&f.from>=f.until)throw Error('INVALID_FILTER');return f;}
function showSummary(data){
 $('summary').replaceChildren();
 for(const s of data.surveys){const box=el('article');box.className='summary-card';box.append(el('h2',names[s.key]+'アンケート'),el('strong',s.total+'件'),el('p',s.scored?`平均 ${s.average}／5（有効回答${s.scored}件）`:'満足度の有効回答はありません'));
  const bars=el('div');bars.className='bars';for(let n=5;n>=1;n--){const bar=el('span',`${n}：${s.distribution[n]||0}件`);bar.className='bar';bars.append(bar);}box.append(bars);
  const freshness=el('p','最終同期：'+date(s.last_successful_sync_at));freshness.className='meta';if(!s.last_successful_sync_at||Date.now()-Date.parse(s.last_successful_sync_at)>900000)freshness.classList.add('warning');box.append(freshness);$('summary').append(box);
 }
}
function append(items){
 for(const r of items){const card=el('button');card.type='button';card.className='card';card.dataset.id=r.id;
  const label=el('span',names[r.survey_key]);label.className='label';card.append(label,el('span',`満足度 ${r.satisfaction??'未回答・該当なし'}`));
  card.append(el('h3',[r.company,r.salon,r.person].filter(Boolean).join(' ／ ')||'記名なし'));
  const m=el('p',`${date(r.submitted_at)} ／ ${r.survey_key==='maker'?'M':'C'}-${String(r.number).padStart(6,'0')}`);m.className='meta';card.append(m);
  if(r.excerpt)card.append(el('p',r.excerpt));if(r.source_missing)card.append(el('p','原本で確認できない回答です'));
  card.onclick=()=>{location.hash='response='+r.id;};$('responses').append(card);
 }shown+=items.length;$('count').textContent=shown?'表示中：'+shown+'件':'条件に合う回答はありません。';
}
async function load(more=false){
 if(!endpoint){$('candidate').hidden=false;lock('本番接続は準備中です。');return;}
 const s=session();if(!s){lock();return;}const seq=++generation;currentToken=s.token;
 $('status').textContent='回答を読み込んでいます…';$('more').disabled=true;
 try{
  if(!more){filter=collect();cursor=null;shown=0;$('responses').replaceChildren();$('detail').replaceChildren();$('detail').hidden=true;}
  const data=await api('list',{...filter,cursor},seq);append(data.items);cursor=data.next;
  admin=data.viewer.role==='admin';$('admin').hidden=!admin;$('workspace').hidden=false;$('gate').hidden=true;$('more').hidden=!cursor;
  if(!more){const summary=await api('summary',{},seq);showSummary(summary);if(!summary.surveys.length){$('status').textContent='アンケート回答は公開準備中です。';$('count').textContent='準備が整うと、ここに回答が表示されます。';return;}}
  $('status').textContent='回答を更新しました。';
  if(!more&&location.hash.startsWith('#response='))await detail(location.hash.slice(10));
 }finally{$('more').disabled=false;}
}
async function detail(id){
 const seq=++generation;$('detail').replaceChildren();$('detail').hidden=true;$('status').textContent='回答の詳細を読み込んでいます…';
 const r=await api('detail',{id},seq);const close=el('button','一覧へ戻る');close.onclick=()=>{location.hash='';$('detail').hidden=true;$('detail').replaceChildren();$('status').textContent='回答一覧です。';};
 $('detail').append(close,el('h2',names[r.survey_key]+'アンケートの回答'),el('p',date(r.submitted_at)));
 if(r.person)$('detail').append(el('p','記名内容（原文）：'+r.person));
 if(r.company)$('detail').append(el('p','会社・ブランド：'+r.company));if(r.salon)$('detail').append(el('p','サロン：'+r.salon));
 if(r.updated_at!==r.first_seen_at)$('detail').append(el('p','原本の変更を反映しています。'));
 if(r.source_missing){const p=el('p','原本で確認できない回答です。管理者が確認してください。');p.className='warning';$('detail').append(p);}
 for(const i of r.items){const box=el('article');box.className='answer';box.append(el('h3',i.title));
  if(i.rows?.length){for(let n=0;n<i.rows.length;n++)box.append(el('p',i.rows[n]+'：'+answer(i.answer?.[n])));}
  else box.append(el('p',answer(i.answer)));$('detail').append(box);
 }$('detail').hidden=false;$('status').textContent='回答の詳細です。';$('detail').scrollIntoView({behavior:'smooth',block:'start'});
}
function answer(v){if(v===null||v===undefined||v===''||(Array.isArray(v)&&!v.length))return '未回答';return Array.isArray(v)?v.map(answer).join(' ／ '):String(v);}
async function health(){if(!admin)return;const seq=++generation;const d=await api('health',{},seq);$('health').replaceChildren();
 for(const s of d.surveys)$('health').append(el('p',`${names[s.key]}：最終同期 ${date(s.last_successful_sync_at)} ／ 通知 ${s.notify_enabled?'ON':'OFF'} ／ バックアップ ${date(s.last_backup_at)}`));
 for(const [key,label] of [['pending','通知待ち'],['retry_wait','再試行待ち'],['sent','API受付成功'],['failed','失敗'],['unknown','結果不明']])$('health').append(el('p',`${label}：${d.delivery[key]||0}件`));
}
$('filters').onsubmit=e=>{e.preventDefault();load().catch(report);};$('more').onclick=()=>load(true).catch(report);$('retry').onclick=()=>load().catch(report);$('health-refresh').onclick=()=>health().catch(report);
addEventListener('hashchange',()=>{if(location.hash.startsWith('#response=')&&currentToken)detail(location.hash.slice(10)).catch(report);else{generation++;$('detail').hidden=true;$('detail').replaceChildren();}});
setInterval(()=>{if(currentToken&&(!session()||session().token!==currentToken))lock('ログイン状態が変わりました。再確認してください。');},1000);
addEventListener('pageshow',e=>{if(e.persisted){lock('ログイン状態を確認しています…');load().catch(report);}else if(currentToken&&!session())lock();});
load().catch(report);
