// 日本語コメント: UIは氏名等をtextContentで表示し、サーバー操作・復旧は状態機械へ委譲する。
import {createWalkin,BUSINESS_TYPES,possibleMatches} from './walkin.mjs';
export const VERSION='0.1.0';
export async function mountWalkin({enqueue,stopScan}={}){
 const el=id=>document.getElementById(id),candidate=window.BFEventCandidate,config=window.BF_EVENT_CANDIDATE_CONFIG||{};
 const call=async(action,data)=>{
  // 通信が返らないまま別タブのロックも保持し続けない。中断後も操作IDは台帳に残る。
  const abort=new AbortController(),timer=setTimeout(()=>abort.abort(),15000);
  try{const r=await candidate.call(action,data,{signal:abort.signal});let json;try{json=await r.json();}catch{throw Error('NETWORK_RESPONSE_UNKNOWN');}if(!json.success)throw Error(json.error||'BACKEND_FAILED');return json.data;}finally{clearTimeout(timer);}
 };
 const errors={WALKIN_OWNER_CHANGED:'前の担当者による処理が残っています。前の担当者でログインし直すか、責任者に引き継いでください。',WALKIN_STORAGE_FAILED:'端末に進捗を保存できません。この端末を消去せず、責任者へ。',WALKIN_RECOVERY_INVALID:'保存済みの進捗を読み取れません。消去せず責任者へ引き継いでください。',WALKIN_EXISTING_REVIEW_REQUIRED:'同じお名前の登録が見つかりました。候補をご確認ください。',SPARE_ALREADY_ASSIGNED:'その名札は他の方に割り当てられました。別の未使用名札を選んでください。',SPARE_UNAVAILABLE:'その名札は使えません。名札一覧を更新してください。',WALKIN_ROSTER_NOT_READY:'割当を名簿で確認できません。登録し直さず、少し待って同じ入場操作を再開してください。',INVALID_APPLICANT:'お名前・サロン名・業態を確認してください。',WALKIN_QUEUE_FAILED:'入場記録を端末に保存できませんでした。登録と名札の割当は保持しています。',WALKIN_IN_PROGRESS:'途中の当日登録があります。続きから進めてください。',VALIDATION_READ_ONLY:'この検証接続は保存停止中です。',FORBIDDEN:'この操作の権限がありません。進捗を消さず責任者へ。'};
 let busy=false,open=false,searched=false,blocked=false;
 const message=e=>{el('walkinMessage').textContent=e?(errors[e.message]||'通信・認証の確認が必要です。再登録せず、表示された再確認・入場ボタンから続けてください。')+'\n'+e.message:'';};
 const fields=()=>({staff_name:el('walkinName').value.trim(),salon_name:el('walkinSalon').value.trim(),business_type:el('walkinBusiness').value});
 const lockName='bf-walkin::'+[config.endpoint,config.eventId,config.backendGeneration].join('|');
 const controller=createWalkin({storage:candidate.storage,call,enqueue,now:()=>{
  const d=new Date(),p=n=>String(n).padStart(2,'0');return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
 },withLock:run=>{if(!navigator.locks)throw Error('WALKIN_LOCK_UNAVAILABLE');return navigator.locks.request(lockName,run);}});
 window.BFEventWalkin={hasUnfinished:()=>controller.hasUnfinished(),clearFinished:()=>controller.next()};
 BUSINESS_TYPES.forEach(t=>{const o=document.createElement('option');o.value=t;o.textContent=t;el('walkinBusiness').append(o);});
 async function loadSpares(){
  const s=controller.state(),rows=await controller.spares(),select=el('walkinSpare');select.replaceChildren();
  const free=rows.filter(r=>!r.retired&&(!r.assigned_app_id||r.assigned_app_id===s.appId));
  const own=free.filter(r=>r.assigned_app_id===s.appId);const choices=own.length?own:free;
  for(const row of choices){const o=document.createElement('option');o.value=row.spare_no;o.textContent=row.spare_no+(row.assigned_app_id?'（この方に割当済み）':'');select.append(o);}
  el('walkinAssign').disabled=busy||!choices.length;
  if(!choices.length)el('walkinMessage').textContent='未使用の予備名札がありません。登録済みの方を作り直さず、責任者へ引き継いでください。';
 }
 async function draw(){
  const s=controller.state();el('walkinOpen').textContent=s?'当日登録の続きを開く':'未申込の方を登録';
  el('walkinPanel').hidden=!open;el('walkinForm').hidden=!!s;
  const phases={creating:'登録の結果を確認中です。同じ操作で再確認します。',registered:'申込登録済み／名札と入場はまだです。',assigning:'名札割当の結果を確認中です。別の札へ変更しないでください。',assigned:'名札割当済み／入場はまだです。',enqueuing:'入場記録を確認中です。同じ記録で続けます。',done:'入場を端末に記録しました。送信状況は上部をご確認ください。'};
  el('walkinStatus').textContent=s?phases[s.phase]:'必要情報を入れ、登録済みか確認してください。';
  el('walkinSummary').hidden=!s;if(s)el('walkinSummary').textContent=s.fields.staff_name+' 様\n'+s.fields.salon_name+(s.appId?'\n受付番号 '+s.appId:'')+(s.spareNo?'／予備名札 '+s.spareNo:'');
  el('walkinSparePanel').hidden=s?.phase!=='registered';el('walkinResumePanel').hidden=!['creating','assigning'].includes(s?.phase);el('walkinCheckinPanel').hidden=!['assigned','enqueuing'].includes(s?.phase);el('walkinDonePanel').hidden=s?.phase!=='done';
  el('walkinConfirm').hidden=!searched;el('walkinRegister').disabled=!el('walkinReviewed').checked;
  if(s?.phase==='registered'&&open)await loadSpares();
 }
 async function run(task){if(busy||blocked)return;busy=true;message(null);el('walkinStatus').textContent='処理中です。繰り返し押さずにお待ちください。';el('walkinPanel').querySelectorAll('button').forEach(b=>b.disabled=true);
  try{await task();}catch(e){message(e);}finally{busy=false;el('walkinPanel').querySelectorAll('button').forEach(b=>b.disabled=false);try{await draw();}catch(e){message(e);}}
 }
 const search=async()=>{
  const f=fields();if(!f.staff_name||!f.salon_name||!f.business_type)throw Error('INVALID_APPLICANT');
  const roster=await call('scanRoster',{}),matches=possibleMatches(roster,f);el('walkinMatches').replaceChildren();
  const note=document.createElement('p');note.className='hint';note.textContent=matches.length?'同じお名前の登録があります。サロン名も確認し、該当する方なら下のボタンから続けます。':'同じお名前の登録は見つかりませんでした。内容を確認して当日登録してください。';el('walkinMatches').append(note);
  for(const person of matches){const row=document.createElement('div');row.className='wiMatch';const name=document.createElement('div');name.textContent=person.staff_name+'／'+person.salon_name+'／'+person.app_id;const b=document.createElement('button');b.className='ghost';b.textContent='この方の名札へ';b.onclick=()=>run(()=>controller.useExisting(person.app_id));row.append(name,b);el('walkinMatches').append(row);}
  searched=true;el('walkinReviewed').checked=false;
 };
 el('walkinSearch').onclick=()=>run(search);
 el('walkinReviewed').onchange=()=>{el('walkinRegister').disabled=!el('walkinReviewed').checked;};
 for(const id of ['walkinName','walkinSalon','walkinBusiness'])el(id).addEventListener('input',()=>{searched=false;el('walkinReviewed').checked=false;el('walkinConfirm').hidden=true;el('walkinMatches').replaceChildren();});
 el('walkinRegister').onclick=()=>run(async()=>{if(!searched||!el('walkinReviewed').checked)throw Error('WALKIN_EXISTING_REVIEW_REQUIRED');await controller.register(fields(),{reviewed:true});});
 el('walkinAssign').onclick=()=>run(()=>controller.assign(el('walkinSpare').value));
 el('walkinReloadSpares').onclick=()=>run(loadSpares);
 el('walkinResume').onclick=()=>run(()=>controller.state().phase==='creating'?controller.register():controller.assign());
 el('walkinCheckin').onclick=()=>run(()=>controller.checkin());
 el('walkinNext').onclick=()=>run(async()=>{await controller.next();searched=false;for(const id of ['walkinName','walkinSalon','walkinBusiness'])el(id).value='';el('walkinReviewed').checked=false;el('walkinMatches').replaceChildren();});
 el('walkinClose').onclick=()=>{open=false;el('walkinPanel').hidden=true;};
 el('walkinOpen').onclick=()=>{stopScan();open=true;run(async()=>{});};
 if(config.validationReadOnly===true||!navigator.locks){blocked=true;el('walkinOpen').disabled=true;const note=document.createElement('p');note.className='hint';note.textContent=config.validationReadOnly?'この接続は保存停止中です。当日登録の保存試験はローカル候補で行います。':'この端末では当日登録の安全な同時操作制御を使えません。管理担当へ引き継いでください。';el('walkinCard').append(note);}
 try{open=!!controller.state();await draw();}catch(e){blocked=true;open=true;el('walkinPanel').hidden=false;el('walkinForm').hidden=true;el('walkinOpen').disabled=true;message(e);}
 return controller;
}
