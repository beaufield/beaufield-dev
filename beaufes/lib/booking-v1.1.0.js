/* セミナー予約。本人と社員で同じ表示・保存契約を使い、来場/QRの処理とは独立させる。 */
(function () {
  'use strict';
  const mock = location.pathname.includes('/test/') || new URLSearchParams(location.search).get('mock') === '1';
  const uid = () => crypto.randomUUID();
  const STAFF_CONTACT_NOTICE = 'ご予約は営業担当にご連絡ください';
  const words = {
    BOOKING_CONFLICT:'別の画面から予約が更新されています。最新の内容を確認してから、もう一度操作してください。',
    BOOKING_REFRESH_REQUIRED:'画面を再読み込みして、予約内容をご確認ください。',
    BOOKING_RECOVERY_REQUIRED:'予約の保存結果を確認しています。再読み込みで現在の予約をご確認ください。',
    APPLICATION_CANCELLED:'来場申込が取り消されているため、予約できません。',
    LOCK_BUSY:'ただいま予約が混み合っています。少し待ってからもう一度お試しください。',
    NO_TOKEN:'ポータルからログインし直してください。',NO_ROLE:'ビューフェスの利用権限がありません。',
    SESSION_INVALID:'ログインの有効期限が切れました。ポータルからログインし直してください。'
  };
  const mockCatalog = [
    {session_id:'S1',title:'オトナ女子の「巡り」に寄り添うフェムケアセミナー',starts_at:'11:30',ends_at:'12:30',capacity:20,booking_channel:'public',overview:'年齢による身体の変化を知り、今日から取り入れられるフェムケアをご紹介します。'},
    {session_id:'S2',title:'整形級！炭酸ガスパック体験会',starts_at:'14:00',ends_at:'15:00',capacity:20,booking_channel:'public',overview:'「整形級！」とも言われ有名芸能人の愛用者も多い炭酸ガスパック。\n毛穴、キメ、ツヤ、お肌の悩みに応えるグローパックが、なぜこんなに愛用されるのか、どのような使い方をすればいいのか、実際に体験しながら徹底ガイドいたします！'},
    {session_id:'B1',slot:'バランス革命 無料計測会',title:'10:00〜10:45の回',starts_at:'10:00',ends_at:'10:45',capacity:1,booking_channel:'public'},
    {session_id:'B2',slot:'バランス革命 無料計測会',title:'10:45〜11:30の回',starts_at:'10:45',ends_at:'11:30',capacity:1,booking_channel:'public'},
    {session_id:'B3',slot:'バランス革命 無料計測会',title:'11:30〜12:15の回',starts_at:'11:30',ends_at:'12:15',capacity:1,booking_channel:'public'},
    {session_id:'B4',slot:'バランス革命 無料計測会',title:'12:45〜13:30の回',starts_at:'12:45',ends_at:'13:30',capacity:1,booking_channel:'public'},
    {session_id:'B5',slot:'バランス革命 無料計測会',title:'13:30〜14:15の回',starts_at:'13:30',ends_at:'14:15',capacity:1,booking_channel:'public'},
    {session_id:'B6',slot:'バランス革命 無料計測会',title:'14:15〜15:00の回',starts_at:'14:15',ends_at:'15:00',capacity:1,booking_channel:'public'},
    {session_id:'B7',slot:'バランス革命 無料計測会',title:'15:00〜15:45の回',starts_at:'15:00',ends_at:'15:45',capacity:1,booking_channel:'public'},
    {session_id:'F1',slot:'Face Maker体験会',title:'10:10〜11:00の回',starts_at:'10:10',ends_at:'11:00',capacity:1,booking_channel:'staff'},
    {session_id:'F2',slot:'Face Maker体験会',title:'11:00〜11:50の回',starts_at:'11:00',ends_at:'11:50',capacity:1,booking_channel:'staff'},
    {session_id:'F3',slot:'Face Maker体験会',title:'11:50〜12:40の回',starts_at:'11:50',ends_at:'12:40',capacity:1,booking_channel:'staff'},
    {session_id:'F4',slot:'Face Maker体験会',title:'12:40〜13:30の回',starts_at:'12:40',ends_at:'13:30',capacity:1,booking_channel:'staff'},
    {session_id:'F5',slot:'Face Maker体験会',title:'13:30〜14:20の回',starts_at:'13:30',ends_at:'14:20',capacity:1,booking_channel:'staff'},
    {session_id:'F6',slot:'Face Maker体験会',title:'14:20〜15:10の回',starts_at:'14:20',ends_at:'15:10',capacity:1,booking_channel:'staff'},
    {session_id:'F7',slot:'Face Maker体験会',title:'15:10〜16:00の回',starts_at:'15:10',ends_at:'16:00',capacity:1,booking_channel:'staff'}
  ];
  const mockSaved={};
  function mockReservedCount(sessionId) {
    return Object.values(mockSaved).filter(ids=>ids.includes(sessionId)).length;
  }
  function mockView(key, actor) {
    const reserved=mockSaved[key] || [];
    const sessions=mockCatalog.map(s=>Object.assign({},s,{
      remaining:Math.max(0,s.capacity-mockReservedCount(s.session_id)),
      can_book:s.booking_channel==='public' || actor==='staff',
      can_cancel:actor==='staff' || s.booking_channel==='public'
    }));
    return {sessions:sessions,reserved:reserved,booking_revision:'mock',booking_open:true};
  }
  async function request(url,action,data) {
    if(mock) {
      const key=data.app_id || data.ticket_token || 'TESTTOKEN';
      const actor=action.startsWith('staff')?'staff':'customer';
      if(action==='reserveSessions'||action==='staffChangeReservation') {
        const current=mockSaved[key] || [];
        const allowed=new Set(mockCatalog.filter(s=>s.booking_channel==='public'||actor==='staff').map(s=>s.session_id));
        mockSaved[key]=data.sessions.filter(id=>allowed.has(id)||current.includes(id));
      }
      return Object.assign(mockView(key,actor),{result:{reserved:mockSaved[key]||[],full:[],conflict:[],invalid:[]},mail_sent:false});
    }
    let last;
    // 同じdata・操作IDを再送する。タイムアウトしても別の予約操作にしない。
    for(let attempt=0;attempt<3;attempt++) {
      if(attempt)await new Promise(resolve=>setTimeout(resolve,attempt*1000));
      const ctrl=new AbortController(),timer=setTimeout(()=>ctrl.abort(),30000);
      try {
        const res=await fetch(url,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({action:action,data:JSON.stringify(data)}),signal:ctrl.signal});
        if(!res.ok)throw new Error('NETWORK');
        const json=JSON.parse(await res.text());
        if(!json.success){const e=new Error(json.error || 'UNKNOWN');e.business=true;throw e;}
        return json.data;
      } catch(e) {last=e;if(e.business)throw e;} finally{clearTimeout(timer);}
    }
    throw last;
  }
  function node(tag,text,cls){const n=document.createElement(tag);if(text)n.textContent=text;if(cls)n.className=cls;return n;}
  function displayTitle(s) {
    return s.slot && s.starts_at && String(s.title).includes(s.starts_at) ? s.slot+' '+s.title : s.title;
  }
  function mount(opts) {
    const root=typeof opts.root==='string'?document.querySelector(opts.root):opts.root;
    if(!root)return null;
    root.classList.add('booking-box');
    let current=null,ready=false,busy=false,confirming=false,requestId=uid(),chosen=[],token=opts.token || '',pending=null,sequence=0;
    function credentials(){return opts.mode==='staff'?{session_token:opts.sessionToken(),app_id:opts.app.app_id}:{ticket_token:token};}
    const status=node('p','','booking-status');status.setAttribute('role','status');
    const options=node('div');
    const title=node('h3',opts.mode==='form'?'セミナーのご予約（任意）':'セミナーのご予約');
    root.replaceChildren(title);
    if(opts.app)root.append(node('p',opts.app.salon_name+' ／ '+opts.app.staff_name+' 様\n受付番号 '+opts.app.app_id,'booking-person'));
    root.append(node('p','ご予約がなくてもビューフェスにご入場いただけます。','booking-help'),options,status);
    const save=node('button','選択した内容で保存する','booking-save');save.type='button';
    const reload=node('button','最新の予約を確認する','booking-reload');reload.type='button';
    if(opts.mode!=='form')root.append(save);
    root.append(reload);
    function render(data) {
      current=data;ready=true;chosen=(data.reserved||[]).slice();options.replaceChildren();
      const sessions=data.sessions||[];
      if(!sessions.length)options.append(node('p','セミナーの予約は準備中です。'));
      sessions.forEach(s=>{
        const reserved=chosen.includes(s.session_id),label=node('label','','booking-option');
        const input=node('input');input.type='checkbox';input.value=s.session_id;input.checked=reserved;
        input.disabled=busy || (reserved ? !s.can_cancel : (!s.can_book || s.remaining===0));
        input.addEventListener('change',()=>{chosen=Array.from(options.querySelectorAll('input:checked')).map(x=>x.value);requestId=uid();pending=null;status.textContent='';});
        const remaining=s.remaining===null?'空きあり':'空き'+s.remaining+'席';
        const state=reserved?'予約済み':s.remaining===0?'満席':s.can_book?'残り'+s.remaining+'席':remaining;
        const copy=node('span');copy.append(node('strong',displayTitle(s)),node('span',s.starts_at+'〜'+s.ends_at+' ／ '+state,'booking-time'));
        if(opts.mode!=='staff'&&s.booking_channel==='staff')copy.append(node('span',STAFF_CONTACT_NOTICE,'booking-time'));
        // 社員限定枠は、お客様画面では枠と空き状況だけを表示する。
        if(s.overview && !(opts.mode!=='staff'&&s.booking_channel==='staff')){const details=node('details');details.append(node('summary','セミナー内容を見る'),node('p',s.overview));copy.append(details);}
        label.append(input,copy);options.append(label);
      });
      save.disabled=busy || !sessions.some(s=>s.can_book || (chosen.includes(s.session_id)&&s.can_cancel));
    }
    async function load() {
      const seq=++sequence;ready=false;save.disabled=true;reload.disabled=true;status.textContent='空き状況と予約内容を確認しています…';
      options.querySelectorAll('input').forEach(x=>x.disabled=true);
      try{
        const data=await request(opts.url,opts.mode==='staff'?'staffGetReservations':'listSessions',credentials());
        if(seq!==sequence)return;
        render(data);requestId=uid();pending=null;status.textContent='';
        if(new URLSearchParams(location.search).has('semissue'))status.textContent='来場のお申し込みは完了しています。セミナー予約は下の現在の内容を確認し、必要ならお取りください。';
      }catch(e){if(seq!==sequence)return;status.textContent=words[e.message]||'予約状況を確認できませんでした。来場申込・入場パスはそのまま利用できます。';}
      finally{reload.disabled=false;}
    }
    function payload(){return ready?{sessions:chosen.slice(),booking_request_id:requestId,booking_revision:current.booking_revision}:{};}
    // ブラウザの警告ダイアログに移らず、対象者と取消内容を画面内で確認する。
    function approve(text){
      confirming=true;save.disabled=true;reload.disabled=true;options.querySelectorAll('input').forEach(x=>x.disabled=true);
      const box=node('div','','booking-confirm'),message=node('p',text),yes=node('button','確認して予約を保存する','booking-save'),no=node('button','戻る','booking-reload');
      yes.type=no.type='button';box.setAttribute('role','alertdialog');box.setAttribute('aria-label','予約内容の確認');box.append(message,yes,no);root.append(box);yes.focus();
      return new Promise(resolve=>{
        function finish(value){box.remove();confirming=false;save.disabled=false;reload.disabled=false;options.querySelectorAll('input').forEach(x=>{const s=current.sessions.find(s=>s.session_id===x.value);x.disabled=x.checked?!s.can_cancel:(!s.can_book||s.remaining===0);});resolve(value);}
        yes.addEventListener('click',()=>finish(true));no.addEventListener('click',()=>finish(false));
      });
    }
    save.addEventListener('click',async()=>{
      if(busy||confirming||!ready)return;
      const removing=(current.reserved||[]).filter(id=>!chosen.includes(id));
      let confirmation=opts.mode==='staff'?opts.app.salon_name+' ／ '+opts.app.staff_name+' 様の予約を保存します。':'';
      if(removing.length)confirmation+=' 選択を外したセミナー予約を取り消します。来場のお申し込みは取り消されません。';
      if(confirmation&&!await approve(confirmation))return;
      pending=pending||Object.assign(credentials(),payload());busy=true;save.disabled=true;reload.disabled=true;options.querySelectorAll('input').forEach(x=>x.disabled=true);status.textContent='予約内容を保存しています…';
      try{
        const data=await request(opts.url,opts.mode==='staff'?'staffChangeReservation':'reserveSessions',pending);
        pending=null;requestId=uid();busy=false;render(data);
        const r=data.result||{};
        status.textContent=(r.full||[]).length?'満席のため変更できませんでした。元の予約を保持しています。':(r.conflict||[]).length?'他の予約と時間が重なるため変更できませんでした。':(r.invalid||[]).length?'受付対象外のため変更できませんでした。':'現在の予約内容を保存しました。';
        if(opts.mode==='staff')status.textContent+=' 自動メールは送信していません。お客様へのご案内をお願いします。';
        else if(data.mail_sent)status.textContent+=' 控えメールをお送りしました。';
        if(opts.changed)opts.changed();
      }catch(e){status.textContent=words[e.message]||'保存結果を確認できませんでした。同じ内容で再試行するか、最新の予約を確認してください。';if(e.message==='BOOKING_CONFLICT'){ready=false;pending=null;}}
      finally{busy=false;save.disabled=!ready;reload.disabled=false;if(ready)options.querySelectorAll('input').forEach(x=>{const s=current.sessions.find(s=>s.session_id===x.value);x.disabled=x.checked?!s.can_cancel:(!s.can_book||s.remaining===0);});}
    });
    reload.addEventListener('click',load);
    load();
    return {payload:payload,setToken(t){token=t;return load();},load:load};
  }
  function staff(opts){
    const dialog=node('dialog','','booking-dialog'),close=node('button','閉じる','booking-reload'),root=node('div');close.type='button';
    dialog.append(close,root);document.body.append(dialog);close.addEventListener('click',()=>dialog.close());dialog.addEventListener('close',()=>dialog.remove());
    dialog.showModal();mount(Object.assign({},opts,{root:root,mode:'staff'}));
  }
  window.BeaufesBooking={mount:mount,staff:staff,isMock:()=>mock,
    mockStaffSummary:()=>({sessions:mockCatalog.map(s=>Object.assign({},s,{is_active:true,reserved_count:mockReservedCount(s.session_id),remaining:Math.max(0,s.capacity-mockReservedCount(s.session_id))})),reservations:Object.keys(mockSaved).flatMap(id=>mockSaved[id].map(s=>({app_id:id,session_id:s,app_status:'confirmed'})))}),
    hasIssue:r=>!!(r&&(r.error||(r.full||[]).length||(r.conflict||[]).length||(r.invalid||[]).length)),
    mockResponse:params=>({success:true,data:params.action==='getPass'?{app_id:'MOCK-001',salon_name:'確認用サロン',staff_name:'確認用',event:{date:'2026-10-26',time:'10:00〜16:00',venue_name:'青島屋2F'}}:{app_id:'MOCK-001',pass_url:'pass.html?t=TESTTOKEN',sessions_result:{}}})
  };
})();
