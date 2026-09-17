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
  const FEMCARE_BULLETS = [
    '更年期は"予防"できる！',
    '結局更年期の原因ってなにから来ているの？',
    '身体の変化に寄り添ったエッセンス新商品のご紹介',
    '今日からできるフェムケア'
  ];
  const FEMCARE_OVERVIEW =
    '年齢を重ねるにつれて感じる、理由のはっきりしないゆらぎや、すっきりしない毎日。\n' +
    '頑張っているのに気分が晴れない、以前とは違う自分に戸惑う、そんなオトナ女子のために生まれたのが「Meguru+（メグルプラス）」です。\n' +
    '「ゆらぐ私に穏やかな巡りを」をコンセプトに、これからの毎日を前向きに、自分らしく過ごすための新しい習慣をご提案します。\n' +
    '本セミナーでは、更年期の仕組みとその解決するための秘密を詳しくご紹介。変化の時期を、美しく心地よく迎えるためのヒントをお届けします。';
  const BALANCE_BULLETS = [
    'コンピューター精密計測で、足裏の圧力バランスと骨盤のゆがみをデータで見える化',
    '微細振動波（圧電体）による体軸の体感テスト',
    'お一人ずつのオーダーメイド インソール型骨格調整具のご提案'
  ];
  const BALANCE_OVERVIEW =
    '長引く「ひざ痛・腰痛・猫背・肩こり・疲労感」の原因は、骨格のゆがみかもしれません。\n' +
    '足裏コンピューター計測で、骨盤と足裏のバランスを精密に測定します。\n' +
    'お一人ずつ丁寧に計測・ご説明を行うため完全予約制です（各回1名様）。';
  const FACE_BULLETS = [
    '低周波の立体電場をベースに、表情筋などへアプローチ',
    '温熱・近赤外線などを組み合わせ、肌や筋肉が反応しやすい環境へ',
    '専用の「ハニカムAgマスク」で顔全体を包み、美容液とのなじみをサポート'
  ];
  const FACE_OVERVIEW =
    '「肌」だけでなく、筋肉・フェイスライン・顔全体の印象までトータルにアプローチするフェイシャル機器「Face Maker（フェイスメーカー）」をご体験いただけます。\n' +
    '目指すのは、スッキリしたフェイスライン、若々しく整った顔の印象、ハリ・うるおいのある肌印象。\n' +
    '体験メニューは「EFFプローブ 30分」＋「ハニカムAgマスク 15分」の約45分です。';
  const mockCatalog = [
    {session_id:'S1',slot:'セミナー 第1部',title:'オトナ女子の"巡り"に寄り添うフェムケアセミナー',starts_at:'11:30',ends_at:'12:30',capacity:20,booking_channel:'public',bullets:FEMCARE_BULLETS,overview:FEMCARE_OVERVIEW},
    {session_id:'S2',slot:'セミナー 第2部',title:'整形級！炭酸ガスパック体験会',starts_at:'14:00',ends_at:'15:00',capacity:20,booking_channel:'public',overview:'「整形級！」とも言われ有名芸能人の愛用者も多い炭酸ガスパック。\n毛穴、キメ、ツヤ、お肌の悩みに応えるグローパックが、なぜこんなに愛用されるのか、どのような使い方をすればいいのか、実際に体験しながら徹底ガイドいたします！'},
    {session_id:'B1',slot:'バランス革命 無料計測会',title:'10:00〜10:45の回',starts_at:'10:00',ends_at:'10:45',capacity:1,booking_channel:'public',bullets:BALANCE_BULLETS,overview:BALANCE_OVERVIEW},
    {session_id:'B2',slot:'バランス革命 無料計測会',title:'10:45〜11:30の回',starts_at:'10:45',ends_at:'11:30',capacity:1,booking_channel:'public',bullets:BALANCE_BULLETS,overview:BALANCE_OVERVIEW},
    {session_id:'B3',slot:'バランス革命 無料計測会',title:'11:30〜12:15の回',starts_at:'11:30',ends_at:'12:15',capacity:1,booking_channel:'public',bullets:BALANCE_BULLETS,overview:BALANCE_OVERVIEW},
    {session_id:'B4',slot:'バランス革命 無料計測会',title:'12:45〜13:30の回',starts_at:'12:45',ends_at:'13:30',capacity:1,booking_channel:'public',bullets:BALANCE_BULLETS,overview:BALANCE_OVERVIEW},
    {session_id:'B5',slot:'バランス革命 無料計測会',title:'13:30〜14:15の回',starts_at:'13:30',ends_at:'14:15',capacity:1,booking_channel:'public',bullets:BALANCE_BULLETS,overview:BALANCE_OVERVIEW},
    {session_id:'B6',slot:'バランス革命 無料計測会',title:'14:15〜15:00の回',starts_at:'14:15',ends_at:'15:00',capacity:1,booking_channel:'public',bullets:BALANCE_BULLETS,overview:BALANCE_OVERVIEW},
    {session_id:'B7',slot:'バランス革命 無料計測会',title:'15:00〜15:45の回',starts_at:'15:00',ends_at:'15:45',capacity:1,booking_channel:'public',bullets:BALANCE_BULLETS,overview:BALANCE_OVERVIEW},
    {session_id:'F1',slot:'Face Maker体験会',title:'10:10〜11:00の回',starts_at:'10:10',ends_at:'11:00',capacity:1,booking_channel:'staff',bullets:FACE_BULLETS,overview:FACE_OVERVIEW},
    {session_id:'F2',slot:'Face Maker体験会',title:'11:00〜11:50の回',starts_at:'11:00',ends_at:'11:50',capacity:1,booking_channel:'staff',bullets:FACE_BULLETS,overview:FACE_OVERVIEW},
    {session_id:'F3',slot:'Face Maker体験会',title:'11:50〜12:40の回',starts_at:'11:50',ends_at:'12:40',capacity:1,booking_channel:'staff',bullets:FACE_BULLETS,overview:FACE_OVERVIEW},
    {session_id:'F4',slot:'Face Maker体験会',title:'12:40〜13:30の回',starts_at:'12:40',ends_at:'13:30',capacity:1,booking_channel:'staff',bullets:FACE_BULLETS,overview:FACE_OVERVIEW},
    {session_id:'F5',slot:'Face Maker体験会',title:'13:30〜14:20の回',starts_at:'13:30',ends_at:'14:20',capacity:1,booking_channel:'staff',bullets:FACE_BULLETS,overview:FACE_OVERVIEW},
    {session_id:'F6',slot:'Face Maker体験会',title:'14:20〜15:10の回',starts_at:'14:20',ends_at:'15:10',capacity:1,booking_channel:'staff',bullets:FACE_BULLETS,overview:FACE_OVERVIEW},
    {session_id:'F7',slot:'Face Maker体験会',title:'15:10〜16:00の回',starts_at:'15:10',ends_at:'16:00',capacity:1,booking_channel:'staff',bullets:FACE_BULLETS,overview:FACE_OVERVIEW}
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
  function groupTitle(s) {
    return s.slot || s.title;
  }
  function inputDisabled(s,reserved,busy) {
    return busy || (reserved ? !s.can_cancel : (!s.can_book || s.remaining===0));
  }
  function stateText(s,reserved) {
    if(reserved)return '予約済み';
    if(s.remaining===0)return '満席';
    if(s.capacity===1)return '空き';
    if(s.remaining===null || s.remaining===undefined)return '空きあり';
    return '残り'+s.remaining+'席';
  }
  function isSeminar(s) {
    return String(s.session_id||'').startsWith('S') || String(s.slot||'').startsWith('セミナー');
  }
  function durationMinutes(s) {
    const start=/^(\d{1,2}):(\d{2})$/.exec(String(s.starts_at||''));
    const end=/^(\d{1,2}):(\d{2})$/.exec(String(s.ends_at||''));
    if(!start||!end)return null;
    return Number(end[1])*60+Number(end[2])-Number(start[1])*60-Number(start[2]);
  }
  function mount(opts) {
    const root=typeof opts.root==='string'?document.querySelector(opts.root):opts.root;
    if(!root)return null;
    root.classList.add('booking-box');
    let current=null,ready=false,busy=false,confirming=false,requestId=uid(),chosen=[],token=opts.token || '',pending=null,sequence=0;
    function credentials(){return opts.mode==='staff'?{session_token:opts.sessionToken(),app_id:opts.app.app_id}:{ticket_token:token};}
    const status=node('p','','booking-status');status.setAttribute('role','status');
    const options=node('div');
    root.replaceChildren();
    if(opts.app)root.append(node('p',opts.app.salon_name+' ／ '+opts.app.staff_name+' 様\n受付番号 '+opts.app.app_id,'booking-person'));
    root.append(options,status);
    const save=node('button','選択した内容で保存する','booking-save');save.type='button';
    const reload=node('button','最新の予約を確認する','booking-reload');reload.type='button';
    if(opts.mode!=='form')root.append(save);
    root.append(reload);
    function render(data) {
      current=data;ready=true;chosen=(data.reserved||[]).slice();options.replaceChildren();
      const sessions=data.sessions||[];
      if(!sessions.length){options.append(node('p','セミナー・体験会の予約は準備中です。'));return;}
      const seminarSessions=sessions.filter(isSeminar),experienceSessions=sessions.filter(s=>!isSeminar(s));
      const seminarSection=node('section','','booking-category booking-category-seminar');
      seminarSection.append(node('h3',opts.mode==='form'?'セミナーのご予約（任意）':'セミナーのご予約'));
      seminarSection.append(node('p','ご予約がなくてもビューフェスにご入場いただけます。','booking-help'));
      const seminarList=node('div');seminarSection.append(seminarList);options.append(seminarSection);
      const experienceSection=node('section','','booking-category booking-category-experience');
      experienceSection.append(node('h3',opts.mode==='form'?'体験会のご予約（任意）':'体験会のご予約'));
      experienceSection.append(node('p','体験会ごとに、ご希望の開始時刻を1つお選びください。','booking-help'));
      const experienceList=node('div');experienceSection.append(experienceList);options.append(experienceSection);
      function changed(input,groupInputs) {
        // 同じ体験会では時間を1つだけ選べる。別の体験会・セミナーは同時に選べる。
        if(input.checked)groupInputs.forEach(other=>{if(other!==input)other.checked=false;});
        chosen=Array.from(options.querySelectorAll('input:checked')).map(x=>x.value);
        requestId=uid();pending=null;status.textContent='';
      }
      function renderCategory(container,categorySessions,isExperience) {
        const order=[],groups={};
        categorySessions.forEach(s=>{
          const key=s.slot || '__'+s.session_id;
          if(!groups[key]){groups[key]=[];order.push(key);}
          groups[key].push(s);
        });
        order.forEach(key=>{
        const group=groups[key];
        if(group.length===1) {
          const s=group[0],reserved=chosen.includes(s.session_id),label=node('label','','booking-option');
          const input=node('input');input.type='checkbox';input.value=s.session_id;input.checked=reserved;
          input.disabled=inputDisabled(s,reserved,busy);
          input.addEventListener('change',()=>changed(input,[input]));
          const copy=node('span');copy.append(node('strong',displayTitle(s)),node('span',s.starts_at+'〜'+s.ends_at+' ／ '+stateText(s,reserved),'booking-time'));
          if(opts.mode!=='staff'&&s.booking_channel==='staff')copy.append(node('span',STAFF_CONTACT_NOTICE,'booking-time'));
          // 社員限定枠は、お客様画面では枠と空き状況だけを表示する。
          if(s.overview && !(opts.mode!=='staff'&&s.booking_channel==='staff')){const details=node('details');details.append(node('summary','セミナー内容を見る'),node('p',s.overview));copy.append(details);}
          label.append(input,copy);container.append(label);
          return;
        }

        const section=node('section','','booking-group');
        section.append(node('h4',groupTitle(group[0])));
        if(isExperience) {
          const description=node('div','','booking-experience-description');
          const bullets=Array.isArray(group[0].bullets)?group[0].bullets:[];
          if(bullets.length){const list=node('ul');bullets.slice(0,3).forEach(text=>list.append(node('li',text)));description.append(list);}
          else if(group[0].overview)description.append(node('p',String(group[0].overview).split('\n')[0]));
          const minutes=durationMinutes(group[0]);
          if(minutes!==null)description.append(node('p','所要時間：約'+minutes+'分','booking-duration'));
          section.append(description);
        }
        const hint=node('p','ご希望の時間を1つお選びください','booking-group-hint');
        const customerStaffOnly=opts.mode!=='staff' && group.every(s=>s.booking_channel==='staff');
        if(customerStaffOnly)hint.textContent='予約枠と空き状況をご確認いただけます。';
        section.append(hint);
        const grid=node('div','','booking-time-grid'),groupInputs=[];
        group.forEach(s=>{
          const reserved=chosen.includes(s.session_id),label=node('label','','booking-time-option');
          const input=node('input');input.type='checkbox';input.value=s.session_id;input.checked=reserved;
          input.disabled=inputDisabled(s,reserved,busy);input.dataset.bookingGroup=key;groupInputs.push(input);
          const time=node('span',s.starts_at+'〜','booking-time-label');
          const state=node('span',stateText(s,reserved),'booking-time-state');
          label.append(input,time,state);grid.append(label);
        });
        groupInputs.forEach(input=>input.addEventListener('change',()=>changed(input,groupInputs)));
        section.append(grid);
        if(customerStaffOnly)section.append(node('p',STAFF_CONTACT_NOTICE,'booking-staff-notice'));
        const sharedOverview=group[0].overview && group.every(s=>s.overview===group[0].overview) ? group[0].overview : '';
        if(sharedOverview){const details=node('details');details.append(node('summary','体験会の詳しい内容を見る'),node('p',sharedOverview));section.append(details);}
        container.append(section);
        });
      }
      renderCategory(seminarList,seminarSessions,false);
      renderCategory(experienceList,experienceSessions,true);
      if(!seminarSessions.length)seminarSection.hidden=true;
      if(!experienceSessions.length)experienceSection.hidden=true;
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
