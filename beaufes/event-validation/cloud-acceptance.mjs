// 日本語コメント: 架空の検証イベントだけで5並列・札競合・同一ID再送を確認する。実顧客や本番には使わない。
export const VERSION='0.1.0';
const KEY='bf-validation-cloud-acceptance-v1';
const EXPECTED='https://ovblkjxlbmnehgkpmyrd.supabase.co/functions/v1/beaufes-event-validation';
export async function runCloudAcceptance({config,call,storage,progress=()=>{},uuid=()=>crypto.randomUUID(),now=()=>new Date().toISOString()}={}){
 if(config?.environment!=='isolated'||config.eventId!=='validation-beaufes'||config.endpoint!==EXPECTED)throw Error('VALIDATION_ONLY');
 const owner=(await call('staffSession')).user_id;if(!owner)throw Error('SESSION_INVALID');
 let state=JSON.parse(storage.getItem(KEY)||'null');
 if(state&&state.owner!==owner)throw Error('OWNER_CHANGED');
 const save=()=>{const raw=JSON.stringify(state);storage.setItem(KEY,raw);if(storage.getItem(KEY)!==raw)throw Error('STORAGE_FAILED');};
 if(!state){
  const free=(await call('listSpareBadges')).spares.filter(s=>!s.assigned_app_id&&!s.retired).map(s=>s.spare_no).sort().reverse().slice(0,6);
  if(free.length<6)throw Error('SIX_TEST_SPARES_REQUIRED');
  state={version:1,owner,run:uuid(),spares:free,createdAt:now(),lanes:Array.from({length:5},(_,i)=>({index:i,create:uuid(),assign:uuid(),fallback:uuid(),checkin:uuid()}))};save();
 }
 if(state.done){progress('既に完了した試験です。新しい登録・入場を増やしません。');return state.summary;}
 progress('1/4 架空の5名を同時に登録しています。');
 await Promise.all(state.lanes.map(async lane=>{
  if(lane.app)return;const r=await call('adminCreateApplication',{request_id:lane.create,staff_name:'並列受付試験'+state.run.slice(0,8)+'-'+lane.index,salon_name:'架空同時受付試験',business_type:'美容室'});
  if(!r.app_id)throw Error('INVALID_RESULT');lane.app=r.app_id;save();
 }));
 if(new Set(state.lanes.map(x=>x.app)).size!==5)throw Error('DUPLICATE_APPLICATION');
 progress('2/4 2接続で同じ札を選び、残りの接続は別札を割り当てています。');
 await Promise.all(state.lanes.map(async lane=>{
  if(lane.spare)return;const chosen=state.spares[lane.index<2?0:lane.index-1];
  try{const r=await call('assignSpare',{request_id:lane.assign,app_id:lane.app,spare_no:chosen});if(r.app_id!==lane.app||r.spare_no!==chosen)throw Error('INVALID_RESULT');lane.spare=chosen;save();}
  catch(e){if(lane.index>=2||e.message!=='SPARE_ALREADY_ASSIGNED')throw e;lane.conflict=true;save();}
 }));
 if(state.lanes.filter(x=>x.conflict).length!==1)throw Error('SPARE_CONFLICT_NOT_DETECTED');
 const loser=state.lanes.find(x=>x.conflict);
 if(!loser.spare){const spare=state.spares[4];const r=await call('assignSpare',{request_id:loser.fallback,app_id:loser.app,spare_no:spare});if(r.app_id!==loser.app||r.spare_no!==spare)throw Error('INVALID_RESULT');loser.spare=spare;save();}
 if(new Set(state.lanes.map(x=>x.spare)).size!==5)throw Error('DUPLICATE_SPARE');
 progress('3/4 5接続から入場を同時保存しています。');
 const rows=state.lanes.map(l=>({checkin_id:l.checkin,entry_code:l.spare,checked_in_at:state.createdAt,device_id:'cloud-lane-'+l.index}));
 const first=await Promise.all(rows.map(row=>call('scanCheckin',{checkins:[row]})));
 for(let i=0;i<5;i++)if(first[i].results?.[0]?.app_id!==state.lanes[i].app)throw Error('CHECKIN_RESOLUTION_FAILED');
 progress('4/4 同じ操作IDを5接続から再送し、重複記録にならないことを確認しています。');
 const replay=await Promise.all(rows.map(row=>call('scanCheckin',{checkins:[row]})));
 if(replay.some(r=>r.results?.[0]?.status!=='duplicate'))throw Error('CHECKIN_REPLAY_FAILED');
 state.done=true;state.summary={status:'PASS_HTTP_FIVE_CONNECTIONS',run:state.run,applications:state.lanes.map(l=>l.app),spares:state.lanes.map(l=>l.spare),checkin_ids:state.lanes.map(l=>l.checkin),spare_conflicts:1,replayed:5,finishedAt:now(),production_cutover:false};save();
 progress('5接続の登録・札競合・入場・同一ID再送が完了しました。実機5台の合格とは別です。');return state.summary;
}
