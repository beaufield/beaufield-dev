// 日本語コメント: 集計は読取結果だけを使う。ブラウザーのタイムゾーンに依存せず日本時間へ統一する。
export const EVENT_DATE='2026-10-26';
export function japanTime(value){
 const text=String(value||'').trim();
 if(!text)return null;
 const match=/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2})(\.\d{1,9})?)?(Z|[+-]\d{2}:?\d{2})?$/.exec(text);
 if(!match)return null;
 const [,y,mo,d,h,mi,s='0',fraction='',zone='']=match;
 const probe=new Date(Date.UTC(+y,+mo-1,+d,+h,+mi,+s));
 if(probe.getUTCFullYear()!==+y||probe.getUTCMonth()!==+mo-1||probe.getUTCDate()!==+d||+h>23||+mi>59||+s>59)return null;
 const normalized=`${y}-${mo.padStart(2,'0')}-${d.padStart(2,'0')}T${h.padStart(2,'0')}:${mi}:${s.padStart(2,'0')}${fraction}${zone||'+09:00'}`;
 const epoch=Date.parse(normalized);if(!Number.isFinite(epoch))return null;
 const j=new Date(epoch+9*3600000);
 return {epoch,date:j.toISOString().slice(0,10),minute:j.getUTCHours()*60+j.getUTCMinutes(),label:j.toISOString().slice(11,19)};
}
export function snapshot(data){
 if(!data||data.schema_version!==1||!Array.isArray(data.applications)||data.booking?.state!=='ready'||!Array.isArray(data.booking.sessions)||!data.booking.by_app||typeof data.booking.by_app!=='object')throw Error('DATA_SCHEMA');
 const seen=new Set(),sessions=new Map(data.booking.sessions.map(s=>[s.session_id,s]));
 const rows=[];
 for(const a of data.applications){
  if(typeof a.app_id!=='string'||!a.app_id||seen.has(a.app_id))throw Error('DATA_SCHEMA');
  seen.add(a.app_id);if(a.status!=='confirmed')continue;
  const raw=data.booking.by_app[a.app_id]?.reserved;
  if(!Array.isArray(raw))throw Error('DATA_SCHEMA');
  const reserved=[...new Set(raw)].map(id=>{const s=sessions.get(id);if(!s)throw Error('DATA_SCHEMA');return s;});
  const checked=String(a.checked_in_at||'').trim();
  rows.push({...a,arrived:!!checked,time:japanTime(checked),reserved});
 }
 return {rows,tantou:Array.isArray(data.tantou_list)?data.tantou_list:[],sessions:[...sessions.values()],capturedAt:data.booking.captured_at||''};
}
export function selectBase(rows,{query='',tantou='',booking='all',session=''}={}){
 const normalize=v=>String(v||'').normalize('NFKC').toLocaleLowerCase('ja');
 const terms=normalize(query).trim().split(/\s+/).filter(Boolean);
 return rows.filter(a=>{
  if(tantou==='__empty__'?!!a.tantou:tantou&&a.tantou!==tantou)return false;
  if(booking==='yes'&&!a.reserved.length||booking==='no'&&a.reserved.length)return false;
  if(session&&!a.reserved.some(s=>s.session_id===session))return false;
  const hay=normalize([a.salon_name,a.staff_name,a.app_id].join(' '));return terms.every(q=>hay.includes(q));
 });
}
export function totals(rows){const arrived=rows.filter(a=>a.arrived).length;return {total:rows.length,arrived,waiting:rows.length-arrived};}
export function histogram(rows,date,interval=10){
 if(![10,60].includes(interval)||!/^\d{4}-\d{2}-\d{2}$/.test(date))throw Error('INVALID_CHART');
 const arrived=rows.filter(a=>a.arrived),dated=arrived.filter(a=>a.time?.date===date);
 const start=Math.min(600,...dated.map(a=>Math.floor(a.time.minute/interval)*interval));
 const end=Math.max(960,...dated.map(a=>(Math.floor(a.time.minute/interval)+1)*interval));
 const bins=[];for(let minute=start;minute<end;minute+=interval)bins.push({minute,count:0});
 for(const a of dated)bins[Math.floor((a.time.minute-start)/interval)].count++;
 // 日本語コメント: 各時間帯の終わりまでの累計。時刻不明・別日を混ぜず、棒と同じ母集団を使う。
 let cumulative=0;for(const bin of bins)bin.cumulative=cumulative+=bin.count;
 return {bins,count:dated.length,unknown:arrived.filter(a=>!a.time).length,otherDate:arrived.filter(a=>a.time&&a.time.date!==date).length,max:Math.max(1,...bins.map(b=>b.count))};
}
export function selectList(rows,{status='arrived',date='',minute=null,interval=10,sort='recent'}={}){
 const selected=rows.filter(a=>{
  if(status==='arrived'&&!a.arrived||status==='waiting'&&a.arrived)return false;
  if(minute!==null&&(!a.arrived||!a.time||a.time.date!==date||a.time.minute<minute||a.time.minute>=minute+interval))return false;
  return true;
 });
 const cmp=(a,b)=>String(a||'').localeCompare(String(b||''),'ja');
 selected.sort((a,b)=>{
  if(sort==='name')return cmp(a.staff_name,b.staff_name)||cmp(a.app_id,b.app_id);
  if(sort==='salon')return cmp(a.salon_name,b.salon_name)||cmp(a.staff_name,b.staff_name)||cmp(a.app_id,b.app_id);
  return (sort==='oldest'?1:-1)*((a.time?.epoch??0)-(b.time?.epoch??0))||cmp(a.app_id,b.app_id);
 });return selected;
}
export const minuteLabel=minute=>`${String(Math.floor(minute/60)).padStart(2,'0')}:${String(minute%60).padStart(2,'0')}`;
// 日本語コメント: 予約枠は時刻だけの値（10:00）を返す。本番APIの形式とISO日時の両方を表示する。
export function sessionClock(value){
 const raw=String(value||'');
 if(/^(?:[01]\d|2[0-3]):[0-5]\d(?::[0-5]\d)?$/.test(raw))return raw.slice(0,5);
 return japanTime(raw)?.label.slice(0,5)||'';
}
