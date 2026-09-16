'use strict';

const assert = require('assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const parent = path.dirname(root);
const orderGas = fs.readFileSync(path.join(root, 'gas/Code.gs'), 'utf8');
const orderHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const bcartGas = fs.readFileSync(path.join(parent, 'bcart-integration/master-tool/gas/Code.gs'), 'utf8');
const bcartHtml = fs.readFileSync(path.join(parent, 'bcart-integration/master-tool/index.html'), 'utf8');
const dashboardGas = fs.readFileSync(path.join(parent, 'project-dashboard/gas/Code.gs'), 'utf8');
const dashboardHtml = fs.readFileSync(path.join(parent, 'project-dashboard/index.html'), 'utf8');

function functionSource(source, name) {
  const start = source.indexOf('function ' + name + '(');
  assert(start >= 0, 'function not found: ' + name);
  const brace = source.indexOf('{', start);
  let depth = 0, quote = '', escaped = false, lineComment = false, blockComment = false;
  for (let i = brace; i < source.length; i++) {
    const c = source[i], n = source[i + 1];
    if (lineComment) { if (c === '\n') lineComment = false; continue; }
    if (blockComment) { if (c === '*' && n === '/') { blockComment = false; i++; } continue; }
    if (quote) {
      if (escaped) escaped = false;
      else if (c === '\\') escaped = true;
      else if (c === quote) quote = '';
      continue;
    }
    if (c === '/' && n === '/') { lineComment = true; i++; continue; }
    if (c === '/' && n === '*') { blockComment = true; i++; continue; }
    if (c === "'" || c === '"' || c === '`') { quote = c; continue; }
    if (c === '{') depth++;
    if (c === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error('unterminated function: ' + name);
}

function makeSheet(rows) {
  return {
    rows,
    getLastRow() { return this.rows.length; },
    getRange(row, col, height, width) {
      return { getValues: () => this.rows.slice(row - 1, row - 1 + height).map(r => r.slice(col - 1, col - 1 + width)) };
    }
  };
}

function testOrderReadsAndNumbering() {
  const target = '20260917-001';
  const detailRows = [['発注No']];
  for (let i = 0; i < 7; i++) detailRows.push([target, 'jan'+i, 'c'+i, 'p'+i, 1, '本', '', 'FALSE']);
  for (let i = 0; i < 700; i++) detailRows.push(['20260916-'+String(i).padStart(3,'0'), 'x', 'x', 'x', 1, '本', '', 'FALSE']);
  for (let i = 7; i < 15; i++) detailRows.push([target, 'jan'+i, 'c'+i, 'p'+i, 1, '本', '', 'FALSE']);
  const history = [['header'], [target, '2026-09-17', '1', 'supplier', '', '', 15, '', '', '', '', '', '', 'COMPLETE']];
  let released = false;
  const sheets = {items: makeSheet(detailRows), history: makeSheet(history)};
  const ctx = vm.createContext({
    Set, String, Number,
    SHEET_ITEMS:'items', SHEET_HISTORY:'history', getSheet:n=>sheets[n],
    LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{released=true;}})},
    findExactRows_:(sh,col,value)=>sh.rows.map((r,i)=>String(r[col-1]||'')===value?i+1:0).filter(Boolean)
  });
  vm.runInContext(functionSource(orderGas,'readRowsForValues_')+'\n'+functionSource(orderGas,'getOrderDetail'),ctx);
  const result = ctx.getOrderDetail(target);
  assert.equal(result.success,true); assert.equal(result.items.length,15); assert(released);
  assert.equal(ctx.readRowsForValues_(makeSheet([['header']]),new Set([target]),6).length,0);
  assert.equal(ctx.readRowsForValues_(makeSheet([['header'],[target,'j','c','n',1,'本']]),new Set([target]),6).length,1);

  const numberRows = [['発注No'],['20260917-001'],...Array.from({length:300},(_,i)=>['20260916-'+String(i+1).padStart(3,'0')]),['20260917-999']];
  ctx.getSheet = ()=>makeSheet(numberRows);
  ctx.SHEET_HISTORY='history';
  vm.runInContext(functionSource(orderGas,'generateOrderNo'),ctx);
  assert.equal(ctx.generateOrderNo('2026-09-17'),'20260917-1000');

  for (const name of ['getOrders','buildPendingOrders','generateOrderNo']) {
    const body = functionSource(orderGas,name);
    assert(!body.includes('readRowsForKey_('), name+' must not use ordered key tail reads');
  }
  assert(!functionSource(orderGas,'buildPendingOrders').includes('readTailRowsUntil_('));
  assert(orderHtml.includes("const LS_PROD_HIST   = 'bf_prod_hist_v2'"));
  assert(orderHtml.includes("const SS_HIST_CACHE = 'bf_hist_sess3'"));
}

function testPendingOrders() {
  const pad=n=>String(n).padStart(2,'0');
  const now=new Date();
  const dateKey=''+now.getFullYear()+pad(now.getMonth()+1)+pad(now.getDate());
  const orders=[dateKey+'-001',dateKey+'-002',dateKey+'-003',dateKey+'-004'];
  const itemRows=[['order','jan','code','name','qty'],
    [orders[0],'j1','C1','one',2],['19990101-001','old','OLD','old',1],
    [orders[2],'j3','C3','pending',1],[orders[1],'j2','C2','received',1],[orders[3],'j4','C4','legacy',3]];
  const histRows=[['h'],[orders[2],'','S','Supplier','','',1,'','','','','','','PENDING'],
    [orders[0],'','S','Supplier','','',1,'','','','','','','COMPLETE'],
    [orders[3],'','S','Supplier','','',1,'','','','','','',''],
    [orders[1],'','S','Supplier','','',1,'','','','','','','COMPLETE']];
  function sheet(rows) {
    return {rows,getLastRow(){return rows.length;},getLastColumn(){return rows[0].length;},
      getDataRange(){return{getValues:()=>rows};},getRange(r,c,h,w){return{getValues:()=>rows.slice(r-1,r-1+h).map(x=>x.slice(c-1,c-1+w))};}};
  }
  const sheets={items:sheet(itemRows),history:sheet(histRows),suppliers:sheet([['code','','','','','lead'],['S','','','','',1]]),products:null};
  const ctx=vm.createContext({Set,Object,Math,Number,String,Date,
    SHEET_ITEMS:'items',SHEET_HISTORY:'history',SHEET_SUPPLIERS:'suppliers',SHEET_PRODUCTS:'products',
    DEFAULT_LEAD_TIME_DAYS:2,DEFAULT_POSTING_LAG_DAYS:3,PENDING_GRACE_DAYS:30,
    getSS:()=>({getSheetByName:n=>sheets[n]}),getSheet:n=>sheets[n],
    getReceivedKeys:()=>new Set([orders[1]+'|C2']),getPostingLagByCode:()=>({S:0}),
    findColIdxGAS:()=>-1,
    Utilities:{formatDate(d,tz,fmt){const y=d.getFullYear(),m=pad(d.getMonth()+1),day=pad(d.getDate());return fmt==='yyyyMMdd'?''+y+m+day:y+'-'+m+'-'+day;}}
  });
  vm.runInContext([functionSource(orderGas,'readRowsForValues_'),functionSource(orderGas,'ymdToEpochDay'),functionSource(orderGas,'epochDayToDateStr'),functionSource(orderGas,'buildPendingOrders')].join('\n'),ctx);
  const result=ctx.buildPendingOrders();
  assert.deepEqual(Array.from(result,x=>x.code).sort(),['C1','C4']);
}

async function testBcartTransportAndApproval() {
  let fetches=0, timers=0, callbacks=0, busyEnds=0, releases=0;
  const transport = vm.createContext({
    Set, console, session:{}, GAS_URL:'gas', _lastClickedBtn:null, _lastClickedAt:0,
    RETRYABLE_READ_ACTIONS:new Set(['getDrafts']), RESPONSE_ARRAY_KEYS:{getDrafts:'drafts'},
    netBusyStart(){},netBusyEnd(){busyEnds++;},releaseBtnBusy(){releases++;},markBtnBusy(){},showToast(){},recordClientDiagnostic(){},
    fetch:async()=>{fetches++;return{json:async()=>({ok:true})};},setTimeout(){timers++;}
  });
  vm.runInContext(functionSource(bcartHtml,'validateBusinessResponse')+'\n'+functionSource(bcartHtml,'makeDiagnosticId')+'\n'+functionSource(bcartHtml,'gasPost'),transport);
  transport.gasPost({action:'approveDraft'},()=>{callbacks++;throw Error('render');});
  await new Promise(resolve=>setImmediate(resolve)); await new Promise(resolve=>setImmediate(resolve));
  assert.equal(fetches,1); assert.equal(callbacks,1); assert.equal(timers,0);
  assert.equal(busyEnds,1); assert.equal(releases,1);
  assert.equal(transport.validateBusinessResponse('getDrafts',{ok:true,version:'x'}).error,'BAD_RESPONSE');
  assert.equal(transport.validateBusinessResponse('loadData',{ok:true,diffs:[]}).error,'BAD_RESPONSE');
  assert.equal(transport.validateBusinessResponse('loadData',{
    ok:true,diffs:[],supplierMap:{},totalCsv:0,totalBcart:0,isOld:false
  }).ok,true);

  const elements={diffList:{innerHTML:'previous'},'cnt-price':{textContent:'1'},'cnt-disc':{textContent:'2'},
    'cnt-unreg':{textContent:'3'},'cnt-ignore':{textContent:'4'},csvWarn:{style:{},textContent:''}};
  let summaryUpdates=0,renders=0;
  const listUi=vm.createContext({allDiffs:[{code:'old'}],Array,
    document:{getElementById:id=>elements[id]},setTimeout:()=>1,clearTimeout(){},
    gasPost:(p,cb)=>cb({ok:false,error:'BAD_RESPONSE'}),showToast(){},
    updateSummary(){summaryUpdates++;},renderDiffs(){renders++;}});
  vm.runInContext(functionSource(bcartHtml,'loadData'),listUi);
  listUi.loadData();
  assert.equal(elements.diffList.innerHTML,'previous');
  assert.equal(summaryUpdates,1); assert.equal(renders,1);

  const actions=[];
  const approval = vm.createContext({draftList:[{draftId:'d',draftType:'new'}],
    collectDraftFormPayload:()=>({productName:'p',categoryId:'c'}),confirm:()=>true,
    showToast(){},showBusyOverlay(){},hideBusyOverlay(){},loadDrafts(){},
    gasPostAsync:async p=>{actions.push(p.action);return{ok:false,error:'REJECTED'};}});
  vm.runInContext(functionSource(bcartHtml,'approveDraftOne'),approval);
  approval.approveDraftOne('d',{disabled:false,textContent:''});
  await new Promise(resolve=>setImmediate(resolve)); await new Promise(resolve=>setImmediate(resolve));
  assert.deepEqual(actions,['updateDraft']);

  let reloads=0;
  const partialButton={disabled:false,textContent:''};
  const partial=vm.createContext({draftList:[{draftId:'d',draftType:'new'}],
    collectDraftFormPayload:()=>({productName:'p',categoryId:'c'}),confirm:()=>true,
    showToast(){},showBusyOverlay(){},hideBusyOverlay(){},loadDrafts(){reloads++;},
    gasPostAsync:async p=>p.action==='updateDraft'?{ok:true}:{ok:true,partial:true,message:'一部成功'}});
  vm.runInContext(functionSource(bcartHtml,'approveDraftOne'),partial);
  partial.approveDraftOne('d',partialButton);
  await new Promise(resolve=>setImmediate(resolve)); await new Promise(resolve=>setImmediate(resolve));
  assert.equal(partialButton.disabled,true); assert.equal(partialButton.textContent,'結果を確認してください'); assert.equal(reloads,0);

  const draftElements={draftStatusFilter:{value:'下書き'},draftSupplierFilter:{value:'A'},
    draftList:{innerHTML:'previous drafts'},draftCount:{textContent:'1件'}};
  const draftUi=vm.createContext({draftLoadGeneration:0,draftDisplayedKey:'下書き|A',draftList:[{draftId:'old'}],
    document:{getElementById:id=>draftElements[id]},ensureDraftMasters:cb=>cb({ok:true}),
    gasPostAsync:async()=>({ok:false,error:'BAD_RESPONSE'}),showToast(){},esc:String,
    rebuildDraftSupplierFilter(){},renderDraftList(){}});
  vm.runInContext(functionSource(bcartHtml,'loadDrafts'),draftUi);
  draftUi.loadDrafts();
  await new Promise(resolve=>setImmediate(resolve)); await new Promise(resolve=>setImmediate(resolve));
  assert.equal(draftElements.draftList.innerHTML,'previous drafts');

  let failedFetches=0, failureCallbacks=0;
  const failed=vm.createContext({Set,console,session:{},GAS_URL:'gas',_lastClickedBtn:null,_lastClickedAt:0,
    RETRYABLE_READ_ACTIONS:new Set(['getDrafts']),RESPONSE_ARRAY_KEYS:{},
    netBusyStart(){},netBusyEnd(){},releaseBtnBusy(){},markBtnBusy(){},showToast(){},recordClientDiagnostic(){},
    fetch:async()=>{failedFetches++;throw Error('lost');},setTimeout(){throw Error('write must not retry');}});
  vm.runInContext(functionSource(bcartHtml,'validateBusinessResponse')+'\n'+functionSource(bcartHtml,'makeDiagnosticId')+'\n'+functionSource(bcartHtml,'gasPost'),failed);
  failed.gasPost({action:'approveDraft'},res=>{failureCallbacks++;assert.equal(res.error,'RESULT_UNKNOWN');});
  await new Promise(resolve=>setImmediate(resolve)); await new Promise(resolve=>setImmediate(resolve));
  assert.equal(failedFetches,1); assert.equal(failureCallbacks,1);
}

function testBcartDataAndAuth() {
  let creates=0,sets=0;
  const base = {PropertiesService:{getScriptProperties:()=>({getProperty:()=> 'configured',setProperty(){sets++;}})},
    SpreadsheetApp:{openById(){throw Error('temporary');},create(){creates++;}},
    SHEET_IGNORE:'ignore'};
  const ctx=vm.createContext(base);
  vm.runInContext(functionSource(bcartGas,'getDataSpreadsheet_')+'\n'+functionSource(bcartGas,'requiredSheetHeaders_')+'\n'+functionSource(bcartGas,'getOrCreateSheet'),ctx);
  assert.throws(()=>ctx.getOrCreateSheet('ignore'),/DATA_UNAVAILABLE/);
  assert.equal(creates,0);assert.equal(sets,0);

  function authResponse(body,status=200) {
    const c=vm.createContext({AUTH_GAS_URL:'auth',UrlFetchApp:{fetch:()=>({getResponseCode:()=>status,getContentText:()=>JSON.stringify(body)})}});
    vm.runInContext(functionSource(bcartGas,'validateSession'),c);
    return c.validateSession({token:'t'});
  }
  assert.equal(authResponse({success:false,error:'INTERNAL_ERROR'}).error,'AUTH_UNAVAILABLE');
  assert.equal(authResponse({ok:true}).error,'AUTH_UNAVAILABLE');
  assert.equal(authResponse({ok:false}).error,'UNAUTHORIZED');
  assert.equal(authResponse({ok:true,user_id:'u',role:'editor',is_admin:false}).ok,true);

  const specials=vm.createContext({Set,bcartGetAll:ep=>({ok:false,error:'API_DOWN'}),getFeatureTypeMap_:()=>({})});
  vm.runInContext(functionSource(bcartGas,'getSpecials'),specials);
  assert.equal(specials.getSpecials().error,'API_DOWN');
  specials.bcartGetAll=()=>({ok:true,data:[{id:1,name:'feature'}]});
  specials.getFeatureTypeMap_=()=>{throw Error('DATA_UNAVAILABLE');};
  assert.throws(()=>specials.getSpecials(),/DATA_UNAVAILABLE/);

  const normalHelper=functionSource(bcartGas,'getOrCreateSheet');
  assert(!normalHelper.includes('insertSheet('));
  assert(!normalHelper.includes('SpreadsheetApp.create('));
  assert(!normalHelper.includes('setProperty('));
}

function testBcartMigrationIsRepeatable() {
  const headers=['group_id','group_name','member_ids','created_at','note','use_view_filter'];
  const dataRow=['g1','Group','1,2','date','note',true];
  const sheet={getLastColumn:()=>headers.length,
    getRange(r,c,h,w){const api={
      getValues:()=>[Array.from({length:w},(_,i)=>headers[c-1+i] || '')],
      setValue:v=>{headers[c-1]=v;return api;},setFontWeight:()=>api,setBackground:()=>api};return api;}};
  const names=['SHEET_IGNORE','SHEET_WIP','SHEET_DESC_SKIP','SHEET_HISTORY','SHEET_SP_GROUPS','SHEET_SP_INDIVIDUAL','SHEET_VF_DETAILS','SHEET_SP_DETAILS','SHEET_FEATURES','SHEET_DRAFT','SHEET_DRAFT_SETS','SHEET_STOCK_LOG','SHEET_STOCK_LOG_DETAIL','SHEET_SP_AUDIT_EXCLUDE'];
  const base={getDataSpreadsheet_:()=>({getSheetByName:()=>sheet}),Date};
  names.forEach(n=>{base[n]=n;});
  const ctx=vm.createContext(base);
  vm.runInContext(functionSource(bcartGas,'requiredSheetHeaders_')+'\n'+functionSource(bcartGas,'setupOrMigrateSheet_'),ctx);
  ctx.setupOrMigrateSheet_('SHEET_SP_GROUPS');
  ctx.setupOrMigrateSheet_('SHEET_SP_GROUPS');
  assert.equal(headers[6],'customer_codes');
  assert.deepEqual(dataRow,['g1','Group','1,2','date','note',true]);
}

function testDashboardTransientRecovery() {
  let unavailable=true,opens=0,puts=[];
  const cacheValues=new Map();
  const sheets={sessions:[['t','u','e'],['token','user',Date.now()+100000]],users:[[],['user','name','',true,'',true]],user_app_roles:[[],['user','project-dashboard','admin']]};
  const ctx=vm.createContext({APP_NAME:'project-dashboard',CACHE_TTL_SESSION:60,Date,
    prop_:()=> 'auth',Logger:{log(){}},
    CacheService:{getScriptCache:()=>({get:k=>cacheValues.get(k)||null,put:(k,v)=>{cacheValues.set(k,v);puts.push(JSON.parse(v));}})},
    SpreadsheetApp:{openById(){opens++;if(unavailable)throw Error('temporary');return{getSheetByName:n=>({getDataRange:()=>({getValues:()=>sheets[n]}),deleteRow(){}})};}}});
  vm.runInContext(functionSource(dashboardGas,'validateSession'),ctx);
  assert.equal(ctx.validateSession('token').transient,true);
  unavailable=false;
  assert.equal(ctx.validateSession('token').valid,true);
  assert.equal(opens,2);assert(!puts.some(x=>x.transient));
  assert(dashboardHtml.includes("Array.isArray(data.projects)"));

  sheets.sessions=[['t','u','e'],['expired','user',Date.now()-1000]];
  assert.equal(ctx.validateSession('expired').valid,false);
  assert(puts.some(x=>x.valid===false));
}

(async()=>{
  testOrderReadsAndNumbering();
  testPendingOrders();
  await testBcartTransportAndApproval();
  testBcartDataAndAuth();
  testBcartMigrationIsRepeatable();
  testDashboardTransientRecovery();
  console.log('R1 stability tests passed');
})().catch(e=>{console.error(e);process.exitCode=1;});
