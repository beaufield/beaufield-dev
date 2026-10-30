// カタログ監査専用の共有取得予算。商品・提案の保存や認可設定は変更しない。
const BMALL_CATALOG_WINDOW_MS_ = 300000;
const BMALL_CATALOG_REQUEST_BUDGET_ = 240; // 300回の上限に対し他処理用60回を残す。

function bridgeCatalogDetail_(path, offset, status, retryAt) {
  const now = Date.now();
  return {state: retryAt > now ? 'WAITING' : 'FAILED', path: path,
    offset: offset, httpStatus: status, retryAt: retryAt || null,
    retryAfterSeconds: retryAt > now ? Math.ceil((retryAt - now) / 1000) : 0};
}
function bridgeCatalogFailure_(code, detail) {
  // 応答本文・URL・ヘッダー・トークンを記録しない。
  try {console.log(JSON.stringify({event:'bmall_catalog_read',version:1,errorCode:'BRIDGE_'+code,...detail}));}catch(_){}
  const error = new Error('BRIDGE_' + code); error.catalogRead = detail; throw error;
}
function bridgeCatalogBudget_(token, count, blockUntil, path, offset, status) {
  // 自動化の保存処理はScriptLockを既に保持するため、独立したUserLockを使う。
  // Webアプリは所有者実行。同じ所有者の監査同士で予算を共有する。
  const lock = LockService.getUserLock();
  if (!lock.tryLock(1000)) bridgeCatalogFailure_('CATALOG_BUDGET_BUSY',bridgeCatalogDetail_(path,offset,status,Date.now()+2000));
  try {
    const key = 'BMALL_CATALOG_BUDGET_V1_' + bridgeHash_(token).slice(0,16);
    const props = PropertiesService.getScriptProperties(), raw = props.getProperty(key), now = Date.now();
    let state = {version:1,requests:[],retryAt:0};
    if (raw) {
      try { state = JSON.parse(raw); } catch (_) { state = null; }
      if (!state || state.version!==1 || !Array.isArray(state.requests) || state.requests.length>240 ||
          state.requests.some(t=>!Number.isSafeInteger(t)||t<0||t>now) ||
          !Number.isSafeInteger(state.retryAt) || state.retryAt<0) {
        bridgeCatalogFailure_('CATALOG_BUDGET_INVALID',bridgeCatalogDetail_(path,offset,status,0));
      }
    }
    state.requests = state.requests.filter(t=>t>now-BMALL_CATALOG_WINDOW_MS_);
    if (blockUntil) {
      state.retryAt = Math.max(state.retryAt,blockUntil);
      props.setProperty(key,JSON.stringify(state));
      return state.retryAt;
    }
    if (state.retryAt>now) bridgeCatalogFailure_('CATALOG_WAITING',bridgeCatalogDetail_(path,offset,status,state.retryAt));
    if (!Number.isSafeInteger(count)||count<1||count>10) bridgeFail_('CATALOG_BUDGET_INVALID');
    if (state.requests.length+count>BMALL_CATALOG_REQUEST_BUDGET_) {
      const retryAt=state.requests[state.requests.length+count-BMALL_CATALOG_REQUEST_BUDGET_-1]+BMALL_CATALOG_WINDOW_MS_+1000;
      bridgeCatalogFailure_('CATALOG_WAITING',bridgeCatalogDetail_(path,offset,status,retryAt));
    }
    for(let i=0;i<count;i++)state.requests.push(now);
    props.setProperty(key,JSON.stringify(state));
  } finally { lock.releaseLock(); }
}
function bridgeCatalogRetryAfter_(response) {
  let headers;try {headers=response.getAllHeaders();}catch(_){return null;}
  const key=Object.keys(headers||{}).find(k=>k.toLowerCase()==='retry-after');
  if(!key)return null;
  const raw=String(headers[key]).trim();
  const seconds=/^\d+$/.test(raw)?Number(raw):Math.ceil((Date.parse(raw)-Date.now())/1000);
  return Number.isSafeInteger(seconds)&&seconds>=0&&seconds<=86400?Math.max(1,seconds):null;
}
function bridgeCatalogFetch_(requests, token, path, offsets) {
  bridgeCatalogBudget_(token,requests.length,0,path,offsets[0],null);
  let responses;
  try {
    if(requests.length===1)responses=[apiFetch_(requests[0].url,requests[0])];
    // 全ページのHTTP応答を同じ予算内で取得。部分データ成功は返さない。
    else responses=UrlFetchApp.fetchAll(requests);
  } catch (_) {
    bridgeCatalogFailure_('CATALOG_NETWORK_FAILED',bridgeCatalogDetail_(path,offsets[0],null,0));
  }
  // 同じ並列バッチに429があれば、先行ページの502等を再試行する前に待機する。
  const limited=responses.findIndex(r=>r.getResponseCode()===429);
  if(limited>=0)bridgeCatalogResponse_(responses[limited],requests[limited],token,path,offsets[limited]);
  return responses;
}
function bridgeCatalogResponse_(response, request, token, path, offset) {
  for(let attempt=0;attempt<3;attempt++) {
    const status=response.getResponseCode();
    if(status===200)return response;
    const retryAfter=bridgeCatalogRetryAfter_(response);
    const detail=bridgeCatalogDetail_(path,offset,status,0);
    if(status===429) {
      const retryAt=bridgeCatalogBudget_(token,0,Date.now()+(retryAfter||300)*1000,path,offset,status);
      bridgeCatalogFailure_('CATALOG_RATE_LIMITED',bridgeCatalogDetail_(path,offset,status,retryAt));
    }
    if(![502,503,504].includes(status)||attempt===2)bridgeCatalogFailure_('CATALOG_HTTP_FAILED',detail);
    // 長いRetry-Afterは実行中に待たず、共有待機として返す。
    if(retryAfter&&retryAfter>5) {
      const retryAt=bridgeCatalogBudget_(token,0,Date.now()+retryAfter*1000,path,offset,status);
      bridgeCatalogFailure_('CATALOG_WAITING',bridgeCatalogDetail_(path,offset,status,retryAt));
    }
    console.log(JSON.stringify({event:'bmall_catalog_read',version:1,outcome:'retry',attempt:attempt+1,...detail}));
    Utilities.sleep(Math.max(attempt===0?1000:3000,(retryAfter||0)*1000));
    response=bridgeCatalogFetch_([request],token,path,[offset])[0];
  }
}
