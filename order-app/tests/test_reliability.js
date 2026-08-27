'use strict';

const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const gas = fs.readFileSync(path.join(ROOT, 'gas', 'Code.gs'), 'utf8');

function section(text, startMarker, endMarker) {
  const start = text.indexOf(startMarker);
  const end = text.indexOf(endMarker, start);
  assert(start >= 0, `start marker not found: ${startMarker}`);
  assert(end > start, `end marker not found: ${endMarker}`);
  return text.slice(start, end);
}

function makeStorage() {
  const values = new Map();
  return {
    getItem: key => values.has(key) ? values.get(key) : null,
    setItem: (key, value) => values.set(key, String(value)),
    removeItem: key => values.delete(key),
    clear: () => values.clear()
  };
}

function makeTransportContext() {
  const listeners = {};
  const document = {
    visibilityState: 'visible',
    addEventListener: (name, fn) => { listeners[name] = fn; },
    getElementById: () => null
  };
  const window = { addEventListener: (name, fn) => { listeners[name] = fn; } };
  const context = vm.createContext({
    console, setTimeout, clearTimeout, AbortController, Date, Promise, Map, Set,
    document, window, location: { replace() {} }, localStorage: makeStorage(),
    fetch: null, listeners
  });
  const source = `
    const APP_VERSION='test';
    const LS_GAS_DIAG='diag';
    const SESSION_KEY='session';
    const PORTAL_URL='portal';
    const GAS_URL='gas';
    let __id=0;
    function _generateRequestId(){ return 'test-id-' + (++__id); }
    function getAuthSession(){ return {token:'token'}; }
    ${section(html, '// 通信1操作の絶対deadline。', 'async function fetchTemplateAsArrayBuffer')}
    globalThis.testApi={
      acquire:_gasAcquire,
      snapshot:_gasDebugSnapshot,
      createOperation:_gasCreateOperation,
      fetchOnce:_gasFetchOnce,
      gasPost,
      listeners
    };
  `;
  vm.runInContext(source, context);
  return context;
}

async function testSemaphoreCancellation() {
  const context = makeTransportContext();
  const api = context.testApi;
  const first = await api.acquire(1000);
  const second = await api.acquire(1000);
  const controllers = Array.from({ length: 100 }, () => new AbortController());
  const waiters = controllers.map(ac => api.acquire(1000, ac.signal));
  const settledPromise = Promise.allSettled(waiters);
  controllers.forEach(ac => ac.abort());
  const settled = await settledPromise;
  assert(settled.every(x => x.status === 'rejected'));
  first.release();
  second.release();
  assert.deepStrictEqual(JSON.parse(JSON.stringify(api.snapshot())), {
    active: 0, broken: false, waiting: 0, cancelled: 0, granted: 0
  });
  const lease1 = await api.acquire(1000);
  const lease2 = await api.acquire(1000);
  lease1.release(); lease1.release(); lease2.release();
  assert.strictEqual(api.snapshot().active, 0);

  // timeoutとreleaseの発火順が前後しても、1枠を二重に渡さない。
  for (const releaseDelay of [0, 1, 2]) {
    for (let i = 0; i < 30; i++) {
      const active1 = await api.acquire(100);
      const active2 = await api.acquire(100);
      const boundary = api.acquire(1).then(lease => ({ granted: true, lease }), () => ({ granted: false }));
      setTimeout(() => active1.release(), releaseDelay);
      const result = await boundary;
      if (result.granted) result.lease.release();
      active2.release();
      await new Promise(resolve => setTimeout(resolve, 3));
      const snapshot = api.snapshot();
      assert.strictEqual(snapshot.active, 0);
      assert.strictEqual(snapshot.broken, false);
    }
  }
}

async function testTransportPoliciesAndDeadline() {
  const context = makeTransportContext();
  const api = context.testApi;
  vm.runInContext(`_gasAbortableDelay = async function(ms, signal) {
    if (signal && signal.aborted) throw new GasDeadlineError();
  };`, context);

  let calls = 0;
  context.fetch = async () => { calls++; throw new TypeError('offline'); };
  await assert.rejects(
    api.gasPost({ action: 'saveSupplier' }, { totalDeadlineMs: 5000, attemptTimeoutMs: 1000 }),
    e => e.name === 'GasOutcomeUnknownError'
  );
  assert.strictEqual(calls, 1, 'WRITE_NO_RETRY must not retry a transport failure');

  calls = 0;
  context.fetch = async () => {
    calls++;
    if (calls < 3) throw new TypeError('temporary');
    return { ok: true, status: 200, text: async () => JSON.stringify({ success: true }) };
  };
  const readResult = await api.gasPost(
    { action: 'getOrderDetail' },
    { totalDeadlineMs: 5000, attemptTimeoutMs: 1000 }
  );
  assert.strictEqual(readResult.success, true);
  assert.strictEqual(calls, 3, 'READ_RETRYABLE must retry up to the shared cap');

  calls = 0;
  context.fetch = async () => {
    calls++;
    const payload = calls === 1 ? { success: false, error: 'AUTH_UNAVAILABLE' } : { success: true };
    return { ok: true, status: 200, text: async () => JSON.stringify(payload) };
  };
  const predispatch = await api.gasPost(
    { action: 'saveSupplier' },
    { totalDeadlineMs: 5000, attemptTimeoutMs: 1000 }
  );
  assert.strictEqual(predispatch.success, true);
  assert.strictEqual(calls, 2, 'pre-dispatch response may be retried for a write');

  calls = 0;
  context.fetch = () => { calls++; return new Promise(() => {}); };
  const started = Date.now();
  await assert.rejects(
    api.gasPost(
      { action: 'getOrderDetail' },
      { totalDeadlineMs: 40, attemptTimeoutMs: 100, maxAttempts: 1 }
    ),
    e => e.name === 'GasDeadlineError'
  );
  assert(Date.now() - started < 300, 'operation cancel must reject without waiting for fetch');
  assert.strictEqual(api.snapshot().active, 0, 'deadline must release the semaphore lease');
}

async function testOutboxSnapshotAndRecovery() {
  const context = vm.createContext({
    console, Date, Map, JSON, localStorage: makeStorage(),
    document: { getElementById: () => null },
    currentUser: { user_id: 'u1' },
    currentOrder: { date: '2026-08-12', supplierCode: '10', supplierName: 'Supplier', fax: '', staff: 'User' },
    cartItems: [{ code: 'A', name: 'Item', qty: 1 }],
    orderRequestId: null, revisionBaseOrderNo: null, savedOrderNo: null,
    allHistoryOrders: [], historyDetailCache: {}, sessionStorage: makeStorage(), SS_HIST_CACHE: 'hist-cache',
    // ⚠️ _completeOrderEntry() が参照する画面側のグローバルは全部ここに置くこと。
    //    1つでも欠けると ReferenceError が _recoverOrderEntry の catch に飲み込まれ、
    //    「recovery incomplete」という無関係な失敗に化ける（2026-08-21 v1.64.0で実際に発生）
    currentScreen: 'order', _historyRefreshing: false, loadHistory: async () => {}, proposalsLoaded: false,
    LS_ORDER_OUTBOX: 'outbox', ORDER_OUTBOX_LIMIT: undefined,
    _generateRequestId: () => 'request-0001',
    showToast() {}, copyToClipboard: async () => true,
    gasPost: async () => ({ success: true, state: 'COMPLETE', orderNo: '20260812-001' }),
    GasOutcomeUnknownError: class GasOutcomeUnknownError extends Error {}
  });
  const outboxSource = section(html, 'const ORDER_OUTBOX_LIMIT = 5;', 'async function saveOrderToGAS(outputType)');
  vm.runInContext(`${outboxSource}\n globalThis.testApi={create:_newOrderOutboxEntry,read:_readOrderOutbox,process:_processOrderEntry};`, context);
  const first = context.testApi.create('CSV');
  context.cartItems[0].qty = 99;
  const second = context.testApi.create('PDF');
  assert.strictEqual(first.requestId, second.requestId);
  assert.strictEqual(second.params.outputType, 'CSV', 'first immutable payload must win');
  assert.strictEqual(JSON.parse(second.params.items)[0].qty, 1, 'cart edits must not mutate an in-flight payload');
  assert.strictEqual(context.testApi.read().length, 1);

  const actions = [];
  let saveCalls = 0;
  context.gasPost = async params => {
    actions.push({ action: params.action, requestId: params.requestId });
    if (params.action === 'saveOrder' && saveCalls++ === 0) {
      const err = new Error('lost response'); err.name = 'GasOutcomeUnknownError'; throw err;
    }
    if (params.action === 'checkOrderByRequestId') return { success:true, state:'NOT_FOUND_NOW' };
    return { success:true, orderNo:'20260812-001' };
  };
  const recoveredOrderNo = await context.testApi.process(first, true);
  assert.strictEqual(recoveredOrderNo, '20260812-001');
  assert.deepStrictEqual(actions.map(x => x.action), ['saveOrder', 'checkOrderByRequestId', 'saveOrder']);
  assert(actions.every(x => x.requestId === 'request-0001'), 'recovery must keep the same requestId');
  assert.strictEqual(context.testApi.read().length, 0, 'COMPLETE removes the durable outbox entry');
}

/* ===================================================
   保存成立後の後片付けで落ちても、成功が失敗に化けないこと（v1.67.0）
   ここが崩れると「保存済みなのに未保存扱い→新しいrequestIdで再送→二重発注」になる
=================================================== */
function makeOutboxContext(overrides = {}) {
  let seq = 0;
  const context = vm.createContext(Object.assign({
    console: { log() {}, warn() {}, error() {} },
    Date, Map, JSON, Promise, localStorage: makeStorage(),
    document: { getElementById: () => null },
    currentUser: { user_id: 'u1' },
    currentOrder: { date: '2026-08-27', supplierCode: '10', supplierName: 'Supplier', fax: '', staff: 'User' },
    cartItems: [{ code: 'A', name: 'Item', qty: 1 }],
    orderRequestId: null, revisionBaseOrderNo: null, savedOrderNo: null,
    allHistoryOrders: [], historyDetailCache: {}, sessionStorage: makeStorage(), SS_HIST_CACHE: 'hist-cache',
    LS_ORDER_OUTBOX: 'outbox', ORDER_OUTBOX_LIMIT: undefined, LS_ORDERED: 'orderApp_orderedCodes',
    currentScreen: 'order', _historyRefreshing: false, loadHistory: async () => {}, proposalsLoaded: false,
    _generateRequestId: () => 'request-' + (++seq),
    showToast() {}, copyToClipboard: async () => true,
    gasPost: async () => ({ success: true, orderNo: '20260827-001' }),
    GasOutcomeUnknownError: class GasOutcomeUnknownError extends Error {}
  }, overrides));
  const src = section(html, 'const ORDER_OUTBOX_LIMIT = 5;', 'async function saveOrderToGAS(outputType)');
  vm.runInContext(src + 'globalThis.testApi={create:_newOrderOutboxEntry,read:_readOrderOutbox,process:_processOrderEntry};', context);
  return context;
}

async function testCompleteOrderEntryNeverTurnsSuccessIntoFailure() {
  // 後片付け（履歴キャッシュ破棄）が落ちる端末を再現する。
  // sessionStorageが使えない環境や容量枯渇で実際に起こりうる
  const brokenSession = makeStorage();
  brokenSession.removeItem = () => { throw new Error('storage disabled'); };
  const context = makeOutboxContext({ sessionStorage: brokenSession });

  const sentRequestIds = [];
  context.gasPost = async params => {
    sentRequestIds.push(params.requestId);
    return { success: true, orderNo: '20260827-001' };
  };

  const entry = context.testApi.create('CSV');
  const orderNo = await context.testApi.process(entry, true);

  assert.strictEqual(orderNo, '20260827-001', 'post-save cleanup failure must not reject a saved order');
  assert.strictEqual(context.savedOrderNo, '20260827-001',
    'savedOrderNo must be set before any fallible cleanup; saveOrderToGAS relies on it to skip re-sending');
  assert.strictEqual(context.testApi.read().length, 0, 'the durable outbox entry must still be removed');
  assert.deepStrictEqual(sentRequestIds, [entry.requestId], 'exactly one saveOrder must reach the server');

  // 後片付けが落ちても発注済みのローカル記録は残る（v1.66.0の「発注済みを隠す」用）
  const orderedLocal = JSON.parse(context.localStorage.getItem('orderApp_orderedCodes') || '[]');
  assert.strictEqual(orderedLocal.length, 1, 'ordered-local record runs before the failing cleanup step');

  // 履歴の自動再取得（待たない処理）が失敗しても、保存結果には影響しない
  const ctx2 = makeOutboxContext({ currentScreen: 'history', loadHistory: async () => { throw new Error('history down'); } });
  const entry2 = ctx2.testApi.create('CSV');
  assert.strictEqual(await ctx2.testApi.process(entry2, true), '20260827-001',
    'a rejected fire-and-forget history refresh must not affect the save result');
}

async function testTerminalServerErrorStillThrows() {
  // tryの範囲を通信だけに絞った（v1.67.0）あとも、終局エラーは今までどおり
  // 例外として伝わり、アウトボックスには要確認として残ること
  const context = makeOutboxContext();
  context.gasPost = async () => ({ success: false, error: 'REQUEST_ID_CONFLICT', message: 'conflict' });
  const entry = context.testApi.create('CSV');
  await assert.rejects(context.testApi.process(entry, true), e => e.code === 'REQUEST_ID_CONFLICT');
  const rows = context.testApi.read();
  assert.strictEqual(rows.length, 1, 'a terminal error must keep the entry for manual review');
  assert.strictEqual(rows[0].status, 'CONFLICT');
  assert.strictEqual(context.savedOrderNo, null, 'a failed save must not mark the order as saved');
}

function testServerCanonicalHashAndGuards() {
  const context = vm.createContext({
    JSON, Number, String,
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (_alg, text) => Array.from(crypto.createHash('sha256').update(text, 'utf8').digest())
    }
  });
  const helpers = section(gas, 'const ORDER_HISTORY_HEADERS', 'function deleteExactRows_');
  vm.runInContext(`${helpers}\n globalThis.testApi={normalize:normalizeOrderItem_,canonical:canonicalOrderPayload_,hash:sha256Hex_};`, context);
  const p1 = { date:'2026-08-12', supplierCode:'10', supplierName:'S', staff:'U', outputType:'CSV' };
  const p2 = { outputType:'CSV', staff:'U', supplierName:'S', supplierCode:'10', date:'2026-08-12' };
  const items1 = [context.testApi.normalize({ code:'A', name:'Item', qty:'2' })];
  const items2 = [context.testApi.normalize({ name:'Item', qty:2, code:'A' })];
  const h1 = context.testApi.hash(context.testApi.canonical(p1, 'u1', items1));
  const h2 = context.testApi.hash(context.testApi.canonical(p2, 'u1', items2));
  assert.strictEqual(h1, h2, 'property order and numeric representation must canonicalize');
  items2[0].qty = 3;
  const h3 = context.testApi.hash(context.testApi.canonical(p2, 'u1', items2));
  assert.notStrictEqual(h1, h3, 'business payload changes must change the hash');

  assert(gas.includes("error: 'REQUEST_ID_REQUIRED'"));
  assert(gas.includes("existingState === 'COMPLETE' && existingHash === requestHash"));
  assert(gas.includes('replaceAndVerifyOrderItems_(itemsSh, orderNo, items, now)'));
  assert(gas.includes("setNumberFormat('@')"), 'string columns must prevent Sheets numeric coercion');
  assert(gas.includes("state === '' || state === 'COMPLETE'"));
  assert(gas.includes("case 'checkOrderByRequestId'"));
}

function makeFakeSheet(initialRows) {
  const sheet = {
    rows: initialRows.map(r => r.slice()),
    getLastRow() { return this.rows.length; },
    appendRow(row) { this.rows.push(row.slice()); return this; },
    deleteRow(row) { this.rows.splice(row - 1, 1); },
    getRange(row, column, numRows = 1, numCols = 1) {
      const owner = this;
      const range = {
        getRow: () => row,
        getValue() { return this.getValues()[0][0]; },
        getValues() {
          const result = [];
          for (let r = 0; r < numRows; r++) {
            const source = owner.rows[row - 1 + r] || [];
            const values = [];
            for (let c = 0; c < numCols; c++) values.push(source[column - 1 + c] === undefined ? '' : source[column - 1 + c]);
            result.push(values);
          }
          return result;
        },
        setValue(value) { return this.setValues([[value]]); },
        setNumberFormat() { return this; },
        setValues(values) {
          for (let r = 0; r < values.length; r++) {
            while (owner.rows.length < row + r) owner.rows.push([]);
            const target = owner.rows[row - 1 + r];
            for (let c = 0; c < values[r].length; c++) target[column - 1 + c] = values[r][c];
          }
          return this;
        },
        createTextFinder(needle) {
          const finder = {
            matchEntireCell() { return this; }, matchCase() { return this; }, useRegularExpression() { return this; },
            findAll() {
              const found = [];
              for (let r = 0; r < numRows; r++) {
                for (let c = 0; c < numCols; c++) {
                  const value = (owner.rows[row - 1 + r] || [])[column - 1 + c];
                  if (String(value === undefined ? '' : value) === String(needle)) {
                    const foundRow = row + r;
                    found.push({ getRow: () => foundRow });
                  }
                }
              }
              return found;
            }
          };
          return finder;
        }
      };
      return range;
    }
  };
  return sheet;
}

function testServerIdempotencyStateMachine() {
  const historyHeaders = ['発注No','発注日','発注先コード','発注先名','FAX番号','担当者','品目数','出力方法','登録日時','user_id','requestId'];
  const itemHeaders = ['発注No','JANコード','Beaufieldコード','商品名','数量','単位','備考','手書きフラグ','登録日時'];
  const hist = makeFakeSheet([historyHeaders]);
  const itemsSheet = makeFakeSheet([itemHeaders]);
  let sequence = 0;
  const lock = { waitLock() {}, tryLock() { return true; }, releaseLock() {} };
  const context = vm.createContext({
    console, JSON, Number, String, Array,
    SHEET_HISTORY: 'history', SHEET_ITEMS: 'items',
    getSheet: name => name === 'history' ? hist : itemsSheet,
    generateOrderNo: date => date.replace(/-/g, '') + '-' + String(++sequence).padStart(3, '0'),
    LockService: { getScriptLock: () => lock },
    SpreadsheetApp: { flush() {} },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (_alg, text) => Array.from(crypto.createHash('sha256').update(text, 'utf8').digest()),
      // 実際の日付を実際のフォーマット文字列どおりに整形する（固定文字列を返す簡易モックだと
      // v1.34.0で追加した発注日の範囲チェック（'yyyy-MM-dd'）が壊れて誤判定するため）
      formatDate: (date, _tz, fmt) => {
        const pad = n => String(n).padStart(2, '0');
        const y = date.getFullYear(), mo = pad(date.getMonth() + 1), d = pad(date.getDate());
        const h = pad(date.getHours()), mi = pad(date.getMinutes()), s = pad(date.getSeconds());
        if (fmt === 'yyyy-MM-dd') return `${y}-${mo}-${d}`;
        if (fmt === 'yyyyMMdd') return `${y}${mo}${d}`;
        return `${y}-${mo}-${d} ${h}:${mi}:${s}`;
      }
    }
  });
  const stateMachine = section(gas, 'const ORDER_HISTORY_HEADERS', 'function deleteOrder(p, user_id)');
  vm.runInContext(`${stateMachine}\n globalThis.testApi={save:saveOrder,check:checkOrderByRequestId};`, context);
  // v1.34.0で追加した発注日の範囲チェック（当日±30日/+7日）に必ず収まるよう、
  // 固定日付ではなく実行時点の「当日」を使う（固定日付だと時間経過でテストが自然に失敗し出す）
  const _now = new Date();
  const todayStr = `${_now.getFullYear()}-${String(_now.getMonth() + 1).padStart(2, '0')}-${String(_now.getDate()).padStart(2, '0')}`;
  const base = {
    requestId: 'request-0001', date: todayStr, supplierCode: '10', supplierName: 'Supplier',
    staff: 'User', outputType: 'CSV',
    items: JSON.stringify([
      { janCode:'1', code:'A', name:'Item A', qty:2, unit:'本', memo:'', isHandwritten:false },
      { janCode:'2', code:'B', name:'Item B', qty:1, unit:'個', memo:'急ぎ', isHandwritten:false }
    ])
  };

  assert.strictEqual(context.testApi.save({ ...base, requestId: '' }, 'u1').error, 'REQUEST_ID_REQUIRED');
  const first = context.testApi.save(base, 'u1');
  assert.strictEqual(first.success, true);
  assert.strictEqual(hist.rows.length, 2);
  assert.strictEqual(hist.rows[1][13], 'COMPLETE');
  assert.strictEqual(itemsSheet.rows.length, 3);

  for (let i = 0; i < 10; i++) {
    const repeated = context.testApi.save(base, 'u1');
    assert.strictEqual(repeated.orderNo, first.orderNo);
  }
  assert.strictEqual(hist.rows.length, 2, 'same requestId must keep one history row');
  assert.strictEqual(itemsSheet.rows.length, 3, 'same requestId must keep one detail set');

  // 明細が揃っていてもPENDINGなら、重複させず同じ内容へ収束させる。
  hist.rows[1][13] = 'PENDING';
  let repaired = context.testApi.save(base, 'u1');
  assert.strictEqual(repaired.success, true);
  assert.strictEqual(hist.rows[1][13], 'COMPLETE');
  assert.strictEqual(itemsSheet.rows.length, 3);

  // 一部欠落・全件欠落のどちらも、存在確認だけで成功扱いせず全件を書き直す。
  hist.rows[1][13] = 'PENDING';
  itemsSheet.rows.splice(1, 1);
  repaired = context.testApi.save(base, 'u1');
  assert.strictEqual(repaired.success, true);
  assert.strictEqual(itemsSheet.rows.length, 3);
  hist.rows[1][13] = 'PENDING';
  itemsSheet.rows.splice(1, 2);
  repaired = context.testApi.save(base, 'u1');
  assert.strictEqual(repaired.success, true);
  assert.strictEqual(itemsSheet.rows.length, 3);

  const changed = { ...base, items: JSON.stringify([{ code:'A', name:'Item A', qty:3 }]) };
  assert.strictEqual(context.testApi.save(changed, 'u1').error, 'REQUEST_ID_CONFLICT');
  assert.strictEqual(context.testApi.check({ requestId: base.requestId }, 'u1').state, 'COMPLETE');
  assert.strictEqual(context.testApi.check({ requestId: 'request-9999' }, 'u1').state, 'NOT_FOUND_NOW');

  const revised = {
    ...base, requestId: 'request-0002', revisionBaseOrderNo: first.orderNo,
    items: JSON.stringify([{ code:'B', name:'Item B', qty:1 }])
  };
  // 新明細作成後・修正元削除前の障害を注入し、次回の同一requestIdで再開できることを確認。
  vm.runInContext(`
    const __deleteOrderUnlockedOriginal = deleteOrderUnlocked_;
    let __deleteFailureOnce = true;
    deleteOrderUnlocked_ = function(...args) {
      if (__deleteFailureOnce) { __deleteFailureOnce = false; throw new Error('injected before old delete'); }
      return __deleteOrderUnlockedOriginal(...args);
    };
  `, context);
  assert.throws(() => context.testApi.save(revised, 'u1'), /injected before old delete/);
  assert.strictEqual(hist.rows.length, 3, 'failed revision keeps old COMPLETE and new PENDING');
  const replacement = context.testApi.save(revised, 'u1');
  assert.strictEqual(replacement.success, true);
  assert.strictEqual(hist.rows.length, 2, 'revision must replace, not duplicate, history');
  assert.strictEqual(hist.rows[1][0], replacement.orderNo);
  assert.strictEqual(hist.rows[1][13], 'COMPLETE');
  assert.strictEqual(itemsSheet.rows.length, 2);
  assert.strictEqual(itemsSheet.rows[1][2], 'B');
}

(async () => {
  await testSemaphoreCancellation();
  await testTransportPoliciesAndDeadline();
  await testOutboxSnapshotAndRecovery();
  await testCompleteOrderEntryNeverTurnsSuccessIntoFailure();
  await testTerminalServerErrorStillThrows();
  testServerCanonicalHashAndGuards();
  testServerIdempotencyStateMachine();
  console.log('All reliability tests passed.');
})().catch(err => {
  console.error(err);
  process.exitCode = 1;
});
