'use strict';
// order-app v1.62.0 の3機能（発注日変更・提案タブの発注済み除外・系列まるごと見送り）の単体テスト。
// 既存 tests/test_reliability.js と同じ方式（index.htmlからマーカー区間を切り出しvmで実行）を使う。
// 実行: node tests/test_order_features.js（test_reliability.js と両方PASSすること）
// 設計原本: 発注日変更・発注済み除外_設計プラン.md §6-1

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const gas  = fs.readFileSync(path.join(ROOT, 'gas', 'Code.gs'), 'utf8');

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

/* ===================================================
   機能1: 発注日変更（isOrderDateAllowed / resolveOrderDate 等）
=================================================== */
function testOrderDateHelpers() {
  const inputDateEl = { value: '' };
  const context = vm.createContext({
    console, String, Number, Date,
    document: { getElementById: id => (id === 'input-date' ? inputDateEl : null) },
    // 抽出範囲の外にある getTodayStr() を最小実装で提供（実装と同じロジック）
    getTodayStr: () => {
      const d = new Date();
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }
  });
  const block = section(html,
    '/* === ORDER DATE HELPERS (test:order-date) === */',
    '/* === /ORDER DATE HELPERS === */');
  vm.runInContext(`${block}\n globalThis.testApi = {
    isOrderDateAllowed, orderDateMin, orderDateMax, shiftDateStr, resolveOrderDate
  };`, context);
  const api = context.testApi;

  const today = context.getTodayStr();
  const min = api.orderDateMin();
  const max = api.orderDateMax();

  assert.strictEqual(api.isOrderDateAllowed(today), true, 'today must be allowed');
  assert.strictEqual(api.isOrderDateAllowed(min), true, 'lower bound (30 days ago) must be allowed');
  assert.strictEqual(api.isOrderDateAllowed(max), true, 'upper bound (7 days ahead) must be allowed');

  const beforeMin = api.shiftDateStr(min, -1);
  const afterMax  = api.shiftDateStr(max, 1);
  assert.strictEqual(api.isOrderDateAllowed(beforeMin), false, '31 days ago must be rejected');
  assert.strictEqual(api.isOrderDateAllowed(afterMax), false, '8 days ahead must be rejected');
  assert.strictEqual(api.isOrderDateAllowed('2026-8-1'), false, 'malformed date (no zero-padding) must be rejected');
  assert.strictEqual(api.isOrderDateAllowed(''), false, 'empty string must be rejected');
  assert.strictEqual(api.isOrderDateAllowed(undefined), false, 'undefined must be rejected');

  // resolveOrderDate: 発注タブの #input-date を引き継ぐ。範囲外・未設定は当日にフォールバック
  inputDateEl.value = today;
  assert.strictEqual(api.resolveOrderDate(), today, 'valid #input-date must be carried over');
  inputDateEl.value = beforeMin;
  assert.strictEqual(api.resolveOrderDate(), today, 'out-of-range #input-date must fall back to today');
  inputDateEl.value = '';
  assert.strictEqual(api.resolveOrderDate(), today, 'empty #input-date must fall back to today');
}

/* ===================================================
   機能2: 発注済みローカル記録・buildOrderedByCode
=================================================== */
function testOrderedCacheAndBuildOrderedByCode() {
  const context = vm.createContext({
    console, JSON, Number, String, Date, Set, Math,
    localStorage: makeStorage(),
    LS_ORDERED: 'orderApp_orderedCodes',
    proposalsData: { pendingOrders: [], analyzedAt: '' }
  });
  const orderedCacheBlock = section(html,
    '/* === ORDERED CACHE (test:ordered-cache) === */',
    '/* === /ORDERED CACHE === */');
  const buildBlock = section(html,
    '/* === BUILD ORDERED BY CODE (test:ordered-cache) === */',
    '/* === /BUILD ORDERED BY CODE === */');
  vm.runInContext(`${orderedCacheBlock}\n${buildBlock}\n globalThis.testApi = {
    readLocal: _readOrderedLocal, writeLocal: _writeOrderedLocal,
    record: _recordOrderedLocal, forget: _forgetOrderedLocal,
    build: buildOrderedByCode, reconcile: _reconcileOrderedLocal,
    isReflected: _isOrderReflected, analyzedDay: _analyzedDayStr, analyzedMs: _analyzedAtMs
  };`, context);
  const api = context.testApi;

  // 保存成功 → ローカル記録 → buildOrderedByCode に反映される
  const entry1 = { params: { date: '2026-08-15', items: JSON.stringify([{ code: 'A', qty: 3 }]) } };
  api.record(entry1, '20260815-001');
  let map = api.build();
  assert.strictEqual(map['A'].qty, 3);
  assert.strictEqual(map['A'].orderDate, '2026-08-15');
  assert.strictEqual(map['A'].orderNo, '20260815-001');

  // 同じ商品を2回発注 → qtyが合算され、orderDateは新しい方になる
  const entry2 = { params: { date: '2026-08-16', items: JSON.stringify([{ code: 'A', qty: 2 }]) } };
  api.record(entry2, '20260816-001');
  map = api.build();
  assert.strictEqual(map['A'].qty, 5, 'qty must be summed across orders');
  assert.strictEqual(map['A'].orderDate, '2026-08-16', 'orderDate must be the newer one');
  assert.strictEqual(map['A'].orderNo, '20260816-001');

  // pendingOrders（GAS側）とローカル記録が同じ発注Noを持つ場合、二重計上されない
  context.proposalsData.pendingOrders = [
    { code: 'A', qty: 5, orderDate: '2026-08-16', orderNo: '20260816-001', isDelayed: false }
  ];
  map = api.build();
  // pendingOrdersの5個 + ローカルのみに残る20260815-001の3個 = 8個（20260816-001はローカル側で重複除外）
  assert.strictEqual(map['A'].qty, 8, 'same orderNo from pendingOrders and local must not be double-counted');

  // isDelayed の伝播（pendingOrders側がtrueならmapにも反映される）
  context.proposalsData.pendingOrders = [
    { code: 'B', qty: 4, orderDate: '2026-08-14', orderNo: '20260814-001', isDelayed: true }
  ];
  map = api.build();
  assert.strictEqual(map['B'].isDelayed, true);

  // _forgetOrderedLocal(orderNo): 該当発注ぶんだけローカル記録から消える
  api.forget('20260815-001');
  const remaining = api.readLocal();
  assert(!remaining.some(r => r.orderNo === '20260815-001'), 'forgotten orderNo must be removed');
  assert(remaining.some(r => r.orderNo === '20260816-001'), 'other orderNo must remain');

  // 修正発注: revisionBaseOrderNoの記録は新しい発注保存時に自動的に消える
  api.forget('20260816-001'); // クリーンアップ
  const entry3 = { params: { date: '2026-08-16', items: JSON.stringify([{ code: 'C', qty: 1 }]), revisionBaseOrderNo: '20260810-001' } };
  api.writeLocal([{ code: 'D', qty: 9, orderDate: '2026-08-10', orderNo: '20260810-001', savedAt: Date.now() }]);
  api.record(entry3, '20260817-001');
  const afterRevision = api.readLocal();
  assert(!afterRevision.some(r => r.orderNo === '20260810-001'), 'revision base orderNo must be dropped on save');
  assert(afterRevision.some(r => r.orderNo === '20260817-001'), 'new order must be recorded');

  // 31日超過の記録は _readOrderedLocal() で自動的に取り除かれる
  const stale = { code: 'E', qty: 1, orderDate: '2020-01-01', orderNo: '20200101-001', savedAt: Date.now() - 31 * 86400000 };
  const fresh = { code: 'F', qty: 1, orderDate: '2026-08-16', orderNo: '20260816-002', savedAt: Date.now() };
  api.writeLocal([stale, fresh]);
  const alive = api.readLocal();
  assert(!alive.some(r => r.orderNo === '20200101-001'), '31-day-old record must be pruned');
  assert(alive.some(r => r.orderNo === '20260816-002'), 'fresh record must remain');

  /* --- v1.66.0: 入荷済みなのに提案から消え続ける不具合の回帰テスト --- */
  // 実害: 2026-08-16のナプラ発注が入荷して在庫に載った後も、ローカル記録が30日残るせいで
  //       8/27時点でナプラの提案85件が「発注済み」として隠れていた

  // ケース1: 発注Noの一部だけが入荷済み → 入荷済みの明細だけが解放され、残りは隠れたまま
  api.writeLocal([
    { code: 'G', qty: 6, orderDate: '2026-08-16', orderNo: '20260816-012', savedAt: Date.now() - 11 * 86400000 },
    { code: 'H', qty: 6, orderDate: '2026-08-16', orderNo: '20260816-012', savedAt: Date.now() - 11 * 86400000 }
  ]);
  context.proposalsData.pendingOrders = [
    { code: 'H', qty: 6, orderDate: '2026-08-16', orderNo: '20260816-012', isDelayed: false }
  ];
  api.reconcile(); // loadProposals が取得成功のたびに呼ぶ突合
  map = api.build();
  assert.strictEqual(map['G'], undefined, 'received line must not be hidden just because the order has other open lines');
  assert.strictEqual(map['H'].qty, 6, 'still-open line must stay hidden');

  // ケース2: 発注が丸ごと入荷済み（サーバーの入荷待ちに無い）→ ローカル記録ごと落ちる
  api.writeLocal([
    { code: 'I', qty: 6, orderDate: '2026-08-16', orderNo: '20260816-012', savedAt: Date.now() - 11 * 86400000 }
  ]);
  context.proposalsData.pendingOrders = [];
  api.reconcile();
  assert.strictEqual(api.readLocal().length, 0, 'record released by the server must be dropped from local');
  assert.strictEqual(api.build()['I'], undefined, 'received item must be proposed again');

  // ケース3: 保存直後（猶予5分以内）は、サーバーがまだ発注を拾えていなくても記録を残す
  api.writeLocal([
    { code: 'J', qty: 6, orderDate: '2026-08-27', orderNo: '20260827-001', savedAt: Date.now() }
  ]);
  context.proposalsData.pendingOrders = [];
  api.reconcile();
  assert.strictEqual(api.readLocal().length, 1, 'just-saved record must survive the grace window');
  assert.strictEqual(api.build()['J'].qty, 6, 'just-saved order must still hide the item');
  api.writeLocal([]);
  context.proposalsData.pendingOrders = [];

  /* --- v1.70.0: 分析より前に発注済みの商品を提案から消さない --- */
  // 実害: 昨日3個だけ発注した商品（推奨15・在庫0）を分析は「あと12個」と提案しているのに、
  //       アプリが「発注済み」として行ごと隠していた（2026-09-04・25件/約20万円が非表示）。
  //       同じ3個を、分析の引き算とアプリの非表示で二重に引いていたのが原因
  context.proposalsData.analyzedAt = '2026-09-04 07:00';

  // 分析より前（前日）の発注 … 提案数量に反映済みなので「隠す判定」の対象外。
  // ただし全件マップ（バッジ表示用）には残る
  context.proposalsData.pendingOrders = [
    { code: 'K', qty: 3, orderDate: '2026-09-03', orderNo: '20260903-001', isDelayed: false }
  ];
  assert.strictEqual(api.build({ unreflectedOnly: true })['K'], undefined,
    'order placed before the analysis must not hide the proposal');
  assert.strictEqual(api.build()['K'].qty, 3,
    'the full map must still carry it so the row can show the badge');

  // 分析と同じ日の発注 … 提案数量に未反映なので従来どおり隠す（二重発注防止）
  context.proposalsData.pendingOrders = [
    { code: 'L', qty: 3, orderDate: '2026-09-04', orderNo: '20260904-001', isDelayed: false }
  ];
  assert.strictEqual(api.build({ unreflectedOnly: true })['L'].qty, 3,
    'order placed on/after the analysis day must still hide the proposal');

  // ローカル記録は発注日ではなく savedAt で前後を判定する（発注日は過去日に変更できるため）
  const analyzedMs = new Date(2026, 8, 4, 7, 0).getTime();
  context.proposalsData.pendingOrders = [];
  api.writeLocal([
    // 過去日で登録したが、保存したのは分析より後 → 未反映＝隠す
    { code: 'M', qty: 2, orderDate: '2026-08-20', orderNo: '20260820-001', savedAt: analyzedMs + 3600000 },
    // 保存も分析より前 → 反映済み＝隠さない
    { code: 'N', qty: 2, orderDate: '2026-09-02', orderNo: '20260902-001', savedAt: analyzedMs - 3600000 }
  ]);
  let unreflected = api.build({ unreflectedOnly: true });
  assert.strictEqual(unreflected['M'].qty, 2, 'local record saved after the analysis must hide the item');
  assert.strictEqual(unreflected['N'], undefined, 'local record saved before the analysis must not hide the item');

  // 分析日時が取れないときは前後を判定せず全件を「未反映」扱い（v1.69.0までと同じ・隠しすぎる側）
  context.proposalsData.analyzedAt = '';
  unreflected = api.build({ unreflectedOnly: true });
  assert.strictEqual(unreflected['M'].qty, 2, 'without analyzedAt everything must fall back to hiding');
  assert.strictEqual(unreflected['N'].qty, 2, 'without analyzedAt everything must fall back to hiding');

  api.writeLocal([]);
  context.proposalsData.pendingOrders = [];
  context.proposalsData.analyzedAt = '';
}

function makeFakeHistorySheet(rows) {
  return {
    rows: rows.map(r => r.slice()),
    getLastRow() { return this.rows.length; },
    getRange(row, column, numRows = 1, numCols = 1) {
      const owner = this;
      return {
        getValues() {
          const result = [];
          for (let r = 0; r < numRows; r++) {
            const source = owner.rows[row - 1 + r] || [];
            const values = [];
            for (let c = 0; c < numCols; c++) values.push(source[column - 1 + c] === undefined ? '' : source[column - 1 + c]);
            result.push(values);
          }
          return result;
        }
      };
    }
  };
}

/* ===================================================
   機能1（GAS側）: generateOrderNo の過去日付フル走査
   （末尾走査だけだと、後日入力で末尾に追記された過去日付の発注Noを見落として
   　重複採番しうる。v1.34.0の修正で「当日以外は全件読み」に変更したことの検証）
=================================================== */
function testGenerateOrderNoPastDateFullScan() {
  // 「今日」を固定する（実行日に依存させない）。2026-08-16 10:00 ローカル時刻固定
  const FIXED_NOW = new Date(2026, 7, 16, 10, 0, 0).getTime();
  class FixedDate extends Date {
    constructor(...args) { if (args.length === 0) super(FIXED_NOW); else super(...args); }
    static now() { return FIXED_NOW; }
  }
  // 発注No（A列）が日付昇順のまま並んでいる通常ケースに加え、末尾に「後日入力」で
  // 過去日付（20260810）の2件目が追記され、日付昇順が崩れているケースを含む
  const hist = makeFakeHistorySheet([
    ['発注No'], // ヘッダー行
    ['20260810-001'],
    ['20260811-001'],
    ['20260812-001'],
    ['20260816-001'], // 当日の発注
    ['20260810-002']  // 後日入力：過去日付(2026-08-10)の2件目が末尾に追記され昇順が崩れている
  ]);
  const context = vm.createContext({
    console, String, Number, parseInt, Math, Array,
    Date: FixedDate,
    SHEET_HISTORY: 'history',
    getSheet: () => hist,
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        const pad = n => String(n).padStart(2, '0');
        const y = date.getFullYear(), mo = pad(date.getMonth() + 1), d = pad(date.getDate());
        return fmt === 'yyyyMMdd' ? `${y}${mo}${d}` : `${y}-${mo}-${d}`;
      }
    }
  });
  const readTailBlock = section(gas,
    'function readTailRowsUntil_(sheet, numCols, isEnough, initialChunk) {',
    '// ヘルパー: セル値を安全に文字列化');
  const generateBlock = section(gas,
    'function generateOrderNo(dateStr) {',
    '// POST: 発注先マスター操作');
  vm.runInContext(`${readTailBlock}\n${generateBlock}\n globalThis.testApi = { generate: generateOrderNo };`, context);

  // 当日ぶん: 末尾走査でも正しく最大値(001)を検出できる
  assert.strictEqual(context.testApi.generate('2026-08-16'), '20260816-002');

  // 過去日付ぶん: 正規の位置(row2)と、後日入力で末尾に追記された(row6)の両方に20260810が
  // 散らばっているが、全件読み（dateKey!==todayKeyの分岐）により両方を検出して正しく003を採番できる。
  // ※このテストの規模（6行）では末尾走査（旧ロジック）でも偶然正しい値を返しうるため、
  //   このテストは「新ロジック（全件読み）が正しいこと」の確認であり、
  //   「旧ロジックが実際に壊れること」の再現ではない（旧ロジックの破綻は数百行規模の
  //   実シートで、対象日より過去の行が非連続に混在したときに起きる。generateOrderNo内の
  //   コメント・発注日変更・発注済み除外_設計プラン.md §1-2参照）
  assert.strictEqual(context.testApi.generate('2026-08-10'), '20260810-003');

  // 未登録の日付は -001 から
  assert.strictEqual(context.testApi.generate('2026-08-09'), '20260809-001');

  // 不正な日付文字列は空文字を返す（呼び出し元 saveOrder 側の日付検証で本来弾かれる想定だが、
  // 万一すり抜けても壊れた発注Noを生成しない防御）
  assert.strictEqual(context.testApi.generate(''), '');
  assert.strictEqual(context.testApi.generate('not-a-date'), '');
}

/* ===================================================
   機能3: 系列まるごと見送り（mgQtyOf / mgRebuild / mgSetPin のガード）
=================================================== */
function testMgSkipGuards() {
  const toastCalls = [];
  const context = vm.createContext({
    console, Math, Object,
    propMgSkipped: new Set(),
    propMgPins: {},
    propMgQty: {},
    showToast: (msg, type) => { toastCalls.push({ msg, type }); }
  });
  const block = section(html,
    '/* === MG SKIP GUARDS (test:mg-skip) === */',
    '/* === /MG SKIP GUARDS === */');
  vm.runInContext(`${block}\n globalThis.testApi = { mgQtyOf, mgRebuild, mgSetPin };`, context);
  const api = context.testApi;

  // 見送り中は常に0本（他商品への配分が残っていても無視する）
  context.propMgSkipped.add('MILFY');
  context.propMgQty['MILFY'] = { X: 12 };
  assert.strictEqual(api.mgQtyOf('MILFY', { code: 'X', proposedQty: 6 }), 0);

  // 見送っていない系列は従来どおりの優先順位（ピン ＞ 組み直し結果 ＞ サーバー提案）を維持する
  context.propMgQty['OTHER'] = { X: 7 };
  assert.strictEqual(api.mgQtyOf('OTHER', { code: 'X', proposedQty: 5 }), 7, 'propMgQty must win over proposedQty');
  context.propMgPins['OTHER'] = { X: 9 };
  assert.strictEqual(api.mgQtyOf('OTHER', { code: 'X', proposedQty: 5 }), 9, 'pin must win over propMgQty');

  // mgRebuild は見送り中だと何もせず、エラートーストだけ出して配分を変更しない
  toastCalls.length = 0;
  api.mgRebuild('MILFY', 2, false);
  assert.strictEqual(toastCalls.length, 1);
  assert.strictEqual(toastCalls[0].type, 'error');
  assert.deepStrictEqual(context.propMgQty['MILFY'], { X: 12 }, 'skipped group allocation must not change');

  // mgSetPin も見送り中は何もしない（±6・チェック操作からの再配分を止める）
  api.mgSetPin('MILFY', 'X', 6);
  assert.strictEqual(context.propMgPins['MILFY'], undefined, 'pin must not be recorded while skipped');
}

/* ===================================================
   機能4（v1.63.0）: カートへの復帰導線
   （再読み込みでcurrentOrderが失われてもカートを捨てずに戻れること）
=================================================== */
function testCartResume() {
  const toastCalls = [];
  const els = {
    'cart-resume-banner': { style: {}, innerHTML: '' },
    'cart-badge-count':   { style: {}, textContent: '' },
    'sel-supplier':       { value: '' }
  };
  const calls = [];
  const context = vm.createContext({
    console, JSON, String, Number, Object, Date,
    cartItems: [],
    currentOrder: { supplierCode: '', supplierName: '', fax: '', staff: '', date: '', outputType: '' },
    currentUser: { name: 'テスト担当' },
    masters: { suppliers: [{ code: '48', name: '株式会社 千代田化学', fax: '000' }] },
    savedOrderNo: null, orderRequestId: null, revisionBaseOrderNo: null,
    sessionStorage: makeStorage(),
    SS_ORDER: 'bf_order',
    document: { getElementById: id => els[id] || null },
    showToast: (msg, type) => { toastCalls.push({ msg, type }) },
    confirm: () => true,
    escHtml: s => String(s == null ? '' : s),
    resolveSupplierName: (code, fallback) => {
      const s = context.masters.suppliers.find(x => x.code === code);
      return (s && s.name) ? s.name : (fallback || code || '');
    },
    resolveOrderDate: () => '2026-08-16',
    // 復帰時に呼ばれる画面遷移系は呼ばれたことだけ記録する
    switchScreen: n => calls.push('switchScreen:' + n),
    goToInput: () => calls.push('goToInput'),
    updateInputHeader: () => calls.push('updateInputHeader'),
    renderItemList: () => calls.push('renderItemList'),
    updateInventoryCartBtn: () => calls.push('updateInventoryCartBtn'),
    saveCart: () => calls.push('saveCart'),
    saveOrderMeta: () => calls.push('saveOrderMeta')
  });
  const block = section(html,
    '/* === CART RESUME (test:cart-resume) === */',
    '/* === /CART RESUME === */');
  vm.runInContext(`${block}\n globalThis.testApi = {
    inferSupplierFromCart, renderCartResumeBanner, resumeCart, discardCart, updateCartNavBadge
  };`, context);
  const api = context.testApi;

  // カートが空ならバナーもバッジも出ない
  api.renderCartResumeBanner();
  assert.strictEqual(els['cart-resume-banner'].style.display, 'none');
  api.updateCartNavBadge();
  assert.strictEqual(els['cart-badge-count'].style.display, 'none');

  // カートあり＋発注先あり → バナーに発注先・品目数・点数が出る
  context.cartItems = [
    { code: 'A', qty: 2, supplierCD: '48' },
    { code: 'B', qty: 3, supplierCD: '48' }
  ];
  context.currentOrder = { supplierCode: '48', supplierName: '株式会社 千代田化学', fax: '', staff: 'x', date: '2026-08-15', outputType: '' };
  api.renderCartResumeBanner();
  assert.strictEqual(els['cart-resume-banner'].style.display, 'block');
  assert(els['cart-resume-banner'].innerHTML.includes('千代田化学'), 'banner must name the supplier');
  assert(els['cart-resume-banner'].innerHTML.includes('2品目・計5点'), 'banner must show item and unit counts');
  assert(els['cart-resume-banner'].innerHTML.includes('resumeCart()'), 'banner must offer the resume action');
  api.updateCartNavBadge();
  assert.strictEqual(els['cart-badge-count'].textContent, '2', 'nav badge shows cart item count');

  // 発注先が確定していれば、そのまま商品入力画面へ戻る
  calls.length = 0;
  api.resumeCart();
  assert(calls.includes('goToInput'), 'resumeCart must return to the input step');
  assert(calls.includes('renderItemList'), 'resumeCart must re-render the cart');

  // ★本命: 再読み込みでcurrentOrderが失われた状態でも、カート商品から発注先を補って復帰できる
  context.currentOrder = { supplierCode: '', supplierName: '', fax: '', staff: '', date: '', outputType: '' };
  assert.strictEqual(api.inferSupplierFromCart(), '48', 'supplier must be inferred from cart items');
  calls.length = 0; toastCalls.length = 0;
  api.resumeCart();
  assert.strictEqual(context.currentOrder.supplierCode, '48', 'resumeCart must restore the supplier');
  assert.strictEqual(context.currentOrder.supplierName, '株式会社 千代田化学');
  assert.strictEqual(context.currentOrder.date, '2026-08-16', 'missing date falls back to resolveOrderDate()');
  assert(calls.includes('goToInput'), 'resumeCart must still reach the input step');
  assert.strictEqual(toastCalls.length, 0, 'no error toast when the supplier can be inferred');

  // 発注先マスターに無いCDしか無い場合は推定せず、プルダウンの選択値を使う
  context.cartItems = [{ code: 'C', qty: 1, supplierCD: '999' }];
  context.currentOrder = { supplierCode: '', supplierName: '', fax: '', staff: '', date: '', outputType: '' };
  assert.strictEqual(api.inferSupplierFromCart(), '', 'unknown supplier CD must not be inferred');
  els['sel-supplier'].value = '48';
  calls.length = 0;
  api.resumeCart();
  assert.strictEqual(context.currentOrder.supplierCode, '48', 'dropdown selection is used as the fallback');

  // 手書きのみ（supplierCD無し）＋プルダウン未選択 → 案内トーストを出して復帰しない
  context.cartItems = [{ code: 'D', qty: 1 }];
  context.currentOrder = { supplierCode: '', supplierName: '', fax: '', staff: '', date: '', outputType: '' };
  els['sel-supplier'].value = '';
  calls.length = 0; toastCalls.length = 0;
  api.resumeCart();
  assert.strictEqual(toastCalls.length, 1);
  assert.strictEqual(toastCalls[0].type, 'error');
  assert(!calls.includes('goToInput'), 'must not jump to the input step without a supplier');

  // 破棄するとカートも発注先も空になる
  context.cartItems = [{ code: 'E', qty: 1, supplierCD: '48' }];
  api.discardCart();
  assert.strictEqual(context.cartItems.length, 0, 'discardCart empties the cart');
  assert.strictEqual(context.currentOrder.supplierCode, '', 'discardCart clears the order info');
}

/* ===================================================
   機能6: 複製して発注（historyItemsToCart / v1.68.0）
   発注履歴は仕入単価・相手商品CDを持たないので、複製時に商品マスターから引き直す。
   v1.67.0以前は purchasePrice を 0 固定にしていたため、複製した行が確認画面の
   合計金額に入らず、あとから追加した行のぶんしか合計に出なかった。
   ※商品名・単価はダミー（実データはコミット禁止。GITHUB-RULES.md参照）
=================================================== */
function testHistoryItemsToCart() {
  const context = vm.createContext({
    console,
    productMaster: [
      { code: 'TEST001', jan: '0000000000011', name: 'テスト商品A',
        unit: '本', supplierCD: '99', makerCode: 'MK-A', purchasePrice: 100 },
      { code: 'TEST002', jan: '0000000000028', name: 'テスト商品B',
        unit: '個', supplierCD: '99', makerCode: 'MK-B', purchasePrice: 250 }
    ]
  });
  const block = section(html,
    '/* === HISTORY TO CART (test:history-to-cart) === */',
    '/* === /HISTORY TO CART === */');
  vm.runInContext(`${block}\n globalThis.testApi = { historyItemsToCart };`, context);
  const api = context.testApi;

  // 商品コードで引ける行は単価・相手商品CDがマスターから入る
  const [a, b] = api.historyItemsToCart([
    { code: 'TEST001', name: 'テスト商品A', qty: 7, unit: '本' },
    { code: 'TEST002', name: 'テスト商品B', qty: 3, unit: '個' }
  ], '99');
  assert.strictEqual(a.purchasePrice, 100, 'copied row must carry the master price');
  assert.strictEqual(a.makerCode, 'MK-A', 'copied row must carry makerCode for PDF output');
  assert.strictEqual(a.janCode, '0000000000011', 'JAN is filled in from the master');
  assert.strictEqual(a.qty, 7, 'quantity comes from the history, not the master');
  assert.strictEqual(b.purchasePrice, 250);

  // 合計金額が全行ぶん出ること（この不具合の実害そのもの）
  const total = [a, b].reduce((s, i) => s + i.qty * (i.purchasePrice || 0), 0);
  assert.strictEqual(total, 7 * 100 + 3 * 250, 'confirm-screen total covers copied rows');

  // JANでしか一致しない行も引ける
  const [c] = api.historyItemsToCart([{ jan: '0000000000028', name: 'B', qty: 2 }], '99');
  assert.strictEqual(c.purchasePrice, 250, 'falls back to a JAN match');

  // 手書き行はマスターを引かず常に0（追加時の仕様と揃える）
  const [d] = api.historyItemsToCart(
    [{ code: 'TEST001', name: '手書き', qty: 3, isHandwritten: true }], '99');
  assert.strictEqual(d.purchasePrice, 0, 'handwritten rows stay unpriced');
  assert.strictEqual(d.isHandwritten, true);

  // マスターに無い商品（終売・マスター未読込）は従来どおり0で通す＝発注自体は成立する
  const [e] = api.historyItemsToCart([{ code: 'NOPE', name: '終売品', qty: 1 }], '99');
  assert.strictEqual(e.purchasePrice, 0, 'unknown products must not throw');
  assert.strictEqual(e.supplierCD, '99', 'falls back to the order supplier code');

  // 明細なし・undefined でも落ちない（vm内の配列なのでlengthで見る）
  assert.strictEqual(api.historyItemsToCart(undefined, '99').length, 0);
  assert.strictEqual(api.historyItemsToCart([], '99').length, 0);
}

(async () => {
  testOrderDateHelpers();
  testOrderedCacheAndBuildOrderedByCode();
  testGenerateOrderNoPastDateFullScan();
  testMgSkipGuards();
  testCartResume();
  testHistoryItemsToCart();
  console.log('All order-feature tests passed.');
})().catch(err => {
  console.error(err);
  process.exit(1);
});
