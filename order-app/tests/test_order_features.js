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
    proposalsData: { pendingOrders: [] }
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
    build: buildOrderedByCode
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

(async () => {
  testOrderDateHelpers();
  testOrderedCacheAndBuildOrderedByCode();
  testGenerateOrderNoPastDateFullScan();
  testMgSkipGuards();
  console.log('All order-feature tests passed.');
})().catch(err => {
  console.error(err);
  process.exit(1);
});
