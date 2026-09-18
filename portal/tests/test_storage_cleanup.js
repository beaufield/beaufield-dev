'use strict';
// ポータル cleanupSharedStorage() の検証。
// 実際のHTMLのスクリプトを実行する。DOM操作は行わず、localStorageだけを模擬する。
// T-4（消してはいけないキーを消さない）が最重要。ここが壊れると業務事故になる。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const html = fs.readFileSync(path.join(__dirname, '..', '..', 'index.html'), 'utf8');
const script = html.match(/<script>\s*([\s\S]*?)<\/script>/)[1];

// localStorageの模擬。fail系フラグでプライベートモード等の例外発生を再現する。
function storage(initial = {}) {
  const values = new Map(Object.entries(initial));
  const keysOrder = Array.from(values.keys());
  return {
    failOnRemoveItem: false,
    failOnGetItem: false,
    failOnLength: false,
    get length() {
      if (this.failOnLength) throw new Error('SecurityError: private mode');
      return values.size;
    },
    key(i) {
      // Map の挿入順を反映する（localStorage.key の実挙動に近い）
      return Array.from(values.keys())[i] ?? null;
    },
    getItem(k) {
      if (this.failOnGetItem) throw new Error('SecurityError');
      return values.has(k) ? values.get(k) : null;
    },
    setItem(k, v) { values.set(k, String(v)); },
    removeItem(k) {
      if (this.failOnRemoveItem) throw new Error('QuotaExceededError-ish');
      values.delete(k);
    },
    _dump() { return Object.fromEntries(values); }
  };
}

function setup(initialKeys = {}) {
  const ls = storage(initialKeys);
  const ctx = vm.createContext({
    console,
    localStorage: ls,
    window: { addEventListener() {} },
    document: {
      addEventListener() {},
      getElementById() { return null; }
    }
  });
  vm.runInContext(script, ctx);
  return { ctx, ls };
}

function run() {
  // T-1: bfc_ で始まる全キーが削除される
  {
    const { ctx, ls } = setup({
      'bfc_history_v1_a': 'x',
      'bfc_summary_week_v1_b': 'y',
      'bfc_x': 'z',
      'bf_session': 'keep-me'
    });
    const result = vm.runInContext('cleanupSharedStorage()', ctx);
    assert.equal(result.removed, 3, 'T-1: bfc_ で始まる3件が削除されること');
    const dump = ls._dump();
    assert.equal(dump['bfc_history_v1_a'], undefined, 'T-1: bfc_history_v1_a が消えていること');
    assert.equal(dump['bfc_summary_week_v1_b'], undefined, 'T-1: bfc_summary_week_v1_b が消えていること');
    assert.equal(dump['bfc_x'], undefined, 'T-1: bfc_x が消えていること');
    assert.equal(dump['bf_session'], 'keep-me', 'T-1: bf_session は残ること');
  }

  // T-2: bf_yoyaku_v* が複数ある場合、最新版だけ残る
  {
    const { ctx, ls } = setup({
      'bf_yoyaku_v1120': 'old1',
      'bf_yoyaku_v1190': 'old2',
      'bf_yoyaku_v1200': 'newest'
    });
    const result = vm.runInContext('cleanupSharedStorage()', ctx);
    assert.equal(result.removed, 2, 'T-2: 旧世代2件が削除されること');
    const dump = ls._dump();
    assert.equal(dump['bf_yoyaku_v1200'], 'newest', 'T-2: 最新版(v1200)が残ること');
    assert.equal(dump['bf_yoyaku_v1120'], undefined, 'T-2: v1120が消えること');
    assert.equal(dump['bf_yoyaku_v1190'], undefined, 'T-2: v1190が消えること');
  }

  // T-3: bf_yoyaku_v* が1件だけの場合は削除されない
  {
    const { ctx, ls } = setup({ 'bf_yoyaku_v1200': 'only-one' });
    const result = vm.runInContext('cleanupSharedStorage()', ctx);
    assert.equal(result.removed, 0, 'T-3: 単独の最新版は削除されないこと');
    assert.equal(ls._dump()['bf_yoyaku_v1200'], 'only-one', 'T-3: 値が保持されること');
  }

  // T-4（最重要）: 消してはいけないキーが1件も削除されない
  {
    const protectedKeys = {
      'bf_session': 'session-data',
      'bf_order_outbox_v1': 'pending-order-data',
      'bf_pm_v2': 'product-master-data',
      'bf_pm_date_v2': '2026-09-16',
      'bf_pm_gas_date_v2': '2026-09-16 01:31',
      'bf_gas_diag_v1': 'diag-data',
      'bf_prod_hist': 'history-data',
      'bf_memo_h': 'memo-data',
      'orderApp_orderedCodes': 'ordered-data'
    };
    const { ctx, ls } = setup(protectedKeys);
    const result = vm.runInContext('cleanupSharedStorage()', ctx);
    assert.equal(result.removed, 0, 'T-4: 保護キーは1件も削除されないこと');
    const dump = ls._dump();
    for (const [k, v] of Object.entries(protectedKeys)) {
      assert.equal(dump[k], v, `T-4: ${k} の値が変わっていないこと`);
    }
  }

  // T-5: 削除対象が0件のとき、記録キーを書かない
  {
    const { ctx, ls } = setup({ 'bf_session': 'x', 'bf_yoyaku_v1200': 'only' });
    const result = vm.runInContext('cleanupSharedStorage()', ctx);
    assert.equal(result.removed, 0, 'T-5: 削除対象なし');
    assert.equal(ls._dump()['bf_portal_cleanup_v1'], undefined, 'T-5: 記録キーが書かれないこと');
  }

  // T-6: localStorage.removeItem が例外を投げても外へ伝播しない
  {
    const { ctx, ls } = setup({ 'bfc_x': 'y', 'bf_session': 'keep' });
    ls.failOnRemoveItem = true;
    assert.doesNotThrow(() => {
      vm.runInContext('cleanupSharedStorage()', ctx);
    }, 'T-6: removeItemの例外が外へ出ないこと');
  }

  // T-7: localStorage.length へのアクセス自体が例外を投げても外へ伝播しない
  //      （プライベートモード等でlocalStorage自体にアクセスできないケースを想定）
  {
    const { ctx, ls } = setup({ 'bfc_x': 'y' });
    ls.failOnLength = true;
    assert.doesNotThrow(() => {
      const result = vm.runInContext('cleanupSharedStorage()', ctx);
      assert.equal(result.removed, 0, 'T-7: 例外時はremoved=0で返ること');
    }, 'T-7: localStorage自体の例外が外へ出ないこと');
  }

  // T-8: 削除後、記録キーの中身が200文字未満であること
  {
    const { ctx, ls } = setup({
      'bfc_history_v1_a': 'x',
      'bfc_history_v1_b': 'y',
      'bf_yoyaku_v1120': 'old',
      'bf_yoyaku_v1200': 'new'
    });
    const result = vm.runInContext('cleanupSharedStorage()', ctx);
    assert.ok(result.removed > 0, 'T-8: 前提として削除が発生していること');
    const record = ls._dump()['bf_portal_cleanup_v1'];
    assert.ok(record, 'T-8: 記録キーが書かれること');
    assert.ok(record.length < 200, `T-8: 記録は200文字未満であること (実際: ${record.length})`);
    // 記録の中身がJSONとしてパースできることも確認
    const parsed = JSON.parse(record);
    assert.equal(parsed.removed, result.removed, 'T-8: 記録内のremoved件数が一致すること');
    assert.ok(typeof parsed.freedChars === 'number', 'T-8: freedCharsが数値であること');
  }

  // T-9: 発注アプリの旧世代キー bf_pm だけを消し、現行の _v2 系は残す
  //      （前方一致や正規表現で実装すると bf_pm_v2 まで巻き込むため、ここで固定する）
  {
    const { ctx, ls } = setup({
      'bf_pm':            'dead-5MB-master',
      'bf_pm_v2':         'CURRENT-master',
      'bf_pm_date_v2':    '2026-09-16',
      'bf_pm_gas_date_v2':'2026-09-16 01:31',
      'bf_prod_hist':     'history-data',
      'bf_session':       'session-data'
    });
    const result = vm.runInContext('cleanupSharedStorage()', ctx);
    const dump = ls._dump();
    assert.equal(dump['bf_pm'], undefined, 'T-9: bf_pm（旧世代・約5.1MB）が削除されること');
    assert.equal(dump['bf_pm_v2'], 'CURRENT-master', 'T-9: bf_pm_v2（現行）は絶対に残ること');
    assert.equal(dump['bf_pm_date_v2'], '2026-09-16', 'T-9: bf_pm_date_v2 は残ること');
    assert.equal(dump['bf_pm_gas_date_v2'], '2026-09-16 01:31', 'T-9: bf_pm_gas_date_v2 は残ること');
    assert.equal(dump['bf_prod_hist'], 'history-data', 'T-9: bf_prod_hist は保護対象。残ること');
    assert.equal(dump['bf_session'], 'session-data', 'T-9: セッションは残ること');
    assert.equal(result.removed, 1, 'T-9: 削除されたのは bf_pm の1件だけであること');
  }

  console.log('Portal storage cleanup: T-1〜T-9 全PASS');
}

run();
