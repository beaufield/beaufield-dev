'use strict';
// 商品マスターのIndexedDB化（v1.74.0）の検証。
// 実IndexedDBはNodeに無いため、IndexedDBのモックは作らず
// _masterStoreGet/_masterStoreSet/_masterStoreDelete をテスト側で差し替える
// （仕様書 ストレージ枯渇対策_実装仕様_Sonnet用.md §4-6 のU-1〜U-7）。
const assert = require('node:assert/strict');
const { setup } = require('./test_master_recovery');

// テスト用のIndexedDB二重（in-memory）。テストごとにリセットする。
function installFakeIdb(t, { failSet = false } = {}) {
  t.run(`
    globalThis.__idbStore = null;      // null = 空。配列 = 保存済みマスター
    globalThis.__idbSetShouldFail = ${failSet};
    globalThis.__idbSetCalls = 0;
    _masterStoreGet = async () => __idbStore;
    _masterStoreSet = async (data) => {
      __idbSetCalls++;
      if (__idbSetShouldFail) return false;
      __idbStore = data;
      _idbMasterPresent = true;
      return true;
    };
    _masterStoreDelete = async () => { __idbStore = null; _idbMasterPresent = false; return true; };
  `);
}

async function main() {
  // U-1: IndexedDBに保存済み → そこから復元。localStorageは読まない
  {
    const t = setup();
    installFakeIdb(t);
    t.run(`__idbStore = [{code:'IDB1', name:'IDB保存商品'}];`);
    t.ctx.localStorage.setItem('bf_pm_v2', JSON.stringify([{ code: 'LS1', name: 'localStorage商品（読まれないはず）' }]));
    const result = await t.run('loadProductMasterFromCache()');
    assert.equal(result, true, 'U-1: 復元に成功すること');
    // vm.createContextの配列/オブジェクトはNode本体と別レルムのため、strictな構造比較はJSON経由で行う
    assert.equal(t.run('JSON.stringify(productMaster)'), JSON.stringify([{ code: 'IDB1', name: 'IDB保存商品' }]), 'U-1: IndexedDBの内容が使われること');
    // localStorageの内容がそのまま残っている＝読みに行っていない証拠（読んでいたら移行処理でIndexedDBが上書きされる）
    assert.equal(
      JSON.parse(t.ctx.localStorage.getItem('bf_pm_v2'))[0].code, 'LS1',
      'U-1: localStorageの旧データが手つかずのまま残ること（＝読んでいない証拠）'
    );
  }

  // U-2: IndexedDBが空・localStorageにbf_pm_v2あり → 復元でき、IndexedDBへ移行され、localStorageのbf_pm_v2が消える
  {
    const t = setup();
    installFakeIdb(t);
    t.ctx.localStorage.setItem('bf_pm_v2', JSON.stringify([{ code: 'LEGACY1', name: '旧マスター商品' }]));
    const result = await t.run('loadProductMasterFromCache()');
    assert.equal(result, true, 'U-2: 復元に成功すること');
    assert.equal(t.run('JSON.stringify(productMaster)'), JSON.stringify([{ code: 'LEGACY1', name: '旧マスター商品' }]), 'U-2: localStorageの内容が使われること');
    assert.equal(t.run('JSON.stringify(__idbStore)'), JSON.stringify([{ code: 'LEGACY1', name: '旧マスター商品' }]), 'U-2: IndexedDBへ移行されること');
    assert.equal(t.ctx.localStorage.getItem('bf_pm_v2'), null, 'U-2: localStorageのbf_pm_v2が消えること');
  }

  // U-3: 移行中に_masterStoreSetがfalse → localStorageのbf_pm_v2を消さない（データを失わない）
  {
    const t = setup();
    installFakeIdb(t, { failSet: true });
    t.ctx.localStorage.setItem('bf_pm_v2', JSON.stringify([{ code: 'LEGACY2', name: '旧マスター商品2' }]));
    const result = await t.run('loadProductMasterFromCache()');
    assert.equal(result, true, 'U-3: 移行に失敗しても、その場では読み込みは成功として扱うこと（データは手元にある）');
    assert.equal(t.run('JSON.stringify(productMaster)'), JSON.stringify([{ code: 'LEGACY2', name: '旧マスター商品2' }]), 'U-3: localStorageの内容が使われること');
    assert.notEqual(t.ctx.localStorage.getItem('bf_pm_v2'), null, 'U-3: 移行失敗時はlocalStorageのbf_pm_v2を消さないこと（データ保全）');
    assert.equal(t.run('__idbStore'), null, 'U-3: IndexedDB側には書き込まれていないこと');
  }

  // U-4: _masterStoreSetがfalse（保存不可） → _productRead.cacheWarning === true・画面に警告
  {
    const t = setup();
    installFakeIdb(t, { failSet: true });
    t.run(`productMaster = [{code:'NEW1', name:'新規商品'}];
      _productRead.gasDate = '2026-09-16 01:31'; _productRead.updatedDate = getTodayStr();`);
    await t.run('persistProductMaster()');
    assert.equal(t.run('_productRead.cacheWarning'), true, 'U-4: cacheWarningが立つこと');
    assert.match(
      t.elements.get('master-badge').querySelector('.master-badge-sub').textContent,
      /端末への保存に失敗/,
      'U-4: 画面に保存失敗の警告が出ること'
    );
  }

  // U-5: 保存成功 → bf_pm_date_v2が最後にセットされる（順序の維持）
  {
    const t = setup();
    installFakeIdb(t);
    t.ctx.localStorage.setItem('bf_pm_date_v2', '2000-01-01'); // 古い日付を仕込んでおく
    t.run(`productMaster = [{code:'NEW2', name:'新規商品2'}];
      _productRead.gasDate = '2026-09-16 01:31'; _productRead.updatedDate = getTodayStr();`);
    await t.run('persistProductMaster()');
    assert.equal(t.run('_productRead.cacheWarning'), false, 'U-5: 保存成功でcacheWarningが立たないこと');
    assert.equal(t.run('JSON.stringify(__idbStore)'), JSON.stringify([{ code: 'NEW2', name: '新規商品2' }]), 'U-5: IndexedDBへ保存されること');
    assert.equal(t.ctx.localStorage.getItem('bf_pm_gas_date_v2'), '2026-09-16 01:31', 'U-5: GAS更新日時が保存されること');
    assert.equal(t.ctx.localStorage.getItem('bf_pm_date_v2'), t.run('getTodayStr()'), 'U-5: 保存成功時は本日の日付が最後にセットされること');
  }

  // U-6: 保存が途中で失敗 → bf_pm_date_v2がセットされない（＝翌起動で「本日更新済」と誤認しない）
  {
    const t = setup();
    installFakeIdb(t, { failSet: true });
    t.ctx.localStorage.setItem('bf_pm_date_v2', '2000-01-01'); // 古い日付（前回成功時のもの）
    t.run(`productMaster = [{code:'NEW3', name:'新規商品3'}];
      _productRead.gasDate = '2026-09-16 01:31'; _productRead.updatedDate = getTodayStr();`);
    await t.run('persistProductMaster()');
    assert.equal(t.run('_productRead.cacheWarning'), true, 'U-6: 保存失敗でcacheWarningが立つこと');
    // 「日付を最後に確定する。途中失敗したキャッシュを翌起動で最新扱いしない」既存設計の維持を確認
    assert.equal(t.ctx.localStorage.getItem('bf_pm_date_v2'), null, 'U-6: 保存失敗時はbf_pm_date_v2が消えたままであること（古い日付で誤認させない）');
  }

  console.log('Product master IndexedDB (U-1〜U-6) PASS');
}

main().catch(e => { console.error(e); process.exitCode = 1; });
