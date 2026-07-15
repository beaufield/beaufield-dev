/**
 * プロジェクト管理ダッシュボード — バックエンドAPI
 *
 * 構成: GitHub Pages（index.html）+ この GAS WebApp + スプレッドシートDB
 * 認証: beaufield-auth 共通セッション（sessions シート照合・is_admin 必須）
 * 同期: dashboard_sync.py が SYNC_TOKEN 付きで POST してくる（読み取り専用ビュー方針。
 *       アプリ側から編集できるのは priority / manual_ball / memo のみ）
 *
 * スクリプトプロパティ（必須・コードに直書きしない）:
 *   DB_SHEET_ID   … ダッシュボードDBスプレッドシートのID
 *   AUTH_SHEET_ID … beaufield-auth スプレッドシートのID
 *   SYNC_TOKEN    … 同期スクリプト用の共有シークレット（ランダム長文字列）
 */

const VERSION = '1.6.0';
const APP_NAME = 'project-dashboard';
const CACHE_TTL_SESSION = 900; // セッション検証キャッシュ 15分（他アプリと同一パターン）

// プロパティ取得（未設定なら明示的にエラー）
function prop_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('スクリプトプロパティ未設定: ' + key);
  return v;
}

// projects シートの列定義（追加時は末尾に足すこと。initDb と upsert の両方が参照する）
const COLS = [
  'id', 'name', 'category', 'status_label', 'ball', 'last_update',
  'summary', 'next_actions', 'relations', 'source', 'path', 'parse_ok',
  'synced_at', 'active',
  // ↓手動管理列（同期で上書きしない）
  'priority', 'manual_ball', 'memo', 'effort',
  // ↓同期列（原本は manual_ops.json → dashboard_sync.py。シート末尾に追加＝v1.4.0）
  'manual_ops',
  // ↓同期列（原本は各プロジェクトの _handoff.md「## 次の一手」→ dashboard_sync.py。シート末尾に追加＝v1.5.0）
  'next_action_short'
];
const MANUAL_COLS = ['priority', 'manual_ball', 'memo', 'effort'];

// assets シートの列定義（v1.6.0）。「発火待ち資産」＝Takashiが自分で起動しないと動かない
// スキル・エージェント・.bat・手動スクリプトのカタログ。原本は SKILL.md や .bat 自身で、
// 収集は asset_scan.py が行う。projects とは別シートにする（資産に ball も status_label も無く、
// 混ぜると matchesTab() の判定と件数バッジが壊れるため）。
// evidence … 最終使用日の出どころ: 'transcript'（起動ログあり・0回なら未使用と断言できる）
//             / 'log'（実行ログのmtime） / 'none'（取得手段が無い＝「記録なし」。未使用と偽らない）
const ASSET_COLS = [
  'id', 'name', 'kind', 'category', 'path', 'trigger', 'summary',
  'last_used', 'use_count', 'evidence', 'synced_at', 'active',
  // ↓手動管理列（同期で上書きしない）
  'memo', 'disposition'
];
const MANUAL_ASSET_COLS = ['memo', 'disposition'];

// ============================================================
// 初期化（GASエディタから一度だけ手動実行する）
// ============================================================
function initDb() {
  const ss = SpreadsheetApp.openById(prop_('DB_SHEET_ID'));
  if (!ss.getSheetByName('projects')) {
    const sh = ss.insertSheet('projects');
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (!ss.getSheetByName('sync_log')) {
    const sh = ss.insertSheet('sync_log');
    sh.getRange(1, 1, 1, 4).setValues([['synced_at', 'project_count', 'device', 'note']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  Logger.log('initDb 完了');
}

// ============================================================
// マイグレーション（v1.2.0・GASエディタから一度だけ手動実行する）
// 既存の projects シートに effort 列（工数感 S/M/L/XL）を末尾追加する
// ============================================================
function migrateAddEffortColumn() {
  const ss = SpreadsheetApp.openById(prop_('DB_SHEET_ID'));
  const sh = ss.getSheetByName('projects');
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('effort') >= 0) {
    Logger.log('effort列は既に存在します（列' + (header.indexOf('effort') + 1) + '）。何もしません');
    return;
  }
  sh.getRange(1, lastCol + 1).setValue('effort');
  Logger.log('effort列を追加しました（列' + (lastCol + 1) + '）');
}

// ============================================================
// マイグレーション（v1.4.0・GASエディタから一度だけ手動実行する）
// 既存の projects シートに manual_ops 列（手動運用バッジの説明文）を末尾追加する
// ============================================================
function migrateAddManualOpsColumn() {
  const ss = SpreadsheetApp.openById(prop_('DB_SHEET_ID'));
  const sh = ss.getSheetByName('projects');
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('manual_ops') >= 0) {
    Logger.log('manual_ops列は既に存在します（列' + (header.indexOf('manual_ops') + 1) + '）。何もしません');
    return;
  }
  sh.getRange(1, lastCol + 1).setValue('manual_ops');
  Logger.log('manual_ops列を追加しました（列' + (lastCol + 1) + '）');
}

// ============================================================
// マイグレーション（v1.5.0・GASエディタから一度だけ手動実行する）
// 既存の projects シートに next_action_short 列（「次の一手」1行表示）を末尾追加する
// ============================================================
function migrateAddNextActionShortColumn() {
  const ss = SpreadsheetApp.openById(prop_('DB_SHEET_ID'));
  const sh = ss.getSheetByName('projects');
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('next_action_short') >= 0) {
    Logger.log('next_action_short列は既に存在します（列' + (header.indexOf('next_action_short') + 1) + '）。何もしません');
    return;
  }
  sh.getRange(1, lastCol + 1).setValue('next_action_short');
  Logger.log('next_action_short列を追加しました（列' + (lastCol + 1) + '）');
}

// ============================================================
// セッション検証（beaufield-auth sessions シート照合・15分キャッシュ）
// ============================================================
function validateSession(token) {
  if (!token) return { valid: false };

  const cache = CacheService.getScriptCache();
  const cacheKey = 'sess_' + token.slice(-32);
  const cached = cache.get(cacheKey);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  try {
    const ss = SpreadsheetApp.openById(prop_('AUTH_SHEET_ID'));
    const sh = ss.getSheetByName('sessions');
    if (!sh) return { valid: false };

    const data = sh.getDataRange().getValues();
    const now = Date.now();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === token) {
        if (Number(data[i][2]) < now) {
          sh.deleteRow(i + 1);
          const r = { valid: false };
          cache.put(cacheKey, JSON.stringify(r), 60);
          return r;
        }
        const userId = String(data[i][1]);
        const r = { valid: true, user_id: userId, is_admin: isAdmin_(ss, userId) };
        cache.put(cacheKey, JSON.stringify(r), CACHE_TTL_SESSION);
        return r;
      }
    }
  } catch (e) {
    Logger.log('セッション検証エラー: ' + e);
  }
  const r = { valid: false };
  cache.put(cacheKey, JSON.stringify(r), 60);
  return r;
}

// users シートの F: is_admin を確認（このアプリは管理者専用）
function isAdmin_(authSs, userId) {
  const users = authSs.getSheetByName('users').getDataRange().getValues();
  for (let i = 1; i < users.length; i++) {
    if (String(users[i][0]) === userId) return users[i][5] === true;
  }
  return false;
}

// 認証必須アクションの共通ガード。失敗時はエラーレスポンス、成功時は null を返す
function authGuard_(token) {
  const auth = validateSession(token);
  if (!auth.valid) {
    return jsonResponse({ success: false, error: 'SESSION_INVALID', message: '認証が必要です。ポータルからログインし直してください。' });
  }
  if (!auth.is_admin) {
    return jsonResponse({ success: false, error: 'FORBIDDEN', message: 'このアプリは管理者専用です。' });
  }
  return null;
}

// ============================================================
// エントリーポイント（GET）: データ取得
// ============================================================
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  try {
    switch (p.action || '') {
      case 'data': {
        const guard = authGuard_(p.session_token || '');
        if (guard) return guard;
        return jsonResponse(getData_());
      }
      case 'version':
        return jsonResponse({ success: true, version: VERSION });
      default:
        return jsonResponse({ success: false, error: '不明なアクション' });
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// エントリーポイント（POST）: 同期・メタ更新
// ============================================================
function doPost(e) {
  let p = {};
  try {
    if (e && e.postData && e.postData.contents) p = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ success: false, error: 'BAD_JSON' });
  }

  try {
    switch (p.action || '') {
      case 'sync': {
        // 同期スクリプト専用（セッションではなく共有トークンで認証）
        if (String(p.token || '') !== prop_('SYNC_TOKEN')) {
          return jsonResponse({ success: false, error: 'SYNC_TOKEN_INVALID' });
        }
        return jsonResponse(syncProjects_(p));
      }
      case 'syncAssets': {
        if (String(p.token || '') !== prop_('SYNC_TOKEN')) {
          return jsonResponse({ success: false, error: 'SYNC_TOKEN_INVALID' });
        }
        return jsonResponse(syncAssets_(p));
      }
      case 'updateMeta': {
        const guard = authGuard_(p.session_token || '');
        if (guard) return guard;
        return jsonResponse(updateMeta_(p));
      }
      case 'reorderPriorities': {
        const guard = authGuard_(p.session_token || '');
        if (guard) return guard;
        return jsonResponse(reorderPriorities_(p));
      }
      default:
        return jsonResponse({ success: false, error: '不明なアクション' });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// 同期: 全プロジェクトを upsert（手動管理列は保持）
// ============================================================
function syncProjects_(p) {
  const projects = p.projects || [];
  if (!projects.length) return { success: false, error: 'プロジェクトが空です' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(prop_('DB_SHEET_ID'));
    const sh = ss.getSheetByName('projects');
    const data = sh.getDataRange().getValues();
    const header = data[0];
    const colIdx = {};
    header.forEach((h, i) => colIdx[h] = i);

    // 既存行を id → {rowIndex, manualValues} でマップ
    const existing = {};
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][colIdx['id']]);
      const manual = {};
      MANUAL_COLS.forEach(c => manual[c] = data[i][colIdx[c]]);
      existing[id] = { row: i + 1, manual: manual };
    }

    const syncedAt = new Date();
    const syncedIds = {};
    const newRows = [];

    projects.forEach(proj => {
      syncedIds[proj.id] = true;
      const rowValues = COLS.map(c => {
        if (MANUAL_COLS.indexOf(c) >= 0) {
          // 手動列: 既に値が入っていれば（Takashiの編集・過去の自動見立て問わず）必ず保持する。
          // 空欄の場合のみ、同期側から届いた初期値（例: effortの自動見立て）で埋める。
          const existingVal = existing[proj.id] ? existing[proj.id].manual[c] : '';
          if (existingVal !== '' && existingVal != null) return existingVal;
          return proj[c] != null ? proj[c] : '';
        }
        switch (c) {
          case 'next_actions': return JSON.stringify(proj.next_actions || []);
          case 'relations':    return JSON.stringify(proj.relations || []);
          case 'parse_ok':     return proj.parse_ok !== false;
          case 'synced_at':    return syncedAt;
          case 'active':       return true;
          default:             return proj[c] != null ? proj[c] : '';
        }
      });
      if (existing[proj.id]) {
        sh.getRange(existing[proj.id].row, 1, 1, COLS.length).setValues([rowValues]);
      } else {
        newRows.push(rowValues);
      }
    });

    if (newRows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, COLS.length).setValues(newRows);
    }

    // 今回の同期に含まれなかった既存プロジェクトは active=FALSE（アーカイブ扱い・行は消さない）
    Object.keys(existing).forEach(id => {
      if (!syncedIds[id]) {
        sh.getRange(existing[id].row, colIdx['active'] + 1).setValue(false);
      }
    });

    // ログ
    const log = ss.getSheetByName('sync_log');
    log.appendRow([syncedAt, projects.length, p.device || '', p.generated_at || '']);

    return { success: true, upserted: projects.length, new_count: newRows.length };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 同期: 発火待ち資産を upsert（手動管理列は保持）— v1.6.0
// syncProjects_() と同じ作法。assets シートが無ければ自動で作る（下記 assetsSheet_ 参照）
// ============================================================
function syncAssets_(p) {
  const assets = p.assets || [];
  if (!assets.length) return { success: false, error: '資産が空です' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = assetsSheet_();
    const data = sh.getDataRange().getValues();
    const header = data[0];
    const colIdx = {};
    header.forEach((h, i) => colIdx[h] = i);

    // 既存行を id → {rowIndex, manualValues} でマップ
    const existing = {};
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][colIdx['id']]);
      const manual = {};
      MANUAL_ASSET_COLS.forEach(c => manual[c] = data[i][colIdx[c]]);
      existing[id] = { row: i + 1, manual: manual };
    }

    const syncedAt = new Date();
    const syncedIds = {};
    const newRows = [];

    assets.forEach(a => {
      syncedIds[a.id] = true;
      const rowValues = ASSET_COLS.map(c => {
        if (MANUAL_ASSET_COLS.indexOf(c) >= 0) {
          // 手動列: Takashiの編集を同期で踏み潰さない（projects側と同じ規約）
          const existingVal = existing[a.id] ? existing[a.id].manual[c] : '';
          if (existingVal !== '' && existingVal != null) return existingVal;
          return a[c] != null ? a[c] : '';
        }
        switch (c) {
          case 'synced_at': return syncedAt;
          case 'active':    return true;
          default:          return a[c] != null ? a[c] : '';
        }
      });
      if (existing[a.id]) {
        sh.getRange(existing[a.id].row, 1, 1, ASSET_COLS.length).setValues([rowValues]);
      } else {
        newRows.push(rowValues);
      }
    });

    if (newRows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, ASSET_COLS.length).setValues(newRows);
    }

    // 消された資産は active=FALSE（行は消さない。memo/disposition を失わないため）
    Object.keys(existing).forEach(id => {
      if (!syncedIds[id]) {
        sh.getRange(existing[id].row, colIdx['active'] + 1).setValue(false);
      }
    });

    return { success: true, upserted: assets.length, new_count: newRows.length };
  } finally {
    lock.releaseLock();
  }
}

// assets シートを取得する。無ければ見出し付きで作る。
// ⚠️ ここで作りきることが重要。列追加（migrateAddXxxColumn 系）は「コード反映・デプロイの前に
//    マイグレーションを実行しないと見出し行が空になる」順序事故を v1.5.0 で起こしている。
//    新規シートは同期時に自動生成すれば、その手順自体が不要になり事故が構造的に起きない。
function assetsSheet_() {
  const ss = SpreadsheetApp.openById(prop_('DB_SHEET_ID'));
  let sh = ss.getSheetByName('assets');
  if (!sh) {
    sh = ss.insertSheet('assets');
    sh.getRange(1, 1, 1, ASSET_COLS.length).setValues([ASSET_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ============================================================
// メタ更新: 手動管理列のみ変更可（読み取り専用ビュー方針の担保）
// projects は priority / manual_ball / memo / effort、assets は memo / disposition
// ============================================================
function updateMeta_(p) {
  const id = String(p.id || '');
  if (!id) return { success: false, error: 'idが必要です' };

  // 資産（id が "skill:xxx" 等）は assets シート側を更新する
  if (p.target === 'asset') return updateAssetMeta_(p, id);

  const ss = SpreadsheetApp.openById(prop_('DB_SHEET_ID'));
  const sh = ss.getSheetByName('projects');
  const data = sh.getDataRange().getValues();
  const header = data[0];
  const colIdx = {};
  header.forEach((h, i) => colIdx[h] = i);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIdx['id']]) === id) {
      MANUAL_COLS.forEach(c => {
        if (p[c] !== undefined) sh.getRange(i + 1, colIdx[c] + 1).setValue(p[c]);
      });
      return { success: true };
    }
  }
  return { success: false, error: 'プロジェクトが見つかりません: ' + id };
}

function updateAssetMeta_(p, id) {
  const sh = assetsSheet_();
  const data = sh.getDataRange().getValues();
  const header = data[0];
  const colIdx = {};
  header.forEach((h, i) => colIdx[h] = i);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIdx['id']]) === id) {
      MANUAL_ASSET_COLS.forEach(c => {
        if (p[c] !== undefined) sh.getRange(i + 1, colIdx[c] + 1).setValue(p[c]);
      });
      return { success: true };
    }
  }
  return { success: false, error: '資産が見つかりません: ' + id };
}

// ============================================================
// 優先度の一括並び替え（PC側のドラッグ&ドロップ操作から呼ばれる）
// updates: [{id, priority}, ...]（priority列のみ更新。他のmanual列には触れない）
// ============================================================
function reorderPriorities_(p) {
  const updates = p.updates || [];
  if (!updates.length) return { success: false, error: '更新対象がありません' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(prop_('DB_SHEET_ID'));
    const sh = ss.getSheetByName('projects');
    const data = sh.getDataRange().getValues();
    const header = data[0];
    const idCol = header.indexOf('id');
    const priCol = header.indexOf('priority');

    const idToRow = {};
    for (let i = 1; i < data.length; i++) idToRow[String(data[i][idCol])] = i + 1;

    let updated = 0;
    updates.forEach(u => {
      const row = idToRow[String(u.id)];
      if (row) { sh.getRange(row, priCol + 1).setValue(u.priority); updated++; }
    });

    return { success: true, updated: updated };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// データ取得: active な全プロジェクト
// ============================================================
function getData_() {
  const ss = SpreadsheetApp.openById(prop_('DB_SHEET_ID'));
  const sh = ss.getSheetByName('projects');
  const data = sh.getDataRange().getValues();
  const header = data[0];

  const projects = [];
  let lastSync = null;
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    header.forEach((h, j) => {
      let v = data[i][j];
      if (h === 'next_actions' || h === 'relations') {
        try { v = JSON.parse(v || '[]'); } catch (e) { v = []; }
      }
      if (v instanceof Date) v = v.toISOString();
      obj[h] = v;
    });
    if (obj.active === true) projects.push(obj);
    if (obj.synced_at && (!lastSync || obj.synced_at > lastSync)) lastSync = obj.synced_at;
  }
  // 資産は同じレスポンスに載せる（往復とセッション検証を増やさないため）
  return { success: true, version: VERSION, synced_at: lastSync,
           projects: projects, assets: getAssets_(ss) };
}

// active な全資産。初回同期前は assets シートが存在しないので、その場合は空配列を返す
// （GET でシートを作らない）
function getAssets_(ss) {
  const sh = ss.getSheetByName('assets');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const header = data[0];

  const assets = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    header.forEach((h, j) => {
      let v = data[i][j];
      if (v instanceof Date) {
        // last_used は "2026-07-15" 文字列で書き込むがシートが日付型に自動変換する。
        // フロントは日付部分だけを期待しているので YYYY-MM-DD に戻す
        v = (h === 'last_used')
          ? Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd')
          : v.toISOString();
      }
      obj[h] = v;
    });
    if (obj.active === true) assets.push(obj);
  }
  return assets;
}

// ============================================================
// 共通: JSONレスポンス
// ============================================================
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
