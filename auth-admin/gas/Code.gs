// ============================================================
// 権限管理アプリ（auth-admin） - Google Apps Script バックエンド
// ============================================================
// 目的: beaufield-auth の user_app_roles を直接編集せず、アプリ×社員の
//       マトリクス画面のタップだけで権限設定を完結させる。
// 利用者: ALLOWED_USER_IDS で指定した本人のみ（想定は前島崇志1名）。
//
// 設計書: 開発・自動化/auth-admin/権限管理アプリ_設計.md（Git非公開）
//
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   AUTH_SHEET_ID     : beaufield-auth スプレッドシートID（共通）
//   ALLOWED_USER_IDS  : このアプリを使える user_id（カンマ区切り可・例: U001）
//                        ※ 未設定の場合は誰も使えない（フェイルクローズ）
//
// 初回セットアップ手順:
//   1. このファイル全文をGASエディタに貼り付けて保存
//   2. 上記2つのスクリプトプロパティを設定する
//   3. 「デプロイ」→ 種類=ウェブアプリ / 実行するユーザー=自分 / アクセス=全員 でデプロイ
//   4. GASエディタの関数選択で setupAuthAdminSheets を選び、実行する
//      → beaufield-auth に apps / role_audit シートを作成し初期データを投入する
//      → ALLOWED_USER_IDS の各ユーザーへ auth-admin の admin 権限を付与する
//      → 何度実行しても安全（既にデータがあればヘッダーの確認だけで終わる）
//   5. scriptId・GAS URL を Dropbox/.claude/handover-secrets.md §8 に記録する
// ============================================================

const _PROPS        = PropertiesService.getScriptProperties();
const VERSION        = 'v1.3.0';
const APP_NAME        = 'auth-admin';
const AUTH_SHEET_ID  = _PROPS.getProperty('AUTH_SHEET_ID');

// セッション検証結果（管理者判定込み）のキャッシュ秒数。
// 🔴 権限マトリクスのデータ自体はキャッシュしない（保存直後に古い値が返るのを防ぐ）。
const CACHE_TTL_AUTH = 60;

// saveMatrix / migrateRoles の冪等化（同じ request_id の再送を無視する）猶予秒数。
const CACHE_TTL_IDEMPOTENCY = 600;

// user_app_roles で意味を持つロール語彙（この3語以外は「旧語彙」として legacy_roles に出す）
const KNOWN_ROLES = ['admin', 'user'];

// 旧語彙→新語彙のエイリアス（migrateRoles で使用）
const ROLE_ALIAS = { manager: 'admin', editor: 'admin', staff: 'user', viewer: 'user' };

// ============================================================
// エントリーポイント
// ============================================================

// GET は稼働確認用のみ。情報は一切返さない。
function doGet(e) {
  return _json({ ok: true, version: VERSION });
}

function doPost(e) {
  let body;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return _json({ success: false, error: 'PARAM_LOST' });
    }
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return _json({ success: false, error: 'PARAM_LOST' });
  }

  const action = String(body.action || '');
  const token  = String(body.session_token || '');

  let auth = _validateAdmin(token);
  if (!auth.ok && auth.transient) {
    // 🔴 誤ログアウト対策: 正常な200でSESSION_INVALIDが返る既知の持病があるため、
    //    一時障害シグナルが立った場合だけ1回だけ再検証してから拒否を確定する。
    auth = _validateAdmin(token);
  }
  if (!auth.ok) {
    return _json({ success: false, error: auth.error || 'SESSION_INVALID' });
  }

  try {
    switch (action) {
      case 'getMatrix':    return _json(_getMatrix());
      case 'saveMatrix':   return _json(_saveMatrix(body, auth.user_id));
      case 'migrateRoles': return _json(_migrateRoles(body, auth.user_id));
      default:              return _json({ success: false, error: 'UNKNOWN_ACTION' });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return _json({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// 認証（ポータルのvalidateSessionは呼ばず、beaufield-authを直接読む。
// order-app/gas/Code.gs の validateSession を移植し、三重ガードを追加したもの。
// 🔴 sessions/usersシートには一切書き込まない・削除しない）
// ============================================================
function _validateAdmin(token) {
  if (!token) return { ok: false, error: 'SESSION_INVALID' };

  const cache    = CacheService.getScriptCache();
  const cacheKey = 'sess_authadmin_v1_' + token.slice(-32);
  const cached   = cache.get(cacheKey);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  if (!AUTH_SHEET_ID) return { ok: false, error: 'SESSION_INVALID', transient: true };

  try {
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

    const sessSh = ss.getSheetByName('sessions');
    if (!sessSh) return { ok: false, error: 'SESSION_INVALID', transient: true };

    const sessions = sessSh.getDataRange().getValues();
    const now = Date.now();
    let userId = null;
    for (let i = 1; i < sessions.length; i++) {
      if (String(sessions[i][0]) === token) {
        if (Number(sessions[i][2]) > now) userId = String(sessions[i][1]);
        break;
      }
    }
    if (!userId) {
      const r = { ok: false, error: 'SESSION_INVALID' };
      cache.put(cacheKey, JSON.stringify(r), CACHE_TTL_AUTH);
      return r;
    }

    const usersSh = ss.getSheetByName('users');
    const rolesSh = ss.getSheetByName('user_app_roles');
    if (!usersSh || !rolesSh) return { ok: false, error: 'SESSION_INVALID', transient: true };

    const users = usersSh.getDataRange().getValues();
    let userRow = null;
    for (let j = 1; j < users.length; j++) {
      if (String(users[j][0]).trim() === userId) { userRow = users[j]; break; }
    }
    const active  = !!userRow && (userRow[3] === true || userRow[3] === 'TRUE');
    const isAdmin = !!userRow && (userRow[5] === true || userRow[5] === 'TRUE');
    if (!userRow || !active) {
      const r = { ok: false, error: 'SESSION_INVALID' };
      cache.put(cacheKey, JSON.stringify(r), CACHE_TTL_AUTH);
      return r;
    }

    // user_app_roles で auth-admin のロールを取得（最初の一致のみ・他アプリと同じ挙動）
    const roles = rolesSh.getDataRange().getValues();
    let role = '';
    for (let k = 1; k < roles.length; k++) {
      if (String(roles[k][0]).trim() === userId && String(roles[k][1]).trim() === APP_NAME) {
        role = normRole_(roles[k][2]);
        break;
      }
    }

    // 三重ガード: role==='admin' ∧ users.is_admin===TRUE ∧ ALLOWED_USER_IDS に含まれる
    const allowed   = _allowedUserIds();
    const permitted = isAdmin && role === 'admin' && allowed.indexOf(userId) !== -1;

    const result = permitted
      ? { ok: true, user_id: userId, name: String(userRow[1] || userId) }
      : { ok: false, error: 'FORBIDDEN' };

    cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_AUTH);
    return result;
  } catch (e) {
    Logger.log('_validateAdmin error: ' + e);
    return { ok: false, error: 'SESSION_INVALID', transient: true };
  }
}

function _allowedUserIds() {
  const raw = String(_PROPS.getProperty('ALLOWED_USER_IDS') || '');
  return raw.split(',').map(s => s.trim()).filter(s => s);
}

// ============================================================
// ロール値の正規化。未登録と 'none' を同一視する。
// 🔴 すべてのロール比較はこれを経由すること（楽観ロックの誤検知防止）。
// ============================================================
function normRole_(v) {
  const s = String(v == null ? '' : v).trim().toLowerCase();
  return (s === '' || s === 'none') ? '' : s;
}

// ============================================================
// getMatrix: マトリクス描画に必要な全データを返す（読み取り専用）
// ============================================================
function _getMatrix() {
  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

  const usersSh = ss.getSheetByName('users');
  const appsSh  = ss.getSheetByName('apps');
  const rolesSh = ss.getSheetByName('user_app_roles');
  if (!appsSh) return { success: false, error: 'APPS_SHEET_MISSING' };
  if (!usersSh || !rolesSh) return { success: false, error: 'INTERNAL_ERROR' };

  // ── users ──────────────────────────────────────────
  const userRows = usersSh.getDataRange().getValues();
  const users = [];
  const knownUserIds = {};
  for (let i = 1; i < userRows.length; i++) {
    const r = userRows[i];
    const uid = String(r[0] || '').trim();
    if (!uid) continue;
    knownUserIds[uid] = true;
    users.push({
      user_id:    uid,
      name:       String(r[1] || uid),
      short_name: String(r[7] || '').trim() || String(r[1] || uid), // H列（存在しない場合はnameを使う）
      active:     r[3] === true || r[3] === 'TRUE',
      is_admin:   r[5] === true || r[5] === 'TRUE'
    });
  }

  // ── apps（archivedは除外・hiddenは含める） ─────────────
  const appRows = appsSh.getDataRange().getValues();
  const apps = [];
  const knownAppNames = {};
  for (let a = 1; a < appRows.length; a++) {
    const ar = appRows[a];
    const appName = String(ar[0] || '').trim();
    if (!appName) continue;
    const status = String(ar[7] || 'active').trim() || 'active';
    if (status === 'archived') continue;
    knownAppNames[appName] = true;

    let roleList = String(ar[4] || 'user').split(',')
      .map(s => s.trim().toLowerCase()).filter(s => s);
    if (!roleList.some(r => KNOWN_ROLES.indexOf(r) !== -1)) roleList = ['user'];

    const roleLabels = {};
    String(ar[5] || '').split(',').forEach(pair => {
      const kv = pair.split('=');
      if (kv.length === 2 && kv[0].trim()) roleLabels[kv[0].trim().toLowerCase()] = kv[1].trim();
    });

    apps.push({
      app_name:    appName,
      label:       String(ar[1] || appName),
      icon:        String(ar[2] || '📱'),
      url:         String(ar[3] || ''),
      roles:       roleList,
      role_labels: roleLabels,
      sort_order:  Number(ar[6]) || 999,
      status:      status
    });
  }
  apps.sort((x, y) => x.sort_order - y.sort_order);

  // 🆕 v1.3.0: アプリごとの許可ロール索引。
  // 🔴 KNOWN_ROLES（グローバル）とは別物。'admin' は KNOWN_ROLES には入っているので
  //    「apps シートでそのアプリに許可されていない admin」は legacy_roles では検出できなかった。
  //    その値は画面の循環（[''].concat(app.roles)）にも無いため、1クリックで黙って消えていた。
  const allowedByApp = {};
  apps.forEach(a => { allowedByApp[a.app_name] = a.roles || []; });

  // ── roles + legacy_roles + duplicate_rows + orphan_* ──
  // 🔴 全アプリが「最初の一致で break」する挙動に合わせ、同一(user_id,app_name)の
  //    2行目以降は roles/legacy_roles に反映しない（重複は duplicate_rows で警告するだけ）。
  const roleRows = rolesSh.getDataRange().getValues();
  const roles = {};
  const legacyRoles = [];
  const outOfVocab = [];   // 🆕 v1.3.0: 値そのものは正しいが、そのアプリでは許可されていないロール
  const seenPairs = {};
  const dupCount = {};
  const orphanAppSet = {};
  const orphanUserSet = {};

  for (let p = 1; p < roleRows.length; p++) {
    const rr = roleRows[p];
    const uid = String(rr[0] || '').trim();
    const appName = String(rr[1] || '').trim();
    if (!uid || !appName) continue;

    const key = uid + '/' + appName;
    if (seenPairs[key]) {
      dupCount[key] = (dupCount[key] || 1) + 1;
    } else {
      seenPairs[key] = true;
      const roleVal = normRole_(rr[2]);
      if (roleVal) {
        if (!roles[uid]) roles[uid] = {};
        roles[uid][appName] = roleVal;
        if (KNOWN_ROLES.indexOf(roleVal) === -1) {
          legacyRoles.push({ user_id: uid, app_name: appName, role: roleVal });
        } else if (allowedByApp[appName] && allowedByApp[appName].indexOf(roleVal) === -1) {
          // 🔴 apps シートの roles 語彙に無い。画面では循環に含まれないので、
          //    放置すると次のクリックで黙って別の値に落ちる
          outOfVocab.push({ user_id: uid, app_name: appName, role: roleVal,
                            allowed: allowedByApp[appName] });
        }
      }
    }

    if (!knownAppNames[appName]) orphanAppSet[appName] = true;
    if (!knownUserIds[uid]) orphanUserSet[uid] = true;
  }

  const duplicateRows = Object.keys(dupCount).map(key => {
    const [uid, appName] = key.split('/');
    return { user_id: uid, app_name: appName, count: dupCount[key] };
  });

  return {
    success: true,
    users: users,
    apps: apps,
    roles: roles,
    legacy_roles: legacyRoles,
    out_of_vocab: outOfVocab,
    duplicate_rows: duplicateRows,
    orphan_apps: Object.keys(orphanAppSet),
    orphan_users: Object.keys(orphanUserSet),
    fetched_at: Date.now(),
    version: VERSION
  };
}

// ============================================================
// saveMatrix: 差分をまとめて反映する
// ============================================================
function _saveMatrix(payload, actorUserId) {
  const requestId = String(payload.request_id || '').trim();
  if (!requestId) return { success: false, error: 'REQUEST_ID_REQUIRED' };

  const cache   = CacheService.getScriptCache();
  const idemKey = 'saveMatrix_' + requestId;
  const cached  = cache.get(idemKey);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const changes = Array.isArray(payload.changes) ? payload.changes : [];
  if (!changes.length) return { success: false, error: 'NO_CHANGES' };

  const lock = LockService.getScriptLock();
  let gotLock = false;
  try { gotLock = lock.tryLock(30000); } catch (e) {}
  if (!gotLock) return { success: false, error: 'BUSY' };

  try {
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const appsSh  = ss.getSheetByName('apps');
    const rolesSh = ss.getSheetByName('user_app_roles');
    if (!appsSh || !rolesSh) return { success: false, error: 'INTERNAL_ERROR' };

    // アプリごとの許可ロールを取得（'none' は常に許可）
    const appRows = appsSh.getDataRange().getValues();
    const allowedRolesByApp = {};
    for (let a = 1; a < appRows.length; a++) {
      const appName = String(appRows[a][0] || '').trim();
      if (!appName) continue;
      let list = String(appRows[a][4] || 'user').split(',')
        .map(s => s.trim().toLowerCase()).filter(s => s);
      if (!list.length) list = ['user'];
      allowedRolesByApp[appName] = list;
    }

    // 同一 (user_id, app_name) が複数あれば最後の1件だけを採用
    const dedup = {};
    const order = [];
    changes.forEach(c => {
      const uid = String(c.user_id || '').trim();
      const appName = String(c.app_name || '').trim();
      if (!uid || !appName) return;
      const key = uid + '/' + appName;
      if (!dedup[key]) order.push(key);
      dedup[key] = {
        user_id:   uid,
        app_name:  appName,
        prev_role: normRole_(c.prev_role),
        next_role: normRole_(c.next_role) || 'none'
      };
    });

    // next_role の妥当性 ＋ 自己ロックアウト防止（1件でも違反があれば全体を中止）
    for (let oi = 0; oi < order.length; oi++) {
      const ch = dedup[order[oi]];
      const allowedRoles = allowedRolesByApp[ch.app_name];
      if (!allowedRoles) return { success: false, error: 'UNKNOWN_APP', detail: ch.app_name };

      const nextIsValid = ch.next_role === 'none' || allowedRoles.indexOf(ch.next_role) !== -1;
      if (!nextIsValid) {
        return { success: false, error: 'INVALID_ROLE', detail: ch.app_name + '=' + ch.next_role };
      }

      if (ch.app_name === 'auth-admin' && ch.user_id === actorUserId && ch.next_role !== 'admin') {
        return { success: false, error: 'SELF_LOCKOUT' };
      }
    }

    // 現在のシート値を読み、(user_id, app_name) → 行番号 の対応を作る（最初の一致のみ）
    const data = rolesSh.getDataRange().getValues(); // [user_id, app_name, role]・1行目ヘッダー
    const rowIndexByKey = {};
    const seen = {};
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim() + '/' + String(data[i][1] || '').trim();
      if (seen[key]) continue; // 重複行の2件目以降は対象にしない（アプリ側のbreakと同じ挙動）
      seen[key] = true;
      rowIndexByKey[key] = i + 1; // シート上の行番号（1始まり）
    }

    const conflicts = [];
    const toUpdate = []; // { row, role }
    const toAppend = [];  // [user_id, app_name, role]
    const auditEntries = []; // [user_id, app_name, before, after]

    order.forEach(key => {
      const ch = dedup[key];
      const rowNum = rowIndexByKey[key];
      const currentRole = rowNum ? normRole_(data[rowNum - 1][2]) : '';

      if (currentRole !== ch.prev_role) {
        conflicts.push({
          user_id: ch.user_id, app_name: ch.app_name,
          expected: ch.prev_role || 'none', actual: currentRole || 'none'
        });
        return;
      }
      if (currentRole === ch.next_role) return; // 実質変化なし

      if (rowNum) {
        toUpdate.push({ row: rowNum, role: ch.next_role });
      } else if (ch.next_role !== 'none') {
        toAppend.push([ch.user_id, ch.app_name, ch.next_role]);
      } else {
        return; // なし→なしは行を作らない
      }
      auditEntries.push([ch.user_id, ch.app_name, currentRole || '', ch.next_role]);
    });

    // 🔴 書き込みはC列だけ。対象行を個別に更新する（シート全体のsetValuesは使わない）
    toUpdate.forEach(u => rolesSh.getRange(u.row, 3, 1, 1).setValue(u.role));
    toAppend.forEach(row => rolesSh.appendRow(row));

    if (auditEntries.length) {
      _appendAudit(ss, actorUserId, auditEntries, requestId, APP_NAME + ' ' + VERSION);
    }

    const result = {
      success: true,
      applied: toUpdate.length + toAppend.length,
      conflicts: conflicts
    };
    cache.put(idemKey, JSON.stringify(result), CACHE_TTL_IDEMPOTENCY);
    return result;
  } catch (e) {
    Logger.log('_saveMatrix error: ' + e);
    return { success: false, error: 'INTERNAL_ERROR' };
  } finally {
    if (gotLock) lock.releaseLock();
  }
}

// ============================================================
// migrateRoles: 旧ロール語彙(manager/editor/staff/viewer)を新語彙へ一括置換する
// 既定はdry run。confirm==='MIGRATE' のときだけ実際に書き込む。
// ============================================================
function _migrateRoles(payload, actorUserId) {
  const requestId = String(payload.request_id || '').trim();
  if (!requestId) return { success: false, error: 'REQUEST_ID_REQUIRED' };

  const dryRun = payload.confirm !== 'MIGRATE';

  const cache   = CacheService.getScriptCache();
  const idemKey = 'migrateRoles_' + requestId;
  if (!dryRun) {
    const cached = cache.get(idemKey);
    if (cached !== null) {
      try { return JSON.parse(cached); } catch (e) {}
    }
  }

  let lock = null, gotLock = false;
  if (!dryRun) {
    lock = LockService.getScriptLock();
    try { gotLock = lock.tryLock(30000); } catch (e) {}
    if (!gotLock) return { success: false, error: 'BUSY' };
  }

  try {
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const rolesSh = ss.getSheetByName('user_app_roles');
    if (!rolesSh) return { success: false, error: 'INTERNAL_ERROR' };

    const data = rolesSh.getDataRange().getValues();
    const plan = [];
    const seen = {};
    for (let i = 1; i < data.length; i++) {
      const uid = String(data[i][0] || '').trim();
      const appName = String(data[i][1] || '').trim();
      if (!uid || !appName) continue;
      const key = uid + '/' + appName;
      if (seen[key]) continue; // 重複行は最初の1件だけを対象にする
      seen[key] = true;

      const raw = String(data[i][2] || '').trim().toLowerCase();
      const mapped = ROLE_ALIAS[raw];
      if (mapped && mapped !== raw) {
        plan.push({ row: i + 1, user_id: uid, app_name: appName, before: raw, after: mapped });
      }
    }

    if (dryRun) {
      return {
        success: true, dry_run: true,
        changes: plan.map(p => ({ user_id: p.user_id, app_name: p.app_name, before: p.before, after: p.after }))
      };
    }

    plan.forEach(p => rolesSh.getRange(p.row, 3, 1, 1).setValue(p.after));

    if (plan.length) {
      const auditEntries = plan.map(p => [p.user_id, p.app_name, p.before, p.after]);
      _appendAudit(ss, actorUserId, auditEntries, requestId, 'role-migration');
    }

    const result = { success: true, dry_run: false, applied: plan.length };
    cache.put(idemKey, JSON.stringify(result), CACHE_TTL_IDEMPOTENCY);
    return result;
  } catch (e) {
    Logger.log('_migrateRoles error: ' + e);
    return { success: false, error: 'INTERNAL_ERROR' };
  } finally {
    if (gotLock) lock.releaseLock();
  }
}

// ============================================================
// role_audit への追記（追記のみ・1回のsetValuesでまとめて書く）
// entries: [ [user_id, app_name, before, after], ... ]
// ============================================================
function _appendAudit(ss, actorUserId, entries, requestId, source) {
  const sh = ss.getSheetByName('role_audit');
  if (!sh || !entries.length) return;

  const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const rows = entries.map(e => [
    ts, actorUserId, e[0], e[1], e[2], e[3], source || (APP_NAME + ' ' + VERSION), requestId || ''
  ]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
}

// ============================================================
// ヘルパー
// ============================================================
function _json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// keepWarm: GASのコールドスタートを防ぐ定期実行用（他アプリと同型）
// ============================================================
function keepWarm() {
  // 何もしない（トリガーによる定期呼び出しでインスタンスをウォームアップするだけ）
}

// ============================================================
// activateAuthAdminTile: auth-admin 自身のタイルを有効化する（一度だけ手動実行）。
// GASエディタの関数選択で選んで実行するだけでよい（引数不要のラッパー）。
// 実行後、ポータルのApp_MASTER読込対応（Phase 5）が本番反映されればタイルが表示される。
// ============================================================
function activateAuthAdminTile() {
  setAppRow('auth-admin', 'https://beaufield.github.io/beaufield-dev/auth-admin/', 'active');
}

// ============================================================
// runMigrateRolesDryRun / runMigrateRolesReal: ロール語彙統一（Phase 3 Step 3）を
// GASエディタから直接実行する。Web経由のmigrateRolesアクションと同じ_migrateRoles()を
// 呼ぶが、セッション検証は行わない（GASエディタで実行できる時点で本人確認済みとみなす。
// setupAuthAdminSheets/activateAuthAdminTileと同じ考え方）。
//
// 使い方:
//   1. まず runMigrateRolesDryRun() を実行し、実行ログの変更予定一覧を確認する
//   2. 内容に問題がなければ runMigrateRolesReal() を実行する（実際に書き込まれる）
//   3. 実行ログ（表示 → 実行数）を確認する。「applied」の件数が予定件数と一致すればOK
// ============================================================
function runMigrateRolesDryRun() {
  const actor  = _allowedUserIds()[0] || 'gas-editor';
  const result = _migrateRoles({ request_id: Utilities.getUuid() }, actor);
  Logger.log(JSON.stringify(result, null, 2));
}

function runMigrateRolesReal() {
  const actor  = _allowedUserIds()[0] || 'gas-editor';
  const result = _migrateRoles({ request_id: Utilities.getUuid(), confirm: 'MIGRATE' }, actor);
  Logger.log(JSON.stringify(result, null, 2));
}

// ============================================================
// setAppRow: apps シートの1行（url・status）を app_name で更新する汎用ヘルパー。
// Phase 6 で saveApp アクションとして画面から使えるようにするまでの暫定版。
// ============================================================
function setAppRow(appName, url, status) {
  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const sh = ss.getSheetByName('apps');
  if (!sh) { Logger.log('🔴 apps シートがありません'); return; }

  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === appName) {
      sh.getRange(i + 1, 4, 1, 1).setValue(url);     // D列: url
      sh.getRange(i + 1, 8, 1, 1).setValue(status);  // H列: status
      Logger.log(appName + ' 行を更新しました: url=' + url + ' / status=' + status);
      return;
    }
  }
  Logger.log('🔴 該当行が見つかりません: ' + appName);
}

// ============================================================
// 初回セットアップ（一度だけ手動実行・ファイル冒頭のコメント参照）
// 何度実行しても安全（既にデータがあればヘッダーの確認だけで終わる）
// ============================================================
function setupAuthAdminSheets() {
  if (!AUTH_SHEET_ID) {
    Logger.log('🔴 NG: AUTH_SHEET_ID が未設定です。スクリプトプロパティを先に設定してください。');
    return;
  }
  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
  _setupAppsSheet(ss);
  _setupRoleAuditSheet(ss);
  _grantSelfAdminRole(ss);
  Logger.log('✅ auth-admin 初期セットアップ完了');
}

function _setupAppsSheet(ss) {
  let sheet = ss.getSheetByName('apps');
  if (!sheet) sheet = ss.insertSheet('apps');

  const header = ['app_name', 'label', 'icon', 'url', 'roles', 'role_labels', 'sort_order', 'status', 'note'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (sheet.getLastRow() > 1) {
    Logger.log('appsシートに既存データがあります（' + (sheet.getLastRow() - 1) + '件）。ヘッダーのみ確認しました。');
    return;
  }

  // 初期データの原本は portal/Code.gs の APP_MASTER（HANDOVER.mdの一覧表は古いURLを含むため使わない）。
  // auth-admin 自身は url 未確定・status='hidden' で登録し、Phase 5（ポータル改修）完了後に
  // 実際のGAS URLを埋めて status='active' に切り替える。
  const rows = [
    ['order-app',         '発注アプリ',              '📦', 'https://beaufield.github.io/beaufield-dev/order-app/',                      'user,admin', '',                           10, 'active', ''],
    ['lending',           '貸出管理',                '🔑', 'https://beaufield.github.io/beaufield-dev/kiki-kanri/',                      'user,admin', '',                           20, 'active', ''],
    ['serial-apps',       'シリアルNo管理',           '🏷️', 'https://beaufield.github.io/beaufield-dev/serial-apps/',                     'user,admin', '',                           30, 'active', ''],
    ['yoyaku-kanri',      '予約管理',                '📋', 'https://beaufield.github.io/beaufield-dev/yoyaku-kanri/',                    'user,admin', 'admin=事務,user=営業',       40, 'active', ''],
    ['bcart-master',      'BCARTマスター管理',        '🛒', 'https://beaufield.github.io/beaufield-dev/bcart-integration/master-tool/',   'user,admin', 'admin=編集可,user=閲覧のみ', 50, 'active', ''],
    ['bcart-orders',      'Bカート受注確認',          '🧾', 'https://beaufield.github.io/beaufield-dev/bcart-integration/order-viewer/',  'user,admin', '',                           60, 'active', ''],
    ['expense-approval',  '経費事前申請',             '💰', 'https://beaufield.github.io/beaufield-dev/expense-approval/',                'user,admin', '',                           70, 'active', ''],
    ['project-dashboard', 'プロジェクトダッシュボード', '📊', 'https://beaufield.github.io/beaufield-dev/project-dashboard/',              'user,admin', '',                           80, 'active', ''],
    ['line-progress',     'LINE登録進捗',             '📈', 'https://beaufield.github.io/beaufield-dev/line-progress/',                   'user,admin', '',                           90, 'active', ''],
    ['route-checker',     'ルート訪問チェッカー',      '🗺️', 'https://beaufield.github.io/beaufield-dev/route-checker/',                  'user',       '',                          100, 'hidden', '当面不使用（portal v1.11.1で一覧から削除・直リンクは有効）'],
    ['beaufes',           'ビューフェス（社員用）',    '🎪', 'https://beaufield.github.io/beaufield-dev/beaufes/badges.html',              'user',       '',                          110, 'hidden', '社員用の入口はbadges.html。index.htmlは公開申込フォーム'],
    ['auth-admin',        '権限管理',                '🔐', '',                                                                             'admin',      '',                          120, 'hidden', 'Phase5完了後にurlを設定しstatusをactiveへ']
  ];
  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);

  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#d9e1f2');
  sheet.autoResizeColumns(1, header.length);
  Logger.log('appsシート: ヘッダー + 初期データ（' + rows.length + '件）を設定しました');
}

function _setupRoleAuditSheet(ss) {
  let sheet = ss.getSheetByName('role_audit');
  if (!sheet) sheet = ss.insertSheet('role_audit');

  const header = ['timestamp', 'actor_user_id', 'target_user_id', 'app_name', 'before', 'after', 'source', 'request_id'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#d9e1f2');
  sheet.autoResizeColumns(1, header.length);
  Logger.log('role_auditシート: ヘッダーを設定しました（既存データは変更していません）');
}

// ALLOWED_USER_IDS の各ユーザーへ auth-admin の admin ロールを付与する（冪等）
function _grantSelfAdminRole(ss) {
  const allowed = _allowedUserIds();
  if (!allowed.length) {
    Logger.log('🔴 ALLOWED_USER_IDS が未設定のため、admin権限の付与をスキップしました。');
    return;
  }

  const rolesSh = ss.getSheetByName('user_app_roles');
  const data = rolesSh.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === APP_NAME) {
      existing[String(data[i][0]).trim()] = i + 1; // 行番号
    }
  }

  let added = 0, updated = 0;
  allowed.forEach(userId => {
    const rowNum = existing[userId];
    if (rowNum) {
      const current = normRole_(data[rowNum - 1][2]);
      if (current !== 'admin') {
        rolesSh.getRange(rowNum, 3, 1, 1).setValue('admin');
        updated++;
      }
    } else {
      rolesSh.appendRow([userId, APP_NAME, 'admin']);
      added++;
    }
  });
  Logger.log('auth-adminのadmin権限付与: 新規' + added + '件・更新' + updated + '件（対象' + allowed.length + '名）');
}
