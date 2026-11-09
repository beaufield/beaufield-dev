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
const VERSION        = 'v1.4.1';
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
    if (!body || typeof body !== 'object' || Array.isArray(body)) return _json({ success: false, error: 'PARAM_LOST' });
  } catch (err) {
    return _json({ success: false, error: 'PARAM_LOST' });
  }

  const action = String(body.action || '');
  const token  = String(body.session_token || '');

  let auth = _validateAdmin(token, action !== 'getMatrix');
  if (!auth.ok && auth.transient) {
    // 🔴 誤ログアウト対策: 正常な200でSESSION_INVALIDが返る既知の持病があるため、
    //    一時障害シグナルが立った場合だけ1回だけ再検証してから拒否を確定する。
    auth = _validateAdmin(token, action !== 'getMatrix');
  }
  if (!auth.ok) {
    return _json({ success: false, error: auth.error || 'SESSION_INVALID' });
  }

  try {
    switch (action) {
      case 'getMatrix':    return _json(_getMatrix());
      case 'saveAppOrder': return _json(_saveAppOrder(body, auth.user_id));
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
function _validateAdmin(token, fresh) {
  if (!token) return { ok: false, error: 'SESSION_INVALID' };

  const cache    = CacheService.getScriptCache();
  const cacheKey = 'sess_authadmin_v1_' + token.slice(-32);
  const cached   = cache.get(cacheKey);
  if (!fresh && cached !== null) {
    try { const r = JSON.parse(cached); if (r.expires_at > Date.now()) return r; } catch (e) {}
  }

  if (!AUTH_SHEET_ID) return { ok: false, error: 'SESSION_INVALID', transient: true };

  try {
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

    const sessSh = ss.getSheetByName('sessions');
    if (!sessSh) return { ok: false, error: 'SESSION_INVALID', transient: true };

    const sessions = sessSh.getDataRange().getValues();
    const now = Date.now();
    let userId = null, expiresAt = 0;
    for (let i = 1; i < sessions.length; i++) {
      if (String(sessions[i][0]) === token) {
        if (Number(sessions[i][2]) > now) { userId = String(sessions[i][1]); expiresAt = Number(sessions[i][2]); }
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
      ? { ok: true, user_id: userId, name: String(userRow[1] || userId), expires_at: expiresAt }
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

// マスターの解釈は読取・保存で共有する。不明な設定をuserへ補完しない。
function _fail(code, detail) { return { success: false, error: code, detail: detail || '' }; }
function _id(value) { return typeof value === 'string' && /^[A-Za-z0-9_-]{1,80}$/.test(value); }
function _true(value) { return value === true || value === 'TRUE'; }
function _hash(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function _rows(ss, name, columns) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('SCHEMA_INVALID:' + name);
  const rows = sheet.getDataRange().getValues();
  Object.keys(columns).forEach(col => {
    if (String((rows[0] || [])[Number(col)] || '').trim() !== columns[col]) throw new Error('SCHEMA_INVALID:' + name);
  });
  return rows;
}
function _appDefinitions(rows) {
  const apps = [], byId = Object.create(null), issues = [];
  rows.slice(1).forEach((r, i) => {
    const id = String(r[0] || '').trim();
    if (!id) {
      if (r.some(v => v !== '' && v != null)) issues.push('appsの' + (i + 2) + '行目にアプリIDがありません');
      return;
    }
    const errors = [];
    if (!_id(id)) errors.push('アプリIDの形式');
    const status = String(r[7] || '').trim();
    if (!['active', 'hidden', 'archived'].includes(status)) errors.push('公開状態');
    const rawRoles = String(r[4] || '').trim();
    const roles = rawRoles.split(',').map(v => v.trim().toLowerCase());
    if (rawRoles && (roles.some(v => !KNOWN_ROLES.includes(v)) || new Set(roles).size !== roles.length)) errors.push('選べる権限');
    const labels = Object.create(null);
    const rawLabels = String(r[5] || '').trim();
    if (rawLabels) rawLabels.split(',').forEach(pair => {
      const parts = pair.split('=');
      const key = parts[0].trim().toLowerCase();
      if (parts.length !== 2 || !roles.includes(key) || !parts[1].trim() || labels[key]) errors.push('権限の表示名');
      else labels[key] = parts[1].trim();
    });
    const sort = Number(r[6]);
    if (r[6] === '' || !Number.isFinite(sort) || !Number.isInteger(sort) || sort < 0) errors.push('並び順');
    const app = { app_name: id, label: String(r[1] || id), icon: String(r[2] || '📱'), url: String(r[3] || ''),
      roles: rawRoles ? roles.filter(v => KNOWN_ROLES.includes(v)) : [], row_key: 'apps-row-' + (i + 2), role_labels: labels, status: status,
      sort_order: Number.isFinite(sort) ? sort : 999, row: i + 2, raw_sort: r[6], raw_roles: rawRoles, config_errors: errors };
    if (!byId[id]) byId[id] = app;
    apps.push(app);
  });
  // 発注・在庫の新旧リンクは、既存運用で同じ権限IDを共有している。
  // URLを限定した既存の2組だけを認め、並び順は行キーで別々に扱う。
  const linkedPaths = {
    'order-app': ['/order-app/', '/order-app/db/'],
    'stock-report': ['/stock-report/', '/stock-report/db/']
  };
  Object.keys(byId).forEach(id => {
    const group = apps.filter(a => a.app_name === id);
    if (group.length > 1) {
      const paths = linkedPaths[id];
      const urls = paths && paths.map(p => 'https://beaufield.github.io/beaufield-dev' + p);
      const isKnownPair = urls && group.length === 2 && urls.every(url => group.filter(a => a.url === url).length === 1)
        && ['hidden', 'archived'].includes(group.find(a => a.url === urls[0]).status)
        && ['active', 'hidden'].includes(group.find(a => a.url === urls[1]).status);
      if (!isKnownPair) group.forEach(a => a.config_errors.push('アプリIDの重複'));
      else {
        const explicit = group.filter(a => a.raw_roles);
        const choices = explicit[0] && explicit[0].roles;
        const sameChoices = choices && explicit.every(a => JSON.stringify(a.roles.slice().sort()) === JSON.stringify(choices.slice().sort()));
        if (!sameChoices) group.forEach(a => a.config_errors.push('共有する権限の設定不一致'));
        else {
          group.forEach(a => { if (!a.raw_roles) a.roles = choices.slice(); a.linked_roles = true; });
          byId[id] = group.find(a => a.url === urls[1]);
        }
      }
    }
    group.forEach(a => {
      if (!a.roles.length) a.config_errors.push('選べる権限');
      if (group.some(g => g.config_errors.length) && !a.config_errors.length) a.config_errors.push('共有先の設定不備');
    });
  });
  apps.forEach(a => { if (a.config_errors.length) issues.push(a.label + ': ' + a.config_errors.join('・')); });
  return { apps: apps, byId: byId, issues: issues };
}
function _readModel(ss) {
  const users = _rows(ss, 'users', { 0: 'user_id', 3: 'active', 5: 'is_admin' });
  const appRows = _rows(ss, 'apps', { 0: 'app_name', 4: 'roles', 5: 'role_labels', 6: 'sort_order', 7: 'status' });
  const roleRows = _rows(ss, 'user_app_roles', { 0: 'user_id', 1: 'app_name', 2: 'role' });
  const definitions = _appDefinitions(appRows);
  const userMap = Object.create(null), userList = [], issues = definitions.issues.slice();
  users.slice(1).forEach(r => {
    const id = String(r[0] || '').trim();
    if (!id) return;
    const user = { user_id: id, name: String(r[1] || id), short_name: String(r[7] || r[1] || id), active: _true(r[3]), is_admin: _true(r[5]) };
    if (!_id(id)) issues.push('社員IDの形式が不正です');
    if (userMap[id]) { userMap[id].duplicate = true; user.duplicate = true; issues.push('社員IDが重複しています: ' + id); }
    else userMap[id] = user;
    userList.push(user);
  });
  const pairs = Object.create(null);
  roleRows.slice(1).forEach((r, i) => {
    const uid = String(r[0] || '').trim(), app = String(r[1] || '').trim();
    if (!uid || !app) return;
    const key = uid + '/' + app;
    if (pairs[key]) pairs[key].count++;
    else pairs[key] = { row: i + 2, role: normRole_(r[2]), count: 1, user_id: uid, app_name: app };
  });
  const visible = definitions.apps.filter(a => a.status !== 'archived')
    .sort((a, b) => a.sort_order - b.sort_order || a.row - b.row);
  // 追加・廃止・同順位・元セルの型・行の入替も含め、取得時のマスターを照合する。
  const revision = _hash(JSON.stringify(appRows));
  return { apps: visible, definitions: definitions, users: userList, userMap: userMap, pairs: pairs,
    issues: issues, appRows: appRows, revision: revision };
}

function _getMatrix() {
  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const m = _readModel(ss), roles = Object.create(null), legacy = [], outside = [], duplicates = [];
  const orphanApps = new Set(), orphanUsers = new Set(), archived = new Set();
  Object.keys(m.pairs).forEach(key => {
    const p = m.pairs[key], app = m.definitions.byId[p.app_name];
    if (p.count > 1) duplicates.push({ user_id: p.user_id, app_name: p.app_name, count: p.count });
    if (p.role) {
      if (!roles[p.user_id]) roles[p.user_id] = Object.create(null);
      roles[p.user_id][p.app_name] = p.role;
      if (!KNOWN_ROLES.includes(p.role)) legacy.push({ user_id: p.user_id, app_name: p.app_name, role: p.role });
      else if (app && app.status !== 'archived' && !app.roles.includes(p.role)) outside.push({ user_id: p.user_id, app_name: p.app_name, role: p.role, allowed: app.roles });
    }
    if (!app) orphanApps.add(p.app_name);
    else if (app.status === 'archived' && p.role) archived.add(p.app_name);
    if (!m.userMap[p.user_id]) orphanUsers.add(p.user_id);
  });
  return { success: true, users: m.users, apps: m.apps, roles: roles, legacy_roles: legacy, out_of_vocab: outside,
    duplicate_rows: duplicates, orphan_apps: Array.from(orphanApps), orphan_users: Array.from(orphanUsers),
    archived_roles: Array.from(archived), configuration_errors: m.issues, order_revision: m.revision, fetched_at: Date.now(), version: VERSION };
}

// 書込は事前検証→試行監査→本体→完了監査の順。途中失敗は成功と扱わない。
// 要求の記録を端末キャッシュに依存せず保持し、不明結果の盲目再実行を防ぐ。
function _runAdminMutation(kind, payload, actor, content, build) {
  if (!_id(payload.request_id) || payload.request_id.length < 8) return _fail('REQUEST_ID_REQUIRED');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return _fail('BUSY');
  let started = false, journalKey = '', fingerprint = '';
  try {
    journalKey = 'auth_request_' + _hash(actor + '/' + kind + '/' + payload.request_id);
    fingerprint = _hash(JSON.stringify(content));
    const previous = _PROPS.getProperty(journalKey);
    if (previous) {
      const record = JSON.parse(previous);
      if (record.fingerprint !== fingerprint) return _fail('REQUEST_REUSED');
      return record.status === 'done' ? record.result : _fail('SAVE_UNCERTAIN');
    }
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const m = _readModel(ss);
    const plan = build(ss, m);
    if (plan.success === false) return plan;
    _rows(ss, 'role_audit', { 0: 'timestamp', 1: 'actor_user_id', 2: 'target_user_id', 3: 'app_name', 4: 'before', 5: 'after', 6: 'source', 7: 'request_id' });
    // 完了済みの記録だけ期限付きで整理。不明な要求は自動で捨てない。
    const all = _PROPS.getProperties();
    Object.keys(all).filter(k => k.indexOf('auth_request_') === 0).forEach(k => {
      try { const r = JSON.parse(all[k]); if (r.status === 'done' && Date.now() - r.at > 7 * 86400000) _PROPS.deleteProperty(k); } catch (_) {}
    });
    const record = { fingerprint: fingerprint, status: 'started', at: Date.now() };
    _PROPS.setProperty(journalKey, JSON.stringify(record));
    started = true;
    _appendAudit(ss, actor, plan.audit, payload.request_id, kind + ':開始 ' + VERSION);
    plan.write();
    SpreadsheetApp.flush();
    _appendAudit(ss, actor, plan.audit, payload.request_id, kind + ':完了 ' + VERSION);
    const result = { success: true, applied: plan.audit.length, conflicts: [] };
    _PROPS.setProperty(journalKey, JSON.stringify({ fingerprint: fingerprint, status: 'done', at: Date.now(), result: result }));
    return result;
  } catch (e) {
    Logger.log('管理保存失敗: ' + kind);
    if (started) return _fail('SAVE_UNCERTAIN');
    return _fail(String(e.message || '').indexOf('SCHEMA_INVALID:') === 0 ? 'SCHEMA_INVALID' : 'INTERNAL_ERROR');
  } finally { lock.releaseLock(); }
}

function _saveMatrix(payload, actor) {
  const changes = payload.changes;
  if (!Array.isArray(changes) || !changes.length || changes.length > 300) return _fail('INVALID_CHANGES');
  return _runAdminMutation('saveMatrix', payload, actor, changes, (ss, m) => {
    const seen = new Set(), writes = [], audit = [];
    for (const c of changes) {
      if (!c || !_id(c.user_id) || !_id(c.app_name) || typeof c.prev_role !== 'string' || typeof c.next_role !== 'string') return _fail('INVALID_CHANGES');
      const key = c.user_id + '/' + c.app_name;
      if (seen.has(key)) return _fail('DUPLICATE_CHANGE', key);
      seen.add(key);
      const user = m.userMap[c.user_id], app = m.definitions.byId[c.app_name], pair = m.pairs[key];
      const next = normRole_(c.next_role), prev = normRole_(c.prev_role);
      if (!user) return _fail('UNKNOWN_USER', c.user_id);
      if (user.duplicate) return _fail('DUPLICATE_USER', c.user_id);
      if (!app) return _fail('UNKNOWN_APP', c.app_name);
      if (app.status === 'archived') return _fail('ARCHIVED_APP', c.app_name);
      if (app.config_errors.length) return _fail('APP_CONFIG_INVALID', c.app_name);
      if (pair && pair.count > 1) return _fail('DUPLICATE_ROLE', key);
      if (next && !app.roles.includes(next)) return _fail('INVALID_ROLE', c.app_name);
      if (!user.active && next) return _fail('INACTIVE_USER', c.user_id);
      if (c.app_name === APP_NAME && c.user_id === actor && next !== 'admin') return _fail('SELF_LOCKOUT');
      const current = pair ? pair.role : '';
      if (current !== prev) return _fail('CONFLICT', key);
      if (current === next) continue;
      audit.push([c.user_id, c.app_name, current || 'none', next || 'none']);
      writes.push({ row: pair ? pair.row : 0, user_id: c.user_id, app_name: c.app_name, role: next || 'none' });
    }
    return { audit: audit, write: () => {
      const sh = ss.getSheetByName('user_app_roles');
      writes.forEach(w => {
        if (w.row) sh.getRange(w.row, 3).setValue(w.role);
        else sh.appendRow([w.user_id, w.app_name, w.role]);
      });
    } };
  });
}

function _saveAppOrder(payload, actor) {
  if (!Array.isArray(payload.order) || !payload.order.length || payload.order.length > 300 || typeof payload.base_revision !== 'string') return _fail('INVALID_ORDER');
  return _runAdminMutation('saveAppOrder', payload, actor, [payload.base_revision, payload.order], (ss, m) => {
    if (m.issues.length) return _fail('APP_CONFIG_INVALID');
    if (m.revision !== payload.base_revision) return _fail('ORDER_CONFLICT');
    if (payload.order.length !== m.apps.length || new Set(payload.order).size !== payload.order.length || payload.order.some(id => !_id(id) || !m.apps.some(a => a.row_key === id))) return _fail('INVALID_ORDER');
    const audit = [], writes = [];
    payload.order.forEach((id, i) => {
      const app = m.apps.find(a => a.row_key === id), next = (i + 1) * 10;
      if (app.raw_sort !== next) { audit.push(['', app.app_name + ':' + app.row_key, String(app.raw_sort), String(next)]); writes.push([app.row, next]); }
    });
    return { audit: audit, write: () => {
      const sh = ss.getSheetByName('apps');
      // 物理行の入替はしない。対象アプリのG列だけを書き換える。
      writes.forEach(w => sh.getRange(w[0], 7).setValue(w[1]));
    } };
  });
}

// 警告原因のマスター整合専用。割当値は読取照合のみで、更新対象はapps.E列。
// 本番承認後にエディタからpreview→applyの順で実行する。自動起動しない。
function previewBeaufesRoleOptions() {
  const result = _alignBeaufesRoleOptions(false);
  Logger.log(JSON.stringify(result)); return result;
}
function applyBeaufesRoleOptions() {
  const result = _alignBeaufesRoleOptions(true);
  Logger.log(JSON.stringify(result)); return result;
}
function _alignBeaufesRoleOptions(apply) {
  const inspect = m => {
    const app = m.definitions.byId.beaufes;
    if (!app || app.config_errors.length || app.status === 'archived') return _fail('APP_CONFIG_INVALID');
    if (app.raw_roles === 'user,admin') return { success: true, changed: false };
    if (app.raw_roles !== 'user') return _fail('CONFIG_CONFLICT');
    return { success: true, changed: true, before: 'user', after: 'user,admin', row: app.row };
  };
  if (!apply) return inspect(_readModel(SpreadsheetApp.openById(AUTH_SHEET_ID)));
  const actor = _allowedUserIds()[0];
  if (!actor) return _fail('FORBIDDEN');
  return _runAdminMutation('alignBeaufesRoles', { request_id: 'beaufes-options-user-admin-20261109' }, actor, ['beaufes', 'user', 'user,admin'], (ss, m) => {
    const p = inspect(m);
    if (!p.success) return p;
    return { audit: p.changed ? [['', 'beaufes', 'roles:user', 'roles:user,admin']] : [],
      write: () => { if (p.changed) ss.getSheetByName('apps').getRange(p.row, 5).setValue('user,admin'); } };
  });
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

    // 旧APIからも同じ検証を使う。外側のロックは一度解放して再取得する。
    if (gotLock) { lock.releaseLock(); gotLock = false; }
    if (!plan.length) return { success: true, dry_run: false, applied: 0 };
    const result = _saveMatrix({ request_id: requestId, changes: plan.map(p => ({
      user_id: p.user_id, app_name: p.app_name, prev_role: p.before, next_role: p.after
    })) }, actorUserId);
    return Object.assign({ dry_run: false }, result);
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
  if (!sh) throw new Error('SCHEMA_INVALID:role_audit');
  if (!entries.length) return;

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
    ['beaufes',           'ビューフェス（社員用）',    '🎪', 'https://beaufield.github.io/beaufield-dev/beaufes/badges.html',              'user,admin', '',                          110, 'hidden', '社員用の入口はbadges.html。index.htmlは公開申込フォーム'],
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
