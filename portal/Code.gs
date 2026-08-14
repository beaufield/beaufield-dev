// ============================================================
// Beaufield ポータル - Google Apps Script
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   AUTH_SHEET_ID : beaufield-auth スプレッドシートID（共通）
//
// ============================================================

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS        = PropertiesService.getScriptProperties();
const VERSION       = 'v1.9.1';
const AUTH_SHEET_ID = _PROPS.getProperty('AUTH_SHEET_ID');

// ロックアウト設定
const MAX_ATTEMPTS = 5;
const LOCK_MINUTES = 10;

// セッション有効期間。共有端末でのトークン残存を抑えるため業務日内に限定する。
const SESSION_HOURS = 12;

// ============================================================
// アプリマスター
// appName は beaufield-auth の user_app_roles シートの値と一致させる
// ============================================================
const APP_MASTER = [
  {
    appName: 'order-app',
    label:   '発注アプリ',
    icon:    '📦',
    url:     'https://beaufield.github.io/beaufield-dev/order-app/'
  },
  {
    appName: 'route-checker',
    label:   'ルート訪問チェッカー',
    icon:    '🗺️',
    url:     'https://beaufield.github.io/beaufield-dev/route-checker/'
  },
  {
    appName: 'lending',
    label:   '貸出管理',
    icon:    '🔑',
    url:     'https://beaufield.github.io/beaufield-dev/kiki-kanri/'
  },
  {
    appName: 'serial-apps',
    label:   'シリアルNo管理',
    icon:    '🏷️',
    url:     'https://beaufield.github.io/beaufield-dev/serial-apps/'
  },
  {
    appName: 'yoyaku-kanri',
    label:   '予約管理',
    icon:    '📋',
    url:     'https://beaufield.github.io/beaufield-dev/yoyaku-kanri/'
  },
  {
    appName: 'bcart-master',
    label:   'BCARTマスター管理',
    icon:    '🛒',
    url:     'https://beaufield.github.io/beaufield-dev/bcart-integration/master-tool/'
  },
  {
    appName: 'bcart-orders',
    label:   'Bカート受注確認',
    icon:    '🧾',
    url:     'https://beaufield.github.io/beaufield-dev/bcart-integration/order-viewer/'
  },
  {
    appName: 'expense-approval',
    label:   '経費事前申請',
    icon:    '💰',
    url:     'https://beaufield.github.io/beaufield-dev/expense-approval/'
  },
  {
    appName: 'project-dashboard',
    label:   'プロジェクトダッシュボード',
    icon:    '📊',
    url:     'https://beaufield.github.io/beaufield-dev/project-dashboard/'
  },
  {
    appName: 'line-progress',
    label:   'LINE登録進捗',
    icon:    '📈',
    url:     'https://beaufield.github.io/beaufield-dev/line-progress/'
  }
];

// ============================================================
// エントリーポイント（GET）
// ============================================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
  const token  = e && e.parameter && e.parameter.session_token ? e.parameter.session_token : '';

  try {
    switch (action) {
      case 'getUsers':    return _json(getUsers());
      case 'getUserApps': return _json(getUserApps(token));
      default:            return _json({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return _json({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// エントリーポイント（POST）
// URL-encoded と JSON body の両方に対応
// ============================================================
function doPost(e) {
  let action = '', data = {};

  try {
    // JSON bodyの場合（他のGASアプリからの内部呼び出し）
    if (e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      if (body.action) { action = body.action; data = body; }
    }
  } catch(err) {}

  // URL-encodedの場合（ポータルHTMLからの呼び出し）
  if (!action && e && e.parameter) {
    action = e.parameter.action || '';
    try {
      data = e.parameter.data ? JSON.parse(e.parameter.data) : {};
    } catch(jsonErr) {
      return _json({ success: false, error: 'INVALID_REQUEST' });
    }
  }

  try {
    switch (action) {
      case 'login':           return _json(login(data));
      case 'resetPin':        return _json(resetPin(data));
      case 'changePin':       return _json(changePin(data));
      case 'validateSession': return _json(validateSession(data));
      case 'getUserApps':     return _json(getUserApps(data.session_token || ''));
      default:                return _json({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return _json({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// ユーザー一覧取得（ログイン画面のグリッド表示用）
// ============================================================
function getUsers() {
  const ss   = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const rows = ss.getSheetByName('users').getDataRange().getValues();
  const users = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[3] === true || row[3] === 'TRUE') {
      users.push({
        user_id:  String(row[0]),
        name:     String(row[1])
      });
    }
  }
  return { success: true, users };
}

// ============================================================
// ログイン処理（ロックアウト付き）
// ============================================================
function login(data) {
  const { user_id, pin } = data;
  if (!user_id || pin === undefined || pin === null || pin === '') {
    return { success: false, message: 'user_idとpinは必須です' };
  }

  // ── ロックアウトチェック ──────────────────────────────────
  const props    = PropertiesService.getScriptProperties();
  const lockKey  = 'lockout_' + user_id;
  const lockData = JSON.parse(props.getProperty(lockKey) || '{"count":0,"until":0}');
  const now      = Date.now();

  if (lockData.until > now) {
    const remaining = Math.ceil((lockData.until - now) / 60000);
    return {
      success: false,
      message: 'PINの誤入力が' + MAX_ATTEMPTS + '回に達しました。' + remaining + '分後に再試行してください。'
    };
  }
  // ─────────────────────────────────────────────────────────

  const ss     = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const rows   = ss.getSheetByName('users').getDataRange().getValues();
  const pinStr = String(pin).padStart(4, '0');

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[0]) === user_id && (row[3] === true || row[3] === 'TRUE')) {
      if (String(row[2]).padStart(4, '0') === pinStr) {
        // ログイン成功 → ロックカウントをリセット
        props.deleteProperty(lockKey);

        // ── セッショントークン発行 ────────────────────────────
        const token     = Utilities.getUuid();
        const expiresAt = now + SESSION_HOURS * 60 * 60 * 1000;
        _saveSession(ss, token, String(row[0]), expiresAt);
        // ────────────────────────────────────────────────────

        // is_admin: F列（row[5]）がTRUEかどうか
        const isAdmin = row[5] === true || row[5] === 'TRUE';

        return {
          success:       true,
          user_id:       String(row[0]),
          name:          String(row[1]),
          session_token: token,
          is_admin:      isAdmin
        };
      } else {
        // PIN不一致 → 失敗カウントを記録
        lockData.count = (lockData.count || 0) + 1;
        if (lockData.count >= MAX_ATTEMPTS) {
          lockData.until = now + LOCK_MINUTES * 60 * 1000;
          lockData.count = 0;
          props.setProperty(lockKey, JSON.stringify(lockData));
          return {
            success: false,
            message: 'PINの誤入力が' + MAX_ATTEMPTS + '回に達しました。' + LOCK_MINUTES + '分間ロックされます。'
          };
        }
        props.setProperty(lockKey, JSON.stringify(lockData));
        const left = MAX_ATTEMPTS - lockData.count;
        return { success: false, message: 'PINが正しくありません（残り' + left + '回）' };
      }
    }
  }
  return { success: false, message: 'ユーザーが見つかりません' };
}

// ============================================================
// PINリセット（管理者専用）
// ============================================================
function resetPin(data) {
  const { session_token, target_user_id, new_pin } = data;

  // 必須チェック
  if (!session_token || !target_user_id || !new_pin) {
    return { success: false, message: '必須パラメータが不足しています' };
  }

  // PIN形式チェック
  const pinStr = String(new_pin).padStart(4, '0');
  if (!/^\d{4}$/.test(pinStr)) {
    return { success: false, message: 'PINは4桁の数字で入力してください' };
  }

  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

  // セッション検証（誰が操作しているか）
  const requestUserId = _getSessionUser(ss, session_token);
  if (!requestUserId) {
    return { success: false, message: 'セッションが無効です。再ログインしてください' };
  }

  // 管理者権限チェック
  if (!_isAdmin(ss, requestUserId)) {
    return { success: false, message: '管理者権限がありません' };
  }

  // 対象ユーザーのPINを更新
  const sh   = ss.getSheetByName('users');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === target_user_id) {
      sh.getRange(i + 1, 3).setValue(pinStr); // C列（PIN）を更新
      // ロックアウトも同時に解除
      PropertiesService.getScriptProperties().deleteProperty('lockout_' + target_user_id);
      // 対象ユーザーの既存セッションをすべて削除（旧PINでのトークンを無効化）
      _deleteUserSessions(ss, target_user_id);
      return { success: true, message: 'PINを更新しました' };
    }
  }
  return { success: false, message: '対象ユーザーが見つかりません' };
}

// ============================================================
// PIN変更（本人による変更）
// ============================================================
function changePin(data) {
  const { session_token, current_pin, new_pin } = data;

  if (!session_token || current_pin === undefined || current_pin === null || !new_pin) {
    return { success: false, message: '必須パラメータが不足しています' };
  }

  const newPinStr = String(new_pin).padStart(4, '0');
  if (!/^\d{4}$/.test(newPinStr)) {
    return { success: false, message: 'PINは4桁の数字で入力してください' };
  }

  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

  // セッション検証
  const userId = _getSessionUser(ss, session_token);
  if (!userId) {
    return { success: false, message: 'セッションが無効です。再ログインしてください' };
  }

  // 現在のPINと照合してから更新
  const sh            = ss.getSheetByName('users');
  const rows          = sh.getDataRange().getValues();
  const currentPinStr = String(current_pin).padStart(4, '0');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === userId) {
      if (String(rows[i][2]).padStart(4, '0') !== currentPinStr) {
        return { success: false, message: '現在のPINが正しくありません' };
      }
      sh.getRange(i + 1, 3).setValue(newPinStr);
      return { success: true, message: 'PINを変更しました' };
    }
  }
  return { success: false, message: 'ユーザーが見つかりません' };
}

// ============================================================
// アクセス可能アプリ一覧取得（セッション必須）
// ============================================================
function getUserApps(token) {
  if (!token) return { success: false, error: 'SESSION_INVALID' };

  const ss     = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const userId = _getSessionUser(ss, token);
  if (!userId) return { success: false, error: 'SESSION_INVALID' };

  const roles = ss.getSheetByName('user_app_roles').getDataRange().getValues();

  const accessMap = {};
  for (let i = 1; i < roles.length; i++) {
    if (String(roles[i][0]) === userId && roles[i][2] !== 'none') {
      accessMap[String(roles[i][1])] = String(roles[i][2]);
    }
  }

  const apps = APP_MASTER
    .filter(app => accessMap[app.appName])
    .map(app => ({
      appName: app.appName,
      label:   app.label,
      icon:    app.icon,
      url:     app.url,
      role:    accessMap[app.appName]
    }));

  return { success: true, apps };
}

// ============================================================
// セッション検証（他のGASアプリからの内部呼び出し用）
// 戻り値: { ok: true/false } ※ success ではなく ok で返す
// ============================================================
function validateSession(data) {
  const token = data.token || '';
  const appName = String(data.app_name || '').trim();
  if (!token) return { ok: false };

  const ss     = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const userId = _getSessionUser(ss, token);
  if (!userId) return { ok: false };

  // ユーザー名・有効状態・管理者状態を users シートから取得する。
  // セッション発行後に利用停止されたユーザーも、ここで必ず拒否する。
  const rows = ss.getSheetByName('users').getDataRange().getValues();
  let userName = userId;
  let isAdmin = false;
  let activeUser = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === userId) {
      userName = String(rows[i][1]) || userId;
      activeUser = rows[i][3] === true || rows[i][3] === 'TRUE';
      isAdmin = rows[i][5] === true || rows[i][5] === 'TRUE';
      break;
    }
  }
  if (!activeUser) return { ok: false };

  // 呼び出し側がアプリ名を指定した場合は、ポータル表示だけでなくAPI側でも
  // user_app_rolesを検証する。未登録・noneは明示的に拒否する。
  let role = '';
  if (appName) {
    const roleRows = ss.getSheetByName('user_app_roles').getDataRange().getValues();
    for (let i = 1; i < roleRows.length; i++) {
      if (String(roleRows[i][0]) === userId && String(roleRows[i][1]) === appName) {
        role = String(roleRows[i][2] || '');
        break;
      }
    }
    if (!role || role === 'none') return { ok: false };
  }

  return { ok: true, user_id: userId, name: userName, is_admin: isAdmin, role: role };
}

// ============================================================
// セッション保存（beaufield-auth の sessions シート）
// ============================================================
function _saveSession(ss, token, user_id, expiresAt) {
  let sh = ss.getSheetByName('sessions');
  if (!sh) {
    // sessions シートが未作成なら自動作成
    sh = ss.insertSheet('sessions');
    sh.appendRow(['token', 'user_id', 'expires_at']);
  }

  // 期限切れセッションを削除（遅延クリーンアップ）
  const data = sh.getDataRange().getValues();
  const now  = Date.now();
  for (let i = data.length - 1; i >= 1; i--) {
    if (Number(data[i][2]) < now) {
      sh.deleteRow(i + 1);
    }
  }

  // 新しいセッションを追記
  sh.appendRow([token, user_id, expiresAt]);
}

// ============================================================
// セッショントークンからuser_idを取得
// ============================================================
function _getSessionUser(ss, token) {
  const sh = ss.getSheetByName('sessions');
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const now  = Date.now();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token && Number(data[i][2]) > now) {
      return String(data[i][1]); // user_id
    }
  }
  return null;
}

// ============================================================
// 対象ユーザーのセッションを全削除（PINリセット・無効化時に使用）
// ============================================================
function _deleteUserSessions(ss, userId) {
  const sh = ss.getSheetByName('sessions');
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === userId) sh.deleteRow(i + 1);
  }
}

// ============================================================
// 管理者チェック（usersシートのF列 is_admin）
// ============================================================
function _isAdmin(ss, user_id) {
  const rows = ss.getSheetByName('users').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === user_id) {
      return rows[i][5] === true || rows[i][5] === 'TRUE';
    }
  }
  return false;
}

// ============================================================
// ヘルパー
// ============================================================
function _json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// keepWarm: GASのコールドスタートを防ぐ定期実行用関数
// ============================================================
function keepWarm() {
  // 何もしない（トリガーによる定期呼び出しでインスタンスをウォームアップするだけ）
}

// ============================================================
// マイグレーション（v1.7.0・GASエディタから一度だけ手動実行する）
// line-progress（友だち登録進捗ダッシュボード）はGAS側で権限判定をせず全社員閲覧可の
// 方針だが、ポータルのホームにタイルを出すには user_app_roles に行が必要
// （getUserApps() が role!='none' のアプリしか返さないため）。
// active な全ユーザーに 'line-progress'/'user' 行を冪等に追加する。
// 既に行がある場合はスキップするので、何度実行しても安全。
// 参照: LINEHarness/友だち登録進捗ダッシュボード_設計.md §5
// ============================================================
function migrateAddLineProgressRoles() {
  const APP_NAME = 'line-progress';
  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

  const usersSh = ss.getSheetByName('users');
  const users = usersSh.getDataRange().getValues();
  const activeUserIds = [];
  for (let i = 1; i < users.length; i++) {
    const active = users[i][3];
    if (active === true || active === 'TRUE') {
      activeUserIds.push(String(users[i][0]));
    }
  }

  const rolesSh = ss.getSheetByName('user_app_roles');
  const roles = rolesSh.getDataRange().getValues();
  const existing = new Set();
  for (let i = 1; i < roles.length; i++) {
    if (String(roles[i][1]) === APP_NAME) existing.add(String(roles[i][0]));
  }

  let added = 0;
  activeUserIds.forEach(function (userId) {
    if (existing.has(userId)) return; // 既に行がある→スキップ（冪等性）
    rolesSh.appendRow([userId, APP_NAME, 'user']);
    added++;
  });

  Logger.log('migrateAddLineProgressRoles 完了: 対象active ' + activeUserIds.length +
    '名中、新規追加 ' + added + '件（既存 ' + (activeUserIds.length - added) + '件はスキップ）');
}
