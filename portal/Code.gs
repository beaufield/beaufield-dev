// ============================================================
// Beaufield ポータル - Google Apps Script
// Version: v1.3.3
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   AUTH_SHEET_ID : beaufield-auth スプレッドシートID（共通）
//
// ============================================================

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS        = PropertiesService.getScriptProperties();
const VERSION       = 'v1.3.4';
const AUTH_SHEET_ID = _PROPS.getProperty('AUTH_SHEET_ID');

// ロックアウト設定
const MAX_ATTEMPTS = 5;
const LOCK_MINUTES = 10;

// セッション有効期間（日）
const SESSION_DAYS = 30;

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
    url:     'https://beaufield.github.io/kiki-kanri/'
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
  }
];

// ============================================================
// エントリーポイント（GET）
// ============================================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
  const data   = e && e.parameter && e.parameter.data   ? JSON.parse(e.parameter.data) : {};

  try {
    switch (action) {
      case 'getUsers':    return _json(getUsers());
      case 'getUserApps': return _json(getUserApps(data));
      default:            return _json({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    return _json({ success: false, error: err.toString() });
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
    data   = e.parameter.data ? JSON.parse(e.parameter.data) : {};
  }

  try {
    switch (action) {
      case 'login':           return _json(login(data));
      case 'resetPin':        return _json(resetPin(data));
      case 'validateSession': return _json(validateSession(data));
      default:                return _json({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    return _json({ success: false, error: err.toString() });
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
        const expiresAt = now + SESSION_DAYS * 24 * 60 * 60 * 1000;
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
      return { success: true, message: 'PINを更新しました' };
    }
  }
  return { success: false, message: '対象ユーザーが見つかりません' };
}

// ============================================================
// アクセス可能アプリ一覧取得
// ============================================================
function getUserApps(data) {
  const { user_id } = data;
  if (!user_id) return { success: false, message: 'user_idは必須です' };

  const ss    = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const roles = ss.getSheetByName('user_app_roles').getDataRange().getValues();

  // そのユーザーがアクセス権を持つアプリ名のセット
  const accessMap = {};
  for (let i = 1; i < roles.length; i++) {
    if (String(roles[i][0]) === user_id && roles[i][2] !== 'none') {
      accessMap[String(roles[i][1])] = String(roles[i][2]); // appName → role
    }
  }

  // APP_MASTER から該当するものだけ返す（定義順を維持）
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
  if (!token) return { ok: false };

  const ss     = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const userId = _getSessionUser(ss, token);
  if (!userId) return { ok: false };

  // ユーザー名を users シートから取得して返す
  const rows = ss.getSheetByName('users').getDataRange().getValues();
  let userName = userId;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === userId) {
      userName = String(rows[i][1]) || userId;
      break;
    }
  }

  return { ok: true, user_id: userId, name: userName };
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
