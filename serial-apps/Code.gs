// ============================================================
// シリアルNo管理アプリ (SerialApps) - Code.gs v1.1.0
// ============================================================

var VERSION = 'v1.1.0';

// ロックアウト設定
var MAX_ATTEMPTS = 5;
var LOCK_MINUTES = 10;
var SESSION_DAYS = 30;

// ---- 設定 --------------------------------------------------
var SHEET_ID  = '';  // ★ 初回セットアップ後にGoogleSheetsのIDを記入
var AUTH_SHEET_ID = '1cCQn16ubEN_Af7XWw8KerBscZtFomBnXHjIIiZUr6V8';
var APP_NAME  = 'serial-apps';

// シート名
var SH_PRODUCT  = 'ProductMaster';
var SH_SHIPPING = 'SerialShipping';

// SerialShipping 列インデックス（0始まり）
var COL = {
  ID:          0,  // A: ID（UUID）
  SHIP_DATE:   1,  // B: 出荷日
  PROD_CODE:   2,  // C: 商品コード
  PROD_NAME:   3,  // D: 商品名
  JAN:         4,  // E: JANコード
  SERIAL:      5,  // F: シリアルNo
  STATUS:      6,  // G: 状態（出荷中/返品済/取消）
  RETURN_DATE: 7,  // H: 返品日
  CANCEL_DATE: 8,  // I: 取消日（未使用）
  REASON:      9,  // J: 取消理由（「返品」「取消」を格納）
  METHOD:      10, // K: 登録方法（単独/連番）
  CUSTOMER:    11, // L: 得意先
  CREATED_AT:  12  // M: 登録日時
};

// ProductMaster 列インデックス（0始まり）
var PCOL = {
  CODE:   0, // A: 商品コード
  NAME:   1, // B: 商品名
  JAN:    2, // C: JANコード
  MAKER:  3, // D: メーカー
  SERIES: 4, // E: 商品シリーズ
  SIZE:   5  // F: サイズ
};

// ============================================================
// WebApp エントリポイント
// ============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('シリアルNo管理アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// 認証（beaufield-auth 共通）
// ============================================================

// ログイン前にユーザー一覧を返す（ドロップダウン用・認証不要）
function getUserList() {
  try {
    var ss    = SpreadsheetApp.openById(AUTH_SHEET_ID);
    var users = ss.getSheetByName('users').getDataRange().getValues();
    var roles = ss.getSheetByName('user_app_roles').getDataRange().getValues();
    var allowedIds = {};
    for (var i = 1; i < roles.length; i++) {
      if (roles[i][1] === APP_NAME && roles[i][2] && roles[i][2] !== 'none') {
        allowedIds[String(roles[i][0])] = true;
      }
    }
    var result = [];
    for (var i = 1; i < users.length; i++) {
      var uid = String(users[i][0]);
      if (users[i][3] === true && allowedIds[uid]) {
        result.push({ userId: uid, name: String(users[i][1]) });
      }
    }
    return { success: true, users: result };
  } catch (e) {
    Logger.log('getUserList error: ' + e);
    return { success: false, users: [] };
  }
}

// userId + PIN でログイン。5回失敗でロックアウト
function loginUser(userId, pin) {
  try {
    var props   = PropertiesService.getScriptProperties();
    var lockKey = 'serial_lockout_' + userId;
    var lockData = JSON.parse(props.getProperty(lockKey) || '{"count":0,"until":0}');
    var now = Date.now();

    if (lockData.until > now) {
      var remaining = Math.ceil((lockData.until - now) / 60000);
      return { success: false, message: 'PINを' + MAX_ATTEMPTS + '回間違えました。' + remaining + '分後に再試行してください。' };
    }

    var ss    = SpreadsheetApp.openById(AUTH_SHEET_ID);
    var users = ss.getSheetByName('users').getDataRange().getValues();

    for (var i = 1; i < users.length; i++) {
      if (String(users[i][0]) !== String(userId)) continue;
      if (users[i][3] !== true) {
        return { success: false, message: 'このアカウントは無効です' };
      }
      if (String(users[i][2]) !== String(pin)) {
        lockData.count = (lockData.count || 0) + 1;
        if (lockData.count >= MAX_ATTEMPTS) {
          lockData.until = now + LOCK_MINUTES * 60 * 1000;
          lockData.count = 0;
          props.setProperty(lockKey, JSON.stringify(lockData));
          return { success: false, message: 'PINを' + MAX_ATTEMPTS + '回間違えました。' + LOCK_MINUTES + '分間ロックされます。' };
        }
        props.setProperty(lockKey, JSON.stringify(lockData));
        return { success: false, message: 'PINが正しくありません（残り' + (MAX_ATTEMPTS - lockData.count) + '回）' };
      }
      // PIN一致 → ロックリセット・セッション発行
      props.deleteProperty(lockKey);
      var role = _getAppRole(ss, userId);
      if (!role || role === 'none') {
        return { success: false, message: 'このアプリへのアクセス権限がありません' };
      }
      var token     = Utilities.getUuid();
      var expiresAt = now + SESSION_DAYS * 24 * 60 * 60 * 1000;
      _saveSession(ss, token, userId, expiresAt);
      return { success: true, userId: userId, name: String(users[i][1]), role: role, sessionToken: token };
    }
    return { success: false, message: 'ユーザーが見つかりません' };
  } catch (e) {
    Logger.log('loginUser error: ' + e);
    return { success: false, message: '認証サーバーへの接続に失敗しました' };
  }
}

function _getAppRole(ss, userId) {
  var roles = ss.getSheetByName('user_app_roles').getDataRange().getValues();
  for (var i = 1; i < roles.length; i++) {
    if (roles[i][0] === userId && roles[i][1] === APP_NAME) {
      return roles[i][2];
    }
  }
  return null;
}

// beaufield-auth.sessions シートに保存。構造: [token, userId, expiresAt(ms)]
function _saveSession(ss, token, userId, expiresAt) {
  ss.getSheetByName('sessions').appendRow([token, userId, expiresAt]);
}

// セッショントークンを検証して userId を返す
function validateSession_(token) {
  if (!token) return { valid: false };
  try {
    var sh   = SpreadsheetApp.openById(AUTH_SHEET_ID).getSheetByName('sessions');
    if (!sh) return { valid: false };
    var rows = sh.getDataRange().getValues();
    var now  = Date.now();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(token)) {
        if (Number(rows[i][2]) < now) return { valid: false };
        return { valid: true, userId: String(rows[i][1]) };
      }
    }
  } catch (e) {
    Logger.log('validateSession_ error: ' + e.message);
  }
  return { valid: false };
}

// ============================================================
// 商品マスタ
// ============================================================

/**
 * 全商品マスタを取得する
 * @returns {Array} 商品オブジェクトの配列
 */
function getProductMaster(sessionToken) {
  var auth = validateSession_(sessionToken);
  if (!auth.valid) return { success: false, message: 'SESSION_INVALID' };
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var sh   = ss.getSheetByName(SH_PRODUCT);
    var data = sh.getDataRange().getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[PCOL.CODE]) continue; // 空行スキップ
      result.push({
        code:   String(row[PCOL.CODE]),
        name:   String(row[PCOL.NAME]),
        jan:    String(row[PCOL.JAN]),
        maker:  String(row[PCOL.MAKER]),
        series: String(row[PCOL.SERIES]),
        size:   String(row[PCOL.SIZE])
      });
    }
    return { success: true, data: result };
  } catch (e) {
    Logger.log('getProductMaster error: ' + e);
    Logger.log('getProductMaster error: ' + e);
    return { success: false, message: '商品マスタの取得に失敗しました' };
  }
}

// ============================================================
// 重複チェック
// ============================================================

/**
 * 指定した商品コード×シリアルNo のうち「出荷中」が存在するものを返す
 * @param {string} productCode
 * @param {Array}  serials - シリアルNoの配列
 * @returns {{ duplicates: Array }} 出荷中が既存のシリアルNo一覧
 */
function checkDuplicates(productCode, serials) {
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var sh   = ss.getSheetByName(SH_SHIPPING);
    var data = sh.getDataRange().getValues();
    var serialSet = {};
    serials.forEach(function(s) { serialSet[s] = true; });

    var duplicates = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[COL.PROD_CODE]) === String(productCode) &&
          row[COL.STATUS] === '出荷中' &&
          serialSet[String(row[COL.SERIAL])]) {
        duplicates.push(String(row[COL.SERIAL]));
      }
    }
    return { success: true, duplicates: duplicates };
  } catch (e) {
    Logger.log('checkDuplicates error: ' + e);
    Logger.log('checkDuplicates error: ' + e);
    return { success: false, message: '重複チェックに失敗しました' };
  }
}

// ============================================================
// 出荷登録
// ============================================================

/**
 * 出荷を登録する
 * @param {object} params
 *   productCode, productName, jan, shipDate (YYYY/MM/DD),
 *   customer, method (単独|連番), serials: Array<string>
 * @returns {{ success, registered, skipped, skippedSerials }}
 */
function registerShipping(params) {
  var auth = validateSession_(params.sessionToken);
  if (!auth.valid) return { success: false, message: 'SESSION_INVALID' };
  try {
    var dupResult = checkDuplicates(params.productCode, params.serials);
    if (!dupResult.success) return dupResult;

    var dupSet = {};
    dupResult.duplicates.forEach(function(s) { dupSet[s] = true; });

    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var sh  = ss.getSheetByName(SH_SHIPPING);
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    var registered = 0;
    var skipped    = 0;
    var skippedSerials = [];

    params.serials.forEach(function(serial) {
      if (dupSet[serial]) {
        skipped++;
        skippedSerials.push(serial);
        return;
      }
      var row = new Array(13).fill('');
      row[COL.ID]          = Utilities.getUuid();
      row[COL.SHIP_DATE]   = params.shipDate;
      row[COL.PROD_CODE]   = params.productCode;
      row[COL.PROD_NAME]   = params.productName;
      row[COL.JAN]         = params.jan || '';
      row[COL.SERIAL]      = serial;
      row[COL.STATUS]      = '出荷中';
      row[COL.METHOD]      = params.method || '単独';
      row[COL.CUSTOMER]    = params.customer || '';
      row[COL.CREATED_AT]  = now;
      sh.appendRow(row);
      registered++;
    });

    return { success: true, registered: registered, skipped: skipped, skippedSerials: skippedSerials };
  } catch (e) {
    Logger.log('registerShipping error: ' + e);
    Logger.log('registerShipping error: ' + e);
    return { success: false, message: '出荷登録に失敗しました' };
  }
}

// ============================================================
// 返品・取消登録
// ============================================================

/**
 * 返品または取消を登録する
 * @param {object} params
 *   productCode, returnDate (YYYY/MM/DD),
 *   kind (返品|取消), serials: Array<string>
 * @returns {{ success, updated, skipped, skippedSerials }}
 */
function registerReturn(params) {
  var auth = validateSession_(params.sessionToken);
  if (!auth.valid) return { success: false, message: 'SESSION_INVALID' };
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var sh   = ss.getSheetByName(SH_SHIPPING);
    var data = sh.getDataRange().getValues();

    var newStatus = (params.kind === '取消') ? '取消' : '返品済';
    var serialSet = {};
    params.serials.forEach(function(s) { serialSet[s] = true; });

    var updated  = 0;
    var skipped  = 0;
    var skippedSerials = [];

    // 行番号は1始まり（0行目はヘッダー）
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[COL.PROD_CODE]) !== String(params.productCode)) continue;
      if (!serialSet[String(row[COL.SERIAL])]) continue;

      if (row[COL.STATUS] !== '出荷中') {
        skipped++;
        skippedSerials.push(String(row[COL.SERIAL]));
        continue;
      }

      // G列（状態）= i+1行目（1始まり）
      sh.getRange(i + 1, COL.STATUS + 1).setValue(newStatus);
      sh.getRange(i + 1, COL.RETURN_DATE + 1).setValue(params.returnDate);
      sh.getRange(i + 1, COL.REASON + 1).setValue(params.kind);
      updated++;
    }

    return { success: true, updated: updated, skipped: skipped, skippedSerials: skippedSerials };
  } catch (e) {
    Logger.log('registerReturn error: ' + e);
    Logger.log('registerReturn error: ' + e);
    return { success: false, message: '返品・取消登録に失敗しました' };
  }
}

// ============================================================
// 検索
// ============================================================

/**
 * 出荷履歴を検索する
 * @param {object} params
 *   dateFrom, dateTo (YYYY/MM/DD),
 *   statuses: Array<string> (空=全件),
 *   productCode, serialNo, productName (前方一致),
 *   maker, series
 * @returns {{ success, data: Array }}
 */
function searchRecords(params) {
  var auth = validateSession_(params.sessionToken);
  if (!auth.valid) return { success: false, message: 'SESSION_INVALID' };
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var sh   = ss.getSheetByName(SH_SHIPPING);
    var data = sh.getDataRange().getValues();

    var from     = params.dateFrom ? new Date(params.dateFrom) : null;
    var to       = params.dateTo   ? new Date(params.dateTo)   : null;
    if (to) to.setHours(23, 59, 59);

    // ProductMaster をメモリに読む（メーカー・シリーズ絞り込み用）
    var prodMap = {};
    if (params.maker || params.series) {
      var pData = ss.getSheetByName(SH_PRODUCT).getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        var pr = pData[p];
        prodMap[String(pr[PCOL.CODE])] = {
          maker:  String(pr[PCOL.MAKER]),
          series: String(pr[PCOL.SERIES])
        };
      }
    }

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.ID]) continue; // 空行スキップ

      // 日付フィルタ（出荷日 or 返品日）
      var shipDate   = row[COL.SHIP_DATE]   ? new Date(row[COL.SHIP_DATE])   : null;
      var returnDate = row[COL.RETURN_DATE] ? new Date(row[COL.RETURN_DATE]) : null;
      var targetDate = shipDate; // 基本は出荷日
      if (from && targetDate && targetDate < from) continue;
      if (to   && targetDate && targetDate > to)   continue;

      // 状態フィルタ
      if (params.statuses && params.statuses.length > 0) {
        if (params.statuses.indexOf(String(row[COL.STATUS])) < 0) continue;
      }

      // 商品コード完全一致
      if (params.productCode && String(row[COL.PROD_CODE]) !== params.productCode) continue;

      // シリアルNo完全一致
      if (params.serialNo && String(row[COL.SERIAL]) !== params.serialNo) continue;

      // 商品名前方一致
      if (params.productName && String(row[COL.PROD_NAME]).indexOf(params.productName) !== 0) continue;

      // メーカー・シリーズ絞り込み
      if (params.maker || params.series) {
        var pm = prodMap[String(row[COL.PROD_CODE])];
        if (!pm) continue;
        if (params.maker  && pm.maker  !== params.maker)  continue;
        if (params.series && pm.series !== params.series) continue;
      }

      result.push({
        id:         String(row[COL.ID]),
        shipDate:   _formatDate(row[COL.SHIP_DATE]),
        productCode: String(row[COL.PROD_CODE]),
        productName: String(row[COL.PROD_NAME]),
        jan:        String(row[COL.JAN]),
        serial:     String(row[COL.SERIAL]),
        status:     String(row[COL.STATUS]),
        returnDate: _formatDate(row[COL.RETURN_DATE]),
        reason:     String(row[COL.REASON] || ''),
        method:     String(row[COL.METHOD] || ''),
        customer:   String(row[COL.CUSTOMER] || ''),
        createdAt:  String(row[COL.CREATED_AT] || '')
      });
    }

    // 登録日時の新しい順に並び替え
    result.sort(function(a, b) {
      return b.createdAt.localeCompare(a.createdAt);
    });

    return { success: true, data: result, total: result.length };
  } catch (e) {
    Logger.log('searchRecords error: ' + e);
    Logger.log('searchRecords error: ' + e);
    return { success: false, message: '検索に失敗しました' };
  }
}

// ============================================================
// ユーティリティ
// ============================================================
function _formatDate(val) {
  if (!val) return '';
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch (e) {
    return String(val);
  }
}

/**
 * バージョン確認用
 */
function getVersion() {
  return VERSION;
}
