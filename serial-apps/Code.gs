// ============================================================
// シリアルNo管理アプリ (SerialApps) - Code.gs v2.1.0
// アーキテクチャ: GitHub Pages (front) + GAS WebApp (API)
// ============================================================
// [重要] コードに機密値を直書きしない。GASスクリプトプロパティに設定すること。
//   GASエディタ → プロジェクトの設定 → スクリプトプロパティ → プロパティを追加
//     SHEET_ID      : このアプリのスプレッドシートID
//     AUTH_SHEET_ID : beaufield-auth スプレッドシートID（共通）
// ============================================================

var VERSION = 'v2.1.0';

var SHEET_ID      = PropertiesService.getScriptProperties().getProperty('SHEET_ID')      || '';
var AUTH_SHEET_ID = PropertiesService.getScriptProperties().getProperty('AUTH_SHEET_ID') || '';

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
  CODE:        0, // A: 商品コード
  NAME:        1, // B: 商品名
  JAN:         2, // C: JANコード
  MAKER:       3, // D: メーカー
  SERIES:      4, // E: 商品シリーズ
  SIZE:        5, // F: サイズ
  SERIAL_TYPE: 6  // G: シリアル種別（バーコード/QR/手入力、空欄=未設定）
};

// ============================================================
// エントリーポイント（GET）
// クエリ: ?action=xxx&session_token=xxx
// ============================================================
function doGet(e) {
  var p      = (e && e.parameter) ? e.parameter : {};
  var action = p.action || '';
  var token  = p.session_token || '';

  // バージョン確認（認証不要）
  if (action === 'getVersion') {
    return jsonResponse({ success: true, version: VERSION });
  }

  var auth = validateSession(token);
  if (!auth.valid) {
    return jsonResponse({ success: false, error: 'SESSION_INVALID', message: '認証が必要です' });
  }

  try {
    switch (action) {
      case 'getProductMaster': return jsonResponse(getProductMaster());
      case 'search':           return jsonResponse(searchRecords(p));
      default:                 return jsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// エントリーポイント（POST）
// application/x-www-form-urlencoded または application/json を受け付ける
// ============================================================
function doPost(e) {
  var p      = {};
  var action = '';

  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      p      = body;
      action = body.action || '';
    } catch (ex) { /* JSON解析失敗 → form-encodedにフォールバック */ }
  }
  if (!action && e && e.parameter) {
    p      = e.parameter;
    action = p.action || '';
  }

  var auth = validateSession(p.session_token || '');
  if (!auth.valid) {
    return jsonResponse({ success: false, error: 'SESSION_INVALID', message: '認証が必要です' });
  }

  try {
    switch (action) {
      case 'registerShipping': return jsonResponse(registerShipping(p));
      case 'registerReturn':   return jsonResponse(registerReturn(p));
      default:                 return jsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// セッション検証（beaufield-auth 共通）
// ============================================================
function validateSession(token) {
  if (!token) return { valid: false };
  try {
    var sh   = SpreadsheetApp.openById(AUTH_SHEET_ID).getSheetByName('sessions');
    if (!sh) return { valid: false };
    var rows = sh.getDataRange().getValues();
    var now  = Date.now();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(token)) {
        if (Number(rows[i][2]) < now) {
          sh.deleteRow(i + 1); // 期限切れ行を削除
          return { valid: false };
        }
        return { valid: true, userId: String(rows[i][1]) };
      }
    }
  } catch (e) {
    Logger.log('validateSession error: ' + e.message);
  }
  return { valid: false };
}

// ============================================================
// 商品マスタ取得
// ============================================================
function getProductMaster() {
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var sh   = ss.getSheetByName(SH_PRODUCT);
    var data = sh.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[PCOL.CODE]) continue;
      result.push({
        code:       String(row[PCOL.CODE]),
        name:       String(row[PCOL.NAME]),
        jan:        String(row[PCOL.JAN]),
        maker:      String(row[PCOL.MAKER]),
        series:     String(row[PCOL.SERIES]),
        size:       String(row[PCOL.SIZE]),
        serialType: String(row[PCOL.SERIAL_TYPE] || '')
      });
    }
    return { success: true, data: result };
  } catch (e) {
    Logger.log('getProductMaster error: ' + e);
    return { success: false, message: '商品マスタの取得に失敗しました' };
  }
}

// ============================================================
// 重複チェック（内部用）
// ============================================================
function checkDuplicates_(productCode, serials) {
  try {
    var sh   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_SHIPPING);
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
    Logger.log('checkDuplicates_ error: ' + e);
    return { success: false, message: '重複チェックに失敗しました' };
  }
}

// ============================================================
// 出荷登録
// params: productCode, productName, jan, shipDate (YYYY/MM/DD),
//         customer, method (単独|連番), serials (JSON文字列 or 配列)
// ============================================================
function registerShipping(params) {
  try {
    var serials   = _parseArray(params.serials);
    var dupResult = checkDuplicates_(params.productCode, serials);
    if (!dupResult.success) return dupResult;

    var dupSet = {};
    dupResult.duplicates.forEach(function(s) { dupSet[s] = true; });

    var sh  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_SHIPPING);
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    var registered = 0, skipped = 0, skippedSerials = [];

    serials.forEach(function(serial) {
      if (dupSet[serial]) { skipped++; skippedSerials.push(serial); return; }
      var row = new Array(13).fill('');
      row[COL.ID]         = Utilities.getUuid();
      row[COL.SHIP_DATE]  = params.shipDate;
      row[COL.PROD_CODE]  = params.productCode;
      row[COL.PROD_NAME]  = params.productName;
      row[COL.JAN]        = params.jan || '';
      row[COL.SERIAL]     = serial;
      row[COL.STATUS]     = '出荷中';
      row[COL.METHOD]     = params.method || '単独';
      row[COL.CUSTOMER]   = params.customer || '';
      row[COL.CREATED_AT] = now;
      sh.appendRow(row);
      registered++;
    });

    return { success: true, registered: registered, skipped: skipped, skippedSerials: skippedSerials };
  } catch (e) {
    Logger.log('registerShipping error: ' + e);
    return { success: false, message: '出荷登録に失敗しました' };
  }
}

// ============================================================
// 返品・取消登録
// params: productCode, returnDate (YYYY/MM/DD),
//         kind (返品|取消), serials (JSON文字列 or 配列)
// ============================================================
function registerReturn(params) {
  try {
    var serials   = _parseArray(params.serials);
    var sh        = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_SHIPPING);
    var data      = sh.getDataRange().getValues();
    var newStatus = (params.kind === '取消') ? '取消' : '返品済';
    var serialSet = {};
    serials.forEach(function(s) { serialSet[s] = true; });

    var updated = 0, skipped = 0, skippedSerials = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[COL.PROD_CODE]) !== String(params.productCode)) continue;
      if (!serialSet[String(row[COL.SERIAL])]) continue;
      if (row[COL.STATUS] !== '出荷中') {
        skipped++;
        skippedSerials.push(String(row[COL.SERIAL]));
        continue;
      }
      sh.getRange(i + 1, COL.STATUS + 1).setValue(newStatus);
      sh.getRange(i + 1, COL.RETURN_DATE + 1).setValue(params.returnDate);
      sh.getRange(i + 1, COL.REASON + 1).setValue(params.kind);
      updated++;
    }

    return { success: true, updated: updated, skipped: skipped, skippedSerials: skippedSerials };
  } catch (e) {
    Logger.log('registerReturn error: ' + e);
    return { success: false, message: '返品・取消登録に失敗しました' };
  }
}

// ============================================================
// 検索
// params: dateFrom, dateTo, statuses (JSON文字列 or 配列),
//         productCode, serialNo, productName, maker, series
// ============================================================
function searchRecords(params) {
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var sh   = ss.getSheetByName(SH_SHIPPING);
    var data = sh.getDataRange().getValues();

    // 日付はYYYY-MM-DD文字列で比較（タイムゾーン問題回避）
    var fromStr = params.dateFrom ? String(params.dateFrom).replace(/\//g, '-').substring(0, 10) : '';
    var toStr   = params.dateTo   ? String(params.dateTo).replace(/\//g, '-').substring(0, 10)   : '';

    var statuses = _parseArray(params.statuses);

    // メーカー・シリーズ・サイズ絞り込み用にProductMasterをメモリ展開
    var prodMap = {};
    if (params.maker || params.series || params.size) {
      var pData = ss.getSheetByName(SH_PRODUCT).getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        var pr = pData[p];
        prodMap[String(pr[PCOL.CODE])] = {
          maker:  String(pr[PCOL.MAKER]),
          series: String(pr[PCOL.SERIES]),
          size:   String(pr[PCOL.SIZE])
        };
      }
    }

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // UUID がない移行データも対象とするため、シリアルNoで空行判定
      if (!row[COL.SERIAL]) continue;

      var shipDateStr = '';
      if (row[COL.SHIP_DATE]) {
        if (row[COL.SHIP_DATE] instanceof Date) {
          // Sheetsがdate型で返した場合（日本時間で YYYY-MM-DD に変換）
          shipDateStr = Utilities.formatDate(row[COL.SHIP_DATE], 'Asia/Tokyo', 'yyyy-MM-dd');
        } else {
          // 文字列で返した場合（"2025/9/19" 形式など 0埋めなしに対応）
          var parts = String(row[COL.SHIP_DATE]).split(/[\/\-]/);
          if (parts.length === 3) {
            shipDateStr = parts[0].padStart(4,'0') + '-' + parts[1].padStart(2,'0') + '-' + parts[2].padStart(2,'0');
          }
        }
      }
      if (fromStr && shipDateStr && shipDateStr < fromStr) continue;
      if (toStr   && shipDateStr && shipDateStr > toStr)   continue;

      if (statuses.length > 0 && statuses.indexOf(String(row[COL.STATUS])) < 0) continue;
      if (params.productCode && _normalizeText(row[COL.PROD_CODE]) !== _normalizeText(params.productCode)) continue;
      if (params.serialNo    && _normalizeText(row[COL.SERIAL])    !== _normalizeText(params.serialNo))    continue;
      if (params.productName && !_matchesQuery(String(row[COL.PROD_NAME]), _normalizeText(params.productName))) continue;

      if (params.maker || params.series || params.size) {
        var pm = prodMap[String(row[COL.PROD_CODE])];
        if (!pm) continue;
        if (params.maker  && _normalizeText(pm.maker)  !== _normalizeText(params.maker))  continue;
        if (params.series && _normalizeText(pm.series) !== _normalizeText(params.series)) continue;
        if (params.size   && _normalizeText(pm.size)   !== _normalizeText(params.size))   continue;
      }

      result.push({
        id:          String(row[COL.ID]),
        shipDate:    _formatDate(row[COL.SHIP_DATE]),
        productCode: String(row[COL.PROD_CODE]),
        productName: String(row[COL.PROD_NAME]),
        jan:         String(row[COL.JAN]),
        serial:      String(row[COL.SERIAL]),
        status:      String(row[COL.STATUS]),
        returnDate:  _formatDate(row[COL.RETURN_DATE]),
        reason:      String(row[COL.REASON]    || ''),
        method:      String(row[COL.METHOD]    || ''),
        customer:    String(row[COL.CUSTOMER]  || ''),
        createdAt:   String(row[COL.CREATED_AT] || '')
      });
    }

    result.sort(function(a, b) { return b.createdAt.localeCompare(a.createdAt); });
    return { success: true, data: result, total: result.length };
  } catch (e) {
    Logger.log('searchRecords error: ' + e);
    return { success: false, message: '検索に失敗しました' };
  }
}

// ============================================================
// ユーティリティ
// ============================================================

// form-encoded で渡ってくる配列はJSON文字列になることがあるため両対応
function _parseArray(val) {
  if (Array.isArray(val)) return val;
  if (!val) return [];
  try { return JSON.parse(val); } catch (e) { return [String(val)]; }
}

// 半角全角・大文字小文字・スペース正規化
function _normalizeText(str) {
  try {
    return String(str || '').normalize('NFKC').toLowerCase().replace(/\s+/g, ' ').trim();
  } catch(e) {
    return String(str || '').toLowerCase().trim();
  }
}

// スペース区切りAND部分一致
function _matchesQuery(text, queryNorm) {
  var t     = _normalizeText(text);
  var terms = queryNorm.split(' ').filter(Boolean);
  return terms.length > 0 && terms.every(function(term) { return t.indexOf(term) >= 0; });
}

function _formatDate(val) {
  if (!val) return '';
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch (e) { return String(val); }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
