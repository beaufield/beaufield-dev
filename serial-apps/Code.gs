// ============================================================
// シリアルNo管理アプリ (SerialApps) - Code.gs v2.6.0
// アーキテクチャ: GitHub Pages (front) + GAS WebApp (API)
// ============================================================
// [重要] コードに機密値を直書きしない。GASスクリプトプロパティに設定すること。
//   GASエディタ → プロジェクトの設定 → スクリプトプロパティ → プロパティを追加
//     SHEET_ID        : このアプリのスプレッドシートID
//     AUTH_SHEET_ID   : beaufield-auth スプレッドシートID（共通）
//     EXPORT_API_KEY  : 月次CSV自動出力（exportShippingCsv）用のAPIキー（ローカルスクリプトと共有）
// ============================================================

var VERSION = 'v2.6.0';

var SHEET_ID       = PropertiesService.getScriptProperties().getProperty('SHEET_ID')       || '';
var AUTH_SHEET_ID  = PropertiesService.getScriptProperties().getProperty('AUTH_SHEET_ID')  || '';
var EXPORT_API_KEY = PropertiesService.getScriptProperties().getProperty('EXPORT_API_KEY') || '';
var CACHE_TTL_SESSION = 900; // 15分（CacheService保持秒数・セッション検証の高速化用）

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
  CREATED_AT:  12, // M: 登録日時
  REGISTERED_BY: 13, // N: 登録者（v2.3.0で追加。登録時のログイン氏名をそのまま保存）
  MAKER:       14, // O: メーカー（v2.3.0で追加。登録時点の商品マスタ値を複製保存）
  SERIES:      15, // P: 商品シリーズ（v2.3.0で追加。同上）
  SIZE:        16  // Q: サイズ（v2.3.0で追加。同上）
};
// ⚠️ N〜Q列は商品マスタの複製値。目的＝将来商品マスタ側の情報が変更・削除されても、
//    出荷済みレコードは登録当時のメーカー/シリーズ/サイズを保持し続けられるようにするため。
//    検索の絞り込み(searchRecords)はこの複製値を優先し、空欄（移行前の旧データ）の場合のみ
//    商品マスタとの突き合わせにフォールバックする。

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

  // メーカー別・月単位のCSV出力（ローカル自動化/スキル用。bf_sessionではなくAPIキーで認証）
  if (action === 'exportShippingCsv') {
    if (!EXPORT_API_KEY || p.api_key !== EXPORT_API_KEY) {
      return ContentService.createTextOutput('UNAUTHORIZED').setMimeType(ContentService.MimeType.TEXT);
    }
    return exportShippingCsv(p);
  }

  var auth = validateSession(token);
  if (!auth.valid) {
    // 認証シートを一時的に読めなかっただけの場合はSESSION_INVALIDにしない。
    // クライアント側はこれを見てログアウトさせず、リトライ対象として扱う
    if (auth.transient) {
      return jsonResponse({ success: false, error: 'AUTH_UNAVAILABLE', message: '認証確認に失敗しました。もう一度お試しください。' });
    }
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
    // 認証シートを一時的に読めなかっただけの場合はSESSION_INVALIDにしない。
    // クライアント側はこれを見てログアウトさせず、リトライ対象として扱う
    if (auth.transient) {
      return jsonResponse({ success: false, error: 'AUTH_UNAVAILABLE', message: '認証確認に失敗しました。もう一度お試しください。' });
    }
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
// CacheService で 15 分間キャッシュしてシート読み込みを削減する
// ============================================================
function validateSession(token) {
  if (!token) return { valid: false };

  var cache    = CacheService.getScriptCache();
  var cacheKey = 'sess_' + token.slice(-32);
  var cached   = cache.get(cacheKey);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  try {
    var sh = SpreadsheetApp.openById(AUTH_SHEET_ID).getSheetByName('sessions');
    // シート取得失敗は「セッション無効」ではなく一時障害。負キャッシュしない
    // （Google側の一時的な応答不良でも起こりうるため。負キャッシュすると
    //  有効なトークンがTTLの間ブロックされ続けてしまう）
    if (!sh) return { valid: false, transient: true };

    var rows = sh.getDataRange().getValues();
    var now  = Date.now();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(token)) {
        if (Number(rows[i][2]) < now) {
          sh.deleteRow(i + 1); // 期限切れ行を削除
          var invalid = { valid: false };
          cache.put(cacheKey, JSON.stringify(invalid), 60);
          return invalid;
        }
        var valid = { valid: true, userId: String(rows[i][1]) };
        cache.put(cacheKey, JSON.stringify(valid), CACHE_TTL_SESSION);
        return valid;
      }
    }
  } catch (e) {
    // 認証シートを読めなかった＝一時障害。ここも負キャッシュしない（上記と同じ理由）
    Logger.log('validateSession error: ' + e.message);
    return { valid: false, transient: true };
  }
  // ここに到達＝シートは読めたがトークンが見つからなかった＝本物の無効
  var result = { valid: false };
  cache.put(cacheKey, JSON.stringify(result), 60);
  return result;
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
// 出荷登録（まとめ書き込み＝高速版）
// params: productCode, productName, jan, shipDate (YYYY/MM/DD),
//         customer, method (単独|連番), serials (JSON文字列 or 配列)
// 【改善点】
//  A. appendRowの1件ずつループを廃止 → setValuesで一括書き込み（数十倍高速）
//  C. 重複チェックと追記でシートを2回開いていたのを1回の読み込みに統合
//  D. 同時登録による行の衝突を防ぐためLockServiceで排他制御
// ============================================================
function registerShipping(params) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 他の登録処理の完了を最大15秒待つ
  } catch (e) {
    return { success: false, message: '他の処理と競合しました。少し待って再度お試しください' };
  }
  try {
    var serials = _parseArray(params.serials);
    var sh      = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_SHIPPING);
    var data    = sh.getDataRange().getValues(); // ← シート読み込みは1回だけ（重複チェックに再利用）

    // 重複（同一商品コード × 出荷中）のシリアルを集合化
    var serialSet = {};
    serials.forEach(function(s) { serialSet[s] = true; });
    var dupSet = {};
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (String(r[COL.PROD_CODE]) === String(params.productCode) &&
          r[COL.STATUS] === '出荷中' &&
          serialSet[String(r[COL.SERIAL])]) {
        dupSet[String(r[COL.SERIAL])] = true;
      }
    }

    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    var newRows = [], skipped = 0, skippedSerials = [];

    serials.forEach(function(serial) {
      if (dupSet[serial]) { skipped++; skippedSerials.push(serial); return; }
      var row = new Array(17).fill('');
      row[COL.ID]           = Utilities.getUuid();
      row[COL.SHIP_DATE]    = params.shipDate;
      row[COL.PROD_CODE]    = params.productCode;
      row[COL.PROD_NAME]    = params.productName;
      row[COL.JAN]          = params.jan || '';
      row[COL.SERIAL]       = serial;
      row[COL.STATUS]       = '出荷中';
      row[COL.METHOD]       = params.method || '単独';
      row[COL.CUSTOMER]     = params.customer || '';
      row[COL.CREATED_AT]   = now;
      row[COL.REGISTERED_BY] = params.registeredBy || '';
      row[COL.MAKER]         = params.maker  || '';
      row[COL.SERIES]        = params.series || '';
      row[COL.SIZE]          = params.size   || '';
      newRows.push(row);
    });

    // 追記は1回のsetValuesで一括（appendRowループの数十倍高速）
    if (newRows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 17).setValues(newRows);
    }

    return { success: true, registered: newRows.length, skipped: skipped, skippedSerials: skippedSerials };
  } catch (e) {
    Logger.log('registerShipping error: ' + e);
    return { success: false, message: '出荷登録に失敗しました' };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 返品・取消登録
// params: productCode, returnDate (YYYY/MM/DD),
//         kind (返品|取消), serials (JSON文字列 or 配列)
// 【改善点】
//  B. 1行につきsetValueを3回呼んでいたのを、変更行だけ1回のsetValuesにまとめて高速化
//  D. 読み書きの競合を防ぐためLockServiceで排他制御
// ============================================================
function registerReturn(params) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, message: '他の処理と競合しました。少し待って再度お試しください' };
  }
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
      // メモリ上で更新し、変更行のG〜J列だけを1回のsetValuesで書き戻す（3回→1回）
      row[COL.STATUS]      = newStatus;
      row[COL.RETURN_DATE] = params.returnDate;
      row[COL.REASON]      = params.kind;
      // G(STATUS)〜J(REASON)の4列を一括書き込み（間のI列CANCEL_DATEは元の値のまま維持）
      sh.getRange(i + 1, COL.STATUS + 1, 1, 4).setValues([row.slice(COL.STATUS, COL.STATUS + 4)]);
      updated++;
    }

    return { success: true, updated: updated, skipped: skipped, skippedSerials: skippedSerials };
  } catch (e) {
    Logger.log('registerReturn error: ' + e);
    return { success: false, message: '返品・取消登録に失敗しました' };
  } finally {
    lock.releaseLock();
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

      // メーカー/シリーズ/サイズは行自体の複製値を優先し、
      // 未設定（複製列導入前の旧データ）の場合のみ商品マスタとの突き合わせにフォールバックする。
      // →商品マスタ側が後で変更・削除されても、登録当時の情報で絞り込みできる。
      var pm        = prodMap[String(row[COL.PROD_CODE])];
      var effMaker  = String(row[COL.MAKER]  || '') || (pm ? pm.maker  : '');
      var effSeries = String(row[COL.SERIES] || '') || (pm ? pm.series : '');
      var effSize   = String(row[COL.SIZE]   || '') || (pm ? pm.size   : '');

      if (params.maker  && _normalizeText(effMaker)  !== _normalizeText(params.maker))  continue;
      if (params.series && _normalizeText(effSeries) !== _normalizeText(params.series)) continue;
      if (params.size   && _normalizeText(effSize)   !== _normalizeText(params.size))   continue;

      result.push({
        id:           String(row[COL.ID]),
        shipDate:     _formatDate(row[COL.SHIP_DATE]),
        productCode:  String(row[COL.PROD_CODE]),
        productName:  String(row[COL.PROD_NAME]),
        jan:          String(row[COL.JAN]),
        serial:       String(row[COL.SERIAL]),
        status:       String(row[COL.STATUS]),
        returnDate:   _formatDate(row[COL.RETURN_DATE]),
        reason:       String(row[COL.REASON]        || ''),
        method:       String(row[COL.METHOD]        || ''),
        customer:     String(row[COL.CUSTOMER]      || ''),
        createdAt:    String(row[COL.CREATED_AT]    || ''),
        registeredBy: String(row[COL.REGISTERED_BY] || ''),
        maker:        effMaker,
        series:       effSeries,
        size:         effSize
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
// メーカー別・月単位のCSV出力
// params: maker（完全一致・商品マスタ突き合わせ含む）, yearMonth（"YYYY-MM"）
// 対象: 出荷日が yearMonth に該当し、状態が「出荷中」の行のみ
// 出力列: 出荷日/商品コード/商品名/JANコード/シリアルNo/状態/返品日/取消日/取消理由/得意先
// ============================================================
function exportShippingCsv(params) {
  try {
    var maker      = params.maker || '';
    var yearMonth  = params.yearMonth || '';
    if (!maker || !/^\d{4}-\d{2}$/.test(yearMonth)) {
      return ContentService.createTextOutput('ERROR: maker, yearMonth(YYYY-MM) は必須です').setMimeType(ContentService.MimeType.TEXT);
    }

    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var sh   = ss.getSheetByName(SH_SHIPPING);
    var data = sh.getDataRange().getValues();

    // メーカー突き合わせ用にProductMasterをメモリ展開（searchRecordsと同じフォールバック方式）
    var pData   = ss.getSheetByName(SH_PRODUCT).getDataRange().getValues();
    var prodMap = {};
    for (var p = 1; p < pData.length; p++) {
      var pr = pData[p];
      prodMap[String(pr[PCOL.CODE])] = { maker: String(pr[PCOL.MAKER]) };
    }

    var makerNorm = _normalizeText(maker);
    var rows = [['出荷日','商品コード','商品名','JANコード','シリアルNo','状態','返品日','取消日','取消理由','得意先']];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.SERIAL]) continue;
      if (String(row[COL.STATUS]) !== '出荷中') continue;

      var shipDateYmd = _shipDateYmd(row[COL.SHIP_DATE]);
      if (!shipDateYmd || shipDateYmd.substring(0, 7) !== yearMonth) continue;

      var pm       = prodMap[String(row[COL.PROD_CODE])];
      var effMaker = String(row[COL.MAKER] || '') || (pm ? pm.maker : '');
      if (_normalizeText(effMaker) !== makerNorm) continue;

      rows.push([
        _formatDate(row[COL.SHIP_DATE]),
        String(row[COL.PROD_CODE]),
        String(row[COL.PROD_NAME]),
        String(row[COL.JAN]),
        String(row[COL.SERIAL]),
        String(row[COL.STATUS]),
        _formatDate(row[COL.RETURN_DATE]),
        _formatDate(row[COL.CANCEL_DATE]),
        String(row[COL.REASON] || ''),
        String(row[COL.CUSTOMER] || '')
      ]);
    }

    var csv = rows.map(function(r) { return r.map(_csvEscape).join(','); }).join('\r\n');
    // 先頭にBOMを付与（ExcelでのUTF-8日本語文字化け対策）。可視文字の混入を避けるためfromCharCodeで生成
    return ContentService.createTextOutput(String.fromCharCode(65279) + csv).setMimeType(ContentService.MimeType.CSV);
  } catch (e) {
    Logger.log('exportShippingCsv error: ' + e);
    return ContentService.createTextOutput('ERROR: ' + e).setMimeType(ContentService.MimeType.TEXT);
  }
}

// 出荷日セルをYYYY-MM-DD文字列に正規化（searchRecordsの日付比較ロジックと同じ方式）
function _shipDateYmd(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  var parts = String(val).split(/[\/\-]/);
  if (parts.length === 3) {
    return parts[0].padStart(4, '0') + '-' + parts[1].padStart(2, '0') + '-' + parts[2].padStart(2, '0');
  }
  return '';
}

// CSVフィールドのエスケープ（カンマ・改行・ダブルクォートを含む場合のみクォート）
function _csvEscape(v) {
  var s = String(v == null ? '' : v);
  if (/[",\r\n]/.test(s)) {
    s = '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
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
