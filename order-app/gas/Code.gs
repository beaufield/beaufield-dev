// ============================================================
// Beaufield 発注アプリ - Google Apps Script バックエンド
// Version: v1.8.0
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   SPREADSHEET_ID  : 発注管理データのスプレッドシートID
//   AUTH_SHEET_ID   : beaufield-auth スプレッドシートID（共通）
//   REORDER_API_KEY : 発注点更新用APIキー（Pythonスクリプトと共有）
//
// ============================================================

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS          = PropertiesService.getScriptProperties();
const SPREADSHEET_ID  = _PROPS.getProperty('SPREADSHEET_ID');
const AUTH_SHEET_ID   = _PROPS.getProperty('AUTH_SHEET_ID');
const VERSION         = 'v1.9.0';

// Google Drive上の商品マスターCSVファイル名
// ※ 同名ファイルが複数ある場合はファイルIDで指定（下記コメント参照）
const PRODUCT_CSV_NAME = '商品.CSV';

// ファイルIDで直接指定する場合はこちらを使う（より確実）
// Google DriveでファイルをID確認後に設定: setProductFileId() を実行
const PRODUCT_FILE_ID_KEY = 'PRODUCT_FILE_ID'; // ScriptPropertiesのキー名

// シート名定数
const SHEET_HISTORY   = '発注履歴';
const SHEET_ITEMS     = '発注明細';
const SHEET_SUPPLIERS = '発注先マスター';
const SHEET_STAFF     = '担当者マスター';
const SHEET_PRODUCTS  = '商品マスター';
const SHEET_REORDER   = '発注点マスター';

// メーカー発注書テンプレート定義（Drive配信用）
// スクリプトプロパティ ORDER_TEMPLATE_FOLDER_ID に Drive フォルダIDを設定すること
const ORDER_TEMPLATES = {
  'grandex':           'グランデックス.pdf',
  'chiyoda':           '千代田化学.pdf',
  'alpenrose':         'アルペンローゼhyumi専用発注書.pdf',
  'melos':             'メロス発注書2026年4月～.pdf',
  'melos_2025':        'メロス発注書2025年価格改定後.pdf',
  'adelans':           'アデランス発注書.pdf',
  'rhythm':            'リズム注文書2023冬～.pdf',
  'hokkaido_natural':  '北海道ナチュラルバイオ.pdf'
};

// ============================================================
// セッション検証
// beaufield-auth の sessions シートでトークンを照合する
// 戻り値: { valid: true, user_id } または { valid: false }
// ============================================================
function validateSession(token) {
  if (!token) return { valid: false };
  try {
    const ss   = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const sh   = ss.getSheetByName('sessions');
    if (!sh) return { valid: false };

    const data = sh.getDataRange().getValues();
    const now  = Date.now();

    for (let i = 1; i < data.length; i++) {
      const rowToken   = String(data[i][0]);
      const rowUserId  = String(data[i][1]);
      const rowExpires = Number(data[i][2]);

      if (rowToken === token) {
        if (rowExpires < now) {
          // 期限切れ → 行を削除してから拒否
          sh.deleteRow(i + 1);
          return { valid: false };
        }
        return { valid: true, user_id: rowUserId };
      }
    }
  } catch(e) {
    Logger.log('セッション検証エラー: ' + e);
  }
  return { valid: false };
}

// ============================================================
// エントリーポイント（GET）
// ============================================================
function doGet(e) {
  const p      = (e && e.parameter) ? e.parameter : {};
  const action = p.action || '';
  const token  = p.session_token || '';

  // セッション検証
  const auth = validateSession(token);
  if (!auth.valid) {
    return jsonResponse({ success: false, error: 'SESSION_INVALID', message: '認証が必要です。ポータルからログインし直してください。' });
  }

  try {
    switch (action) {
      case 'getMasters':        return jsonResponse(getMasters());
      case 'getProductMaster':  return jsonResponse(getProductMaster());
      case 'getOrders':         return jsonResponse(getOrders(p.supplierCode || ''));
      case 'getOrderDetail':    return jsonResponse(getOrderDetail(p.orderNo));
      case 'getOrderTemplate':  return jsonResponse(getOrderTemplate(p.makerKey || ''));
      default:                  return jsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch(err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
// エントリーポイント（POST）
// application/x-www-form-urlencoded または application/json を受け付ける
// ============================================================
function doPost(e) {
  // JSON ボディを優先して解析（Power AutomateからのPOST対応）
  let p = {};
  let action = '';
  if (e && e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      p = body;
      action = body.action || '';
    } catch(ex) {
      // JSON解析失敗 → form-encodedとして処理
    }
  }
  // form-encoded フォールバック
  if (!action && e && e.parameter) {
    p = e.parameter;
    action = p.action || '';
  }

  // updateProductMaster: シークレットキー認証（Power Automate用・セッション不要）
  if (action === 'updateProductMaster') {
    if (p.secret !== UPDATE_SECRET) {
      return jsonResponse({ success: false, error: 'UNAUTHORIZED' });
    }
    try {
      return jsonResponse(updateProductMaster(p.data || ''));
    } catch(err) {
      return jsonResponse({ success: false, error: err.toString() });
    }
  }

  // updateReorderPoints: APIキー認証（Pythonスクリプト用・セッション不要）
  if (action === 'updateReorderPoints') {
    const apiKey = p.api_key || '';
    if (!apiKey || apiKey !== _PROPS.getProperty('REORDER_API_KEY')) {
      return jsonResponse({ success: false, error: 'UNAUTHORIZED' });
    }
    try {
      const products = p.products || [];
      return jsonResponse(updateReorderPoints(products));
    } catch(err) {
      return jsonResponse({ success: false, error: err.toString() });
    }
  }

  // 通常のセッション検証
  const token = p.session_token || '';
  const auth = validateSession(token);
  if (!auth.valid) {
    return jsonResponse({ success: false, error: 'SESSION_INVALID', message: '認証が必要です。ポータルからログインし直してください。' });
  }

  try {
    switch (action) {
      case 'saveOrder':    return jsonResponse(saveOrder(p, auth.user_id));
      case 'deleteOrder':  return jsonResponse(deleteOrder(p, auth.user_id));
      case 'saveSupplier': return jsonResponse(saveSupplier(p, auth.user_id));
      case 'saveStaff':    return jsonResponse(saveStaff(p, auth.user_id));
      default:             return jsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch(err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
// ヘルパー: JSONレスポンス生成
// ============================================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ヘルパー: シート取得
// ============================================================
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name);
  return sh;
}

// ============================================================
// ヘルパー: セル値を安全に文字列化
// スプレッドシートが日付・日時型として認識したセルは getValues() で
// Dateオブジェクトとして返ってくる。そのまま String() すると
// "Thu Mar 26 2026 00:00:00 GMT+0900..." のGMT形式になってしまうため、
// Utilities.formatDate() を使って明示的にフォーマットする。
//
// fmt: 省略時は 'yyyy/MM/dd HH:mm'（日時用）
//       日付のみの場合は 'yyyy/MM/dd' を指定すること
// ============================================================
function cellToStr(val, fmt) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Tokyo', fmt || 'yyyy/MM/dd HH:mm');
  }
  return String(val || '');
}

// ============================================================
// GET: マスターデータ一括取得
// レスポンス: { success: true, suppliers: [...], staff: [...] }
// ============================================================
function getMasters() {
  const suppSheet  = getSheet(SHEET_SUPPLIERS);
  const staffSheet = getSheet(SHEET_STAFF);

  const suppData  = suppSheet.getDataRange().getValues();
  const suppliers = suppData.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      code:          String(r[0]).trim(),
      name:          String(r[1]).trim(),
      fax:           String(r[2] || '').trim(),
      // 発注方法: カンマ区切り文字列 → 配列。空欄は空配列（アプリ側で全ボタン表示）
      outputMethods: String(r[4] || '').trim()
                       .split(',')
                       .map(s => s.trim())
                       .filter(Boolean)
    }));

  const staffData = staffSheet.getDataRange().getValues();
  const staff = staffData.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({ name: String(r[0]).trim() }));

  return { success: true, suppliers, staff };
}

// ============================================================
// GET: 商品マスター取得
// Google Sheetsの「商品マスター」シートから読み込んで返す
// 「発注点マスター」シートのデータをJOINして reorderPoint フィールドを付与する
// レスポンス: { success: true, products: [...], updatedAt: '...' }
// ============================================================
function getProductMaster() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sh || sh.getLastRow() === 0) {
    return { success: true, products: [], updatedAt: '' };
  }

  const data    = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const colMap = {
    code:            findColIdxGAS(headers, 'コード'),
    name:            findColIdxGAS(headers, '商品名'),
    kana:            findColIdxGAS(headers, 'かな'),
    unit:            findColIdxGAS(headers, '単位名'),
    supplierCD:      findColIdxGAS(headers, '仕入先CD'),
    supplierName:    findColIdxGAS(headers, '仕入先名'),
    makerCode:       findColIdxGAS(headers, '相手商品CD'),
    jan:             findColIdxGAS(headers, 'JANCD'),
    purchasePrice:   findColIdxGAS(headers, '仕入単価'),
    discontinued:    findColIdxGAS(headers, '廃番'),
    stockManagement: findColIdxGAS(headers, '在庫有無'),
    lastSaleDate:    findColIdxGAS(headers, '最終売上日'),
    stock:           findColIdxGAS(headers, '在庫数')
  };

  const products = [];
  for (let i = 1; i < data.length; i++) {
    const r    = data[i];
    const name = colMap.name !== -1 ? String(r[colMap.name] || '').trim() : '';
    if (!name) continue;
    products.push({
      code:            colMap.code !== -1            ? String(r[colMap.code]            || '').trim()                                        : '',
      name:            name,
      kana:            colMap.kana !== -1            ? String(r[colMap.kana]            || '').trim()                                        : '',
      unit:            colMap.unit !== -1            ? String(r[colMap.unit]            || '').trim()                                        : '',
      supplierCD:      colMap.supplierCD !== -1      ? String(r[colMap.supplierCD]      || '').trim()                                        : '',
      supplierName:    colMap.supplierName !== -1    ? String(r[colMap.supplierName]    || '').trim()                                        : '',
      makerCode:       colMap.makerCode !== -1       ? String(r[colMap.makerCode]       || '').trim()                                        : '',
      jan:             colMap.jan !== -1             ? String(r[colMap.jan]             || '').trim()                                        : '',
      purchasePrice:   colMap.purchasePrice !== -1   ? (parseFloat(String(r[colMap.purchasePrice] || '0').replace(/,/g, '')) || 0)          : 0,
      discontinued:    colMap.discontinued !== -1    ? String(r[colMap.discontinued]    || '').trim()                                        : '',
      stockManagement: colMap.stockManagement !== -1 ? String(r[colMap.stockManagement] || '').trim()                                        : '',
      lastSaleDate:    colMap.lastSaleDate !== -1    ? String(r[colMap.lastSaleDate]    || '').trim()                                        : '',
      stock:           colMap.stock !== -1           ? String(r[colMap.stock]           || '').trim()                                        : '',
      reorderPoint:    null,
      reorderUpdatedAt: ''
    });
  }

  // 発注点マスターをJOIN
  try {
    const rsh = ss.getSheetByName(SHEET_REORDER);
    if (rsh && rsh.getLastRow() > 1) {
      const rdata = rsh.getDataRange().getValues();
      // ヘッダー: [0]=商品コード [1]=月平均出荷数 [2]=更新日時
      const reorderMap = {};
      for (let i = 1; i < rdata.length; i++) {
        const code = String(rdata[i][0] || '').trim();
        if (code) {
          reorderMap[code] = {
            reorderPoint:     parseFloat(rdata[i][1]) || 0,
            reorderUpdatedAt: cellToStr(rdata[i][2], 'yyyy/MM/dd')
          };
        }
      }
      products.forEach(p => {
        if (p.code && reorderMap[p.code]) {
          p.reorderPoint     = reorderMap[p.code].reorderPoint;
          p.reorderUpdatedAt = reorderMap[p.code].reorderUpdatedAt;
        }
      });
    }
  } catch(e) {
    Logger.log('発注点マスター読込エラー（無視）: ' + e);
  }

  const updatedAt = PropertiesService.getScriptProperties().getProperty('PM_UPDATED_AT') || '';
  return { success: true, products, updatedAt };
}

// ============================================================
// POST: 発注点マスター更新（Pythonスクリプトからの自動実行用）
// APIキー認証のみ（セッション不要）
// リクエスト: { action: 'updateReorderPoints', api_key: '...', products: [{code, reorderPoint, updatedAt}] }
// レスポンス: { success: true, count: N }
// ============================================================
function updateReorderPoints(products) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_REORDER);
  if (!sh) sh = ss.insertSheet(SHEET_REORDER);

  sh.clearContents();
  sh.getRange(1, 1, 1, 3).setValues([['商品コード', '月平均出荷数', '更新日時']]);

  if (!products || products.length === 0) {
    return { success: true, count: 0 };
  }

  const rows = products.map(p => [
    String(p.code       || '').trim(),
    parseFloat(p.reorderPoint) || 0,
    String(p.updatedAt  || '')
  ]);
  sh.getRange(2, 1, rows.length, 3).setValues(rows);

  Logger.log('✅ 発注点マスター更新完了: ' + rows.length + '件');
  return { success: true, count: rows.length };
}

// ============================================================
// タイマートリガー: Google Driveから商品マスターCSVを読み込んでSheets更新
// GASエディタのトリガー設定、または setDailyTrigger() で毎朝6時に自動実行される
// ============================================================
function updateProductMasterFromDrive() {
  let file = null;

  // まずScriptPropertiesにファイルIDが保存されていればそちらを優先
  const savedId = PropertiesService.getScriptProperties().getProperty(PRODUCT_FILE_ID_KEY);
  if (savedId) {
    try {
      file = DriveApp.getFileById(savedId);
    } catch(e) {
      Logger.log('保存済みファイルID無効。名前検索にフォールバック: ' + e);
    }
  }

  // ファイルIDがなければファイル名で検索（更新日時が最新のものを使用）
  if (!file) {
    const files = DriveApp.getFilesByName(PRODUCT_CSV_NAME);
    let latest = null;
    while (files.hasNext()) {
      const f = files.next();
      if (!latest || f.getLastUpdated() > latest.getLastUpdated()) latest = f;
    }
    if (!latest) throw new Error('Google Driveに「' + PRODUCT_CSV_NAME + '」が見つかりません');
    file = latest;
    // 次回以降はファイルIDで直接取得するよう保存
    PropertiesService.getScriptProperties().setProperty(PRODUCT_FILE_ID_KEY, file.getId());
  }

  const csvText = file.getBlob().getDataAsString('Shift-JIS');
  const rows    = parseCSVText(csvText);
  if (rows.length < 2) throw new Error('CSVが空か1行のみです');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sh) sh = ss.insertSheet(SHEET_PRODUCTS);

  sh.clearContents();
  // 大量データは一括書き込みで高速化
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  const updatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  PropertiesService.getScriptProperties().setProperty('PM_UPDATED_AT', updatedAt);
  Logger.log('✅ 商品マスター更新完了: ' + (rows.length - 1) + '件 (' + updatedAt + ')');
}

// ============================================================
// 毎朝6時の自動トリガーを設定する（GASエディタから【1回だけ】手動実行）
// 既存の同名トリガーがあれば先に削除してから再登録する
// ============================================================
function setDailyTrigger() {
  // 既存トリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'updateProductMasterFromDrive') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 毎朝6時（日本時間）に登録
  ScriptApp.newTrigger('updateProductMasterFromDrive')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone('Asia/Tokyo')
    .create();
  Logger.log('✅ 毎朝6時トリガーを設定しました');
}

// ============================================================
// POST: 商品マスターCSVをSheets更新（base64経由・旧Power Automate用）
// ※ Google Drive直接読み込みに移行したため通常は使わない
// ============================================================
function updateProductMaster(base64Data) {
  if (!base64Data) return { success: false, error: 'dataが空です' };

  const bytes   = Utilities.base64Decode(base64Data);
  const blob    = Utilities.newBlob(bytes, 'text/plain', 'products.csv');
  const csvText = blob.getDataAsString('Shift-JIS');

  const rows = parseCSVText(csvText);
  if (rows.length < 2) return { success: false, error: 'CSVが空か1行のみです' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sh) sh = ss.insertSheet(SHEET_PRODUCTS);

  sh.clearContents();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  const updatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  PropertiesService.getScriptProperties().setProperty('PM_UPDATED_AT', updatedAt);

  return { success: true, rows: rows.length - 1, updatedAt };
}

// ============================================================
// ヘルパー: CSVテキストを行×列の2次元配列に変換
// ============================================================
function parseCSVText(text) {
  const rows = [];
  const lines = text.split('\n');
  lines.forEach(line => {
    const clean = line.replace(/\r/g, '');
    if (clean.trim() === '') return;
    rows.push(splitCSVLineGAS(clean));
  });
  return rows;
}

function splitCSVLineGAS(line) {
  const result = []; let cur = '', inQ = false;
  for (const ch of line) {
    if (ch === '"') { inQ = !inQ; }
    else if (ch === ',' && !inQ) { result.push(cur); cur = ''; }
    else { cur += ch; }
  }
  result.push(cur);
  return result;
}

function findColIdxGAS(headers, name) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === name || headers[i].indexOf(name) !== -1) return i;
  }
  return -1;
}

// ============================================================
// GET: 発注履歴取得（直近20件・新しい順）
//      supplierCode を指定するとそのメーカーの直近20件を返す
//      + 取得した20件の明細から商品ごとの最新注文情報（productHistory）を返す
// レスポンス: { success: true, orders: [...], productHistory: {...} }
// ============================================================
function getOrders(filterSupplierCode) {
  filterSupplierCode = String(filterSupplierCode || '').trim();

  const sh   = getSheet(SHEET_HISTORY);
  const data = sh.getDataRange().getValues();

  if (data.length <= 1) return { success: true, orders: [], productHistory: {} };

  // メーカー指定がある場合は先にフィルターしてから直近20件を取得する
  const orders = data.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
    .filter(r => !filterSupplierCode || String(r[2] || '').trim() === filterSupplierCode)
    .map(r => ({
      orderNo:      String(r[0] || ''),
      // r[1] は発注日。スプレッドシートが日付型として認識するため cellToStr で変換
      date:         cellToStr(r[1], 'yyyy/MM/dd'),
      supplierCode: String(r[2] || ''),
      supplierName: String(r[3] || ''),
      fax:          String(r[4] || ''),
      staff:        String(r[5] || ''),
      itemCount:    r[6] || 0,
      outputType:   String(r[7] || ''),
      // r[8] は登録日時。同じく cellToStr で変換（デフォルト: 日時フォーマット）
      createdAt:    cellToStr(r[8])
    }))
    .reverse()
    .slice(0, 20);

  // 直近20件の発注明細から商品ごとの最新注文情報を構築
  // 追加のAPI通信なし（発注明細シートをここで一括読込）
  const orderNos     = new Set(orders.map(o => o.orderNo));
  const orderDateMap = {};
  orders.forEach(o => { orderDateMap[o.orderNo] = o.date; });

  const itemsSh   = getSheet(SHEET_ITEMS);
  const itemsData = itemsSh.getDataRange().getValues();
  const productHistory = {}; // キー: 商品コードまたはJANコード → { date, qty, unit }

  itemsData.slice(1).forEach(r => {
    const orderNo = String(r[0] || '').trim();
    if (!orderNos.has(orderNo)) return; // 直近20件以外はスキップ

    const jan  = String(r[1] || '').trim();
    const code = String(r[2] || '').trim();
    const qty  = r[4] || 0;
    const unit = String(r[5] || '').trim();
    const date = orderDateMap[orderNo] || '';

    // 商品コードとJANコードの両方をキーとして登録（より新しい日付で上書き）
    [code, jan].filter(Boolean).forEach(key => {
      if (!productHistory[key] || date > productHistory[key].date) {
        productHistory[key] = { date, qty, unit };
      }
    });
  });

  return { success: true, orders, productHistory };
}

// ============================================================
// GET: 発注明細取得
// レスポンス: { success: true, items: [...] }
// ============================================================
function getOrderDetail(orderNo) {
  if (!orderNo) return { success: false, error: 'orderNoが未指定です' };
  const sh   = getSheet(SHEET_ITEMS);
  const data = sh.getDataRange().getValues();

  const items = data.slice(1)
    .filter(r => String(r[0]).trim() === String(orderNo).trim())
    .map(r => ({
      jan:           String(r[1] || ''),
      code:          String(r[2] || ''),
      name:          String(r[3] || ''),
      qty:           r[4] || 0,
      unit:          String(r[5] || ''),
      memo:          String(r[6] || ''),
      isHandwritten: r[7] === 'TRUE'
    }));

  return { success: true, items };
}

// ============================================================
// POST: 発注保存
// user_id: validateSession() から取得した実際のユーザーID（改ざん不可）
// ============================================================
function saveOrder(p, user_id) {
  const date                = p.date         || '';
  const supplierCode        = p.supplierCode || '';
  const supplierName        = p.supplierName || '';
  const fax                 = p.fax          || '';
  const staff               = p.staff        || '';
  const outputType          = p.outputType   || '';
  const items               = JSON.parse(p.items || '[]');
  // 修正発注の場合は元の発注Noが渡される（保存後に削除する）
  const revisionBaseOrderNo = String(p.revisionBaseOrderNo || '').trim();

  if (!date || !supplierCode || !supplierName || !staff) {
    return { success: false, error: '必須項目が不足しています (date, supplierCode, supplierName, staff)' };
  }

  const orderNo = generateOrderNo(date);
  const now     = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  const histSh = getSheet(SHEET_HISTORY);
  // user_id をサーバー側から記録（フロントから渡されたstaffとは別に監査用として保持）
  histSh.appendRow([orderNo, date, supplierCode, supplierName, fax, staff, items.length, outputType, now, user_id]);

  const itemsSh = getSheet(SHEET_ITEMS);
  items.forEach(item => {
    itemsSh.appendRow([
      orderNo,
      item.janCode       || '',
      item.code          || '',
      item.name          || '',
      item.qty           || 0,
      item.unit          || '',
      item.memo          || '',
      item.isHandwritten ? 'TRUE' : 'FALSE',
      now
    ]);
  });

  // 修正発注の場合：新規保存が完了してから元の発注を削除する
  if (revisionBaseOrderNo) {
    try {
      // 発注履歴から削除
      const histData = histSh.getDataRange().getValues();
      for (let i = histData.length - 1; i >= 1; i--) {
        if (String(histData[i][0]).trim() === revisionBaseOrderNo) {
          histSh.deleteRow(i + 1);
          break;
        }
      }
      // 発注明細から削除
      const itemsData = itemsSh.getDataRange().getValues();
      for (let i = itemsData.length - 1; i >= 1; i--) {
        if (String(itemsData[i][0]).trim() === revisionBaseOrderNo) {
          itemsSh.deleteRow(i + 1);
        }
      }
    } catch(e) {
      // 削除エラーは無視して保存成功として返す
      Logger.log('修正前履歴削除エラー（無視）: ' + e);
    }
  }

  return { success: true, orderNo };
}

// ============================================================
// POST: 発注履歴削除（管理者のみ）
// 発注履歴シートと発注明細シートから該当orderNoの行を物理削除する
// ============================================================
function deleteOrder(p, user_id) {
  const orderNo = String(p.orderNo || '').trim();
  if (!orderNo) return { success: false, error: '発注Noが未指定です' };

  // 管理者チェック
  if (!getIsAdmin(user_id)) {
    return { success: false, error: '削除は管理者のみ実行できます' };
  }

  // 発注履歴シートから削除（1行）
  const histSh   = getSheet(SHEET_HISTORY);
  const histData = histSh.getDataRange().getValues();
  let deletedHist = false;
  for (let i = histData.length - 1; i >= 1; i--) {
    if (String(histData[i][0]).trim() === orderNo) {
      histSh.deleteRow(i + 1);
      deletedHist = true;
      break; // 発注Noはユニーク
    }
  }
  if (!deletedHist) {
    return { success: false, error: '発注No「' + orderNo + '」が見つかりません' };
  }

  // 発注明細シートから削除（複数行）
  const itemsSh   = getSheet(SHEET_ITEMS);
  const itemsData = itemsSh.getDataRange().getValues();
  let deletedCount = 0;
  // 後ろから削除しないと行番号がズレる
  for (let i = itemsData.length - 1; i >= 1; i--) {
    if (String(itemsData[i][0]).trim() === orderNo) {
      itemsSh.deleteRow(i + 1);
      deletedCount++;
    }
  }

  return { success: true, deletedItems: deletedCount };
}

// ============================================================
// ヘルパー: ユーザーが管理者かどうかを確認
// beaufield-auth の users シートの列F（is_admin）を参照する
// ============================================================
function getIsAdmin(user_id) {
  try {
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const sh = ss.getSheetByName('users');
    if (!sh) return false;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(user_id).trim()) {
        // 列F（index 5）が TRUE または 'TRUE' の場合に管理者と判定
        return data[i][5] === true || String(data[i][5]).toUpperCase() === 'TRUE';
      }
    }
  } catch(e) {
    Logger.log('管理者チェックエラー: ' + e);
  }
  return false;
}

// 発注No採番（YYYYMMDD-NNN）
function generateOrderNo(dateStr) {
  const dateKey = dateStr.replace(/-/g, '');
  const sh      = getSheet(SHEET_HISTORY);
  const data    = sh.getDataRange().getValues();
  let maxSeq = 0;
  data.slice(1).forEach(r => {
    const no = String(r[0] || '');
    if (no.startsWith(dateKey + '-')) {
      const seq = parseInt(no.split('-')[1]) || 0;
      if (seq > maxSeq) maxSeq = seq;
    }
  });
  return dateKey + '-' + String(maxSeq + 1).padStart(3, '0');
}

// ============================================================
// POST: 発注先マスター操作
// ============================================================
function saveSupplier(p, user_id) {
  if (!getIsAdmin(user_id)) return { success: false, error: 'FORBIDDEN', message: '管理者権限が必要です' };
  const mode          = p.mode          || '';
  const code          = String(p.code          || '').trim();
  const name          = String(p.name          || '').trim();
  const fax           = String(p.fax           || '').trim();
  const outputMethods = String(p.outputMethods || '').trim(); // カンマ区切り文字列で受け取る

  if (!code) return { success: false, error: 'コードが未入力です' };

  const sh   = getSheet(SHEET_SUPPLIERS);
  const data = sh.getDataRange().getValues();
  const now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const exists = data.slice(1).some(r => String(r[0]).trim() === code);
    if (exists) return { success: false, error: 'コード「' + code + '」はすでに登録されています' };
    sh.appendRow([code, name, fax, now, outputMethods]);
    return { success: true };
  } else if (mode === 'update') {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.getRange(i + 1, 1, 1, 5).setValues([[code, name, fax, now, outputMethods]]);
        return { success: true };
      }
    }
    return { success: false, error: 'コード「' + code + '」が見つかりません' };
  } else if (mode === 'delete') {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'コード「' + code + '」が見つかりません' };
  } else {
    return { success: false, error: '不明なmode: ' + mode };
  }
}

// ============================================================
// POST: 担当者マスター操作
// ============================================================
function saveStaff(p, user_id) {
  if (!getIsAdmin(user_id)) return { success: false, error: 'FORBIDDEN', message: '管理者権限が必要です' };
  const mode = p.mode || '';
  const name = String(p.name || '').trim();

  if (!name) return { success: false, error: '担当者名が未入力です' };

  const sh   = getSheet(SHEET_STAFF);
  const data = sh.getDataRange().getValues();
  const now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const exists = data.slice(1).some(r => String(r[0]).trim() === name);
    if (exists) return { success: false, error: '「' + name + '」はすでに登録されています' };
    sh.appendRow([name, now]);
    return { success: true };
  } else if (mode === 'delete') {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === name) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: '「' + name + '」が見つかりません' };
  } else {
    return { success: false, error: '不明なmode: ' + mode };
  }
}

// ============================================================
// 初期設定関数（GASエディタから【1回だけ】手動実行）
// ============================================================
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function ensureSheet(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); Logger.log('シートを作成しました: ' + name); }
    if (sh.getLastRow() === 0) { sh.appendRow(headers); Logger.log('ヘッダーを設定しました: ' + name); }
    return sh;
  }

  ensureSheet(SHEET_HISTORY,  ['発注No','発注日','発注先コード','発注先名','FAX番号','担当者','品目数','出力方法','登録日時','user_id']);
  ensureSheet(SHEET_ITEMS,    ['発注No','JANコード','Beaufieldコード','商品名','数量','単位','備考','手書きフラグ','登録日時']);
  ensureSheet(SHEET_REORDER,  ['商品コード','月平均出荷数','更新日時']);

  const suppSh = ensureSheet(SHEET_SUPPLIERS, ['コード','名称','FAX','更新日時','発注方法']);
  if (suppSh.getLastRow() <= 1) {
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    suppSh.getRange(2, 1, 9, 4).setValues([
      ['10', 'デミコスメティクス',     '',             now],
      ['15', 'アペティート化粧品',      '052-883-1222', now],
      ['24', 'シュワルツコフ',          '',             now],
      ['48', '千代田化学',             '',             now],
      ['62', 'プレジール',             '0948-24-9801', now],
      ['67', 'ナプラ',                 '',             now],
      ['77', 'earth walk republic',   '078-200-6869', now],
      ['81', 'GO-ON',                 '078-200-6678', now],
      ['58', 'パシフィックプロダクツ', '03-5299-0435', now],
    ]);
    Logger.log('発注先マスターに初期データを登録しました');
  }

  const staffSh = ensureSheet(SHEET_STAFF, ['名前','更新日時']);
  if (staffSh.getLastRow() <= 1) {
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    staffSh.appendRow(['前島', now]);
    Logger.log('担当者マスターに初期データを登録しました');
  }

  Logger.log('✅ 初期設定完了 (Version: ' + VERSION + ')');
}

// ============================================================
// メーカー発注書テンプレート配信（Drive経由・認証付き）
// ============================================================
// 用途: ブラウザがGitHub Pages経由で公開PDFをfetchするのを廃止し、
//      Driveに保管したテンプレートをセッション認証経由で配信する。
//
// 前提:
//   - Driveに「Beaufield発注書テンプレート」フォルダを作成し、
//     スクリプトプロパティ ORDER_TEMPLATE_FOLDER_ID にIDを設定
//   - フォルダ内に ORDER_TEMPLATES で定義した名前の PDF をアップロード
//   - GASの実行ユーザー設定が「自分」になっていれば、利用者側のGoogle権限は不要
//
// レスポンス:
//   { success: true, filename, mimeType: 'application/pdf', data: <base64> }
//   または { success: false, error }
// ============================================================
function getOrderTemplate(makerKey) {
  const fileName = ORDER_TEMPLATES[makerKey];
  if (!fileName) {
    return { success: false, error: '不明なメーカーキー: ' + makerKey };
  }

  const folderId = _PROPS.getProperty('ORDER_TEMPLATE_FOLDER_ID');
  if (!folderId) {
    return { success: false, error: 'ORDER_TEMPLATE_FOLDER_ID 未設定（GASスクリプトプロパティ）' };
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files  = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      return { success: false, error: 'テンプレートPDFが見つかりません: ' + fileName };
    }

    const file   = files.next();
    const blob   = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    return {
      success: true,
      filename: fileName,
      mimeType: 'application/pdf',
      data: base64
    };
  } catch(err) {
    return { success: false, error: 'テンプレート取得エラー: ' + err.toString() };
  }
}
