// ============================================================
// Beaufield 発注アプリ - Google Apps Script バックエンド
// Version: v1.8.0
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   SPREADSHEET_ID    : 発注管理データのスプレッドシートID
//   AUTH_SHEET_ID     : beaufield-auth スプレッドシートID（共通）
//   REORDER_API_KEY   : 発注点・発注提案更新用APIキー（Pythonスクリプトと共有）
//   LINEWORKS_WEBHOOK : LINE WORKS Incoming Webhook URL（任意・発注提案の通知用）
//
// 発注先マスターの拡張列（シートに直接入力・空欄なら既定値7日）:
//   F列 = リードタイム(日)   発注してから入荷するまでの日数
//   G列 = 発注サイクル(日)   そのメーカーへ発注する間隔（週1なら7）
//
// ============================================================

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS          = PropertiesService.getScriptProperties();
const SPREADSHEET_ID  = _PROPS.getProperty('SPREADSHEET_ID');
const AUTH_SHEET_ID   = _PROPS.getProperty('AUTH_SHEET_ID');
const UPDATE_SECRET   = _PROPS.getProperty('UPDATE_SECRET');   // 商品マスター更新用（Power Automate連携）
const VERSION         = 'v1.14.0';
const CACHE_TTL_SESSION = 900; // 15分（CacheService保持秒数・セッション検証の高速化用）

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
const SHEET_PROPOSALS = '発注提案';       // analyze_demand.py が週次で書き込む発注提案リスト
const SHEET_PROPOSAL_EXCL = '提案除外設定'; // 発注提案の対象外にする商品（カタログ等のノイズ除外用）
const SHEET_EXCESS     = '過剰在庫';       // analyze_demand.py が週次で書き込む過剰在庫リスト
const SHEET_EXCESS_ACK = '過剰在庫確認済み'; // 持ちすぎを承知の上で確認済みにした商品（意図的なまとめ仕入等）

// 発注先マスターの拡張列（F=リードタイム(日), G=発注サイクル(日)）
// 未入力の場合のフォールバック値。analyze_demand.py の計算に使う
const DEFAULT_LEAD_TIME_DAYS  = 7;
const DEFAULT_ORDER_CYCLE_DAYS = 7;

// メーカー発注書テンプレート定義（Drive配信用）
// スクリプトプロパティ ORDER_TEMPLATE_FOLDER_ID に Drive フォルダIDを設定すること
const ORDER_TEMPLATES = {
  'grandex':           '54.pdf',
  'chiyoda':           '48.pdf',
  'alpenrose':         '57.pdf',
  'melos':             '2.pdf',
  'melos_2025':        'メロス発注書2025年価格改定後.pdf',
  'adelans':           '82.pdf',
  'rhythm':            'リズム注文書2023冬～.pdf',
  'hokkaido_natural':  '北海道ナチュラルバイオ.pdf'
};

// ============================================================
// セッション検証
// beaufield-auth の sessions シートでトークンを照合する
// CacheService で 15 分間キャッシュしてシート読み込みを削減する
// （ログアウト即時反映が必要な運用ではない前提。他アプリと同一パターン）
// 戻り値: { valid: true, user_id } または { valid: false }
// ============================================================
function validateSession(token) {
  if (!token) return { valid: false };

  const cache    = CacheService.getScriptCache();
  const cacheKey = 'sess_' + token.slice(-32);
  const cached   = cache.get(cacheKey);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch(e) {}
  }

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
          const r = { valid: false };
          cache.put(cacheKey, JSON.stringify(r), 60);
          return r;
        }
        const r = { valid: true, user_id: rowUserId };
        cache.put(cacheKey, JSON.stringify(r), CACHE_TTL_SESSION);
        return r;
      }
    }
  } catch(e) {
    Logger.log('セッション検証エラー: ' + e);
  }
  const r = { valid: false };
  cache.put(cacheKey, JSON.stringify(r), 60);
  return r;
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
      case 'getOrderProposals': return jsonResponse(getOrderProposals());
      default:                  return jsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch(err) {
    Logger.log('doGet error: ' + err);
    return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
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
      Logger.log('updateProductMaster error: ' + err);
      return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
    }
  }

  // APIキー認証アクション（Pythonスクリプト用・セッション不要）
  const API_KEY_ACTIONS = ['updateReorderPoints', 'getReorderConfig', 'updateOrderProposals', 'updateProposalExplanations', 'testNotify'];
  if (API_KEY_ACTIONS.indexOf(action) !== -1) {
    const apiKey = p.api_key || '';
    if (!apiKey || apiKey !== _PROPS.getProperty('REORDER_API_KEY')) {
      return jsonResponse({ success: false, error: 'UNAUTHORIZED' });
    }
    try {
      switch (action) {
        case 'updateReorderPoints':        return jsonResponse(updateReorderPoints(p.products || []));
        case 'getReorderConfig':           return jsonResponse(getReorderConfig());
        case 'updateOrderProposals':       return jsonResponse(updateOrderProposals(p));
        case 'updateProposalExplanations': return jsonResponse(updateProposalExplanations(p));
        case 'testNotify':                 return jsonResponse(testNotify());
      }
    } catch(err) {
      Logger.log(action + ' error: ' + err);
      return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
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
      case 'saveProposalExclusion': return jsonResponse(saveProposalExclusion(p, auth.user_id));
      case 'saveExcessAck': return jsonResponse(saveExcessAck(p, auth.user_id));
      default:             return jsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch(err) {
    Logger.log('doPost error: ' + err);
    return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
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
// SpreadsheetApp.openById() は1リクエスト内で複数回呼ぶとオーバーヘッドになるため
// 実行コンテキスト内でキャッシュして使い回す
// ============================================================
let _ssCache = null;
function getSS() {
  if (!_ssCache) _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ssCache;
}
function getSheet(name) {
  const sh = getSS().getSheetByName(name);
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
  const ss = getSS();
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

  // 同時保存による発注No重複を防ぐため、採番〜履歴書き込みをロックで保護
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, error: '現在別の処理が実行中です。数秒後に再度お試しください。' };
  }

  try {
    const orderNo = generateOrderNo(date);
    const now     = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    const histSh = getSheet(SHEET_HISTORY);
    // user_id をサーバー側から記録（フロントから渡されたstaffとは別に監査用として保持）
    histSh.appendRow([orderNo, date, supplierCode, supplierName, fax, staff, items.length, outputType, now, user_id]);

    const itemsSh = getSheet(SHEET_ITEMS);
    if (items.length > 0) {
      const rows = items.map(item => [
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
      itemsSh.getRange(itemsSh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
    }

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
        Logger.log('修正前履歴削除エラー（無視）: ' + e);
      }
    }

    return { success: true, orderNo };
  } finally {
    lock.releaseLock();
  }
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
// 発注提案機能（analyze_demand.py 連携）
// ============================================================
// フロー:
//   1. analyze_demand.py が getReorderConfig で設定取得
//      （発注先別リードタイム・除外商品・過去発注からのロット推定材料）
//   2. 売上データを分析して updateOrderProposals で「発注提案」「過剰在庫」シートを更新
//      提案・過剰在庫があれば LINEWORKS_WEBHOOK（スクリプトプロパティ・任意）へ通知
//   3. アプリは getOrderProposals で提案・過剰在庫を表示
//   4. 不要な商品は saveProposalExclusion で除外登録（以後提案されない）
//   5. 持ちすぎを承知の上の商品は saveExcessAck で確認済み登録（次回分析後もリストから消える）
//   6. Claude Code が updateProposalExplanations でAI説明を追記（任意）
// ============================================================

// 発注提案シートの列定義（updateOrderProposals / getOrderProposals で共有）
const PROPOSAL_HEADERS = ['商品コード','商品名','仕入先コード','仕入先名','パターン',
                          '現在庫','発注済','推奨在庫','提案数量','推定ロット','月平均',
                          '注文P95','最大注文','根拠メモ','AI説明','分析日時','仕入単価','提案金額','ABCランク'];

// 過剰在庫シートの列定義（updateOrderProposals / getOrderProposals で共有）
const EXCESS_HEADERS = ['商品コード','商品名','仕入先コード','仕入先名','現在庫','推奨在庫',
                        '過剰数量','仕入単価','過剰金額','在庫月数','ABCランク','パターン','分析日時'];

// POST(APIキー): 分析に必要な設定を返す
// レスポンス: { success, suppliers: [{code,name,leadTimeDays,orderCycleDays}],
//              exclusions: [商品コード], lotStats: {code: {orderCount,minQty,gcdQty}} }
function getReorderConfig() {
  // 発注先マスター（F列=リードタイム(日), G列=発注サイクル(日)。未入力はnull→Python側で既定値）
  const suppData = getSheet(SHEET_SUPPLIERS).getDataRange().getValues();
  const suppliers = suppData.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      code:           String(r[0]).trim(),
      name:           String(r[1]).trim(),
      leadTimeDays:   parseFloat(r[5]) > 0 ? parseFloat(r[5]) : null,
      orderCycleDays: parseFloat(r[6]) > 0 ? parseFloat(r[6]) : null
    }));

  // 提案除外設定（シート未作成なら空）
  let exclusions = [];
  const exSh = getSS().getSheetByName(SHEET_PROPOSAL_EXCL);
  if (exSh && exSh.getLastRow() > 1) {
    exclusions = exSh.getDataRange().getValues().slice(1)
      .map(r => String(r[0]).trim()).filter(Boolean);
  }

  // 過去の発注明細から商品別のロット推定材料を集計
  // gcdQty: 全発注数量の最大公約数（ケース単位の推定に使う）
  // あわせて直近60日の発注（発注済み・未入荷の可能性がある分）も商品別に返す
  // 発注日は発注No（YYYYMMDD-NNN）の先頭8桁から取得
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 60);
  const cutoffStr = Utilities.formatDate(cutoff, 'Asia/Tokyo', 'yyyyMMdd');

  const itemsData = getSheet(SHEET_ITEMS).getDataRange().getValues();
  const lotStats = {};
  const recentOrders = {};
  itemsData.slice(1).forEach(r => {
    const code = String(r[2] || '').trim();
    const qty  = Math.round(parseFloat(r[4]) || 0);
    if (!code || qty <= 0) return;
    if (!lotStats[code]) lotStats[code] = { orderCount: 0, minQty: qty, gcdQty: 0 };
    const s = lotStats[code];
    s.orderCount++;
    if (qty < s.minQty) s.minQty = qty;
    s.gcdQty = gcdInt(s.gcdQty, qty);

    const orderDate = String(r[0] || '').split('-')[0];  // 発注No先頭のYYYYMMDD
    if (orderDate.length === 8 && orderDate >= cutoffStr) {
      if (!recentOrders[code]) recentOrders[code] = [];
      recentOrders[code].push({ date: orderDate, qty });
    }
  });

  return {
    success: true,
    suppliers,
    exclusions,
    lotStats,
    recentOrders,
    defaults: { leadTimeDays: DEFAULT_LEAD_TIME_DAYS, orderCycleDays: DEFAULT_ORDER_CYCLE_DAYS }
  };
}

function gcdInt(a, b) {
  a = Math.abs(a); b = Math.abs(b);
  while (b) { const t = b; b = a % b; a = t; }
  return a;
}

// 金額を「◯◯.◯万円」表記にする（LINE WORKS通知用）
function formatManYen(yen) {
  return (yen / 10000).toFixed(1) + '万円';
}

// POST(APIキー): 発注提案・過剰在庫シートを全面書き換え
// リクエスト: { proposals: [{code,name,supplierCode,supplierName,pattern,stock,recommended,
//              proposedQty,lot,meanMonthly,p95Order,maxOrder,note}],
//              excess: [{code,name,supplierCode,supplierName,stock,recommended,excessQty,
//              unitCost,excessAmount,monthsOfStock,abcRank,pattern}], analyzedAt }
function updateOrderProposals(p) {
  const proposals  = p.proposals || [];
  const excess     = p.excess || [];
  const analyzedAt = String(p.analyzedAt || '');

  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_PROPOSALS);
  if (!sh) sh = ss.insertSheet(SHEET_PROPOSALS);

  sh.clearContents();
  sh.getRange(1, 1, 1, PROPOSAL_HEADERS.length).setValues([PROPOSAL_HEADERS]);

  if (proposals.length > 0) {
    const rows = proposals.map(x => [
      String(x.code         || '').trim(),
      String(x.name         || ''),
      String(x.supplierCode || ''),
      String(x.supplierName || ''),
      String(x.pattern      || ''),
      parseFloat(x.stock)       || 0,
      parseFloat(x.onOrder)     || 0,
      parseFloat(x.recommended) || 0,
      parseFloat(x.proposedQty) || 0,
      parseFloat(x.lot)         || 1,
      parseFloat(x.meanMonthly) || 0,
      parseFloat(x.p95Order)    || 0,
      parseFloat(x.maxOrder)    || 0,
      String(x.note || ''),
      '',           // AI説明（updateProposalExplanations で追記）
      analyzedAt,
      parseFloat(x.unitCost) || 0,
      parseFloat(x.amount)   || 0,
      String(x.abcRank || '')
    ]);
    sh.getRange(2, 1, rows.length, PROPOSAL_HEADERS.length).setValues(rows);
  }

  // 過剰在庫シート（全面書き換え）
  let exSh = ss.getSheetByName(SHEET_EXCESS);
  if (!exSh) exSh = ss.insertSheet(SHEET_EXCESS);
  exSh.clearContents();
  exSh.getRange(1, 1, 1, EXCESS_HEADERS.length).setValues([EXCESS_HEADERS]);
  if (excess.length > 0) {
    const exRows = excess.map(x => [
      String(x.code         || '').trim(),
      String(x.name         || ''),
      String(x.supplierCode || ''),
      String(x.supplierName || ''),
      parseFloat(x.stock)         || 0,
      parseFloat(x.recommended)   || 0,
      parseFloat(x.excessQty)     || 0,
      parseFloat(x.unitCost)      || 0,
      parseFloat(x.excessAmount)  || 0,
      parseFloat(x.monthsOfStock) || 0,
      String(x.abcRank || ''),
      String(x.pattern || ''),
      analyzedAt
    ]);
    exSh.getRange(2, 1, exRows.length, EXCESS_HEADERS.length).setValues(exRows);
  }

  // LINE WORKS通知（LINEWORKS_WEBHOOK 未設定なら何もしない）
  if (proposals.length > 0) {
    const bySupplier = {};
    proposals.forEach(x => {
      const key = String(x.supplierName || '不明');
      bySupplier[key] = (bySupplier[key] || 0) + 1;
    });
    const lines = Object.keys(bySupplier)
      .sort((a, b) => bySupplier[b] - bySupplier[a])
      .slice(0, 8)
      .map(name => '・' + name + ': ' + bySupplier[name] + '件');
    const more = Object.keys(bySupplier).length > 8 ? '\n…ほか' + (Object.keys(bySupplier).length - 8) + '社' : '';
    const totalAmount = proposals.reduce((s, x) => s + (parseFloat(x.amount) || 0), 0);
    const excessAmount = excess.reduce((s, x) => s + (parseFloat(x.excessAmount) || 0), 0);
    const excessLine = excess.length > 0
      ? ('\n⚠️ 持ちすぎ在庫: ' + excess.length + '件・過剰' + formatManYen(excessAmount))
      : '';
    notifyLineWorks(
      '📋 発注提案の分析が完了しました（' + analyzedAt + '）\n' +
      '提案: ' + proposals.length + '件・合計' + formatManYen(totalAmount) + '\n' + lines.join('\n') + more +
      excessLine + '\n' +
      '発注アプリの「発注提案」タブで確認してください。'
    );
  }

  Logger.log('✅ 発注提案更新完了: ' + proposals.length + '件 / 過剰在庫 ' + excess.length + '件');
  return { success: true, count: proposals.length, excessCount: excess.length };
}

// POST(APIキー): AI説明の追記（Claude Code から実行）
// リクエスト: { explanations: [{code, text}] }
function updateProposalExplanations(p) {
  const explanations = p.explanations || [];
  if (explanations.length === 0) return { success: true, count: 0 };

  const sh = getSS().getSheetByName(SHEET_PROPOSALS);
  if (!sh || sh.getLastRow() < 2) return { success: false, error: '発注提案シートが空です' };

  const data = sh.getDataRange().getValues();
  const colAi = PROPOSAL_HEADERS.indexOf('AI説明');  // 0-based
  const map = {};
  explanations.forEach(x => { map[String(x.code || '').trim()] = String(x.text || ''); });

  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][0]).trim();
    if (map[code] !== undefined) {
      sh.getRange(i + 1, colAi + 1).setValue(map[code]);
      count++;
    }
  }
  return { success: true, count };
}

// GET(セッション): 発注提案の取得（アプリの発注提案タブ用）
function getOrderProposals() {
  const ss = getSS();
  const sh = ss.getSheetByName(SHEET_PROPOSALS);
  let proposals = [];
  let analyzedAt = '';
  if (sh && sh.getLastRow() > 1) {
    const data = sh.getDataRange().getValues();
    proposals = data.slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:         String(r[0]).trim(),
        name:         String(r[1] || ''),
        supplierCode: String(r[2] || '').trim(),
        supplierName: String(r[3] || ''),
        pattern:      String(r[4] || ''),
        stock:        parseFloat(r[5])  || 0,
        onOrder:      parseFloat(r[6])  || 0,
        recommended:  parseFloat(r[7])  || 0,
        proposedQty:  parseFloat(r[8])  || 0,
        lot:          parseFloat(r[9])  || 1,
        meanMonthly:  parseFloat(r[10]) || 0,
        p95Order:     parseFloat(r[11]) || 0,
        maxOrder:     parseFloat(r[12]) || 0,
        note:         String(r[13] || ''),
        aiNote:       String(r[14] || ''),
        unitCost:     parseFloat(r[16]) || 0,
        amount:       parseFloat(r[17]) || 0,
        abcRank:      String(r[18] || '')
      }));
    analyzedAt = cellToStr(data[1][15], 'yyyy-MM-dd HH:mm');
  }

  // 除外リスト（除外管理UIでの表示・解除用）
  let exclusions = [];
  const exSh = ss.getSheetByName(SHEET_PROPOSAL_EXCL);
  if (exSh && exSh.getLastRow() > 1) {
    exclusions = exSh.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:   String(r[0]).trim(),
        name:   String(r[1] || ''),
        reason: String(r[2] || ''),
        addedBy: String(r[3] || ''),
        addedAt: cellToStr(r[4], 'yyyy-MM-dd HH:mm')
      }));
  }

  // 過剰在庫の確認済みリスト（先に読み、確認済みの商品を過剰在庫リストから除外する）
  let excessAcks = [];
  const ackSh = ss.getSheetByName(SHEET_EXCESS_ACK);
  if (ackSh && ackSh.getLastRow() > 1) {
    excessAcks = ackSh.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:    String(r[0]).trim(),
        name:    String(r[1] || ''),
        reason:  String(r[2] || ''),
        ackedBy: String(r[3] || ''),
        ackedAt: cellToStr(r[4], 'yyyy-MM-dd HH:mm')
      }));
  }
  const ackedCodes = new Set(excessAcks.map(x => x.code));

  // 過剰在庫リスト（確認済みは除く）
  let excess = [];
  const excSh = ss.getSheetByName(SHEET_EXCESS);
  if (excSh && excSh.getLastRow() > 1) {
    excess = excSh.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:          String(r[0]).trim(),
        name:          String(r[1] || ''),
        supplierCode:  String(r[2] || '').trim(),
        supplierName:  String(r[3] || ''),
        stock:         parseFloat(r[4]) || 0,
        recommended:   parseFloat(r[5]) || 0,
        excessQty:     parseFloat(r[6]) || 0,
        unitCost:      parseFloat(r[7]) || 0,
        excessAmount:  parseFloat(r[8]) || 0,
        monthsOfStock: parseFloat(r[9]) || 0,
        abcRank:       String(r[10] || ''),
        pattern:       String(r[11] || '')
      }))
      .filter(x => !ackedCodes.has(x.code));
  }

  return { success: true, proposals, analyzedAt, exclusions, excess, excessAcks };
}

// POST(セッション): 過剰在庫の確認済み登録/解除
// リクエスト: { action:'saveExcessAck', mode:'add'|'delete', code, name, reason }
function saveExcessAck(p, user_id) {
  const mode   = p.mode || '';
  const code   = String(p.code   || '').trim();
  const name   = String(p.name   || '').trim();
  const reason = String(p.reason || '').trim();
  if (!code) return { success: false, error: '商品コードが未指定です' };

  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_EXCESS_ACK);
  if (!sh) {
    sh = ss.insertSheet(SHEET_EXCESS_ACK);
    sh.appendRow(['商品コード','商品名','理由','確認者','確認日時']);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const data = sh.getDataRange().getValues();
    const exists = data.slice(1).some(r => String(r[0]).trim() === code);
    if (exists) return { success: false, error: '商品コード「' + code + '」はすでに確認済みです' };
    sh.appendRow([code, name, reason, user_id, now]);
    return { success: true };
  } else if (mode === 'delete') {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: '商品コード「' + code + '」が見つかりません' };
  }
  return { success: false, error: '不明なmode: ' + mode };
}

// POST(セッション): 提案除外の登録/解除
// リクエスト: { action:'saveProposalExclusion', mode:'add'|'delete', code, name, reason }
function saveProposalExclusion(p, user_id) {
  const mode   = p.mode || '';
  const code   = String(p.code   || '').trim();
  const name   = String(p.name   || '').trim();
  const reason = String(p.reason || '').trim();
  if (!code) return { success: false, error: '商品コードが未指定です' };

  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_PROPOSAL_EXCL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PROPOSAL_EXCL);
    sh.appendRow(['商品コード','商品名','理由','登録者','登録日時']);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const data = sh.getDataRange().getValues();
    const exists = data.slice(1).some(r => String(r[0]).trim() === code);
    if (exists) return { success: false, error: '商品コード「' + code + '」はすでに除外登録されています' };
    sh.appendRow([code, name, reason, user_id, now]);
    // 現在の提案シートからも即時削除（次回分析を待たずに消す）
    const prSh = ss.getSheetByName(SHEET_PROPOSALS);
    if (prSh && prSh.getLastRow() > 1) {
      const prData = prSh.getDataRange().getValues();
      for (let i = prData.length - 1; i >= 1; i--) {
        if (String(prData[i][0]).trim() === code) prSh.deleteRow(i + 1);
      }
    }
    return { success: true };
  } else if (mode === 'delete') {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: '商品コード「' + code + '」が見つかりません' };
  }
  return { success: false, error: '不明なmode: ' + mode };
}

// ヘルパー: LINE WORKS Incoming Webhook 通知（失敗しても本処理は継続）
// ペイロードは {body:{text}} 形式が正（LINE WORKS Incoming Webhook仕様）
// 戻り値: { sent, hasWebhook, httpStatus, responseBody } — testNotify のデバッグにも使う
function notifyLineWorks(text) {
  const result = { sent: false, hasWebhook: false, httpStatus: null, responseBody: '' };
  try {
    const webhook = _PROPS.getProperty('LINEWORKS_WEBHOOK');
    if (!webhook) return result;
    result.hasWebhook = true;
    const res = UrlFetchApp.fetch(webhook.trim(), {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ body: { text: text } }),
      muteHttpExceptions: true
    });
    result.httpStatus = res.getResponseCode();
    result.responseBody = String(res.getContentText() || '').slice(0, 300);
    result.sent = result.httpStatus >= 200 && result.httpStatus < 300;
    if (!result.sent) Logger.log('LINE WORKS通知 HTTPエラー: ' + result.httpStatus + ' ' + result.responseBody);
  } catch(e) {
    Logger.log('LINE WORKS通知エラー（無視）: ' + e);
    result.responseBody = String(e).slice(0, 300);
  }
  return result;
}

// POST(APIキー): LINE WORKS通知の疎通テスト（デバッグ用）
// レスポンスに webhook の設定有無・HTTPステータス・応答本文を返す
function testNotify() {
  const r = notifyLineWorks('🔔 発注アプリからのテスト通知です（testNotify）');
  return { success: true, ...r };
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

  let folderId = _PROPS.getProperty('ORDER_TEMPLATE_FOLDER_ID');
  if (!folderId) {
    return { success: false, error: 'ORDER_TEMPLATE_FOLDER_ID 未設定（GASスクリプトプロパティ）' };
  }
  // URLが設定されていた場合もIDを正しく抽出
  const m = folderId.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) folderId = m[1];

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
    Logger.log('getOrderTemplate error: ' + err);
    return { success: false, error: 'INTERNAL_ERROR' };
  }
}
