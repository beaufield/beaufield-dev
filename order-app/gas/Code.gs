// ============================================================
// Beaufield 発注アプリ - Google Apps Script バックエンド
// Version: v1.1.3
// ============================================================
// [重要] SPREADSHEET_ID を必ず自分のスプレッドシートIDに変更してください
//   → Googleスプレッドシートを開き、URLの /d/XXXXX/edit の
//     XXXXX 部分をコピーして下記に貼り付けてください
// ============================================================

const SPREADSHEET_ID  = 'ここにスプレッドシートIDを貼り付けてください';
const VERSION         = 'v1.1.3';

// シート名定数
const SHEET_HISTORY   = '発注履歴';
const SHEET_ITEMS     = '発注明細';
const SHEET_SUPPLIERS = '発注先マスター';
const SHEET_STAFF     = '担当者マスター';

// ============================================================
// エントリーポイント（GET）
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  try {
    switch (action) {
      case 'getMasters':      return jsonResponse(getMasters());
      case 'getOrders':       return jsonResponse(getOrders());
      case 'getOrderDetail':  return jsonResponse(getOrderDetail(e.parameter.orderNo));
      default:                return jsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch(err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
// エントリーポイント（POST）
// application/x-www-form-urlencoded 形式で受け取る
// ============================================================
function doPost(e) {
  const p      = (e && e.parameter) ? e.parameter : {};
  const action = p.action || '';
  try {
    switch (action) {
      case 'saveOrder':    return jsonResponse(saveOrder(p));
      case 'saveSupplier': return jsonResponse(saveSupplier(p));
      case 'saveStaff':    return jsonResponse(saveStaff(p));
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
      code: String(r[0]).trim(),
      name: String(r[1]).trim(),
      fax:  String(r[2] || '').trim()
    }));

  const staffData = staffSheet.getDataRange().getValues();
  const staff = staffData.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({ name: String(r[0]).trim() }));

  return { success: true, suppliers, staff };
}

// ============================================================
// GET: 発注履歴取得（直近20件・新しい順）
//      + 直近20件の明細から商品ごとの最新注文情報（productHistory）を返す
// レスポンス: { success: true, orders: [...], productHistory: {...} }
// ============================================================
function getOrders() {
  const sh   = getSheet(SHEET_HISTORY);
  const data = sh.getDataRange().getValues();

  if (data.length <= 1) return { success: true, orders: [], productHistory: {} };

  const orders = data.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
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
// ============================================================
function saveOrder(p) {
  const date         = p.date         || '';
  const supplierCode = p.supplierCode || '';
  const supplierName = p.supplierName || '';
  const fax          = p.fax          || '';
  const staff        = p.staff        || '';
  const outputType   = p.outputType   || '';
  const items        = JSON.parse(p.items || '[]');

  if (!date || !supplierCode || !supplierName || !staff) {
    return { success: false, error: '必須項目が不足しています (date, supplierCode, supplierName, staff)' };
  }

  const orderNo = generateOrderNo(date);
  const now     = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  const histSh = getSheet(SHEET_HISTORY);
  histSh.appendRow([orderNo, date, supplierCode, supplierName, fax, staff, items.length, outputType, now]);

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

  return { success: true, orderNo };
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
function saveSupplier(p) {
  const mode = p.mode || '';
  const code = String(p.code || '').trim();
  const name = String(p.name || '').trim();
  const fax  = String(p.fax  || '').trim();

  if (!code) return { success: false, error: 'コードが未入力です' };

  const sh   = getSheet(SHEET_SUPPLIERS);
  const data = sh.getDataRange().getValues();
  const now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const exists = data.slice(1).some(r => String(r[0]).trim() === code);
    if (exists) return { success: false, error: 'コード「' + code + '」はすでに登録されています' };
    sh.appendRow([code, name, fax, now]);
    return { success: true };
  } else if (mode === 'update') {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.getRange(i + 1, 1, 1, 4).setValues([[code, name, fax, now]]);
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
function saveStaff(p) {
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

  ensureSheet(SHEET_HISTORY, ['発注No','発注日','発注先コード','発注先名','FAX番号','担当者','品目数','出力方法','登録日時']);
  ensureSheet(SHEET_ITEMS,   ['発注No','JANコード','Beaufieldコード','商品名','数量','単位','備考','手書きフラグ','登録日時']);

  const suppSh = ensureSheet(SHEET_SUPPLIERS, ['コード','名称','FAX','更新日時']);
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
