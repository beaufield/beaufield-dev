// BCARTマスター管理ツール - バックエンド
//
// [スクリプトプロパティに設定が必要]
//   BCART_TOKEN       : BCARTアクセストークン
//   GEMINI_API_KEY    : Google Gemini APIキー
//   LINEWORKS_WEBHOOK : LINE WORKS Webhook URL（任意）
//   CSV_FOLDER_ID     : 商品.CSV保管Driveフォルダ ID
//   AUTH_GAS_URL      : portal GAS WebApp URL（セッション検証用）

const VERSION = 'v2.4.2';

// ===================== 設定 =====================
const BCART_BASE_URL = 'https://api.bcart.jp/api/v1';
const CSV_FILENAME = '商品.CSV';
const GEMINI_MODEL = 'gemini-2.5-flash-lite';

// スクリプトプロパティから機密値を取得（コード直書き禁止）
const _BCART_PROPS = PropertiesService.getScriptProperties();
const CSV_FOLDER_ID = _BCART_PROPS.getProperty('CSV_FOLDER_ID');
const AUTH_GAS_URL  = _BCART_PROPS.getProperty('AUTH_GAS_URL');

// シート名
const SHEET_IGNORE      = '対応不要';
const SHEET_WIP         = '作業中';
const SHEET_HISTORY     = '更新履歴';
const SHEET_SP_GROUPS   = '特別価格_顧客グループ';
const SHEET_SP_DETAILS  = '特別価格_明細';
const SHEET_VF_DETAILS  = '例外表示_明細';

// ===================== エントリポイント =====================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;

    const noAuthActions = ['getVersion'];
    let userName = '不明';
    if (!noAuthActions.includes(action)) {
      const authResult = validateSession(params.session);
      if (!authResult.ok) return jsonResponse({ ok: false, error: 'UNAUTHORIZED' });
      if (authResult.user) {
        userName = authResult.user.name || authResult.user.user_id || '不明';
      }
    }
    params._userName = userName;

    switch (action) {
      case 'getVersion':           return jsonResponse({ ok: true, version: VERSION });
      case 'loadData':             return jsonResponse(loadData());
      case 'updatePrice':          return jsonResponse(updatePrice(params));
      case 'updateJodai':          return jsonResponse(updateJodai(params));
      case 'updateJan':            return jsonResponse(updateJan(params));
      case 'updateAll':            return jsonResponse(updateAll(params));
      case 'hideProductSet':       return jsonResponse(hideProductSet(params));
      case 'setStock':             return jsonResponse(setStock(params));
      case 'markIgnore':           return jsonResponse(markIgnore(params));
      case 'unmarkIgnore':         return jsonResponse(unmarkIgnore(params));
      case 'getIgnoreList':        return jsonResponse(getIgnoreList());
      case 'markWip':              return jsonResponse(markWip(params));
      case 'unmarkWip':            return jsonResponse(unmarkWip(params));
      case 'bulkUpdate':           return jsonResponse(bulkUpdate(params));
      case 'bulkIgnore':           return jsonResponse(bulkIgnore(params));
      case 'bulkMarkWip':          return jsonResponse(bulkMarkWip(params));
      case 'searchProducts':       return jsonResponse(searchProducts(params));
      case 'getSpecials':          return jsonResponse(getSpecials());
      case 'updateProduct':        return jsonResponse(updateProductAction(params));
      case 'deleteProduct':        return jsonResponse(deleteProduct(params));
      case 'getHistory':           return jsonResponse(getHistory());
      case 'debugData':            return jsonResponse(debugData());
      case 'debugCode':            return jsonResponse(debugCode(params));
      case 'debugProduct':         return jsonResponse(debugProduct());
      // 機能A: 新規登録
      case 'getCategories':        return jsonResponse(getCategories());
      case 'getFeatures':          return jsonResponse(getSpecials());
      case 'registerProduct':      return jsonResponse(registerProduct(params));
      // 機能B: 特別価格管理
      case 'getSpecialPriceData':       return jsonResponse(getSpecialPriceData());
      case 'saveCustomerGroup':         return jsonResponse(saveCustomerGroup(params));
      case 'deleteCustomerGroup':       return jsonResponse(deleteCustomerGroup(params));
      case 'getProductSetsForFeature':  return jsonResponse(getProductSetsForFeature(params));
      case 'searchProductSets':         return jsonResponse(searchProductSets(params));
      case 'applyGroupPrices':          return jsonResponse(applyGroupPrices(params));
      case 'saveSpecialPriceDetails':   return jsonResponse(saveSpecialPriceDetails(params));
      case 'deleteSpecialPriceDetail':  return jsonResponse(deleteSpecialPriceDetail(params));
      case 'saveViewFilterDetails':     return jsonResponse(saveViewFilterDetails(params));
      case 'getSpecialPriceCurrent':    return jsonResponse(getSpecialPriceCurrent(params));
      case 'getViewFilterCurrent':      return jsonResponse(getViewFilterCurrent(params));
      case 'applyViewFilters':          return jsonResponse(applyViewFilters(params));
      case 'deleteViewFilterDetail':    return jsonResponse(deleteViewFilterDetail(params));
      case 'getMembers':                return jsonResponse(getMembers());
      // 機能C: 説明文生成
      case 'getProductsForDescription': return jsonResponse(getProductsForDescription());
      case 'generateDescription':       return jsonResponse(generateDescription(params));
      case 'factCheckDescription':      return jsonResponse(factCheckDescription(params));
      case 'applyDescription':          return jsonResponse(applyDescription(params));
      default:                         return jsonResponse({ ok: false, error: 'UNKNOWN_ACTION' });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doGet(e) {
  return jsonResponse({ ok: true, version: VERSION });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================== セッション検証 =====================
function validateSession(session) {
  if (!session || !session.token) return { ok: false };
  try {
    const res = UrlFetchApp.fetch(AUTH_GAS_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ action: 'validateSession', token: session.token }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    if (!data.ok) return { ok: false };
    return { ok: true, user: { user_id: data.user_id, name: data.name || data.user_id || '不明' } };
  } catch (e) {
    return { ok: false };
  }
}

// ===================== メインデータ読み込み =====================
function loadData() {
  const csvData = loadCsvFromDrive();
  if (!csvData.ok) return csvData;

  const bcartProducts = bcartGetAll('/products');
  if (!bcartProducts.ok) return bcartProducts;

  Utilities.sleep(500);

  const bcartSets = bcartGetAll('/product_sets');
  if (!bcartSets.ok) return bcartSets;

  const ignoreMap = getIgnoreMap();
  const wipMap = getWipMap();

  const diffs = calcDiffs(csvData.rows, bcartProducts.data, bcartSets.data, ignoreMap, wipMap);

  const csvHeaders = csvData.rows.length > 0 ? Object.keys(csvData.rows[0]) : [];
  const csvSample = csvData.rows.length > 0 ? {
    コード: csvData.rows[0]['コード'],
    売上単価: csvData.rows[0]['売上単価'],
    廃番: csvData.rows[0]['廃番']
  } : {};

  const supplierMap = {};
  csvData.rows.forEach(row => {
    if (row['コード'] && row['仕入先名']) {
      const key = String(parseInt(row['コード'], 10) || row['コード']);
      if (!supplierMap[key]) supplierMap[key] = row['仕入先名'];
    }
  });

  return {
    ok: true,
    diffs: diffs,
    supplierMap: supplierMap,
    csvUpdatedAt: csvData.updatedAt,
    isOld: csvData.isOld,
    totalCsv: csvData.rows.length,
    totalBcart: bcartSets.data.length,
    csvHeaders: csvHeaders,
    csvSample: csvSample
  };
}

// ===================== デバッグ =====================
function debugData() {
  const csvResult = loadCsvFromDrive();
  let csvDebug;
  if (csvResult.ok && csvResult.rows.length > 0) {
    const sample5 = csvResult.rows.slice(0, 5);
    csvDebug = {
      totalRows: csvResult.rows.length,
      headers: Object.keys(csvResult.rows[0]),
      sampleCodes: sample5.map(r => r['コード']),
      sampleCodesStripped: sample5.map(r => String(parseInt(r['コード'], 10) || r['コード'])),
      sample1Full: csvResult.rows[0]
    };
  } else {
    csvDebug = { error: csvResult.error };
  }

  const setsResult = bcartGetAll('/product_sets');
  let bcartDebug;
  if (setsResult.ok) {
    const data = setsResult.data;
    let matchCount = 0;
    if (csvResult.ok) {
      const bcartSetMap = {};
      data.forEach(s => { bcartSetMap[s.product_no] = true; });
      csvResult.rows.forEach(r => {
        const stripped = String(parseInt(r['コード'], 10) || r['コード']);
        const raw = r['コード'];
        if (bcartSetMap[stripped] || bcartSetMap[raw]) matchCount++;
      });
    }
    bcartDebug = {
      total: data.length,
      sampleProductNos: data.slice(0, 5).map(s => s.product_no),
      sample1Full: data[0],
      matchCountWithStrip: matchCount
    };
  } else {
    bcartDebug = { error: setsResult.error };
  }

  return { ok: true, csv: csvDebug, bcart: bcartDebug };
}

function debugCode(params) {
  const targetCode = String(params.code || '153');
  const csvResult = loadCsvFromDrive();
  if (!csvResult.ok) return csvResult;

  const matches = csvResult.rows.filter(r => {
    const key = String(parseInt(r['コード'], 10) || r['コード']);
    return key === targetCode || r['コード'] === targetCode;
  });

  const setsResult = bcartGetAll('/product_sets');
  const bcartMatch = setsResult.ok ? setsResult.data.find(s => s.product_no === targetCode) : null;

  return {
    ok: true,
    targetCode: targetCode,
    csv: {
      matchCount: matches.length,
      rows: matches.map(r => ({
        コード: r['コード'],
        商品名: r['商品名'],
        売上単価: r['売上単価'],
        仕入単価: r['仕入単価'],
        廃番: r['廃番'],
        仕入先名: r['仕入先名']
      }))
    },
    bcart: bcartMatch ? {
      id: bcartMatch.id,
      product_no: bcartMatch.product_no,
      name: bcartMatch.name,
      unit_price: bcartMatch.unit_price
    } : null
  };
}

// ===================== CSV読み込み =====================
function loadCsvFromDrive() {
  try {
    const folder = DriveApp.getFolderById(CSV_FOLDER_ID);
    const files = folder.getFilesByName(CSV_FILENAME);
    if (!files.hasNext()) return { ok: false, error: 'CSV_NOT_FOUND' };

    const file = files.next();
    const updatedAt = file.getLastUpdated().toLocaleString('ja-JP');
    const daysDiff = (new Date() - file.getLastUpdated()) / (1000 * 60 * 60 * 24);
    const isOld = daysDiff > 3;

    const content = file.getBlob().getDataAsString('Shift_JIS');
    const rows = parseCsv(content);

    return { ok: true, rows: rows, updatedAt: updatedAt, isOld: isOld };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function parseCsv(content) {
  const lines = content.split('\n').filter(l => l.trim());
  if (lines.length < 2) return [];

  const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
  const rows = [];

  for (let i = 1; i < lines.length; i++) {
    const values = splitCsvLine(lines[i]);
    if (values.length < headers.length) continue;
    const row = {};
    headers.forEach((h, idx) => row[h] = (values[idx] || '').trim().replace(/^"|"$/g, ''));
    rows.push(row);
  }
  return rows;
}

function splitCsvLine(line) {
  const result = [];
  let current = '';
  let inQuote = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { inQuote = !inQuote; }
    else if (ch === ',' && !inQuote) { result.push(current); current = ''; }
    else { current += ch; }
  }
  result.push(current);
  return result;
}

// ===================== 突合処理 =====================
function calcDiffs(csvRows, bcartProducts, bcartSets, ignoreMap, wipMap) {
  const bcartSetMap = {};
  bcartSets.forEach(s => { bcartSetMap[s.product_no] = s; });

  const bcartProductMap = {};
  bcartProducts.forEach(p => { bcartProductMap[p.id] = p; });

  const diffs = [];

  csvRows.forEach(row => {
    const code = row['コード'];
    if (!code) return;

    const codeKey = String(parseInt(code, 10) || code);

    const ignoreEntry = ignoreMap[codeKey];
    const isIgnored = !!ignoreEntry;
    const isWip = !!wipMap[codeKey];
    const bcartSet = bcartSetMap[codeKey];

    if (!bcartSet) {
      const isAlsoDiscontinued = row['廃番'] === '1' || row['廃番'] === 'TRUE' || row['廃番'] === '廃番';
      if (isAlsoDiscontinued) return;
      const unregPrice = parseFloat(String(row['売上単価'] || '').replace(/,/g, '')) || 0;
      if (!unregPrice) return;
      diffs.push({
        type: 'unregistered',
        code: codeKey,
        name: row['商品名'] || row['略称'] || '',
        supplier: row['仕入先名'] || '',
        csvPrice:  unregPrice,
        csvKouri:  parseFloat(String(row['定価１'] || row['定価1'] || '').replace(/,/g, '')) || 0,
        csvShiire: parseFloat(String(row['仕入単価'] || '').replace(/,/g, '')) || 0,
        csvJan:    (row['JANCD'] || '').trim(),
        csvUnit:   (row['単位名'] || '').trim(),
        stockManagement: row['在庫有無'] || '',
        isIgnored: isIgnored,
        isWip: isWip,
        ignoreReason: ignoreEntry ? (ignoreEntry.reason || '') : '',
        ignoreDate:   ignoreEntry ? (ignoreEntry.date   || '') : '',
        bcartSetId: null,
        bcartProductId: null
      });
      return;
    }

    const bcartProduct = bcartProductMap[bcartSet.product_id] || {};
    const csvPrice  = parseFloat(String(row['売上単価'] || '').replace(/,/g, '')) || 0;
    const csvKouri  = parseFloat(String(row['定価１']   || row['定価1'] || '').replace(/,/g, '')) || 0;
    const csvShiire = parseFloat(String(row['仕入単価'] || '').replace(/,/g, '')) || 0;
    const bcartPrice = parseFloat(bcartSet.unit_price) || 0;
    const bcartJodai = parseFloat(bcartSet.jodai) || 0;
    const bcartJan   = (bcartSet.jan_code || '').trim();
    const csvJan     = (row['JANCD'] || '').trim();
    const isDiscontinued = row['廃番'] === '1' || row['廃番'] === 'TRUE' || row['廃番'] === '廃番';
    const bcartSetVisible = bcartSet.set_flag !== '非表示';

    const issues = [];

    if (csvPrice > 0 && Math.abs(csvPrice - bcartPrice) > 0) {
      issues.push({ type: 'price', csvPrice: csvPrice, bcartPrice: bcartPrice });
    }
    if (isDiscontinued && bcartSetVisible) {
      issues.push({ type: 'discontinued', bcartProductFlag: bcartProduct.flag || '' });
    }
    // ② 廃番済み・セット非表示だが親商品がまだ表示中の場合を検出
    if (isDiscontinued && !bcartSetVisible && bcartProduct.flag === '表示') {
      issues.push({ type: 'parent_visible' });
    }
    if (csvKouri > 0 && Math.abs(csvKouri - bcartJodai) > 0) {
      issues.push({ type: 'jodai', csvJodai: csvKouri, bcartJodai: bcartJodai });
    }
    if (csvJan && csvJan !== bcartJan) {
      issues.push({ type: 'jan', csvJan: csvJan, bcartJan: bcartJan });
    }

    if (issues.length > 0) {
      diffs.push({
        type: 'diff',
        code: codeKey,
        name: row['商品名'] || row['略称'] || '',
        supplier: row['仕入先名'] || '',
        issues: issues,
        stockManagement: row['在庫有無'] || '',
        csvKouri: csvKouri,
        csvShiire: csvShiire,
        isIgnored: isIgnored,
        isWip: isWip,
        ignoreReason: ignoreEntry ? (ignoreEntry.reason || '') : '',
        ignoreDate:   ignoreEntry ? (ignoreEntry.date   || '') : '',
        bcartSetId: bcartSet.id,
        bcartProductId: bcartSet.product_id
      });
    }
  });

  return diffs;
}

// ===================== BCART API =====================
function getBcartToken() {
  const token = PropertiesService.getScriptProperties().getProperty('BCART_TOKEN');
  if (!token) throw new Error('BCARTトークンが設定されていません（スクリプトプロパティ: BCART_TOKEN）');
  return token;
}

function bcartGetAll(path) {
  try {
    const token = getBcartToken();
    const allData = [];
    const limit = 100;
    let offset = 0;

    while (true) {
      const url = BCART_BASE_URL + path + '?limit=' + limit + '&offset=' + offset;
      let res;
      for (let retry = 0; retry <= 3; retry++) {
        res = UrlFetchApp.fetch(url, {
          method: 'get',
          headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' },
          muteHttpExceptions: true
        });
        const code = res.getResponseCode();
        if (code === 429) {
          if (retry < 3) { Utilities.sleep(5000 * (retry + 1)); continue; }
          return { ok: false, error: 'BCART_API_ERROR: 429 レート制限（しばらく待ってから再読み込みしてください）' };
        }
        if (code === 503 || code === 502) {
          if (retry < 3) { Utilities.sleep(5000 * (retry + 1)); continue; }
          return { ok: false, error: 'BCART_API_ERROR: ' + code + ' 帯域幅エラー（しばらく待ってから再読み込みしてください）\n' + res.getContentText().slice(0, 300) };
        }
        break;
      }
      if (res.getResponseCode() !== 200) {
        return { ok: false, error: 'BCART_API_ERROR: ' + res.getResponseCode() + '\n' + res.getContentText().slice(0, 300) };
      }
      const parsed = JSON.parse(res.getContentText());
      if (parsed.message || parsed.error) {
        return { ok: false, error: 'BCART_API_ERROR: ' + (parsed.message || parsed.error) };
      }
      const page = parsed.data || parsed.product_sets || parsed.products || parsed.specials || parsed.product_stock || parsed.categories || parsed.product_features || parsed.features || parsed;
      if (!Array.isArray(page) || page.length === 0) break;
      allData.push(...page);
      if (page.length < limit) break;
      offset += limit;
      Utilities.sleep(600);
    }

    return { ok: true, data: allData };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function bcartGet(path, params) {
  try {
    const token = getBcartToken();
    let url = BCART_BASE_URL + path;
    if (params) {
      const qs = Object.entries(params).map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join('&');
      url += '?' + qs;
    }
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { ok: false, error: 'BCART_API_ERROR: ' + res.getResponseCode() };
    const data = JSON.parse(res.getContentText());
    return { ok: true, data: data.data || data.product_set || data };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function bcartPatch(path, body) {
  try {
    const token = getBcartToken();
    const res = UrlFetchApp.fetch(BCART_BASE_URL + path, {
      method: 'patch',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code !== 200 && code !== 204) return { ok: false, error: 'BCART_API_ERROR: ' + code + ' ' + res.getContentText() };
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function bcartDelete(path) {
  try {
    const token = getBcartToken();
    const res = UrlFetchApp.fetch(BCART_BASE_URL + path, {
      method: 'delete',
      headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code !== 200 && code !== 204) return { ok: false, error: 'BCART_API_ERROR: ' + code };
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function bcartPost(path, body) {
  try {
    const token = getBcartToken();
    const res = UrlFetchApp.fetch(BCART_BASE_URL + path, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code !== 200 && code !== 201) {
      return { ok: false, error: 'BCART_API_ERROR: ' + code + ' ' + res.getContentText() };
    }
    return { ok: true, data: JSON.parse(res.getContentText()) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ===================== 更新処理 =====================
function updatePrice(params) {
  const body = { unit_price: params.price };
  if (params.csvKouri || params.csvShiire) {
    body.group_price = {};
    if (params.csvKouri)  body.group_price['1']  = { fixed_price: params.csvKouri };
    if (params.csvShiire) body.group_price['10'] = { fixed_price: params.csvShiire };
  }
  const res = bcartPatch('/product_sets/' + params.bcartSetId, body);
  addHistory({
    userName: params._userName,
    code: params.code || '',
    name: params.name || '',
    type: '価格更新',
    before: params.beforePrice ? params.beforePrice + '円' : '',
    after: params.price + '円',
    result: res.ok ? '成功' : ('失敗: ' + res.error)
  });
  return res;
}

function updateJodai(params) {
  const body = { jodai: params.jodai };
  if (params.jodaiType) body.jodai_type = params.jodaiType;
  const res = bcartPatch('/product_sets/' + params.bcartSetId, body);
  addHistory({
    userName: params._userName,
    code: params.code || '',
    name: params.name || '',
    type: '上代更新',
    before: params.beforeJodai ? params.beforeJodai + '円' : '',
    after: params.jodai + '円',
    result: res.ok ? '成功' : ('失敗: ' + res.error)
  });
  return res;
}

function updateJan(params) {
  const res = bcartPatch('/product_sets/' + params.bcartSetId, { jan_code: params.janCode });
  addHistory({
    userName: params._userName,
    code: params.code || '',
    name: params.name || '',
    type: 'JAN更新',
    before: params.beforeJan || '',
    after: params.janCode || '',
    result: res.ok ? '成功' : ('失敗: ' + res.error)
  });
  return res;
}

function updateAll(params) {
  const body = {};
  if (params.price   !== undefined) body.unit_price = params.price;
  if (params.jodai   !== undefined) body.jodai = params.jodai;
  if (params.jodai   !== undefined && params.jodaiType) body.jodai_type = params.jodaiType;
  if (params.janCode !== undefined) body.jan_code = params.janCode;
  if (params.csvKouri || params.csvShiire) {
    body.group_price = {};
    if (params.csvKouri)  body.group_price['1']  = { fixed_price: params.csvKouri };
    if (params.csvShiire) body.group_price['10'] = { fixed_price: params.csvShiire };
  }
  const res = bcartPatch('/product_sets/' + params.bcartSetId, body);
  addHistory({
    userName: params._userName,
    code: params.code || '',
    name: params.name || '',
    type: '一括更新（価格/上代/JAN）',
    before: '',
    after: JSON.stringify(body),
    result: res.ok ? '成功' : ('失敗: ' + res.error)
  });
  return res;
}

function hideProductSet(params) {
  const res = bcartPatch('/product_sets/' + params.bcartSetId, { set_flag: '非表示' });
  addHistory({
    userName: params._userName,
    code: params.code || '',
    name: params.name || '',
    type: '廃番非表示（セット）',
    before: '表示',
    after: '非表示',
    result: res.ok ? '成功' : ('失敗: ' + res.error)
  });
  return res;
}

// ⑧ 品番存在確認付き欠品設定
function setStock(params) {
  // 品番がBCARTに存在するか事前確認
  const checkRes = bcartGet('/product_stock/' + params.productNo);
  if (!checkRes.ok) {
    return { ok: false, error: '品番「' + params.productNo + '」はBCARTに見つかりませんでした（' + checkRes.error + '）' };
  }
  const stockData = checkRes.data;
  const isEmpty = !stockData ||
    (Array.isArray(stockData) && stockData.length === 0) ||
    (typeof stockData === 'object' && !Array.isArray(stockData) && Object.keys(stockData).length === 0);
  if (isEmpty) {
    return { ok: false, error: '品番「' + params.productNo + '」は在庫管理が設定されていないか、BCARTに存在しません' };
  }

  const res = bcartPatch('/product_stock', [{ product_no: params.productNo, stock: String(params.stock) }]);
  addHistory({
    userName: params._userName,
    code: params.productNo || '',
    name: params.name || '',
    type: '欠品設定',
    before: '',
    after: params.stock === 0 ? '欠品（在庫0）' : '在庫あり（在庫1）',
    result: res.ok ? '成功' : ('失敗: ' + res.error)
  });
  return res;
}

function bulkUpdate(params) {
  const results = [];
  params.items.forEach(item => {
    let res;
    if (item.type === 'price') {
      const body = { unit_price: item.price };
      if (item.csvKouri || item.csvShiire) {
        body.group_price = {};
        if (item.csvKouri)  body.group_price['1']  = { fixed_price: item.csvKouri };
        if (item.csvShiire) body.group_price['10'] = { fixed_price: item.csvShiire };
      }
      res = bcartPatch('/product_sets/' + item.bcartSetId, body);
      addHistory({
        userName: params._userName,
        code: item.code || '',
        name: item.name || '',
        type: '価格更新（一括）',
        before: item.beforePrice ? item.beforePrice + '円' : '',
        after: item.price + '円',
        result: res.ok ? '成功' : ('失敗: ' + res.error)
      });
    } else if (item.type === 'discontinued') {
      res = bcartPatch('/product_sets/' + item.bcartSetId, { set_flag: '非表示' });
      addHistory({
        userName: params._userName,
        code: item.code || '',
        name: item.name || '',
        type: '廃番非表示（セット・一括）',
        before: '表示',
        after: '非表示',
        result: res.ok ? '成功' : ('失敗: ' + res.error)
      });
    } else if (item.type === 'jodai') {
      const jodaiBody = { jodai: item.jodai };
      if (item.jodaiType) jodaiBody.jodai_type = item.jodaiType;
      res = bcartPatch('/product_sets/' + item.bcartSetId, jodaiBody);
    } else if (item.type === 'jan') {
      res = bcartPatch('/product_sets/' + item.bcartSetId, { jan_code: item.janCode });
    }
    results.push({ code: item.code, ok: res ? res.ok : false, error: res ? res.error : '' });
  });
  return { ok: true, results: results };
}

function bulkIgnore(params) {
  const results = [];
  (params.items || []).forEach(item => {
    const res = markIgnore({ code: item.code, name: item.name, reason: params.reason || '', supplier: item.supplier || '' });
    results.push({ code: item.code, ok: res.ok });
  });
  return { ok: true, results: results };
}

function bulkMarkWip(params) {
  const results = [];
  (params.items || []).forEach(item => {
    const res = markWip({ code: item.code, name: item.name });
    results.push({ code: item.code, ok: res.ok });
  });
  return { ok: true, results: results };
}

// ===================== 商品検索 =====================
function searchProducts(params) {
  const products = bcartGetAll('/products');
  if (!products.ok) return products;

  const sets = bcartGetAll('/product_sets');
  if (!sets.ok) return sets;

  const setByProductId = {};
  sets.data.forEach(s => {
    if (!setByProductId[s.product_id]) setByProductId[s.product_id] = s;
  });

  const now = new Date();
  let filtered = products.data;

  if (params.keyword) {
    const kw = params.keyword.toLowerCase();
    filtered = filtered.filter(p =>
      (p.name || '').toLowerCase().includes(kw) ||
      (p.product_no || '').toLowerCase().includes(kw)
    );
  }

  const getDisplayEnd = p => p.hanbai_end || '';
  if (params.expired) {
    filtered = filtered.filter(p => {
      const de = getDisplayEnd(p);
      if (!de) return false;
      try { return new Date(de) < now; } catch(e) { return false; }
    });
  }

  if (params.status === '公開') {
    filtered = filtered.filter(p => p.flag === '表示');
  } else if (params.status === '非表示') {
    filtered = filtered.filter(p => p.flag === '非表示');
  }

  if (params.stockZero) {
    filtered = filtered.filter(p => {
      const s = setByProductId[p.id];
      if (!s) return false;
      const stock = s.stock !== undefined ? s.stock : s.inventory;
      return stock !== undefined && (parseInt(stock) === 0 || String(stock) === '0');
    });
  }

  if (params.specialId) {
    const fid = String(params.specialId);
    filtered = filtered.filter(p =>
      String(p.feature_id1 || '') === fid ||
      String(p.feature_id2 || '') === fid ||
      String(p.feature_id3 || '') === fid
    );
  }

  const result = filtered.map(p => {
    const s = setByProductId[p.id] || {};
    return {
      id: p.id,
      product_no: p.product_no || p.main_no || '',
      name: p.name || '',
      flag: p.flag || '',
      display_start: p.hanbai_start || '',
      display_end: getDisplayEnd(p),
      unit_price: s.unit_price || '',
      stock: s.stock !== undefined ? s.stock : (s.inventory !== undefined ? s.inventory : ''),
      set_id: s.id || ''
    };
  });

  return { ok: true, products: result, total: result.length };
}

function getSpecials() {
  const endpoints = ['/product_features', '/features'];
  for (const ep of endpoints) {
    try {
      const res = bcartGetAll(ep);
      if (res.ok && res.data) {
        const raw = res.data;
        const list = raw.product_features || raw.features || raw.data || (Array.isArray(raw) ? raw : []);
        if (list.length > 0) {
          const specials = list.map(f => ({
            id:   f.id         || f.feature_id   || f.featureId,
            name: f.name       || f.feature_name || f.title || f.featureName || String(f.id || f.feature_id || '')
          })).filter(f => f.id);
          if (specials.length > 0) return { ok: true, specials: specials };
        }
      }
    } catch(e) {}
  }

  try {
    const products = bcartGetAll('/products');
    if (!products.ok) return { ok: true, specials: [] };
    const featureIds = new Set();
    products.data.forEach(p => {
      if (p.feature_id1) featureIds.add(p.feature_id1);
      if (p.feature_id2) featureIds.add(p.feature_id2);
      if (p.feature_id3) featureIds.add(p.feature_id3);
    });
    const featureList = [...featureIds].sort((a, b) => a - b).map(id => {
      const matched = products.data.filter(p =>
        p.feature_id1 == id || p.feature_id2 == id || p.feature_id3 == id
      );
      const sample = matched.length > 0 ? String(matched[0].name || '').substring(0, 12) : '';
      return { id: id, name: '特集' + id + ': ' + sample + ' 他' + matched.length + '件' };
    });
    return { ok: true, specials: featureList };
  } catch(e) {
    return { ok: true, specials: [] };
  }
}

function debugProduct() {
  const products = bcartGetAll('/products');
  if (!products.ok) return products;
  const sample = products.data.length > 0 ? products.data[0] : null;
  return {
    ok: true,
    total: products.data.length,
    fields: sample ? Object.keys(sample) : [],
    sample: sample
  };
}

function updateProductAction(params) {
  const results = [];
  if (params.productId && params.productFields) {
    const res = bcartPatch('/products/' + params.productId, params.productFields);
    results.push({ target: 'product', ok: res.ok, error: res.error || '' });
    if (res.ok) {
      addHistory({
        userName: params._userName || '不明',
        code: params.code || '',
        name: params.name || '',
        type: '親商品非表示',
        before: '表示',
        after: '非表示',
        result: '成功'
      });
    }
  }
  if (params.setId && params.setFields) {
    const res = bcartPatch('/product_sets/' + params.setId, params.setFields);
    results.push({ target: 'set', ok: res.ok, error: res.error || '' });
  }
  const allOk = results.length > 0 && results.every(r => r.ok);
  return { ok: allOk, results: results };
}

function deleteProduct(params) {
  return bcartDelete('/products/' + params.productId);
}

// ===================== 機能A: 新規登録 =====================
function getCategories() {
  const result = bcartGetAll('/categories');
  if (!result.ok) return result;
  return {
    ok: true,
    categories: result.data.map(c => ({ id: c.id, name: c.name || String(c.id) }))
  };
}

// ① 孤立商品の自動ロールバック付き新規登録
function registerProduct(params) {
  const productBody = {
    products: [{
      name:        params.productName,
      category_id: params.categoryId,
      flag:        params.productFlag || '非表示',
      feature_id1: params.featureId1 || null,
      feature_id2: params.featureId2 || null,
      feature_id3: params.featureId3 || null
    }]
  };

  const step1 = bcartPost('/products', productBody);
  if (!step1.ok) {
    addHistory({
      userName: params._userName,
      code: params.code || '',
      name: params.productName || '',
      type: '新規登録（商品作成失敗）',
      before: '', after: '',
      result: '失敗: ' + step1.error
    });
    return step1;
  }

  const createdProductId = step1.data && step1.data.products && step1.data.products[0]
    ? step1.data.products[0].id : null;
  if (!createdProductId) {
    return { ok: false, error: '商品IDが取得できませんでした（レスポンス: ' + JSON.stringify(step1.data) + '）' };
  }

  const setBody = {
    product_sets: [{
      product_id:  createdProductId,
      product_no:  params.code,
      name:        params.setName,
      jan_code:    params.janCode || '',
      unit_price:  params.csvPrice,
      jodai:       params.csvKouri || 0,
      group_price: {
        '1':  { fixed_price: params.csvKouri || 0 },
        '10': { fixed_price: params.csvShiire || 0 }
      },
      unit:        params.csvUnit || '',
      quantity:    1,
      min_order:   1,
      stock_flag:  1,
      tax_type_id: params.taxTypeId || 1,
      set_flag:    params.setFlag || '非表示'
    }]
  };

  const step2 = bcartPost('/product_sets', setBody);
  if (!step2.ok) {
    // 自動ロールバック: step1で作成した商品を削除する
    const rollback = bcartDelete('/products/' + createdProductId);
    if (rollback.ok) {
      addHistory({
        userName: params._userName,
        code: params.code || '',
        name: params.productName || '',
        type: '新規登録（ロールバック完了）',
        before: '', after: '商品ID: ' + createdProductId + ' を削除',
        result: '失敗: ' + step2.error
      });
      return { ok: false, error: '登録失敗（自動ロールバック完了）: ' + step2.error };
    } else {
      // ロールバックも失敗した場合は孤立商品IDを返す
      addHistory({
        userName: params._userName,
        code: params.code || '',
        name: params.productName || '',
        type: '新規登録（孤立商品: 手動削除要）',
        before: '', after: '孤立商品ID: ' + createdProductId,
        result: '失敗: ' + step2.error
      });
      return { ok: false, error: step2.error, orphanProductId: createdProductId };
    }
  }

  const createdSetId = step2.data && step2.data.product_sets && step2.data.product_sets[0]
    ? step2.data.product_sets[0].id : null;

  unmarkWip({ code: params.code });

  addHistory({
    userName: params._userName,
    code: params.code || '',
    name: params.productName || '',
    type: '新規登録',
    before: '',
    after: '商品ID: ' + createdProductId + ' / セットID: ' + createdSetId,
    result: '成功'
  });

  return { ok: true, productId: createdProductId, setId: createdSetId };
}

// ===================== 会員取得 =====================
function getMembers() {
  try {
    const token = getBcartToken();
    const allMembers = [];
    let offset = 0;
    const limit = 100;

    while (true) {
      const url = BCART_BASE_URL + '/customers?limit=' + limit + '&offset=' + offset;
      const res = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' },
        muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) {
        return { ok: false, error: 'BCART_API_ERROR: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200) };
      }
      const parsed = JSON.parse(res.getContentText());
      const page = parsed.customers || parsed.data || (Array.isArray(parsed) ? parsed : null);
      if (!page || !Array.isArray(page) || page.length === 0) break;
      allMembers.push(...page);
      if (page.length < limit) break;
      offset += limit;
      Utilities.sleep(300);
    }

    const members = allMembers.map(m => ({
      id:            String(m.id || ''),
      name:          m.comp_name || m.company_name || m.name || String(m.id || ''),
      ext_id:        String(m.ext_id || ''),
      comp_name:     String(m.comp_name || ''),
      view_group_id: String(m.view_group_id || ''),
      memo:          String(m.memo || ''),
      email:         String(m.email || ''),
      code:          String(m.code || m.customer_no || '')
    })).filter(m => m.id);

    return { ok: true, members: members };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ===================== 商品セット検索 =====================
function searchProductSets(params) {
  const keyword = (params.keyword || '').toLowerCase().trim();
  if (!keyword) return { ok: false, error: 'キーワードを入力してください' };

  const sets = bcartGetAll('/product_sets');
  if (!sets.ok) return sets;

  const filtered = sets.data.filter(s =>
    (String(s.product_no || '')).toLowerCase().includes(keyword) ||
    (String(s.name || '')).toLowerCase().includes(keyword)
  ).slice(0, 30);

  return {
    ok: true,
    sets: filtered.map(s => ({
      id:         s.id,
      product_no: s.product_no || '',
      name:       s.name || '',
      unit_price: s.unit_price || 0
    }))
  };
}

// ===================== 機能B: 特別価格管理 =====================

// ③ applied_at列を含めてデータを返す
function getSpecialPriceData() {
  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  const groups = [];
  for (let i = 1; i < groupRows.length; i++) {
    if (groupRows[i][0]) {
      groups.push({
        group_id:       String(groupRows[i][0]),
        group_name:     String(groupRows[i][1]),
        member_ids:     String(groupRows[i][2] || ''),
        created_at:     String(groupRows[i][3]),
        note:           groupRows[i][4] || '',
        use_view_filter: groupRows[i][5] === true || groupRows[i][5] === 'TRUE'
      });
    }
  }

  const detailSheet = getOrCreateSheet(SHEET_SP_DETAILS);
  const detailRows = detailSheet.getDataRange().getValues();
  const details = [];
  for (let i = 1; i < detailRows.length; i++) {
    if (detailRows[i][0]) {
      details.push({
        detail_id:        String(detailRows[i][0]),
        group_id:         String(detailRows[i][1]),
        product_set_id:   detailRows[i][2],
        product_no:       String(detailRows[i][3]),
        product_set_name: String(detailRows[i][4]),
        unit_price:       detailRows[i][5],
        updated_at:       String(detailRows[i][6]),
        applied_at:       String(detailRows[i][7] || '')  // ③ 追加
      });
    }
  }

  const vfSheet = getOrCreateSheet(SHEET_VF_DETAILS);
  const vfRows = vfSheet.getDataRange().getValues();
  const vfDetails = [];
  for (let i = 1; i < vfRows.length; i++) {
    if (vfRows[i][0]) {
      vfDetails.push({
        detail_id:        String(vfRows[i][0]),
        group_id:         String(vfRows[i][1]),
        product_set_id:   vfRows[i][2],
        product_no:       String(vfRows[i][3]),
        product_set_name: String(vfRows[i][4]),
        applied_at:       String(vfRows[i][5] || '')  // ③ VFは既存列だが意味を「反映時刻」に統一
      });
    }
  }

  return { ok: true, groups: groups, details: details, vfDetails: vfDetails };
}

function saveCustomerGroup(params) {
  const sheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const rows = sheet.getDataRange().getValues();
  const memberIds = params.member_ids || '';

  const uvf = params.use_view_filter === true || params.use_view_filter === 'TRUE' ? 'TRUE' : 'FALSE';
  if (params.group_id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === params.group_id) {
        sheet.getRange(i + 1, 2).setValue(params.group_name);
        const mCell = sheet.getRange(i + 1, 3);
        mCell.setNumberFormat('@');
        mCell.setValue(memberIds);
        sheet.getRange(i + 1, 5).setValue(params.note || '');
        sheet.getRange(i + 1, 6).setValue(uvf);
        return { ok: true, group_id: params.group_id };
      }
    }
  }

  const newId = 'G' + new Date().getTime().toString(36).toUpperCase();
  sheet.appendRow([newId, params.group_name, '', new Date().toLocaleString('ja-JP'), params.note || '', uvf]);
  const lastRow = sheet.getLastRow();
  const mCell = sheet.getRange(lastRow, 3);
  mCell.setNumberFormat('@');
  mCell.setValue(memberIds);
  return { ok: true, group_id: newId };
}

function deleteCustomerGroup(params) {
  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  let memberIds = [];
  let groupRowIdx = -1;
  for (let i = 1; i < groupRows.length; i++) {
    if (String(groupRows[i][0]) === params.group_id) {
      memberIds = String(groupRows[i][2] || '').split(',').map(s => s.trim()).filter(s => s);
      groupRowIdx = i + 1;
      break;
    }
  }

  const detailSheet = getOrCreateSheet(SHEET_SP_DETAILS);
  const detailRows = detailSheet.getDataRange().getValues();
  const detailRowsToDelete = [];
  const productSetIds = [];
  for (let i = 1; i < detailRows.length; i++) {
    if (String(detailRows[i][1]) === params.group_id) {
      detailRowsToDelete.push(i + 1);
      productSetIds.push(detailRows[i][2]);
    }
  }

  if (memberIds.length > 0 && productSetIds.length > 0) {
    const allSets = bcartGetAll('/product_sets');
    if (allSets.ok) {
      productSetIds.forEach(setId => {
        const bcartSet = allSets.data.find(s => s.id == setId);
        if (!bcartSet) return;
        const newSp = Object.assign({}, bcartSet.special_price || {});
        memberIds.forEach(mid => { delete newSp[String(mid)]; });
        bcartPatch('/product_sets/' + setId, { special_price: newSp });
        Utilities.sleep(100);
      });
    }
  }

  // 例外表示設定のクリーンアップ
  const vfSheet = getOrCreateSheet(SHEET_VF_DETAILS);
  const vfRows = vfSheet.getDataRange().getValues();
  const vfRowsToDelete = [];
  for (let i = 1; i < vfRows.length; i++) {
    if (String(vfRows[i][1]) === params.group_id) {
      vfRowsToDelete.push({ rowIdx: i + 1, setId: vfRows[i][2] });
    }
  }
  if (memberIds.length > 0 && vfRowsToDelete.length > 0) {
    vfRowsToDelete.forEach(entry => {
      const res = bcartGet('/product_sets/' + entry.setId);
      if (res.ok && res.data) {
        let ids = String(res.data.visible_customer_id || '').split(',').map(s => s.trim()).filter(s => s);
        ids = ids.filter(id => !memberIds.includes(id));
        const patch = { visible_customer_id: ids.join(',') };
        if (ids.length === 0) patch.view_group_filter = '';
        bcartPatch('/product_sets/' + entry.setId, patch);
        Utilities.sleep(100);
      }
    });
  }
  vfRowsToDelete.sort((a, b) => b.rowIdx - a.rowIdx).forEach(entry => vfSheet.deleteRow(entry.rowIdx));

  detailRowsToDelete.sort((a, b) => b - a).forEach(rowIdx => detailSheet.deleteRow(rowIdx));
  if (groupRowIdx > 0) groupSheet.deleteRow(groupRowIdx);

  addHistory({
    userName: params._userName,
    code: '', name: params.group_id,
    type: '顧客グループ削除', before: '', after: '', result: '成功'
  });
  return { ok: true };
}

function getProductSetsForFeature(params) {
  const featureId = String(params.featureId);
  const products = bcartGetAll('/products');
  if (!products.ok) return products;
  const allSets = bcartGetAll('/product_sets');
  if (!allSets.ok) return allSets;

  const setByProductId = {};
  allSets.data.forEach(s => {
    if (!setByProductId[s.product_id]) setByProductId[s.product_id] = [];
    setByProductId[s.product_id].push(s);
  });

  const matchProducts = products.data.filter(p =>
    String(p.feature_id1 || '') === featureId ||
    String(p.feature_id2 || '') === featureId ||
    String(p.feature_id3 || '') === featureId
  );

  const sets = [];
  matchProducts.forEach(p => {
    (setByProductId[p.id] || []).forEach(s => {
      sets.push({
        id:            s.id,
        product_no:    s.product_no || '',
        name:          s.name || p.name || '',
        unit_price:    s.unit_price || 0,
        special_price: s.special_price || {}
      });
    });
  });

  sets.sort((a, b) => String(a.product_no).localeCompare(String(b.product_no)));
  return { ok: true, sets: sets };
}

// ③ アプリ保存時にapplied_atをクリア（未反映状態にする）
function saveSpecialPriceDetails(params) {
  const detailSheet = getOrCreateSheet(SHEET_SP_DETAILS);
  let detailRows = detailSheet.getDataRange().getValues();

  params.items.forEach(item => {
    let found = false;
    for (let i = 1; i < detailRows.length; i++) {
      if (String(detailRows[i][1]) === params.group_id && String(detailRows[i][2]) === String(item.product_set_id)) {
        detailSheet.getRange(i + 1, 5).setValue(item.product_set_name || '');
        detailSheet.getRange(i + 1, 6).setValue(item.unit_price);
        detailSheet.getRange(i + 1, 7).setValue(new Date().toLocaleString('ja-JP'));
        detailSheet.getRange(i + 1, 8).setValue('');  // ③ applied_atをクリア（未反映）
        detailRows[i][5] = item.unit_price;
        found = true;
        break;
      }
    }
    if (!found) {
      const newId = 'D' + new Date().getTime().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 5);
      const newRow = [newId, params.group_id, item.product_set_id, item.product_no || '', item.product_set_name || '', item.unit_price, new Date().toLocaleString('ja-JP'), ''];  // ③ applied_at空で追加
      detailSheet.appendRow(newRow);
      detailRows.push(newRow);
    }
    Utilities.sleep(50);
  });

  addHistory({
    userName: params._userName,
    code: '', name: params.group_id,
    type: '特別価格アプリ保存',
    before: '', after: params.items.length + '件',
    result: '成功'
  });
  return { ok: true, savedCount: params.items.length };
}

// ③ BCART反映時にapplied_atを記録
function applyGroupPrices(params) {
  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  let memberIds = [];
  for (let i = 1; i < groupRows.length; i++) {
    if (String(groupRows[i][0]) === params.group_id) {
      memberIds = String(groupRows[i][2] || '').split(',').map(s => s.trim()).filter(s => s);
      break;
    }
  }
  if (memberIds.length === 0) return { ok: false, error: 'グループの会員IDが設定されていません' };

  const detailSheet = getOrCreateSheet(SHEET_SP_DETAILS);
  let detailRows = detailSheet.getDataRange().getValues();

  const allSets = bcartGetAll('/product_sets');
  if (!allSets.ok) return allSets;
  const setMap = {};
  allSets.data.forEach(s => { setMap[s.id] = s; });

  let successCount = 0, failCount = 0;
  const errors = [];

  params.items.forEach(item => {
    const bcartSet = setMap[item.product_set_id];
    if (!bcartSet) {
      failCount++;
      errors.push('setId ' + item.product_set_id + ': 商品セットが見つかりません');
      return;
    }

    const newSp = Object.assign({}, bcartSet.special_price || {});
    Object.keys(newSp).forEach(key => {
      if (key.indexOf(',') !== -1) {
        const ids = key.split(',').map(s => s.trim()).filter(s => s);
        if (ids.some(id => memberIds.includes(id))) {
          delete newSp[key];
        }
      }
    });
    memberIds.forEach(mid => { newSp[String(mid)] = { unit_price: item.unit_price }; });

    const res = bcartPatch('/product_sets/' + item.product_set_id, { special_price: newSp });
    if (res.ok) {
      successCount++;
      const nowStr = new Date().toLocaleString('ja-JP');
      let found = false;
      for (let i = 1; i < detailRows.length; i++) {
        if (String(detailRows[i][1]) === params.group_id && detailRows[i][2] == item.product_set_id) {
          detailSheet.getRange(i + 1, 5).setValue(item.product_set_name || '');
          detailSheet.getRange(i + 1, 6).setValue(item.unit_price);
          detailSheet.getRange(i + 1, 7).setValue(nowStr);
          detailSheet.getRange(i + 1, 8).setValue(nowStr);  // ③ applied_atに反映時刻を記録
          detailRows[i][5] = item.unit_price;
          found = true;
          break;
        }
      }
      if (!found) {
        const newId = 'D' + new Date().getTime().toString(36).toUpperCase();
        const newRow = [newId, params.group_id, item.product_set_id, item.product_no || '', item.product_set_name || '', item.unit_price, nowStr, nowStr];
        detailSheet.appendRow(newRow);
        detailRows.push(newRow);
      }
    } else {
      failCount++;
      errors.push('setId ' + item.product_set_id + ': ' + res.error);
    }
    Utilities.sleep(150);
  });

  addHistory({
    userName: params._userName,
    code: '', name: params.group_id,
    type: '特別価格適用',
    before: '',
    after: '成功: ' + successCount + '件 / 失敗: ' + failCount + '件',
    result: failCount > 0 ? '一部失敗' : '成功'
  });

  return { ok: true, successCount: successCount, failCount: failCount, errors: errors };
}

function deleteSpecialPriceDetail(params) {
  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  let memberIds = [];
  for (let i = 1; i < groupRows.length; i++) {
    if (String(groupRows[i][0]) === params.group_id) {
      memberIds = String(groupRows[i][2] || '').split(',').map(s => s.trim()).filter(s => s);
      break;
    }
  }

  if (memberIds.length > 0 && params.product_set_id) {
    const allSets = bcartGetAll('/product_sets');
    if (allSets.ok) {
      const bcartSet = allSets.data.find(s => s.id == params.product_set_id);
      if (bcartSet) {
        const newSp = Object.assign({}, bcartSet.special_price || {});
        memberIds.forEach(mid => { delete newSp[String(mid)]; });
        bcartPatch('/product_sets/' + params.product_set_id, { special_price: newSp });
      }
    }
  }

  const detailSheet = getOrCreateSheet(SHEET_SP_DETAILS);
  const detailRows = detailSheet.getDataRange().getValues();
  for (let i = 1; i < detailRows.length; i++) {
    if (String(detailRows[i][0]) === params.detail_id) {
      detailSheet.deleteRow(i + 1);
      break;
    }
  }

  addHistory({
    userName: params._userName,
    code: '', name: 'group:' + params.group_id + ' / set:' + params.product_set_id,
    type: '特別価格削除', before: '', after: '', result: '成功'
  });
  return { ok: true };
}

// ===================== 例外表示設定 =====================
const VF_FILTER_VALUE = '非会員,通常会員,1,2';

function getSpecialPriceCurrent(params) {
  const setIds = (params.product_set_ids || []).map(String);

  const allSets = bcartGetAll('/product_sets');
  if (!allSets.ok) return allSets;

  const results = setIds.map(setId => {
    const setData = allSets.data.find(s => String(s.id) === setId);
    if (!setData) return { product_set_id: setId, error: '商品セットが見つかりません' };
    return {
      product_set_id: setId,
      unit_price: setData.unit_price,
      all_special: setData.special_price || {}
    };
  });

  return { ok: true, results: results };
}

function getViewFilterCurrent(params) {
  const memberIds = params.member_ids || [];
  const results = [];
  (params.product_set_ids || []).forEach(setId => {
    const res = bcartGet('/product_sets/' + setId);
    if (res.ok && res.data) {
      const cur = res.data;
      const curVcid  = String(cur.visible_customer_id || '');
      const curFilter = String(cur.view_group_filter || '');
      const existingIds = curVcid.split(',').map(s => s.trim()).filter(s => s);
      const mergedIds   = [...new Set([...existingIds, ...memberIds])];
      results.push({
        product_set_id:            setId,
        before_view_group_filter:  curFilter,
        before_visible_customer_id: curVcid,
        after_view_group_filter:   VF_FILTER_VALUE,
        after_visible_customer_id: mergedIds.join(','),
        already_set:               memberIds.every(id => existingIds.includes(id)) && curFilter === VF_FILTER_VALUE
      });
    } else {
      results.push({ product_set_id: setId, error: res.error || '取得失敗' });
    }
    Utilities.sleep(100);
  });
  return { ok: true, results: results };
}

// ③ アプリ保存時にapplied_atをクリア（VF）
function saveViewFilterDetails(params) {
  const sheet = getOrCreateSheet(SHEET_VF_DETAILS);
  let rows = sheet.getDataRange().getValues();

  params.items.forEach(item => {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === params.group_id && String(rows[i][2]) === String(item.product_set_id)) {
        sheet.getRange(i + 1, 5).setValue(item.product_set_name || '');
        sheet.getRange(i + 1, 6).setValue('');  // ③ applied_atをクリア（未反映）
        rows[i][4] = item.product_set_name;
        found = true;
        break;
      }
    }
    if (!found) {
      const newId = 'V' + new Date().getTime().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 4);
      const newRow = [newId, params.group_id, item.product_set_id, item.product_no || '', item.product_set_name || '', ''];  // ③ applied_at空で追加
      sheet.appendRow(newRow);
      rows.push(newRow);
    }
    Utilities.sleep(50);
  });

  addHistory({
    userName: params._userName, code: '', name: params.group_id,
    type: '例外表示アプリ保存', before: '', after: params.items.length + '件', result: '成功'
  });
  return { ok: true, savedCount: params.items.length };
}

// ③ BCART反映時にapplied_atを記録（VF）
function applyViewFilters(params) {
  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  let memberIds = [];
  for (let i = 1; i < groupRows.length; i++) {
    if (String(groupRows[i][0]) === params.group_id) {
      memberIds = String(groupRows[i][2] || '').split(',').map(s => s.trim()).filter(s => s);
      break;
    }
  }
  if (memberIds.length === 0) return { ok: false, error: 'グループの会員IDが設定されていません' };

  const vfSheet = getOrCreateSheet(SHEET_VF_DETAILS);
  let vfRows = vfSheet.getDataRange().getValues();

  let successCount = 0, failCount = 0;
  const errors = [];

  params.items.forEach(item => {
    const cur = bcartGet('/product_sets/' + item.product_set_id);
    let existingIds = [];
    if (cur.ok && cur.data) {
      existingIds = String(cur.data.visible_customer_id || '').split(',').map(s => s.trim()).filter(s => s);
    }
    const merged = [...new Set([...existingIds, ...memberIds])].join(',');

    const res = bcartPatch('/product_sets/' + item.product_set_id, {
      view_group_filter:   VF_FILTER_VALUE,
      visible_customer_id: merged
    });
    if (res.ok) {
      successCount++;
      // ③ applied_atに反映時刻を記録
      const nowStr = new Date().toLocaleString('ja-JP');
      for (let i = 1; i < vfRows.length; i++) {
        if (String(vfRows[i][1]) === params.group_id && String(vfRows[i][2]) === String(item.product_set_id)) {
          vfSheet.getRange(i + 1, 6).setValue(nowStr);
          vfRows[i][5] = nowStr;
          break;
        }
      }
    } else {
      failCount++;
      errors.push('setId ' + item.product_set_id + ': ' + res.error);
    }
    Utilities.sleep(150);
  });

  addHistory({
    userName: params._userName, code: '', name: params.group_id,
    type: '例外表示適用',
    before: '', after: '成功: ' + successCount + '件 / 失敗: ' + failCount + '件',
    result: failCount > 0 ? '一部失敗' : '成功'
  });
  return { ok: true, successCount: successCount, failCount: failCount, errors: errors };
}

function deleteViewFilterDetail(params) {
  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  let memberIds = [];
  for (let i = 1; i < groupRows.length; i++) {
    if (String(groupRows[i][0]) === params.group_id) {
      memberIds = String(groupRows[i][2] || '').split(',').map(s => s.trim()).filter(s => s);
      break;
    }
  }

  if (params.product_set_id && memberIds.length > 0) {
    const cur = bcartGet('/product_sets/' + params.product_set_id);
    if (cur.ok && cur.data) {
      let ids = String(cur.data.visible_customer_id || '').split(',').map(s => s.trim()).filter(s => s);
      ids = ids.filter(id => !memberIds.includes(id));
      const patch = { visible_customer_id: ids.join(',') };
      if (ids.length === 0) patch.view_group_filter = '';
      bcartPatch('/product_sets/' + params.product_set_id, patch);
    }
  }

  const sheet = getOrCreateSheet(SHEET_VF_DETAILS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === params.detail_id) { sheet.deleteRow(i + 1); break; }
  }

  addHistory({
    userName: params._userName, code: '', name: 'group:' + params.group_id + ' / set:' + params.product_set_id,
    type: '例外表示削除', before: '', after: '', result: '成功'
  });
  return { ok: true };
}

// ===================== 更新履歴 =====================
function addHistory(entry) {
  try {
    const sheet = getOrCreateSheet(SHEET_HISTORY);
    sheet.appendRow([
      new Date().toLocaleString('ja-JP'),
      entry.userName || '不明',
      entry.code || '',
      entry.name || '',
      entry.type || '',
      entry.before || '',
      entry.after || '',
      entry.result || '成功'
    ]);
  } catch(e) {
    Logger.log('履歴書き込みエラー: ' + e);
  }
}

function getHistory() {
  const sheet = getOrCreateSheet(SHEET_HISTORY);
  const rows = sheet.getDataRange().getValues();
  const list = [];
  for (let i = rows.length - 1; i >= 1 && list.length < 200; i--) {
    if (rows[i][0]) {
      list.push({
        date:     String(rows[i][0]),
        userName: rows[i][1],
        code:     rows[i][2],
        name:     rows[i][3],
        type:     rows[i][4],
        before:   rows[i][5],
        after:    rows[i][6],
        result:   rows[i][7]
      });
    }
  }
  return { ok: true, list: list };
}

// ===================== シート管理 =====================
function getOrCreateSheet(sheetName) {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('MASTER_TOOL_SS_ID');
  let ss;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ssId = null; }
  }
  if (!ssId) {
    ss = SpreadsheetApp.create('BCARTマスター管理ツール_データ');
    props.setProperty('MASTER_TOOL_SS_ID', ss.getId());
  }

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === SHEET_IGNORE) {
      sheet.appendRow(['商品コード', '商品名', '理由', '登録日時', '仕入先名']);
    } else if (sheetName === SHEET_WIP) {
      sheet.appendRow(['商品コード', '商品名', '登録日時']);
    } else if (sheetName === SHEET_HISTORY) {
      sheet.appendRow(['日時', '操作者', '商品コード', '商品名', '操作種別', '変更前', '変更後', '結果']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_SP_GROUPS) {
      sheet.appendRow(['group_id', 'group_name', 'member_ids', 'created_at', 'note', 'use_view_filter']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_VF_DETAILS) {
      sheet.appendRow(['detail_id', 'group_id', 'product_set_id', 'product_no', 'product_set_name', 'applied_at']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_SP_DETAILS) {
      // ③ applied_at列を含む8列構成
      sheet.appendRow(['detail_id', 'group_id', 'product_set_id', 'product_no', 'product_set_name', 'unit_price', 'updated_at', 'applied_at']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f3f4f6');
    }
  } else {
    // 既存シートのマイグレーション処理
    if (sheetName === SHEET_IGNORE) {
      const lastCol = sheet.getLastColumn();
      if (lastCol < 5) {
        sheet.getRange(1, 5).setValue('仕入先名');
      } else {
        const headerVal = sheet.getRange(1, 5).getValue();
        if (!headerVal) sheet.getRange(1, 5).setValue('仕入先名');
      }
    } else if (sheetName === SHEET_SP_GROUPS) {
      // use_view_filter 列の自動追加
      if (sheet.getLastColumn() < 6) {
        sheet.getRange(1, 6).setValue('use_view_filter');
        sheet.getRange(1, 6).setFontWeight('bold').setBackground('#f3f4f6');
      }
    } else if (sheetName === SHEET_SP_DETAILS) {
      // ③ applied_at 列の自動追加（既存シートへのマイグレーション）
      if (sheet.getLastColumn() < 8) {
        sheet.getRange(1, 8).setValue('applied_at');
        sheet.getRange(1, 8).setFontWeight('bold').setBackground('#f3f4f6');
      }
    }
  }
  return sheet;
}

function getIgnoreMap() {
  const sheet = getOrCreateSheet(SHEET_IGNORE);
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      const rawDate = String(rows[i][3] || '');
      map[rows[i][0]] = {
        reason: rows[i][2] || '',
        date:   rawDate.split(' ')[0] || rawDate
      };
    }
  }
  return map;
}

function getWipMap() {
  const sheet = getOrCreateSheet(SHEET_WIP);
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) map[rows[i][0]] = true;
  }
  return map;
}

function markIgnore(params) {
  const sheet = getOrCreateSheet(SHEET_IGNORE);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === params.code) {
      sheet.getRange(i + 1, 3).setValue(params.reason);
      if (params.supplier) sheet.getRange(i + 1, 5).setValue(params.supplier);
      return { ok: true };
    }
  }
  sheet.appendRow([params.code, params.name, params.reason, new Date().toLocaleString('ja-JP'), params.supplier || '']);
  return { ok: true };
}

function unmarkIgnore(params) {
  const sheet = getOrCreateSheet(SHEET_IGNORE);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === params.code) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: true };
}

function getIgnoreList() {
  const sheet = getOrCreateSheet(SHEET_IGNORE);
  const lastCol = sheet.getLastColumn();
  const rows = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      list.push({
        code:         rows[i][0],
        name:         rows[i][1],
        reason:       rows[i][2],
        registeredAt: rows[i][3],
        supplier:     lastCol >= 5 ? (rows[i][4] || '') : ''
      });
    }
  }
  return { ok: true, list: list };
}

function markWip(params) {
  const sheet = getOrCreateSheet(SHEET_WIP);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(params.code)) return { ok: true };
  }
  sheet.appendRow([params.code, params.name, new Date().toLocaleString('ja-JP')]);
  return { ok: true };
}

function unmarkWip(params) {
  const sheet = getOrCreateSheet(SHEET_WIP);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(params.code)) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: true };
}

// ===================== 機能C: 説明文生成（Gemini API） =====================

function getProductsForDescription() {
  const products = bcartGetAll('/products');
  if (!products.ok) return products;

  const categoriesRes = bcartGetAll('/categories');
  const categoryMap = {};
  if (categoriesRes.ok) {
    categoriesRes.data.forEach(c => { categoryMap[String(c.id)] = c.name || String(c.id); });
  }

  const allMapped = products.data.map(p => ({
    id: p.id,
    name: p.name || '',
    category_id: String(p.category_id || ''),
    category_name: categoryMap[String(p.category_id)] || '',
    detail: p.description || '',
    flag: p.flag || '',
    feature_id1: p.feature_id1 || null,
    feature_id2: p.feature_id2 || null,
    feature_id3: p.feature_id3 || null
  }));

  const noDetail  = allMapped.filter(p => !p.detail || p.detail.trim() === '').sort((a, b) => b.id - a.id);
  const hasDetail = allMapped.filter(p => p.detail && p.detail.trim() !== '');
  // ※ detail は内部的に p.description をマッピングしたもの

  return { ok: true, products: noDetail, withDetail: hasDetail.length, total: allMapped.length };
}

// groundingMetadata からソースURL・検索クエリを抽出する共通ヘルパー
function extractGroundingInfo(candidate) {
  const sources = [];
  const queries = [];
  try {
    const meta = candidate.groundingMetadata || {};
    Logger.log('groundingMetadata: ' + JSON.stringify(meta).slice(0, 1000));

    // ソースURL: groundingChunks（標準）
    (meta.groundingChunks || []).forEach(function(chunk) {
      if (chunk.web && chunk.web.uri) {
        sources.push({ uri: chunk.web.uri, title: chunk.web.title || chunk.web.uri });
      }
    });
    // ソースURL: groundingAttributions（旧API）
    if (!sources.length) {
      (meta.groundingAttributions || []).forEach(function(attr) {
        if (attr.web && attr.web.uri) {
          sources.push({ uri: attr.web.uri, title: attr.web.title || attr.web.uri });
        }
      });
    }
    // 検索クエリ（ソースURLがない場合のフォールバック表示用）
    (meta.webSearchQueries || []).forEach(function(q) { queries.push(q); });
  } catch(e) {
    Logger.log('extractGroundingInfo error: ' + e.message);
  }
  return { sources: sources, queries: queries };
}

function generateDescription(params) {
  const apiKey = (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '').trim();
  if (!apiKey) return { ok: false, error: 'GEMINI_API_KEYが設定されていません（スクリプトプロパティ: GEMINI_API_KEY）' };

  const productName  = params.productName  || '';
  const categoryName = params.categoryName || '美容商材';
  const useSearch    = params.useSearch !== false;

  const prompt = useSearch
    ? 'あなたは美容商材の専門ライターです。\n' +
      'Web検索で「' + productName + '」の商品情報を調べ、美容室・エステサロン向けBtoB商材の説明文を日本語で作成してください。\n\n' +
      '商品名: ' + productName + '\n' +
      'カテゴリ: ' + categoryName + '\n\n' +
      '・検索で得た実際の商品情報をもとに、100〜150文字程度で簡潔に説明してください。\n' +
      '・見出しや箇条書きは使わず、自然な文章で書いてください。\n' +
      '・余計な前置き・後書きは不要です。説明文の本文のみ出力してください。\n' +
      '・検索で確認できない情報（成分・数値・効能等）は記載しないこと。'
    : 'あなたは美容商材の専門ライターです。\n' +
      '以下の美容商材（美容室・エステサロン向けBtoB）の商品説明文を日本語で作成してください。\n\n' +
      '商品名: ' + productName + '\n' +
      'カテゴリ: ' + categoryName + '\n\n' +
      '・どのような商品かを簡潔に説明する文章を100〜150文字程度で作成してください。\n' +
      '・見出しや箇条書きは使わず、自然な文章で書いてください。\n' +
      '・余計な前置き・後書きは不要です。説明文の本文のみ出力してください。\n' +
      '・商品名から合理的に推測できる内容のみで作成し、不確かな数値・成分・効能は記載しないこと。';

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey);
  const body = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { maxOutputTokens: 600, temperature: 0.7 }
  };
  if (useSearch) body.tools = [{ google_search: {} }];

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    const httpCode = res.getResponseCode();
    const rawBody  = res.getContentText();

    if (httpCode !== 200) {
      let errCode = '', errStatus = '', errMessage = '', errDetails = '';
      let quotaMetric = '', quotaId = '', quotaDimensions = '', limitValue = null;
      try {
        const errData = JSON.parse(rawBody);
        const e = errData.error || {};
        errCode    = e.code    != null ? String(e.code)    : '';
        errStatus  = e.status  != null ? String(e.status)  : '';
        errMessage = e.message != null ? String(e.message) : '';
        if (Array.isArray(e.details)) {
          errDetails = JSON.stringify(e.details);
          e.details.forEach(function(d) {
            const dtype = d['@type'] || '';
            if (dtype.indexOf('QuotaFailure') !== -1) {
              const v = (d.violations || [])[0] || {};
              quotaMetric     = v.quotaMetric     || v.subject || quotaMetric;
              quotaId         = v.quotaId         || quotaId;
              quotaDimensions = v.quotaDimensions ? JSON.stringify(v.quotaDimensions) : quotaDimensions;
              if (v.quotaValue !== undefined) limitValue = v.quotaValue;
            }
            if (dtype.indexOf('ErrorInfo') !== -1) {
              const m = d.metadata || {};
              if (!quotaMetric) quotaMetric = m.quota_metric || m.quotaMetric || '';
              if (!quotaId)     quotaId     = m.quota_id     || m.quotaId     || '';
              if (limitValue === null && m.limit !== undefined) limitValue = m.limit;
            }
          });
        }
      } catch(parseErr) {}

      var msg = 'Gemini API エラー HTTP ' + httpCode;
      if (errCode)         msg += '\ncode: '              + errCode;
      if (errStatus)       msg += '\nstatus: '            + errStatus;
      if (errMessage)      msg += '\nmessage: '           + errMessage;
      if (errDetails)      msg += '\ndetails: '           + errDetails;
      if (quotaMetric)     msg += '\nquota_metric: '      + quotaMetric;
      if (quotaId)         msg += '\nquota_id: '          + quotaId;
      if (quotaDimensions) msg += '\nquota_dimensions: '  + quotaDimensions;
      if (limitValue !== null) msg += '\nlimit: '         + limitValue;

      if (limitValue === 0 || limitValue === '0') {
        msg += '\n\n→ このプロジェクト/モデルの無料入力トークン枠が0に設定されています（使いすぎではありません）。' +
               'Google AI StudioのプロジェクトでGemini APIの無料枠クォータを確認・申請してください。';
      }

      msg += '\n\nraw: ' + rawBody.slice(0, 1000);
      return { ok: false, error: msg };
    }

    const data = JSON.parse(rawBody);
    const text = data.candidates && data.candidates[0] && data.candidates[0].content &&
                 data.candidates[0].content.parts && data.candidates[0].content.parts[0]
                 ? data.candidates[0].content.parts[0].text : '';
    if (!text) return { ok: false, error: '説明文の生成に失敗しました（空のレスポンス）\n\nraw: ' + rawBody.slice(0, 1000) };

    const grounding = extractGroundingInfo(data.candidates[0]);
    return { ok: true, text: text.trim(), sources: grounding.sources, queries: grounding.queries };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function factCheckDescription(params) {
  const apiKey = (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '').trim();
  if (!apiKey) return { ok: false, error: 'GEMINI_API_KEYが設定されていません' };

  const productName  = params.productName  || '';
  const categoryName = params.categoryName || '美容商材';
  const description  = params.description  || '';
  if (!description) return { ok: false, error: '説明文が指定されていません' };

  const prompt =
    'あなたは商品情報の事実確認専門家です。\n' +
    'Web検索で以下の商品を調査し、説明文の各記述が正確かどうかを判定してください。\n\n' +
    '商品名: ' + productName + '\n' +
    'カテゴリ: ' + categoryName + '\n' +
    '確認する説明文:\n' + description + '\n\n' +
    '結果を必ず以下のJSON形式のみで出力してください（マークダウンのコードブロック不要）:\n' +
    '{"verdict":"ok","summary":"判定コメント（20文字以内）","issues":[]}\n\n' +
    'verdict の選択基準:\n' +
    '"ok"      → 記述内容がWeb検索で確認でき、正確\n' +
    '"warning" → 確認できない記述が一部あるが、明らかな誤りはない\n' +
    '"caution" → 明らかな誤りまたは確認できない重要な記述がある\n' +
    'issues: 問題点を日本語の配列で記載。okの場合は空配列[]。';

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey);
  const body = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { maxOutputTokens: 400, temperature: 0.1 },
    tools: [{ google_search: {} }]
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    const httpCode = res.getResponseCode();
    const rawBody  = res.getContentText();
    if (httpCode !== 200) return { ok: false, error: 'Gemini API エラー HTTP ' + httpCode + '\n' + rawBody.slice(0, 500) };

    const data = JSON.parse(rawBody);
    const text = data.candidates && data.candidates[0] && data.candidates[0].content &&
                 data.candidates[0].content.parts && data.candidates[0].content.parts[0]
                 ? data.candidates[0].content.parts[0].text : '';
    if (!text) return { ok: false, error: 'チェック結果が空でした' };

    const grounding = extractGroundingInfo(data.candidates[0]);

    // JSONを抽出・パース（コードブロック等をフォールバック除去）
    let result;
    try {
      result = JSON.parse(text.trim());
    } catch(e) {
      const m = text.match(/\{[\s\S]*?\}/);
      if (m) { try { result = JSON.parse(m[0]); } catch(e2) {} }
    }
    if (!result) return { ok: false, error: 'チェック結果の解析に失敗しました\n\n' + text.slice(0, 300) };

    return {
      ok: true,
      verdict: result.verdict || 'warning',
      summary: result.summary || '',
      issues:  Array.isArray(result.issues) ? result.issues : [],
      sources: grounding.sources,
      queries: grounding.queries
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function applyDescription(params) {
  if (!params.productId) return { ok: false, error: '商品IDが指定されていません' };
  const description = params.description || '';
  if (!description.trim()) return { ok: false, error: '説明文が空です' };

  const res = bcartPatch('/products/' + params.productId, { description: description });
  addHistory({
    userName: params._userName,
    code:   params.code || '',
    name:   params.name || '',
    type:   '説明文更新',
    before: '',
    after:  description.slice(0, 50) + (description.length > 50 ? '...' : ''),
    result: res.ok ? '成功' : ('失敗: ' + res.error)
  });
  return res;
}

// ===================== LINE WORKS通知（週次タイマー用） =====================
function weeklyCheck() {
  const result = loadData();
  if (!result.ok) return;

  const activeDiffs = result.diffs.filter(d => !d.isIgnored);
  if (activeDiffs.length === 0) return;

  const priceCount = activeDiffs.filter(d => d.issues && d.issues.some(i => i.type === 'price')).length;
  const discCount  = activeDiffs.filter(d => d.issues && d.issues.some(i => i.type === 'discontinued')).length;
  const parentCount = activeDiffs.filter(d => d.issues && d.issues.some(i => i.type === 'parent_visible')).length;
  const unregCount = activeDiffs.filter(d => d.type === 'unregistered').length;

  let msg = `【BCARTマスター管理】差異が${activeDiffs.length}件あります\n💰 価格差異: ${priceCount}件\n🚫 廃番未処理: ${discCount}件\n👻 親商品表示中: ${parentCount}件\n❌ 未登録: ${unregCount}件\n\nBCARTマスター管理ツールで確認してください。`;

  const webhook = PropertiesService.getScriptProperties().getProperty('LINEWORKS_WEBHOOK');
  if (!webhook) return;

  try {
    UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ content: msg }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('LINE WORKS通知エラー: ' + e);
  }
}
