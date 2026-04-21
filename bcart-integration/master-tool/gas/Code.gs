// BCARTマスター管理ツール - バックエンド
// Version: v1.6.0

const VERSION = 'v1.6.0';

// ===================== 設定 =====================
const BCART_BASE_URL = 'https://api.bcart.jp/api/v1';
const CSV_FOLDER_ID = '12QedyGwHcpXF-lEQBgJo5sIC_i3Iv42P';
const CSV_FILENAME = '商品.CSV';

// 認証（beaufield-auth GAS）
const AUTH_GAS_URL = 'https://script.google.com/macros/s/AKfycbzNVW7AaPUTuwneE-M40DTN1clO5VT2yLCHq7cjYvaHqfMfXgVi38UAOsDZQbmmN3wOzw/exec';

// シート名
const SHEET_IGNORE      = '対応不要';
const SHEET_WIP         = '作業中';
const SHEET_HISTORY     = '更新履歴';
const SHEET_SP_GROUPS   = '特別価格_顧客グループ';
const SHEET_SP_DETAILS  = '特別価格_明細';

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
      case 'getSpecialPriceData':      return jsonResponse(getSpecialPriceData());
      case 'saveCustomerGroup':        return jsonResponse(saveCustomerGroup(params));
      case 'deleteCustomerGroup':      return jsonResponse(deleteCustomerGroup(params));
      case 'getProductSetsForFeature': return jsonResponse(getProductSetsForFeature(params));
      case 'applyGroupPrices':         return jsonResponse(applyGroupPrices(params));
      case 'deleteSpecialPriceDetail': return jsonResponse(deleteSpecialPriceDetail(params));
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

    const isIgnored = !!ignoreMap[codeKey];
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
        csvKouri:  parseFloat(String(row['定価1'] || '').replace(/,/g, '')) || 0,
        csvShiire: parseFloat(String(row['仕入単価'] || '').replace(/,/g, '')) || 0,
        csvJan:    (row['JANCD'] || '').trim(),
        csvUnit:   (row['単位名'] || '').trim(),
        stockManagement: row['在庫有無'] || '',
        isIgnored: isIgnored,
        isWip: isWip,
        ignoreReason: ignoreMap[codeKey] || '',
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
      issues.push({ type: 'discontinued' });
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
        ignoreReason: ignoreMap[codeKey] || '',
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
      for (let retry = 0; retry <= 2; retry++) {
        res = UrlFetchApp.fetch(url, {
          method: 'get',
          headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' },
          muteHttpExceptions: true
        });
        if (res.getResponseCode() === 429) {
          if (retry < 2) { Utilities.sleep(3000); continue; }
          return { ok: false, error: 'BCART_API_ERROR: 429 レート制限（しばらく待ってから再読み込みしてください）' };
        }
        break;
      }
      if (res.getResponseCode() !== 200) {
        return { ok: false, error: 'BCART_API_ERROR: ' + res.getResponseCode() };
      }
      const parsed = JSON.parse(res.getContentText());
      const page = parsed.data || parsed.product_sets || parsed.products || parsed.specials || parsed.product_stock || parsed.categories || parsed;
      if (!Array.isArray(page) || page.length === 0) break;
      allData.push(...page);
      if (page.length < limit) break;
      offset += limit;
      if (offset > 0) Utilities.sleep(300);
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
    return { ok: true, data: data.data || data };
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
  const res = bcartPatch('/product_sets/' + params.bcartSetId, { jodai: params.jodai });
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

function setStock(params) {
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
      res = bcartPatch('/product_sets/' + item.bcartSetId, { jodai: item.jodai });
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
      const res = bcartGet(ep);
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
      min_order:   1,
      stock_flag:  1,
      tax_type_id: params.taxTypeId || 1,
      set_flag:    params.setFlag || '非表示'
    }]
  };

  const step2 = bcartPost('/product_sets', setBody);
  if (!step2.ok) {
    addHistory({
      userName: params._userName,
      code: params.code || '',
      name: params.productName || '',
      type: '新規登録（セット作成失敗・孤立商品ID: ' + createdProductId + '）',
      before: '', after: '商品ID: ' + createdProductId,
      result: '失敗: ' + step2.error
    });
    return { ok: false, error: step2.error, orphanProductId: createdProductId };
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

// ===================== 機能B: 特別価格管理 =====================

function getSpecialPriceData() {
  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  const groups = [];
  for (let i = 1; i < groupRows.length; i++) {
    if (groupRows[i][0]) {
      groups.push({
        group_id:   String(groupRows[i][0]),
        group_name: String(groupRows[i][1]),
        member_ids: String(groupRows[i][2] || ''),
        created_at: String(groupRows[i][3]),
        note:       groupRows[i][4] || ''
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
        updated_at:       String(detailRows[i][6])
      });
    }
  }

  return { ok: true, groups: groups, details: details };
}

function saveCustomerGroup(params) {
  const sheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const rows = sheet.getDataRange().getValues();

  if (params.group_id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === params.group_id) {
        sheet.getRange(i + 1, 2).setValue(params.group_name);
        sheet.getRange(i + 1, 3).setValue(params.member_ids);
        sheet.getRange(i + 1, 5).setValue(params.note || '');
        return { ok: true, group_id: params.group_id };
      }
    }
  }

  const newId = 'G' + new Date().getTime().toString(36).toUpperCase();
  sheet.appendRow([newId, params.group_name, params.member_ids, new Date().toLocaleString('ja-JP'), params.note || '']);
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
    memberIds.forEach(mid => { newSp[String(mid)] = { unit_price: item.unit_price }; });

    const res = bcartPatch('/product_sets/' + item.product_set_id, { special_price: newSp });
    if (res.ok) {
      successCount++;
      let found = false;
      for (let i = 1; i < detailRows.length; i++) {
        if (String(detailRows[i][1]) === params.group_id && detailRows[i][2] == item.product_set_id) {
          detailSheet.getRange(i + 1, 5).setValue(item.product_set_name || '');
          detailSheet.getRange(i + 1, 6).setValue(item.unit_price);
          detailSheet.getRange(i + 1, 7).setValue(new Date().toLocaleString('ja-JP'));
          detailRows[i][5] = item.unit_price;
          found = true;
          break;
        }
      }
      if (!found) {
        const newId = 'D' + new Date().getTime().toString(36).toUpperCase();
        const newRow = [newId, params.group_id, item.product_set_id, item.product_no || '', item.product_set_name || '', item.unit_price, new Date().toLocaleString('ja-JP')];
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
      sheet.appendRow(['group_id', 'group_name', 'member_ids', 'created_at', 'note']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_SP_DETAILS) {
      sheet.appendRow(['detail_id', 'group_id', 'product_set_id', 'product_no', 'product_set_name', 'unit_price', 'updated_at']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f3f4f6');
    }
  } else if (sheetName === SHEET_IGNORE) {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 5) {
      sheet.getRange(1, 5).setValue('仕入先名');
    } else {
      const headerVal = sheet.getRange(1, 5).getValue();
      if (!headerVal) sheet.getRange(1, 5).setValue('仕入先名');
    }
  }
  return sheet;
}

function getIgnoreMap() {
  const sheet = getOrCreateSheet(SHEET_IGNORE);
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) map[rows[i][0]] = rows[i][2] || '';
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
    if (rows[i][0] === params.code) return { ok: true };
  }
  sheet.appendRow([params.code, params.name, new Date().toLocaleString('ja-JP')]);
  return { ok: true };
}

function unmarkWip(params) {
  const sheet = getOrCreateSheet(SHEET_WIP);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === params.code) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: true };
}

// ===================== LINE WORKS通知（週次タイマー用） =====================
function weeklyCheck() {
  const result = loadData();
  if (!result.ok) return;

  const activeDiffs = result.diffs.filter(d => !d.isIgnored);
  if (activeDiffs.length === 0) return;

  const priceCount = activeDiffs.filter(d => d.issues && d.issues.some(i => i.type === 'price')).length;
  const discCount  = activeDiffs.filter(d => d.issues && d.issues.some(i => i.type === 'discontinued')).length;
  const unregCount = activeDiffs.filter(d => d.type === 'unregistered').length;

  const msg = `【BCARTマスター管理】差異が${activeDiffs.length}件あります\n💰 価格差異: ${priceCount}件\n🚫 廃番未処理: ${discCount}件\n❌ 未登録: ${unregCount}件\n\nBCARTマスター管理ツールで確認してください。`;

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
