// BCARTマスター管理ツール - バックエンド
//
// [スクリプトプロパティに設定が必要]
//   BCART_TOKEN       : BCARTアクセストークン
//   GEMINI_API_KEY    : Google Gemini APIキー
//   LINEWORKS_WEBHOOK : LINE WORKS Webhook URL（任意）
//   CSV_FOLDER_ID     : 商品.CSV保管Driveフォルダ ID
//   AUTH_GAS_URL      : portal GAS WebApp URL（セッション検証用）

const VERSION = 'v2.17.0';

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
const SHEET_SP_INDIVIDUAL = '特別価格_個別';
const SHEET_VF_DETAILS  = '例外表示_明細';
const SHEET_FEATURES    = '特集_管理';
const SHEET_DESC_SKIP   = '説明文不要';
const SHEET_DRAFT       = '登録ドラフト';
const SHEET_DRAFT_SETS  = '登録ドラフト_セット';

// ===================== エントリポイント =====================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;

    const noAuthActions = ['getVersion'];
    const claudeActions = ['previewSuffixName', 'applySuffixName', 'previewHanbaiEnd', 'applyHanbaiEnd', 'previewSetDescription', 'applySetDescription', 'previewProductFields', 'applyProductFields', 'previewSetFields', 'applySetFields', 'previewProductSort', 'applyProductSort', 'getDraftSupplierSummary', 'getDraftCandidates', 'getRegisteredExamples', 'saveDrafts'];
    let userName = '不明';
    if (claudeActions.includes(action)) {
      const claudeKey = _BCART_PROPS.getProperty('CLAUDE_API_KEY');
      if (!claudeKey || params.apiKey !== claudeKey) return jsonResponse({ ok: false, error: 'UNAUTHORIZED' });
      userName = 'Claude(API)';
    } else if (!noAuthActions.includes(action)) {
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
      case 'bulkUpdateVisibility': return jsonResponse(bulkUpdateVisibility(params));
      case 'bulkIgnore':           return jsonResponse(bulkIgnore(params));
      case 'bulkMarkWip':          return jsonResponse(bulkMarkWip(params));
      case 'searchProducts':       return jsonResponse(searchProducts(params));
      case 'cloneProduct':         return jsonResponse(cloneProduct(params));
      case 'getSpecials':          return jsonResponse(getSpecials());
      case 'updateProduct':        return jsonResponse(updateProductAction(params));
      case 'deleteProduct':        return jsonResponse(deleteProduct(params));
      case 'getHistory':           return jsonResponse(getHistory());
      case 'debugData':
      case 'debugCode':
      case 'debugProduct':
        if (PropertiesService.getScriptProperties().getProperty('DEBUG_ENABLED') !== 'true') {
          return jsonResponse({ ok: false, error: 'UNAUTHORIZED' });
        }
        if (action === 'debugData')    return jsonResponse(debugData());
        if (action === 'debugCode')    return jsonResponse(debugCode(params));
        return jsonResponse(debugProduct());
      // 機能A: 新規登録
      case 'getCategories':        return jsonResponse(getCategories());
      case 'getFeatures':          return jsonResponse(getSpecials());
      case 'registerProduct':      return jsonResponse(registerProduct(params));
      case 'addSetToProduct':      return jsonResponse(addSetToProduct(params));
      case 'bulkRegisterProduct':  return jsonResponse(bulkRegisterProduct(params));
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
      case 'saveIndividualPrices':      return jsonResponse(saveIndividualPrices(params));
      case 'applyIndividualPrices':     return jsonResponse(applyIndividualPrices(params));
      case 'deleteIndividualPrice':     return jsonResponse(deleteIndividualPrice(params));
      case 'auditSpecialPrices':        return jsonResponse(auditSpecialPrices(params));
      // 機能C: 説明文生成
      case 'getProductsForDescription': return jsonResponse(getProductsForDescription(params));
      case 'getSimilarProducts':        return jsonResponse(getSimilarProducts(params));
      case 'generateDescription':       return jsonResponse(generateDescription(params));
      case 'factCheckDescription':      return jsonResponse(factCheckDescription(params));
      case 'applyDescription':          return jsonResponse(applyDescription(params));
      case 'markDescSkip':              return jsonResponse(markDescSkip(params));
      case 'unmarkDescSkip':            return jsonResponse(unmarkDescSkip(params));
      // 機能D: 特集管理
      case 'getFeatureList':          return jsonResponse(getFeatureList());
      case 'createFeature':           return jsonResponse(createFeature(params));
      case 'updateFeature':           return jsonResponse(updateFeature(params));
      case 'bulkUpdateFeatureOrder':  return jsonResponse(bulkUpdateFeatureOrder(params));
      case 'saveFeatureType':         return jsonResponse(saveFeatureType(params));
      case 'bulkSaveFeatureTypes':    return jsonResponse(bulkSaveFeatureTypes(params));
      // 機能E: Claudeチャット直接操作
      case 'previewSuffixName':      return jsonResponse(previewSuffixName(params));
      case 'applySuffixName':        return jsonResponse(applySuffixName(params));
      case 'previewHanbaiEnd':       return jsonResponse(previewHanbaiEnd(params));
      case 'applyHanbaiEnd':         return jsonResponse(applyHanbaiEnd(params));
      case 'previewSetDescription':  return jsonResponse(previewSetDescription(params));
      case 'applySetDescription':    return jsonResponse(applySetDescription(params));
      case 'previewProductFields':   return jsonResponse(previewProductFields(params));
      case 'applyProductFields':     return jsonResponse(applyProductFields(params));
      case 'previewSetFields':       return jsonResponse(previewSetFields(params));
      case 'applySetFields':         return jsonResponse(applySetFields(params));
      case 'previewProductSort':     return jsonResponse(previewProductSort(params));
      case 'applyProductSort':       return jsonResponse(applyProductSort(params));
      // 商品登録ドラフト（AI叩き台作成・フェーズ1）
      case 'getDraftSupplierSummary': return jsonResponse(getDraftSupplierSummary());
      case 'getDraftCandidates':      return jsonResponse(getDraftCandidates(params));
      case 'getRegisteredExamples':   return jsonResponse(getRegisteredExamples(params));
      case 'saveDrafts':              return jsonResponse(saveDrafts(params));
      // 登録ドラフト フェーズ2（ポータルセッション認証必須。claudeActionsに追加禁止）
      case 'getDrafts':               return jsonResponse(getDrafts(params));
      case 'updateDraft':             return jsonResponse(updateDraft(params));
      case 'approveDraft':            return jsonResponse(approveDraft(params));
      case 'rejectDraft':             return jsonResponse(rejectDraft(params));
      case 'publishDraft':            return jsonResponse(publishDraft(params));
      default:                         return jsonResponse({ ok: false, error: 'UNKNOWN_ACTION' });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return jsonResponse({ ok: false, error: 'INTERNAL_ERROR' });
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
    Logger.log('loadCsvFromDrive error: ' + e.message);
    return { ok: false, error: 'INTERNAL_ERROR' };
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

function bcartGetAll(path, extraParams) {
  try {
    const token = getBcartToken();
    const allData = [];
    const limit = 100;
    let offset = 0;
    const extraQs = extraParams
      ? Object.entries(extraParams).map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join('&')
      : '';

    while (true) {
      const url = BCART_BASE_URL + path + '?limit=' + limit + '&offset=' + offset + (extraQs ? '&' + extraQs : '');
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
          Logger.log('bcartGetAll 502/503: ' + res.getContentText().slice(0, 300));
          return { ok: false, error: 'BCART_API_ERROR: ' + code + ' 帯域幅エラー（しばらく待ってから再読み込みしてください）' };
        }
        break;
      }
      if (res.getResponseCode() !== 200) {
        Logger.log('bcartGetAll error: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
        return { ok: false, error: 'BCART_API_ERROR: ' + res.getResponseCode() };
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
    Logger.log('bcartGetAll error: ' + e.message);
    return { ok: false, error: 'INTERNAL_ERROR' };
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
    return { ok: true, data: data.data || data.product_set || data.product || data };
  } catch (e) {
    Logger.log('bcartGet error: ' + e.message);
    return { ok: false, error: 'INTERNAL_ERROR' };
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
    if (code !== 200 && code !== 204) {
      const bodyText = res.getContentText().slice(0, 300);
      Logger.log('bcartPatch error: ' + code + ' ' + bodyText);
      return { ok: false, error: 'BCART_API_ERROR: ' + code + ' ' + bodyText.slice(0, 200) };
    }
    return { ok: true };
  } catch (e) {
    Logger.log('bcartPatch error: ' + e.message);
    return { ok: false, error: 'INTERNAL_ERROR' };
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
    Logger.log('bcartDelete error: ' + e.message);
    return { ok: false, error: 'INTERNAL_ERROR' };
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
      const bodyText = res.getContentText().slice(0, 300);
      Logger.log('bcartPost error: ' + code + ' ' + bodyText);
      return { ok: false, error: 'BCART_API_ERROR: ' + code + ' ' + bodyText.slice(0, 200) };
    }
    return { ok: true, data: JSON.parse(res.getContentText()) };
  } catch (e) {
    Logger.log('bcartPost error: ' + e.message);
    return { ok: false, error: 'INTERNAL_ERROR' };
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
  // 品番が空だとBCARTが422（product_no必須）を返すため、手前で弾いて分かりやすく案内する
  if (!params.productNo) {
    return { ok: false, error: 'この商品は品番が登録されていないため、欠品設定できません（BCART側で品番を設定してください）' };
  }
  // 品番がBCARTに存在するか事前確認
  const checkRes = bcartGet('/product_stock/' + encodeURIComponent(params.productNo));
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

  // 仕様: 「欠品（在庫0）」= 通常在庫管理(stock_flag:0)で在庫0＝売り切れ表示。
  //       「在庫あり」= 無制限(stock_flag:1)＝在庫数に関係なく常に購入可（stockは送らない/BCART仕様上無視）。
  // 公式仕様: 本文は { product_stock: [...] } で包むのが必須。
  const stockEntry = (params.stock === 0)
    ? { product_no: params.productNo, stock_flag: 0, stock: 0 }
    : { product_no: params.productNo, stock_flag: 1 };
  const res = bcartPatch('/product_stock', { product_stock: [stockEntry] });
  addHistory({
    userName: params._userName,
    code: params.productNo || '',
    name: params.name || '',
    type: '欠品設定',
    before: '',
    after: params.stock === 0 ? '欠品（在庫0）' : '在庫あり（無制限）',
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

function bulkUpdateVisibility(params) {
  const flag = params.flag === '表示' ? '表示' : '非表示';
  const results = [];
  (params.items || []).forEach(item => {
    const res = bcartPatch('/products/' + item.id, { flag: flag });
    addHistory({
      userName: params._userName,
      code: item.code || '',
      name: item.name || '',
      type: '表示切替（一括）',
      before: item.beforeFlag || '',
      after: flag,
      result: res.ok ? '成功' : ('失敗: ' + res.error)
    });
    results.push({ id: item.id, ok: res.ok, error: res.ok ? '' : res.error });
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

  if (params.specialId) {
    const fid = String(params.specialId);
    filtered = filtered.filter(p =>
      String(p.feature_id1 || '') === fid ||
      String(p.feature_id2 || '') === fid ||
      String(p.feature_id3 || '') === fid
    );
  }

  // 実際に登録されている在庫(stock_flag/stock)を product_stock から取得してマップ化
  // - withStock : 欠品設定タブで現在の在庫状態を表示するため
  // - stockZero : 商品検索タブの「在庫0のみ」で、無制限(stock_flag:1)等を除いた
  //               「本当の欠品（通常在庫管理 stock_flag:0 かつ stock:0）」だけに絞り込むため
  // どちらの指定もない通常の商品検索では取得しない（速度を落とさないため）
  let stockByNo = null;
  if ((params.withStock || params.stockZero) && filtered.length) {
    const ps = bcartGetAll('/product_stock');
    if (ps.ok && Array.isArray(ps.data)) {
      stockByNo = {};
      ps.data.forEach(row => {
        if (row.product_no === undefined || row.product_no === null || row.product_no === '') return;
        const key = String(row.product_no);
        const prev = stockByNo[key];
        // 同一品番が複数行ある場合は通常在庫管理(stock_flag:0)の行を優先（欠品設定で操作する対象に合わせる）
        if (!prev || (String(prev.stock_flag) !== '0' && String(row.stock_flag) === '0')) {
          stockByNo[key] = { stock_flag: row.stock_flag, stock: row.stock };
        }
      });
    }
  }

  // 在庫0のみ: 通常在庫管理(stock_flag:0)かつ在庫数0＝実際に欠品している商品だけを残す。
  // 無制限(1)・別在庫参照(100)は在庫0表示でも常に購入可なので除外。
  // 在庫情報(product_stock)が無い商品も「欠品と確認できない」ため除外（仕様確認済 2026-06-14）。
  if (params.stockZero) {
    filtered = filtered.filter(p => {
      const s = setByProductId[p.id] || {};
      // 品番は親商品(products)側が空のセット商品があるためセット側へフォールバック（mapと同一ロジック）
      const pno = p.product_no || p.main_no || s.product_no || '';
      const live = stockByNo ? stockByNo[String(pno)] : null;
      if (!live) return false;                            // 在庫情報なし → 欠品扱いしない
      if (String(live.stock_flag) !== '0') return false;  // 無制限/別在庫参照 → 除外
      const st = live.stock;
      return st !== undefined && st !== null && st !== '' && parseInt(st) === 0;
    });
  }

  const result = filtered.map(p => {
    const s = setByProductId[p.id] || {};
    // 品番は親商品(products)側が空のセット商品があるため、セット(product_sets)側の品番でフォールバック
    const pno = p.product_no || p.main_no || s.product_no || '';
    const live = stockByNo ? stockByNo[String(pno)] : null;
    return {
      id: p.id,
      product_no: pno,
      name: p.name || '',
      flag: p.flag || '',                    // 親商品(products)の表示/非表示
      set_flag: s.set_flag || '',            // 商品セット(product_sets)の表示/非表示
      display_start: p.hanbai_start || '',
      display_end: getDisplayEnd(p),
      unit_price: s.unit_price || '',
      // 在庫数: withStock時は product_stock の実値、それ以外は従来通りセット側の値
      stock: live ? live.stock : (s.stock !== undefined ? s.stock : (s.inventory !== undefined ? s.inventory : '')),
      // 在庫フラグ: withStock時のみ product_stock から付与（0=通常/1=無制限/100=別在庫参照）
      stock_flag: live ? live.stock_flag : undefined,
      set_id: s.id || '',
      category_id: p.category_id || null,
      feature_id1: p.feature_id1 || null,
      feature_id2: p.feature_id2 || null,
      feature_id3: p.feature_id3 || null
    };
  });

  return { ok: true, products: result, total: result.length };
}

// 商品複製（商品基本情報 + 全セット、全フィールドを複製）
// 親商品は必ず非表示で作成（確認後に手動公開する運用）。商品セットは元の表示状態(set_flag)をそのまま複製する。
function cloneProduct(params) {
  const productId = String(params.productId);

  // 元商品を全フィールド取得
  const srcRes = bcartGet('/products/' + productId);
  if (!srcRes.ok) return srcRes;
  const src = srcRes.data;
  if (!src || !src.id) {
    return { ok: false, error: '元商品が見つかりませんでした（ID: ' + productId + '）' };
  }

  // 元商品の全セットを取得して複製対象を絞り込み
  const setsRes = bcartGetAll('/product_sets', { product_id: productId });
  if (!setsRes.ok) return setsRes;
  const srcSets = setsRes.data.filter(s => String(s.product_id) === productId);

  // sub_images: GETはオブジェクト形式 {1:{image,caption},...} → POSTは配列形式
  const subImages = [];
  if (src.sub_images && typeof src.sub_images === 'object') {
    Object.keys(src.sub_images).forEach(k => {
      const si = src.sub_images[k];
      if (si && (si.image || si.caption)) {
        subImages.push({ image: si.image || '', caption: si.caption || '' });
      }
    });
  }

  // 新商品作成（表示状態(flag)以外は元商品の全フィールドをそのまま複製）
  const productBody = {
    products: [{
      main_no:              src.main_no || '',
      name:                 params.productName || src.name || '',
      catch_copy:           src.catch_copy || '',
      category_id:          src.category_id || null,
      sub_category_id:      src.sub_category_id || '',
      feature_id1:          src.feature_id1 || null,
      feature_id2:          src.feature_id2 || null,
      feature_id3:          src.feature_id3 || null,
      made_in:              src.made_in || '',
      size:                 src.size || '',
      sozai:                src.sozai || '',
      caution:              src.caution || '',
      tag:                  src.tag || '',
      description:          src.description || '',
      meta_title:           src.meta_title || '',
      meta_keywords:        src.meta_keywords || '',
      meta_description:     src.meta_description || '',
      image:                src.image || '',
      view_group_filter:    src.view_group_filter || '',
      visible_customer_id:  src.visible_customer_id || '',
      hide_customer_id:     src.hide_customer_id || '',
      prepend_text:         src.prepend_text || '',
      middle_text:          src.middle_text || '',
      append_text:          src.append_text || '',
      rv_prepend_text:      src.rv_prepend_text || '',
      rv_middle_text:       src.rv_middle_text || '',
      rv_append_text:       src.rv_append_text || '',
      file_download:        src.file_download || null,
      customs:              src.customs || [],
      hanbai_start:         src.hanbai_start || null,
      hanbai_end:           src.hanbai_end || null,
      recommend_product_id: src.recommend_product_id || '',
      view_pattern:         src.view_pattern || 0,
      priority:             src.priority || 0,
      flag:                 '非表示',
      sub_images:           subImages
    }]
  };

  const step1 = bcartPost('/products', productBody);
  if (!step1.ok) {
    addHistory({
      userName: params._userName,
      code: params.productNo || '',
      name: params.productName || '',
      type: '商品複製（商品作成失敗）',
      before: '元商品ID: ' + productId,
      after: '',
      result: '失敗: ' + step1.error
    });
    return step1;
  }

  const newProductId = step1.data && step1.data.products && step1.data.products[0]
    ? step1.data.products[0].id : null;
  if (!newProductId) {
    return { ok: false, error: '新商品IDが取得できませんでした（レスポンス: ' + JSON.stringify(step1.data) + '）' };
  }

  // 全セットを新商品へ複製（表示状態(set_flag)は元のセットのまま複製）
  const results = [];
  let successCount = 0;

  for (const s of srcSets) {
    // group_price: GETレスポンスにはname/rate等の付随情報が混じるため、POSTに必要な
    // fixed_price/volume_discountだけに整形する（fixed_price未設定＝掛け率運用のグループは送らない）
    const groupPrice = {};
    if (s.group_price && typeof s.group_price === 'object') {
      Object.keys(s.group_price).forEach(gid => {
        const gp = s.group_price[gid] || {};
        if (gp.fixed_price === undefined || gp.fixed_price === null) return;
        const entry = { fixed_price: gp.fixed_price };
        if (gp.volume_discount) entry.volume_discount = gp.volume_discount;
        groupPrice[gid] = entry;
      });
    }

    const setBody = {
      product_sets: [{
        product_id:          newProductId,
        product_no:          s.product_no || '',
        jan_code:            s.jan_code || '',
        location_no:         s.location_no || '',
        jodai_type:          s.jodai_type || 'メーカー希望小売価格',
        jodai:               s.jodai || 0,
        name:                s.name || '',
        unit_price:          s.unit_price !== undefined ? s.unit_price : 0,
        min_order:           s.min_order || null,
        max_order:           s.max_order || null,
        group_price:         groupPrice,
        special_price:       s.special_price || {},
        volume_discount:     s.volume_discount || {},
        quantity:            s.quantity || 1,
        unit:                s.unit || '',
        description:         s.description || '',
        stock:               s.stock !== undefined ? s.stock : null,
        stock_flag:          s.stock_flag !== undefined ? s.stock_flag : 1,
        stock_parent:        s.stock_parent || null,
        stock_view_id:       s.stock_view_id || null,
        stock_few:           s.stock_few || 0,
        view_group_filter:   s.view_group_filter || '',
        visible_customer_id: s.visible_customer_id || '',
        hide_customer_id:    s.hide_customer_id || '',
        customs:             s.customs || [],
        option_ids:          s.option_ids || [],
        shipping_group_id:   s.shipping_group_id || null,
        shipping_size:       s.shipping_size || 0,
        priority:            s.priority || 0,
        tax_type_id:         s.tax_type_id || 1,
        set_flag:            s.set_flag || '表示'
      }]
    };
    const res = bcartPost('/product_sets', setBody);
    const setId = res.ok && res.data && res.data.product_sets && res.data.product_sets[0]
      ? res.data.product_sets[0].id : null;
    results.push({ product_no: s.product_no || '', ok: res.ok, setId: setId, error: res.ok ? null : (res.error || '') });
    if (res.ok) successCount++;
  }

  // 全セット失敗時はロールバック
  if (successCount === 0 && srcSets.length > 0) {
    bcartDelete('/products/' + newProductId);
    addHistory({
      userName: params._userName,
      code: params.productNo || '',
      name: params.productName || '',
      type: '商品複製（ロールバック完了）',
      before: '元商品ID: ' + productId,
      after: '新商品ID: ' + newProductId + ' を削除',
      result: '失敗: 全セット複製失敗'
    });
    return { ok: false, error: '全セットの複製に失敗したため商品を削除しました' };
  }

  addHistory({
    userName: params._userName,
    code: params.productNo || '',
    name: params.productName || '',
    type: '商品複製',
    before: '元商品ID: ' + productId,
    after: '新商品ID: ' + newProductId + ' / ' + successCount + '/' + srcSets.length + 'セット複製',
    result: successCount === srcSets.length ? '成功' : '一部失敗'
  });

  return { ok: true, newProductId: newProductId, setCount: srcSets.length, successCount: successCount, results: results };
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
          if (specials.length > 0) {
            const typeMap = getFeatureTypeMap_();
            specials.forEach(f => { f.type = typeMap[String(f.id)] || ''; });
            return { ok: true, specials: specials };
          }
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

// 上代(csvKouri)が無い(0/null)場合は「参考上代タイプ: 非表示」とし、jodai/group_priceの
// メーカー希望小売価格を送らない（0円をメーカー希望小売価格として送るとBCART APIの
// バリデーション(422)に弾かれるため。2026-07-12 登録ドラフト承認時に発覚）
function buildJodaiFields_(csvKouri, csvShiire, jodaiType) {
  const fields = {};
  if (csvKouri) {
    fields.jodai = csvKouri;
    fields.jodai_type = jodaiType || 'メーカー希望小売価格';
    fields.group_price = { '1': { fixed_price: csvKouri } };
  } else {
    fields.jodai_type = '非表示';
  }
  if (csvShiire) {
    fields.group_price = fields.group_price || {};
    fields.group_price['10'] = { fixed_price: csvShiire };
  }
  return fields;
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
    product_sets: [Object.assign({
      product_id:  createdProductId,
      product_no:  params.code,
      name:        params.setName,
      jan_code:    params.janCode || '',
      unit_price:  params.csvPrice,
      unit:        params.csvUnit || '',
      quantity:    1,
      min_order:   1,
      stock_flag:  1,
      tax_type_id: params.taxTypeId || 1,
      set_flag:    params.setFlag || '非表示'
    }, buildJodaiFields_(params.csvKouri, params.csvShiire, params.jodaiType))]
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

// 既存のBCART商品にセットを追加
function addSetToProduct(params) {
  const setBody = {
    product_sets: [Object.assign({
      product_id:  params.productId,
      product_no:  params.code,
      name:        params.setName,
      jan_code:    params.janCode || '',
      unit_price:  params.csvPrice,
      unit:        params.csvUnit || '',
      quantity:    1,
      min_order:   1,
      stock_flag:  1,
      tax_type_id: params.taxTypeId || 1,
      set_flag:    params.setFlag || '非表示'
    }, buildJodaiFields_(params.csvKouri, params.csvShiire, params.jodaiType))]
  };

  const res = bcartPost('/product_sets', setBody);
  if (!res.ok) {
    addHistory({
      userName: params._userName,
      code: params.code || '',
      name: params.setName || '',
      type: '既存商品にセット追加（失敗）',
      before: '', after: '',
      result: '失敗: ' + res.error
    });
    return res;
  }

  const createdSetId = res.data && res.data.product_sets && res.data.product_sets[0]
    ? res.data.product_sets[0].id : null;

  unmarkWip({ code: params.code });

  addHistory({
    userName: params._userName,
    code: params.code || '',
    name: params.setName || '',
    type: '既存商品にセット追加',
    before: '',
    after: '商品ID: ' + params.productId + ' / セットID: ' + createdSetId,
    result: '成功'
  });

  return { ok: true, setId: createdSetId };
}

// 複数商品セットを紐付けた新規BCART商品を一括登録
function bulkRegisterProduct(params) {
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
      code: (params.sets || []).map(s => s.code).join(','),
      name: params.productName,
      type: '一括新規登録（商品作成失敗）',
      before: '', after: '', result: '失敗: ' + step1.error
    });
    return step1;
  }

  const createdProductId = step1.data && step1.data.products && step1.data.products[0]
    ? step1.data.products[0].id : null;
  if (!createdProductId) {
    return { ok: false, error: '商品IDが取得できませんでした' };
  }

  // 各セットを順次登録
  const sets = params.sets || [];
  const results = [];
  let successCount = 0;

  for (const s of sets) {
    const setBody = {
      product_sets: [Object.assign({
        product_id:  createdProductId,
        product_no:  s.code,
        name:        s.setName,
        jan_code:    s.janCode || '',
        unit_price:  s.csvPrice,
        unit:        s.csvUnit || '',
        quantity:    1,
        min_order:   1,
        stock_flag:  1,
        tax_type_id: params.taxTypeId || 1,
        set_flag:    params.setFlag || '非表示'
      }, buildJodaiFields_(s.csvKouri, s.csvShiire, params.jodaiType))]
    };
    const res = bcartPost('/product_sets', setBody);
    const setId = res.ok && res.data && res.data.product_sets && res.data.product_sets[0]
      ? res.data.product_sets[0].id : null;
    results.push({ code: s.code, ok: res.ok, setId, error: res.ok ? null : (res.error || '') });
    if (res.ok) successCount++;
  }

  // 全セット失敗時はロールバック
  if (successCount === 0 && sets.length > 0) {
    bcartDelete('/products/' + createdProductId);
    const firstError = (results.find(r => !r.ok) || {}).error || '不明';
    addHistory({
      userName: params._userName,
      code: sets.map(s => s.code).join(','),
      name: params.productName,
      type: '一括新規登録（ロールバック完了）',
      before: '', after: '商品ID: ' + createdProductId + ' を削除',
      result: '失敗: 全セット登録失敗（' + firstError + '）'
    });
    return { ok: false, error: '全セットの登録に失敗したため商品を削除しました（' + firstError + '）', results: results };
  }

  // 成功したセットのWIPを解除
  results.filter(r => r.ok).forEach(r => unmarkWip({ code: r.code }));

  addHistory({
    userName: params._userName,
    code: sets.map(s => s.code).join(','),
    name: params.productName,
    type: '一括新規登録',
    before: '',
    after: '商品ID: ' + createdProductId + ' / ' + successCount + '/' + sets.length + 'セット成功',
    result: successCount === sets.length ? '成功' : '一部失敗'
  });

  return { ok: true, productId: createdProductId, results };
}

// ===================== 商品登録ドラフト（AI叩き台作成・フェーズ1） =====================

// 未登録商品の抽出（calcDiffsのunregistered判定と同一基準。ドラフト用に項目を拡張）
function getUnregisteredForDraft() {
  const csvData = loadCsvFromDrive();
  if (!csvData.ok) return csvData;
  const bcartSets = bcartGetAll('/product_sets');
  if (!bcartSets.ok) return bcartSets;
  const ignoreMap = getIgnoreMap();

  const bcartSetMap = {};
  bcartSets.data.forEach(s => { bcartSetMap[s.product_no] = s; });

  const rows = [];
  csvData.rows.forEach(row => {
    const code = row['コード'];
    if (!code) return;
    const codeKey = String(parseInt(code, 10) || code);
    if (bcartSetMap[codeKey]) return;
    if (ignoreMap[codeKey]) return;
    const isDiscontinued = row['廃番'] === '1' || row['廃番'] === 'TRUE' || row['廃番'] === '廃番';
    if (isDiscontinued) return;
    if ((row['在庫有無'] || '') !== 'する') return;  // 差異一覧のデフォルトフィルターに合わせる（HANDOVER.md 差異検出除外ルール）
    const price = parseFloat(String(row['売上単価'] || '').replace(/,/g, '')) || 0;
    if (!price) return;

    rows.push({
      code: codeKey,
      name: row['商品名'] || row['略称'] || '',
      kana: row['かな'] || '',
      ryakusho: row['略称'] || '',
      unit: (row['単位名'] || '').trim(),
      supplierCd: (row['仕入先CD'] || '').trim(),
      supplierName: row['仕入先名'] || '',
      price: price,
      kouri: parseFloat(String(row['定価１'] || row['定価1'] || '').replace(/,/g, '')) || 0,
      shiire: parseFloat(String(row['仕入単価'] || '').replace(/,/g, '')) || 0,
      jan: (row['JANCD'] || '').trim(),
      lastSaleKey: lastSaleDateKey_(row['最終売上日']),
      stockManagement: row['在庫有無'] || ''
    });
  });

  return { ok: true, rows: rows };
}

// 最終売上日をYYYYMMDD文字列に正規化（比較・ソート用）。不明/実績なしはnull。
// ⚠️ 商品.CSVでの実際の表記は未検証（設計書§7-9）。同源データのsales-dbはYYYYMMDD8桁・実績なしは00000000。
// スラッシュ/ハイフン区切りにもフォールバック対応するが、デプロイ後に実データで要検証。
function lastSaleDateKey_(raw) {
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return null;
  if (/^\d{8}$/.test(s)) return s === '00000000' ? null : s;
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    const key = m[1] + ('0' + m[2]).slice(-2) + ('0' + m[3]).slice(-2);
    return key === '00000000' ? null : key;
  }
  return null;
}

function dateKeyDaysAgo_(days) {
  const d = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
  return '' + d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
}

function dateKeyToSlash_(key) {
  if (!key) return '';
  return key.slice(0, 4) + '/' + key.slice(4, 6) + '/' + key.slice(6, 8);
}

// getDraftSupplierSummary — 仕入先別の未登録件数サマリー（バッチ計画用）
function getDraftSupplierSummary() {
  const data = getUnregisteredForDraft();
  if (!data.ok) return data;

  const cutoff1y = dateKeyDaysAgo_(365);
  const cutoff2y = dateKeyDaysAgo_(730);
  const bySupplier = {};

  data.rows.forEach(r => {
    const key = r.supplierCd || '(不明)';
    if (!bySupplier[key]) {
      bySupplier[key] = { supplierCd: r.supplierCd, supplierName: r.supplierName, countWithin1y: 0, countWithin2y: 0, countTotal: 0 };
    }
    bySupplier[key].countTotal++;
    if (r.lastSaleKey && r.lastSaleKey >= cutoff1y) bySupplier[key].countWithin1y++;
    if (r.lastSaleKey && r.lastSaleKey >= cutoff2y) bySupplier[key].countWithin2y++;
  });

  const suppliers = Object.keys(bySupplier).map(k => bySupplier[k])
    .sort((a, b) => b.countWithin1y - a.countWithin1y);

  return { ok: true, totalUnregistered: data.rows.length, suppliers: suppliers };
}

// getDraftCandidates — ドラフト対象候補の取得（仕入先・最終売上日で絞り込み、最終売上日降順）
function getDraftCandidates(params) {
  const data = getUnregisteredForDraft();
  if (!data.ok) return data;

  const supplierCd = params.supplierCd != null ? String(params.supplierCd).trim() : '';
  const withinDays = params.lastSaleWithinDays || 365;
  const cutoff = dateKeyDaysAgo_(withinDays);
  const limit = params.limit || 50;
  const offset = params.offset || 0;

  const filtered = data.rows.filter(r => {
    if (supplierCd && r.supplierCd !== supplierCd) return false;
    if (!r.lastSaleKey || r.lastSaleKey < cutoff) return false;
    return true;
  });

  filtered.sort((a, b) => (b.lastSaleKey || '').localeCompare(a.lastSaleKey || ''));

  const page = filtered.slice(offset, offset + limit);

  return {
    ok: true,
    total: filtered.length,
    candidates: page.map(r => ({
      code: r.code, name: r.name, kana: r.kana, ryakusho: r.ryakusho,
      unit: r.unit, supplierCd: r.supplierCd, supplierName: r.supplierName,
      price: r.price, kouri: r.kouri, shiire: r.shiire, jan: r.jan,
      lastSaleDate: dateKeyToSlash_(r.lastSaleKey),
      stockManagement: r.stockManagement
    }))
  };
}

// getRegisteredExamples — シリーズ推定用のfew-shot教師データ（既存の登録実績）
function getRegisteredExamples(params) {
  const supplierCd = params.supplierCd != null ? String(params.supplierCd).trim() : '';
  const maxProducts = params.maxProducts || 80;

  const csvData = loadCsvFromDrive();
  if (!csvData.ok) return csvData;
  const codeToSupplier = {};
  csvData.rows.forEach(row => {
    const code = row['コード'];
    if (!code) return;
    const key = String(parseInt(code, 10) || code);
    codeToSupplier[key] = (row['仕入先CD'] || '').trim();
  });

  const productsRes = bcartGetAll('/products');
  if (!productsRes.ok) return productsRes;
  Utilities.sleep(500);
  const setsRes = bcartGetAll('/product_sets');
  if (!setsRes.ok) return setsRes;

  const productMap = {};
  productsRes.data.forEach(p => { productMap[p.id] = p; });

  const setsByProduct = {};
  setsRes.data.forEach(s => {
    const pid = String(s.product_id);
    if (!setsByProduct[pid]) setsByProduct[pid] = [];
    setsByProduct[pid].push(s);
  });

  let matchedIds = [];
  if (supplierCd) {
    matchedIds = Object.keys(setsByProduct).filter(pid =>
      setsByProduct[pid].some(s => codeToSupplier[String(parseInt(s.product_no, 10) || s.product_no)] === supplierCd)
    );
  }

  let fallback = false;
  if (matchedIds.length === 0) {
    fallback = true;
    matchedIds = Object.keys(setsByProduct).filter(pid => setsByProduct[pid].length > 1);
  }

  matchedIds.sort((a, b) => setsByProduct[b].length - setsByProduct[a].length);
  matchedIds = matchedIds.slice(0, maxProducts);

  const examples = matchedIds.map(pid => {
    const p = productMap[pid] || {};
    return {
      productId: Number(pid),
      productName: p.name || '',
      categoryId: p.category_id || null,
      featureIds: [p.feature_id1 || null, p.feature_id2 || null, p.feature_id3 || null],
      sets: setsByProduct[pid].map(s => ({ setId: s.id, productNo: s.product_no, setName: s.name }))
    };
  });

  const categoriesRes = getCategories();

  return {
    ok: true,
    examples: examples,
    categories: categoriesRes.ok ? categoriesRes.categories : [],
    fallback: fallback
  };
}

// saveDrafts — ドラフト保存（登録ドラフト／登録ドラフト_セットシートへ書き込み。部分保存はしない）
function saveDrafts(params) {
  const drafts = params.drafts || [];
  if (drafts.length === 0) return { ok: false, error: '保存するドラフトがありません' };

  const parentSheet = getOrCreateSheet(SHEET_DRAFT);
  const setSheet = getOrCreateSheet(SHEET_DRAFT_SETS);

  const parentRows = parentSheet.getDataRange().getValues();
  const parentStatusById = {};
  for (let i = 1; i < parentRows.length; i++) {
    if (parentRows[i][0]) parentStatusById[parentRows[i][0]] = parentRows[i][1];
  }

  // 重複ガード: 既存の下書き/承認/登録済ドラフトに含まれるcodeを収集
  const existingSetRows = setSheet.getDataRange().getValues();
  const activeCodes = {};
  for (let i = 1; i < existingSetRows.length; i++) {
    const draftId = existingSetRows[i][0];
    const code = existingSetRows[i][1];
    if (!draftId || !code) continue;
    const status = parentStatusById[draftId];
    if (status === '下書き' || status === '承認' || status === '登録済') {
      activeCodes[String(code)] = draftId;
    }
  }

  // 本日分の既存連番の最大値を取得
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  let maxSeq = 0;
  for (let i = 1; i < parentRows.length; i++) {
    const id = String(parentRows[i][0] || '');
    if (id.indexOf('D' + todayStr + '-') === 0) {
      const seq = parseInt(id.split('-')[1], 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }

  const nowStr = new Date().toLocaleString('ja-JP');
  const savedIds = [];
  const skipped = [];
  const supplierNames = {};
  let savedCount = 0;

  drafts.forEach(d => {
    const sets = d.sets || [];
    if (sets.length === 0) {
      skipped.push({ code: '', reason: 'セットが指定されていません' });
      return;
    }
    const conflictCode = sets.map(s => s.code).find(c => activeCodes[String(c)]);
    if (conflictCode) {
      skipped.push({ code: conflictCode, reason: '既存ドラフト' + activeCodes[String(conflictCode)] + 'に含まれるためスキップ' });
      return;
    }

    maxSeq++;
    const draftId = 'D' + todayStr + '-' + ('00' + maxSeq).slice(-3);

    parentSheet.appendRow([
      draftId, '下書き', d.draftType || 'new_product', d.targetProductId || '',
      d.productName || '', d.categoryId || '',
      d.featureId1 || '', d.featureId2 || '', d.featureId3 || '',
      d.description || '', d.confidence || '', d.reasoning || '',
      (d.refUrls || []).join('\n'), d.supplierCd || '', d.supplierName || '',
      nowStr, '', ''
    ]);

    sets.forEach(s => {
      setSheet.appendRow([
        draftId, s.code || '', s.setName || '', s.jan || '',
        s.unitPrice || '', s.jodai || '', s.shiire || '', s.unit || '',
        s.lastSaleDate || ''
      ]);
      activeCodes[String(s.code)] = draftId;
    });

    savedIds.push(draftId);
    savedCount++;
    if (d.supplierName) supplierNames[d.supplierName] = true;
  });

  if (savedCount > 0) {
    const webhook = PropertiesService.getScriptProperties().getProperty('LINEWORKS_WEBHOOK');
    if (webhook) {
      const supplierLabel = Object.keys(supplierNames).join('、') || '不明';
      try {
        UrlFetchApp.fetch(webhook, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({ content: `【BCART登録ドラフト】${savedCount}件保存（仕入先: ${supplierLabel}）\nレビューをお願いします。` }),
          muteHttpExceptions: true
        });
      } catch (e) {
        Logger.log('LINE WORKS通知エラー: ' + e);
      }
    }
  }

  return { ok: true, saved: savedCount, skipped: skipped, draftIds: savedIds };
}

// ===================== 登録ドラフト（フェーズ2: レビュー・登録実行。すべてポータルセッション認証） =====================

// 登録ドラフト列インデックス(0始まり): 0 draft_id / 1 status / 2 draft_type / 3 target_product_id /
// 4 product_name / 5 category_id / 6-8 feature_id1-3 / 9 description / 10 confidence / 11 reasoning /
// 12 ref_urls / 13 supplier_cd / 14 supplier_name / 15 created_at / 16 reviewed_at / 17 registered_product_id /
// 18 jodai_type / 19 tax_type_id

function getDrafts(params) {
  const status = params.status !== undefined ? params.status : '下書き';
  const supplierCd = params.supplierCd != null ? String(params.supplierCd) : '';

  const parentRows = getOrCreateSheet(SHEET_DRAFT).getDataRange().getValues();
  const setRows = getOrCreateSheet(SHEET_DRAFT_SETS).getDataRange().getValues();

  const setsByDraft = {};
  for (let i = 1; i < setRows.length; i++) {
    const r = setRows[i];
    if (!r[0]) continue;
    if (!setsByDraft[r[0]]) setsByDraft[r[0]] = [];
    setsByDraft[r[0]].push({
      code: String(r[1]), setName: r[2], jan: String(r[3] || ''), unitPrice: r[4], jodai: r[5], shiire: r[6], unit: r[7], lastSaleDate: r[8]
    });
  }

  const drafts = [];
  for (let i = 1; i < parentRows.length; i++) {
    const r = parentRows[i];
    if (!r[0]) continue;
    if (status && r[1] !== status) continue;
    if (supplierCd && String(r[13]) !== supplierCd) continue;
    drafts.push({
      draftId: r[0], status: r[1], draftType: r[2], targetProductId: r[3] || null,
      productName: r[4], categoryId: r[5] || null,
      featureId1: r[6] || null, featureId2: r[7] || null, featureId3: r[8] || null,
      description: r[9], confidence: r[10], reasoning: r[11],
      refUrls: r[12] ? String(r[12]).split('\n').filter(Boolean) : [],
      supplierCd: r[13], supplierName: r[14],
      createdAt: r[15], reviewedAt: r[16] || '', registeredProductId: r[17] || null,
      jodaiType: r[18] || '', taxTypeId: r[19] || null,
      sets: setsByDraft[r[0]] || []
    });
  }
  drafts.sort((a, b) => b.draftId.localeCompare(a.draftId));

  return { ok: true, drafts: drafts };
}

function updateDraft(params) {
  const sheet = getOrCreateSheet(SHEET_DRAFT);
  const rows = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === params.draftId) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) return { ok: false, error: 'draftIdが見つかりません' };

  if (params.productName !== undefined) sheet.getRange(rowIdx, 5).setValue(params.productName);
  if (params.categoryId !== undefined) sheet.getRange(rowIdx, 6).setValue(params.categoryId);
  if (params.featureId1 !== undefined) sheet.getRange(rowIdx, 7).setValue(params.featureId1);
  if (params.featureId2 !== undefined) sheet.getRange(rowIdx, 8).setValue(params.featureId2);
  if (params.featureId3 !== undefined) sheet.getRange(rowIdx, 9).setValue(params.featureId3);
  if (params.description !== undefined) sheet.getRange(rowIdx, 10).setValue(params.description);
  if (params.jodaiType !== undefined) sheet.getRange(rowIdx, 19).setValue(params.jodaiType);
  if (params.taxTypeId !== undefined) sheet.getRange(rowIdx, 20).setValue(params.taxTypeId);

  if (params.sets) {
    const setSheet = getOrCreateSheet(SHEET_DRAFT_SETS);
    const setRows = setSheet.getDataRange().getValues();
    for (let i = setRows.length; i >= 2; i--) {
      if (setRows[i - 1][0] === params.draftId) setSheet.deleteRow(i);
    }
    params.sets.forEach(s => {
      setSheet.appendRow([params.draftId, String(s.code || ''), s.setName || '', String(s.jan || ''), s.unitPrice || '', s.jodai || '', s.shiire || '', s.unit || '', s.lastSaleDate || '']);
    });
  }

  addHistory({ userName: params._userName, code: '', name: params.draftId, type: '登録ドラフト編集', before: '', after: '', result: '成功' });
  return { ok: true };
}

// 承認＝登録実行。CSV最新値で価格をリフレッシュし、シートの現在値を正として非表示登録する
function approveDraft(params) {
  const parentSheet = getOrCreateSheet(SHEET_DRAFT);
  const parentRows = parentSheet.getDataRange().getValues();
  let rowIdx = -1, draft = null;
  for (let i = 1; i < parentRows.length; i++) {
    if (parentRows[i][0] === params.draftId) { rowIdx = i + 1; draft = parentRows[i]; break; }
  }
  if (rowIdx === -1) return { ok: false, error: 'draftIdが見つかりません' };
  if (draft[1] !== '下書き') return { ok: false, error: 'このドラフトは下書き状態ではありません（status: ' + draft[1] + '）' };

  const setRows = getOrCreateSheet(SHEET_DRAFT_SETS).getDataRange().getValues();
  const rawSets = [];
  for (let i = 1; i < setRows.length; i++) {
    if (setRows[i][0] === params.draftId) {
      // Sheetsが数値化した値をBCART API用に文字列へ正規化（product_no/jan_codeは文字列型必須。数値のまま送ると422になる）
      rawSets.push({
        code: String(setRows[i][1]), setName: String(setRows[i][2] || ''), janCode: String(setRows[i][3] || ''),
        csvPrice: Number(setRows[i][4]) || 0, csvKouri: Number(setRows[i][5]) || 0,
        csvShiire: Number(setRows[i][6]) || 0, csvUnit: String(setRows[i][7] || '')
      });
    }
  }
  if (rawSets.length === 0) return { ok: false, error: 'セットが登録されていません' };

  // CSV最新値でのリフレッシュ（ドラフト保存後に販売管理側で価格改定された場合の取りこぼし防止）
  const notes = [];
  let sets = rawSets;
  const csvData = loadCsvFromDrive();
  if (csvData.ok) {
    const csvMap = {};
    csvData.rows.forEach(row => {
      const code = row['コード'];
      if (!code) return;
      csvMap[String(parseInt(code, 10) || code)] = row;
    });
    sets = rawSets.map(s => {
      const row = csvMap[String(parseInt(s.code, 10) || s.code)];
      if (!row) {
        notes.push('code ' + s.code + ': CSVに見つからないためドラフト保存時の値のまま登録');
        return s;
      }
      const freshPrice  = parseFloat(String(row['売上単価'] || '').replace(/,/g, '')) || 0;
      const freshKouri  = parseFloat(String(row['定価１'] || row['定価1'] || '').replace(/,/g, '')) || 0;
      const freshShiire = parseFloat(String(row['仕入単価'] || '').replace(/,/g, '')) || 0;
      const freshJan    = (row['JANCD'] || '').trim();
      const freshUnit   = (row['単位名'] || '').trim();

      if (freshPrice && freshPrice !== s.csvPrice) notes.push('code ' + s.code + ': 売価 ' + s.csvPrice + '→' + freshPrice + '円（CSV最新値）');
      if (freshKouri !== s.csvKouri) notes.push('code ' + s.code + ': 上代 ' + s.csvKouri + '→' + freshKouri + '円（CSV最新値）');
      if (freshShiire && freshShiire !== s.csvShiire) notes.push('code ' + s.code + ': 仕入 ' + s.csvShiire + '→' + freshShiire + '円（CSV最新値）');

      return {
        code: s.code, setName: s.setName,
        janCode: freshJan || s.janCode,
        csvPrice: freshPrice || s.csvPrice,
        csvKouri: freshKouri,
        csvShiire: freshShiire || s.csvShiire,
        csvUnit: freshUnit || s.csvUnit
      };
    });
  } else {
    notes.push('CSV読み込み失敗のためドラフト保存時の値で登録: ' + csvData.error);
  }

  const draftType = draft[2];
  const productName = draft[4];
  const categoryId = draft[5];
  const featureId1 = draft[6] || null, featureId2 = draft[7] || null, featureId3 = draft[8] || null;
  const description = draft[9];
  const jodaiType = draft[18] || null;
  const taxTypeId = draft[19] || null;

  let productId, results;
  if (draftType === 'add_to_existing') {
    const targetProductId = draft[3];
    results = sets.map(s => {
      const res = addSetToProduct({
        _userName: params._userName, productId: targetProductId, code: s.code, setName: s.setName,
        janCode: s.janCode, csvPrice: s.csvPrice, csvKouri: s.csvKouri, csvShiire: s.csvShiire, csvUnit: s.csvUnit,
        jodaiType: jodaiType, taxTypeId: taxTypeId, setFlag: '非表示'
      });
      return { code: s.code, ok: res.ok, setId: res.setId || null, error: res.ok ? null : res.error };
    });
    productId = Number(targetProductId);
  } else {
    const bulkRes = bulkRegisterProduct({
      _userName: params._userName, productName: productName, categoryId: categoryId,
      featureId1: featureId1, featureId2: featureId2, featureId3: featureId3,
      productFlag: '非表示', setFlag: '非表示', jodaiType: jodaiType, taxTypeId: taxTypeId, sets: sets
    });
    if (!bulkRes.ok) return Object.assign(bulkRes, { notes: notes });  // 全セット失敗＝ロールバック済み。ドラフトは下書きのまま
    productId = bulkRes.productId;
    results = bulkRes.results;
  }

  const allOk = results.every(r => r.ok);
  if (!allOk) {
    addHistory({
      userName: params._userName, code: sets.map(s => s.code).join(','), name: productName || '',
      type: '登録ドラフト承認（一部失敗）', before: '', after: '商品ID: ' + productId, result: '一部失敗'
    });
    return {
      ok: true, partial: true, productId: productId, results: results, notes: notes,
      message: '一部のセットが登録に失敗しました。ステータスは下書きのままです。BCART管理画面で商品ID ' + productId + ' を確認してください。'
    };
  }

  // description反映（registerProduct/bulkRegisterProduct系はdescriptionを扱わないため追加PATCH）
  let descriptionApplied = true;
  if (description) {
    const descRes = bcartPatch('/products/' + productId, { description: description });
    descriptionApplied = descRes.ok;
    if (!descRes.ok) Logger.log('approveDraft: description PATCH失敗 draftId=' + params.draftId + ' error=' + descRes.error);
  }

  parentSheet.getRange(rowIdx, 2).setValue('登録済');
  parentSheet.getRange(rowIdx, 17).setValue(new Date().toLocaleString('ja-JP'));
  parentSheet.getRange(rowIdx, 18).setValue(productId);

  addHistory({
    userName: params._userName, code: sets.map(s => s.code).join(','), name: productName || '',
    type: '登録ドラフト承認・登録', before: '', after: '商品ID: ' + productId + (notes.length ? '（' + notes.join('；') + '）' : ''), result: '成功'
  });

  return { ok: true, productId: productId, results: results, notes: notes, descriptionApplied: descriptionApplied };
}

// 却下＝対応不要マークへ。以後の候補抽出から自動的に消える
function rejectDraft(params) {
  const parentSheet = getOrCreateSheet(SHEET_DRAFT);
  const rows = parentSheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === params.draftId) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) return { ok: false, error: 'draftIdが見つかりません' };

  const supplierName = rows[rowIdx - 1][14];
  parentSheet.getRange(rowIdx, 2).setValue('却下');
  parentSheet.getRange(rowIdx, 17).setValue(new Date().toLocaleString('ja-JP'));

  const setRows = getOrCreateSheet(SHEET_DRAFT_SETS).getDataRange().getValues();
  const codes = [];
  for (let i = 1; i < setRows.length; i++) {
    if (setRows[i][0] === params.draftId) codes.push({ code: String(setRows[i][1]), setName: setRows[i][2] });
  }
  codes.forEach(c => {
    markIgnore({ code: c.code, name: c.setName, reason: '登録ドラフト却下: ' + (params.reason || ''), supplier: supplierName });
  });

  addHistory({
    userName: params._userName, code: codes.map(c => c.code).join(','), name: params.draftId,
    type: '登録ドラフト却下', before: '', after: '', result: '成功'
  });
  return { ok: true };
}

// 公開＝表示化（登録済ドラフトのみ。画像アップ完了後にTakashiが押す想定）
function publishDraft(params) {
  const parentSheet = getOrCreateSheet(SHEET_DRAFT);
  const rows = parentSheet.getDataRange().getValues();
  let rowIdx = -1, draft = null;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === params.draftId) { rowIdx = i + 1; draft = rows[i]; break; }
  }
  if (rowIdx === -1) return { ok: false, error: 'draftIdが見つかりません' };
  if (draft[1] !== '登録済') return { ok: false, error: 'このドラフトは登録済み状態ではありません（status: ' + draft[1] + '）' };

  const draftType = draft[2];
  const registeredProductId = draft[17];
  if (!registeredProductId) return { ok: false, error: '登録済み商品IDが記録されていません' };

  const setRows = getOrCreateSheet(SHEET_DRAFT_SETS).getDataRange().getValues();
  const codes = [];
  for (let i = 1; i < setRows.length; i++) {
    if (setRows[i][0] === params.draftId) codes.push(String(setRows[i][1]));
  }

  const allSets = bcartGetAll('/product_sets');
  if (!allSets.ok) return allSets;
  const setIdByCode = {};
  allSets.data.forEach(s => { setIdByCode[String(s.product_no)] = s.id; });

  const results = codes.map(code => {
    const setId = setIdByCode[code];
    if (!setId) return { code: code, ok: false, error: 'BCART上にセットが見つかりません' };
    const res = bcartPatch('/product_sets/' + setId, { set_flag: '表示' });
    return { code: code, ok: res.ok, error: res.ok ? null : res.error };
  });

  if (draftType === 'new_product') {
    const productRes = bcartPatch('/products/' + registeredProductId, { flag: '表示' });
    if (!productRes.ok) return { ok: false, error: '親商品の表示化に失敗しました: ' + productRes.error, results: results };
  }

  const allOk = results.every(r => r.ok);
  parentSheet.getRange(rowIdx, 2).setValue('公開済');

  addHistory({
    userName: params._userName, code: codes.join(','), name: draft[4] || '',
    type: '登録ドラフト公開', before: '', after: '商品ID: ' + registeredProductId, result: allOk ? '成功' : '一部失敗'
  });

  return { ok: true, allOk: allOk, results: results };
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
        Logger.log('getMembers error: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
        return { ok: false, error: 'BCART_API_ERROR: ' + res.getResponseCode() };
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
    Logger.log('getMembers error: ' + e.message);
    return { ok: false, error: 'INTERNAL_ERROR' };
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

// 会員一覧キャッシュ（1リクエスト内で使い回す）
let _spCustomersCache = null;
function fetchCustomersCached() {
  if (_spCustomersCache) return _spCustomersCache;
  const res = getMembers();
  if (!res.ok) return res;
  const byExtId = {};
  const byId = {};
  res.members.forEach(m => {
    if (m.ext_id) byExtId[m.ext_id] = m;
    byId[m.id] = m;
  });
  _spCustomersCache = { ok: true, members: res.members, byExtId: byExtId, byId: byId };
  return _spCustomersCache;
}

// 得意先コード（会員のext_id）→ 会員ID解決。未登録は unresolved に残す
function resolveCustomerCodes(codes) {
  if (!codes || codes.length === 0) return { ok: true, memberIds: [], unresolved: [] };
  const c = fetchCustomersCached();
  if (!c.ok) return { ok: false, error: c.error, memberIds: [], unresolved: codes.slice() };
  const memberIds = [];
  const unresolved = [];
  codes.forEach(code => {
    const m = c.byExtId[code];
    if (m) memberIds.push(String(m.id));
    else unresolved.push(code);
  });
  return { ok: true, memberIds: memberIds, unresolved: unresolved };
}

// グループシートから該当行を探す（行の値配列を返す）
function findGroupRowValues(groupId) {
  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  for (let i = 1; i < groupRows.length; i++) {
    if (String(groupRows[i][0]) === String(groupId)) return groupRows[i];
  }
  return null;
}

// グループの実効会員 = 手入力member_ids ∪ 得意先コード解決分。未登録コードは unresolvedCodes
function getGroupEffectiveMembers(groupId) {
  const row = findGroupRowValues(groupId);
  if (!row) return { memberIds: [], unresolvedCodes: [], notFound: true };
  const manual = String(row[2] || '').split(',').map(s => s.trim()).filter(s => s);
  const codes  = String(row[6] || '').split(',').map(s => s.trim()).filter(s => s);
  const r = resolveCustomerCodes(codes);
  // 会員APIエラー時はコード分を保留扱いにして手入力分だけで動作継続
  const merged = [...new Set(manual.concat(r.memberIds))];
  return { memberIds: merged, unresolvedCodes: r.unresolved, apiError: r.ok === false };
}

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
        use_view_filter: groupRows[i][5] === true || groupRows[i][5] === 'TRUE',
        customer_codes: String(groupRows[i][6] || '')
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

  const indSheet = getOrCreateSheet(SHEET_SP_INDIVIDUAL);
  const indRows = indSheet.getDataRange().getValues();
  const individuals = [];
  for (let i = 1; i < indRows.length; i++) {
    if (indRows[i][0]) {
      individuals.push({
        detail_id:        String(indRows[i][0]),
        customer_code:    String(indRows[i][1]),
        member_id:        String(indRows[i][2] || ''),
        customer_name:    String(indRows[i][3] || ''),
        product_set_id:   indRows[i][4],
        product_no:       String(indRows[i][5]),
        product_set_name: String(indRows[i][6]),
        unit_price:       indRows[i][7],
        updated_at:       String(indRows[i][8] || ''),
        applied_at:       String(indRows[i][9] || ''),
        note:             String(indRows[i][10] || '')
      });
    }
  }

  return { ok: true, groups: groups, details: details, vfDetails: vfDetails, individuals: individuals };
}

// 既存シート（6列時代）にcustomer_codes列ヘッダーを補う
function ensureSpGroupCodesHeader(sheet) {
  if (String(sheet.getRange(1, 7).getValue() || '') !== 'customer_codes') {
    sheet.getRange(1, 7).setValue('customer_codes').setFontWeight('bold').setBackground('#f3f4f6');
  }
}

function saveCustomerGroup(params) {
  const sheet = getOrCreateSheet(SHEET_SP_GROUPS);
  ensureSpGroupCodesHeader(sheet);
  const rows = sheet.getDataRange().getValues();
  const memberIds = params.member_ids || '';
  // customer_codes は渡されたときだけ更新（会員選択モーダル等からの保存で消さない）
  const hasCodes = params.customer_codes !== undefined && params.customer_codes !== null;

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
        if (hasCodes) {
          const cCell = sheet.getRange(i + 1, 7);
          cCell.setNumberFormat('@');
          cCell.setValue(String(params.customer_codes));
        }
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
  if (hasCodes) {
    const cCell = sheet.getRange(lastRow, 7);
    cCell.setNumberFormat('@');
    cCell.setValue(String(params.customer_codes));
  }
  return { ok: true, group_id: newId };
}

function deleteCustomerGroup(params) {
  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  let groupRowIdx = -1;
  for (let i = 1; i < groupRows.length; i++) {
    if (String(groupRows[i][0]) === params.group_id) {
      groupRowIdx = i + 1;
      break;
    }
  }
  const memberIds = getGroupEffectiveMembers(params.group_id).memberIds;

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
  const eff = getGroupEffectiveMembers(params.group_id);
  const memberIds = eff.memberIds;
  if (memberIds.length === 0) {
    if (eff.unresolvedCodes.length > 0) {
      return { ok: false, error: '有効な会員がいません（未登録の得意先コード: ' + eff.unresolvedCodes.join(', ') + '）' };
    }
    return { ok: false, error: 'グループの会員ID・得意先コードが設定されていません' };
  }

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
    after: '成功: ' + successCount + '件 / 失敗: ' + failCount + '件' + (eff.unresolvedCodes.length > 0 ? ' / 未登録保留: ' + eff.unresolvedCodes.length + '件' : ''),
    result: failCount > 0 ? '一部失敗' : '成功'
  });

  return { ok: true, successCount: successCount, failCount: failCount, errors: errors, memberCount: memberIds.length, unresolvedCodes: eff.unresolvedCodes };
}

function deleteSpecialPriceDetail(params) {
  const memberIds = getGroupEffectiveMembers(params.group_id).memberIds;

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
  const eff = getGroupEffectiveMembers(params.group_id);
  const memberIds = eff.memberIds;
  if (memberIds.length === 0) {
    if (eff.unresolvedCodes.length > 0) {
      return { ok: false, error: '有効な会員がいません（未登録の得意先コード: ' + eff.unresolvedCodes.join(', ') + '）' };
    }
    return { ok: false, error: 'グループの会員ID・得意先コードが設定されていません' };
  }

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
    before: '', after: '成功: ' + successCount + '件 / 失敗: ' + failCount + '件' + (eff.unresolvedCodes.length > 0 ? ' / 未登録保留: ' + eff.unresolvedCodes.length + '件' : ''),
    result: failCount > 0 ? '一部失敗' : '成功'
  });
  return { ok: true, successCount: successCount, failCount: failCount, errors: errors, memberCount: memberIds.length, unresolvedCodes: eff.unresolvedCodes };
}

function deleteViewFilterDetail(params) {
  const memberIds = getGroupEffectiveMembers(params.group_id).memberIds;

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

// ===================== 個別特価（1得意先×1商品） =====================

// 品番の正規化（数値なら先頭0を除去して統一。CSVコード⇔product_no照合用）
function normalizeProductNo(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  const n = parseInt(s, 10);
  return (!isNaN(n) && /^\d+$/.test(s)) ? String(n) : s.toLowerCase();
}

// 一括保存（貼り付けインポート兼用）。items: [{customer_code, product_no, unit_price, note}]
// 会員未登録の得意先コードは member_id 空欄＝保留として保持し、反映時に再解決する
function saveIndividualPrices(params) {
  const items = params.items || [];
  if (items.length === 0) return { ok: false, error: '登録データがありません' };
  if (items.length > 500) return { ok: false, error: '一度に登録できるのは500件までです' };

  const allSets = bcartGetAll('/product_sets');
  if (!allSets.ok) return allSets;
  const setByNo = {};
  allSets.data.forEach(s => {
    const key = normalizeProductNo(s.product_no);
    if (key && !setByNo[key]) setByNo[key] = s;
  });

  const customers = fetchCustomersCached();  // 取得失敗でも保留として登録は続行

  const sheet = getOrCreateSheet(SHEET_SP_INDIVIDUAL);
  const rows = sheet.getDataRange().getValues();
  const rowKey = {};  // customer_code|product_set_id → 行番号
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) rowKey[String(rows[i][1]) + '|' + String(rows[i][4])] = i + 1;
  }

  let added = 0, updated = 0;
  const unknownProducts = [];
  const pendingCodes = [];
  const nowStr = new Date().toLocaleString('ja-JP');

  items.forEach(item => {
    const code = String(item.customer_code || '').trim();
    const pno = normalizeProductNo(item.product_no);
    const price = Number(item.unit_price);
    if (!code || !pno || !(price > 0)) return;
    const set = setByNo[pno];
    if (!set) { unknownProducts.push(String(item.product_no)); return; }

    let memberId = '', memberName = '';
    if (customers.ok) {
      const m = customers.byExtId[code];
      if (m) { memberId = String(m.id); memberName = m.comp_name || m.name || ''; }
    }
    if (!memberId) pendingCodes.push(code);

    const key = code + '|' + String(set.id);
    if (rowKey[key]) {
      const r = rowKey[key];
      sheet.getRange(r, 3).setValue(memberId);
      if (memberName) sheet.getRange(r, 4).setValue(memberName);
      sheet.getRange(r, 8).setValue(price);
      sheet.getRange(r, 9).setValue(nowStr);
      sheet.getRange(r, 10).setValue('');  // 価格変更＝未反映に戻す
      if (item.note !== undefined && item.note !== null) sheet.getRange(r, 11).setValue(String(item.note));
      updated++;
    } else {
      const newId = 'I' + new Date().getTime().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 6);
      sheet.appendRow([newId, code, memberId, memberName, set.id, String(set.product_no || ''), set.name || '', price, nowStr, '', String(item.note || '')]);
      rowKey[key] = sheet.getLastRow();
      added++;
    }
  });

  addHistory({
    userName: params._userName, code: '', name: '個別特価',
    type: '個別特価アプリ保存',
    before: '',
    after: '新規' + added + '件 / 更新' + updated + '件' + (pendingCodes.length > 0 ? ' / 未登録保留' + [...new Set(pendingCodes)].length + '件' : ''),
    result: unknownProducts.length > 0 ? '一部スキップ' : '成功'
  });

  return {
    ok: true, added: added, updated: updated,
    pendingCodes: [...new Set(pendingCodes)],
    unknownProducts: [...new Set(unknownProducts)]
  };
}

// BCART反映。product_set単位でPATCHをまとめ、保留行（会員未登録）は反映前に再解決を試みる
function applyIndividualPrices(params) {
  const sheet = getOrCreateSheet(SHEET_SP_INDIVIDUAL);
  const rows = sheet.getDataRange().getValues();
  const onlyIds = params.detail_ids && params.detail_ids.length > 0 ? params.detail_ids.map(String) : null;

  const customers = fetchCustomersCached();

  const targets = [];  // {rowIdx, memberId, setId, price}
  let resolvedNow = 0;
  const pending = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (onlyIds && onlyIds.indexOf(String(rows[i][0])) === -1) continue;
    const code = String(rows[i][1] || '');
    let memberId = String(rows[i][2] || '');
    if (!memberId && customers.ok) {
      const m = customers.byExtId[code];
      if (m) {
        memberId = String(m.id);
        sheet.getRange(i + 1, 3).setValue(memberId);
        sheet.getRange(i + 1, 4).setValue(m.comp_name || m.name || '');
        resolvedNow++;
      }
    }
    if (!memberId) { pending.push(code); continue; }
    targets.push({ rowIdx: i + 1, memberId: memberId, setId: String(rows[i][4]), price: Number(rows[i][7]) });
  }

  if (targets.length === 0) {
    return { ok: true, successCount: 0, failCount: 0, resolvedNow: resolvedNow, skippedPending: [...new Set(pending)], errors: [] };
  }

  const allSets = bcartGetAll('/product_sets');
  if (!allSets.ok) return allSets;
  const setMap = {};
  allSets.data.forEach(s => { setMap[String(s.id)] = s; });

  const bySet = {};
  targets.forEach(t => { (bySet[t.setId] = bySet[t.setId] || []).push(t); });

  let successCount = 0, failCount = 0;
  const errors = [];
  const nowStr = new Date().toLocaleString('ja-JP');

  Object.keys(bySet).forEach(setId => {
    const list = bySet[setId];
    const bcartSet = setMap[setId];
    if (!bcartSet) {
      failCount += list.length;
      errors.push('setId ' + setId + ': 商品セットが見つかりません');
      return;
    }
    const newSp = Object.assign({}, bcartSet.special_price || {});
    list.forEach(t => {
      // カンマ区切りキーに対象会員が含まれる場合は個別キーに分解してから上書き（他会員の価格は維持）
      Object.keys(newSp).forEach(key => {
        if (key.indexOf(',') !== -1) {
          const ids = key.split(',').map(s => s.trim()).filter(s => s);
          if (ids.indexOf(t.memberId) !== -1) {
            const val = newSp[key];
            delete newSp[key];
            ids.forEach(id => { if (id !== t.memberId) newSp[id] = val; });
          }
        }
      });
      newSp[t.memberId] = { unit_price: t.price };
    });
    const res = bcartPatch('/product_sets/' + setId, { special_price: newSp });
    if (res.ok) {
      successCount += list.length;
      list.forEach(t => { sheet.getRange(t.rowIdx, 10).setValue(nowStr); });
    } else {
      failCount += list.length;
      errors.push('setId ' + setId + ': ' + res.error);
    }
    Utilities.sleep(150);
  });

  addHistory({
    userName: params._userName, code: '', name: '個別特価',
    type: '個別特価適用',
    before: '',
    after: '成功: ' + successCount + '件 / 失敗: ' + failCount + '件' + (pending.length > 0 ? ' / 未登録保留: ' + [...new Set(pending)].length + '件' : ''),
    result: failCount > 0 ? '一部失敗' : '成功'
  });

  return {
    ok: true, successCount: successCount, failCount: failCount,
    resolvedNow: resolvedNow, skippedPending: [...new Set(pending)], errors: errors
  };
}

function deleteIndividualPrice(params) {
  const sheet = getOrCreateSheet(SHEET_SP_INDIVIDUAL);
  const rows = sheet.getDataRange().getValues();
  let rowIdx = -1, row = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(params.detail_id)) { rowIdx = i + 1; row = rows[i]; break; }
  }
  if (rowIdx < 0) return { ok: false, error: '対象が見つかりません' };

  const memberId = String(row[2] || '');
  const setId = String(row[4] || '');
  const applied = String(row[9] || '') !== '';
  if (memberId && setId && applied) {
    const cur = bcartGet('/product_sets/' + setId);
    if (cur.ok && cur.data) {
      const newSp = Object.assign({}, cur.data.special_price || {});
      let changed = false;
      Object.keys(newSp).forEach(key => {
        const ids = key.split(',').map(s => s.trim()).filter(s => s);
        if (ids.indexOf(memberId) !== -1) {
          const val = newSp[key];
          delete newSp[key];
          ids.forEach(id => { if (id !== memberId) newSp[id] = val; });
          changed = true;
        }
      });
      if (changed) bcartPatch('/product_sets/' + setId, { special_price: newSp });
    }
  }
  sheet.deleteRow(rowIdx);

  addHistory({
    userName: params._userName, code: String(row[1] || ''), name: String(row[5] || '') + ' ' + String(row[6] || ''),
    type: '個別特価削除', before: '', after: '', result: '成功'
  });
  return { ok: true };
}

// ===================== 特価突合チェック =====================
// BCART実態（special_price）とアプリ内マスター（グループ明細＋個別特価）を突合する
function auditSpecialPrices(params) {
  const allSets = bcartGetAll('/product_sets');
  if (!allSets.ok) return allSets;
  const customers = fetchCustomersCached();
  const memberName = function(id) {
    return (customers.ok && customers.byId[id]) ? (customers.byId[id].comp_name || customers.byId[id].name || '') : '';
  };
  const memberCode = function(id) {
    return (customers.ok && customers.byId[id]) ? customers.byId[id].ext_id : '';
  };

  // 期待値マップ expected[setId][memberId] = {price, source}
  const expected = {};
  const pendingCodes = {};  // 未登録得意先コード → 出所

  const groupSheet = getOrCreateSheet(SHEET_SP_GROUPS);
  const groupRows = groupSheet.getDataRange().getValues();
  const groupInfo = {};
  for (let i = 1; i < groupRows.length; i++) {
    if (!groupRows[i][0]) continue;
    const gid = String(groupRows[i][0]);
    const manual = String(groupRows[i][2] || '').split(',').map(s => s.trim()).filter(s => s);
    const codes = String(groupRows[i][6] || '').split(',').map(s => s.trim()).filter(s => s);
    const r = resolveCustomerCodes(codes);
    r.unresolved.forEach(code => { pendingCodes[code] = 'グループ: ' + String(groupRows[i][1]); });
    groupInfo[gid] = { name: String(groupRows[i][1]), memberIds: [...new Set(manual.concat(r.memberIds))] };
  }

  const detailSheet = getOrCreateSheet(SHEET_SP_DETAILS);
  const detailRows = detailSheet.getDataRange().getValues();
  for (let i = 1; i < detailRows.length; i++) {
    if (!detailRows[i][0]) continue;
    const g = groupInfo[String(detailRows[i][1])];
    if (!g) continue;
    const setId = String(detailRows[i][2]);
    const price = Number(detailRows[i][5]);
    if (!expected[setId]) expected[setId] = {};
    g.memberIds.forEach(mid => { expected[setId][mid] = { price: price, source: 'グループ: ' + g.name }; });
  }

  const indSheet = getOrCreateSheet(SHEET_SP_INDIVIDUAL);
  const indRows = indSheet.getDataRange().getValues();
  for (let i = 1; i < indRows.length; i++) {
    if (!indRows[i][0]) continue;
    const mid = String(indRows[i][2] || '');
    const code = String(indRows[i][1] || '');
    if (!mid) { if (!pendingCodes[code]) pendingCodes[code] = '個別特価'; continue; }
    const setId = String(indRows[i][4]);
    if (!expected[setId]) expected[setId] = {};
    expected[setId][mid] = { price: Number(indRows[i][7]), source: '個別特価' };
  }

  // BCART実態と突合
  const unmanaged = [], mismatch = [], missing = [];
  let bcartEntries = 0, matched = 0;
  const LIMIT = 300;
  allSets.data.forEach(s => {
    const setId = String(s.id);
    const sp = s.special_price || {};
    const actual = {};
    Object.keys(sp).forEach(key => {
      const price = (sp[key] && sp[key].unit_price !== undefined) ? Number(sp[key].unit_price) : NaN;
      key.split(',').map(x => x.trim()).filter(x => x).forEach(mid => { actual[mid] = price; });
    });
    const exp = expected[setId] || {};
    Object.keys(actual).forEach(mid => {
      bcartEntries++;
      if (exp[mid] === undefined) {
        if (unmanaged.length < LIMIT) unmanaged.push({ product_no: String(s.product_no || ''), set_name: s.name || '', member_id: mid, member_name: memberName(mid), customer_code: memberCode(mid), bcart_price: actual[mid] });
      } else if (Number(exp[mid].price) !== Number(actual[mid])) {
        if (mismatch.length < LIMIT) mismatch.push({ product_no: String(s.product_no || ''), set_name: s.name || '', member_id: mid, member_name: memberName(mid), customer_code: memberCode(mid), bcart_price: actual[mid], expected_price: exp[mid].price, source: exp[mid].source });
      } else {
        matched++;
      }
    });
    Object.keys(exp).forEach(mid => {
      if (actual[mid] === undefined) {
        if (missing.length < LIMIT) missing.push({ product_no: String(s.product_no || ''), set_name: s.name || '', member_id: mid, member_name: memberName(mid), customer_code: memberCode(mid), expected_price: exp[mid].price, source: exp[mid].source });
      }
    });
    delete expected[setId];
  });
  // アプリ内マスターにあるがBCARTに商品セット自体が存在しない
  Object.keys(expected).forEach(setId => {
    Object.keys(expected[setId]).forEach(mid => {
      if (missing.length < LIMIT) missing.push({ product_no: '(セットID:' + setId + ')', set_name: '商品セットがBCARTに存在しません', member_id: mid, member_name: memberName(mid), customer_code: memberCode(mid), expected_price: expected[setId][mid].price, source: expected[setId][mid].source });
    });
  });

  const pending = Object.keys(pendingCodes).map(code => ({ customer_code: code, source: pendingCodes[code] }));

  return {
    ok: true,
    summary: {
      sets: allSets.data.length,
      bcartEntries: bcartEntries,
      matched: matched,
      unmanaged: unmanaged.length,
      mismatch: mismatch.length,
      missing: missing.length,
      pending: pending.length
    },
    unmanaged: unmanaged, mismatch: mismatch, missing: missing, pending: pending,
    truncated: unmanaged.length >= LIMIT || mismatch.length >= LIMIT || missing.length >= LIMIT
  };
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
    } else if (sheetName === SHEET_DESC_SKIP) {
      sheet.appendRow(['商品ID', '商品名', '登録日時']);
    } else if (sheetName === SHEET_HISTORY) {
      sheet.appendRow(['日時', '操作者', '商品コード', '商品名', '操作種別', '変更前', '変更後', '結果']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_SP_GROUPS) {
      sheet.appendRow(['group_id', 'group_name', 'member_ids', 'created_at', 'note', 'use_view_filter', 'customer_codes']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_SP_INDIVIDUAL) {
      sheet.appendRow(['detail_id', 'customer_code', 'member_id', 'customer_name', 'product_set_id', 'product_no', 'product_set_name', 'unit_price', 'updated_at', 'applied_at', 'note']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#f3f4f6');
      sheet.getRange('B2:C').setNumberFormat('@');  // 得意先コード・会員IDを文字列扱い（287-1等の日付誤変換防止）
    } else if (sheetName === SHEET_VF_DETAILS) {
      sheet.appendRow(['detail_id', 'group_id', 'product_set_id', 'product_no', 'product_set_name', 'applied_at']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_SP_DETAILS) {
      // ③ applied_at列を含む8列構成
      sheet.appendRow(['detail_id', 'group_id', 'product_set_id', 'product_no', 'product_set_name', 'unit_price', 'updated_at', 'applied_at']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_FEATURES) {
      sheet.appendRow(['feature_id', 'type', 'updated_at']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_DRAFT) {
      sheet.appendRow(['draft_id', 'status', 'draft_type', 'target_product_id', 'product_name', 'category_id', 'feature_id1', 'feature_id2', 'feature_id3', 'description', 'confidence', 'reasoning', 'ref_urls', 'supplier_cd', 'supplier_name', 'created_at', 'reviewed_at', 'registered_product_id', 'jodai_type', 'tax_type_id']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 20).setFontWeight('bold').setBackground('#f3f4f6');
    } else if (sheetName === SHEET_DRAFT_SETS) {
      sheet.appendRow(['draft_id', 'code', 'set_name', 'jan', 'unit_price', 'jodai', 'shiire', 'unit', 'last_sale_date']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f3f4f6');
      sheet.getRange('B2:B').setNumberFormat('@');  // code
      sheet.getRange('D2:D').setNumberFormat('@');  // jan（先頭ゼロ・大桁数の誤変換防止）
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
    } else if (sheetName === SHEET_DRAFT) {
      // jodai_type / tax_type_id 列の自動追加（既存シートへのマイグレーション）
      if (sheet.getLastColumn() < 20) {
        sheet.getRange(1, 19, 1, 2).setValues([['jodai_type', 'tax_type_id']]);
        sheet.getRange(1, 19, 1, 2).setFontWeight('bold').setBackground('#f3f4f6');
      }
    } else if (sheetName === SHEET_DRAFT_SETS) {
      // code/jan列をテキスト書式に統一（既存シートへのマイグレーション。数値化された既存値はapproveDraft側のString()正規化で吸収）
      sheet.getRange('B2:B').setNumberFormat('@');
      sheet.getRange('D2:D').setNumberFormat('@');
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

// ===================== 機能C-補助: 説明文不要フラグ =====================
function getDescSkipMap() {
  const sheet = getOrCreateSheet(SHEET_DESC_SKIP);
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) map[String(rows[i][0])] = true;
  }
  return map;
}

function markDescSkip(params) {
  const sheet = getOrCreateSheet(SHEET_DESC_SKIP);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(params.id)) return { ok: true };
  }
  sheet.appendRow([params.id, params.name, new Date().toLocaleString('ja-JP')]);
  return { ok: true };
}

function unmarkDescSkip(params) {
  const sheet = getOrCreateSheet(SHEET_DESC_SKIP);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(params.id)) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: true };
}

// ===================== 機能C: 説明文生成（Gemini API） =====================

function getProductsForDescription(params) {
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

  const skipMap = getDescSkipMap();

  // 説明文不要リストのみ表示モード
  if (params && params.skipOnly) {
    const skipList = allMapped.filter(p => skipMap[String(p.id)]).sort((a, b) => b.id - a.id);
    return { ok: true, products: skipList, skipOnly: true, total: allMapped.length };
  }

  // 通常モード: 説明文なし かつ 不要フラグなし
  const noDetail  = allMapped.filter(p => (!p.detail || p.detail.trim() === '') && !skipMap[String(p.id)]).sort((a, b) => b.id - a.id);
  const hasDetail = allMapped.filter(p => p.detail && p.detail.trim() !== '');
  const skipCount = Object.keys(skipMap).length;

  return { ok: true, products: noDetail, withDetail: hasDetail.length, total: allMapped.length, skipCount: skipCount };
}

function getSimilarProducts(params) {
  const name      = params && params.name      ? String(params.name)      : '';
  const excludeId = params && params.excludeId ? String(params.excludeId) : '';
  if (!name) return { ok: true, products: [] };

  const baseName = _getBaseNameGas(name);
  if (!baseName || baseName.length < 4) return { ok: true, products: [] };

  const res = bcartGetAll('/products');
  if (!res.ok) return res;

  const matches = res.data
    .filter(p => {
      if (String(p.id) === excludeId) return false;
      if (!p.description || !p.description.trim()) return false;
      return _getBaseNameGas(p.name || '').indexOf(baseName) >= 0;
    })
    .map(p => ({ id: p.id, name: p.name, detail: p.description }));

  return { ok: true, products: matches };
}

function _getBaseNameGas(name) {
  return (name || '')
    .replace(/[\d]+(\.\d+)?\s*(g|ml|mL|L|ℓ|kg|cc|本|枚|個|set|セット|pack|パック|step|ステップ|oz)/gi, '')
    .replace(/[\d]+/g, '')
    .replace(/\b(mini|max|super|pro|lite|light|plus|compact|large|extra|jumbo|ex)\b/gi, '')
    .replace(/(ミニ|マックス|スーパー|プロ|ライト|プラス|コンパクト|ラージ|エクストラ|ジャンボ)/g, '')
    .replace(/\s+/g, ' ')
    .trim();
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
    Logger.log('generateDescription error: ' + e.message);
    return { ok: false, error: 'INTERNAL_ERROR' };
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
    Logger.log('factCheckDescription error: ' + e.message);
    return { ok: false, error: 'INTERNAL_ERROR' };
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

// ===================== 機能D: 特集管理 =====================

function getFeatureList() {
  const res = bcartGetAll('/product_features');
  if (!res.ok) return res;
  const raw = res.data;
  const list = (Array.isArray(raw) ? raw : []).map(f => ({
    id: f.id,
    name: f.name || '',
    rv_description: f.rv_description || '',
    priority: f.priority !== undefined ? Number(f.priority) : 0,
    flag: f.flag !== undefined ? Number(f.flag) : 1,
    type: ''
  }));

  const typeMap = getFeatureTypeMap_();
  list.forEach(f => { f.type = typeMap[String(f.id)] || ''; });
  list.sort((a, b) => a.priority - b.priority);
  return { ok: true, features: list };
}

function createFeature(params) {
  if (!params.name) return { ok: false, error: '特集名は必須です' };
  const payload = {
    product_features: [{
      name: params.name,
      priority: Number(params.priority || 0),
      flag: Number(params.flag !== undefined ? params.flag : 1),
      rv_description: params.rv_description || ''
    }]
  };
  const res = bcartPost('/product_features', payload);
  if (!res.ok) return res;

  // 作成された特集IDを取得して種類を保存
  try {
    const created = res.data && res.data.product_features;
    const newId = created && created[0] && created[0].id;
    if (newId && params.type) saveFeatureTypeInternal_(newId, params.type);
  } catch (e) {
    Logger.log('createFeature: 種類保存エラー ' + e.message);
  }

  addHistory({ userName: params._userName, code: '', name: params.name, type: '特集作成', before: '', after: params.name, result: '成功' });
  return { ok: true };
}

function updateFeature(params) {
  if (!params.id) return { ok: false, error: 'IDが指定されていません' };
  const payload = {
    name: params.name,
    priority: Number(params.priority || 0),
    flag: Number(params.flag !== undefined ? params.flag : 1),
    rv_description: params.rv_description || ''
  };
  const res = bcartPatch('/product_features/' + params.id, payload);
  if (!res.ok) return res;

  if (params.type !== undefined) saveFeatureTypeInternal_(params.id, params.type);
  addHistory({ userName: params._userName, code: '', name: params.name, type: '特集更新', before: '', after: params.name, result: '成功' });
  return { ok: true };
}

function bulkUpdateFeatureOrder(params) {
  if (!params.features || !params.features.length) return { ok: false, error: 'featuresが空です' };
  const payload = {
    product_features: params.features.map(f => ({ id: Number(f.id), priority: Number(f.priority) }))
  };
  const res = bcartPatch('/product_features/', payload);
  if (!res.ok) return res;
  addHistory({ userName: params._userName, code: '', name: '', type: '特集順序変更', before: '', after: params.features.length + '件', result: '成功' });
  return { ok: true };
}

function saveFeatureType(params) {
  if (!params.featureId) return { ok: false, error: 'featureIdが指定されていません' };
  saveFeatureTypeInternal_(params.featureId, params.type || '');
  return { ok: true };
}

function bulkSaveFeatureTypes(params) {
  if (!params.items || !params.items.length) return { ok: false, error: 'itemsが空です' };
  params.items.forEach(item => {
    if (item.featureId) saveFeatureTypeInternal_(item.featureId, item.type || '');
  });
  return { ok: true };
}

// ---- 内部ヘルパー ----

function getFeatureTypeMap_() {
  const sheet = getOrCreateSheet(SHEET_FEATURES);
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== '') map[String(data[i][0])] = data[i][1];
  }
  return map;
}

function saveFeatureTypeInternal_(featureId, type) {
  const sheet = getOrCreateSheet(SHEET_FEATURES);
  const data = sheet.getDataRange().getValues();
  const now = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(featureId)) {
      sheet.getRange(i + 1, 2).setValue(type);
      sheet.getRange(i + 1, 3).setValue(now);
      return;
    }
  }
  sheet.appendRow([featureId, type, now]);
}

// ===================== 機能E: Claudeチャット直接操作 =====================

function fetchProductsAndSets_(productIds) {
  const idSet = new Set(productIds.map(id => String(id)));
  const products = bcartGetAll('/products');
  if (!products.ok) return products;
  const sets = bcartGetAll('/product_sets');
  if (!sets.ok) return sets;
  const targetProducts = products.data.filter(p => idSet.has(String(p.id)));
  const targetSets = sets.data.filter(s => idSet.has(String(s.product_id)));
  const notFound = productIds.filter(id => !targetProducts.some(p => String(p.id) === String(id)));
  return { ok: true, products: targetProducts, sets: targetSets, notFound };
}

// ---- 操作1: 名前末尾テキスト追加 ----

function previewSuffixName(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };
  if (!params.suffix)
    return { ok: false, error: 'suffix が未指定です' };

  const d = fetchProductsAndSets_(params.productIds);
  if (!d.ok) return d;

  const suffix = params.suffix;
  const preview = d.products.map(p => {
    const currentName = p.name || '';
    const skip = currentName.endsWith(suffix);
    const pSets = d.sets
      .filter(s => String(s.product_id) === String(p.id))
      .map(s => {
        const cur = s.name || '';
        const sSkip = cur.endsWith(suffix);
        return { setId: s.id, currentName: cur, newName: sSkip ? cur : cur + suffix, skip: sSkip };
      });
    return {
      productId: p.id,
      productNo: p.product_no || p.main_no || '',
      currentName,
      newName: skip ? currentName : currentName + suffix,
      skip,
      sets: pSets
    };
  });

  const totalChanges = preview.reduce(
    (sum, p) => sum + (p.skip ? 0 : 1) + p.sets.reduce((s2, s) => s2 + (s.skip ? 0 : 1), 0), 0
  );
  return { ok: true, preview, totalChanges, notFound: d.notFound, suffix };
}

function applySuffixName(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };
  if (!params.suffix)
    return { ok: false, error: 'suffix が未指定です' };

  const d = fetchProductsAndSets_(params.productIds);
  if (!d.ok) return d;

  const suffix = params.suffix;
  const userName = params._userName || 'Claude(API)';
  const results = [];

  // 商品名（products）一括PATCH
  const productUpdates = d.products
    .filter(p => !(p.name || '').endsWith(suffix))
    .map(p => ({ id: p.id, name: (p.name || '') + suffix }));
  if (productUpdates.length > 0) {
    const res = bcartPatch('/products/', { products: productUpdates });
    results.push({ target: 'products', count: productUpdates.length, ok: res.ok, error: res.error || '' });
    if (res.ok) {
      productUpdates.forEach(u => {
        const p = d.products.find(p2 => p2.id === u.id);
        addHistory({ userName, code: p ? (p.product_no || p.main_no || '') : '', name: p ? (p.name || '') : '', type: '商品名末尾追加', before: p ? (p.name || '') : '', after: u.name, result: '成功' });
      });
    }
  }

  // セット名（product_sets）個別PATCH（unit_price を同時送信で API 要件を満たす）
  let setOk = 0;
  let setFail = 0;
  for (const s of d.sets) {
    if ((s.name || '').endsWith(suffix)) continue;
    const newName = (s.name || '') + suffix;
    const res = bcartPatch('/product_sets/' + s.id, { name: newName, unit_price: s.unit_price });
    if (res.ok) {
      setOk++;
      const p = d.products.find(p2 => String(p2.id) === String(s.product_id));
      addHistory({ userName, code: p ? (p.product_no || p.main_no || '') : '', name: s.name || '', type: 'セット名末尾追加', before: s.name || '', after: newName, result: '成功' });
    } else {
      setFail++;
      Logger.log('applySuffixName set PATCH failed: id=' + s.id + ' ' + res.error);
    }
    Utilities.sleep(300);
  }
  if (setOk > 0 || setFail > 0)
    results.push({ target: 'product_sets', success: setOk, failed: setFail, ok: setFail === 0 });

  return { ok: results.every(r => r.ok !== false), results, suffix };
}

// ---- 操作2: 販売終了日時の設定 ----

function previewHanbaiEnd(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };

  const hanbaiEnd = Object.prototype.hasOwnProperty.call(params, 'hanbaiEnd') ? params.hanbaiEnd : null;

  const all = bcartGetAll('/products');
  if (!all.ok) return all;

  const idSet = new Set(params.productIds.map(id => String(id)));
  const targets = all.data.filter(p => idSet.has(String(p.id)));
  const notFound = params.productIds.filter(id => !targets.some(p => String(p.id) === String(id)));

  const preview = targets.map(p => ({
    productId: p.id,
    productNo: p.product_no || p.main_no || '',
    name: p.name || '',
    currentHanbaiEnd: p.hanbai_end || null,
    newHanbaiEnd: hanbaiEnd
  }));

  return { ok: true, preview, totalChanges: preview.length, notFound, hanbaiEnd };
}

function applyHanbaiEnd(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };

  const hanbaiEnd = Object.prototype.hasOwnProperty.call(params, 'hanbaiEnd') ? params.hanbaiEnd : null;

  const all = bcartGetAll('/products');
  if (!all.ok) return all;

  const idSet = new Set(params.productIds.map(id => String(id)));
  const targets = all.data.filter(p => idSet.has(String(p.id)));
  const userName = params._userName || 'Claude(API)';

  if (targets.length === 0) return { ok: true, count: 0, message: '対象商品が見つかりませんでした' };

  const updates = targets.map(p => ({ id: p.id, hanbai_end: hanbaiEnd }));
  const res = bcartPatch('/products/', { products: updates });

  if (res.ok) {
    targets.forEach(p => {
      addHistory({ userName, code: p.product_no || p.main_no || '', name: p.name || '', type: '販売終了日時設定', before: p.hanbai_end || '（未設定）', after: hanbaiEnd || '（解除）', result: '成功' });
    });
  }

  return { ok: res.ok, count: updates.length, error: res.error || '', hanbaiEnd };
}

// ---- 操作3: 商品セット説明欄 末尾追加 ----

function previewSetDescription(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };
  if (!params.appendText)
    return { ok: false, error: 'appendText が未指定です' };

  const d = fetchProductsAndSets_(params.productIds);
  if (!d.ok) return d;

  const appendText = params.appendText;
  const preview = d.products.map(p => {
    const pSets = d.sets
      .filter(s => String(s.product_id) === String(p.id))
      .map(s => {
        const cur = s.description || '';
        const skip = cur.includes(appendText);
        return {
          setId: s.id,
          setName: s.name || '',
          currentDescriptionPreview: cur.slice(0, 100) + (cur.length > 100 ? '…' : ''),
          skip
        };
      });
    return { productId: p.id, productNo: p.product_no || p.main_no || '', name: p.name || '', sets: pSets };
  });

  const totalChanges = preview.reduce((sum, p) => sum + p.sets.reduce((s2, s) => s2 + (s.skip ? 0 : 1), 0), 0);
  return { ok: true, preview, totalChanges, notFound: d.notFound, appendText };
}

function applySetDescription(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };
  if (!params.appendText)
    return { ok: false, error: 'appendText が未指定です' };

  const d = fetchProductsAndSets_(params.productIds);
  if (!d.ok) return d;

  const appendText = params.appendText;
  const userName = params._userName || 'Claude(API)';
  let setOk = 0;
  let setFail = 0;

  for (const s of d.sets) {
    const cur = s.description || '';
    if (cur.includes(appendText)) continue;
    const newDesc = cur + appendText;
    const res = bcartPatch('/product_sets/' + s.id, { description: newDesc });
    if (res.ok) {
      setOk++;
      const p = d.products.find(p2 => String(p2.id) === String(s.product_id));
      addHistory({ userName, code: p ? (p.product_no || p.main_no || '') : '', name: s.name || '', type: 'セット説明末尾追加', before: cur.slice(0, 50) + (cur.length > 50 ? '…' : ''), after: '末尾追加: ' + appendText.slice(0, 30) + (appendText.length > 30 ? '…' : ''), result: '成功' });
    } else {
      setFail++;
      Logger.log('applySetDescription PATCH failed: id=' + s.id + ' ' + res.error);
    }
    Utilities.sleep(300);
  }

  return { ok: setFail === 0, success: setOk, failed: setFail };
}

// ---- 操作4: 商品フィールド一括設定（特集3・商品特徴・画像） ----

function previewProductFields(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };

  const d = fetchProductsAndSets_(params.productIds);
  if (!d.ok) return d;

  const featureId3   = Object.prototype.hasOwnProperty.call(params, 'feature_id3') ? params.feature_id3 : null;
  const tag          = Object.prototype.hasOwnProperty.call(params, 'tag')          ? params.tag          : null;
  const imageBasePath = Object.prototype.hasOwnProperty.call(params, 'imageBasePath') ? params.imageBasePath : null;

  const preview = d.products.map(p => {
    const item = { productId: p.id, productNo: p.product_no || p.main_no || '', productName: p.name || '' };
    if (featureId3 !== null)   item.feature_id3    = { before: p.feature_id3 || null, after: featureId3 };
    if (tag !== null)          item.tag             = { before: p.tag || null, after: tag };
    if (imageBasePath !== null) item.image          = { before: p.image || null, after: imageBasePath + p.id + '.png' };
    return item;
  });

  return { ok: true, preview, totalChanges: d.products.length, notFound: d.notFound };
}

function applyProductFields(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };

  const d = fetchProductsAndSets_(params.productIds);
  if (!d.ok) return d;

  const featureId3    = Object.prototype.hasOwnProperty.call(params, 'feature_id3')   ? params.feature_id3   : undefined;
  const tag           = Object.prototype.hasOwnProperty.call(params, 'tag')            ? params.tag            : undefined;
  const imageBasePath = Object.prototype.hasOwnProperty.call(params, 'imageBasePath') ? params.imageBasePath : undefined;
  const userName      = params._userName || 'Claude(API)';

  const productUpdates = d.products.map(p => {
    const update = { id: p.id };
    if (featureId3    !== undefined) update.feature_id3 = featureId3;
    if (tag           !== undefined) update.tag         = tag;
    if (imageBasePath !== undefined) update.image       = imageBasePath + p.id + '.png';
    return update;
  });

  const res = bcartPatch('/products/', { products: productUpdates });
  if (res.ok) {
    d.products.forEach(p => {
      const u = productUpdates.find(x => x.id === p.id);
      addHistory({ userName, code: p.product_no || p.main_no || '', name: p.name || '', type: '商品フィールド更新(Claude)', before: '', after: JSON.stringify(u), result: '成功' });
    });
  }
  return { ok: res.ok, count: productUpdates.length, error: res.error || '' };
}

// ---- 操作5: セットフィールド一括設定（上代タイプ・入数・状態） ----

function previewSetFields(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };

  const d = fetchProductsAndSets_(params.productIds);
  if (!d.ok) return d;

  const jodaiType = Object.prototype.hasOwnProperty.call(params, 'jodai_type') ? params.jodai_type : null;
  const quantity  = Object.prototype.hasOwnProperty.call(params, 'quantity')   ? params.quantity   : null;
  const setFlag   = Object.prototype.hasOwnProperty.call(params, 'set_flag')   ? params.set_flag   : null;

  const preview = d.sets.map(s => {
    const item = { setId: s.id, productId: s.product_id, setName: s.name || '' };
    if (jodaiType !== null) item.jodai_type = { before: s.jodai_type || null, after: jodaiType };
    if (quantity  !== null) item.quantity   = { before: s.quantity   || null, after: quantity };
    if (setFlag   !== null) item.set_flag   = { before: s.set_flag   || null, after: setFlag };
    return item;
  });

  return { ok: true, preview, totalChanges: d.sets.length, notFound: d.notFound };
}

function applySetFields(params) {
  if (!params.productIds || !Array.isArray(params.productIds) || params.productIds.length === 0)
    return { ok: false, error: 'productIds が未指定です' };

  const d = fetchProductsAndSets_(params.productIds);
  if (!d.ok) return d;

  const jodaiType = Object.prototype.hasOwnProperty.call(params, 'jodai_type') ? params.jodai_type : undefined;
  const quantity  = Object.prototype.hasOwnProperty.call(params, 'quantity')   ? params.quantity   : undefined;
  const setFlag   = Object.prototype.hasOwnProperty.call(params, 'set_flag')   ? params.set_flag   : undefined;
  const userName  = params._userName || 'Claude(API)';

  let success = 0;
  let failed  = 0;

  for (const s of d.sets) {
    const body = {};
    if (jodaiType !== undefined) body.jodai_type = jodaiType;
    if (quantity  !== undefined) body.quantity   = quantity;
    if (setFlag   !== undefined) body.set_flag   = setFlag;
    if (Object.keys(body).length === 0) continue;

    const res = bcartPatch('/product_sets/' + s.id, body);
    const p = d.products.find(p2 => String(p2.id) === String(s.product_id));
    if (res.ok) {
      success++;
      addHistory({ userName, code: p ? (p.product_no || p.main_no || '') : '', name: s.name || '', type: 'セットフィールド更新(Claude)', before: '', after: JSON.stringify(body), result: '成功' });
    } else {
      failed++;
      Logger.log('applySetFields PATCH failed: setId=' + s.id + ' ' + res.error);
    }
    Utilities.sleep(300);
  }

  return { ok: failed === 0, success, failed };
}

// ---- 操作6: 商品の表示順（表示優先度）一括設定 ----

function previewProductSort(params) {
  if (!params.items || !Array.isArray(params.items) || params.items.length === 0)
    return { ok: false, error: 'items が未指定です（[{ id, sort }] の配列）' };

  const productIds = params.items.map(it => it.id);
  const d = fetchProductsAndSets_(productIds);
  if (!d.ok) return d;

  // 表示順フィールドの候補（実データに存在するものだけ現在値を拾い、フィールド名を特定する）
  const candidates = ['sort', 'priority', 'disp_no', 'sort_no', 'sort_order', 'display_order',
                      'display_priority', 'disp_order', 'view_priority', 'order', 'order_no'];

  const sortMap = {};
  params.items.forEach(it => { sortMap[String(it.id)] = it.sort; });

  const preview = d.products.map(p => {
    const cand = {};
    candidates.forEach(k => { if (Object.prototype.hasOwnProperty.call(p, k)) cand[k] = p[k]; });
    return {
      productId: p.id,
      productNo: p.product_no || p.main_no || '',
      productName: p.name || '',
      candidateFields: cand,        // 候補フィールドの現在値（ここから表示順フィールド名を特定）
      after: sortMap[String(p.id)]  // 設定したい表示順
    };
  });

  // 先頭商品の全フィールド名（候補に該当が無い場合の特定用）
  const allKeys = d.products.length > 0 ? Object.keys(d.products[0]) : [];

  return { ok: true, preview, allKeys, totalChanges: d.products.length, notFound: d.notFound };
}

function applyProductSort(params) {
  if (!params.items || !Array.isArray(params.items) || params.items.length === 0)
    return { ok: false, error: 'items が未指定です（[{ id, sort }] の配列）' };
  if (!params.sortField)
    return { ok: false, error: 'sortField が未指定です（previewで特定したフィールド名を渡してください）' };

  const productIds = params.items.map(it => it.id);
  const d = fetchProductsAndSets_(productIds);
  if (!d.ok) return d;

  const sortField = params.sortField;
  const sortMap = {};
  params.items.forEach(it => { sortMap[String(it.id)] = it.sort; });
  const userName = params._userName || 'Claude(API)';

  const productUpdates = d.products
    .filter(p => sortMap[String(p.id)] !== undefined)
    .map(p => {
      const u = { id: p.id };
      u[sortField] = sortMap[String(p.id)];
      return u;
    });

  if (productUpdates.length === 0)
    return { ok: false, error: '対象商品が見つかりません', notFound: d.notFound };

  const res = bcartPatch('/products/', { products: productUpdates });
  if (res.ok) {
    d.products.forEach(p => {
      if (sortMap[String(p.id)] === undefined) return;
      addHistory({
        userName,
        code: p.product_no || p.main_no || '',
        name: p.name || '',
        type: '表示順更新(Claude)',
        before: '',
        after: sortField + '=' + sortMap[String(p.id)],
        result: '成功'
      });
    });
  }
  return { ok: res.ok, count: productUpdates.length, error: res.error || '', notFound: d.notFound };
}
