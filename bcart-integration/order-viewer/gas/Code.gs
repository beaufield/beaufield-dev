/**
 * Bカート受注確認 - 参照専用バックエンド
 *
 * 必須スクリプトプロパティ:
 *   AUTH_GAS_URL            BeaufieldポータルGASのWebアプリURL
 *   BCART_TOKEN             既存BカートAPIトークン
 *   BCART_SORT_CONFIRMED_AT Bカートの受注商品並び順を確認した日（YYYY-MM-DD）
 *
 * 安全方針:
 *   - BカートAPIはGETだけを使用する
 *   - 許可するパスは /orders と /order_products だけ
 *   - ブラウザからURL、パス、HTTPメソッドを受け取らない
 *   - 受注データをSheets、Drive、Properties、Cacheへ保存しない
 */

const VERSION = 'v0.1.2';
const APP_NAME = 'bcart-orders';
// ロール語彙統一（2026-08-21〜）: user_app_roles の値を admin/user/none の3語に揃える移行中。
// 旧語彙（viewer）が残っているシート行も従来どおり通すためのエイリアス。
// 移行後もシート直編集で旧語彙が復活した場合の安全網として残す。
const ROLE_ALIAS = { viewer: 'user' };
const BCART_BASE_URL = 'https://api.bcart.jp/api/v1';
const BCART_ADMIN_ORDER_BASE = 'https://beaufieldec.i16.bcart.jp/admin/order';
const SESSION_CACHE_SECONDS = 300;
const MAX_ORDER_LIST_LIMIT = 100;
const MAX_ORDER_PRODUCT_ROWS = 1000;
const ORDER_PRODUCT_PAGE_SIZE = 100;
const MAX_SEARCH_DAYS = 90;

const BCART_ALLOWED_PATHS = Object.freeze({
  '/orders': true,
  '/order_products': true
});

const ORDER_FIELDS = [
  'id',
  'code',
  'customer_comp_name',
  'ordered_at',
  'status',
  'final_price'
].join(',');

const ORDER_PRODUCT_FIELDS = [
  'id',
  'order_id',
  'product_id',
  'main_no',
  'product_no',
  'jan_code',
  'product_name',
  'product_set_id',
  'set_name',
  'unit_price',
  'set_quantity',
  'set_unit',
  'order_pro_count',
  'tax_rate',
  'tax_incl',
  'item_type'
].join(',');

// ============================================================
// エントリーポイント
// ============================================================

function doGet(e) {
  const action = e && e.parameter ? String(e.parameter.action || '') : '';
  if (action === 'version' || action === '') {
    return jsonResponse_({ success: true, version: VERSION, app_name: APP_NAME });
  }
  return jsonResponse_({ success: false, error: 'USE_POST', message: 'データ取得はPOSTを使用してください。' });
}

function doPost(e) {
  const requestId = Utilities.getUuid();
  const startedAt = Date.now();
  let params = {};

  try {
    params = parsePostBody_(e);
    const action = String(params.action || '');

    if (action === 'version') {
      return jsonResponse_({ success: true, version: VERSION, app_name: APP_NAME, request_id: requestId });
    }

    const auth = authorizeRequest_(String(params.session_token || ''));
    if (!auth.ok) {
      logRequest_(requestId, action, auth.user_id || '', '', startedAt, auth.error || 'AUTH_FAILED');
      return jsonResponse_({
        success: false,
        error: auth.error,
        message: auth.message,
        request_id: requestId
      });
    }

    let result;
    switch (action) {
      case 'health':
        result = getHealth_();
        break;
      case 'listOrders':
        result = listOrders_(params);
        break;
      case 'getOrderDetail':
        result = getOrderDetail_(params);
        break;
      default:
        return jsonResponse_({
          success: false,
          error: 'INVALID_REQUEST',
          message: '不明なアクションです。',
          request_id: requestId
        });
    }

    result.success = true;
    result.version = VERSION;
    result.request_id = requestId;
    result.fetched_at = nowIsoJapan_();
    logRequest_(requestId, action, auth.user_id, params.order_id || '', startedAt, 'OK');
    return jsonResponse_(result);
  } catch (error) {
    const safeError = normalizeError_(error);
    logRequest_(requestId, String(params.action || ''), '', params.order_id || '', startedAt, safeError.error);
    return jsonResponse_({
      success: false,
      error: safeError.error,
      message: safeError.message,
      request_id: requestId
    });
  }
}

function parsePostBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw appError_('INVALID_REQUEST', 'リクエスト本文がありません。');
  }

  try {
    const parsed = JSON.parse(e.postData.contents);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error('object required');
    }
    return parsed;
  } catch (error) {
    throw appError_('INVALID_REQUEST', 'JSON形式のリクエスト本文が必要です。');
  }
}

// ============================================================
// Beaufield-auth認証・アプリ権限
// ============================================================

function authorizeRequest_(token) {
  if (!token) {
    return {
      ok: false,
      error: 'SESSION_INVALID',
      message: '認証が必要です。ポータルからログインし直してください。'
    };
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'auth_' + token.slice(-32);
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      // 壊れたキャッシュは無視して再検証する。
    }
  }

  const appAuth = portalValidateSession_(token, APP_NAME);
  if (appAuth.ok) {
    const rawRole = String(appAuth.role || '').toLowerCase();
    const role = ROLE_ALIAS[rawRole] || rawRole;
    if (role !== 'user' && role !== 'admin') {
      const deniedRole = {
        ok: false,
        error: 'APP_ACCESS_DENIED',
        message: 'このアプリの閲覧権限がありません。',
        user_id: String(appAuth.user_id || '')
      };
      cache.put(cacheKey, JSON.stringify(deniedRole), 60);
      return deniedRole;
    }

    const allowed = {
      ok: true,
      user_id: String(appAuth.user_id || ''),
      name: String(appAuth.name || ''),
      role: role,
      is_admin: Boolean(appAuth.is_admin)
    };
    cache.put(cacheKey, JSON.stringify(allowed), SESSION_CACHE_SECONDS);
    return allowed;
  }

  // app_nameなしでは有効なら、セッションではなくアプリ権限が不足している。
  const baseAuth = portalValidateSession_(token, '');
  if (baseAuth.ok) {
    const denied = {
      ok: false,
      error: 'APP_ACCESS_DENIED',
      message: 'このアプリの閲覧権限がありません。',
      user_id: String(baseAuth.user_id || '')
    };
    cache.put(cacheKey, JSON.stringify(denied), 60);
    return denied;
  }

  const invalid = {
    ok: false,
    error: 'SESSION_INVALID',
    message: '認証が必要です。ポータルからログインし直してください。'
  };
  cache.put(cacheKey, JSON.stringify(invalid), 30);
  return invalid;
}

function portalValidateSession_(token, appName) {
  const authUrl = requiredProperty_('AUTH_GAS_URL');
  const payload = {
    action: 'validateSession',
    token: token
  };
  if (appName) payload.app_name = appName;

  let lastError = null;
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const response = UrlFetchApp.fetch(authUrl, {
        method: 'post',
        contentType: 'text/plain;charset=utf-8',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        followRedirects: true
      });
      const status = response.getResponseCode();
      const body = response.getContentText();
      if (status === 200) {
        const data = JSON.parse(body);
        if (data && typeof data.ok === 'boolean') return data;
      }
      lastError = new Error('AUTH_HTTP_' + status);
    } catch (error) {
      lastError = error;
    }

    if (attempt === 0) Utilities.sleep(500);
  }

  Logger.log('認証サービス確認失敗: ' + String(lastError && lastError.message || 'unknown'));
  throw appError_('AUTH_SERVICE_ERROR', '認証サービスへ接続できませんでした。少し待って再実行してください。');
}

// ============================================================
// 受注一覧・詳細
// ============================================================

function listOrders_(params) {
  const filters = params.filters && typeof params.filters === 'object' ? params.filters : {};
  const paging = params.paging && typeof params.paging === 'object' ? params.paging : {};
  const validated = validateOrderFilters_(filters, paging);

  const query = {
    limit: validated.limit,
    offset: validated.offset,
    fields: ORDER_FIELDS,
    ordered_at_min: validated.orderedAtMin,
    ordered_at_max: validated.orderedAtMax
  };
  if (validated.code) query.code = validated.code;
  if (validated.status) query.status = validated.status;

  const data = bcartGet_('/orders', query);
  if (!Array.isArray(data.orders) || !data.meta || typeof data.meta !== 'object') {
    throw appError_('RESPONSE_INVALID', 'Bカートの受注一覧レスポンスが不完全です。');
  }

  const orders = data.orders.map(sanitizeOrder_).sort(compareOrdersNewestFirst_);
  const total = toNonNegativeInteger_(data.meta.total, orders.length);

  return {
    orders: orders,
    paging: {
      limit: validated.limit,
      offset: validated.offset,
      returned: orders.length,
      total: total,
      has_next: validated.offset + orders.length < total
    }
  };
}

function getOrderDetail_(params) {
  const orderId = toPositiveInteger_(params.order_id, '受注ID');

  const orderData = bcartGet_('/orders', {
    ids: orderId,
    limit: 1,
    offset: 0,
    fields: ORDER_FIELDS
  });
  if (!Array.isArray(orderData.orders)) {
    throw appError_('RESPONSE_INVALID', 'Bカートの受注レスポンスが不完全です。');
  }
  if (orderData.orders.length === 0) {
    throw appError_('ORDER_NOT_FOUND', '指定した受注が見つかりません。');
  }

  const apiRows = fetchAllOrderProducts_(orderId);
  return buildOrderDetail_(sanitizeOrder_(orderData.orders[0]), apiRows);
}

function fetchAllOrderProducts_(orderId) {
  const rows = [];

  for (let offset = 0; offset < MAX_ORDER_PRODUCT_ROWS; offset += ORDER_PRODUCT_PAGE_SIZE) {
    const data = bcartGet_('/order_products', {
      order_id: orderId,
      limit: ORDER_PRODUCT_PAGE_SIZE,
      offset: offset,
      fields: ORDER_PRODUCT_FIELDS
    });
    if (!Array.isArray(data.order_products) || !data.meta || typeof data.meta !== 'object') {
      throw appError_('RESPONSE_INVALID', 'Bカートの受注明細レスポンスが不完全です。');
    }

    data.order_products.forEach(function (row) {
      rows.push(sanitizeOrderProduct_(row));
    });

    const total = toNonNegativeInteger_(data.meta.total, rows.length);
    if (rows.length >= total || data.order_products.length < ORDER_PRODUCT_PAGE_SIZE) {
      return rows;
    }
  }

  throw appError_('TOO_MANY_ROWS', '受注明細が安全上限の1,000行以上あるため表示を停止しました。');
}

function buildOrderDetail_(order, apiRows) {
  const sourceRows = apiRows.map(function (row, index) {
    const copy = clonePlainObject_(row);
    copy.source_index = index + 1;
    copy.line_amount = calculateLineAmount_(copy);
    return copy;
  });

  let productSourceIndex = 0;
  const products = [];
  const nonProducts = [];

  sourceRows.forEach(function (row) {
    if (row.item_type === 'product') {
      productSourceIndex++;
      row.source_product_index = productSourceIndex;
      products.push(row);
    } else {
      nonProducts.push(row);
    }
  });

  products.sort(compareOrderProductsForBcartScreen_);
  let movedCount = 0;
  products.forEach(function (row, index) {
    row.display_index = index + 1;
    row.moved = row.display_index !== row.source_product_index;
    if (row.moved) movedCount++;
  });
  nonProducts.sort(function (a, b) { return Number(a.id) - Number(b.id); });

  const confirmedAt = optionalProperty_('BCART_SORT_CONFIRMED_AT');
  const confirmation = evaluateSortConfirmation_(confirmedAt);

  return {
    order: order,
    sort: {
      mode: 'product_id_desc',
      label: 'Bカート商品ID降順',
      rule: 'product_id DESC, order_product.id ASC, null product_id last',
      confirmed_at: confirmation.confirmedAt,
      confirmation_status: confirmation.status,
      confirmation_message: confirmation.message,
      api_row_count: sourceRows.length,
      product_row_count: products.length,
      non_product_row_count: nonProducts.length,
      moved_product_count: movedCount
    },
    products: products,
    non_products: nonProducts,
    bcart_admin_url: BCART_ADMIN_ORDER_BASE + '/' + order.id + '/view'
  };
}

/**
 * 現在のBカート管理画面設定「Bカート商品ID降順」の再現規則。
 * product_set_idは比較に使用しない。
 */
function compareOrderProductsForBcartScreen_(a, b) {
  const aMissing = isMissingProductId_(a.product_id);
  const bMissing = isMissingProductId_(b.product_id);

  if (aMissing !== bMissing) return aMissing ? 1 : -1;

  if (!aMissing) {
    const productIdDiff = Number(b.product_id) - Number(a.product_id);
    if (productIdDiff !== 0) return productIdDiff;
  }

  return Number(a.id) - Number(b.id);
}

// ============================================================
// BカートAPI参照専用クライアント
// この区間では method: 'get' 以外を追加しない。
// ============================================================

function bcartGet_(path, query) {
  if (!BCART_ALLOWED_PATHS[path]) {
    throw appError_('INVALID_REQUEST', '許可されていないBカートAPIパスです。');
  }

  const token = requiredProperty_('BCART_TOKEN');
  const url = BCART_BASE_URL + path + '?' + buildQueryString_(query);
  let lastStatus = 0;

  for (let attempt = 0; attempt < 2; attempt++) {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Accept: 'application/json',
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true,
      followRedirects: true
    });
    const status = response.getResponseCode();
    lastStatus = status;

    if (status === 200) {
      const text = response.getContentText();
      try {
        const parsed = JSON.parse(text);
        if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
          throw new Error('object required');
        }
        return parsed;
      } catch (error) {
        throw appError_('RESPONSE_INVALID', 'BカートAPIからJSON以外の応答が返りました。');
      }
    }

    if (status === 401 || status === 403) {
      throw appError_('BCART_AUTH_ERROR', 'BカートAPIトークンまたは権限を確認してください。');
    }
    if (status === 429) {
      throw appError_('BCART_RATE_LIMIT', 'BカートAPIの利用上限に達しました。少し待って再実行してください。');
    }
    if (status < 500 || attempt === 1) break;
    Utilities.sleep(500);
  }

  throw appError_('BCART_API_ERROR', 'BカートAPIの取得に失敗しました（HTTP ' + lastStatus + '）。');
}

function buildQueryString_(query) {
  return Object.keys(query).filter(function (key) {
    return query[key] !== '' && query[key] !== null && query[key] !== undefined;
  }).map(function (key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(String(query[key]));
  }).join('&');
}

// ============================================================
// 入力検証・整形
// ============================================================

function validateOrderFilters_(filters, paging) {
  const orderedAtMin = validateDateTime_(filters.ordered_at_min, '開始日時');
  const orderedAtMax = validateDateTime_(filters.ordered_at_max, '終了日時');
  const minDate = new Date(orderedAtMin.replace(' ', 'T') + '+09:00');
  const maxDate = new Date(orderedAtMax.replace(' ', 'T') + '+09:00');
  if (minDate.getTime() > maxDate.getTime()) {
    throw appError_('INVALID_REQUEST', '開始日時は終了日時以前にしてください。');
  }
  if ((maxDate.getTime() - minDate.getTime()) / 86400000 > MAX_SEARCH_DAYS) {
    throw appError_('INVALID_REQUEST', '検索期間は90日以内にしてください。');
  }

  const limit = toIntegerInRange_(paging.limit === undefined ? 50 : paging.limit, 1, MAX_ORDER_LIST_LIMIT, '表示件数');
  const offset = toIntegerInRange_(paging.offset === undefined ? 0 : paging.offset, 0, 1000000, '開始位置');
  const code = validateShortText_(filters.code, 255, '受注番号');
  const status = validateShortText_(filters.status, 255, '対応状況');

  return {
    orderedAtMin: orderedAtMin,
    orderedAtMax: orderedAtMax,
    limit: limit,
    offset: offset,
    code: code,
    status: status
  };
}

function validateDateTime_(value, label) {
  const text = String(value || '');
  if (!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(text)) {
    throw appError_('INVALID_REQUEST', label + 'は YYYY-MM-DD HH:mm:ss 形式で指定してください。');
  }
  const date = new Date(text.replace(' ', 'T') + '+09:00');
  if (isNaN(date.getTime())) {
    throw appError_('INVALID_REQUEST', label + 'が正しい日時ではありません。');
  }
  const normalized = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  if (normalized !== text) {
    throw appError_('INVALID_REQUEST', label + 'が正しい日時ではありません。');
  }
  return text;
}

function validateShortText_(value, maxLength, label) {
  const text = String(value === undefined || value === null ? '' : value).trim();
  if (text.length > maxLength) {
    throw appError_('INVALID_REQUEST', label + 'が長すぎます。');
  }
  return text;
}

function toPositiveInteger_(value, label) {
  return toIntegerInRange_(value, 1, 2147483647, label);
}

function toIntegerInRange_(value, min, max, label) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < min || number > max) {
    throw appError_('INVALID_REQUEST', label + 'が不正です。');
  }
  return number;
}

function toNonNegativeInteger_(value, fallback) {
  const number = Number(value);
  return Number.isInteger(number) && number >= 0 ? number : fallback;
}

function sanitizeOrder_(row) {
  return {
    id: toApiPositiveInteger_(row.id, '受注ID'),
    code: String(row.code === undefined || row.code === null ? '' : row.code),
    customer_comp_name: String(row.customer_comp_name || ''),
    ordered_at: String(row.ordered_at || ''),
    status: String(row.status || ''),
    final_price: toFiniteNumber_(row.final_price, 0)
  };
}

function sanitizeOrderProduct_(row) {
  return {
    id: toApiPositiveInteger_(row.id, '受注商品ID'),
    order_id: toApiPositiveInteger_(row.order_id, '受注ID'),
    product_id: normalizeNullableInteger_(row.product_id),
    main_no: String(row.main_no || ''),
    product_no: String(row.product_no || ''),
    jan_code: String(row.jan_code || ''),
    product_name: String(row.product_name || ''),
    product_set_id: normalizeNullableInteger_(row.product_set_id),
    set_name: String(row.set_name || ''),
    unit_price: toFiniteNumber_(row.unit_price, 0),
    set_quantity: toFiniteNumber_(row.set_quantity, 0),
    set_unit: String(row.set_unit || ''),
    order_pro_count: toFiniteNumber_(row.order_pro_count, 0),
    tax_rate: String(row.tax_rate === undefined || row.tax_rate === null ? '' : row.tax_rate),
    tax_incl: toFiniteNumber_(row.tax_incl, 0),
    item_type: String(row.item_type || 'unknown')
  };
}

function normalizeNullableInteger_(value) {
  if (value === null || value === undefined || value === '') return null;
  const number = Number(value);
  return Number.isInteger(number) ? number : null;
}

function toApiPositiveInteger_(value, label) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < 1 || number > 2147483647) {
    throw appError_('RESPONSE_INVALID', 'BカートAPIの' + label + 'が不正です。');
  }
  return number;
}

function toFiniteNumber_(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function isMissingProductId_(value) {
  return value === null || value === undefined || value === '';
}

function calculateLineAmount_(row) {
  const unitPrice = toFiniteNumber_(row.unit_price, 0);
  const orderCount = toFiniteNumber_(row.order_pro_count, 0);
  if (row.item_type === 'product') {
    const setQuantity = toFiniteNumber_(row.set_quantity, 0);
    return unitPrice * setQuantity * orderCount;
  }
  return unitPrice * (orderCount || 1);
}

function compareOrdersNewestFirst_(a, b) {
  const dateDiff = String(b.ordered_at).localeCompare(String(a.ordered_at));
  return dateDiff !== 0 ? dateDiff : Number(b.id) - Number(a.id);
}

function clonePlainObject_(value) {
  return JSON.parse(JSON.stringify(value));
}

function evaluateSortConfirmation_(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))) {
    return {
      confirmedAt: '',
      status: 'missing',
      message: 'Bカートの受注商品並び順設定が未確認です。'
    };
  }

  const confirmedDate = new Date(String(value) + 'T00:00:00+09:00');
  const ageDays = Math.floor((Date.now() - confirmedDate.getTime()) / 86400000);
  if (isNaN(confirmedDate.getTime()) || ageDays < 0) {
    return {
      confirmedAt: String(value),
      status: 'invalid',
      message: '並び順設定の確認日が不正です。'
    };
  }
  if (ageDays > 90) {
    return {
      confirmedAt: String(value),
      status: 'stale',
      message: 'Bカートの並び順設定を90日以上確認していません。'
    };
  }
  return {
    confirmedAt: String(value),
    status: 'current',
    message: 'Bカートの並び順設定を確認済みです。'
  };
}

// ============================================================
// ヘルス・共通処理
// ============================================================

function getHealth_() {
  return {
    health: {
      auth_url_configured: Boolean(optionalProperty_('AUTH_GAS_URL')),
      bcart_token_configured: Boolean(optionalProperty_('BCART_TOKEN')),
      sort_confirmed_at: optionalProperty_('BCART_SORT_CONFIRMED_AT'),
      bcart_allowed_paths: Object.keys(BCART_ALLOWED_PATHS)
    }
  };
}

function requiredProperty_(key) {
  const value = optionalProperty_(key);
  if (!value) throw appError_('CONFIG_ERROR', 'スクリプトプロパティ ' + key + ' が未設定です。');
  return value;
}

function optionalProperty_(key) {
  return String(PropertiesService.getScriptProperties().getProperty(key) || '').trim();
}

function nowIsoJapan_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function appError_(code, message) {
  const error = new Error(message);
  error.appCode = code;
  return error;
}

function normalizeError_(error) {
  const knownCodes = [
    'AUTH_SERVICE_ERROR',
    'INVALID_REQUEST',
    'ORDER_NOT_FOUND',
    'BCART_AUTH_ERROR',
    'BCART_RATE_LIMIT',
    'BCART_API_ERROR',
    'RESPONSE_INVALID',
    'TOO_MANY_ROWS',
    'CONFIG_ERROR'
  ];
  const code = error && knownCodes.indexOf(error.appCode) !== -1 ? error.appCode : 'INTERNAL_ERROR';
  const message = code === 'INTERNAL_ERROR'
    ? '予期しないエラーが発生しました。request_idを管理者へ連絡してください。'
    : String(error.message || 'エラーが発生しました。');
  return { error: code, message: message };
}

function jsonResponse_(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

function logRequest_(requestId, action, userId, orderId, startedAt, result) {
  Logger.log(JSON.stringify({
    request_id: requestId,
    action: String(action || ''),
    user_id: String(userId || ''),
    order_id: String(orderId || ''),
    duration_ms: Date.now() - startedAt,
    result: String(result || '')
  }));
}
