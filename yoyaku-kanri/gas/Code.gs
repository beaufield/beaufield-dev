// ============================================================
// Beaufield 予約管理アプリ - Google Apps Script
// Version: 1.1.0
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   SPREADSHEET_ID  : 予約管理データのスプレッドシートID
//   AUTH_SHEET_ID   : beaufield-auth スプレッドシートID（共通）
//
// ============================================================

const VERSION  = '1.1.0';
const APP_NAME = 'yoyaku-kanri';

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS         = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = _PROPS.getProperty('SPREADSHEET_ID');
const AUTH_SHEET_ID  = _PROPS.getProperty('AUTH_SHEET_ID');

// シート名定数
const SHEET_PRODUCTS     = 'products';
const SHEET_RESERVATIONS = 'reservations';

// ============================================================
// 起動時チェック（プロパティ未設定を早期検知）
// ============================================================
function _checkProps() {
  if (!SPREADSHEET_ID) throw new Error('スクリプトプロパティ SPREADSHEET_ID が未設定です');
  if (!AUTH_SHEET_ID)  throw new Error('スクリプトプロパティ AUTH_SHEET_ID が未設定です');
}

// ============================================================
// セッション検証 + ユーザー情報取得（AUTH_SHEET を 1 回だけ開く）
// ※ 旧: validateSession + getUserInfo の 2 回オープンを統合
//
// 戻り値: { valid: true, user_id, name, is_admin } または { valid: false }
// ============================================================
function validateAndGetUser(token) {
  if (!token) return { valid: false };
  try {
    _checkProps();
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

    // --- 1. セッション検証 ---
    const sh = ss.getSheetByName('sessions');
    if (!sh) return { valid: false };
    const sessions = sh.getDataRange().getValues();
    const now = Date.now();
    let userId = null;

    for (let i = 1; i < sessions.length; i++) {
      if (String(sessions[i][0]) === String(token)) {
        if (Number(sessions[i][2]) < now) {
          sh.deleteRow(i + 1);
          return { valid: false };
        }
        userId = String(sessions[i][1]);
        break;
      }
    }
    if (!userId) return { valid: false };

    // --- 2. ユーザー情報取得（同じ SS を使い回す） ---
    const ush = ss.getSheetByName('users');
    if (!ush) return { valid: false };
    const uRows = ush.getDataRange().getValues();

    for (let i = 1; i < uRows.length; i++) {
      if (String(uRows[i][0]) === userId) {
        return {
          valid:    true,
          user_id:  userId,
          name:     String(uRows[i][1]),
          is_admin: uRows[i][5] === true || String(uRows[i][5]).toUpperCase() === 'TRUE'
        };
      }
    }
  } catch(e) {
    Logger.log('validateAndGetUser エラー: ' + e);
  }
  return { valid: false };
}

// ============================================================
// エントリーポイント（GET）
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  const data   = (e && e.parameter && e.parameter.data)   ? JSON.parse(e.parameter.data) : {};
  const token  = (e && e.parameter && e.parameter.session_token) ? e.parameter.session_token : '';

  // AUTH_SHEET を 1 回だけ開いてセッション検証＋ユーザー情報を取得
  const auth = validateAndGetUser(token);
  if (!auth.valid) return _jsonResponse(_err('SESSION_INVALID'));

  // ハンドラに渡す（ハンドラ内で getUserInfo を再呼び出ししない）
  data._userId   = auth.user_id;
  data._userInfo = auth;

  try {
    switch (action) {
      case 'init':            return _jsonResponse(initApp(data));
      case 'getProducts':     return _jsonResponse(getProducts(data));
      case 'getReservations': return _jsonResponse(getReservations(data));
      case 'getStats':        return _jsonResponse(getStats(data));
      case 'getUsers':        return _jsonResponse(getUsers(data));
      case 'getUserInfo':     return _jsonResponse(_ok({ user_id: auth.user_id, name: auth.name, is_admin: auth.is_admin }));
      default:                return _jsonResponse(_err('不明なアクション: ' + action));
    }
  } catch (err) {
    return _jsonResponse(_err(err.toString()));
  }
}

// ============================================================
// エントリーポイント（POST）
// ============================================================
function doPost(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';
  const data   = params.data ? JSON.parse(params.data) : {};
  const token  = params.session_token || '';

  const auth = validateAndGetUser(token);
  if (!auth.valid) return _jsonResponse(_err('SESSION_INVALID'));

  data._userId   = auth.user_id;
  data._userInfo = auth;

  try {
    switch (action) {
      case 'saveProduct':         return _jsonResponse(saveProduct(data));
      case 'toggleProductActive': return _jsonResponse(toggleProductActive(data));
      case 'saveReservation':     return _jsonResponse(saveReservation(data));
      case 'deleteReservation':   return _jsonResponse(deleteReservation(data));
      case 'updateStatus':        return _jsonResponse(updateStatus(data));
      default:                    return _jsonResponse(_err('不明なアクション: ' + action));
    }
  } catch (err) {
    return _jsonResponse(_err(err.toString()));
  }
}

// ============================================================
// ユーティリティ
// ============================================================
function _jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function _ok(data) { return { success: true,  data: data }; }
function _err(msg) { return { success: false, error: msg }; }
function _now() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

// ============================================================
// 初期化 API（高速化の核心）
// 起動時に必要な全データを 1 回のリクエストで返す
// ============================================================
function initApp(data) {
  const productsRes     = getProducts(data);
  const reservationsRes = getReservations(data);
  const usersRes        = getUsers(data);

  const auth = data._userInfo;
  return _ok({
    user:         { user_id: auth.user_id, name: auth.name, is_admin: auth.is_admin },
    products:     productsRes.data     || [],
    reservations: reservationsRes.data || [],
    users:        usersRes.data        || []
  });
}

// ============================================================
// スプレッドシート初期セットアップ
// GASエディタから手動で一度だけ実行すること
// ============================================================
function setupSheets() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // --- products シート ---
  let ps = ss.getSheetByName(SHEET_PRODUCTS);
  if (!ps) {
    ps = ss.insertSheet(SHEET_PRODUCTS);
    ps.getRange(1, 1, 1, 6).setValues([[
      'product_id', 'name', 'stock_limit', 'deadline', 'is_active', 'created_at'
    ]]);
    ps.setFrozenRows(1);
    ps.setColumnWidth(1, 160);
    ps.setColumnWidth(2, 200);
    Logger.log('productsシート作成完了');
  } else {
    Logger.log('productsシートは既に存在します');
  }

  // --- reservations シート ---
  let rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs) {
    rs = ss.insertSheet(SHEET_RESERVATIONS);
    rs.getRange(1, 1, 1, 13).setValues([[
      'reservation_no', 'salon_name', 'product_id', 'product_name',
      'quantity', 'status', 'staff_id', 'staff_name',
      'operator_id', 'operator_name', 'delivery_method',
      'reserved_at', 'updated_at'
    ]]);
    rs.setFrozenRows(1);
    Logger.log('reservationsシート作成完了');
  } else {
    Logger.log('reservationsシートは既に存在します');
  }

  Logger.log('セットアップ完了 ✅');
}

// ============================================================
// 商品マスター CRUD
// ============================================================

/**
 * 商品一覧取得
 * data.activeOnly = true の場合、受付中・期限内のみ返す（予約登録画面用）
 */
function getProducts(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ps = ss.getSheetByName(SHEET_PRODUCTS);
  if (!ps || ps.getLastRow() < 2) return _ok([]);

  const pRows = ps.getRange(2, 1, ps.getLastRow() - 1, 6).getValues();

  // 予約済み数を商品ごとに集計（ステータス問わず合算）
  const reservedMap = {};
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (rs && rs.getLastRow() >= 2) {
    const rRows = rs.getRange(2, 1, rs.getLastRow() - 1, 6).getValues();
    rRows.forEach(r => {
      if (r[0] === '' || r[0] === null) return;
      const pid = String(r[2]);
      const qty = Number(r[4]) || 0;
      reservedMap[pid] = (reservedMap[pid] || 0) + qty;
    });
  }

  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  const products = pRows
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => {
      const pid      = String(r[0]);
      const limit    = Number(r[2]) || 0;
      const reserved = reservedMap[pid] || 0;
      const deadlineRaw = r[3];
      const deadline = deadlineRaw
        ? Utilities.formatDate(new Date(deadlineRaw), 'Asia/Tokyo', 'yyyy-MM-dd')
        : '';
      const isExpired = deadline && deadline < todayStr;
      const isActive  = r[4] === true || String(r[4]).toUpperCase() === 'TRUE';

      return {
        product_id:     pid,
        name:           String(r[1]),
        stock_limit:    limit,
        reserved_total: reserved,
        remaining:      limit > 0 ? Math.max(0, limit - reserved) : null, // null=無制限
        deadline:       deadline,
        is_active:      isActive,
        is_expired:     !!isExpired,
        created_at:     r[5] ? String(r[5]) : ''
      };
    });

  // activeOnly=true のとき：受付中かつ期限内のみ（フロントエンド側でも計算可能だが互換性維持）
  if (data && data.activeOnly) {
    return _ok(products.filter(p => p.is_active && !p.is_expired));
  }
  return _ok(products);
}

/**
 * 商品登録・更新
 * product_id あり → 更新、なし → 新規登録
 */
function saveProduct(data) {
  const userInfo = data._userInfo;
  if (!userInfo || !userInfo.is_admin) return _err('権限がありません');

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ps   = ss.getSheetByName(SHEET_PRODUCTS);
  const name = String(data.name || '').trim();
  if (!name) return _err('商品名は必須です');

  const stockLimit = Number(data.stock_limit) || 0;
  const deadline   = data.deadline || '';
  const isActive   = data.is_active !== false;

  if (data.product_id) {
    // 更新
    if (!ps || ps.getLastRow() < 2) return _err('商品が見つかりません');
    const rows = ps.getRange(2, 1, ps.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.product_id)) {
        ps.getRange(i + 2, 2, 1, 4).setValues([[name, stockLimit, deadline, isActive]]);
        return _ok({ product_id: data.product_id, message: '更新しました' });
      }
    }
    return _err('商品が見つかりません');
  } else {
    // 新規
    const newId = 'P' + new Date().getTime();
    ps.appendRow([newId, name, stockLimit, deadline, isActive, _now()]);
    return _ok({ product_id: newId, message: '登録しました' });
  }
}

/**
 * 商品の有効 / 無効を切り替える
 */
function toggleProductActive(data) {
  const userInfo = data._userInfo;
  if (!userInfo || !userInfo.is_admin) return _err('権限がありません');

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ps   = ss.getSheetByName(SHEET_PRODUCTS);
  if (!ps || ps.getLastRow() < 2) return _err('商品が見つかりません');

  const rows = ps.getRange(2, 1, ps.getLastRow() - 1, 5).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.product_id)) {
      const newActive = !(rows[i][4] === true || String(rows[i][4]).toUpperCase() === 'TRUE');
      ps.getRange(i + 2, 5).setValue(newActive);
      return _ok({ is_active: newActive });
    }
  }
  return _err('商品が見つかりません');
}

// ============================================================
// 予約 CRUD
// ============================================================

/**
 * 予約一覧取得
 * 営業（is_admin=false）: 自分の予約のみ
 * 事務（is_admin=true） : 全件
 */
function getReservations(data) {
  const userInfo = data._userInfo;
  if (!userInfo) return _err('ユーザー情報が取得できません');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs || rs.getLastRow() < 2) return _ok([]);

  const rows = rs.getRange(2, 1, rs.getLastRow() - 1, 13).getValues();

  let list = rows
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      reservation_no:  Number(r[0]),
      salon_name:      String(r[1]),
      product_id:      String(r[2]),
      product_name:    String(r[3]),
      quantity:        Number(r[4]),
      status:          String(r[5]),
      staff_id:        String(r[6]),
      staff_name:      String(r[7]),
      operator_id:     String(r[8]),
      operator_name:   String(r[9]),
      delivery_method: String(r[10]),
      reserved_at:     r[11] ? String(r[11]) : '',
      updated_at:      r[12] ? String(r[12]) : ''
    }));

  // 営業は自分の予約のみ表示
  if (!userInfo.is_admin) {
    list = list.filter(r => r.staff_id === String(data._userId));
  }

  // 予約No 降順ソート
  list.sort((a, b) => b.reservation_no - a.reservation_no);
  return _ok(list);
}

/**
 * 予約登録・更新
 * reservation_no あり → 更新、なし → 新規登録
 * 権限チェック:
 *   営業: 自分の「予約」ステータスのみ変更可
 *   事務: 全件変更可
 */
function saveReservation(data) {
  const userInfo = data._userInfo;
  if (!userInfo) return _err('ユーザー情報が取得できません');

  // 入力検証
  const salonName      = String(data.salon_name      || '').trim();
  const productId      = String(data.product_id      || '');
  const quantity       = Number(data.quantity)       || 0;
  const staffId        = String(data.staff_id        || '');
  const staffName      = String(data.staff_name      || '');
  const deliveryMethod = String(data.delivery_method || '未定');

  if (!salonName)   return _err('サロン名は必須です');
  if (!productId)   return _err('商品を選択してください');
  if (quantity < 1) return _err('数量は1以上を指定してください');
  if (!staffId)     return _err('担当者を選択してください');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ps = ss.getSheetByName(SHEET_PRODUCTS);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);

  // 商品情報取得
  let product = null;
  if (ps && ps.getLastRow() >= 2) {
    const pRows = ps.getRange(2, 1, ps.getLastRow() - 1, 5).getValues();
    for (const pr of pRows) {
      if (String(pr[0]) === productId) {
        product = {
          name:        String(pr[1]),
          stock_limit: Number(pr[2]) || 0,
          deadline:    pr[3] ? Utilities.formatDate(new Date(pr[3]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
          is_active:   pr[4] === true || String(pr[4]).toUpperCase() === 'TRUE'
        };
        break;
      }
    }
  }
  if (!product)           return _err('商品が見つかりません');
  if (!product.is_active) return _err('この商品は現在受付停止中です');

  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  if (product.deadline && product.deadline < todayStr) {
    return _err('この商品の受付期限が過ぎています');
  }

  // 在庫チェック（stock_limit > 0 の場合のみ）
  if (product.stock_limit > 0) {
    let alreadyReserved = 0;
    if (rs && rs.getLastRow() >= 2) {
      const rRows = rs.getRange(2, 1, rs.getLastRow() - 1, 6).getValues();
      for (const rr of rRows) {
        if (String(rr[2]) === productId) {
          // 更新時は自分の予約を除いてカウント
          if (data.reservation_no && Number(rr[0]) === Number(data.reservation_no)) continue;
          alreadyReserved += Number(rr[4]) || 0;
        }
      }
    }
    const remaining = product.stock_limit - alreadyReserved;
    if (quantity > remaining) {
      return _err(`予約可能数を超えています（残り ${remaining} 個）`);
    }
  }

  const now = _now();

  if (data.reservation_no) {
    // ---- 更新 ----
    if (!rs || rs.getLastRow() < 2) return _err('予約が見つかりません');
    const rows = rs.getRange(2, 1, rs.getLastRow() - 1, 13).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (Number(rows[i][0]) === Number(data.reservation_no)) {
        // 権限チェック
        if (!userInfo.is_admin) {
          if (String(rows[i][6]) !== String(data._userId)) return _err('他の担当者の予約は変更できません');
          if (String(rows[i][5]) !== '予約') return _err('確定済みの予約は変更できません');
        }
        // B〜K列（2〜11列目）を更新。ステータスは変更しない
        rs.getRange(i + 2, 2, 1, 10).setValues([[
          salonName, productId, product.name, quantity,
          rows[i][5], // ステータス維持
          staffId, staffName,
          String(data._userId), userInfo.name,
          deliveryMethod
        ]]);
        rs.getRange(i + 2, 13).setValue(now); // updated_at
        return _ok({
          reservation_no: Number(data.reservation_no),
          message: '更新しました',
          product_name: product.name
        });
      }
    }
    return _err('予約が見つかりません');

  } else {
    // ---- 新規登録 ----
    let newNo = 1;
    if (rs && rs.getLastRow() >= 2) {
      const nums = rs.getRange(2, 1, rs.getLastRow() - 1, 1).getValues()
        .map(r => Number(r[0]) || 0);
      newNo = Math.max(...nums) + 1;
    }
    rs.appendRow([
      newNo, salonName, productId, product.name, quantity, '予約',
      staffId, staffName,
      String(data._userId), userInfo.name,
      deliveryMethod, now, now
    ]);
    return _ok({
      reservation_no: newNo,
      message: '予約を登録しました',
      // フロントエンドでローカル更新するために必要な情報を返す
      reservation: {
        reservation_no:  newNo,
        salon_name:      salonName,
        product_id:      productId,
        product_name:    product.name,
        quantity:        quantity,
        status:          '予約',
        staff_id:        staffId,
        staff_name:      staffName,
        operator_id:     String(data._userId),
        operator_name:   userInfo.name,
        delivery_method: deliveryMethod,
        reserved_at:     now,
        updated_at:      now
      }
    });
  }
}

/**
 * 予約削除
 * 営業: 自分の「予約」ステータスのみ削除可
 * 事務: 全件削除可
 */
function deleteReservation(data) {
  const userInfo = data._userInfo;
  if (!userInfo) return _err('ユーザー情報が取得できません');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs || rs.getLastRow() < 2) return _err('予約が見つかりません');

  const rows = rs.getRange(2, 1, rs.getLastRow() - 1, 7).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(data.reservation_no)) {
      if (!userInfo.is_admin) {
        if (String(rows[i][6]) !== String(data._userId)) return _err('他の担当者の予約は削除できません');
        if (String(rows[i][5]) !== '予約') return _err('確定済みの予約は削除できません');
      }
      rs.deleteRow(i + 2);
      return _ok({ message: '削除しました' });
    }
  }
  return _err('予約が見つかりません');
}

/**
 * ステータス変更（事務のみ）
 * status: '予約' | '確定'
 */
function updateStatus(data) {
  const userInfo = data._userInfo;
  if (!userInfo || !userInfo.is_admin) return _err('権限がありません');

  const validStatuses = ['予約', '確定'];
  if (!validStatuses.includes(data.status)) return _err('無効なステータスです');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs || rs.getLastRow() < 2) return _err('予約が見つかりません');

  const rows = rs.getRange(2, 1, rs.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(data.reservation_no)) {
      rs.getRange(i + 2, 6).setValue(data.status);
      rs.getRange(i + 2, 13).setValue(_now());
      return _ok({ message: 'ステータスを更新しました' });
    }
  }
  return _err('予約が見つかりません');
}

// ============================================================
// ユーザー一覧（担当者ドロップダウン用）
// is_admin=FALSE かつ active=TRUE の営業スタッフのみ返す
// ============================================================
function getUsers(data) {
  try {
    const ss   = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const sh   = ss.getSheetByName('users');
    if (!sh || sh.getLastRow() < 2) return _ok([]);

    const rows  = sh.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < rows.length; i++) {
      const active   = rows[i][3] === true || String(rows[i][3]).toUpperCase() === 'TRUE';
      const is_admin = rows[i][5] === true || String(rows[i][5]).toUpperCase() === 'TRUE';
      if (active && !is_admin) {
        users.push({ user_id: String(rows[i][0]), name: String(rows[i][1]) });
      }
    }
    return _ok(users);
  } catch(e) {
    return _err('ユーザー取得エラー: ' + e);
  }
}

// ============================================================
// 集計データ取得（事務のみ）
// ============================================================
function getStats(data) {
  const userInfo = data._userInfo;
  if (!userInfo || !userInfo.is_admin) return _err('権限がありません');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  const ps = ss.getSheetByName(SHEET_PRODUCTS);

  const empty = {
    byProduct: [], byStaff: [], bySalon: [],
    crossTable: { staffList: [], salonList: [], data: {} }
  };
  if (!rs || rs.getLastRow() < 2) return _ok(empty);

  const allRows = rs.getRange(2, 1, rs.getLastRow() - 1, 13).getValues()
    .filter(r => r[0] !== '' && r[0] !== null);

  // --- 商品別集計 ---
  const productMap = {};
  allRows.forEach(r => {
    const pid  = String(r[2]);
    const name = String(r[3]);
    const qty  = Number(r[4]) || 0;
    const st   = String(r[5]);
    if (!productMap[pid]) {
      productMap[pid] = { product_id: pid, product_name: name, total: 0, confirmed: 0, stock_limit: 0, remaining: null };
    }
    productMap[pid].total += qty;
    if (st === '確定') productMap[pid].confirmed += qty;
  });

  // 在庫上限・残数を商品マスターから補完
  if (ps && ps.getLastRow() >= 2) {
    const pRows = ps.getRange(2, 1, ps.getLastRow() - 1, 3).getValues();
    pRows.forEach(pr => {
      const pid   = String(pr[0]);
      const limit = Number(pr[2]) || 0;
      if (productMap[pid] && limit > 0) {
        productMap[pid].stock_limit = limit;
        productMap[pid].remaining   = Math.max(0, limit - productMap[pid].total);
      }
    });
  }
  const byProduct = Object.values(productMap).sort((a, b) => b.total - a.total);

  // --- 担当別集計 ---
  const staffMap = {};
  allRows.forEach(r => {
    const key = String(r[6]);
    if (!staffMap[key]) staffMap[key] = { staff_id: key, staff_name: String(r[7]), total: 0 };
    staffMap[key].total += Number(r[4]) || 0;
  });
  const byStaff = Object.values(staffMap).sort((a, b) => b.total - a.total);

  // --- サロン別集計 ---
  const salonMap = {};
  allRows.forEach(r => {
    const key = String(r[1]);
    if (!salonMap[key]) salonMap[key] = { salon_name: key, total: 0 };
    salonMap[key].total += Number(r[4]) || 0;
  });
  const bySalon = Object.values(salonMap).sort((a, b) => b.total - a.total);

  // --- クロス集計（担当 × サロン）---
  const filterPid  = data.filter_product_id || '';
  const crossRows  = filterPid ? allRows.filter(r => String(r[2]) === filterPid) : allRows;

  const crossData     = {};
  const staffNamesMap = {};
  const salonNamesSet = {};
  crossRows.forEach(r => {
    const sid   = String(r[6]);
    const sname = String(r[7]);
    const salon = String(r[1]);
    const qty   = Number(r[4]) || 0;
    staffNamesMap[sid]    = sname;
    salonNamesSet[salon]  = true;
    if (!crossData[sid]) crossData[sid] = {};
    crossData[sid][salon] = (crossData[sid][salon] || 0) + qty;
  });

  const staffList = Object.keys(staffNamesMap).map(id => ({ id, name: staffNamesMap[id] }));
  const salonList = Object.keys(salonNamesSet).sort();

  return _ok({
    byProduct,
    byStaff,
    bySalon,
    crossTable: { staffList, salonList, data: crossData }
  });
}
