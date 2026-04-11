// ============================================================
// Beaufield 予約管理アプリ - Google Apps Script
// Version: 1.4.0
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   SPREADSHEET_ID  : 予約管理データのスプレッドシートID
//   AUTH_SHEET_ID   : beaufield-auth スプレッドシートID（共通）
//
// ============================================================

const VERSION  = '1.7.1';
const APP_NAME = 'yoyaku-kanri';

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS         = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = _PROPS.getProperty('SPREADSHEET_ID');
const AUTH_SHEET_ID  = _PROPS.getProperty('AUTH_SHEET_ID');

// シート名定数
const SHEET_PRODUCTS     = 'products';
const SHEET_RESERVATIONS = 'reservations';

// CacheService キャッシュ時間（秒）
const CACHE_TTL_AUTH  = 900;  // セッション検証キャッシュ: 15分（増加でシート読み込み頻度削減）
const CACHE_TTL_USERS = 600;  // ユーザー一覧キャッシュ: 10分（_getUsersFromAuth の二重シート読み込みを防ぐ）

// ============================================================
// コールドスタート対策 ── Keep-Warm トリガー
// ============================================================
// GASは一定時間使われないとスクリプト環境が破棄され、
// 次回アクセス時に「コールドスタート」（15〜30秒の待ち）が発生する。
// 対策: この関数を「時間ベーストリガー」で5〜10分ごとに実行するよう設定する。
//
// 設定手順（GASエディタ）:
//   ① 左メニュー「トリガー（時計アイコン）」→「トリガーを追加」
//   ② 実行する関数: keepWarm
//   ③ イベントのソース: 時間主導型
//   ④ 時間ベースのトリガーのタイプ: 分ベースのタイマー → 5分おき
//   ⑤ 保存
//
// ※ GAS実行クォータへの影響: 1回あたり0.1秒以下のため、1日で約30秒消費（問題なし）
// ============================================================
function keepWarm() {
  // スクリプト環境を常にウォーム状態に保つ軽量関数
  // キャッシュに何も書き込まず、プロパティを1つ読むだけで十分
  PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

// ============================================================
// 起動時チェック（プロパティ未設定を早期検知）
// ============================================================
function _checkProps() {
  if (!SPREADSHEET_ID) throw new Error('スクリプトプロパティ SPREADSHEET_ID が未設定です');
  if (!AUTH_SHEET_ID)  throw new Error('スクリプトプロパティ AUTH_SHEET_ID が未設定です');
}

// ============================================================
// ① セッション検証 + ユーザー情報取得
//    beaufield-auth の user_app_roles シートからアプリ固有ロール取得
//    yoyaku_role: "admin"（事務） / "staff"（営業） / null（未登録）
//    CacheService で 5 分キャッシュ → 連続リクエスト時のシート読み込みを削減
// ============================================================
function validateAndGetUser(token) {
  if (!token) return { valid: false };
  try {
    _checkProps();

    // キャッシュヒット確認
    const cache    = CacheService.getScriptCache();
    const cacheKey = 'auth_' + token;
    const cached   = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

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
          // 期限切れ: ここでは削除せず invalid を返す（削除は cleanExpiredSessions バッチで行う）
          return { valid: false };
        }
        userId = String(sessions[i][1]);
        break;
      }
    }
    if (!userId) return { valid: false };

    // --- 2. ユーザー情報取得 ---
    const ush = ss.getSheetByName('users');
    if (!ush) return { valid: false };
    const uRows = ush.getDataRange().getValues();
    let userName = null;

    for (let i = 1; i < uRows.length; i++) {
      if (String(uRows[i][0]) === userId) {
        userName = String(uRows[i][1]);
        break;
      }
    }
    if (!userName) return { valid: false };

    // --- 3. user_app_roles からアプリ固有ロール取得 ---
    //    "admin" → 事務（商品管理・全予約管理）
    //    "staff" → 営業（担当者として登録・自分の予約のみ操作）
    //    未登録  → null（アクセス不可）
    const arSh = ss.getSheetByName('user_app_roles');
    let yoyakuRole = null;
    if (arSh && arSh.getLastRow() >= 2) {
      const arRows = arSh.getDataRange().getValues();
      for (let i = 1; i < arRows.length; i++) {
        if (String(arRows[i][0]) === userId && String(arRows[i][1]) === APP_NAME) {
          yoyakuRole = String(arRows[i][2]).trim().toLowerCase() || null;
          break;
        }
      }
    }

    const result = {
      valid:       true,
      user_id:     userId,
      name:        userName,
      yoyaku_role: yoyakuRole   // "admin" / "staff" / null
    };
    cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_AUTH);
    return result;

  } catch(e) {
    Logger.log('validateAndGetUser エラー: ' + e);
  }
  return { valid: false };
}

// ============================================================
// 期限切れセッション削除バッチ（GASトリガーで夜間に実行推奨）
// GASエディタ → トリガー → cleanExpiredSessions → 時間ベース → 毎日
// ============================================================
function cleanExpiredSessions() {
  try {
    _checkProps();
    const ss  = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const sh  = ss.getSheetByName('sessions');
    if (!sh || sh.getLastRow() < 2) return;
    const now  = Date.now();
    const rows = sh.getDataRange().getValues();
    // 後ろから削除（行番号ずれ防止）
    for (let i = rows.length - 1; i >= 1; i--) {
      if (Number(rows[i][2]) < now) sh.deleteRow(i + 1);
    }
    Logger.log('期限切れセッション削除完了');
  } catch(e) {
    Logger.log('cleanExpiredSessions エラー: ' + e);
  }
}

// ============================================================
// エントリーポイント（GET）
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  const data   = (e && e.parameter && e.parameter.data)   ? JSON.parse(e.parameter.data) : {};
  const token  = (e && e.parameter && e.parameter.session_token) ? e.parameter.session_token : '';

  const auth = validateAndGetUser(token);
  if (!auth.valid) return _jsonResponse(_err('SESSION_INVALID'));

  data._userId   = auth.user_id;
  data._userInfo = auth;

  try {
    switch (action) {
      case 'init':            return _jsonResponse(initApp(data));
      case 'getProducts':     return _jsonResponse(getProducts(data));
      case 'getReservations': return _jsonResponse(getReservations(data));
      case 'getUsers':        return _jsonResponse(getUsers(data));
      case 'getUserInfo':     return _jsonResponse(_ok({ user_id: auth.user_id, name: auth.name, yoyaku_role: auth.yoyaku_role }));
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
      case 'bulkDelete':          return _jsonResponse(bulkDelete(data));
      case 'bulkUpdateStatus':    return _jsonResponse(bulkUpdateStatus(data));
      case 'bulkUpdateDelivery':  return _jsonResponse(bulkUpdateDelivery(data));
      case 'processArrival':      return _jsonResponse(processArrival(data));
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
// ② 初期化 API
//    user_app_roles 未登録ユーザーはここで ACCESS_DENIED を返す
//    SPREADSHEET_ID を 1 回だけ openById → products/reservations に使い回す
// ============================================================
function initApp(data) {
  _checkProps();
  const auth = data._userInfo;

  // user_app_roles に yoyaku-kanri のエントリがないユーザーはアクセス不可
  if (!auth.yoyaku_role) return _err('ACCESS_DENIED');

  // SPREADSHEET_ID を 1 回だけオープン
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  return _ok({
    user: {
      user_id:     auth.user_id,
      name:        auth.name,
      yoyaku_role: auth.yoyaku_role  // "admin" or "staff"
    },
    products:     _getProductsFromSS(ss),
    reservations: _getReservationsFromSS(ss),
    users:        _getUsersFromAuth()
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
 * 内部ヘルパー: 既に開いた ss を受け取って商品一覧を返す
 * reserved_total はフロントエンドのローカルデータから計算するため GAS では集計しない
 */
function _getProductsFromSS(ss) {
  const ps = ss.getSheetByName(SHEET_PRODUCTS);
  if (!ps || ps.getLastRow() < 2) return [];

  const pRows    = ps.getRange(2, 1, ps.getLastRow() - 1, 6).getValues();
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  return pRows
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => {
      const pid         = String(r[0]);
      const limit       = Number(r[2]) || 0;
      const deadlineRaw = r[3];
      const deadline    = deadlineRaw
        ? Utilities.formatDate(new Date(deadlineRaw), 'Asia/Tokyo', 'yyyy-MM-dd')
        : '';
      const isExpired = deadline && deadline < todayStr;
      const isActive  = r[4] === true || String(r[4]).toUpperCase() === 'TRUE';

      return {
        product_id:     pid,
        name:           String(r[1]),
        stock_limit:    limit,
        reserved_total: 0,      // フロントエンドが reservations から計算して上書きする
        remaining:      limit > 0 ? limit : null, // 同上（フロントが更新）
        deadline:       deadline,
        is_active:      isActive,
        is_expired:     !!isExpired,
        created_at:     r[5] ? String(r[5]) : ''
      };
    });
}

/** 公開 API: getProducts（単独呼び出し用） */
function getProducts(data) {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ok(_getProductsFromSS(ss));
}

/**
 * 商品登録・更新（adminのみ）
 */
function saveProduct(data) {
  const userInfo = data._userInfo;
  if (!userInfo || userInfo.yoyaku_role !== 'admin') return _err('権限がありません');

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ps   = ss.getSheetByName(SHEET_PRODUCTS);
  const name = String(data.name || '').trim();
  if (!name) return _err('商品名は必須です');

  const stockLimit = Number(data.stock_limit) || 0;
  const deadline   = data.deadline || '';
  const isActive   = data.is_active !== false;

  if (data.product_id) {
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
    const newId = 'P' + new Date().getTime();
    ps.appendRow([newId, name, stockLimit, deadline, isActive, _now()]);
    return _ok({ product_id: newId, message: '登録しました' });
  }
}

/**
 * 商品の有効 / 無効を切り替える（adminのみ）
 */
function toggleProductActive(data) {
  const userInfo = data._userInfo;
  if (!userInfo || userInfo.yoyaku_role !== 'admin') return _err('権限がありません');

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
 * 内部ヘルパー: 既に開いた ss を受け取って予約一覧を返す
 */
function _getReservationsFromSS(ss) {
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs || rs.getLastRow() < 2) return [];

  const rows = rs.getRange(2, 1, rs.getLastRow() - 1, 13).getValues();
  const list = rows
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
      // セル値がDate型になる場合があるためUtilitiesで書式変換（String()はGMT表記になるため使用禁止）
      reserved_at:     r[11] ? Utilities.formatDate(new Date(r[11]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : '',
      updated_at:      r[12] ? Utilities.formatDate(new Date(r[12]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : ''
    }));

  list.sort((a, b) => b.reservation_no - a.reservation_no);
  return list;
}

/** 公開 API: getReservations（単独呼び出し用） */
function getReservations(data) {
  const userInfo = data._userInfo;
  if (!userInfo) return _err('ユーザー情報が取得できません');
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ok(_getReservationsFromSS(ss));
}

/**
 * 予約登録・更新
 */
function saveReservation(data) {
  const userInfo = data._userInfo;
  if (!userInfo) return _err('ユーザー情報が取得できません');

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

  _checkProps();
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

  const now    = _now();
  const nowFmt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');

  if (data.reservation_no) {
    // ---- 更新 ----
    if (!rs || rs.getLastRow() < 2) return _err('予約が見つかりません');
    const rows = rs.getRange(2, 1, rs.getLastRow() - 1, 13).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (Number(rows[i][0]) === Number(data.reservation_no)) {
        if (userInfo.yoyaku_role !== 'admin') {
          if (String(rows[i][6]) !== String(data._userId)) return _err('他の担当者の予約は変更できません');
          if (String(rows[i][5]) !== '予約') return _err('確定済みの予約は変更できません');
        }
        // 数量増加は予約優先度の公平性を損なうため拒否（admin含む全ユーザー）
        const originalQty = Number(rows[i][4]);
        if (quantity > originalQty) {
          return _err(`数量を増やすことはできません（現在: ${originalQty}個）`);
        }
        rs.getRange(i + 2, 2, 1, 10).setValues([[
          salonName, productId, product.name, quantity,
          rows[i][5],
          staffId, staffName,
          String(data._userId), userInfo.name,
          deliveryMethod
        ]]);
        rs.getRange(i + 2, 13).setValue(now);
        return _ok({
          reservation_no: Number(data.reservation_no),
          message:        '更新しました',
          product_name:   product.name,
          reservation: {
            reservation_no:  Number(data.reservation_no),
            salon_name:      salonName,
            product_id:      productId,
            product_name:    product.name,
            quantity:        quantity,
            status:          String(rows[i][5]),
            staff_id:        staffId,
            staff_name:      staffName,
            operator_id:     String(data._userId),
            operator_name:   userInfo.name,
            delivery_method: deliveryMethod,
            reserved_at:     rows[i][11]
              ? Utilities.formatDate(new Date(rows[i][11]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')
              : nowFmt,
            updated_at: nowFmt
          }
        });
      }
    }
    return _err('予約が見つかりません');

  } else {
    // ---- 新規登録 ----
    const newNo = rs && rs.getLastRow() >= 2
      ? Math.max(...rs.getRange(2, 1, rs.getLastRow() - 1, 1).getValues().map(r => Number(r[0]) || 0)) + 1
      : 1;

    rs.appendRow([
      newNo, salonName, productId, product.name, quantity, '予約',
      staffId, staffName,
      String(data._userId), userInfo.name,
      deliveryMethod, now, now
    ]);
    return _ok({
      reservation_no: newNo,
      message: '予約を登録しました',
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
        reserved_at:     nowFmt,
        updated_at:      nowFmt
      }
    });
  }
}

/**
 * 予約削除
 */
function deleteReservation(data) {
  const userInfo = data._userInfo;
  if (!userInfo) return _err('ユーザー情報が取得できません');

  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs || rs.getLastRow() < 2) return _err('予約が見つかりません');

  const rows = rs.getRange(2, 1, rs.getLastRow() - 1, 7).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(data.reservation_no)) {
      if (userInfo.yoyaku_role !== 'admin') {
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
 * ステータス変更（adminのみ）
 */
function updateStatus(data) {
  const userInfo = data._userInfo;
  if (!userInfo || userInfo.yoyaku_role !== 'admin') return _err('権限がありません');

  const validStatuses = ['予約', '確定'];
  if (!validStatuses.includes(data.status)) return _err('無効なステータスです');

  _checkProps();
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
// 一括操作
// ============================================================

/**
 * 複数予約を一括削除
 */
function bulkDelete(data) {
  const userInfo = data._userInfo;
  if (!userInfo) return _err('ユーザー情報が取得できません');

  const nos = data.reservation_nos;
  if (!Array.isArray(nos) || nos.length === 0) return _err('削除対象が指定されていません');

  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs || rs.getLastRow() < 2) return _err('予約が見つかりません');

  const rows   = rs.getRange(2, 1, rs.getLastRow() - 1, 7).getValues();
  const nosSet = new Set(nos.map(Number));
  let deleted  = 0;
  const errors = [];

  for (let i = rows.length - 1; i >= 0; i--) {
    const no = Number(rows[i][0]);
    if (!nosSet.has(no)) continue;

    if (userInfo.yoyaku_role !== 'admin') {
      if (String(rows[i][6]) !== String(data._userId)) {
        errors.push(`No.${no}: 他の担当者の予約は削除できません`);
        continue;
      }
      if (String(rows[i][5]) !== '予約') {
        errors.push(`No.${no}: 確定済みの予約は削除できません`);
        continue;
      }
    }
    rs.deleteRow(i + 2);
    deleted++;
  }

  return _ok({ deleted, errors });
}

/**
 * 複数予約を一括ステータス変更（adminのみ）
 */
function bulkUpdateStatus(data) {
  const userInfo = data._userInfo;
  if (!userInfo || userInfo.yoyaku_role !== 'admin') return _err('権限がありません');

  const nos    = data.reservation_nos;
  const status = data.status;
  const validStatuses = ['予約', '確定'];
  if (!validStatuses.includes(status))         return _err('無効なステータスです');
  if (!Array.isArray(nos) || nos.length === 0) return _err('対象が指定されていません');

  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs || rs.getLastRow() < 2) return _err('予約が見つかりません');

  const rows   = rs.getRange(2, 1, rs.getLastRow() - 1, 1).getValues();
  const nosSet = new Set(nos.map(Number));
  const now    = _now();
  let updated  = 0;

  for (let i = 0; i < rows.length; i++) {
    if (nosSet.has(Number(rows[i][0]))) {
      rs.getRange(i + 2, 6).setValue(status);
      rs.getRange(i + 2, 13).setValue(now);
      updated++;
    }
  }

  return _ok({ updated });
}

/**
 * 複数予約の発送方法を一括変更
 * admin: 全件変更可
 * staff: 自分の「予約」ステータスのみ変更可
 */
function bulkUpdateDelivery(data) {
  const userInfo = data._userInfo;
  if (!userInfo) return _err('ユーザー情報が取得できません');

  const nos            = data.reservation_nos;
  const deliveryMethod = String(data.delivery_method || '').trim();
  const validMethods   = ['発送', '持参', '未定'];
  if (!validMethods.includes(deliveryMethod))       return _err('無効な発送方法です');
  if (!Array.isArray(nos) || nos.length === 0)      return _err('対象が指定されていません');

  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs || rs.getLastRow() < 2) return _err('予約が見つかりません');

  const rows   = rs.getRange(2, 1, rs.getLastRow() - 1, 7).getValues();
  const nosSet = new Set(nos.map(Number));
  const now    = _now();
  let updated  = 0;
  const errors = [];

  for (let i = 0; i < rows.length; i++) {
    const no = Number(rows[i][0]);
    if (!nosSet.has(no)) continue;

    if (userInfo.yoyaku_role !== 'admin') {
      if (String(rows[i][6]) !== String(data._userId)) {
        errors.push(`No.${no}: 他の担当者の予約は変更できません`);
        continue;
      }
      if (String(rows[i][5]) !== '予約') {
        errors.push(`No.${no}: 確定済みの予約は変更できません`);
        continue;
      }
    }
    rs.getRange(i + 2, 11).setValue(deliveryMethod);  // K列: delivery_method
    rs.getRange(i + 2, 13).setValue(now);              // M列: updated_at
    updated++;
  }

  return _ok({ updated, errors });
}

// ============================================================
// 入荷処理（adminのみ）
//   allocations: [{reservation_no, confirm_qty}]
//   confirm_qty === quantity  → ステータスを「確定」に変更
//   confirm_qty < quantity    → 元行: 数量を (quantity - confirm_qty) に減らす（予約継続）
//                               新行: confirm_qty / ステータス「確定」/ 新No
//   元行のNoを維持することで予約順位が保たれる（スプリット設計）
// ============================================================
function processArrival(data) {
  const userInfo = data._userInfo;
  if (!userInfo || userInfo.yoyaku_role !== 'admin') return _err('権限がありません');

  const productId   = String(data.product_id || '');
  const allocations = data.allocations; // [{reservation_no, confirm_qty}]

  if (!productId)                                              return _err('商品IDは必須です');
  if (!Array.isArray(allocations) || allocations.length === 0) return _err('割り当て情報がありません');

  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!rs || rs.getLastRow() < 2) return _err('予約データが見つかりません');

  const rows = rs.getRange(2, 1, rs.getLastRow() - 1, 13).getValues();
  const now  = _now();

  // allocations を Map 化（reservation_no → confirm_qty）
  const allocMap = {};
  for (const a of allocations) {
    const no  = Number(a.reservation_no);
    const qty = Number(a.confirm_qty);
    if (no > 0 && qty > 0) allocMap[no] = qty;
  }

  if (Object.keys(allocMap).length === 0) return _err('有効な割り当て情報がありません');

  // 現在の最大No（スプリット時の新No発番用）
  let maxNo = Math.max(...rows.map(r => Number(r[0]) || 0));

  // スプリット行バッファ（appendRow をまとめて最後に実行）
  const newRows = [];
  let processed = 0;

  for (let i = 0; i < rows.length; i++) {
    const no = Number(rows[i][0]);
    if (allocMap[no] === undefined) continue;

    const confirmQty = allocMap[no];
    const curQty     = Number(rows[i][4]);
    const curStatus  = String(rows[i][5]);

    // 念のためガード（確定済みや数量超過は処理しない）
    if (curStatus === '確定')        continue;
    if (confirmQty > curQty)         continue;
    if (confirmQty <= 0)             continue;

    if (confirmQty === curQty) {
      // ── 全数確定: ステータスだけ「確定」に変更 ────────────
      rs.getRange(i + 2, 6).setValue('確定');
      rs.getRange(i + 2, 13).setValue(now);
    } else {
      // ── 部分確定: スプリット処理 ──────────────────────────
      // 元行: 残数（= curQty - confirmQty）に更新、ステータスは「予約」のまま（No変わらず=優先順位維持）
      rs.getRange(i + 2, 5).setValue(curQty - confirmQty);
      rs.getRange(i + 2, 13).setValue(now);

      // 新行: 確定分を新No で追記
      maxNo++;
      newRows.push([
        maxNo,
        rows[i][1],              // salon_name
        rows[i][2],              // product_id
        rows[i][3],              // product_name
        confirmQty,              // 確定数量
        '確定',
        rows[i][6],              // staff_id (コピー)
        rows[i][7],              // staff_name (コピー)
        String(userInfo.user_id), // operator_id（入荷処理実行者）
        userInfo.name,            // operator_name
        rows[i][10],             // delivery_method (コピー)
        rows[i][11],             // reserved_at (元の予約日時をコピー)
        now                      // updated_at
      ]);
    }
    processed++;
  }

  // スプリット行を一括追加
  for (const row of newRows) {
    rs.appendRow(row);
  }

  if (processed === 0) return _err('処理対象の予約が見つかりませんでした');

  return _ok({
    message:  `入荷処理が完了しました（${processed}件処理、うちスプリット${newRows.length}件）`,
    processed: processed,
    splits:    newRows.length
  });
}

// ============================================================
// ユーザー一覧
//   user_app_roles シートで yoyaku-kanri に登録済みのユーザーのみ返す
//   yoyaku_role: "admin"（事務・担当者非表示）/ "staff"（営業・担当者として表示）
//
// ※ CacheService を使わない理由:
//   user_app_roles に新規スタッフを追加してもキャッシュが10分間残ると
//   担当者フィルタに反映されないため、毎回シートから読む。
//   呼び出しは initApp（ページロード時のみ）に限定されるため
//   フロントの localStorage 5分キャッシュで十分に保護される。
// ============================================================
function _getUsersFromAuth() {
  try {
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

    // user_app_roles から yoyaku-kanri 登録ユーザーとロールを取得
    const arSh = ss.getSheetByName('user_app_roles');
    const roleMap = {};
    if (arSh && arSh.getLastRow() >= 2) {
      const arRows = arSh.getDataRange().getValues();
      for (let i = 1; i < arRows.length; i++) {
        if (String(arRows[i][1]) === APP_NAME) {
          const uid  = String(arRows[i][0]);
          const role = String(arRows[i][2]).trim().toLowerCase();  // 大文字小文字を吸収
          if (uid && role) roleMap[uid] = role;
        }
      }
    }

    const sh = ss.getSheetByName('users');
    if (!sh || sh.getLastRow() < 2) return [];

    const rows  = sh.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < rows.length; i++) {
      const active = rows[i][3] === true || String(rows[i][3]).toUpperCase() === 'TRUE';
      if (!active) continue;
      const uid        = String(rows[i][0]);
      const yoyakuRole = roleMap[uid];
      if (!yoyakuRole) continue;  // yoyaku-kanri 未登録はスキップ

      const shortName = String(rows[i][7] || '').trim();
      users.push({
        user_id:     uid,
        name:        String(rows[i][1]),
        short_name:  shortName || String(rows[i][1]),
        yoyaku_role: yoyakuRole  // "admin" or "staff"（常に小文字）
      });
    }
    return users;
  } catch(e) {
    Logger.log('_getUsersFromAuth エラー: ' + e);
    return [];
  }
}

/** 公開 API: getUsers（単独呼び出し用） */
function getUsers(data) {
  return _ok(_getUsersFromAuth());
}
