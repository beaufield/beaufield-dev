// ============================================================
// Beaufield ルート訪問チェッカー - Google Apps Script
// Version: 1.1.0
// ============================================================
// [重要] デプロイ前に SPREADSHEET_ID を必ず設定してください
// ============================================================

const SPREADSHEET_ID = '1yVd3yI9v8acjyKaM-fCs_VoBnOx44mnRDOBVGzBB288';
const VERSION = '1.1.4';

// シート名定数
const SHEET_USERS      = 'users';
const SHEET_SALONS     = 'salons';
const SHEET_VISIT_LOGS = 'visit_logs';

// 曜日マッピング（日本語曜日 → JS の getDay() 値）
const DAY_MAP   = { '日': 0, '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6 };
const DAY_NAMES = ['日', '月', '火', '水', '木', '金', '土'];

// ============================================================
// エントリーポイント（GET）
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  const data   = (e && e.parameter && e.parameter.data)   ? JSON.parse(e.parameter.data) : {};

  try {
    switch (action) {
      case 'login':           return _jsonResponse(login(data));
      case 'getPublicUsers':  return _jsonResponse(getPublicUsers());
      case 'getMyRoute':      return _jsonResponse(getMyRoute(data));
      case 'getMySalons':     return _jsonResponse(getMySalons(data));
      case 'getTodayLogs':    return _jsonResponse(getTodayLogs(data));
      case 'getVisitHistory': return _jsonResponse(getVisitHistory(data));
      case 'getUsers':        return _jsonResponse(getUsers(data));
      case 'getAllSalons':    return _jsonResponse(getAllSalons(data));
      default:                return _jsonResponse(_err('不明なアクション: ' + action));
    }
  } catch (err) {
    return _jsonResponse(_err(err.toString()));
  }
}

// ============================================================
// エントリーポイント（POST）
// application/x-www-form-urlencoded 形式で受け取る
// パラメータ: action（操作名）、data（JSONエンコードされた文字列）
// ============================================================
function doPost(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';
  const data   = params.data ? JSON.parse(params.data) : {};

  try {
    switch (action) {
      case 'checkVisit':       return _jsonResponse(checkVisit(data));
      case 'addSalon':         return _jsonResponse(addSalon(data));
      case 'updateSalon':      return _jsonResponse(updateSalon(data));
      case 'deactivateSalon':  return _jsonResponse(deactivateSalon(data));
      case 'updateSortOrder':  return _jsonResponse(updateSortOrder(data));
      case 'addUser':          return _jsonResponse(addUser(data));
      case 'updateUser':       return _jsonResponse(updateUser(data));
      case 'resetPin':         return _jsonResponse(resetPin(data));
      case 'changePin':        return _jsonResponse(changePin(data));
      case 'updateSalonAdmin': return _jsonResponse(updateSalonAdmin(data));
      case 'updateUserOrder':  return _jsonResponse(updateUserOrder(data));
      default:                 return _jsonResponse(_err('不明なアクション: ' + action));
    }
  } catch (err) {
    return _jsonResponse(_err(err.toString()));
  }
}

// ============================================================
// GETアクション実装
// ============================================================

/**
 * getPublicUsers: アクティブなユーザーの氏名・ID・チーム・ロールを返す（認証不要・ログイン/履歴フィルタ用）
 * 入力: なし
 * 出力: { users: [{ user_id, name, role, team, display_order }] }
 * display_order 列があれば昇順ソート（なければ行順）
 */
function getPublicUsers() {
  const rows  = _readSheet(SHEET_USERS);
  const users = rows
    .filter(r => r.active === true)
    .sort((a, b) => (Number(a.display_order) || 9999) - (Number(b.display_order) || 9999))
    .map(r => ({
      user_id:       r.user_id,
      name:          r.name,
      role:          r.role,
      team:          r.team,
      display_order: Number(r.display_order) || 9999
    }));
  return _ok({ users: users });
}

/**
 * login: user_id + PIN で認証し、ユーザー情報を返す
 * 入力: { user_id, pin }
 * 出力: { user_id, name, role, team }
 */
function login(data) {
  const { user_id, pin } = data;
  if (!user_id || pin === undefined || pin === null || pin === '') {
    return _err('user_idとpinは必須です');
  }

  const rows = _readSheet(SHEET_USERS);
  // Sheetsが'0000'を数値0として保存する場合があるため padStart で正規化して比較
  const user = rows.find(r =>
    r.user_id === user_id &&
    String(r.pin).padStart(4, '0') === String(pin) &&
    r.active === true
  );

  if (!user) return _err('user_idまたはPINが正しくありません');

  // PIN はレスポンスに含めない
  return _ok({
    user_id: user.user_id,
    name:    user.name,
    role:    user.role,
    team:    user.team
  });
}

/**
 * getMyRoute: 今日の曜日に該当する担当サロン一覧を返す
 * 入力: { user_id }
 * 出力: { salons: [{ salon_id, salon_name, visit_day, sort_order }] }
 */
function getMyRoute(data) {
  const { user_id } = data;
  if (!user_id) return _err('user_idは必須です');

  const today  = _todayDayName(); // 例: '月'
  const salons = _readSheet(SHEET_SALONS);

  const myRoute = salons
    .filter(s => s.owner_user_id === user_id && s.visit_day === today && s.active === true)
    .sort((a, b) => Number(a.sort_order) - Number(b.sort_order));

  return _ok({ salons: myRoute });
}

/**
 * getMySalons: 担当者の全サロンを取得（マイサロン管理画面用）
 * 入力: { user_id }
 * 出力: { salons: [...] }
 */
function getMySalons(data) {
  const { user_id } = data;
  if (!user_id) return _err('user_idは必須です');

  const dayOrder = ['月', '火', '水', '木', '金', '土', '他'];
  const salons   = _readSheet(SHEET_SALONS);

  const mySalons = salons
    .filter(s => s.owner_user_id === user_id && s.active === true)
    .sort((a, b) => {
      // 曜日順（他は末尾）→ 表示順（sort_order）
      const ai = dayOrder.indexOf(a.visit_day); const bi = dayOrder.indexOf(b.visit_day);
      const dayDiff = (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi);
      if (dayDiff !== 0) return dayDiff;
      return Number(a.sort_order) - Number(b.sort_order);
    });

  return _ok({ salons: mySalons });
}

/**
 * getTodayLogs: 担当者の当日チェック済みサロンID一覧を返す
 * 入力: { user_id }
 * 出力: { salon_ids: ['S0001', ...] }
 */
function getTodayLogs(data) {
  const { user_id } = data;
  if (!user_id) return _err('user_idは必須です');

  const today = _todayStr(); // YYYY-MM-DD
  const logs  = _readSheet(SHEET_VISIT_LOGS);

  const checkedIds = logs
    .filter(l => l.user_id === user_id && l.visit_date === today)
    .map(l => l.salon_id);

  return _ok({ salon_ids: checkedIds });
}

/**
 * getVisitHistory: 履歴マトリクス用データを返す
 * 入力: { user_ids: ['U001', ...], month: 'YYYY-MM', day_of_week: '月' or 'all' }
 * 出力: { weeks: ['3/第1月', ...], rows: [{ salon_id, salon_name, user_name, visit_day, cells: [{week_label, date, visited}], alert }] }
 */
function getVisitHistory(data) {
  const { user_ids, month, day_of_week } = data;
  if (!user_ids || !month) return _err('user_idsとmonthは必須です');

  // 対象月の日付範囲を取得
  const [year, mon] = month.split('-').map(Number);
  const startDate   = new Date(year, mon - 1, 1);
  const endDate     = new Date(year, mon, 0);  // 月末日
  const startStr    = _dateToStr(startDate);
  const endStr      = _dateToStr(endDate);

  // 各シートを読み込む
  const allSalons = _readSheet(SHEET_SALONS);
  const allUsers  = _readSheet(SHEET_USERS);
  const allLogs   = _readSheet(SHEET_VISIT_LOGS);

  // 対象ユーザーのアクティブサロンを絞り込む
  let targetSalons = allSalons.filter(s =>
    user_ids.includes(s.owner_user_id) && s.active === true
  );

  // 曜日フィルタ（指定がある場合）
  if (day_of_week && day_of_week !== 'all') {
    targetSalons = targetSalons.filter(s => s.visit_day === day_of_week);
  }

  // 対象期間のログをサロンIDでインデックス化（高速化）
  const logsBySalon = {};
  allLogs
    .filter(l => l.visit_date >= startStr && l.visit_date <= endStr && user_ids.includes(l.user_id))
    .forEach(l => {
      if (!logsBySalon[l.salon_id]) logsBySalon[l.salon_id] = new Set();
      logsBySalon[l.salon_id].add(l.visit_date);
    });

  // 月内の曜日別週ラベルを生成（例: { '月': [{label:'3/第1月', date:'2026-03-02'}, ...] }）
  const weekLabels = _generateWeekLabels(year, mon);

  const today = new Date();

  // ユーザー情報をインデックス化
  const userIndex = {};
  allUsers.forEach(u => { userIndex[u.user_id] = u; });

  // 各サロンの行データを生成
  const dayOrder = ['月', '火', '水', '木', '金', '土', '他'];
  const rows = targetSalons
    .sort((a, b) => {
      // ユーザー順 → 曜日順（他は末尾）→ sort_order順
      const userDiff = user_ids.indexOf(a.owner_user_id) - user_ids.indexOf(b.owner_user_id);
      if (userDiff !== 0) return userDiff;
      const ai = dayOrder.indexOf(a.visit_day); const bi = dayOrder.indexOf(b.visit_day);
      const dayDiff = (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi);
      if (dayDiff !== 0) return dayDiff;
      return Number(a.sort_order) - Number(b.sort_order);
    })
    .map(salon => {
      const user      = userIndex[salon.owner_user_id];
      const userName  = user ? user.name : '不明';
      const visited   = logsBySalon[salon.salon_id] || new Set();

      let cells, alert, visitCount;

      if (salon.visit_day === '他') {
        // 曜日なし：週ラベルなし・訪問回数のみ集計・アラートなし
        visitCount = [...visited].filter(d => d >= startStr && d <= endStr).length;
        cells      = [];
        alert      = false;
      } else {
        const labels = weekLabels[salon.visit_day] || [];
        visitCount   = 0;
        cells        = labels.map(wl => ({
          week_label: wl.label,
          date:       wl.date,
          visited:    visited.has(wl.date)
        }));
        // 4週連続未訪問アラート判定
        alert = _checkFourWeekAlert(salon, allLogs, today);
      }

      return {
        salon_id:    salon.salon_id,
        salon_name:  salon.salon_name,
        code:        salon.code || '',
        user_id:     salon.owner_user_id,
        user_name:   userName,
        visit_day:   salon.visit_day,
        sort_order:  Number(salon.sort_order),
        cells:       cells,
        alert:       alert,
        visit_count: visitCount  // 「他」サロン用：当月の実訪問数
      };
    });

  // ヘッダー用の週ラベル一覧（「他」は除外、表示対象曜日のみ）
  const fixedDayOrder = ['月', '火', '水', '木', '金', '土'];
  const targetDays = (day_of_week && day_of_week !== 'all')
    ? (day_of_week === '他' ? [] : [day_of_week])
    : fixedDayOrder;

  const weeks = [];
  targetDays.forEach(day => {
    (weekLabels[day] || []).forEach(wl => weeks.push(wl.label));
  });

  return _ok({ weeks: weeks, rows: rows });
}

/**
 * getUsers: ユーザー一覧取得（admin のみ）
 * 入力: { user_id }
 * 出力: { users: [...] }（PIN は含めない）
 */
function getUsers(data) {
  const { user_id } = data;
  if (!user_id) return _err('user_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser || requestUser.role !== 'admin') return _err('権限がありません');

  const rows  = _readSheet(SHEET_USERS);
  const users = rows
    .sort((a, b) => (Number(a.display_order) || 9999) - (Number(b.display_order) || 9999))
    .map(r => ({
      user_id:       r.user_id,
      name:          r.name,
      role:          r.role,
      team:          r.team,
      active:        r.active,
      display_order: Number(r.display_order) || 9999
    }));

  return _ok({ users: users });
}

/**
 * getAllSalons: 全サロン取得（admin のみ）
 * 入力: { user_id }
 * 出力: { salons: [...] }
 */
function getAllSalons(data) {
  const { user_id } = data;
  if (!user_id) return _err('user_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser || requestUser.role !== 'admin') return _err('権限がありません');

  const salons = _readSheet(SHEET_SALONS);
  return _ok({ salons: salons });
}

// ============================================================
// POSTアクション実装
// ============================================================

/**
 * checkVisit: 訪問チェックを visit_logs に記録する
 * 入力: { user_id, salon_id }
 * 出力: { log_id, visit_date } または { duplicate: true }
 */
function checkVisit(data) {
  const { user_id, salon_id } = data;
  if (!user_id || !salon_id) return _err('user_idとsalon_idは必須です');

  // 排他ロックで重複書き込みを防止
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return _err('サーバーが混み合っています。しばらくしてからお試しください。');
  }

  try {
    const today = _todayStr();
    const logs  = _readSheet(SHEET_VISIT_LOGS);

    // 同日・同ユーザー・同サロンの重複チェック
    const duplicate = logs.find(l =>
      l.user_id === user_id && l.salon_id === salon_id && l.visit_date === today
    );
    if (duplicate) return _ok({ duplicate: true, message: '本日はすでにチェック済みです' });

    // 新規ログを追記
    const sheet     = _getSheet(SHEET_VISIT_LOGS);
    const now       = new Date();
    const logId     = 'L' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMddHHmmssSSS');
    const visitedAt = Utilities.formatDate(now, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");

    sheet.appendRow([logId, visitedAt, today, user_id, salon_id]);

    return _ok({ log_id: logId, visit_date: today });
  } finally {
    lock.releaseLock();
  }
}

/**
 * addSalon: 新規サロン登録
 * 入力: { user_id, salon_name, visit_day, code? }
 * 出力: { salon_id }
 * salons シート列: salon_id(1), salon_name(2), owner_user_id(3), visit_day(4), sort_order(5), active(6), code(7)
 */
function addSalon(data) {
  const { user_id, salon_name, visit_day, code } = data;
  if (!user_id || !salon_name || !visit_day) {
    return _err('user_id、salon_name、visit_dayは必須です');
  }
  if (!DAY_MAP.hasOwnProperty(visit_day) && visit_day !== '他') {
    return _err('visit_dayは月/火/水/木/金/土/他のいずれかを指定してください');
  }

  const sheet  = _getSheet(SHEET_SALONS);
  const salons = _readSheet(SHEET_SALONS);

  // 同担当者の最大 sort_order を取得して +1
  const mySalons  = salons.filter(s => s.owner_user_id === user_id);
  const maxOrder  = mySalons.reduce((max, s) => Math.max(max, Number(s.sort_order) || 0), 0);

  // タイムスタンプベースの salon_id（重複回避）
  const salonId = 'S' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MMddHHmmssSSS');

  sheet.appendRow([salonId, salon_name, user_id, visit_day, maxOrder + 1, true, code || '']);

  return _ok({ salon_id: salonId });
}

/**
 * updateSalon: サロン情報更新（名称・曜日・表示順・コード）
 * 入力: { user_id, salon_id, salon_name?, visit_day?, sort_order?, code? }
 * 出力: { salon_id }
 * salons シート列: salon_id(1), salon_name(2), owner_user_id(3), visit_day(4), sort_order(5), active(6), code(7)
 */
function updateSalon(data) {
  const { user_id, salon_id, salon_name, visit_day, sort_order, code } = data;
  if (!user_id || !salon_id) return _err('user_idとsalon_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser) return _err('ユーザーが見つかりません');

  const sheet = _getSheet(SHEET_SALONS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === salon_id) {
      // 権限チェック：自分の担当サロンのみ変更可（admin は全件可）
      if (rows[i][2] !== user_id && requestUser.role !== 'admin') {
        return _err('自分の担当サロンのみ変更できます');
      }
      if (salon_name !== undefined) sheet.getRange(i + 1, 2).setValue(salon_name);
      if (visit_day  !== undefined) sheet.getRange(i + 1, 4).setValue(visit_day);
      if (sort_order !== undefined) sheet.getRange(i + 1, 5).setValue(Number(sort_order));
      if (code       !== undefined) sheet.getRange(i + 1, 7).setValue(code);
      return _ok({ salon_id: salon_id });
    }
  }

  return _err('サロンが見つかりません');
}

/**
 * deactivateSalon: サロンを無効化（active = FALSE）
 * 入力: { user_id, salon_id }
 * 出力: { salon_id }
 */
function deactivateSalon(data) {
  const { user_id, salon_id } = data;
  if (!user_id || !salon_id) return _err('user_idとsalon_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser) return _err('ユーザーが見つかりません');

  const sheet = _getSheet(SHEET_SALONS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === salon_id) {
      if (rows[i][2] !== user_id && requestUser.role !== 'admin') {
        return _err('自分の担当サロンのみ変更できます');
      }
      sheet.getRange(i + 1, 6).setValue(false);
      return _ok({ salon_id: salon_id });
    }
  }

  return _err('サロンが見つかりません');
}

/**
 * updateSortOrder: 表示順一括更新
 * 入力: { user_id, order_list: [{ salon_id, sort_order }, ...] }
 * 出力: { updated: N }
 */
function updateSortOrder(data) {
  const { user_id, order_list } = data;
  if (!user_id || !order_list || !Array.isArray(order_list)) {
    return _err('user_idとorder_list（配列）は必須です');
  }

  const sheet = _getSheet(SHEET_SALONS);
  const rows  = sheet.getDataRange().getValues();

  // salon_id → 行番号のインデックスを作成
  const idToRowNum = {};
  rows.forEach((row, i) => {
    if (i > 0) idToRowNum[row[0]] = i + 1;
  });

  order_list.forEach(item => {
    const rowNum = idToRowNum[item.salon_id];
    if (rowNum) {
      sheet.getRange(rowNum, 5).setValue(Number(item.sort_order));
    }
  });

  return _ok({ updated: order_list.length });
}

/**
 * addUser: ユーザー追加（admin のみ）
 * 入力: { user_id, name, pin, role, team }
 * 出力: { user_id: 新規ID }
 */
function addUser(data) {
  const { user_id, name, pin, role, team } = data;
  if (!user_id) return _err('user_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser || requestUser.role !== 'admin') return _err('権限がありません');
  if (!name || !pin || !role || !team) return _err('name、pin、role、teamは必須です');
  if (!/^\d{4}$/.test(String(pin))) return _err('PINは4桁の数字を指定してください');

  const sheet = _getSheet(SHEET_USERS);
  const users = _readSheet(SHEET_USERS);

  // タイムスタンプベースの user_id
  const newUserId = 'U' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MMddHHmmss');

  sheet.appendRow([newUserId, name, pin, role, team, true]);

  return _ok({ user_id: newUserId });
}

/**
 * updateUser: ユーザー更新（admin のみ）
 * 入力: { user_id, target_user_id, name?, role?, team?, active? }
 * 出力: { user_id: target_user_id }
 */
function updateUser(data) {
  const { user_id, target_user_id, name, role, team, active } = data;
  if (!user_id || !target_user_id) return _err('user_idとtarget_user_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser || requestUser.role !== 'admin') return _err('権限がありません');

  const sheet = _getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === target_user_id) {
      if (name   !== undefined) sheet.getRange(i + 1, 2).setValue(name);
      if (role   !== undefined) sheet.getRange(i + 1, 4).setValue(role);
      if (team   !== undefined) sheet.getRange(i + 1, 5).setValue(team);
      if (active !== undefined) sheet.getRange(i + 1, 6).setValue(active === true || active === 'true');
      return _ok({ user_id: target_user_id });
    }
  }

  return _err('ユーザーが見つかりません');
}

/**
 * resetPin: PINリセット（admin のみ、0000 に戻す）
 * 入力: { user_id, target_user_id }
 * 出力: { user_id: target_user_id }
 */
function resetPin(data) {
  const { user_id, target_user_id } = data;
  if (!user_id || !target_user_id) return _err('user_idとtarget_user_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser || requestUser.role !== 'admin') return _err('権限がありません');

  const sheet = _getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === target_user_id) {
      sheet.getRange(i + 1, 3).setValue('0000');
      return _ok({ user_id: target_user_id });
    }
  }

  return _err('ユーザーが見つかりません');
}

/**
 * changePin: 本人によるPIN変更
 * 入力: { user_id, current_pin, new_pin }
 * 出力: { user_id }
 */
function changePin(data) {
  const { user_id, current_pin, new_pin } = data;
  if (!user_id || current_pin === undefined || !new_pin) {
    return _err('user_id、current_pin、new_pinは必須です');
  }
  if (!/^\d{4}$/.test(String(new_pin))) return _err('新しいPINは4桁の数字を指定してください');

  const sheet = _getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === user_id) {
      if (String(rows[i][2]).padStart(4, '0') !== String(current_pin)) {
        return _err('現在のPINが正しくありません');
      }
      sheet.getRange(i + 1, 3).setValue(String(new_pin));
      return _ok({ user_id: user_id });
    }
  }

  return _err('ユーザーが見つかりません');
}

/**
 * updateUserOrder: ユーザー表示順の一括更新（admin のみ）
 * 入力: { user_id, order_list: [{ user_id, display_order }, ...] }
 * 出力: { updated: N }
 */
function updateUserOrder(data) {
  const { user_id, order_list } = data;
  if (!user_id || !order_list || !Array.isArray(order_list)) {
    return _err('user_idとorder_list（配列）は必須です');
  }

  const requestUser = _findUser(user_id);
  if (!requestUser || requestUser.role !== 'admin') return _err('権限がありません');

  const sheet   = _getSheet(SHEET_USERS);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];

  // display_order 列のインデックスを取得（なければエラー）
  const colIdx = headers.indexOf('display_order');
  if (colIdx === -1) return _err('display_order列が見つかりません。initDisplayOrder()を先に実行してください');

  // user_id → 行番号のインデックス
  const idToRowNum = {};
  rows.forEach((row, i) => { if (i > 0) idToRowNum[row[0]] = i + 1; });

  order_list.forEach(item => {
    const rowNum = idToRowNum[item.user_id];
    if (rowNum) sheet.getRange(rowNum, colIdx + 1).setValue(Number(item.display_order));
  });

  return _ok({ updated: order_list.length });
}

/**
 * initSalonCode: salons シートに code 列（7列目）を追加する
 * ※ 既存のシートに code 列がない場合に1回だけ手動実行する
 */
function initSalonCode() {
  const sheet   = _getSheet(SHEET_SALONS);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  const colIdx = headers.indexOf('code');
  if (colIdx !== -1) {
    Logger.log('✅ code列はすでに存在します（列' + (colIdx + 1) + '）');
    return;
  }

  // 末尾に code 列を追加
  const newColNum = headers.length + 1;
  sheet.getRange(1, newColNum).setValue('code');
  // 既存行には空文字を設定
  if (data.length > 1) {
    const emptyValues = Array(data.length - 1).fill(['']);
    sheet.getRange(2, newColNum, data.length - 1, 1).setValues(emptyValues);
  }
  Logger.log('✅ code列をsalonsシートに追加しました（列' + newColNum + '）');
}

/**
 * initDisplayOrder: users シートに display_order 列を追加して連番を設定する
 * ※ 既存のシートに display_order 列がない場合に1回だけ手動実行する
 */
function initDisplayOrder() {
  const sheet   = _getSheet(SHEET_USERS);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  let colIdx = headers.indexOf('display_order');

  if (colIdx === -1) {
    // 列が存在しない → 末尾に追加
    const newColNum = headers.length + 1;
    sheet.getRange(1, newColNum).setValue('display_order');
    for (let i = 1; i < data.length; i++) {
      sheet.getRange(i + 1, newColNum).setValue(i);
    }
    Logger.log('✅ display_order列を追加しました（' + (data.length - 1) + '件）');
  } else {
    // 列が存在する → 空のセルだけ連番で補完
    for (let i = 1; i < data.length; i++) {
      if (data[i][colIdx] === '' || data[i][colIdx] === null) {
        sheet.getRange(i + 1, colIdx + 1).setValue(i);
      }
    }
    Logger.log('✅ display_orderを補完しました');
  }
}

/**
 * updateSalonAdmin: サロン更新（admin 専用・全件対象）
 * 入力: { user_id, salon_id, salon_name?, owner_user_id?, visit_day?, active?, code? }
 * 出力: { salon_id }
 * salons シート列: salon_id(1), salon_name(2), owner_user_id(3), visit_day(4), sort_order(5), active(6), code(7)
 */
function updateSalonAdmin(data) {
  const { user_id, salon_id, salon_name, owner_user_id, visit_day, active, code } = data;
  if (!user_id || !salon_id) return _err('user_idとsalon_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser || requestUser.role !== 'admin') return _err('権限がありません');

  const sheet = _getSheet(SHEET_SALONS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === salon_id) {
      if (salon_name    !== undefined) sheet.getRange(i + 1, 2).setValue(salon_name);
      if (owner_user_id !== undefined) sheet.getRange(i + 1, 3).setValue(owner_user_id);
      if (visit_day     !== undefined) sheet.getRange(i + 1, 4).setValue(visit_day);
      if (active        !== undefined) sheet.getRange(i + 1, 6).setValue(active === true || active === 'true');
      if (code          !== undefined) sheet.getRange(i + 1, 7).setValue(code);
      return _ok({ salon_id: salon_id });
    }
  }

  return _err('サロンが見つかりません');
}

// ============================================================
// ユーティリティ関数
// ============================================================

/** JSONレスポンスを生成する */
function _jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 成功レスポンスを生成する */
function _ok(result) {
  return { status: 'ok', result: result };
}

/** エラーレスポンスを生成する */
function _err(message) {
  return { status: 'error', message: message };
}

/** シートオブジェクトを取得する */
function _getSheet(sheetName) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
}

/**
 * シートのデータをオブジェクト配列として読み込む
 * 1行目をヘッダー（フィールド名）として使用する
 */
function _readSheet(sheetName) {
  const sheet = _getSheet(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      // Date型はYYYY-MM-DD文字列に変換
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
      }
      // スプレッドシートのTRUE/FALSEを真偽値に変換
      if (val === 'TRUE'  || val === true)  val = true;
      if (val === 'FALSE' || val === false) val = false;
      obj[h] = val;
    });
    return obj;
  });
}

/** 今日の日付文字列を返す（YYYY-MM-DD） */
function _todayStr() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

/** 今日の曜日名を返す（月/火/水/木/金/土/日） */
function _todayDayName() {
  return DAY_NAMES[new Date().getDay()];
}

/** Dateオブジェクトを YYYY-MM-DD 文字列に変換する */
function _dateToStr(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}

/** user_id でアクティブなユーザーを1件取得する */
function _findUser(userId) {
  const users = _readSheet(SHEET_USERS);
  return users.find(u => u.user_id === userId && u.active === true) || null;
}

/**
 * 指定月の曜日別週ラベルを生成する
 * 例: { '月': [{label:'3/第1月', date:'2026-03-02'}, ...], ... }
 */
function _generateWeekLabels(year, month) {
  const result     = {};
  const dayNums    = { '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6 };
  const daysInMon  = new Date(year, month, 0).getDate(); // 月末日を取得

  Object.entries(dayNums).forEach(([dayName, dayNum]) => {
    result[dayName] = [];
    let count = 1;
    for (let d = 1; d <= daysInMon; d++) {
      const date = new Date(year, month - 1, d);
      if (date.getDay() === dayNum) {
        const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
        result[dayName].push({
          label: month + '/第' + count + dayName,
          date:  dateStr
        });
        count++;
      }
    }
  });

  return result;
}

/**
 * 4週連続未訪問アラートを判定する
 * 対象サロンの訪問曜日で直近4週遡り、全て未訪問なら true を返す
 */
function _checkFourWeekAlert(salon, allLogs, today) {
  const visitDayNum = DAY_MAP[salon.visit_day];
  if (visitDayNum === undefined) return false;

  // 過去の該当曜日を最大60日分遡って4件収集
  const recentDates = [];
  const checkDate   = new Date(today.getTime()); // コピーして元のDateを破壊しない

  let attempts = 0;
  while (recentDates.length < 4 && attempts < 60) {
    checkDate.setDate(checkDate.getDate() - 1);
    attempts++;
    if (checkDate.getDay() === visitDayNum) {
      recentDates.push(_dateToStr(checkDate));
    }
  }

  // 4週分揃わなければアラートなし
  if (recentDates.length < 4) return false;

  // 対象サロンのログ日付セット
  const salonLogDates = new Set(
    allLogs
      .filter(l => l.salon_id === salon.salon_id)
      .map(l => l.visit_date)
  );

  // 4週全て空白ならアラートON
  return recentDates.every(d => !salonLogDates.has(d));
}

// ============================================================
// 初回セットアップ（スクリプトエディタから一度だけ手動実行）
// ============================================================

/**
 * setupSheets: 3シートのヘッダーと初期データを一括セットアップする
 * Google Apps Script エディタの「実行」ボタンから手動で1回だけ実行してください
 */
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // users シートの作成
  _initSheet(ss, SHEET_USERS, ['user_id', 'name', 'pin', 'role', 'team', 'active']);

  // salons シートの作成
  _initSheet(ss, SHEET_SALONS, ['salon_id', 'salon_name', 'owner_user_id', 'visit_day', 'sort_order', 'active']);

  // visit_logs シートの作成
  _initSheet(ss, SHEET_VISIT_LOGS, ['log_id', 'visited_at', 'visit_date', 'user_id', 'salon_id']);

  // users に初期データを投入
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  if (usersSheet.getLastRow() === 1) {
    const initUsers = [
      ['U001', 'Takashi（管理者）', '0000', 'admin',    'all', true],
      ['U002', '統括者',           '0000', 'director', 'all', true],
      ['U003', 'A部長',            '0000', 'manager',  'A',   true],
      ['U004', 'B部長',            '0000', 'manager',  'B',   true],
      ['U005', '営業A1',           '0000', 'sales',    'A',   true],
      ['U006', '営業A2',           '0000', 'sales',    'A',   true],
      ['U007', '営業B1',           '0000', 'sales',    'B',   true],
      ['U008', '営業B2',           '0000', 'sales',    'B',   true],
    ];
    // PIN列（C列）を書式なしテキストに設定してから書き込む（'0000'が数値0に変換されるのを防止）
    usersSheet.getRange(2, 3, initUsers.length, 1).setNumberFormat('@STRING@');
    usersSheet.getRange(2, 1, initUsers.length, initUsers[0].length).setValues(initUsers);
  }

  SpreadsheetApp.flush();
  Logger.log('✅ セットアップ完了: users / salons / visit_logs シートを作成しました');
}

/** シートが存在しなければ作成し、ヘッダー行を設定する */
function _initSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#c9daf8');
    sheet.setFrozenRows(1);
  }
}
