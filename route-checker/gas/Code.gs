// ============================================================
// Beaufield ルート訪問チェッカー - Google Apps Script
// Version: 1.11.1
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   SPREADSHEET_ID : 訪問履歴データのスプレッドシートID
//   AUTH_SHEET_ID  : beaufield-auth スプレッドシートID（共通）
//
// ============================================================

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS         = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = _PROPS.getProperty('SPREADSHEET_ID');
const VERSION        = '1.11.1';

// beaufield-auth 共通認証設定
const AUTH_SHEET_ID = _PROPS.getProperty('AUTH_SHEET_ID');
const APP_NAME      = 'route-checker';

// シート名定数
const SHEET_USERS      = 'users';
const SHEET_SALONS     = 'salons';
const SHEET_VISIT_LOGS = 'visit_logs';

// 曜日マッピング（日本語曜日 → JS の getDay() 値）
const DAY_MAP   = { '日': 0, '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6 };
const DAY_NAMES = ['日', '月', '火', '水', '木', '金', '土'];

// CacheService キャッシュ時間（秒）
const CACHE_TTL_SESSION      = 900;  // セッション検証: 15分
const CACHE_TTL_PUBLIC_USERS = 600;  // ユーザー一覧: 10分

// ============================================================
// コールドスタート対策 ── Keep-Warm トリガー
// ============================================================
// GASは一定時間使われないとスクリプト環境が破棄され、
// 次回アクセス時に「コールドスタート」（15〜30秒の待ち）が発生する。
// 対策: この関数を「時間ベーストリガー」で5分ごとに実行するよう設定する。
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
  // プロパティを1つ読むだけで十分（シートへのアクセス不要）
  PropertiesService.getScriptProperties().getProperty('keepWarm');
}

// ============================================================
// セッション検証
// beaufield-auth の sessions シートでトークンを照合する
// CacheService で 10 分間キャッシュし、毎回スプレッドシートを
// 開くオーバーヘッドを回避する（高速化 v1.7.0）
// 戻り値: { valid: true, user_id } または { valid: false }
// ============================================================
function validateSession(token) {
  if (!token) return { valid: false };

  // ── キャッシュ確認（CacheService スクリプトキャッシュ・10分） ──
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'sess_' + token.slice(-32); // キー長制限対策で末尾32文字を使用
  const cached   = cache.get(cacheKey);
  if (cached !== null) {
    try {
      return JSON.parse(cached); // { valid: true, user_id } or { valid: false }
    } catch(e) { /* 壊れたキャッシュは無視してシートで再検証 */ }
  }

  // ── キャッシュなし → スプレッドシートで検証 ──
  try {
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const sh = ss.getSheetByName('sessions');
    if (!sh) return { valid: false };

    const data = sh.getDataRange().getValues();
    const now  = Date.now();

    for (let i = 1; i < data.length; i++) {
      const rowToken   = String(data[i][0]);
      const rowUserId  = String(data[i][1]);
      const rowExpires = Number(data[i][2]);

      if (rowToken === token) {
        if (rowExpires < now) {
          sh.deleteRow(i + 1);
          const result = { valid: false };
          cache.put(cacheKey, JSON.stringify(result), 60);
          return result;
        }
        const result = { valid: true, user_id: rowUserId };
        cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_SESSION); // 15分キャッシュ
        return result;
      }
    }
  } catch(e) {
    Logger.log('セッション検証エラー: ' + e);
  }

  const result = { valid: false };
  cache.put(cacheKey, JSON.stringify(result), 60); // 無効は1分だけキャッシュ
  return result;
}

// ============================================================
// エントリーポイント（GET）
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  const data   = (e && e.parameter && e.parameter.data)   ? JSON.parse(e.parameter.data) : {};
  const token  = (e && e.parameter && e.parameter.session_token) ? e.parameter.session_token : '';

  // login・getPublicUsers は認証不要（ログイン画面・ユーザー名表示用）
  const publicActions = ['login', 'getPublicUsers'];
  if (!publicActions.includes(action)) {
    const auth = validateSession(token);
    if (!auth.valid) {
      return _jsonResponse(_err('SESSION_INVALID'));
    }
  }

  try {
    switch (action) {
      case 'login':           return _jsonResponse(login(data));
      case 'getPublicUsers':  return _jsonResponse(getPublicUsers());
      case 'getMyRoute':      return _jsonResponse(getMyRoute(data));
      case 'getMySalons':     return _jsonResponse(getMySalons(data));
      case 'getTodayLogs':        return _jsonResponse(getTodayLogs(data));
      case 'getCheckScreenData':  return _jsonResponse(getCheckScreenData(data));
      case 'getVisitHistory': return _jsonResponse(getVisitHistory(data));
      case 'getSummary':     return _jsonResponse(getSummary(data));
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
  const token  = params.session_token || '';

  // 全POSTアクションでセッション検証
  const auth = validateSession(token);
  if (!auth.valid) {
    return _jsonResponse(_err('SESSION_INVALID'));
  }

  try {
    switch (action) {
      case 'checkVisit':       return _jsonResponse(checkVisit(data));
      case 'uncheckVisit':     return _jsonResponse(uncheckVisit(data));
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
 * CacheService で 10 分間キャッシュ（起動時の毎回読み込みを防ぐ）
 */
function getPublicUsers() {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'public_users_v1';
  const cached   = cache.get(cacheKey);
  if (cached) {
    try { return _ok({ users: JSON.parse(cached) }); } catch(e) { /* 壊れたキャッシュは無視 */ }
  }

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

  cache.put(cacheKey, JSON.stringify(users), CACHE_TTL_PUBLIC_USERS);
  return _ok({ users: users });
}

/**
 * login: 名前選択 + PIN で認証する（beaufield-auth を使用）
 * 入力: { user_id, pin }
 * 出力: { user_id, name, role, team }
 *
 * 認証フロー:
 *   1. beaufield-auth.users で選択した user_id の PIN を照合
 *   2. beaufield-auth.user_app_roles でこのアプリへのアクセス権を確認
 *   3. route-checker の users シートから role/team を取得（UIの権限制御に使用）
 */
function login(data) {
  const { user_id, pin } = data;
  if (!user_id || pin === undefined || pin === null || pin === '') {
    return _err('user_idとpinは必須です');
  }

  // ── ロックアウトチェック ──────────────────────────────────────
  const MAX_ATTEMPTS = 5;
  const LOCK_MINUTES = 10;
  const props    = PropertiesService.getScriptProperties();
  const lockKey  = 'lockout_' + user_id;
  const lockData = JSON.parse(props.getProperty(lockKey) || '{"count":0,"until":0}');
  const now      = Date.now();

  if (lockData.until > now) {
    const remaining = Math.ceil((lockData.until - now) / 60000);
    return _err('PINの誤入力が' + MAX_ATTEMPTS + '回に達しました。' + remaining + '分後に再試行してください。');
  }
  // ─────────────────────────────────────────────────────────────

  try {
    const authSs = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const pinStr = String(pin).padStart(4, '0');

    // Step 1: beaufield-auth.users で user_id + PIN を照合
    const authRows = authSs.getSheetByName('users').getDataRange().getValues();
    let authUser   = null;
    for (let i = 1; i < authRows.length; i++) {
      const row = authRows[i];
      // row[0]=user_id, row[2]=pin, row[3]=active
      if (String(row[0]) === user_id &&
          (row[3] === true || row[3] === 'TRUE')) {
        if (String(row[2]).padStart(4, '0') === pinStr) {
          authUser = { user_id: String(row[0]), name: String(row[1]) };
        } else {
          // PIN不一致 → 失敗回数を記録
          lockData.count = (lockData.count || 0) + 1;
          if (lockData.count >= MAX_ATTEMPTS) {
            lockData.until = now + LOCK_MINUTES * 60 * 1000;
            lockData.count = 0;
            props.setProperty(lockKey, JSON.stringify(lockData));
            return _err('PINの誤入力が' + MAX_ATTEMPTS + '回に達しました。' + LOCK_MINUTES + '分間ロックされます。');
          }
          props.setProperty(lockKey, JSON.stringify(lockData));
          const left = MAX_ATTEMPTS - lockData.count;
          return _err('PINが正しくありません（残り' + left + '回）');
        }
        break;
      }
    }
    if (!authUser) return _err('ユーザーが見つかりません');

    // Step 2: beaufield-auth.user_app_roles でアクセス権確認
    const roleRows = authSs.getSheetByName('user_app_roles').getDataRange().getValues();
    let hasAccess  = false;
    for (let i = 1; i < roleRows.length; i++) {
      if (String(roleRows[i][0]) === user_id &&
          roleRows[i][1] === APP_NAME &&
          roleRows[i][2] !== 'none') {
        hasAccess = true;
        break;
      }
    }
    if (!hasAccess) return _err('このアプリへのアクセス権限がありません');

    // Step 3: route-checker の users シートから role/team を取得
    const localUser = _findUser(user_id);
    if (!localUser) {
      return _err('ユーザー情報が見つかりません（usersシートにユーザーが登録されているか確認してください）');
    }

    // ログイン成功 → 失敗カウントをリセット
    props.deleteProperty(lockKey);

    return _ok({
      user_id: authUser.user_id,
      name:    authUser.name,
      role:    localUser.role,
      team:    localUser.team
    });

  } catch (e) {
    Logger.log('login error: ' + e);
    return _err('認証エラー: ' + e.toString());
  }
}

/**
 * getMyRoute: 今日の曜日に該当する担当サロン一覧を返す
 * 入力: { user_id }
 * 出力: { salons: [{ salon_id, salon_name, visit_day, sort_order }] }
 */
function getMyRoute(data) {
  const { user_id, date } = data;
  if (!user_id) return _err('user_idは必須です');

  const dayName = _dayNameFromDate(date || ''); // 指定日の曜日（省略時は今日）
  const salons  = _readSheet(SHEET_SALONS);

  const myRoute = salons
    .filter(s => s.owner_user_id === user_id && s.visit_day === dayName && s.active === true)
    .sort((a, b) => Number(a.sort_order) - Number(b.sort_order));

  return _ok({ salons: myRoute, day_name: dayName });
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
 * getTodayLogs: 担当者の当日チェック済みサロン一覧を返す
 * 入力: { user_id, date? }
 * 出力: { logs: [{ salon_id, time }] }  ※ time は "HH:MM" 形式
 */
function getTodayLogs(data) {
  const { user_id, date } = data;
  if (!user_id) return _err('user_idは必須です');

  const targetDate = date || _todayStr();
  const logs       = _readSheet(SHEET_VISIT_LOGS);

  const checkedLogs = logs
    .filter(l => l.user_id === user_id && l.visit_date === targetDate)
    .map(l => ({ salon_id: l.salon_id, time: _extractTime(l.visited_at) }));

  return _ok({ logs: checkedLogs });
}

/**
 * getCheckScreenData: チェック画面に必要な全データを1回のAPI呼び出しで返す（高速化用）
 * 入力: { user_id, date? }
 * 出力: { route, day_name, salon_ids, salons }
 *   route:      当日ルートのサロン一覧（sort_order順）
 *   day_name:   対象日の曜日名（例: '火'）
 *   salon_ids:  当日チェック済みのサロンID一覧
 *   salons:     担当者の全アクティブサロン（曜日→sort_order順）
 * ※ salonsシートは1回だけ読み込んでrouteとsalons両方に使う
 */
function getCheckScreenData(data) {
  const { user_id, date } = data;
  if (!user_id) return _err('user_idは必須です');

  // salonsシートを1回だけ読み込む（getMyRoute + getMySalons を統合）
  const dayName = _dayNameFromDate(date || '');
  const salons  = _readSheet(SHEET_SALONS);

  // ルート：当日曜日に一致するサロンを sort_order 順で返す
  const route = salons
    .filter(s => s.owner_user_id === user_id && s.visit_day === dayName && s.active === true)
    .sort((a, b) => Number(a.sort_order) - Number(b.sort_order));

  // 全担当サロン：曜日順 → sort_order順（マイサロン管理・イレギュラー追加用）
  const dayOrder = ['月', '火', '水', '木', '金', '土', '他'];
  const mySalons = salons
    .filter(s => s.owner_user_id === user_id && s.active === true)
    .sort((a, b) => {
      const ai = dayOrder.indexOf(a.visit_day); const bi = dayOrder.indexOf(b.visit_day);
      const dayDiff = (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi);
      if (dayDiff !== 0) return dayDiff;
      return Number(a.sort_order) - Number(b.sort_order);
    });

  // 当日チェック済みログ（salon_id と時刻を返す）
  const targetDate = date || _todayStr();
  const logs = _readSheet(SHEET_VISIT_LOGS);
  const checkedLogs = logs
    .filter(l => l.user_id === user_id && l.visit_date === targetDate)
    .map(l => ({ salon_id: l.salon_id, time: _extractTime(l.visited_at) }));

  return _ok({
    route:    route,
    day_name: dayName,
    logs:     checkedLogs,
    salons:   mySalons
  });
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

  // 対象期間のログをサロンIDでインデックス化（date → time のマップ）
  const logsBySalon = {};
  allLogs
    .filter(l => l.visit_date >= startStr && l.visit_date <= endStr && user_ids.includes(l.user_id))
    .forEach(l => {
      if (!logsBySalon[l.salon_id]) logsBySalon[l.salon_id] = {};
      logsBySalon[l.salon_id][l.visit_date] = _extractTime(l.visited_at);
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
      const visited   = logsBySalon[salon.salon_id] || {}; // { 'YYYY-MM-DD': 'HH:MM' }

      let cells, alert, visitCount;

      if (salon.visit_day === '他') {
        // 曜日なし：週ラベルなし・訪問回数のみ集計・アラートなし
        visitCount = Object.keys(visited).filter(d => d >= startStr && d <= endStr).length;
        cells      = [];
        alert      = false;
      } else {
        const labels = weekLabels[salon.visit_day] || [];
        visitCount   = 0;
        cells        = labels.map(wl => {
          const exactVisit = wl.date in visited;
          let offDate = null;
          if (!exactVisit) {
            // 同じ週（月〜日）に曜日外訪問があれば反映する
            const wr = _getWeekRange(wl.date);
            offDate = Object.keys(visited).find(d => d >= wr.start && d <= wr.end) || null;
          }
          return {
            week_label:   wl.label,
            date:         wl.date,
            visited:      exactVisit || !!offDate,
            off_schedule: !exactVisit && !!offDate,  // 曜日外訪問フラグ
            time:         exactVisit ? (visited[wl.date] || '') :
                          offDate   ? (visited[offDate]  || '') : ''
          };
        });
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
 * getSummary: 集計画面用データを返す（manager / director / admin のみ）
 * 入力: { user_id, month?: 'YYYY-MM', week_start?: 'YYYY-MM-DD'（月曜日） }
 *   - month のみ指定  → その月全体（1日〜末日）
 *   - week_start 指定 → その月曜日から6日間（月〜日）
 * 出力: { matrix: [...], alerts: [...] }
 *   matrix 要素: { user_id, user_name, days:{火:n,...}, salons:{火:n,...}, total, total_salons }
 *   alerts 要素: { user_id, user_name, salon_id, salon_name, visit_day, last_visit, weeks_since }
 */
function getSummary(data) {
  const { user_id, month, week_start } = data;
  if (!user_id) return _err('user_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser) return _err('ユーザーが見つかりません');

  const role = requestUser.role;
  if (!['sales', 'manager', 'director', 'admin'].includes(role)) return _err('権限がありません');

  // ── 集計対象の日付範囲を決定 ──
  let startStr, endStr;
  if (week_start) {
    // 週モード: week_start（月曜日）から6日間
    const parts     = week_start.split('-').map(Number);
    const startDate = new Date(parts[0], parts[1] - 1, parts[2]);
    const endDate   = new Date(startDate.getTime());
    endDate.setDate(endDate.getDate() + 6);
    startStr = _dateToStr(startDate);
    endStr   = _dateToStr(endDate);
  } else if (month) {
    // 月モード
    const [year, mon] = month.split('-').map(Number);
    startStr = _dateToStr(new Date(year, mon - 1, 1));
    endStr   = _dateToStr(new Date(year, mon, 0));
  } else {
    return _err('month または week_start のいずれかは必須です');
  }

  // ── 集計対象ユーザーを決定 ──
  const allUsers   = _readSheet(SHEET_USERS).filter(u => u.active === true);
  let   targetUsers;
  if (role === 'sales') {
    // 自分のみ
    targetUsers = [requestUser];
  } else if (role === 'manager') {
    // 自チームの sales のみ
    targetUsers = allUsers.filter(u => u.team === requestUser.team && u.role === 'sales');
  } else {
    // director / admin: 全 sales + manager
    targetUsers = allUsers.filter(u => ['sales', 'manager'].includes(u.role));
  }
  // 表示順でソート
  targetUsers.sort((a, b) => (Number(a.display_order) || 9999) - (Number(b.display_order) || 9999));

  const allSalons = _readSheet(SHEET_SALONS);
  const allLogs   = _readSheet(SHEET_VISIT_LOGS);
  const today     = new Date();

  // 対象期間の訪問済みサロンを「user_id_salon_id」セットで管理（サロンごとに1回でも訪問→カウント）
  const visitedSet = new Set(
    allLogs
      .filter(l => l.visit_date >= startStr && l.visit_date <= endStr)
      .map(l => l.user_id + '_' + l.salon_id)
  );

  const VISIT_DAYS = ['火', '水', '木', '金', '土'];

  // ── matrix を構築 ──
  const matrix = targetUsers.map(user => {
    const userSalons = allSalons.filter(s => s.owner_user_id === user.user_id && s.active === true);

    const days   = {};
    const salons = {};
    let total = 0, totalSalons = 0;

    VISIT_DAYS.forEach(d => {
      const daySalons  = userSalons.filter(s => s.visit_day === d);
      const visitCount = daySalons.filter(s => visitedSet.has(user.user_id + '_' + s.salon_id)).length;
      days[d]   = visitCount;
      salons[d] = daySalons.length;
      total       += visitCount;
      totalSalons += daySalons.length;
    });

    return {
      user_id:      user.user_id,
      user_name:    user.name,
      days:         days,
      salons:       salons,
      total:        total,
      total_salons: totalSalons
    };
  });

  // ── アラート（4週以上未訪問）: 期間に関わらず今日基準で判定 ──
  const alerts = [];
  targetUsers.forEach(user => {
    const userSalons = allSalons.filter(
      s => s.owner_user_id === user.user_id && s.active === true && s.visit_day !== '他'
    );
    userSalons.forEach(salon => {
      if (!_checkFourWeekAlert(salon, allLogs, today)) return;

      // 最終訪問日を取得
      const salonLogs = allLogs
        .filter(l => l.salon_id === salon.salon_id)
        .sort((a, b) => b.visit_date.localeCompare(a.visit_date));
      const lastVisit = salonLogs.length > 0 ? salonLogs[0].visit_date : null;

      // 最終訪問からの週数
      let weeksSince = null;
      if (lastVisit) {
        const msPerWeek = 7 * 24 * 60 * 60 * 1000;
        weeksSince = Math.floor((today.getTime() - new Date(lastVisit).getTime()) / msPerWeek);
      }

      alerts.push({
        user_id:     user.user_id,
        user_name:   user.name,
        salon_id:    salon.salon_id,
        salon_name:  salon.salon_name,
        visit_day:   salon.visit_day,
        last_visit:  lastVisit,
        weeks_since: weeksSince
      });
    });
  });

  return _ok({ matrix, alerts });
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
    const targetDate = data.date || _todayStr(); // 指定日（省略時は今日）
    const logs       = _readSheet(SHEET_VISIT_LOGS);

    // 同日・同ユーザー・同サロンの重複チェック
    const duplicate = logs.find(l =>
      l.user_id === user_id && l.salon_id === salon_id && l.visit_date === targetDate
    );
    if (duplicate) return _ok({ duplicate: true, message: '同日はすでにチェック済みです' });

    // 新規ログを追記
    const sheet     = _getSheet(SHEET_VISIT_LOGS);
    const now       = new Date();
    const logId     = 'L' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMddHHmmssSSS');
    const visitedAt = Utilities.formatDate(now, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");

    sheet.appendRow([logId, visitedAt, targetDate, user_id, salon_id]);

    return _ok({ log_id: logId, visit_date: targetDate, visited_at: visitedAt });
  } finally {
    lock.releaseLock();
  }
}

/**
 * uncheckVisit: 訪問チェックを visit_logs から削除する（取り消し）
 * 入力: { user_id, salon_id, date? }
 * 出力: { deleted: true/false }
 */
function uncheckVisit(data) {
  const { user_id, salon_id, date } = data;
  if (!user_id || !salon_id) return _err('user_idとsalon_idは必須です');

  const targetDate = date || _todayStr();

  // 排他ロックで競合書き込みを防止
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return _err('サーバーが混み合っています。しばらくしてからお試しください。');
  }

  try {
    const sheet = _getSheet(SHEET_VISIT_LOGS);
    const rows  = sheet.getDataRange().getValues();

    // 末尾から検索して最新のログを削除（同日に複数ある場合も考慮）
    // 列: log_id(0), visited_at(1), visit_date(2), user_id(3), salon_id(4)
    for (let i = rows.length - 1; i >= 1; i--) {
      // visit_date 列は Sheets が Date 型に自動変換する場合があるため文字列に正規化
      const visitDate = rows[i][2] instanceof Date
        ? Utilities.formatDate(rows[i][2], 'Asia/Tokyo', 'yyyy-MM-dd')
        : String(rows[i][2]);

      if (String(rows[i][3]) === user_id &&
          String(rows[i][4]) === salon_id &&
          visitDate === targetDate) {
        sheet.deleteRow(i + 1);
        return _ok({ deleted: true });
      }
    }
    return _ok({ deleted: false });
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

  // ユーザー変更のため public_users キャッシュを削除
  CacheService.getScriptCache().remove('public_users_v1');

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
      // ユーザー変更のため public_users キャッシュを削除
      CacheService.getScriptCache().remove('public_users_v1');
      return _ok({ user_id: target_user_id });
    }
  }

  return _err('ユーザーが見つかりません');
}

/**
 * resetPin: PINリセット（admin のみ、0000 に戻す）
 * beaufield-auth の users シートを更新する
 * 入力: { user_id, target_user_id }
 * 出力: { user_id: target_user_id }
 */
function resetPin(data) {
  const { user_id, target_user_id } = data;
  if (!user_id || !target_user_id) return _err('user_idとtarget_user_idは必須です');

  const requestUser = _findUser(user_id);
  if (!requestUser || requestUser.role !== 'admin') return _err('権限がありません');

  // beaufield-auth の users シートを更新
  const authSs = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const sheet  = authSs.getSheetByName('users');
  const rows   = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === target_user_id) {
      sheet.getRange(i + 1, 3).setValue('0000'); // C列: pin
      return _ok({ user_id: target_user_id });
    }
  }

  return _err('ユーザーが見つかりません');
}

/**
 * changePin: 本人によるPIN変更
 * beaufield-auth の users シートを更新する
 * 入力: { user_id, current_pin, new_pin }
 * 出力: { user_id }
 */
function changePin(data) {
  const { user_id, current_pin, new_pin } = data;
  if (!user_id || current_pin === undefined || !new_pin) {
    return _err('user_id、current_pin、new_pinは必須です');
  }
  if (!/^\d{4}$/.test(String(new_pin))) return _err('新しいPINは4桁の数字を指定してください');

  // beaufield-auth の users シートで現在PINを確認してから更新
  const authSs = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const sheet  = authSs.getSheetByName('users');
  const rows   = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === user_id) {
      if (String(rows[i][2]).padStart(4, '0') !== String(current_pin).padStart(4, '0')) {
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
      // Date型を文字列に変換
      // 時刻情報がある場合（visited_at等）は日時形式、なければ日付のみ（visit_date等）
      if (val instanceof Date) {
        var hasTime = val.getHours() !== 0 || val.getMinutes() !== 0 || val.getSeconds() !== 0;
        val = Utilities.formatDate(val, 'Asia/Tokyo', hasTime ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
      }
      // スプレッドシートのTRUE/FALSEを真偽値に変換
      if (val === 'TRUE'  || val === true)  val = true;
      if (val === 'FALSE' || val === false) val = false;
      obj[h] = val;
    });
    return obj;
  });
}

/**
 * visited_at（"2026-04-03T14:32:00" 形式）から "M/D HH:MM" を返す
 * 例: "2026-04-03T14:32:00" → "4/3 14:32"
 * 形式が不正な場合は空文字を返す
 */
function _extractTime(visitedAt) {
  if (!visitedAt) return '';
  const match = String(visitedAt).match(/(\d{2})-(\d{2})T(\d{2}:\d{2})/);
  if (!match) return '';
  return parseInt(match[1]) + '/' + parseInt(match[2]) + ' ' + match[3];
}

/** 今日の日付文字列を返す（YYYY-MM-DD） */
function _todayStr() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

/** 今日の曜日名を返す（月/火/水/木/金/土/日） */
function _todayDayName() {
  return DAY_NAMES[new Date().getDay()];
}

/** YYYY-MM-DD 文字列から曜日名を返す（空の場合は今日） */
function _dayNameFromDate(dateStr) {
  if (!dateStr) return _todayDayName();
  var parts = dateStr.split('-');
  var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  return DAY_NAMES[d.getDay()];
}

/** Dateオブジェクトを YYYY-MM-DD 文字列に変換する */
function _dateToStr(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}

/**
 * 指定日が含まれる週（月〜日）の開始・終了日文字列を返す
 * 例: '2026-04-08'（水）→ { start:'2026-04-06'（月）, end:'2026-04-12'（日）}
 */
function _getWeekRange(dateStr) {
  var parts = dateStr.split('-');
  var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  var day = d.getDay(); // 0=日, 1=月, ..., 6=土
  var diffToMon = (day === 0) ? -6 : 1 - day; // 月曜日へのオフセット
  var mon = new Date(d.getTime());
  mon.setDate(d.getDate() + diffToMon);
  var sun = new Date(mon.getTime());
  sun.setDate(mon.getDate() + 6);
  return { start: _dateToStr(mon), end: _dateToStr(sun) };
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

  // 今日を含む直近4回分の該当曜日を収集（今日訪問済みの場合もアラート対象外にするため今日を含める）
  const recentDates = [];
  const checkDate   = new Date(today.getTime()); // コピーして元のDateを破壊しない

  let attempts = 0;
  while (recentDates.length < 4 && attempts < 62) {
    if (checkDate.getDay() === visitDayNum) {
      recentDates.push(_dateToStr(checkDate));
    }
    checkDate.setDate(checkDate.getDate() - 1);
    attempts++;
  }

  // 4週分揃わなければアラートなし
  if (recentDates.length < 4) return false;

  // 対象サロンのログ日付セット
  const salonLogDates = new Set(
    allLogs
      .filter(l => l.salon_id === salon.salon_id)
      .map(l => l.visit_date)
  );

  // 4週全て空白ならアラートON（曜日外訪問も同じ週なら訪問済みとみなす）
  return recentDates.every(d => {
    if (salonLogDates.has(d)) return false; // 当日訪問あり
    const wr = _getWeekRange(d);
    return ![...salonLogDates].some(ld => ld >= wr.start && ld <= wr.end);
  });
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
