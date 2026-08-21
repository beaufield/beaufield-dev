/**
 * 友だち登録進捗ダッシュボード — バックエンドAPI
 *
 * 構成: GitHub Pages（index.html）+ この GAS WebApp + LINE Harness REST API
 * 認証: beaufield-auth 共通セッション。全社員閲覧可だが、activeユーザーかつ
 *       user_app_roles に line-progress が明示登録されていることを必須とする。
 * データ源: LINE Harness REST API（Bearer認証）。友だちデータはこの GAS の中だけで
 *           扱い、リポジトリにも公開HTMLのソースにも一切埋め込まない。
 *
 * スクリプトプロパティ（必須・コードに直書きしない。GASエディタ →
 * 「プロジェクトの設定」→「スクリプトプロパティ」）:
 *   AUTH_SHEET_ID         … beaufield-auth スプレッドシートのID
 *   LINE_HARNESS_API_URL  … https://bfline-harness.tak-maejima.workers.dev
 *   LINE_HARNESS_API_KEY  … LINE Harness Bearer APIキー（bcart-approval/.dev.vars と同一値）
 *   TAG_FORM              … 「顧客認証済み」タグID（①初期フォーム登録の判定に使用）
 *   TAG_CONFIRM           … 「確認番号検証済み」タグID（②確認番号の入力の判定に使用）
 *   TAG_BMALL             … 「Bモール利用中」タグID（③Bモール登録の判定に使用）
 *   TAG_STAFF             … 「社員」タグID（集計対象から除外）
 *   TAG_PARTNER           … 「協力会社」タグID（集計対象から除外）
 *   TAG_BEAUFES           … 「ビューフェス2026申込」タグID
 *   FORM_AUTH_ID           … 認証フォーム（段階1）ID
 *   FORM_BMALL_ID          … Bモール登録フォーム（段階2）ID
 *   SALES_REP_TAGS         … 🆕 営業担当タグの名前→タグID対応（JSON文字列・任意）。
 *                             例: {"井戸川":"4fa97b9c-...","中村":"42210f2c-...",
 *                                  "脇本":"569b4e18-...","嵐":"1de8a1f0-...",
 *                                  "松田":"a892abcb-...","その他":"8ca733f8-..."}
 *                             （タグIDの原本: LINEHarness/tools/tag-sales-rep/README.md）
 *                             未設定でもダッシュボードは動く（担当別分析が非表示になるだけ）。
 *                             タグは手動運用（自動トリガーなし・2026-08-16 Takashi決定）のため、
 *                             タグ付け替えのタイミングでずれることがある前提で見る。
 *
 * 参照: LINEHarness/友だち登録進捗ダッシュボード_設計.md
 */

const VERSION = '1.4.0';
const APP_NAME = 'line-progress';
const CACHE_TTL_SESSION = 60; // 権限変更・ログアウトを最大1分で反映
const CACHE_TTL_DATA = 90;     // 集計結果キャッシュ 90秒（§8-2「開くたび最新」の実務的な下限）

// プロパティ取得（未設定なら明示的にエラー。project-dashboardと同一パターン）
function prop_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('スクリプトプロパティ未設定: ' + key);
  return v;
}

// ============================================================
// セッション検証（beaufield-auth sessions シート照合・15分キャッシュ）
// is_admin は要求しないが、active とアプリ明示ロールは必須。
// ============================================================
function validateSession_(token) {
  if (!token) return { valid: false };

  const cache = CacheService.getScriptCache();
  const cacheKey = 'sess_line_progress_v2_' + token.slice(-32);
  const cached = cache.get(cacheKey);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  try {
    const ss = SpreadsheetApp.openById(prop_('AUTH_SHEET_ID'));
    const sh = ss.getSheetByName('sessions');
    if (!sh) return { valid: false };

    const data = sh.getDataRange().getValues();
    const now = Date.now();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === token) {
        if (Number(data[i][2]) < now) {
          const r = { valid: false };
          cache.put(cacheKey, JSON.stringify(r), 60);
          return r;
        }
        const userId = String(data[i][1]);
        const users = ss.getSheetByName('users').getDataRange().getValues();
        let userRow = null;
        for (let j = 1; j < users.length; j++) {
          if (String(users[j][0]) === userId) { userRow = users[j]; break; }
        }
        if (!userRow || !(userRow[3] === true || userRow[3] === 'TRUE')) return { valid: false };

        const roles = ss.getSheetByName('user_app_roles').getDataRange().getValues();
        let role = '';
        for (let j = 1; j < roles.length; j++) {
          if (String(roles[j][0]) === userId && String(roles[j][1]) === APP_NAME) {
            role = String(roles[j][2] || '').trim().toLowerCase();
            break;
          }
        }
        if (!role || role === 'none') return { valid: false };

        const r = {
          valid: true,
          user_id: userId,
          name: String(userRow[1] || userId),
          role: role
        };
        cache.put(cacheKey, JSON.stringify(r), CACHE_TTL_SESSION);
        return r;
      }
    }
  } catch (e) {
    Logger.log('セッション検証エラー: ' + e);
  }
  const r = { valid: false };
  cache.put(cacheKey, JSON.stringify(r), 60);
  return r;
}

// 認証必須アクションの共通ガード。失敗時はエラーレスポンス、成功時は null を返す
function authGuard_(token) {
  const auth = validateSession_(token);
  if (!auth.valid) {
    return jsonResponse_({
      success: false,
      error: 'SESSION_INVALID',
      message: '認証が必要です。ポータルからログインし直してください。'
    });
  }
  return null;
}

// ============================================================
// エントリーポイント（GET）: 疎通確認・ヘルスチェック用
// ============================================================
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  try {
    switch (p.action || '') {
      case 'version':
        return jsonResponse_({ success: true, version: VERSION });
      case 'health': {
        const guard = authGuard_(p.session_token || '');
        if (guard) return guard;
        return jsonResponse_(healthCheck_());
      }
      default:
        return jsonResponse_({ success: false, error: '不明なアクション' });
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return jsonResponse_({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// エントリーポイント（POST）: データ取得
// session_token をURLクエリではなくPOSTボディで送る設計（§3-2）
// ============================================================
function doPost(e) {
  let p = {};
  try {
    if (e && e.postData && e.postData.contents) p = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ success: false, error: 'BAD_JSON' });
  }

  try {
    switch (p.action || '') {
      case 'data': {
        const guard = authGuard_(p.session_token || '');
        if (guard) return guard;
        return jsonResponse_(getData_(!!p.fresh));
      }
      case 'health': {
        const guard = authGuard_(p.session_token || '');
        if (guard) return guard;
        return jsonResponse_(healthCheck_());
      }
      default:
        return jsonResponse_({ success: false, error: '不明なアクション' });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return jsonResponse_({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// LINE Harness REST API クライアント
// ============================================================
function lhFetch_(path) {
  const base = prop_('LINE_HARNESS_API_URL').replace(/\/+$/, '');
  const res = UrlFetchApp.fetch(base + path, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + prop_('LINE_HARNESS_API_KEY') },
    muteHttpExceptions: true
  });
  const status = res.getResponseCode();
  if (status !== 200) {
    throw new Error('LINE Harness API失敗 status=' + status + ' path=' + path);
  }
  return JSON.parse(res.getContentText());
}

// friends を全件取得（デフォルトlimit=50の既知の落とし穴を踏まないよう明示指定。
// ?page= は無視され offset が正、という実測済みの仕様に従う）
function fetchAllFriends_() {
  const items = [];
  let offset = 0;
  const limit = 200;
  for (let page = 0; page < 5; page++) { // 上限ガード（§3-4-1）
    const r = lhFetch_('/api/friends?limit=' + limit + '&offset=' + offset);
    const data = r.data || {};
    const pageItems = data.items || [];
    items.push.apply(items, pageItems);
    if (!data.hasNextPage) break;
    offset += limit;
  }
  return items;
}

function fetchSubmissions_(formId) {
  const r = lhFetch_('/api/forms/' + encodeURIComponent(formId) + '/submissions');
  const data = r.data;
  const items = Array.isArray(data) ? data : (data && data.items) || [];
  return items;
}

// ============================================================
// 🆕 営業担当タグ（名前→タグID）。SALES_REP_TAGS 未設定なら null を返し、
// 呼び出し側は「担当別分析なし」として静かにスキップする（必須機能にしない）。
// ============================================================
function getSalesRepTagMap_() {
  const raw = PropertiesService.getScriptProperties().getProperty('SALES_REP_TAGS');
  if (!raw) return null;
  try {
    const map = JSON.parse(raw);
    return (map && typeof map === 'object') ? map : null;
  } catch (e) {
    Logger.log('SALES_REP_TAGS のJSONパースに失敗: ' + e);
    return null;
  }
}

// ============================================================
// タグ実在チェック（§7 R2: タグを作り直すとIDが変わり集計が静かに0になる事故対策）
// ============================================================
function healthCheck_() {
  const r = lhFetch_('/api/tags');
  const ids = {};
  (r.data || []).forEach(function (t) { ids[t.id] = true; });

  const need = {
    TAG_FORM: prop_('TAG_FORM'),
    TAG_CONFIRM: prop_('TAG_CONFIRM'),
    TAG_BMALL: prop_('TAG_BMALL'),
    TAG_STAFF: prop_('TAG_STAFF'),
    TAG_PARTNER: prop_('TAG_PARTNER'),
    TAG_BEAUFES: prop_('TAG_BEAUFES')
  };
  const missing = [];
  Object.keys(need).forEach(function (key) {
    if (!ids[need[key]]) missing.push(key);
  });

  const repMap = getSalesRepTagMap_();
  if (repMap) {
    Object.keys(repMap).forEach(function (name) {
      if (!ids[repMap[name]]) missing.push('SALES_REP_TAGS:' + name);
    });
  }

  return { success: true, ok: missing.length === 0, missing: missing };
}

// ============================================================
// 集計本体（90秒キャッシュ。fresh=true でバイパス＝画面の［更新］ボタン用）
// ============================================================
function getData_(fresh) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'data_v' + VERSION;
  if (!fresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) {}
    }
  }

  const TAG_FORM = prop_('TAG_FORM');
  const TAG_CONFIRM = prop_('TAG_CONFIRM');
  const TAG_BMALL = prop_('TAG_BMALL');
  const TAG_STAFF = prop_('TAG_STAFF');
  const TAG_PARTNER = prop_('TAG_PARTNER');
  const TAG_BEAUFES = prop_('TAG_BEAUFES');
  const FORM_AUTH_ID = prop_('FORM_AUTH_ID');
  const FORM_BMALL_ID = prop_('FORM_BMALL_ID');

  // 🆕 営業担当タグ: タグID→担当名の逆引き（未設定ならnullのまま＝全員salesRep:null）
  const repMap = getSalesRepTagMap_();
  const repNameByTagId_ = {};
  if (repMap) {
    Object.keys(repMap).forEach(function (name) { repNameByTagId_[repMap[name]] = name; });
  }

  const friends = fetchAllFriends_();
  const authSubs = fetchSubmissions_(FORM_AUTH_ID);
  const bmallSubs = fetchSubmissions_(FORM_BMALL_ID);

  // friendId → 最古のcreatedAt（同一friendの重複送信は最初の1件だけを完了時刻とみなす）
  const firstByFriend_ = function (subs) {
    const m = {};
    subs.forEach(function (s) {
      if (!s.friendId) return;
      const cur = m[s.friendId];
      if (!cur || s.createdAt < cur) m[s.friendId] = s.createdAt;
    });
    return m;
  };
  const formAtMap = firstByFriend_(authSubs);
  const bmallAtMap = firstByFriend_(bmallSubs);

  const customers = [];
  const staff = [];
  const partner = [];
  const now = Date.now();

  friends.forEach(function (f) {
    const tagIds = {};
    (f.tags || []).forEach(function (t) { tagIds[t.id] = true; });
    const has_ = function (id) { return !!tagIds[id]; };
    const meta = f.metadata || {};

    // 🆕 営業担当: friendが持つタグのうちSALES_REP_TAGSに載っているものを1つ採用
    // （通常は前方一致で単一タグのみ付与される運用。複数付いていた場合はタグ配列の先頭優先）
    let salesRep = null;
    if (repMap) {
      (f.tags || []).some(function (t) {
        if (repNameByTagId_[t.id]) { salesRep = repNameByTagId_[t.id]; return true; }
        return false;
      });
    }

    const row = {
      id: f.id,
      name: meta.staff_name || f.displayName || '',
      lineName: f.displayName || '', // 🆕 LINEの表示名（初期登録フォームの氏名とは別に常時保持。登録前の旧名での識別用）
      salon: meta.salon_name || '',
      form: has_(TAG_FORM),
      confirm: has_(TAG_CONFIRM),
      bmall: has_(TAG_BMALL),
      beaufes: has_(TAG_BEAUFES),
      salesRep: salesRep,
      addedAt: f.createdAt || null,
      formAt: formAtMap[f.id] || null,
      bmallAt: bmallAtMap[f.id] || null
    };

    if (has_(TAG_STAFF)) {
      row.group = 'staff';
      staff.push(row);
    } else if (has_(TAG_PARTNER)) {
      row.group = 'partner';
      partner.push(row);
    } else {
      row.group = 'customer';
      // 停滞日数: 最後に進んだステップの日時からの経過日数（未完了者向け表示）
      const lastProgressAt = row.bmallAt || row.formAt || row.addedAt;
      row.stalledDays = lastProgressAt
        ? Math.floor((now - new Date(lastProgressAt).getTime()) / 86400000)
        : null;
      customers.push(row);
    }
  });

  const series = buildSeries_(customers);

  // ③Bモール登録の一部は、このフォームを経由せず直接タグ付与されているケースがある
  // （bcart-approval の承認フローが即タグを付ける経路など）。series はフォーム送信記録
  // に基づく実測のため、bmall=true だが bmallAt が無い＝series には出てこない人数を
  // 明示してフロントで注記表示できるようにする（数字を黙って小さく見せない・§7 R8方針）
  const bmallUnmatchedCount = customers.filter(function (c) {
    return c.bmall && !c.bmallAt;
  }).length;

  const result = {
    success: true,
    synced_at: new Date().toISOString(),
    customers: customers,
    staff: staff,
    partner: partner,
    series: series,
    bmallUnmatchedCount: bmallUnmatchedCount,
    // 🆕 SALES_REP_TAGS のキー順（=タグ定義の表示順）。未設定なら空配列＝フロントは
    // 担当別分析セクションを表示しない
    salesReps: repMap ? Object.keys(repMap) : []
  };

  cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_DATA);
  return result;
}

// ============================================================
// 日次推移シリーズ（友だち数・①初期フォーム・③Bモール登録の累積）
// ②確認番号は完了時刻の記録が無いため含めない（§4-4・§7 R8）
//
// 🆕 customers（=集計対象の顧客。1人1行に既に重複排除済み）の formAt/bmallAt から
// 組み立てる。フォーム送信記録を生のまま数えると「同一人物の再送信」や「社員の送信」が
// 混入し、統計タイルの人数と一致しなくなるため使わない（実測で発覚: 生の送信件数は
// 認証140件/Bモール93件だが、重複排除した顧客の人数は111人/最大82人）
// ============================================================
function buildSeries_(customers) {
  // LINE HarnessのcreatedAtは +09:00 付きのISO文字列。日付部分だけ切り出す
  const toDateStr_ = function (iso) { return String(iso).slice(0, 10); };

  const addedDates = customers
    .map(function (c) { return c.addedAt ? toDateStr_(c.addedAt) : null; })
    .filter(Boolean);
  const formDates = customers
    .map(function (c) { return c.formAt ? toDateStr_(c.formAt) : null; })
    .filter(Boolean);
  const bmallDates = customers
    .map(function (c) { return c.bmallAt ? toDateStr_(c.bmallAt) : null; })
    .filter(Boolean);

  const allDates = addedDates.concat(formDates, bmallDates);
  if (!allDates.length) return { dates: [], friends: [], form: [], bmall: [] };

  allDates.sort();
  const start = allDates[0];
  // "今日"はJSTで判定する（toISOString()はUTC変換されるため、JST早朝(0-9時)に実行すると
  // 前日扱いになり当日の列が抜け落ちるバグを踏む。Utilities.formatDateで明示的にJST化する）
  const end = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  // 欠測日も0埋めして連続させる（§3-4-5: 日付が飛ぶと折れ線が嘘をつくため）
  const dates = [];
  let cur = new Date(start + 'T00:00:00+09:00');
  const endDate = new Date(end + 'T00:00:00+09:00');
  while (cur.getTime() <= endDate.getTime()) {
    dates.push(Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy-MM-dd'));
    cur = new Date(cur.getTime() + 86400000);
  }

  const countByDate_ = function (dateList) {
    const perDay = {};
    dateList.forEach(function (d) { perDay[d] = (perDay[d] || 0) + 1; });
    let running = 0;
    return dates.map(function (d) {
      running += (perDay[d] || 0);
      return running;
    });
  };

  return {
    dates: dates,
    friends: countByDate_(addedDates),
    form: countByDate_(formDates),
    bmall: countByDate_(bmallDates)
  };
}

// ============================================================
// ヘルパー
// ============================================================
function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// keepWarm: GASのコールドスタートを防ぐ定期実行用関数（任意・トリガー未設定なら何もしない）
function keepWarm() {}
