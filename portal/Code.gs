// ============================================================
// Beaufield ポータル - Google Apps Script
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   AUTH_SHEET_ID : beaufield-auth スプレッドシートID（共通）
//
// ============================================================

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS        = PropertiesService.getScriptProperties();
const VERSION       = 'v1.14.0';
// ポータル画面（GitHub Pages）のURL。旧HTML向けの更新案内タイルのリンク先に使う。
const PORTAL_URL    = 'https://beaufield.github.io/beaufield-dev/';
const AUTH_SHEET_ID = _PROPS.getProperty('AUTH_SHEET_ID');

// ロックアウト設定（PIN総当たり対策・2026-10-09 セキュリティレビューH1）
// 誤入力が MAX_ATTEMPTS 回たまるごとにロックし、ロックのたびに時間を延ばす。
// 失敗回数はロック明けでも戻さず、ログインに成功したとき（または管理者のPINリセット時）だけ0に戻す。
// 旧方式（5回で10分・ロック明けに0から再開）では1日で全社員のうち誰かのPINが当たる確率が約5割あった。
const MAX_ATTEMPTS       = 5;
const LOCK_STEPS_MINUTES = [10, 60, 1440]; // 1回目10分 → 2回目60分 → 3回目以降は毎回24時間

// 全ユーザー合計の失敗回数がこの時間内にこの回数へ達したら管理者へ通知する（総当たりの早期発見用）
const GLOBAL_FAIL_ALERT_COUNT = 30;
const GLOBAL_FAIL_WINDOW_SEC  = 3600;

// 通知先: スクリプトプロパティ LOGIN_ALERT_WEBHOOK（LINE WORKS Incoming Webhook のURL）。
// 未設定なら通知はせずログに残すだけ（ログイン処理には影響しない）。
// ⚠️ 全社員向けの通知枠は使わないこと（管理者だけが見る枠を設定する）。

// セッション有効期間。共有端末でのトークン残存を抑えつつ、頻繁な再ログインを避けるため7日とする。
const SESSION_HOURS = 24 * 7;

// ============================================================
// アプリマスター
// appName は beaufield-auth の user_app_roles シートの値と一致させる
// ============================================================
const APP_MASTER = [
  {
    appName: 'order-app',
    label:   '発注アプリ',
    icon:    '📦',
    url:     'https://beaufield.github.io/beaufield-dev/order-app/'
  },
  {
    appName: 'lending',
    label:   '貸出管理',
    icon:    '🔑',
    url:     'https://beaufield.github.io/beaufield-dev/kiki-kanri/'
  },
  {
    appName: 'serial-apps',
    label:   'シリアルNo管理',
    icon:    '🏷️',
    url:     'https://beaufield.github.io/beaufield-dev/serial-apps/'
  },
  {
    appName: 'yoyaku-kanri',
    label:   '予約管理',
    icon:    '📋',
    url:     'https://beaufield.github.io/beaufield-dev/yoyaku-kanri/'
  },
  {
    appName: 'bcart-master',
    label:   'BCARTマスター管理',
    icon:    '🛒',
    url:     'https://beaufield.github.io/beaufield-dev/bcart-integration/master-tool/'
  },
  {
    appName: 'bcart-orders',
    label:   'Bカート受注確認',
    icon:    '🧾',
    url:     'https://beaufield.github.io/beaufield-dev/bcart-integration/order-viewer/'
  },
  {
    appName: 'expense-approval',
    label:   '経費事前申請',
    icon:    '💰',
    url:     'https://beaufield.github.io/beaufield-dev/expense-approval/'
  },
  {
    appName: 'project-dashboard',
    label:   'プロジェクトダッシュボード',
    icon:    '📊',
    url:     'https://beaufield.github.io/beaufield-dev/project-dashboard/'
  },
  {
    appName: 'line-progress',
    label:   'LINE登録進捗',
    icon:    '📈',
    url:     'https://beaufield.github.io/beaufield-dev/line-progress/'
  }
];

// ============================================================
// アプリマスターの読み込み（beaufield-auth の apps シートを正本とする）
// ============================================================
// 2026-08-21〜: アプリ一覧の原本を APP_MASTER（コード直書き）から
// beaufield-auth の apps シートへ移設。以後のアプリ追加・並べ替え・非表示は
// シート編集だけで完結し、ポータルの再デプロイが不要になる。
// APP_MASTER は下記のとおり三重フォールバックの安全弁として削除せず残す
// （シート未作成／空／例外時は自動的に現状のAPP_MASTERへ戻る）。
// キャッシュは60秒（他アプリのロールキャッシュと同じ粒度）。
function _loadAppMaster() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('app_master_v1');
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }
  try {
    const sh = SpreadsheetApp.openById(AUTH_SHEET_ID).getSheetByName('apps');
    if (!sh || sh.getLastRow() < 2) return APP_MASTER; // シート未作成 → 現状維持
    const rows = sh.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < rows.length; i++) {
      const appName = String(rows[i][0] || '').trim();       // A: app_name
      const label   = String(rows[i][1] || '').trim();       // B: label
      const url     = String(rows[i][3] || '').trim();       // D: url
      const status  = String(rows[i][7] || 'active').trim(); // H: status
      if (!appName || !label || !url) continue;              // 不正行はスキップ
      if (status !== 'active') continue;                     // hidden / archived は出さない
      list.push({
        appName: appName,
        label:   label,
        icon:    String(rows[i][2] || '📱'),                 // C: icon
        url:     url,
        sort:    Number(rows[i][6]) || 999                   // G: sort_order
      });
    }
    if (!list.length) return APP_MASTER;                     // 全滅 → 現状維持
    list.sort((a, b) => a.sort - b.sort);
    cache.put('app_master_v1', JSON.stringify(list), 60);
    return list;
  } catch (e) {
    Logger.log('_loadAppMaster failed: ' + e);
    return APP_MASTER;                                       // 例外 → 現状維持
  }
}

// ============================================================
// エントリーポイント（GET）
// ============================================================
// getUsers    : ログイン画面のユーザー一覧。認証前に必要なため現行フロントもGETで呼ぶ。
// getUserApps : 現行フロント（HTML v1.3.0以降）はPOSTで呼ぶ。ここへGETで来るのは
//               v1.2.1以前のHTMLを開いたまま放置している端末だけ。
//               アプリ一覧は返さず、更新案内タイルを返す（_legacyUpdateNotice 参照）。
function doGet(e) {
  const action = e && e.parameter && e.parameter.action ? e.parameter.action : '';

  try {
    switch (action) {
      case 'getUsers':    return _json(getUsers());
      case 'getUserApps': return _json(_legacyUpdateNotice());
      default:            return _json({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return _json({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// 旧HTML（v1.2.1以前）向けの更新案内タイル
// ============================================================
// 背景:
//   2026-08-12にgetUserAppsをPOST専用化したところ、ポータルを開いたまま放置していた
//   端末が古いJSでGETを呼び続け、画面に「利用可能なアプリがありません／管理者に
//   お問い合わせください」とだけ出た（2026-08-13の障害）。旧HTMLはサーバーが返す
//   error文字列を読まずに握りつぶすため、理由を伝える手段がない。
// 対策:
//   旧HTMLは apps の各要素をタイルとして描画する。そこで一覧の代わりに案内タイルを
//   1枚返し、利用者がタップするだけで新しいHTMLを取得できるようにする。
//   URLに ?v= を付けるのは、ブラウザキャッシュ（GitHub Pagesは max-age=600）を
//   確実に回避して必ず最新HTMLを取りに行かせるため。
// 撤去条件:
//   全13名の端末がポータルv1.3.0以降になったことを確認したら、doGetの
//   case 'getUserApps' ごと削除してよい（残っていても実害はない）。
// ⚠️ 本番確認時の注意:
//   GET getUserApps は `USE_POST` ではなく、この案内（success:true・タイル1枚）を
//   返すのが正常。8/13以前のドキュメントには「USE_POSTが正常」と書かれているので注意。
function _legacyUpdateNotice() {
  return {
    success: true,
    apps: [{
      appName: '_portal_update',
      label:   '更新が必要です（タップしてください）',
      icon:    '🔄',
      url:     PORTAL_URL + '?v=' + VERSION.replace(/[^0-9]/g, '')
    }]
  };
}

// ============================================================
// エントリーポイント（POST）
// URL-encoded と JSON body の両方に対応
// ============================================================
// ============================================================
// 冪等キー（requestId）による二重実行防止
// ============================================================
// GASのWebアプリは二段構え:
//   ① script.google.com/.../exec        … ここでスクリプトが実行され、302が返る
//   ② script.googleusercontent.com/...  … ここで結果だけが配信される
//
// ②のURLは【使い捨て】かつ【寿命が5〜30秒の間】で、切れると①へ差し戻される。
// その差し戻しはGETなのでPOSTのボディが消え、クライアントには
// 「不明なアクション: 」が返る。
// ⚠️ **このとき①の処理は既に実行済み**（2026-09-19にカウンタで実測確認）。
//    つまりクライアントから見た「失敗」は、実際には「成功したが答えを受け取れなかった」。
//    iPhoneで302の直後にアプリを切り替えるとタブが止まるため、特に踏みやすい。
//
// したがってクライアントが同じ requestId で再送してきたら、処理をやり直さず
// 前回の結果をそのまま返す。これで再送が安全になる。
// 検証の詳細: portal/PLAN-stuck-spinner-fix.md
const IDEMPOTENCY_TTL_SEC = 600;  // 10分。再送は数秒〜数十秒以内に来る

function _withIdempotency(requestId, fn, action, data) {
  const rid = String(requestId || '').trim();
  if (rid && !/^[A-Za-z0-9-]{8,100}$/.test(rid)) return {success: false, error: 'INVALID_REQUEST'};
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return {success: false, error: 'BUSY'}; }
  try {
    // 旧画面のキーなし操作もロックする。従来互換は保持するが、安全な再送は保証しない。
    if (!rid) return fn();
    const props = PropertiesService.getScriptProperties();
    const prefix = 'portal_idem_v2_';
    const digest = function(value) {
      return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
        .map(function(b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
    };
    // キー、操作、利用者・セッション、内容を結び付ける。PINやトークンそのものは保存しない。
    const canonical = function(value) {
      if (Array.isArray(value)) return value.map(canonical);
      if (value && typeof value === 'object') {
        const sorted = {};
        Object.keys(value).sort().forEach(function(k) { if (k !== 'requestId') sorted[k] = canonical(value[k]); });
        return sorted;
      }
      return value;
    };
    const fingerprint = digest(JSON.stringify([action || '', canonical(data || {})]));
    const key = prefix + digest(rid);
    const now = Date.now();
    const raw = props.getProperty(key);
    if (raw) {
      const entry = JSON.parse(raw);
      if (entry.expiresAt > now) {
        if (entry.fingerprint !== fingerprint) return {success: false, error: 'REQUEST_CONFLICT'};
        if (entry.state !== 'done') return {success: false, error: 'RESULT_UNKNOWN'};
        // ログアウト・利用停止後に、過去のログイン結果から失効トークンを復活させない。
        if (action === 'login' && entry.result.success &&
            !validateSession({token: entry.result.session_token}).ok) {
          return {success: false, error: 'SESSION_INVALID'};
        }
        return Object.assign({}, entry.result, {replayed: true});
      }
    }
    // 保存容量を制限する。処理中の記録を追い出して再実行を許可しない。
    const all = props.getProperties(); let live = 0;
    Object.keys(all).filter(function(k) { return k.indexOf(prefix) === 0; }).forEach(function(k) {
      const entry = JSON.parse(all[k]);
      if (entry.expiresAt <= now) props.deleteProperty(k); else live++;
    });
    if (live >= 256) return {success: false, error: 'BUSY'};
    // 更新前に予約を永続保存。更新後に応答の保存が失敗した場合は、不明扱いで再更新を止める。
    const entry = {fingerprint: fingerprint, state: 'pending', expiresAt: now + 86400000};
    props.setProperty(key, JSON.stringify(entry));
    const result = fn();
    entry.state = 'done'; entry.result = result;
    entry.expiresAt = Date.now() + IDEMPOTENCY_TTL_SEC * 1000;
    props.setProperty(key, JSON.stringify(entry));
    return result;
  } finally { lock.releaseLock(); }
}

function doPost(e) {
  let action = '', data = {};

  try {
    // JSON bodyの場合（他のGASアプリからの内部呼び出し）
    if (e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      if (body.action) { action = body.action; data = body; }
    }
  } catch(err) {}

  // URL-encodedの場合（ポータルHTMLからの呼び出し）
  if (!action && e && e.parameter) {
    action = e.parameter.action || '';
    try {
      data = e.parameter.data ? JSON.parse(e.parameter.data) : {};
    } catch(jsonErr) {
      return _json({ success: false, error: 'INVALID_REQUEST' });
    }
  }

  try {
    switch (action) {
      // 書き込みは冪等キーで包む。同じ requestId の再送は前回の結果を返すだけになる
      case 'login':           return _json(_withIdempotency(data.requestId, function(){ return login(data); }, 'login', data));
      case 'resetPin':        return _json(_withIdempotency(data.requestId, function(){ return resetPin(data); }, 'resetPin', data));
      case 'changePin':       return _json(_withIdempotency(data.requestId, function(){ return changePin(data); }, 'changePin', data));
      case 'logout':          return _json(logout(data));
      case 'validateSession': return _json(validateSession(data));
      case 'getUserApps':     return _json(getUserApps(data.session_token || ''));
      default:                return _json({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return _json({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// ユーザー一覧取得（ログイン画面のグリッド表示用）
// ============================================================
function getUsers() {
  const ss   = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const rows = ss.getSheetByName('users').getDataRange().getValues();
  const users = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[3] === true || row[3] === 'TRUE') {
      users.push({
        user_id:  String(row[0]),
        name:     String(row[1])
      });
    }
  }
  return { success: true, users };
}

// ============================================================
// ログイン処理（ロックアウト付き）
// ============================================================
function login(data) {
  const { user_id, pin } = data;
  if (!user_id || pin === undefined || pin === null || pin === '') {
    return { success: false, message: 'user_idとpinは必須です' };
  }

  // ── ロックアウトチェック ──────────────────────────────────
  const props    = PropertiesService.getScriptProperties();
  const lockKey  = 'lockout_' + user_id;
  const lockData = _readLockData(props, lockKey);
  const now      = Date.now();

  if (lockData.until > now) {
    return {
      success: false,
      message: 'PINの誤入力が続いたためロックしています。' + _formatWait(lockData.until - now) +
               '後に再試行してください。急ぎの場合は管理者にPINリセットを依頼してください。'
    };
  }
  // ─────────────────────────────────────────────────────────

  const ss     = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const rows   = ss.getSheetByName('users').getDataRange().getValues();
  const pinStr = String(pin).padStart(4, '0');

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[0]) === user_id && (row[3] === true || row[3] === 'TRUE')) {
      if (String(row[2]).padStart(4, '0') === pinStr) {
        // ログイン成功 → 失敗回数をリセット（段階ロックの段階も最初に戻る）
        props.deleteProperty(lockKey);

        // ── セッショントークン発行 ────────────────────────────
        const token     = Utilities.getUuid();
        const expiresAt = now + SESSION_HOURS * 60 * 60 * 1000;
        _saveSession(ss, token, String(row[0]), expiresAt);
        // ────────────────────────────────────────────────────

        // is_admin: F列（row[5]）がTRUEかどうか
        const isAdmin = row[5] === true || row[5] === 'TRUE';

        return {
          success:       true,
          user_id:       String(row[0]),
          name:          String(row[1]),
          session_token: token,
          is_admin:      isAdmin
        };
      } else {
        // PIN不一致 → 失敗回数を記録（ロック明けでも0に戻さない）
        _countGlobalFailure();
        lockData.fails = lockData.fails + 1;
        if (lockData.fails % MAX_ATTEMPTS === 0) {
          const lockMinutes = _lockMinutesFor(lockData.fails);
          lockData.until = now + lockMinutes * 60 * 1000;
          props.setProperty(lockKey, JSON.stringify(lockData));
          _sendLoginAlert(
            '🔒 ポータル: PIN誤入力によるロック\n' +
            '対象: ' + String(row[1]) + '（' + String(row[0]) + '）\n' +
            '連続失敗: ' + lockData.fails + '回 → ' + _formatWait(lockMinutes * 60 * 1000) + 'ロック\n' +
            '本人の入力ミスでなければ、総当たり攻撃の可能性があります。'
          );
          return {
            success: false,
            message: 'PINの誤入力が続いたため' + _formatWait(lockMinutes * 60 * 1000) + 'ロックします。' +
                     '急ぎの場合は管理者にPINリセットを依頼してください。'
          };
        }
        props.setProperty(lockKey, JSON.stringify(lockData));
        const left = MAX_ATTEMPTS - (lockData.fails % MAX_ATTEMPTS);
        return { success: false, message: 'PINが正しくありません（あと' + left + '回でロック）' };
      }
    }
  }
  // 存在しない・無効化済みのuser_idへの試行も、全体の失敗回数には数える
  _countGlobalFailure();
  return { success: false, message: 'ユーザーが見つかりません' };
}

// ============================================================
// ロックアウト補助（v1.14.0〜 段階ロック）
// ============================================================

// ロック情報を読む。形式: { fails: 最後の成功以降の失敗回数, until: ロック解除時刻(ms) }
// 旧形式 { count, until } も読める（count をそのまま fails として引き継ぐ）。
// 壊れた値でログインできなくなる事故を避けるため、読めなければ初期値に戻す。
function _readLockData(props, lockKey) {
  try {
    const raw = JSON.parse(props.getProperty(lockKey) || '{}');
    const fails = Number(raw.fails !== undefined ? raw.fails : raw.count) || 0;
    const until = Number(raw.until) || 0;
    return { fails: Math.max(0, Math.floor(fails)), until: until };
  } catch (e) {
    Logger.log('_readLockData: 壊れたロック情報を初期化します key=' + lockKey);
    return { fails: 0, until: 0 };
  }
}

// 何回目のロックかに応じたロック時間（分）。5回目=1段目、10回目=2段目、15回目以降=最終段
function _lockMinutesFor(fails) {
  const stage = Math.floor(fails / MAX_ATTEMPTS); // 1, 2, 3, ...
  const idx = Math.min(stage, LOCK_STEPS_MINUTES.length) - 1;
  return LOCK_STEPS_MINUTES[Math.max(0, idx)];
}

// 残り時間を画面表示用の文字列にする（60分未満は「N分」、それ以上は「約N時間」）
function _formatWait(ms) {
  const minutes = Math.max(1, Math.ceil(ms / 60000));
  if (minutes < 60) return minutes + '分';
  return '約' + Math.ceil(minutes / 60) + '時間';
}

// 全ユーザー合計の失敗回数を1時間単位で数え、閾値に達した時点で1回だけ通知する。
// 呼び出し元の login は _withIdempotency のスクリプトロック内で動くため、加算は直列化されている。
// CacheService は消えることがあるが、用途は「気づくための通知」なので数え漏れは許容する。
function _countGlobalFailure() {
  try {
    const cache = CacheService.getScriptCache();
    const key = 'login_fail_h_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHH');
    const count = (Number(cache.get(key)) || 0) + 1;
    cache.put(key, String(count), GLOBAL_FAIL_WINDOW_SEC + 600);
    if (count === GLOBAL_FAIL_ALERT_COUNT) {
      _sendLoginAlert(
        '⚠️ ポータル: ログイン失敗が多発しています\n' +
        'この1時間（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d H') + '時台）の失敗が' +
        GLOBAL_FAIL_ALERT_COUNT + '回に達しました（全ユーザー合計）。\n' +
        '総当たり攻撃の可能性があります。'
      );
    }
  } catch (e) {
    Logger.log('_countGlobalFailure error: ' + e);
  }
}

// GASエディタから手動で1回だけ実行する（v1.14.0の反映時）。
// 目的: ①UrlFetchApp（外部通信）の権限をこのプロジェクトで承認する ②通知先と文面を確かめる。
// 送り先は LOGIN_ALERT_WEBHOOK（管理者個別の枠）。⚠️ 全社員向けの枠が設定されていないことを確認してから実行する。
// 関数名の末尾に _ を付けない（付けるとエディタの実行メニューに出ない）。
function testLoginAlert() {
  if (!_PROPS.getProperty('LOGIN_ALERT_WEBHOOK')) {
    Logger.log('LOGIN_ALERT_WEBHOOK が未設定です。通知は送られません。');
    return;
  }
  _sendLoginAlert('🧪 ポータル: ログイン監視通知のテストです（' + VERSION + '）。\nこのメッセージが届けば通知の設定は完了です。');
  Logger.log('テスト通知を送信しました。LINE WORKSで届いたか確認してください。');
}

// 管理者向けのLINE WORKS通知。失敗してもログイン処理には影響させない。
// ⚠️ PINやセッショントークンは本文に入れないこと。
function _sendLoginAlert(text) {
  Logger.log('login alert: ' + text);
  const url = _PROPS.getProperty('LOGIN_ALERT_WEBHOOK');
  if (!url) return;
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ body: { text: text } }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('_sendLoginAlert error: ' + e);
  }
}

// ============================================================
// PINリセット（管理者専用）
// ============================================================
function resetPin(data) {
  const { session_token, target_user_id, new_pin } = data;

  // 必須チェック
  if (!session_token || !target_user_id || !new_pin) {
    return { success: false, message: '必須パラメータが不足しています' };
  }

  // PIN形式チェック
  const pinStr = String(new_pin).padStart(4, '0');
  if (!/^\d{4}$/.test(pinStr)) {
    return { success: false, message: 'PINは4桁の数字で入力してください' };
  }

  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

  // セッション検証（誰が操作しているか）
  const requestUserId = _getSessionUser(ss, session_token);
  if (!requestUserId) {
    return { success: false, message: 'セッションが無効です。再ログインしてください' };
  }

  // 現在も有効な管理者か確認する（発行済みセッションだけでは許可しない）。
  if (!_isAdmin(ss, requestUserId)) {
    return { success: false, message: '管理者権限がありません' };
  }

  // 対象ユーザーのPINを更新
  const sh   = ss.getSheetByName('users');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === target_user_id) {
      sh.getRange(i + 1, 3).setValue(pinStr); // C列（PIN）を更新
      // ロックアウトも同時に解除
      PropertiesService.getScriptProperties().deleteProperty('lockout_' + target_user_id);
      // 対象ユーザーの既存セッションをすべて削除（旧PINでのトークンを無効化）
      _deleteUserSessions(ss, target_user_id);
      return { success: true, message: 'PINを更新しました' };
    }
  }
  return { success: false, message: '対象ユーザーが見つかりません' };
}

// ============================================================
// PIN変更（本人による変更）
// ============================================================
function changePin(data) {
  const { session_token, current_pin, new_pin } = data;

  if (!session_token || current_pin === undefined || current_pin === null || !new_pin) {
    return { success: false, message: '必須パラメータが不足しています' };
  }

  const newPinStr = String(new_pin).padStart(4, '0');
  if (!/^\d{4}$/.test(newPinStr)) {
    return { success: false, message: 'PINは4桁の数字で入力してください' };
  }

  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

  // セッション検証
  const userId = _getSessionUser(ss, session_token);
  if (!userId) {
    return { success: false, message: 'セッションが無効です。再ログインしてください' };
  }

  // 現在のPINと照合してから更新
  const sh            = ss.getSheetByName('users');
  const rows          = sh.getDataRange().getValues();
  const currentPinStr = String(current_pin).padStart(4, '0');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === userId) {
      // セッション発行後に無効化されていたら、PIN照合・書込の前に拒否する。
      if (!(rows[i][3] === true || rows[i][3] === 'TRUE')) {
        return { success: false, message: 'セッションが無効です。再ログインしてください' };
      }
      if (String(rows[i][2]).padStart(4, '0') !== currentPinStr) {
        return { success: false, message: '現在のPINが正しくありません' };
      }
      sh.getRange(i + 1, 3).setValue(newPinStr);
      // PIN変更後は旧PINで発行済みの全セッションを失効させる。
      _deleteUserSessions(ss, userId);
      return { success: true, message: 'PINを変更しました。すべての端末で再ログインしてください。', reauth_required: true };
    }
  }
  return { success: false, message: 'ユーザーが見つかりません' };
}

// ============================================================
// アクセス可能アプリ一覧取得（セッション必須）
// ============================================================
function getUserApps(token) {
  if (!token) return { success: false, error: 'SESSION_INVALID' };

  const ss     = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const userId = _getSessionUser(ss, token);
  if (!userId) return { success: false, error: 'SESSION_INVALID' };

  // セッション発行後に無効化されたユーザーは、一覧取得でも拒否する。
  const users = ss.getSheetByName('users').getDataRange().getValues();
  let activeUser = false;
  for (let i = 1; i < users.length; i++) {
    if (String(users[i][0]) === userId) {
      activeUser = users[i][3] === true || users[i][3] === 'TRUE';
      break;
    }
  }
  if (!activeUser) return { success: false, error: 'SESSION_INVALID' };

  const roles = ss.getSheetByName('user_app_roles').getDataRange().getValues();

  const accessMap = {};
  for (let i = 1; i < roles.length; i++) {
    if (String(roles[i][0]) === userId && roles[i][2] !== 'none') {
      accessMap[String(roles[i][1])] = String(roles[i][2]);
    }
  }

  const apps = _loadAppMaster()
    .filter(app => accessMap[app.appName])
    .map(app => ({
      appName: app.appName,
      label:   app.label,
      icon:    app.icon,
      url:     app.url,
      role:    accessMap[app.appName]
    }));

  return { success: true, apps };
}

// ============================================================
// セッション検証（他のGASアプリからの内部呼び出し用）
// 戻り値: { ok: true/false } ※ success ではなく ok で返す
// ============================================================
function validateSession(data) {
  const token = data.token || '';
  const appName = String(data.app_name || '').trim();
  if (!token) return { ok: false };

  const ss     = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const userId = _getSessionUser(ss, token);
  if (!userId) return { ok: false };

  // ユーザー名・有効状態・管理者状態を users シートから取得する。
  // セッション発行後に利用停止されたユーザーも、ここで必ず拒否する。
  const rows = ss.getSheetByName('users').getDataRange().getValues();
  let userName = userId;
  let isAdmin = false;
  let activeUser = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === userId) {
      userName = String(rows[i][1]) || userId;
      activeUser = rows[i][3] === true || rows[i][3] === 'TRUE';
      isAdmin = rows[i][5] === true || rows[i][5] === 'TRUE';
      break;
    }
  }
  if (!activeUser) return { ok: false };

  // 呼び出し側がアプリ名を指定した場合は、ポータル表示だけでなくAPI側でも
  // user_app_rolesを検証する。未登録・noneは明示的に拒否する。
  let role = '';
  if (appName) {
    const roleRows = ss.getSheetByName('user_app_roles').getDataRange().getValues();
    for (let i = 1; i < roleRows.length; i++) {
      if (String(roleRows[i][0]) === userId && String(roleRows[i][1]) === appName) {
        role = String(roleRows[i][2] || '');
        break;
      }
    }
    if (!role || role === 'none') return { ok: false };
  }

  return { ok: true, user_id: userId, name: userName, is_admin: isAdmin, role: role };
}

// ============================================================
// セッション保存（beaufield-auth の sessions シート）
// ============================================================
function _saveSession(ss, token, user_id, expiresAt) {
  let sh = ss.getSheetByName('sessions');
  if (!sh) {
    // sessions シートが未作成なら自動作成
    sh = ss.insertSheet('sessions');
    sh.appendRow(['token', 'user_id', 'expires_at']);
  }

  // 期限切れセッションを削除（遅延クリーンアップ）
  const data = sh.getDataRange().getValues();
  const now  = Date.now();
  for (let i = data.length - 1; i >= 1; i--) {
    if (Number(data[i][2]) < now) {
      sh.deleteRow(i + 1);
    }
  }

  // 新しいセッションを追記
  sh.appendRow([token, user_id, expiresAt]);
}

// ============================================================
// セッショントークンからuser_idを取得
// ============================================================
function _getSessionUser(ss, token) {
  const sh = ss.getSheetByName('sessions');
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const now  = Date.now();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token && Number(data[i][2]) > now) {
      return String(data[i][1]); // user_id
    }
  }
  return null;
}

// ============================================================
// 対象ユーザーのセッションを全削除（PINリセット・無効化時に使用）
// ============================================================
function _deleteUserSessions(ss, userId) {
  const sh = ss.getSheetByName('sessions');
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === userId) sh.deleteRow(i + 1);
  }
}

// ============================================================
// 現在の端末で使用中のセッションだけを削除する（サーバー側ログアウト）
// ============================================================
function logout(data) {
  const token = String((data && (data.session_token || data.token)) || '');
  if (!token) return { success: true };

  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
  const sh = ss.getSheetByName('sessions');
  if (!sh) return { success: true };

  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === token) sh.deleteRow(i + 1);
  }
  return { success: true };
}

// ============================================================
// 有効な管理者チェック（usersシートのD列 active・F列 is_admin）
// ============================================================
function _isAdmin(ss, user_id) {
  const rows = ss.getSheetByName('users').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === user_id) {
      // 管理者フラグが残っていても、無効化済みユーザーには操作させない。
      return (rows[i][3] === true || rows[i][3] === 'TRUE') &&
        (rows[i][5] === true || rows[i][5] === 'TRUE');
    }
  }
  return false;
}

// ============================================================
// ヘルパー
// ============================================================
function _json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// keepWarm: GASのコールドスタートを防ぐ定期実行用関数
// ============================================================
function keepWarm() {
  // 何もしない（トリガーによる定期呼び出しでインスタンスをウォームアップするだけ）
}

// ============================================================
// マイグレーション（v1.7.0・GASエディタから一度だけ手動実行する）
// line-progress（友だち登録進捗ダッシュボード）はGAS側で権限判定をせず全社員閲覧可の
// 方針だが、ポータルのホームにタイルを出すには user_app_roles に行が必要
// （getUserApps() が role!='none' のアプリしか返さないため）。
// active な全ユーザーに 'line-progress'/'user' 行を冪等に追加する。
// 既に行がある場合はスキップするので、何度実行しても安全。
// 参照: LINEHarness/友だち登録進捗ダッシュボード_設計.md §5
// ============================================================
function migrateAddLineProgressRoles() {
  const APP_NAME = 'line-progress';
  const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);

  const usersSh = ss.getSheetByName('users');
  const users = usersSh.getDataRange().getValues();
  const activeUserIds = [];
  for (let i = 1; i < users.length; i++) {
    const active = users[i][3];
    if (active === true || active === 'TRUE') {
      activeUserIds.push(String(users[i][0]));
    }
  }

  const rolesSh = ss.getSheetByName('user_app_roles');
  const roles = rolesSh.getDataRange().getValues();
  const existing = new Set();
  for (let i = 1; i < roles.length; i++) {
    if (String(roles[i][1]) === APP_NAME) existing.add(String(roles[i][0]));
  }

  let added = 0;
  activeUserIds.forEach(function (userId) {
    if (existing.has(userId)) return; // 既に行がある→スキップ（冪等性）
    rolesSh.appendRow([userId, APP_NAME, 'user']);
    added++;
  });

  Logger.log('migrateAddLineProgressRoles 完了: 対象active ' + activeUserIds.length +
    '名中、新規追加 ' + added + '件（既存 ' + (activeUserIds.length - added) + '件はスキップ）');
}
