/**
 * 営業別 回収率アプリ（collection-rate）— バックエンドAPI v1.2（回収遅延を追加）
 *
 * 構成: GitHub Pages（index.html）+ この GAS WebApp + Googleシート（DB）
 * 仕様: collection-rate/仕様書_v1.1.md・仕様書_v1.2_回収遅延.md・実装計画_v1.1.md・実装計画_v1.2_回収遅延.md
 * 認証: beaufield-auth 共通セッション。user_app_roles に collection-rate のロール
 *       （user / admin）が明示登録されている在籍ユーザーだけが見られる。
 *       取り込み・読み戻し・巻き戻し・診断は、ローカルの取り込みスクリプトから
 *       スクリプトプロパティ IMPORT_KEY で認証する（delayImport＝回収遅延の取り込みも同じ）。
 *       回収遅延の閲覧（action: delay）は、画面と同じポータルのセッション（user / admin）。
 * データ: 得意先名・金額はこのGASとDBシートの中だけで扱い、公開HTMLには埋め込まない。
 *
 * 記録と確定の仕組み（実装計画 §4-5）:
 *   rows_YYYYMMDD_HHmmss シート … ある時点の全取り込み行。取り込みのたびに新しく作り、書き換えない
 *   history シート            … 操作1回につき1行。その操作の後の状態（状態JSON）を丸ごと持つ。追記だけ
 *   スクリプトプロパティ STATE … 今の確定がどの記録行かを指す。これを書いた瞬間がその操作の確定
 *   commits シート            … 確定した版の STATE の写し（確定台帳）。復旧の根拠にだけ使う
 *   delay_YYYYMMDD_HHmmss シート … 回収遅延一覧表1回分（1基準日）。取り込みのたびに新しく作り、書き換えない。
 *                                状態 JSON の delays から参照する。自動でも手でも消さない（v1.2 R4）
 *
 * スクリプトプロパティ（コードに直書きしない）:
 *   AUTH_SHEET_ID   … beaufield-auth スプレッドシートのID
 *   DB_SHEET_ID     … 回収率DBスプレッドシートのID
 *   IMPORT_KEY      … 取り込み用キー（32文字以上のランダム文字列）
 *   NON_SALES_REPS  … 任意。担当者比較から外す担当名をカンマ区切り（未設定なら「店売り」）
 *   STATE           … 確定ポインタ（このコードだけが書く）
 */

const VERSION = '1.2.0';
const APP_NAME = 'collection-rate';
const CACHE_TTL_SESSION = 60;           // 権限変更・ログアウトを最大1分で反映
const CACHE_TTL_DATA = 300;             // 集計結果。キーに版と確定日時を入れるので古い値は出ない
const CACHE_CHUNK_CHARS = 30000;        // 1件 90KB 以下（日本語は1文字最大3バイト）
const CACHE_MAX_CHUNKS = 20;
const START_MONTH = '2026-07';          // 取り込み開始月（新しい期）
const CLOSES = ['15日', '末日'];
const VIEWS = ['合算', '末日', '15日'];
const ACTIONS = ['import', 'exclude_on', 'exclude_off', 'target_set', 'target_clear', 'rollback', 'delay_import'];
const STATE_MAX_CHARS = 45000;
const KEEP_CHAIN_SHEETS = 5;             // 系列の新しい方から残す rows シートの数
const KEEP_RECENT_MS = 15 * 60 * 1000;  // 作ってから15分以内の rows シートは残す
const HISTORY_LIMIT = 30;
const ROWS_NAME_RE = /^rows_[0-9]{8}_[0-9]{6}(_[0-9]+)?$/;
const DELAY_NAME_RE = /^delay_[0-9]{8}_[0-9]{6}(_[0-9]+)?$/;
const DELAY_START = '2026-11-07';          // 回収遅延の取り込み開始日（仕様書 v1.2 D4）
const DELAY_METHODS = ['現', '振', '引'];
const DELAY_CLOSES = ['15', '末'];
const CTRL_RE = /[\u0000-\u001f\u007f]/;
const NUM_COLS = ['prev', 'paid', 'adj', 'carry'];
const DELAY_TEXT_COLS = ['asOf', 'rep', 'code', 'name', 'tel', 'method', 'close', 'due'];
const DELAY_NUM_COLS = ['m3', 'm2', 'm1', 'm0', 'total', 'pay1', 'pay2', 'bal'];

const SHEETS = {
  rows:    ['month', 'close', 'closeDate', 'repCode', 'rep', 'code', 'name', 'prev', 'paid', 'adj', 'carry'],
  history: ['seq', 'parent', 'at', 'by', 'action', 'target', 'before', 'after', 'note', 'rows', 'state'],
  commits: ['seq', 'row', 'rows', 'at', 'recordedAt'],
  delay:   ['asOf', 'rep', 'code', 'name', 'tel', 'method', 'close', 'due', 'm3', 'm2', 'm1', 'm0', 'total', 'pay1', 'pay2', 'bal']
};

// 系列をたどった回数（試験で無限ループにならないことを確かめるため）
let LAST_CHAIN_STEPS_ = 0;

function prop_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('スクリプトプロパティ未設定: ' + key);
  return v;
}

// 定型コードの例外を作る（理由はログにだけ出す）
function fail_(code, detail) {
  if (detail) Logger.log(code + ': ' + detail);
  const e = new Error(code);
  e.code = code;
  return e;
}

function isObj_(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }

// ============================================================
// セッション検証（beaufield-auth。line-progress と同じ照合方法）
// キャッシュの失敗は無認証扱いにせず、シートで照合する
// ============================================================
function validateSession_(token) {
  if (!token) return { valid: false };
  const cacheKey = 'sess_cr_v1_' + String(token).slice(-32);
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }
  if (cache) {
    let cached = null;
    try { cached = cache.get(cacheKey); } catch (e) { cached = null; Logger.log('セッションキャッシュの読み込み失敗（シートで照合）: ' + e); }
    if (cached !== null && cached !== undefined) {
      try { return JSON.parse(cached); } catch (e) {}
    }
  }
  let r = { valid: false };
  try {
    const ss = SpreadsheetApp.openById(prop_('AUTH_SHEET_ID'));
    const sessions = ss.getSheetByName('sessions').getDataRange().getValues();
    const now = Date.now();
    for (let i = 1; i < sessions.length; i++) {
      if (String(sessions[i][0]) !== String(token)) continue;
      if (Number(sessions[i][2]) < now) break;
      const userId = String(sessions[i][1]);
      const users = ss.getSheetByName('users').getDataRange().getValues();
      let userRow = null;
      for (let j = 1; j < users.length; j++) {
        if (String(users[j][0]) === userId) { userRow = users[j]; break; }
      }
      if (!userRow || !(userRow[3] === true || userRow[3] === 'TRUE')) break;
      const roles = ss.getSheetByName('user_app_roles').getDataRange().getValues();
      let role = '';
      for (let j = 1; j < roles.length; j++) {
        if (String(roles[j][0]) === userId && String(roles[j][1]) === APP_NAME) {
          role = String(roles[j][2] || '').trim().toLowerCase();
          break;
        }
      }
      if (role !== 'user' && role !== 'admin') break;
      r = { valid: true, user_id: userId, name: String(userRow[1] || userId), role: role };
      break;
    }
  } catch (e) {
    Logger.log('セッション検証エラー: ' + e);
    return { valid: false }; // 一時的な失敗はキャッシュしない
  }
  if (cache) {
    try { cache.put(cacheKey, JSON.stringify(r), r.valid ? CACHE_TTL_SESSION : 30); } catch (e) {}
  }
  return r;
}

function sessionInvalid_() {
  return jsonResponse_({ success: false, error: 'SESSION_INVALID',
    message: '認証が必要です。ポータルからログインし直してください。' });
}

// 取り込み用キーの照合（長さを揃えた比較）
function importKeyOk_(key) {
  const want = prop_('IMPORT_KEY');
  if (!key || String(key).length !== want.length) return false;
  let diff = 0;
  for (let i = 0; i < want.length; i++) diff |= want.charCodeAt(i) ^ String(key).charCodeAt(i);
  return diff === 0;
}

// ============================================================
// エントリーポイント
// ============================================================
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  if (p.action === 'version') return jsonResponse_({ success: true, version: VERSION });
  return jsonResponse_({ success: false, error: 'USE_POST' });
}

function doPost(e) {
  let p = {};
  try {
    if (e && e.postData && e.postData.contents) p = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ success: false, error: 'BAD_JSON' });
  }
  if (!isObj_(p)) return jsonResponse_({ success: false, error: 'BAD_JSON' });
  try {
    switch (p.action || '') {
      case 'version':
        return jsonResponse_({ success: true, version: VERSION });

      // ---- 画面から（ポータルのセッション） ----
      case 'data': {
        const auth = validateSession_(p.session_token);
        if (!auth.valid) return sessionInvalid_();
        const d = getData_({ nocache: auth.role === 'admin' && p.nocache === true });
        const out = Object.assign({}, d, { me: { name: auth.name, role: auth.role } });
        if (auth.role !== 'admin') { // 検索用一覧と変更履歴は admin だけ
          delete out.customers;
          delete out.history;
          delete out.historyOk;
        }
        return jsonResponse_(out);
      }
      case 'delay': {
        const auth = validateSession_(p.session_token);
        if (!auth.valid) return sessionInvalid_();
        return jsonResponse_(getDelay_({ asOf: p.asOf, nocache: auth.role === 'admin' && p.nocache === true }));
      }
      case 'setExcluded': {
        const auth = validateSession_(p.session_token);
        if (!auth.valid) return sessionInvalid_();
        if (auth.role !== 'admin') return jsonResponse_({ success: false, error: 'FORBIDDEN' });
        const code = (p.code === null || p.code === undefined) ? '' : String(p.code).trim();
        return jsonResponse_(setExcluded_(code, !!p.on, String(p.note || ''), auth.name));
      }
      case 'setTarget': {
        const auth = validateSession_(p.session_token);
        if (!auth.valid) return sessionInvalid_();
        if (auth.role !== 'admin') return jsonResponse_({ success: false, error: 'FORBIDDEN' });
        return jsonResponse_(setTarget_(p.view || '合算', p.rate, auth.name));
      }

      // ---- 取り込みスクリプトから（IMPORT_KEY） ----
      case 'import':
        if (!importKeyOk_(p.key)) return jsonResponse_({ success: false, error: 'BAD_KEY' });
        return jsonResponse_(importRows_(p));
      case 'delayImport':
        if (!importKeyOk_(p.key)) return jsonResponse_({ success: false, error: 'BAD_KEY' });
        return jsonResponse_(importDelay_(p));
      case 'importStatus':
        if (!importKeyOk_(p.key)) return jsonResponse_({ success: false, error: 'BAD_KEY' });
        return jsonResponse_(importStatus_(p));
      case 'rollback':
        if (!importKeyOk_(p.key)) return jsonResponse_({ success: false, error: 'BAD_KEY' });
        return jsonResponse_(rollback_(p));
      case 'diagnose':
        if (!importKeyOk_(p.key)) return jsonResponse_({ success: false, error: 'BAD_KEY' });
        return jsonResponse_(diagnose_(p));

      default:
        return jsonResponse_({ success: false, error: 'UNKNOWN_ACTION' });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err + '\n' + (err && err.stack));
    // 定型コード（英大文字と_だけ）はそのまま返す。それ以外の本文は伏せる
    const msg = String(err && err.message || '');
    return jsonResponse_({ success: false, error: /^[A-Z_]+$/.test(msg) ? msg : 'INTERNAL_ERROR' });
  }
}

// ============================================================
// DBシートの基本操作
// ============================================================
function db_() { return SpreadsheetApp.openById(prop_('DB_SHEET_ID')); }

// history / commits シート（無ければ作る）
function sheetFor_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const rng = sh.getRange(1, 1, 1, SHEETS[name].length);
    rng.setNumberFormat('@');
    rng.setValues([SHEETS[name]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// シートの行数が足りなければ足す（新しいシートは1,000行しかない）
function ensureRows_(sh, need) {
  const max = sh.getMaxRows();
  if (max < need) sh.insertRowsAfter(max, need - max);
}

// v1.0 からの数値変換（取り込みの入力確認とシートの読み戻しには使わない）
function int_(v) {
  const n = Number(String(v).replace(/,/g, ''));
  if (!isFinite(n) || Math.floor(n) !== n) throw new Error('BAD_NUMBER');
  return n;
}

function hex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// 取り込みスクリプト（kaishu_import.py の digest）と同じ指紋（変えない）
function digest_(rows) {
  const sorted = rows.slice().sort(function (a, b) {
    if (a.code !== b.code) return a.code < b.code ? -1 : 1;
    return a.repCode < b.repCode ? -1 : (a.repCode > b.repCode ? 1 : 0);
  });
  const body = sorted.map(function (r) {
    return [r.repCode, r.code, r.prev, r.paid, r.adj, r.carry].join('|');
  }).join('\n');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, body, Utilities.Charset.UTF_8);
  return hex_(bytes);
}

// ============================================================
// 全文指紋と入力の決まり（実装計画 §4-5-11。kaishu_import.py の full_digest と同じ計算）
// ============================================================
function cmpCodeRep_(a, b) {
  const ac = String(a.code), bc = String(b.code);
  if (ac !== bc) return ac < bc ? -1 : 1;
  const ar = String(a.repCode), br = String(b.repCode);
  return ar < br ? -1 : (ar > br ? 1 : 0);
}

// 1つのキー（月×締め）の行すべての全11列から作る指紋
function fullDigest_(rows) {
  const body = rows.slice().sort(cmpCodeRep_).map(function (r) {
    return SHEETS.rows.map(function (k) { return String(r[k]); }).join('\u001f');
  }).join('\n');
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, body, Utilities.Charset.UTF_8));
}

// 送られてきた1行の型と形を確かめる。外れたらエラーコード、通れば ''
function validateImportRow_(r) {
  if (!isObj_(r)) return 'BAD_TEXT';
  const codes = ['repCode', 'code'];
  for (let i = 0; i < codes.length; i++) {
    const v = r[codes[i]];
    if (typeof v !== 'string' || !/^[0-9]+$/.test(v)) return 'BAD_TEXT';
  }
  const texts = ['rep', 'name'];
  for (let i = 0; i < texts.length; i++) {
    const v = r[texts[i]];
    if (typeof v !== 'string' || CTRL_RE.test(v)) return 'BAD_TEXT';
  }
  for (let i = 0; i < NUM_COLS.length; i++) {
    const v = r[NUM_COLS[i]];
    if (typeof v !== 'number' || !Number.isSafeInteger(v)) return 'BAD_NUMBER';
  }
  return '';
}

// ============================================================
// 回収遅延の入力の決まり（実装計画 v1.2 §1。kaishu_import.py の delay_full_digest と同じ計算）
// ============================================================
// YYYY-MM-DD の形で、実在する日か（繰り上がりで通さない）
function dateReal_(s) {
  if (typeof s !== 'string' || !/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(s)) return false;
  const y = Number(s.slice(0, 4)), m = Number(s.slice(5, 7)), d = Number(s.slice(8, 10));
  const t = new Date(Date.UTC(y, m - 1, d));
  return t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d;
}

// 基準日として受け付けるか（形・実在・開始日以降）
function delayAsOfOk_(s) { return dateReal_(s) && s >= DELAY_START; }

// 基準日から4つの請求月を計算する（年をまたぐ）。m0 = 基準月、m1 = その前月…
function delayMonths_(asOf) {
  const base = Number(asOf.slice(0, 4)) * 12 + Number(asOf.slice(5, 7)) - 1;
  const mm = function (k) {
    const n = base - k;
    return String(Math.floor(n / 12)).padStart(4, '0') + '-' + String(n % 12 + 1).padStart(2, '0');
  };
  return { m3: mm(3), m2: mm(2), m1: mm(1), m0: mm(0) };
}

function monthsEqual_(a, b) {
  if (!isObj_(a) || !isObj_(b)) return false;
  return ['m3', 'm2', 'm1', 'm0'].every(function (k) { return typeof a[k] === 'string' && a[k] === b[k]; });
}

// 文字列の列の内容の決まり。外れたら false
function delayTextRuleOk_(k, s) {
  if (CTRL_RE.test(s)) return false;
  switch (k) {
    case 'asOf': return true; // 外側との一致は別に確かめる
    case 'code': return /^[0-9]+$/.test(s);
    case 'tel': return s === '' || /^[0-9-]+$/.test(s);
    case 'method': return DELAY_METHODS.indexOf(s) >= 0;
    case 'close': return DELAY_CLOSES.indexOf(s) >= 0;
    default: return s !== ''; // rep・name・due
  }
}

// 送られてきた1行の型と形を確かめる。外れたらエラーコード、通れば ''
function validateDelayRow_(r, asOf) {
  if (!isObj_(r)) return 'BAD_TEXT';
  for (let i = 0; i < DELAY_TEXT_COLS.length; i++) {
    if (typeof r[DELAY_TEXT_COLS[i]] !== 'string' || CTRL_RE.test(r[DELAY_TEXT_COLS[i]])) return 'BAD_TEXT';
  }
  for (let i = 0; i < DELAY_NUM_COLS.length; i++) {
    const v = r[DELAY_NUM_COLS[i]];
    if (typeof v !== 'number' || !Number.isSafeInteger(v)) return 'BAD_NUMBER';
  }
  if (r.asOf !== asOf) return 'ASOF_MISMATCH';
  for (let i = 0; i < DELAY_TEXT_COLS.length; i++) {
    if (!delayTextRuleOk_(DELAY_TEXT_COLS[i], r[DELAY_TEXT_COLS[i]])) return 'BAD_TEXT';
  }
  return '';
}

// 安全な整数の範囲を超えないように足す。超えたら BAD_NUMBER
function addSafe_(a, b) {
  const n = a + b;
  if (!Number.isSafeInteger(n)) throw fail_('BAD_NUMBER', '合計が安全な整数の範囲を超えた');
  return n;
}

function cmpDelayCode_(a, b) {
  const x = String(a.code), y = String(b.code);
  return x < y ? -1 : (x > y ? 1 : 0);
}

// 遅延の全文指紋: 16列。文字列はそのまま、整数は10進文字列、列は \u001f、行は \n、code の文字列順
function delayFullDigest_(rows) {
  const body = rows.slice().sort(cmpDelayCode_).map(function (r) {
    return SHEETS.delay.map(function (k) { return String(r[k]); }).join('\u001f');
  }).join('\n');
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, body, Utilities.Charset.UTF_8));
}

// 検証済みの行から8項目の合計を作る
function delaySums_(rows) {
  const s = {};
  DELAY_NUM_COLS.forEach(function (k) { s[k] = 0; });
  rows.forEach(function (r) { DELAY_NUM_COLS.forEach(function (k) { s[k] = addSafe_(s[k], r[k]); }); });
  return s;
}

// シートから読み戻した数値の文字列を確かめる（'-0'・'007'・小数・空・カンマ入りは拒否）
function parseStoredInt_(v) {
  const s = String(v);
  if (!/^(0|-?[1-9][0-9]*)$/.test(s)) throw new Error('BAD_STORED_INT');
  const n = Number(s);
  if (!Number.isSafeInteger(n)) throw new Error('BAD_STORED_INT');
  return n;
}

// rows シートの1行（文字列）を確かめて行オブジェクトにする。外れたら errCode
function parseStoredRow_(r, rowNo, errCode) {
  const o = {};
  for (let j = 0; j < SHEETS.rows.length; j++) {
    const name = SHEETS.rows[j];
    const s = String(r[j]);
    let ok = true;
    if (NUM_COLS.indexOf(name) >= 0) {
      try { o[name] = parseStoredInt_(s); } catch (e) { ok = false; }
    } else {
      if (CTRL_RE.test(s)) ok = false;
      else if ((name === 'repCode' || name === 'code') && !/^[0-9]+$/.test(s)) ok = false;
      else if (name === 'month' && !/^[0-9]{4}-[0-9]{2}$/.test(s)) ok = false;
      else if (name === 'close' && CLOSES.indexOf(s) < 0) ok = false;
      else if (name === 'closeDate' && !/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(s)) ok = false;
      o[name] = s;
    }
    if (!ok) throw fail_(errCode, 'rows シートの値が正しくない: 行' + rowNo + ' 列' + name);
  }
  return o;
}

// rows シートの全行を読む（errCode は取り込みの確認中なら VERIFY_FAILED、確定データなら STATE_BROKEN）
function readRowsSheet_(sh, errCode) {
  const v = sh.getDataRange().getValues();
  if (!v.length || SHEETS.rows.some(function (h, i) { return String(v[0][i]) !== h; })) {
    throw fail_(errCode, 'rows シートの見出しが違う: ' + sh.getName());
  }
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(function (c) { return c === '' || c === null; })) continue;
    out.push(parseStoredRow_(v[i], i + 1, errCode));
  }
  return out;
}

function rowStrings_(r) { return SHEETS.rows.map(function (k) { return String(r[k]); }); }

// 遅延シートを読む（形と型だけを確かめる。件数・指紋は比べない。実装計画 v1.2 §1-3）
// 外れたら DELAY_UNREADABLE（理由はログ）。見出しが16列とも一致・各行の型・全行の asOf が引数と同じ・コード重複なし・1行以上
function loadDelaySheet_(sh, asOf) {
  const v = sh.getDataRange().getValues();
  const name = sh.getName();
  if (!v.length || SHEETS.delay.some(function (h, i) { return String(v[0][i]) !== h; })) {
    throw fail_('DELAY_UNREADABLE', 'delay シートの見出しが違う: ' + name);
  }
  const out = [], seen = {};
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(function (c) { return c === '' || c === null; })) continue;
    const o = {};
    for (let j = 0; j < SHEETS.delay.length; j++) {
      const k = SHEETS.delay[j];
      const s = String(v[i][j]);
      let good = true;
      if (DELAY_NUM_COLS.indexOf(k) >= 0) {
        try { o[k] = parseStoredInt_(s); } catch (e) { good = false; }
      } else {
        good = delayTextRuleOk_(k, s) && (k !== 'asOf' || s === asOf);
        o[k] = s;
      }
      if (!good) throw fail_('DELAY_UNREADABLE', 'delay シートの値が正しくない: ' + name + ' 行' + (i + 1) + ' 列' + k);
    }
    if (seen[o.code]) throw fail_('DELAY_UNREADABLE', 'delay シートにコード重複: ' + name);
    seen[o.code] = true;
    out.push(o);
  }
  if (!out.length) throw fail_('DELAY_UNREADABLE', 'delay シートに行が無い: ' + name);
  return out;
}

// 読めた行から件数と全文指紋を計算する
function delayActual_(rows) { return { rows: rows.length, full: delayFullDigest_(rows) }; }

function cmpFull_(a, b) {
  if (a.month !== b.month) return a.month < b.month ? -1 : 1;
  if (a.close !== b.close) return a.close < b.close ? -1 : 1;
  return cmpCodeRep_(a, b);
}

// 2つの行の集まりを §4-5-11 の正規化・並びで1行ずつ全11列比べる。違えば最初の場所を返す
function fullRowsEqual_(a, b) {
  if (a.length !== b.length) return { row: -1, col: '行数', detail: a.length + ' / ' + b.length };
  const x = a.slice().sort(cmpFull_), y = b.slice().sort(cmpFull_);
  for (let i = 0; i < x.length; i++) {
    const sx = rowStrings_(x[i]), sy = rowStrings_(y[i]);
    for (let j = 0; j < sx.length; j++) {
      if (sx[j] !== sy[j]) return { row: i, col: SHEETS.rows[j], detail: x[i].month + '|' + x[i].close + ' code ' + x[i].code };
    }
  }
  return null;
}

// 変更記録の全11列を String どうしで比べ、最初に違った列名を返す（同じなら ''）
function historyRowEqual_(written, read) {
  for (let i = 0; i < SHEETS.history.length; i++) {
    if (String(read[i]) !== String(written[i])) return SHEETS.history[i];
  }
  return '';
}

// ============================================================
// 状態 JSON と確定ポインタ（実装計画 §4-5-2・§4-5-5）
// ============================================================
function emptyState_() {
  return { v: 1, rows: '', rowsCount: 0, imports: {}, excluded: {},
           targets: { '合算': null, '末日': null, '15日': null }, excludedAt: null, targetsAt: null };
}

// 状態 JSON の正しさ。外れたら理由、正しければ ''
function stateInvalid_(st) {
  if (!isObj_(st)) return '状態がオブジェクトでない';
  if (st.v !== 1) return '状態の v が 1 でない';
  if (typeof st.rows !== 'string' || (st.rows !== '' && !ROWS_NAME_RE.test(st.rows))) return '状態の rows の形が違う';
  if (!isObj_(st.imports) || !isObj_(st.excluded) || !isObj_(st.targets)) return '状態の imports・excluded・targets の形が違う';
  if (!Number.isInteger(st.rowsCount) || st.rowsCount < 0) return '状態の rowsCount が整数でない';
  const keys = Object.keys(st.imports);
  let sum = 0;
  for (let i = 0; i < keys.length; i++) {
    const im = st.imports[keys[i]];
    if (!isObj_(im) || !Number.isInteger(im.rows) || im.rows < 0) return '取り込み記録の形が違う: ' + keys[i];
    if (keys[i] !== im.month + '|' + im.close) return '取り込み記録のキーが違う: ' + keys[i];
    sum += im.rows;
  }
  if (st.rows === '') {
    if (st.rowsCount !== 0 || keys.length) return 'rows が空なのに取り込み記録がある';
  } else if (st.rowsCount !== sum) {
    return 'rowsCount が取り込み記録の行数の合計と違う';
  }
  for (let i = 0; i < VIEWS.length; i++) {
    const t = st.targets[VIEWS[i]];
    if (!(t === null || (typeof t === 'number' && isFinite(t)))) return '目標値の形が違う';
  }
  if (!(st.excludedAt === null || typeof st.excludedAt === 'string')) return 'excludedAt の形が違う';
  if (!(st.targetsAt === null || typeof st.targetsAt === 'string')) return 'targetsAt の形が違う';
  // 回収遅延（v1.2。無ければ {} として扱う）
  if (st.delays !== undefined) {
    if (!isObj_(st.delays)) return '状態の delays の形が違う';
    const dk = Object.keys(st.delays);
    for (let i = 0; i < dk.length; i++) {
      const d = st.delays[dk[i]];
      if (!isObj_(d)) return '遅延の記録の形が違う: ' + dk[i];
      if (!delayAsOfOk_(dk[i]) || d.asOf !== dk[i]) return '遅延の記録のキーが基準日として正しくない: ' + dk[i];
      if (typeof d.sheet !== 'string' || !DELAY_NAME_RE.test(d.sheet)) return '遅延の sheet の形が違う: ' + dk[i];
      if (!Number.isInteger(d.rows) || d.rows < 1) return '遅延の rows が1以上の整数でない: ' + dk[i];
      if (typeof d.full !== 'string' || !/^[0-9a-f]{64}$/.test(d.full)) return '遅延の full の形が違う: ' + dk[i];
      if (!monthsEqual_(d.months, delayMonths_(dk[i]))) return '遅延の months が基準日と合わない: ' + dk[i];
      if (!isObj_(d.sums) || !DELAY_NUM_COLS.every(function (k) { return Number.isSafeInteger(d.sums[k]); })) return '遅延の sums の形が違う: ' + dk[i];
    }
  }
  return '';
}

// 状態 JSON が指すシートがすべてあるか（rows シートと遅延の全シート。中身は読まない）。無ければ理由、あれば ''
// 「状態が正しい」と判断する場所はすべてこの関数を使う（実装計画 v1.2 R8）
// found（任意）が渡されたら、見つけた rows シートを found.rows に入れる（シートの取得を1回で済ませるため）
function stateSheetsOk_(ss, st, found) {
  if (st.rows !== '') {
    const rs = ss.getSheetByName(st.rows);
    if (!rs) return '状態の rows シートが無い';
    if (found) found.rows = rs;
  }
  const d = st.delays || {};
  const keys = Object.keys(d);
  for (let i = 0; i < keys.length; i++) {
    if (!ss.getSheetByName(d[keys[i]].sheet)) return '状態の遅延シートが無い: ' + keys[i];
  }
  return '';
}

function isPosInt_(v) { return typeof v === 'number' && Number.isInteger(v) && v >= 1; }

// 確定状態の対応を確かめる。通れば {ok:true, seq,row,rows,at,state,stateChars,rowsSheet}、外れたら {ok:false, reason}
// シートの読み込み中の例外（削除されたシート等）はそのまま投げる
function inspectState_(ss, raw) {
  if (raw === null || raw === undefined || raw === '') {
    return { ok: true, seq: 0, row: 0, rows: '', at: null, state: emptyState_(), stateChars: 0, rowsSheet: null };
  }
  const bad = function (reason) { return { ok: false, reason: reason }; };
  let S;
  try { S = JSON.parse(raw); } catch (e) { return bad('STATE が JSON として読めない'); }
  if (!isObj_(S)) return bad('STATE がオブジェクトでない');
  if (!isPosInt_(S.seq)) return bad('STATE.seq が1以上の整数でない');
  if (!(isPosInt_(S.row) && S.row >= 2)) return bad('STATE.row が2以上の整数でない');
  if (typeof S.rows !== 'string') return bad('STATE.rows が文字列でない');
  if (typeof S.at !== 'string') return bad('STATE.at が文字列でない');
  const hs = ss.getSheetByName('history');
  if (!hs) return bad('history シートが無い');
  if (S.row > hs.getLastRow()) return bad('STATE.row の行が history シートに無い');
  const v = hs.getRange(S.row, 1, 1, SHEETS.history.length).getValues()[0].map(function (c) { return String(c); });
  if (v[0] !== String(S.seq)) return bad('記録行の seq が STATE と違う');
  if (v[9] !== S.rows) return bad('記録行の rows が STATE と違う');
  if (v[2] !== S.at) return bad('記録行の at が STATE と違う');
  let st;
  try { st = JSON.parse(v[10]); } catch (e) { return bad('状態 JSON が読めない'); }
  const why = stateInvalid_(st);
  if (why) return bad(why);
  if (st.rows !== v[9]) return bad('状態 JSON の rows が記録行の rows と違う');
  const found = { rows: null };
  const noSheet = stateSheetsOk_(ss, st, found);
  if (noSheet) return bad(noSheet);
  const rowsSheet = found.rows;
  return { ok: true, seq: S.seq, row: S.row, rows: S.rows, at: S.at, state: st, stateChars: v[10].length, rowsSheet: rowsSheet };
}

// 確定状態を読む。壊れていれば STATE_BROKEN
function readState_(ss, raw) {
  if (raw === undefined) raw = PropertiesService.getScriptProperties().getProperty('STATE');
  const r = inspectState_(ss, raw);
  if (!r.ok) throw fail_('STATE_BROKEN', r.reason);
  return r;
}

// ============================================================
// 系列の検査（実装計画 §4-5-12）
// ============================================================
function readHistoryCols_(hs) {
  const last = hs.getLastRow();
  if (last < 2) return [];
  return hs.getRange(2, 1, last - 1, 10).getValues();
}

function chainFail_(seq, reason) {
  const e = fail_('HISTORY_BROKEN', '系列の検査で外れた: #' + seq + ' ' + reason);
  e.seq = seq;
  e.reason = reason;
  return e;
}

// cols は history シートの2行目以降の1〜10列目、head は {seq, row}。系列を新しい順に返す
function buildChain_(cols, head) {
  LAST_CHAIN_STEPS_ = 0;
  if (head.seq === 0) return [];
  const bySeq = {};
  for (let i = 0; i < cols.length; i++) {
    const s = String(cols[i][0]);
    if (!/^[1-9][0-9]*$/.test(s)) continue; // 正の整数として読めない行は未確定行として無視
    (bySeq[s] = bySeq[s] || []).push({ c: cols[i], row: i + 2 });
  }
  const chain = [];
  let cur = head.seq;
  while (cur !== 0) {
    LAST_CHAIN_STEPS_++;
    if (LAST_CHAIN_STEPS_ > head.seq) throw chainFail_(cur, 'たどる回数が版の数を超えた');
    const hits = bySeq[String(cur)];
    if (!hits) throw chainFail_(cur, '行が無い（欠けた親）');
    if (hits.length > 1) throw chainFail_(cur, '同じ seq の行が2つ以上ある');
    const c = hits[0].c.map(function (x) { return String(x); });
    if (!/^(0|[1-9][0-9]*)$/.test(c[1])) throw chainFail_(cur, 'parent が0以上の整数でない');
    const parent = Number(c[1]);
    if (parent >= cur) throw chainFail_(cur, 'parent が自分以上（自己参照・逆向き）');
    if (ACTIONS.indexOf(c[4]) < 0) throw chainFail_(cur, 'action が決まった7種類以外');
    if (c[2] === '' || isNaN(new Date(c[2]).getTime())) throw chainFail_(cur, 'at が日時として読めない');
    if (c[9] !== '' && !ROWS_NAME_RE.test(c[9])) throw chainFail_(cur, 'rows 列の形が違う');
    if (cur === head.seq && hits[0].row !== head.row) throw chainFail_(cur, '先頭の行番号が STATE と違う');
    chain.push({ seq: cur, parent: parent, at: c[2], by: c[3], action: c[4], target: c[5],
                 before: c[6], after: c[7], note: c[8], rows: c[9], row: hits[0].row });
    cur = parent;
  }
  return chain;
}

function loadChain_(ss, st) {
  if (st.seq === 0) return [];
  const hs = ss.getSheetByName('history');
  if (!hs) throw chainFail_(st.seq, 'history シートが無い');
  return buildChain_(readHistoryCols_(hs), st);
}

function histOut_(c) {
  return { seq: c.seq, at: c.at, by: c.by, action: c.action, target: c.target, before: c.before, after: c.after, note: c.note };
}

// ============================================================
// 確定台帳（commits シート。実装計画 §4-5-2・§4-5-3 手順5b）
// ============================================================
function ledgerRows_(ss) {
  const sh = ss.getSheetByName('commits');
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues().map(function (v, i) {
    return { sheetRow: i + 2, seq: String(v[0]), row: String(v[1]), rows: String(v[2]), at: String(v[3]), recordedAt: String(v[4]) };
  });
}

// 書いた順番で最後の行
function ledgerLast_(ss) {
  const sh = ss.getSheetByName('commits');
  if (!sh) return null;
  const last = sh.getLastRow();
  if (last < 2) return null;
  const v = sh.getRange(last, 1, 1, 5).getValues()[0].map(function (c) { return String(c); });
  return { sheetRow: last, seq: v[0], row: v[1], rows: v[2], at: v[3], recordedAt: v[4] };
}

function ledgerMatches_(L, st) {
  return !!L && L.seq === String(st.seq) && L.row === String(st.row) && L.rows === String(st.rows) && L.at === String(st.at);
}

// 確定台帳に1行足して読み戻す。違えば例外
function appendLedger_(ss, seq, row, rows, at) {
  const sh = sheetFor_(ss, 'commits');
  const r = sh.getLastRow() + 1;
  ensureRows_(sh, r);
  const vals = [String(seq), String(row), String(rows), String(at), new Date().toISOString()];
  const rng = sh.getRange(r, 1, 1, 5);
  rng.setNumberFormat('@');
  rng.setValues([vals]);
  SpreadsheetApp.flush();
  const back = sh.getRange(r, 1, 1, 5).getValues()[0];
  for (let i = 0; i < 5; i++) {
    if (String(back[i]) !== vals[i]) throw fail_('LEDGER_VERIFY_FAILED', '確定台帳の読み戻しが違う: 列' + SHEETS.commits[i]);
  }
}

// 確定台帳の補い: 最後の行が今の STATE と違えば（空のときも）今の STATE を1行足す。失敗はログだけ
function syncLedger_(ss, st) {
  const out = { synced: false, ledgerSeq: null, matches: false };
  if (st.seq > 0) {
    try {
      if (!ledgerMatches_(ledgerLast_(ss), st)) {
        appendLedger_(ss, st.seq, st.row, st.rows, st.at);
        out.synced = true;
      }
    } catch (e) {
      Logger.log('確定台帳の補いに失敗（操作は続ける）: ' + e);
    }
  }
  try {
    const L = ledgerLast_(ss);
    out.ledgerSeq = L && /^[1-9][0-9]*$/.test(L.seq) ? Number(L.seq) : null;
    out.matches = st.seq === 0 ? !L : ledgerMatches_(L, st);
  } catch (e) {
    Logger.log('確定台帳の読み込みに失敗: ' + e);
  }
  return out;
}

// ============================================================
// 確定の手順（データを変える操作すべてに共通。実装計画 §4-5-3）
// ============================================================
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw fail_('BUSY');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// 書き込みの前置き: 確定状態・系列を確かめ、台帳を補い、この操作の日時を1回だけ作る
function beginWrite_(ss) {
  const st = readState_(ss);
  const chain = loadChain_(ss, st);
  syncLedger_(ss, st);
  const at = new Date().toISOString();
  return { st: st, chain: chain, at: at };
}

function copyState_(st) { return JSON.parse(JSON.stringify(st)); }

function commit_(ss, cur, action, target, before, after, note, by, nextState, at, afterCommit) {
  // 1. 大きさ
  const json = JSON.stringify(nextState);
  if (json.length > STATE_MAX_CHARS) throw fail_('STATE_TOO_LARGE', '状態 JSON ' + json.length + '文字');
  // 2. 変更記録に1行（seq はシートにある最大＋1。未確定行の番号も避ける）
  const hs = sheetFor_(ss, 'history');
  const last = hs.getLastRow();
  let maxSeq = 0;
  if (last >= 2) {
    hs.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      const s = String(r[0]);
      if (/^[1-9][0-9]*$/.test(s)) maxSeq = Math.max(maxSeq, Number(s));
    });
  }
  const seq = maxSeq + 1;
  const row = last + 1;
  const vals = [String(seq), String(cur.seq), String(at), String(by), String(action), String(target),
                String(before), String(after), String(note), String(nextState.rows), json];
  ensureRows_(hs, row);
  const rng = hs.getRange(row, 1, 1, SHEETS.history.length);
  rng.setNumberFormat('@');
  rng.setValues([vals]);
  // 3. 反映（失敗したら例外のまま。確定していない）
  SpreadsheetApp.flush();
  // 4. 読み戻して全11列を確かめる
  const bad = historyRowEqual_(vals, hs.getRange(row, 1, 1, SHEETS.history.length).getValues()[0]);
  if (bad) throw fail_('VERIFY_FAILED', '変更記録の読み戻しが違う: 行' + row + ' 列' + bad);
  // 5. 確定
  PropertiesService.getScriptProperties().setProperty('STATE', JSON.stringify({ seq: seq, row: row, rows: nextState.rows, at: at }));
  // ここから後は確定済み。失敗はログだけ
  // 5b. 確定台帳
  try { appendLedger_(ss, seq, row, nextState.rows, at); } catch (e) { Logger.log('確定台帳の追記に失敗（次の操作か読み戻しで補う）: ' + e); }
  // 6. 後片付け
  if (afterCommit) {
    try { afterCommit(ss); } catch (e) { Logger.log('後片付けに失敗（確定は済んでいる）: ' + e); }
  }
  // 7. 反映
  try { SpreadsheetApp.flush(); } catch (e) { Logger.log('確定後の反映に失敗（確定は済んでいる）: ' + e); }
  return { success: true, changed: true, seq: seq, stateAt: at };
}

// ============================================================
// 取り込み（実装計画 §4-5-4）
// ============================================================
function importRows_(p) {
  // 1. 入力の確認（ロックの前）
  const month = p.month, close = p.close, closeDate = p.closeDate;
  if (typeof month !== 'string' || !/^[0-9]{4}-[0-9]{2}$/.test(month) || month < START_MONTH) throw fail_('BAD_MONTH');
  if (typeof close !== 'string' || CLOSES.indexOf(close) < 0) throw fail_('BAD_CLOSE');
  if (typeof closeDate !== 'string' || !/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(closeDate)) throw fail_('BAD_TEXT', 'closeDate の形');
  if (!Array.isArray(p.rows) || !p.rows.length) throw fail_('NO_ROWS');
  for (let i = 0; i < p.rows.length; i++) {
    const bad = validateImportRow_(p.rows[i]);
    if (bad) throw fail_(bad, '送られた行 ' + (i + 1) + ' の型か形が違う');
  }
  const rows = p.rows.map(function (r) {
    return { month: month, close: close, closeDate: closeDate, repCode: r.repCode, rep: r.rep, code: r.code, name: r.name,
             prev: r.prev, paid: r.paid, adj: r.adj, carry: r.carry };
  });
  const dg = digest_(rows);
  if (p.digest && p.digest !== dg) throw fail_('DIGEST_MISMATCH');
  const full = fullDigest_(rows);
  if (p.full && p.full !== full) throw fail_('FULL_MISMATCH', 'Python と GAS の全文指紋が違う');
  const key = month + '|' + close;
  const source = (p.source === null || p.source === undefined) ? '' : String(p.source);

  return withLock_(function () {
    const ss = db_();
    // 2. 前置き
    const w = beginWrite_(ss);
    const cur = w.st;
    const curRows = cur.rows === '' ? [] : readRowsSheet_(cur.rowsSheet, 'STATE_BROKEN');
    // 3. 同じ取り込みの再送か
    const old = cur.state.imports[key];
    if (old && old.digest === dg && old.full === full) return { success: true, changed: false, seq: cur.seq };
    // 4. 新しい全件
    const all = curRows.filter(function (r) { return !(r.month === month && r.close === close); }).concat(rows);
    // 5. 新しいシート
    const base = 'rows_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    let name = base;
    for (let n = 2; ss.getSheetByName(name); n++) {
      if (n > 99) throw fail_('NAME_EXHAUSTED');
      name = base + '_' + n;
    }
    const sh = ss.insertSheet(name);
    ensureRows_(sh, all.length + 1);
    const rng = sh.getRange(1, 1, all.length + 1, SHEETS.rows.length);
    rng.setNumberFormat('@');
    rng.setValues([SHEETS.rows].concat(all.map(rowStrings_)));
    SpreadsheetApp.flush();
    // 6. 読み戻して全列を確かめる
    const back = readRowsSheet_(sh, 'VERIFY_FAILED');
    if (back.length !== all.length) throw fail_('VERIFY_FAILED', '新しい rows シートの行数が違う: ' + back.length + ' / ' + all.length);
    const diff = fullRowsEqual_(all, back);
    if (diff) throw fail_('VERIFY_FAILED', '新しい rows シートの読み戻しが違う: ' + diff.detail + ' 列' + diff.col);
    const groups = function (list) {
      const g = {};
      list.forEach(function (r) { (g[r.month + '|' + r.close] = g[r.month + '|' + r.close] || []).push(r); });
      return g;
    };
    const ga = groups(all), gb = groups(back);
    Object.keys(ga).concat(Object.keys(gb)).forEach(function (k) {
      if (!ga[k] || !gb[k] || ga[k].length !== gb[k].length || digest_(ga[k]) !== digest_(gb[k])) {
        throw fail_('VERIFY_FAILED', '新しい rows シートのキー ' + k + ' の件数か指紋が違う');
      }
    });
    // 7. 新しい状態
    const next = copyState_(cur.state);
    next.rows = name;
    next.rowsCount = all.length;
    next.imports[key] = {
      month: month, close: close, closeDate: closeDate, source: source, rows: rows.length, digest: dg, full: full,
      sumPrev: rows.reduce(function (s, r) { return s + r.prev; }, 0),
      sumPaid: rows.reduce(function (s, r) { return s + r.paid; }, 0),
      importedAt: w.at
    };
    // 8. 確定
    const before = old ? old.rows + '件/' + String(old.digest).slice(0, 8) : '';
    const after = rows.length + '件/' + dg.slice(0, 8);
    return commit_(ss, cur, 'import', month + '×' + close, before, after, source, 'import', next, w.at, cleanupRows_);
  });
}

// ============================================================
// 回収遅延の取り込み（実装計画 v1.2 §3 importDelay_）。新しい操作の種類 delay_import
// 回収率の rows シート・集計には触れない（R2）
// ============================================================
function importDelay_(p) {
  // ① 入力の確認（ロックの前）
  const asOf = p.asOf;
  if (!delayAsOfOk_(asOf)) throw fail_('BAD_ASOF', '基準日の形・実在・開始日');
  const months = delayMonths_(asOf);
  if (!monthsEqual_(p.months, months)) throw fail_('BAD_MONTHS', 'months がサーバーの計算と違う');
  if (!Array.isArray(p.rows) || !p.rows.length) throw fail_('NO_ROWS');
  for (let i = 0; i < p.rows.length; i++) {
    const bad = validateDelayRow_(p.rows[i], asOf);
    if (bad) throw fail_(bad, '送られた行 ' + (i + 1) + ' の型か形が違う');
  }
  const seen = {};
  p.rows.forEach(function (r) {
    if (seen[r.code]) throw fail_('DUP_CODE', '同じ得意先コードが複数行ある');
    seen[r.code] = true;
  });
  const rows = p.rows.map(function (r) {
    const o = {};
    SHEETS.delay.forEach(function (k) { o[k] = r[k]; });
    return o;
  });
  rows.forEach(function (r) {
    const sum = addSafe_(addSafe_(addSafe_(r.m3, r.m2), r.m1), r.m0);
    if (sum !== r.total) throw fail_('TOTAL_MISMATCH', 'total が4列の合計と違う: code ' + r.code);
  });
  const sums = delaySums_(rows); // 8項目の合計。安全な整数を超えたら BAD_NUMBER
  const full = delayFullDigest_(rows);
  if (typeof p.full !== 'string' || p.full !== full) throw fail_('FULL_MISMATCH', 'Python と GAS の全文指紋が違う');
  const source = (p.source === null || p.source === undefined) ? '' : String(p.source);

  return withLock_(function () {
    const ss = db_();
    // ② 前置き
    const w = beginWrite_(ss);
    const cur = w.st;
    // ③ 同じ基準日の記録があり、内容が同じなら変更なし
    const old = (cur.state.delays || {})[asOf];
    if (old && old.full === full) return { success: true, changed: false, seq: cur.seq };
    // ④ 新しいシートに書いて読み戻す
    const base = 'delay_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    let name = base;
    for (let n = 2; ss.getSheetByName(name); n++) {
      if (n > 99) throw fail_('NAME_EXHAUSTED');
      name = base + '_' + n;
    }
    const sh = ss.insertSheet(name);
    ensureRows_(sh, rows.length + 1);
    const rng = sh.getRange(1, 1, rows.length + 1, SHEETS.delay.length);
    rng.setNumberFormat('@');
    rng.setValues([SHEETS.delay].concat(rows.map(function (r) { return SHEETS.delay.map(function (k) { return String(r[k]); }); })));
    SpreadsheetApp.flush();
    let back;
    try {
      back = loadDelaySheet_(sh, asOf);
    } catch (e) {
      if (e.message === 'DELAY_UNREADABLE') throw fail_('VERIFY_FAILED', '新しい delay シートが読み戻せない');
      throw e;
    }
    const act = delayActual_(back);
    if (act.rows !== rows.length || act.full !== full) throw fail_('VERIFY_FAILED', '新しい delay シートの件数か指紋が違う');
    // ⑤ 新しい状態（rows・imports・除外・目標は変えない）
    const next = copyState_(cur.state);
    next.delays = next.delays || {};
    next.delays[asOf] = {
      asOf: asOf, sheet: name, rows: rows.length, full: full, months: months, sums: sums,
      source: source, importedAt: w.at
    };
    // ⑥ 確定
    const before = old ? old.rows + '件/' + String(old.full).slice(0, 8) : '';
    const after = rows.length + '件/' + full.slice(0, 8);
    return commit_(ss, cur, 'delay_import', 'delay×' + asOf, before, after, source, 'import', next, w.at, null);
  });
}

// ============================================================
// 掃除（取り込みの確定の後だけ。実装計画 §4-5-8。delay_ シートは消さない＝v1.2 R4）
// ============================================================
function rowsSheetTime_(name) {
  const m = /^rows_([0-9]{4})([0-9]{2})([0-9]{2})_([0-9]{2})([0-9]{2})([0-9]{2})/.exec(name);
  if (!m) return null;
  return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]) - 9, Number(m[5]), Number(m[6]));
}

function cleanupRows_(ss) {
  let st, chain;
  try {
    st = readState_(ss);
    chain = loadChain_(ss, st);
  } catch (e) {
    Logger.log('掃除を中止（' + e.message + '）。rows シートは1枚も消していません');
    return;
  }
  const keep = {};
  let n = 0;
  chain.forEach(function (c) {
    if (c.rows && !keep[c.rows] && n < KEEP_CHAIN_SHEETS) { keep[c.rows] = true; n++; }
  });
  if (st.rows) keep[st.rows] = true;
  const now = Date.now();
  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    if (!ROWS_NAME_RE.test(name) || keep[name]) return;
    const t = rowsSheetTime_(name);
    if (t === null || now - t < KEEP_RECENT_MS) return;
    try {
      ss.deleteSheet(sh);
      Logger.log('rows シートを削除: ' + name);
    } catch (e) {
      Logger.log('rows シートの削除に失敗: ' + name + ' ' + e);
    }
  });
}

// ============================================================
// 復旧・巻き戻しの候補の判定（診断・読み戻しで共通。実装計画 v1.2 R8）
// ============================================================
// history シートの row 行の状態 JSON が正しく、rows と遅延の両方のシートがそろうか。正しければ ''、外れたら理由
function restorableReason_(ss, row, rowsCol) {
  try {
    const hs = ss.getSheetByName('history');
    const T = JSON.parse(String(hs.getRange(row, 11).getValue()));
    const why = stateInvalid_(T);
    if (why) return why;
    if (T.rows !== rowsCol) return '状態 JSON の rows が記録行と違う';
    return stateSheetsOk_(ss, T);
  } catch (e) {
    return '状態 JSON が読めない';
  }
}

// ============================================================
// 回収遅延の読み戻し（1つの基準日分。実装計画 v1.2 §3 importStatus_・§1-3 の表）
// rec は状態に登録された記録。無ければ未確定として返す。例外にしない（読めないときは readable:false）
// ============================================================
function delayCheck_(ss, asOf, rec) {
  const c = { asOf: asOf, committed: !!rec, readable: null, consistent: false,
              actualRows: null, actualFull: null,
              recordRows: rec ? rec.rows : null, recordFull: rec ? rec.full : null, error: null };
  if (!rec) return c;
  try {
    const rows = loadDelaySheet_(ss.getSheetByName(rec.sheet), asOf);
    const act = delayActual_(rows);
    c.readable = true;
    c.actualRows = act.rows;
    c.actualFull = act.full;
    c.consistent = act.rows === rec.rows && act.full === rec.full;
  } catch (e) {
    if (e.message !== 'DELAY_UNREADABLE') throw e;
    c.readable = false;
    c.error = 'DELAY_UNREADABLE';
  }
  return c;
}

// ============================================================
// 読み戻し（実装計画 §4-5-6）
// ============================================================
function importStatus_(p) {
  // 基準日の指定があれば、形の確認は先に行う（形が違えば BAD_ASOF）
  const wantAsOf = (p.asOf === undefined || p.asOf === null) ? null : p.asOf;
  if (wantAsOf !== null && !dateReal_(wantAsOf)) throw fail_('BAD_ASOF', 'status の基準日の形');
  return withLock_(function () {
    const ss = db_();
    const st = readState_(ss);
    const chain = loadChain_(ss, st);
    const lg = syncLedger_(ss, st);
    const rows = st.rows === '' ? [] : readRowsSheet_(st.rowsSheet, 'STATE_BROKEN');
    const byKey = {};
    rows.forEach(function (r) { (byKey[r.month + '|' + r.close] = byKey[r.month + '|' + r.close] || []).push(r); });
    let keys;
    if (p.month && p.close) {
      keys = [String(p.month) + '|' + String(p.close)];
    } else {
      const seen = {};
      Object.keys(st.state.imports).concat(Object.keys(byKey)).forEach(function (k) { seen[k] = true; });
      keys = Object.keys(seen).sort();
    }
    const checks = keys.map(function (k) {
      const parts = k.split('|');
      const actual = byKey[k] || [];
      const rec = st.state.imports[k] || null;
      const c = {
        month: parts[0], close: parts[1], committed: !!rec, consistent: false,
        actualRows: actual.length,
        actualDigest: actual.length ? digest_(actual) : null,
        actualFull: actual.length ? fullDigest_(actual) : null,
        recordRows: rec ? rec.rows : null,
        recordDigest: rec ? rec.digest : null,
        recordFull: rec ? rec.full : null
      };
      c.consistent = c.committed && c.actualRows === c.recordRows && c.actualDigest === c.recordDigest && c.actualFull === c.recordFull;
      return c;
    });
    // 回収遅延（登録された基準日。新しい順。未取り込みなら [] と {}）
    const delays = st.state.delays || {};
    const delayDates = Object.keys(delays).sort().reverse();
    const delayChecks = {};
    (wantAsOf !== null ? [wantAsOf] : delayDates).forEach(function (d) {
      delayChecks[d] = delayCheck_(ss, d, Object.prototype.hasOwnProperty.call(delays, d) ? delays[d] : null);
    });
    // 巻き戻し先に使える版: rollback_ が受け付ける版だけ（rows と遅延の両方のシートがそろう）
    const rollbackable = [];
    for (let i = 0; i < chain.length && rollbackable.length < 10; i++) {
      const c = chain[i];
      if (!restorableReason_(ss, c.row, c.rows)) rollbackable.push({ seq: c.seq, at: c.at, action: c.action, target: c.target });
    }
    return {
      success: true, version: VERSION, stateSeq: st.seq, stateAt: st.at, stateChars: st.stateChars,
      ledgerSeq: lg.ledgerSeq, ledgerSynced: lg.synced, ledgerMatchesState: lg.matches,
      checks: checks, delayDates: delayDates, delayChecks: delayChecks, rollbackable: rollbackable
    };
  });
}

// ============================================================
// 巻き戻し（実装計画 §4-5-7）
// ============================================================
function exKeySet_(st) { return Object.keys(st.excluded).sort().join(','); }
function targetsKey_(st) { return VIEWS.map(function (v) { return String(st.targets[v]); }).join(','); }

// キーの並びに左右されない JSON（imports の比較用）
function stableJson_(v) {
  if (Array.isArray(v)) return '[' + v.map(stableJson_).join(',') + ']';
  if (isObj_(v)) {
    return '{' + Object.keys(v).sort().map(function (k) { return JSON.stringify(k) + ':' + stableJson_(v[k]); }).join(',') + '}';
  }
  return JSON.stringify(v);
}

function rollback_(p) {
  const toSeq = typeof p.toSeq === 'string' && /^[1-9][0-9]*$/.test(p.toSeq) ? Number(p.toSeq) : p.toSeq;
  if (!isPosInt_(toSeq)) throw fail_('BAD_SEQ');
  const reason = (p.reason === null || p.reason === undefined) ? '' : String(p.reason);
  return withLock_(function () {
    const ss = db_();
    const w = beginWrite_(ss);
    const C = w.st.state;
    const entry = w.chain.filter(function (c) { return c.seq === toSeq; })[0];
    if (!entry) throw fail_('BAD_SEQ', '確定の系列に #' + toSeq + ' が無い');
    const hs = ss.getSheetByName('history');
    let T;
    try { T = JSON.parse(String(hs.getRange(entry.row, 11).getValue())); } catch (e) { throw fail_('ROLLBACK_UNAVAILABLE', '#' + toSeq + ' の状態 JSON が読めない'); }
    const why = stateInvalid_(T);
    if (why) throw fail_('ROLLBACK_UNAVAILABLE', '#' + toSeq + ' の状態が正しくない: ' + why);
    if (T.rows !== entry.rows) throw fail_('ROLLBACK_UNAVAILABLE', '#' + toSeq + ' の状態の rows が記録行と違う');
    const noSheet = stateSheetsOk_(ss, T);
    if (noSheet) throw fail_('ROLLBACK_UNAVAILABLE', '#' + toSeq + ' のシートが消えている: ' + noSheet);

    const next = copyState_(C);
    const parts = [];
    next.rows = T.rows;
    next.rowsCount = T.rowsCount;
    next.imports = copyState_(T.imports);
    next.delays = copyState_(T.delays || {});
    if (T.rows !== C.rows || stableJson_(T.imports) !== stableJson_(C.imports)) parts.push('データ');
    if (stableJson_(T.delays || {}) !== stableJson_(C.delays || {})) parts.push('遅延');
    if (exKeySet_(T) !== exKeySet_(C)) {
      next.excluded = copyState_(T.excluded);
      next.excludedAt = w.at;
      parts.push('除外');
    }
    if (targetsKey_(T) !== targetsKey_(C)) {
      next.targets = copyState_(T.targets);
      next.targetsAt = w.at;
      parts.push('目標');
    }
    if (!parts.length) return { success: true, changed: false, seq: w.st.seq };
    return commit_(ss, w.st, 'rollback', '#' + toSeq, '#' + w.st.seq, '#' + toSeq + 'の状態（' + parts.join('・') + '）',
                   reason, 'import', next, w.at, null);
  });
}

// ============================================================
// 管理画面の書き込み
// ============================================================
function setExcluded_(code, on, note, byName) {
  if (typeof code !== 'string' || !/^[0-9]+$/.test(code)) throw fail_('BAD_CODE');
  return withLock_(function () {
    const ss = db_();
    const w = beginWrite_(ss);
    const C = w.st.state;
    const has = Object.prototype.hasOwnProperty.call(C.excluded, code);
    if (on === has) return { success: true, changed: false, seq: w.st.seq }; // メモだけの違いも同じとみなす
    const next = copyState_(C);
    next.excludedAt = w.at;
    if (on) {
      const rows = w.st.rows === '' ? [] : readRowsSheet_(w.st.rowsSheet, 'STATE_BROKEN');
      let hit = null;
      rows.forEach(function (r) {
        if (r.code !== code) return;
        if (!hit || r.month > hit.month || (r.month === hit.month && r.close === '末日')) hit = r;
      });
      if (!hit) throw fail_('CODE_NOT_FOUND');
      const memo = note.slice(0, 100);
      next.excluded[code] = { name: hit.name, note: memo, at: w.at, by: byName };
      return commit_(ss, w.st, 'exclude_on', code, '', '除外: ' + memo, memo, byName, next, w.at, null);
    }
    const oldNote = String(C.excluded[code].note || '');
    delete next.excluded[code];
    return commit_(ss, w.st, 'exclude_off', code, '除外: ' + oldNote, '', note.slice(0, 100), byName, next, w.at, null);
  });
}

function setTarget_(view, rate, byName) {
  if (VIEWS.indexOf(view) < 0) throw fail_('BAD_VIEW');
  // 空かどうかは数値にする前に判定する（Number('') は 0 になる）
  const empty = rate === null || rate === undefined || (typeof rate === 'string' && rate.trim() === '');
  let val = null;
  if (empty) {
    if (view === '合算') throw fail_('BAD_RATE');
  } else {
    if (typeof rate !== 'number' && typeof rate !== 'string') throw fail_('BAD_RATE');
    const n = Number(typeof rate === 'string' ? rate.trim() : rate);
    if (!isFinite(n) || n < 1 || n > 150) throw fail_('BAD_RATE'); // 丸める前の値で判定
    val = Math.round(n * 10) / 10;
  }
  return withLock_(function () {
    const ss = db_();
    const w = beginWrite_(ss);
    const C = w.st.state;
    const old = C.targets[view];
    if (old === val) return { success: true, changed: false, seq: w.st.seq };
    const next = copyState_(C);
    next.targets[view] = val;
    next.targetsAt = w.at;
    return commit_(ss, w.st, val === null ? 'target_clear' : 'target_set', view,
                   old === null ? '' : String(old), val === null ? '' : String(val), '', byName, next, w.at, null);
  });
}

// ============================================================
// 閲覧（実装計画 §4-5-5）。キャッシュがあっても毎回、確定状態と系列を検査する
// ============================================================
function cacheGetChunks_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(key + '_n'));
    if (!Number.isInteger(n) || n < 1 || n > CACHE_MAX_CHUNKS) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const got = cache.getAll(keys);
    let s = '';
    for (let i = 0; i < n; i++) {
      const part = got[keys[i]];
      if (part === undefined || part === null) return null;
      s += part;
    }
    return JSON.parse(s);
  } catch (e) {
    Logger.log('集計キャッシュの読み込みを省略: ' + e);
    return null;
  }
}

function cachePutChunks_(key, obj) {
  try {
    const s = JSON.stringify(obj);
    const n = Math.ceil(s.length / CACHE_CHUNK_CHARS) || 1;
    if (n > CACHE_MAX_CHUNKS) { Logger.log('集計が大きいのでキャッシュしない: ' + s.length + '文字'); return; }
    const map = {};
    map[key + '_n'] = String(n);
    for (let i = 0; i < n; i++) map[key + '_' + i] = s.slice(i * CACHE_CHUNK_CHARS, (i + 1) * CACHE_CHUNK_CHARS);
    CacheService.getScriptCache().putAll(map, CACHE_TTL_DATA);
  } catch (e) {
    Logger.log('集計キャッシュの保存を省略: ' + e);
  }
}

function getData_(opts) {
  opts = opts || {};
  const t0 = Date.now();
  const raws = [];
  let lastErr = null;
  for (let attempt = 0; attempt < 3; attempt++) {
    if (attempt > 0) Utilities.sleep(500);
    try {
      // この試行で読む値はすべてここから読み直す
      const raw = PropertiesService.getScriptProperties().getProperty('STATE');
      raws.push(raw);
      const ss = db_();
      const st = readState_(ss, raw);
      let history = [], historyOk = true;
      try {
        history = loadChain_(ss, st).slice(0, HISTORY_LIMIT).map(histOut_);
      } catch (e) {
        if (e.message !== 'HISTORY_BROKEN') throw e;
        historyOk = false;
        history = [];
      }
      const key = 'cr_data_' + VERSION + '_' + st.seq + '_' + st.at;
      let agg = opts.nocache ? null : cacheGetChunks_(key);
      const cacheUse = agg ? 'hit' : 'miss';
      if (!agg) {
        const rows = st.rows === '' ? [] : readRowsSheet_(st.rowsSheet, 'STATE_BROKEN');
        agg = buildData_(rows, st.state);
        cachePutChunks_(key, agg);
      }
      const S = st.state;
      return {
        success: true,
        version: VERSION,
        generatedAt: new Date().toISOString(),
        stateSeq: st.seq,
        stateAt: st.at,
        months: agg.months,
        targets: { '合算': S.targets['合算'], '末日': S.targets['末日'], '15日': S.targets['15日'] },
        byMonth: agg.byMonth,
        excluded: Object.keys(S.excluded).sort().map(function (c) {
          const x = S.excluded[c];
          return { code: c, name: String(x.name), note: String(x.note || ''), updatedAt: x.at || null, updatedBy: x.by || '' };
        }),
        excludedUpdatedAt: S.excludedAt,
        targetsUpdatedAt: S.targetsAt,
        imports: Object.keys(S.imports).map(function (k) { return S.imports[k]; }).sort(function (a, b) {
          if (a.month !== b.month) return a.month < b.month ? -1 : 1;
          return a.close < b.close ? -1 : (a.close > b.close ? 1 : 0);
        }).map(function (i) {
          return { month: i.month, close: i.close, closeDate: i.closeDate, source: i.source, rows: i.rows, importedAt: i.importedAt };
        }),
        customers: agg.customers,
        history: history,
        historyOk: historyOk,
        timing: { ms: Date.now() - t0, cache: cacheUse }
      };
    } catch (e) {
      lastErr = e;
      Logger.log('閲覧の読み込みに失敗（' + (attempt + 1) + '回目）: ' + e);
    }
  }
  if (raws.some(function (r) { return r !== raws[0]; })) throw fail_('DATA_UNAVAILABLE', '読み込みの間に版が変わった');
  if (lastErr && lastErr.message === 'STATE_BROKEN') throw fail_('STATE_BROKEN');
  throw fail_('DATA_UNAVAILABLE', String(lastErr));
}

// ============================================================
// 回収遅延の閲覧（実装計画 v1.2 §3-2）。getData_ と同じ作り: ロックなし・3回まで試行・1回の試行の値だけで応答を作る
// ============================================================
// 遅延シートを読んで記録と突き合わせる（読めない・違う → STATE_BROKEN）
function readDelayVerified_(ss, rec) {
  let rows;
  try {
    rows = loadDelaySheet_(ss.getSheetByName(rec.sheet), rec.asOf);
  } catch (e) {
    if (e.message === 'DELAY_UNREADABLE') throw fail_('STATE_BROKEN', '遅延シートが読めない: ' + rec.asOf);
    throw e;
  }
  const act = delayActual_(rows);
  if (act.rows !== rec.rows || act.full !== rec.full) throw fail_('STATE_BROKEN', '遅延シートの件数か指紋が記録と違う: ' + rec.asOf);
  return rows;
}

function getDelay_(opts) {
  opts = opts || {};
  const t0 = Date.now();
  // 入力の誤りは試行の前に確かめて即座に返す（再試行しない）
  const want = (opts.asOf === undefined || opts.asOf === null) ? null : opts.asOf;
  if (want !== null && !delayAsOfOk_(want)) throw fail_('BAD_ASOF', '基準日の形・実在・開始日');
  const raws = [];
  let lastErr = null;
  for (let attempt = 0; attempt < 3; attempt++) {
    if (attempt > 0) Utilities.sleep(500);
    try {
      // この試行で読む値はすべてここから読み直す
      const raw = PropertiesService.getScriptProperties().getProperty('STATE');
      raws.push(raw);
      const ss = db_();
      const st = readState_(ss, raw);
      let historyOk = true;
      try {
        loadChain_(ss, st);
      } catch (e) {
        if (e.message !== 'HISTORY_BROKEN') throw e;
        historyOk = false;
      }
      const S = st.state;
      const delays = S.delays || {};
      const list = Object.keys(delays).sort().reverse();
      let sel = null;
      if (want !== null) {
        if (!Object.prototype.hasOwnProperty.call(delays, want)) throw fail_('BAD_ASOF', '登録の無い基準日');
        sel = want;
      } else if (list.length) {
        sel = list[0];
      }
      const prev = sel ? (list[list.indexOf(sel) + 1] || null) : null;
      let cacheUse = 'miss', body = { rows: [], prevRows: [] };
      if (sel) {
        const key = 'cr_delay_' + VERSION + '_' + st.seq + '_' + st.at + '_' + sel;
        let c = opts.nocache ? null : cacheGetChunks_(key);
        if (c && !(isObj_(c) && Array.isArray(c.rows) && Array.isArray(c.prevRows))) c = null;
        if (c) {
          cacheUse = 'hit';
          body = c;
        } else {
          const rows = readDelayVerified_(ss, delays[sel]);
          const prevRows = prev ? readDelayVerified_(ss, delays[prev]).map(function (r) {
            return { rep: r.rep, code: r.code, name: r.name, m3: r.m3, m2: r.m2, m1: r.m1, m0: r.m0, bal: r.bal };
          }) : [];
          body = { rows: rows.sort(cmpDelayCode_), prevRows: prevRows.sort(cmpDelayCode_) };
          cachePutChunks_(key, body);
        }
      }
      return {
        success: true,
        version: VERSION,
        stateSeq: st.seq,
        stateAt: st.at,
        historyOk: historyOk,
        dates: list.map(function (d) {
          const r = delays[d];
          return { asOf: d, rows: r.rows, importedAt: r.importedAt, source: r.source, months: r.months };
        }),
        asOf: sel,
        prevAsOf: prev,
        months: sel ? delays[sel].months : null,
        rows: body.rows,
        prevRows: body.prevRows,
        excluded: Object.keys(S.excluded).sort(),
        timing: { ms: Date.now() - t0, cache: cacheUse }
      };
    } catch (e) {
      if (e.message === 'BAD_ASOF') throw e; // 登録の無い日は再試行しない
      lastErr = e;
      Logger.log('遅延の閲覧の読み込みに失敗（' + (attempt + 1) + '回目）: ' + e);
    }
  }
  if (raws.some(function (r) { return r !== raws[0]; })) throw fail_('DATA_UNAVAILABLE', '読み込みの間に版が変わった');
  if (lastErr && lastErr.message === 'STATE_BROKEN') throw fail_('STATE_BROKEN');
  throw fail_('DATA_UNAVAILABLE', String(lastErr));
}

// rows（確定状態の rows シートの行）と状態 JSON（除外先）から、キャッシュに入れる集計を作る
function buildData_(rowsRaw, state) {
  const nonSales = (PropertiesService.getScriptProperties().getProperty('NON_SALES_REPS') || '店売り')
    .split(',').map(function (s) { return s.trim(); }).filter(String);
  const ex = state.excluded || {};
  const isEx = function (code) { return Object.prototype.hasOwnProperty.call(ex, code); };

  // ① 月 → 締め → repCode に振り分けて系統ビューの数字と明細を作る
  const sys = {};
  rowsRaw.forEach(function (r) {
    const m = String(r.month);
    if (m < START_MONTH) return;
    const close = String(r.close);
    const repCode = String(r.repCode);
    const code = String(r.code);
    const M = sys[m] = sys[m] || {};
    const C = M[close] = M[close] || {};
    const R = C[repCode] = C[repCode] || { rep: String(r.rep), repCode: repCode, prev: 0, paid: 0, exCarry: 0, exPaid: 0, exCount: 0, details: [] };
    const prev = Number(r.prev), paid = Number(r.paid), carry = Number(r.carry);
    const x = isEx(code);
    if (x) {
      R.exCarry += carry; R.exPaid += paid; R.exCount++;
    } else {
      R.prev += prev; R.paid += paid;
    }
    if ((prev > 0 && paid < prev) || (x && carry !== 0)) {
      R.details.push({ code: code, name: String(r.name), close: close, prev: prev, paid: paid, carry: carry, ex: x });
    }
  });

  const rate = function (o) { return o.prev > 0 ? o.paid / o.prev : null; };
  const byRepCode = function (a, b) { return Number(a.repCode) - Number(b.repCode); };
  const makeView = function (reps) {
    const company = { prev: 0, paid: 0, rate: null, exCarry: 0, exPaid: 0 };
    reps.sort(byRepCode);
    reps.forEach(function (R) {
      R.rate = rate(R);
      company.prev += R.prev; company.paid += R.paid; company.exCarry += R.exCarry; company.exPaid += R.exPaid;
    });
    company.rate = rate(company);
    return { company: company, reps: reps };
  };

  const months = Object.keys(sys).sort();
  const byMonth = {};
  months.forEach(function (m) {
    const closes = CLOSES.filter(function (c) { return !!sys[m][c]; }).sort();
    const views = { '合算': null, '末日': null, '15日': null };
    CLOSES.forEach(function (c) {
      if (!sys[m][c]) return;
      const reps = Object.keys(sys[m][c]).map(function (k) {
        const R = sys[m][c][k];
        R.details.sort(function (a, b) { return b.carry - a.carry || (a.code < b.code ? -1 : (a.code > b.code ? 1 : 0)); });
        return { rep: R.rep, repCode: R.repCode, sales: nonSales.indexOf(R.rep) < 0, prev: R.prev, paid: R.paid, rate: null,
                 exCarry: R.exCarry, exPaid: R.exPaid, exCount: R.exCount, details: R.details };
      });
      views[c] = makeView(reps);
    });
    // ② 系統ビューの数字を repCode ごとに足して合算ビュー（details は付けない）
    const all = {};
    ['末日', '15日'].forEach(function (c) {
      if (!views[c]) return;
      views[c].reps.forEach(function (R) {
        const A = all[R.repCode] = all[R.repCode] || { rep: R.rep, repCode: R.repCode, prev: 0, paid: 0, rate: null, exCarry: 0, exPaid: 0, exCount: 0 };
        A.prev += R.prev; A.paid += R.paid; A.exCarry += R.exCarry; A.exPaid += R.exPaid; A.exCount += R.exCount;
      });
    });
    const allReps = Object.keys(all).map(function (k) {
      const A = all[k];
      return { rep: A.rep, repCode: A.repCode, sales: nonSales.indexOf(A.rep) < 0, prev: A.prev, paid: A.paid, rate: null,
               exCarry: A.exCarry, exPaid: A.exPaid, exCount: A.exCount };
    });
    views['合算'] = makeView(allReps);
    byMonth[m] = { closes: closes, complete: CLOSES.every(function (c) { return closes.indexOf(c) >= 0; }), views: views };
  });

  // ⑤ 管理画面の検索用: 全期間・得意先コードごとに1件（最後に載った月。同じ月に両系統あれば末日〆）
  const last = {};
  rowsRaw.forEach(function (r) {
    const code = String(r.code), m = String(r.month), c = String(r.close);
    if (m < START_MONTH) return;
    const L = last[code];
    if (!L || m > L.month || (m === L.month && c === '末日' && L.close !== '末日')) last[code] = r;
  });
  const customers = Object.keys(last).map(function (code) {
    const r = last[code];
    return { code: code, name: String(r.name), rep: String(r.rep), repCode: String(r.repCode),
             lastMonth: String(r.month), lastClose: String(r.close), carry: Number(r.carry) };
  });
  customers.sort(function (a, b) { return b.carry - a.carry || (a.code < b.code ? -1 : (a.code > b.code ? 1 : 0)); });

  return { months: months, byMonth: byMonth, customers: customers };
}

// ============================================================
// 診断（読むだけ。実装計画 §7-2）
// 得意先名・担当者名・金額・状態 JSON の中身は返さない
// ============================================================
function diagnose_(p) {
  const ss = db_();
  const out = { success: true, version: VERSION };
  const hs = ss.getSheetByName('history');
  const cols = hs ? readHistoryCols_(hs) : [];
  const lastRow = hs ? hs.getLastRow() : 0;

  // 1行の変更記録の列（1〜10列目）
  const histAt = function (row) { return (row >= 2 && row <= lastRow) ? cols[row - 2].map(function (c) { return String(c); }) : null; };
  // その行の状態 JSON が正しく、rows シートがある（'' なら不要）か
  const restorable = function (row, rowsCol) { return restorableReason_(ss, row, rowsCol); };
  // 根拠の行（台帳・E1）を今の DB で照合する。mode: 'values'＝4つの値だけ、'chain'＝＋系列、'full'＝＋状態 JSON と rows シート
  const verify = function (E, mode) {
    if (!/^[1-9][0-9]*$/.test(String(E.seq))) return { ok: false, reason: 'seq が正の整数でない' };
    if (!/^[1-9][0-9]*$/.test(String(E.row)) || Number(E.row) < 2) return { ok: false, reason: 'row が2以上の整数でない' };
    const h = histAt(Number(E.row));
    if (!h) return { ok: false, reason: '変更記録にその行が無い' };
    if (h[0] !== String(E.seq)) return { ok: false, reason: '変更記録の seq が違う' };
    if (h[9] !== String(E.rows)) return { ok: false, reason: '変更記録の rows が違う' };
    if (h[2] !== String(E.at)) return { ok: false, reason: '変更記録の at が違う' };
    if (mode === 'values') return { ok: true, reason: null };
    const full = mode === 'full';
    let chain;
    try { chain = buildChain_(cols, { seq: Number(E.seq), row: Number(E.row) }); } catch (e) { return { ok: false, reason: '系列の検査が通らない: ' + (e.reason || e.message) }; }
    if (full) {
      const why = restorable(Number(E.row), h[9]);
      if (why) return { ok: false, reason: why };
    }
    return { ok: true, reason: null, chain: chain };
  };
  const brief = function (c) { return { seq: c.seq, at: c.at, action: c.action, target: c.target, by: c.by }; };

  // state
  const raw = PropertiesService.getScriptProperties().getProperty('STATE');
  let value = null;
  try { const S = JSON.parse(raw); if (isObj_(S)) value = { seq: S.seq, row: S.row, rows: S.rows, at: S.at }; } catch (e) {}
  let insp;
  try { insp = inspectState_(ss, raw); } catch (e) { insp = { ok: false, reason: '読み込み中の例外' }; }
  out.state = { present: !!raw, value: raw ? value : null, ok: insp.ok, reason: insp.ok ? null : insp.reason };

  // chain
  if (!insp.ok) {
    out.chain = { ok: false, seq: null, reason: '確定状態が読めないので検査していません' };
  } else {
    try { buildChain_(cols, insp); out.chain = { ok: true, seq: null, reason: null }; }
    catch (e) { out.chain = { ok: false, seq: e.seq || null, reason: e.reason || e.message }; }
  }

  // ledger
  const ledger = ledgerRows_(ss);
  out.ledger = ledger.slice(-50).reverse().map(function (L) {
    const v = verify(L, 'full');
    return { seq: L.seq, row: L.row, rows: L.rows, at: L.at, recordedAt: L.recordedAt, verified: v.ok, reason: v.reason };
  });

  // latestEvidence（書いた順番で最後の行）と candidate（その系列の中で戻せる一番新しい版）
  const L = ledger.length ? ledger[ledger.length - 1] : null;
  let latestChain = null;
  out.latestEvidence = null;
  out.candidate = null;
  out.lostConfirmed = [];
  if (L) {
    const v = verify(L, 'chain');
    out.latestEvidence = { seq: L.seq, row: L.row, rows: L.rows, at: L.at, recordedAt: L.recordedAt, verified: v.ok, reason: v.reason };
    if (v.ok) {
      latestChain = v.chain;
      let idx = -1;
      for (let i = 0; i < latestChain.length; i++) {
        if (!restorable(latestChain[i].row, latestChain[i].rows)) { idx = i; break; }
      }
      if (idx >= 0) {
        const c = latestChain[idx];
        out.candidate = { seq: c.seq, row: c.row, rows: c.rows, at: c.at,
                          stateValue: JSON.stringify({ seq: c.seq, row: c.row, rows: c.rows, at: c.at }) };
      }
      out.lostConfirmed = latestChain.slice(0, idx >= 0 ? idx : latestChain.length).map(brief);
    }
  }

  // undetermined: 最新の根拠の系列に入っておらず、at が根拠を台帳に記録した日時より後の行
  const inChain = {};
  (latestChain || []).forEach(function (c) { inChain[c.row] = true; });
  const since = L ? new Date(L.recordedAt).getTime() : -Infinity;
  out.undetermined = [];
  cols.forEach(function (c, i) {
    const row = i + 2;
    const s = String(c[0]);
    if (!/^[1-9][0-9]*$/.test(s) || inChain[row]) return;
    const t = new Date(String(c[2])).getTime();
    if (!(isNaN(t) || t > since)) return;
    out.undetermined.push({ seq: Number(s), at: String(c[2]), action: String(c[4]), target: String(c[5]), by: String(c[3]) });
  });

  // otherBranches: 台帳の前の行のうち、採用中の枝に入っていない版（自動の候補にしない）
  const chainKey = {};
  (latestChain || []).forEach(function (c) { chainKey[c.seq + '|' + c.at] = true; });
  const seen = {};
  out.otherBranches = [];
  ledger.slice(0, -1).forEach(function (E) {
    const k = E.seq + '|' + E.at;
    if (chainKey[k] || seen[k]) return;
    seen[k] = true;
    out.otherBranches.push({ seq: /^[1-9][0-9]*$/.test(E.seq) ? Number(E.seq) : E.seq, at: E.at, branch: '以前の枝' });
  });

  // sheets
  const nowRefs = {};
  if (insp.ok) Object.keys(insp.state.delays || {}).forEach(function (k) { nowRefs[insp.state.delays[k].sheet] = true; });
  out.sheets = ss.getSheets().filter(function (sh) { return ROWS_NAME_RE.test(sh.getName()) || DELAY_NAME_RE.test(sh.getName()); }).map(function (sh) {
    const name = sh.getName(), rowsN = Math.max(0, sh.getLastRow() - 1);
    if (!DELAY_NAME_RE.test(name)) return { name: name, rows: rowsN };
    // delay_ シートは削除の判断に使わない（v1.2 R4）。「今の状態から参照されているか」だけを示す
    const ref = insp.ok ? !!nowRefs[name] : null;
    return { name: name, rows: rowsN, referenced: ref,
             note: ref === null ? '確定状態が読めないので判定していません' : (ref ? '今の状態から参照されている' : '今の状態からは参照されていない') };
  });

  // e1（照合できるかと、最新の根拠と同じ版か。親子関係は返さない）
  if (p.e1) {
    const E = isObj_(p.e1) ? p.e1 : {};
    const v = verify({ seq: String(E.seq), row: String(E.row), rows: String(E.rows), at: String(E.at) }, 'values');
    out.e1 = {
      verified: v.ok,
      sameAsLatestEvidence: !!L && String(E.seq) === L.seq && String(E.row) === L.row && String(E.rows) === L.rows && String(E.at) === L.at
    };
  }
  return out;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
