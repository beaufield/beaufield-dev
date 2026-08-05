// ============================================================
// ビューフェス申込アプリ - Google Apps Script
// Version: 0.6.1
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   SPREADSHEET_ID  : ビューフェス申込データのスプレッドシートID（新規に「beaufes2026」という名前で作成する）
//
// 初回セットアップ手順:
//   1. 新規Googleスプレッドシートを作成（名前: beaufes2026）
//   2. 拡張機能 → Apps Script でこのファイルの内容を貼り付け
//   3. プロジェクトの設定でスクリプトプロパティ SPREADSHEET_ID を設定
//   4. GASエディタの関数選択で setupSheets を選び、一度だけ実行（シート・見出し・configの初期値を作成）
//   5. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
//        実行ユーザー: 自分（tak.maejima@gmail.com。スプレッドシート作成に使ったGoogleアカウント）
//        アクセスできるユーザー: 全員（お客様の申込フォームが認証なしで動く必要があるため）
//   6. 発行されたウェブアプリURLを index.html / pass.html の GAS_URL に設定
//
// メール送信元（beaufes@gmail.com）を実際に使うには、事前に
// tak.maejima@gmail.com のGmail設定で「他のメールアドレスを追加」から
// beaufes@gmail.com を送信元アドレス（send-as）として登録しておくこと。
// 未登録の場合は自動的に MailApp（送信元は tak.maejima@ のまま・差出人名とReply-Toで代替）にフォールバックする。
//
// 🔴 送信量の注意（2026-08-04 実測確認済み）: tak.maejima@gmail.com はGoogle Workspace Business
// Standardを契約しているにもかかわらず、Apps Scriptの送信枠は個人アカウント扱いで1日100通だった
//（MailApp.getRemainingDailyQuota()で実測確認済み。契約と実測の食い違いの原因は未特定・追跡はせず
// 実測値を設計上の正として採用。checkMailQuota()関数でいつでも再確認できる）。
// 申込受付中の分散送信は問題ないが、前日リマインド等で200名に一括送信する機能を作る際は、
// 50〜80件ずつ複数回に分けて送るなどの対策が必須。
// 詳細: LINEHarness/ビューフェス申込_設計.md §7-0-1〜7-0-2
//
// 🆕 v0.5.0（2026-08-05・設計書v3対応）: applications シートに business_type 列（U列=21列目）を追加。
// 既存の本番シート（beaufes2026）には自動で列が増えないため、デプロイ後に一度だけ
// migrateAddBusinessType() をGASエディタから手動実行してヘッダーを追加すること。
// 新規にsetupSheetsでシートを作る場合はヘッダーに最初から含まれるため不要。
//
// 🆕 v0.6.0（2026-08-06）: エリア項目をフォームから廃止（用途は地域分布の集計のみで、
// お客様に選ばせる必要がないと判断。Takashiさん指示）。J列(area)は列位置維持のため
// 残すが常に空文字を書き込む。詳細: LINEHarness/ビューフェス申込_設計.md §4-1
// 🆕 v0.6.1: 別端末で先行push済みのbusiness_type機能(v0.5.0)とのマージ調整のみ。機能変更なし
// ============================================================

const VERSION  = '0.6.1';
const APP_NAME = 'beaufes';

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS         = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = _PROPS.getProperty('SPREADSHEET_ID');

// シート名定数
const SHEET_APPLICATIONS = 'applications';
const SHEET_SESSIONS     = 'sessions';
const SHEET_RESERVATIONS = 'reservations';
const SHEET_CHECKINS     = 'checkins';
const SHEET_MAIL_LOG     = 'mail_log';
const SHEET_CONFIG       = 'config';

// メール送信元（機密ではないため直書きでよい。§7-0-2で確定）
const MAIL_FROM_ADDR = 'beaufes@gmail.com';
const MAIL_FROM_NAME = '株式会社ビューフィールド ビューフェス事務局';

// GitHub Pagesの公開URL（QRコード・メール内のパスリンク生成に使用）
const SITE_BASE_URL = 'https://beaufield.github.io/beaufield-dev/beaufes/';

// ============================================================
// 起動時チェック（プロパティ未設定を早期検知）
// ============================================================
function _checkProps() {
  if (!SPREADSHEET_ID) throw new Error('スクリプトプロパティ SPREADSHEET_ID が未設定です');
}

// ============================================================
// エントリーポイント（GET） ―― 全て認証なし・公開アクション
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  let data = {};
  try {
    if (e && e.parameter && e.parameter.data) data = JSON.parse(e.parameter.data);
  } catch (jsonErr) {
    return _jsonResponse(_err('INVALID_REQUEST'));
  }

  try {
    switch (action) {
      case 'getPass': return _jsonResponse(getPass(data));
      default:         return _jsonResponse(_err('不明なアクション: ' + action));
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return _jsonResponse(_err('INTERNAL_ERROR'));
  }
}

// ============================================================
// エントリーポイント（POST） ―― 全て認証なし・公開アクション
// ============================================================
function doPost(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';
  let data = {};
  try {
    if (params.data) data = JSON.parse(params.data);
  } catch (jsonErr) {
    return _jsonResponse(_err('INVALID_REQUEST'));
  }

  try {
    switch (action) {
      case 'apply': return _jsonResponse(applyApplication(data));
      default:      return _jsonResponse(_err('不明なアクション: ' + action));
    }
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return _jsonResponse(_err('INTERNAL_ERROR'));
  }
}

// ============================================================
// ユーティリティ
// ============================================================
function _jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function _ok(data)  { return { success: true,  data: data }; }
function _err(msg)  { return { success: false, error: msg }; }
function _now() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}
function _getSheet(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name + '（setupSheetsを実行してください）');
  return sh;
}
function _genToken() {
  // QR用ランダム16文字（推測不能・§3設計どおり）
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}
function _escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// 既存app_idの最大連番+1を発番（例: F2026-0007 → F2026-0008）
function _nextAppId(rows) {
  const year = new Date().getFullYear();
  let maxSeq = 0;
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0] || '');
    const m  = id.match(/-(\d+)$/);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxSeq) maxSeq = n;
    }
  }
  const seq = maxSeq + 1;
  return 'F' + year + '-' + ('0000' + seq).slice(-4);
}

// config シートを key/value オブジェクトとして取得
function _getConfig() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = ss.getSheetByName(SHEET_CONFIG);
  const cfg = {};
  if (!sh) return cfg;
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) cfg[String(rows[i][0])] = rows[i][1];
  }
  return cfg;
}

// ============================================================
// ① 申込受付（doPost: action=apply）
//    キーは email_norm + staff_name（§4-3）
//    一致する既存申込があれば更新（内容変更・パス再送）、なければ新規登録
// ============================================================
function applyApplication(data) {
  _checkProps();

  const salonName      = String(data.salon_name || '').trim();
  const staffName      = String(data.staff_name || '').trim();
  const email          = String(data.email || '').trim();
  const phone          = String(data.phone || '').trim();
  const businessType   = String(data.business_type || '').trim(); // 🆕 U列。名札の色分けに使用（§4-1-2）
  const hasTransaction = String(data.has_transaction || '').trim(); // 'yes' / 'no'
  const address        = String(data.address || '').trim();        // 新規客のみ
  const referrer       = String(data.referrer || '').trim();       // 新規客のみ
  const agree          = !!data.agree_capability;
  const note           = String(data.note || '').trim();

  if (!salonName) return _err('サロン名を入力してください');
  if (!staffName) return _err('お名前を入力してください');
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return _err('メールアドレスの形式が正しくありません');
  if (!phone) return _err('お電話番号を入力してください');
  if (!businessType) return _err('業態を選択してください');
  if (!agree) return _err('美容従事者であることの確認にチェックしてください');

  const emailNorm = email.toLowerCase();

  let appId, ticketToken, isUpdate;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh  = _getSheet(ss, SHEET_APPLICATIONS);
    const rows = sh.getDataRange().getValues();
    const now  = _now();

    // 既存申込を探す（email_norm + staff_name 完全一致。§4-3）
    let foundRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][7]).toLowerCase() === emailNorm && String(rows[i][5]) === staffName) {
        foundRow = i + 1; // 1-indexed 行番号
        break;
      }
    }

    if (foundRow > 0) {
      // --- 既存申込を更新（内容変更・パス再送の導線を兼ねる） ---
      isUpdate    = true;
      appId       = rows[foundRow - 1][0];
      ticketToken = rows[foundRow - 1][16];

      sh.getRange(foundRow, 3, 1, 1).setValue(now); // updated_at
      sh.getRange(foundRow, 5, 1, 9).setValues([[
        // J列(area)はフォームから削除済み（2026-08-06）。列位置を保つため空文字を書き続ける
        salonName, staffName, email, emailNorm, phone, '', hasTransaction, address, referrer
      ]]);
      sh.getRange(foundRow, 14, 1, 1).setValue(agree);
      sh.getRange(foundRow, 18, 1, 1).setValue('confirmed');
      sh.getRange(foundRow, 21, 1, 1).setValue(businessType); // U列（§4-1-2）
    } else {
      // --- 新規申込 ---
      isUpdate    = false;
      appId       = _nextAppId(rows);
      ticketToken = _genToken();

      sh.appendRow([
        appId, now, now, 'web',
        salonName, staffName, email, emailNorm, phone, '', // J列(area)はフォームから削除済み（2026-08-06）
        hasTransaction, address, referrer, agree,
        '', '',                 // line_friend_id / line_user_id（LIFF連動はP2で使用）
        ticketToken, 'confirmed', '', note,
        businessType             // U列（§4-1-2）
      ]);
    }
  } finally {
    lock.releaseLock();
  }

  const passUrl = SITE_BASE_URL + 'pass.html?t=' + ticketToken;
  _sendConfirmationMail(email, salonName, staffName, passUrl, isUpdate);

  return _ok({ app_id: appId, pass_url: passUrl, is_update: isUpdate });
}

// ============================================================
// ② 入場パス取得（doGet: action=getPass）―― 完全公開・認証なし
//    token を知っている人だけが自分の情報を見られる
// ============================================================
function getPass(data) {
  _checkProps();
  const token = String(data.ticket_token || '').trim();
  if (!token) return _err('INVALID_TOKEN');

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = _getSheet(ss, SHEET_APPLICATIONS);
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][16]) === token) {
      const cfg = _getConfig();
      return _ok({
        app_id:         rows[i][0],
        salon_name:     rows[i][4],
        staff_name:     rows[i][5],
        status:         rows[i][17],
        checked_in_at:  rows[i][18] || null,
        event: {
          date:       cfg.event_date       || '',
          time:       cfg.event_time       || '',
          venue_name: cfg.venue_name       || '',
          venue_addr: cfg.venue_addr       || ''
        },
        pass_url: SITE_BASE_URL + 'pass.html?t=' + token
      });
    }
  }
  return _err('NOT_FOUND');
}

// ============================================================
// メール送信
//    送信元は beaufes@gmail.com（send-as登録済み前提）。
//    GmailAppでの送信に失敗した場合（send-as未登録等）は
//    MailApp（送信元は実行アカウントのまま・差出人名とReply-Toで代替）にフォールバックする。
// ============================================================
function _sendConfirmationMail(email, salonName, staffName, passUrl, isUpdate) {
  const cfg       = _getConfig();
  const eventDate = cfg.event_date || '2026-10-26';
  const eventTime = cfg.event_time || '10:00〜16:00';
  const venueName = cfg.venue_name || '青島屋（AOSHIMAYA）';
  const venueAddr = cfg.venue_addr || '宮崎市青島2丁目12-11';

  const subject = isUpdate
    ? '【ビューフェス2026】お申し込み内容を更新しました'
    : '【ビューフェス2026】お申し込みありがとうございます（入場パス）';

  const textBody =
    salonName + '\n' +
    staffName + ' 様\n\n' +
    'ビューフェス2026へのお申し込みを' + (isUpdate ? '更新' : '受付') + 'いたしました。\n\n' +
    '  日時 : ' + eventDate + ' ' + eventTime + '\n' +
    '  会場 : ' + venueName + ' ' + venueAddr + '\n\n' +
    '▼ 当日は入場口でこちらの入場パスをご提示ください\n' +
    '  ' + passUrl + '\n\n' +
    '※このメールを保存いただくか、上のリンクをスマホのホーム画面に追加しておくと当日スムーズです\n' +
    '\n' +
    '--\n' +
    'ビューフェス事務局（' + MAIL_FROM_ADDR + '）\n';

  const htmlBody =
    '<p>' + _escapeHtml(salonName) + '<br>' + _escapeHtml(staffName) + ' 様</p>' +
    '<p>ビューフェス2026へのお申し込みを' + (isUpdate ? '更新' : '受付') + 'いたしました。</p>' +
    '<p>日時: ' + _escapeHtml(eventDate) + ' ' + _escapeHtml(eventTime) + '<br>' +
    '会場: ' + _escapeHtml(venueName) + ' ' + _escapeHtml(venueAddr) + '</p>' +
    '<p><a href="' + passUrl + '">▼ 入場パスを開く</a></p>' +
    '<p style="color:#666;font-size:13px;">' +
    '※このメールを保存いただくか、上のリンクをスマホのホーム画面に追加しておくと当日スムーズです</p>' +
    '<p style="color:#999;font-size:12px;">ビューフェス事務局（' + MAIL_FROM_ADDR + '）</p>';

  _sendMail(email, subject, textBody, htmlBody, isUpdate ? 'resend' : 'apply');
}

function _sendMail(to, subject, textBody, htmlBody, type) {
  let status = 'ok', error = '';
  try {
    GmailApp.sendEmail(to, subject, textBody, {
      name:     MAIL_FROM_NAME,
      from:     MAIL_FROM_ADDR,
      replyTo:  MAIL_FROM_ADDR,
      htmlBody: htmlBody
    });
  } catch (e1) {
    // send-as未登録・Gmailサービス未使用等でGmailAppが使えない場合のフォールバック
    try {
      MailApp.sendEmail({
        to:       to,
        subject:  subject,
        name:     MAIL_FROM_NAME,
        replyTo:  MAIL_FROM_ADDR,
        body:     textBody,
        htmlBody: htmlBody
      });
      status = 'fallback';
      error  = 'GmailApp失敗のためMailAppにフォールバック: ' + e1;
    } catch (e2) {
      status = 'error';
      error  = String(e2);
    }
  }
  _logMail(to, type, status, error);
}

function _logMail(to, type, status, error) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = _getSheet(ss, SHEET_MAIL_LOG);
    sh.appendRow([_now(), to, type, status, error]);
  } catch (e) {
    Logger.log('mail_log書き込み失敗: ' + e);
  }
}

// ============================================================
// スプレッドシート初期セットアップ
// GASエディタから手動で一度だけ実行すること
// ============================================================
function setupSheets() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // --- applications シート（1行 = 来場者1人）---
  let appSh = ss.getSheetByName(SHEET_APPLICATIONS);
  if (!appSh) {
    appSh = ss.insertSheet(SHEET_APPLICATIONS);
    appSh.getRange(1, 1, 1, 21).setValues([[
      'app_id', 'created_at', 'updated_at', 'source',
      'salon_name', 'staff_name', 'email', 'email_norm', 'phone', 'area', // areaは2026-08-06にフォームから削除・列は維持（空文字のみ）
      'has_transaction', 'address', 'referrer', 'agree_capability',
      'line_friend_id', 'line_user_id',
      'ticket_token', 'status', 'checked_in_at', 'note',
      'business_type'          // 🆕 U列（§4-1-2・v0.5.0で追加）
    ]]);
    appSh.setFrozenRows(1);
    appSh.setColumnWidth(1, 110);
    Logger.log('applicationsシート作成完了');
  } else {
    Logger.log('applicationsシートは既に存在します');
  }

  // --- sessions シート（セミナー枠マスタ。今回は保留のため空のまま用意）---
  let sesSh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sesSh) {
    sesSh = ss.insertSheet(SHEET_SESSIONS);
    sesSh.getRange(1, 1, 1, 9).setValues([[
      'session_id', 'slot', 'title', 'speaker', 'room',
      'starts_at', 'ends_at', 'capacity', 'is_active'
    ]]);
    sesSh.setFrozenRows(1);
    Logger.log('sessionsシート作成完了');
  } else {
    Logger.log('sessionsシートは既に存在します');
  }

  // --- reservations シート（1行 = 1人 × 1セミナー）---
  let resSh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!resSh) {
    resSh = ss.insertSheet(SHEET_RESERVATIONS);
    resSh.getRange(1, 1, 1, 6).setValues([[
      'res_id', 'app_id', 'session_id', 'created_at', 'status', 'attended_at'
    ]]);
    resSh.setFrozenRows(1);
    Logger.log('reservationsシート作成完了');
  } else {
    Logger.log('reservationsシートは既に存在します');
  }

  // --- checkins シート（監査用の生ログ）---
  let ckSh = ss.getSheetByName(SHEET_CHECKINS);
  if (!ckSh) {
    ckSh = ss.insertSheet(SHEET_CHECKINS);
    ckSh.getRange(1, 1, 1, 8).setValues([[
      'checkin_id', 'app_id', 'ticket_token', 'session_id',
      'checked_at', 'staff_user_id', 'device', 'note'
    ]]);
    ckSh.setFrozenRows(1);
    Logger.log('checkinsシート作成完了');
  } else {
    Logger.log('checkinsシートは既に存在します');
  }

  // --- mail_log シート ---
  let mlSh = ss.getSheetByName(SHEET_MAIL_LOG);
  if (!mlSh) {
    mlSh = ss.insertSheet(SHEET_MAIL_LOG);
    mlSh.getRange(1, 1, 1, 5).setValues([[
      'sent_at', 'to', 'type', 'status', 'error'
    ]]);
    mlSh.setFrozenRows(1);
    Logger.log('mail_logシート作成完了');
  } else {
    Logger.log('mail_logシートは既に存在します');
  }

  // --- config シート（イベント情報。コードを触らず運用で変更できる）---
  let cfgSh = ss.getSheetByName(SHEET_CONFIG);
  if (!cfgSh) {
    cfgSh = ss.insertSheet(SHEET_CONFIG);
    cfgSh.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    cfgSh.getRange(2, 1, 6, 2).setValues([
      ['event_date',     '2026年10月26日（月）'],
      ['event_time',     '10:00〜16:00'],
      ['venue_name',     '青島屋（AOSHIMAYA）'],
      ['venue_addr',     '宮崎市青島2丁目12-11'],
      ['apply_deadline', '2026年10月20日（火）'],
      ['mail_from',      MAIL_FROM_ADDR]
    ]);
    cfgSh.setFrozenRows(1);
    Logger.log('configシート作成完了（初期値を投入済み）');
  } else {
    Logger.log('configシートは既に存在します');
  }

  Logger.log('setupSheets完了');
}

// ============================================================
// 🆕 マイグレーション: 既存の applications シートに business_type 列（U列）を追加する
// 2026-08-05・設計書v3（§4-1-2）対応。
// 【本番シートに対して一度だけ手動実行すること】GASエディタの関数選択で
// migrateAddBusinessType を選び、▷実行する。setupSheetsとは別に必要（setupSheetsは
// シートが既に存在する場合は何もしないため、既存シートへの列追加はこの関数が担う）。
// 新規にsetupSheetsでシートを作る場合はヘッダーに最初から含まれるため実行不要。
// ============================================================
function migrateAddBusinessType() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _getSheet(ss, SHEET_APPLICATIONS);
  const header = sh.getRange(1, 21).getValue();
  if (header === 'business_type') {
    Logger.log('business_type列は既にU1に存在します。何もしませんでした。');
    return;
  }
  if (header) {
    throw new Error('U1に想定外の値が入っています（"' + header + '"）。手動で確認してください。');
  }
  sh.getRange(1, 21).setValue('business_type');
  Logger.log('business_type列をU1に追加しました。');
}

// ============================================================
// 本日のメール送信残数を確認する（診断用・手動実行）
// 2026-08-04実測: tak.maejima@gmail.comはGoogle Workspace Business Standard契約にも
// かかわらず、Apps Script上では個人アカウント扱いで100通/日だった（原因未特定）
// ============================================================
function checkMailQuota() {
  Logger.log('本日の残りメール送信可能数: ' + MailApp.getRemainingDailyQuota());
}
