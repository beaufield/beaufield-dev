// =========================================
// 経費承認フロー GAS バックエンド
// --- スクリプトプロパティに設定する値 ---
// LW_CLIENT_ID       : LINE WORKS Client ID
// LW_CLIENT_SECRET   : Client Secret
// LW_SA_ID           : Service Account ID
// LW_PRIVATE_KEY     : RSA秘密鍵 PEM全文
// LW_BOT_ID          : Bot ID
// TEST_USER_ID       : テスト送信先 LW userId
// WEBHOOK_SECRET     : LINE WORKS Callback URL に付与するランダム文字列
// DB_SHEET_ID        : Google スプレッドシートのID
// AUTH_SHEET_ID      : beaufield-auth スプレッドシートID（portal と共有）
// =========================================

const VERSION = 'v1.8.1';
const APP_NAME = 'expense-approval';

// --- シート名 ---
const SHEET_REQUESTS  = '申請一覧';
const SHEET_QUEUE     = 'コメントキュー';
const SHEET_APPROVERS = '承認者マスタ';
const SHEET_SETTINGS  = '設定';

// --- 申請一覧 列番号（1始まり）---
const COL_REQ_ID       = 1;  // A: RequestId
const COL_REQ_DATE     = 2;  // B: 申請日時
const COL_REQ_USER_ID  = 3;  // C: user_id
const COL_REQ_NAME     = 4;  // D: 申請者名
const COL_REQ_LW_ID    = 5;  // E: 申請者LWユーザーID
const COL_REQ_TYPE     = 6;  // F: 経費種別
const COL_REQ_PURPOSE  = 7;  // G: 目的
const COL_REQ_USE_DATE = 8;  // H: 予定日
const COL_REQ_AMOUNT   = 9;  // I: 予定額
const COL_APR_NAME     = 10; // J: 承認者名
const COL_APR_LW_ID    = 11; // K: 承認者LWユーザーID
const COL_STATUS       = 12; // L: ステータス（申請中/承認/却下）
const COL_COMMENT      = 13; // M: 承認コメント
const COL_DONE_DATE    = 14; // N: 完了日時
const COL_ACCT_STATUS  = 15; // O: 経理ステータス（未処理/処理済）
const COL_CONFIRMED    = 16; // P: 確定金額

// --- コメントキュー 列番号 ---
const COL_Q_REQ_ID     = 1;  // A: RequestId
const COL_Q_ACTION     = 2;  // B: Action（approve/reject）
const COL_Q_FROM_USER  = 3;  // C: 承認者LWユーザーID
const COL_Q_STATUS     = 4;  // D: Status（待機中/完了）
const COL_Q_CREATED_AT = 5;  // E: CreatedAt

// --- 承認者マスタ 列番号 ---
const COL_M_USER_ID    = 1;  // A: user_id
const COL_M_NAME       = 2;  // B: 申請者名
const COL_M_LW_ID      = 3;  // C: 申請者LWユーザーID
const COL_M_APR_NAME   = 4;  // D: 承認者名
const COL_M_APR_LW_ID  = 5;  // E: 承認者LWユーザーID

// =========================================
// エントリーポイント
// =========================================

function doGet(e) {
  return jsonResponse_({ ok: false, error: 'USE_POST' });
}

function doPost(e) {
  try {
    const type = (e.parameter || {}).type;

    if (type === 'list') {
      const body = e.postData ? JSON.parse(e.postData.contents) : {};
      const auth = validateSession_(body.session_token);
      if (!auth.valid) return jsonResponse_({ ok: false, error: 'SESSION_INVALID' });
      return getRequestList_(auth);
    } else if (type === 'status') {
      const body = e.postData ? JSON.parse(e.postData.contents) : {};
      const auth = validateSession_(body.session_token);
      if (!auth.valid) return jsonResponse_({ ok: false, error: 'SESSION_INVALID' });
      return getSubmissionStatus_(body, auth);
    } else if (type === 'form') {
      const body = e.postData ? JSON.parse(e.postData.contents) : {};
      const auth = validateSession_(body.session_token);
      if (!auth.valid) return jsonResponse_({ ok: false, error: 'SESSION_INVALID' });
      return handleFormSubmit_(body, auth);
    } else if (type === 'accounting') {
      const body = e.postData ? JSON.parse(e.postData.contents) : {};
      const auth = validateSession_(body.session_token);
      if (!auth.valid) return jsonResponse_({ ok: false, error: 'SESSION_INVALID' });
      return updateAccounting_(body, auth);
    } else {
      // LINE WORKS コールバック：WEBHOOK_SECRET で検証
      const secret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET');
      if ((e.parameter || {}).secret !== secret) {
        return jsonResponse_({ ok: false, error: 'unauthorized' });
      }
      const body = e.postData ? JSON.parse(e.postData.contents) : {};
      return handleLwCallback_(body);
    }
  } catch (err) {
    Logger.log('doPost エラー: ' + err.message);
    return jsonResponse_({ ok: false, error: 'INTERNAL_ERROR' });
  }
}

// =========================================
// 申請フォーム受信
// =========================================

function handleFormSubmit_(data, auth) {
  const checked = validateExpenseInput_(data);
  if (!checked.ok) return jsonResponse_(checked);
  const value = checked.value;
  const clientId = String(data.client_request_id || '').trim();
  if (clientId && !/^[A-Za-z0-9-]{8,100}$/.test(clientId)) return jsonResponse_({ok: false, error: 'INVALID_REQUEST'});
  const requestId = 'REQ-' + (clientId || Utilities.getUuid());
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return jsonResponse_({ok: false, error: 'BUSY'}); }
  let masterInfo;
  try {
    const sheet = getDb_().getSheetByName(SHEET_REQUESTS);
    const rows = sheet.getDataRange().getValues();
    const old = rows.find(function(row, i) { return i > 0 && String(row[0]) === requestId; });
    if (old) {
      if (!sameExpenseSubmission_(old, value, auth)) return jsonResponse_({ok: false, error: 'REQUEST_CONFLICT'});
      const notice = getExpenseNotice_(requestId, 'approval');
      return jsonResponse_({ok: true, requestId: requestId, replayed: true, notification_warning: !!notice && notice.state !== 'done'});
    }
    masterInfo = lookupApprover_(auth.user_id);
    if (!masterInfo || !masterInfo.approverLwId) return jsonResponse_({ok: false, error: 'APPROVER_NOT_CONFIGURED'});
    // 通知予約を先に永続化し、追記直後に実行が止まっても既存5分トリガーで回復できるようにする。
    ensureExpenseNotice_(requestId, 'approval');
    // 内容検査・重複検査を済ませてから、一回の追記で申請を確定する。
    sheet.appendRow([requestId, new Date(), auth.user_id, expenseLiteral_(auth.name || auth.user_id),
      masterInfo.applicantLwId, value.expense_type, expenseLiteral_(value.purpose), value.use_date,
      value.amount, expenseLiteral_(masterInfo.approverName), masterInfo.approverLwId,
      '申請中', '', '', '未処理', '']);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  // 通知の失敗を申請保存の失敗にしない。二重送信を避け、再送時には通知しない。
  const notificationWarning = !deliverExpenseNotice_(requestId, 'approval');
  return jsonResponse_({ok: true, requestId: requestId, notification_warning: notificationWarning});
}

// =========================================
// LINE WORKS コールバック受信
// ボタンクリックは message タイプで届く
// =========================================

function handleLwCallback_(body) {
  const content  = body.content || {};
  const fromUser = (body.source || {}).userId;

  if (content.type === 'text') {
    return handleTextMessage_(content.text, fromUser);
  }

  return jsonResponse_({ ok: true });
}

// --- テキストメッセージ処理（ボタンクリック or コメント返信）---
function handleTextMessage_(text, fromUser) {
  // ボタンクリックによる構造化コマンドの判定
  // 形式: "承認|REQ-xxx" / "却下|REQ-xxx" / "承認+コメント|REQ-xxx" / "却下+コメント|REQ-xxx"
  const cmdMatch = (text || '').match(/^(承認\+コメント|却下\+コメント|承認|却下)\|(REQ-.+)$/);
  if (cmdMatch) {
    const action    = cmdMatch[1];
    const requestId = cmdMatch[2];

    if (action === '承認' || action === '却下') {
      return jsonResponse_(finalizeRequest_(requestId, action === '承認', '', fromUser));
    } else if (action === '承認+コメント' || action === '却下+コメント') {
      const queued = enqueueComment_(requestId, action === '承認+コメント' ? 'approve' : 'reject', fromUser);
      if (!queued.ok || queued.replayed) return jsonResponse_(queued);
      const token = getLwAccessToken_();
      if (token) {
        sendLwMessage_(token, fromUser, 'コメントを入力して返信してください。');
      }
    }
    return jsonResponse_({ ok: true });
  }

  // 自由テキスト → コメントキューの待機中エントリを探す
  const qSheet = getDb_().getSheetByName(SHEET_QUEUE);
  const rows   = qSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[COL_Q_FROM_USER - 1] === fromUser && row[COL_Q_STATUS - 1] === '待機中') {
      const requestId = row[COL_Q_REQ_ID - 1];
      const action    = row[COL_Q_ACTION - 1];

      const result = finalizeRequest_(requestId, action === 'approve', text, fromUser);
      if (result.ok) completeExpenseQueue_(requestId, fromUser);
      return jsonResponse_(result);
    }
  }

  return jsonResponse_({ ok: true });
}

// =========================================
// 申請確定処理
// =========================================

function finalizeRequest_(requestId, approved, comment, fromUser) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return {ok: false, error: 'BUSY'}; }
  let row;
  try {
    const sheet = getDb_().getSheetByName(SHEET_REQUESTS);
    const rows = sheet.getDataRange().getValues();
    const i = rows.findIndex(function(r, idx) { return idx > 0 && r[COL_REQ_ID - 1] === requestId; });
    if (i < 0) return {ok: false, error: 'NOT_FOUND'};
    row = rows[i];
    const approver = String(row[COL_APR_LW_ID - 1] || '').trim();
    if (!approver || approver !== String(fromUser || '')) return {ok: false, error: 'FORBIDDEN'};
    // 確定済みは上書きも再通知もしない。古いボタン・タイムアウト・再送を同じ規則で扱う。
    if (row[COL_STATUS - 1] !== '申請中') return {ok: true, alreadyProcessed: true};
    ensureExpenseNotice_(requestId, 'result');
    sheet.getRange(i + 1, COL_STATUS, 1, 3).setValues([
      [approved ? '承認' : '却下', expenseLiteral_(comment || ''), new Date()]
    ]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return {ok: true, notification_warning: !deliverExpenseNotice_(requestId, 'result')};
}

// =========================================
// タイムアウト監視（5分毎トリガーで実行）
// =========================================

function checkTimeouts() {
  // 再送の異常が、コメント待ち申請のタイムアウト処理を止めないようにする。
  let noticeProblem = '';
  try { if (retryExpenseNotices_()) noticeProblem = '経費通知の再送上限に達した申請があります。'; }
  catch (e) { noticeProblem = '経費通知の配送状態を確認してください。'; }
  const rows = getDb_().getSheetByName(SHEET_QUEUE).getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[COL_Q_STATUS - 1] !== '待機中') continue;
    const elapsed = (now - new Date(row[COL_Q_CREATED_AT - 1])) / 60000;
    if (!Number.isFinite(elapsed) || elapsed < 10) continue;
    const result = finalizeRequest_(row[COL_Q_REQ_ID - 1], row[COL_Q_ACTION - 1] === 'approve',
      '（コメントなし・タイムアウト）', row[COL_Q_FROM_USER - 1]);
    if (result.ok) completeExpenseQueue_(row[COL_Q_REQ_ID - 1], row[COL_Q_FROM_USER - 1]);
  }
  // 上限到達はGoogleの既存エラー通知で分かるようにする。新しい通知先は作らない。
  if (noticeProblem) throw new Error(noticeProblem);
}

// =========================================
// LINE WORKS メッセージ送信
// =========================================

// 承認依頼（ボタン4つ・message タイプ）
function sendApprovalRequest_(token, requestId, approverLwId,
                               applicantName, expenseType, purpose, useDate, amount) {
  const text = [
    '【経費事前申請】',
    '申請者: ' + applicantName,
    '種別: ' + expenseType,
    '目的: ' + purpose,
    '予定日: ' + useDate,
    '予定額: ' + Number(amount).toLocaleString() + ' 円'
  ].join('\n');

  const message = {
    content: {
      type: 'button_template',
      contentText: text,
      actions: [
        { type: 'message', label: '承認',             text: '承認|' + requestId },
        { type: 'message', label: '承認+コメント',     text: '承認+コメント|' + requestId },
        { type: 'message', label: '却下',             text: '却下|' + requestId },
        { type: 'message', label: '却下+コメント',     text: '却下+コメント|' + requestId }
      ]
    }
  };

  return sendLwMessage_(token, approverLwId, null, message);
}

// 結果通知（申請者 + 承認時は経理担当者にも送信）
function sendResultNotice_(token, applicantLwId, applicantName,
                            expenseType, purpose, useDate, amount, approved, comment) {
  const status = approved ? '✅ 承認' : '❌ 却下';
  const lines  = [
    '【経費申請 結果通知】',
    '申請者: ' + applicantName,
    '種別: ' + expenseType,
    '目的: ' + purpose,
    '予定日: ' + useDate,
    '予定額: ' + Number(amount).toLocaleString() + ' 円',
    '結果: ' + status
  ];
  if (comment) lines.push('コメント: ' + comment);
  const text = lines.join('\n');

  const applicantDelivered = sendLwMessage_(token, applicantLwId, text) !== false;
  let accountingDelivered = true;

  if (approved) {
    const accountingLwId = getSetting_('経理担当者LWユーザーID');
    if (accountingLwId) {
      accountingDelivered = sendLwMessage_(token, accountingLwId, text) !== false;
    }
  }
  return applicantDelivered && accountingDelivered;
}

// LW メッセージ送信
function sendLwMessage_(token, userId, text, messageObj) {
  const botId   = PropertiesService.getScriptProperties().getProperty('LW_BOT_ID');
  const url     = 'https://www.worksapis.com/v1.0/bots/' + botId + '/users/' + userId + '/messages';
  const payload = messageObj || { content: { type: 'text', text: text } };

  const res = UrlFetchApp.fetch(url, {
    method            : 'post',
    contentType       : 'application/json',
    headers           : { Authorization: 'Bearer ' + token },
    payload           : JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 201 && res.getResponseCode() !== 200) {
    Logger.log('LW送信失敗 HTTP=' + res.getResponseCode());
    return false;
  }
  return true;
}

// =========================================
// LINE WORKS アクセストークン取得
// トークンは1時間有効なため、CacheServiceで50分キャッシュしてJWT署名生成と
// UrlFetchApp呼び出し（承認/却下操作のたびに発生していた）を削減する
// =========================================

const CACHE_TTL_LW_TOKEN = 3000; // 50分（秒）

function getLwAccessToken_() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('lw_access_token');
  if (cached) return cached;

  const token = getLwAccessTokenFresh_();
  if (token) cache.put('lw_access_token', token, CACHE_TTL_LW_TOKEN);
  return token;
}

function getLwAccessTokenFresh_() {
  const props         = PropertiesService.getScriptProperties().getProperties();
  const clientId      = props.LW_CLIENT_ID;
  const clientSecret  = props.LW_CLIENT_SECRET;
  const saId          = props.LW_SA_ID;
  const privateKeyPem = normalizePem_(props.LW_PRIVATE_KEY);

  const now        = Math.floor(Date.now() / 1000);
  const headerB64  = base64urlEncode_(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payloadB64 = base64urlEncode_(JSON.stringify({
    iss: clientId,
    sub: saId,
    iat: now,
    exp: now + 3600
  }));
  const unsignedJwt = headerB64 + '.' + payloadB64;

  let sigBytes;
  try {
    sigBytes = Utilities.computeRsaSha256Signature(unsignedJwt, privateKeyPem);
  } catch (err) {
    Logger.log('JWT署名エラー: ' + err.message);
    return null;
  }

  const jwt = unsignedJwt + '.' + base64urlEncodeBytes_(sigBytes);

  const res = UrlFetchApp.fetch('https://auth.worksmobile.com/oauth2/v2.0/token', {
    method: 'post',
    payload: {
      grant_type   : 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion    : jwt,
      client_id    : clientId,
      client_secret: clientSecret,
      scope        : 'bot'
    },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    Logger.log('トークン取得失敗: ' + res.getContentText());
    return null;
  }

  return JSON.parse(res.getContentText()).access_token;
}

// =========================================
// ヘルパー関数
// =========================================

function getDb_() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('DB_SHEET_ID');
  return SpreadsheetApp.openById(sheetId);
}

function lookupApprover_(userId) {
  const sheet = getDb_().getSheetByName(SHEET_APPROVERS);
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL_M_USER_ID - 1] === userId) {
      return {
        applicantLwId: rows[i][COL_M_LW_ID - 1],
        approverName : rows[i][COL_M_APR_NAME - 1],
        approverLwId : rows[i][COL_M_APR_LW_ID - 1]
      };
    }
  }
  return null;
}

function enqueueComment_(requestId, action, fromUser) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return {ok: false, error: 'BUSY'}; }
  try {
    const db = getDb_();
    const request = db.getSheetByName(SHEET_REQUESTS).getDataRange().getValues()
      .find(function(row, i) { return i > 0 && row[COL_REQ_ID - 1] === requestId; });
    if (!request || request[COL_STATUS - 1] !== '申請中') return {ok: false, error: 'ALREADY_PROCESSED'};
    if (!request[COL_APR_LW_ID - 1] || String(request[COL_APR_LW_ID - 1]) !== String(fromUser || '')) return {ok: false, error: 'FORBIDDEN'};
    const sheet = db.getSheetByName(SHEET_QUEUE);
    const rows = sheet.getDataRange().getValues();
    const i = rows.findIndex(function(row, idx) { return idx > 0 && row[0] === requestId && row[2] === fromUser && row[3] === '待機中'; });
    if (i > 0) {
      if (rows[i][1] === action) return {ok: true, replayed: true};
      sheet.getRange(i + 1, 1, 1, 5).setValues([[requestId, action, fromUser, '待機中', new Date()]]);
    } else sheet.appendRow([requestId, action, fromUser, '待機中', new Date()]);
    return {ok: true};
  } finally { lock.releaseLock(); }
}

function getSetting_(key) {
  const db    = getDb_();
  const sheet = db.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    Logger.log('設定シートが見つかりません。存在シート: ' + db.getSheets().map(s => s.getName()).join(' / '));
    return null;
  }
  const rows = sheet.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) return rows[i][1];
  }
  return null;
}

function normalizePem_(rawPem) {
  const base64 = rawPem
    .replace(/\\n/g, '')
    .replace(/\r\n/g, '')
    .replace(/\r/g, '')
    .replace(/\n/g, '')
    .replace('-----BEGIN PRIVATE KEY-----', '')
    .replace('-----END PRIVATE KEY-----', '')
    .replace(/\s/g, '');

  const lines = base64.match(/.{1,64}/g) || [];
  return '-----BEGIN PRIVATE KEY-----\n' + lines.join('\n') + '\n-----END PRIVATE KEY-----';
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// beaufield-auth の sessions シートでトークンを検証する
// sessions シート構造: [token, user_id, expiresAt(ms)]
function validateSession_(token) {
  if (!token) return { valid: false };
  try {
    const authSheetId = PropertiesService.getScriptProperties().getProperty('AUTH_SHEET_ID');
    if (!authSheetId) {
      Logger.log('AUTH_SHEET_ID が未設定です');
      return { valid: false };
    }
    const authSs = SpreadsheetApp.openById(authSheetId);
    const sh   = authSs.getSheetByName('sessions');
    if (!sh) return { valid: false };
    const rows = sh.getDataRange().getValues();
    const now  = Date.now();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(token)) {
        if (Number(rows[i][2]) < now) return { valid: false };
        const userId = String(rows[i][1]);
        const usersSh = authSs.getSheetByName('users');
        if (!usersSh) return { valid: false };
        const users = usersSh.getDataRange().getValues();
        for (let j = 1; j < users.length; j++) {
          if (String(users[j][0]) !== userId) continue;
          const active = users[j][3] === true || users[j][3] === 'TRUE';
          if (!active) return { valid: false };

          const rolesSh = authSs.getSheetByName('user_app_roles');
          if (!rolesSh) return { valid: false };
          const roles = rolesSh.getDataRange().getValues();
          let role = '';
          for (let k = 1; k < roles.length; k++) {
            if (String(roles[k][0]) === userId && String(roles[k][1]) === APP_NAME) {
              role = String(roles[k][2] || '').trim().toLowerCase();
              break;
            }
          }
          if (!role || role === 'none') return { valid: false };
          return {
            valid: true,
            user_id: userId,
            name: String(users[j][1] || userId),
            is_admin: users[j][5] === true || users[j][5] === 'TRUE',
            role: role
          };
        }
        return { valid: false };
      }
    }
  } catch (e) {
    Logger.log('validateSession_ error: ' + e.message);
  }
  return { valid: false };
}

function base64urlEncode_(str) {
  return Utilities.base64EncodeWebSafe(str).replace(/=+$/, '');
}

function base64urlEncodeBytes_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function updateAccounting_(data, auth) {
  const userId          = auth.user_id;  // セッション検証済みの user_id を使用
  const requestId       = data.request_id;
  const confirmedAmount = Number(data.confirmed_amount);
  if (data.confirmed_amount === '' || data.confirmed_amount == null || !Number.isFinite(confirmedAmount) || confirmedAmount < 0) {
    return jsonResponse_({ok: false, error: 'INVALID_REQUEST'});
  }

  const accountingUserId = getSetting_('経理担当者user_id');
  const adminUserId      = getSetting_('管理者user_id');
  if (!accountingUserId && !adminUserId) {
    Logger.log('経理担当者user_id / 管理者user_id が設定シートに未登録です');
    return jsonResponse_({ ok: false, error: 'CONFIG_ERROR' });
  }
  if (userId !== accountingUserId && userId !== adminUserId) {
    return jsonResponse_({ ok: false, error: '権限がありません' });
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return jsonResponse_({ok: false, error: 'BUSY'}); }
  try {
  const sheet = getDb_().getSheetByName(SHEET_REQUESTS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL_REQ_ID - 1] !== requestId) continue;
    sheet.getRange(i + 1, COL_ACCT_STATUS, 1, 2).setValues([['処理済', confirmedAmount]]);
    return jsonResponse_({ ok: true });
  }
  return jsonResponse_({ ok: false, error: '申請が見つかりません' });
  } finally { lock.releaseLock(); }
}

function getRequestList_(auth) {
  const db    = getDb_();
  const sheet = db.getSheetByName(SHEET_REQUESTS);
  if (!sheet) {
    Logger.log('シート「' + SHEET_REQUESTS + '」が見つかりません。存在シート: ' + db.getSheets().map(s => s.getName()).join(' / '));
    return jsonResponse_({ ok: false, error: 'データシートが見つかりません。管理者に連絡してください。' });
  }
  const rows  = sheet.getDataRange().getValues();
  const list  = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[COL_REQ_ID - 1]) continue;
    list.push({
      requestId      : String(row[COL_REQ_ID - 1]),
      applyDate      : formatDateTime_(row[COL_REQ_DATE - 1]),
      name           : row[COL_REQ_NAME - 1],
      expenseType    : row[COL_REQ_TYPE - 1],
      purpose        : row[COL_REQ_PURPOSE - 1],
      useDate        : formatDate_(row[COL_REQ_USE_DATE - 1]),
      amount         : Number(row[COL_REQ_AMOUNT - 1]) || 0,
      approverName   : String(row[COL_APR_NAME - 1]).trim(),
      status         : row[COL_STATUS - 1],
      comment        : row[COL_COMMENT - 1],
      doneDate       : formatDateTime_(row[COL_DONE_DATE - 1]),
      acctStatus     : row[COL_ACCT_STATUS - 1] || '未処理',
      confirmedAmount: Number(row[COL_CONFIRMED - 1]) || 0
    });
  }

  list.reverse();
  // 経理担当者IDは内部情報のため外部に返さず、呼び出し元ユーザーが経理担当者かどうかの真偽値のみ返す
  const accountingUserId = getSetting_('経理担当者user_id') || '';
  const isAccountingUser = !!(auth && auth.user_id && auth.user_id === accountingUserId);
  return jsonResponse_({ ok: true, requests: list, isAccountingUser: isAccountingUser });
}

function formatDateTime_(d) {
  if (!d) return '';
  const date = (d instanceof Date) ? d : new Date(d);
  if (isNaN(date.getTime())) return String(d);
  const y   = date.getFullYear();
  const m   = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const h   = String(date.getHours()).padStart(2, '0');
  const min = String(date.getMinutes()).padStart(2, '0');
  return y + '/' + m + '/' + day + ' ' + h + ':' + min;
}

// デバッグ用：GASエディタから直接実行してシート名を確認する
function debugSheets() {
  const db    = getDb_();
  const names = db.getSheets().map(s => s.getName());
  Logger.log('スプレッドシートID: ' + db.getId());
  Logger.log('シート一覧: ' + names.join(' / '));
}

function formatDate_(d) {
  if (!d) return '';
  const date = (d instanceof Date) ? d : new Date(d);
  if (isNaN(date.getTime())) return String(d);
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + day;
}


// 数式として解釈される文字列を、文字として保存する。
function expenseLiteral_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
function expenseDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(value || '').replace(/\//g, '-').split('T')[0];
}
function validateExpenseInput_(data) {
  const type = String(data.expense_type || '');
  const purpose = String(data.purpose || '').trim();
  const useDate = String(data.use_date || '');
  const amount = Number(data.amount);
  const date = new Date(useDate + 'T00:00:00Z');
  if (!['交通費', '接待費', '消耗品', 'その他'].includes(type) || !purpose || purpose.length > 5000 ||
      !/^\d{4}-\d{2}-\d{2}$/.test(useDate) || !Number.isFinite(date.getTime()) ||
      date.toISOString().slice(0, 10) !== useDate || !Number.isFinite(amount) || amount <= 0) {
    return {ok: false, error: 'INVALID_REQUEST'};
  }
  return {ok: true, value: {expense_type: type, purpose: purpose, use_date: useDate, amount: amount}};
}
function sameExpenseSubmission_(row, value, auth) {
  // getValues()は文字列の先頭アポストロフィを返さない。モックの場合も同様に比較する。
  return String(row[COL_REQ_USER_ID - 1]) === String(auth.user_id) &&
    String(row[COL_REQ_TYPE - 1]) === value.expense_type &&
    String(row[COL_REQ_PURPOSE - 1]) === value.purpose &&
    expenseDate_(row[COL_REQ_USE_DATE - 1]) === value.use_date &&
    Number(row[COL_REQ_AMOUNT - 1]) === value.amount;
}
function getSubmissionStatus_(data, auth) {
  const clientId = String(data.client_request_id || '');
  if (!/^[A-Za-z0-9-]{8,100}$/.test(clientId)) return jsonResponse_({ok: false, error: 'INVALID_REQUEST'});
  const checked = validateExpenseInput_(data);
  if (!checked.ok) return jsonResponse_(checked);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return jsonResponse_({ok: false, error: 'BUSY'}); }
  try {
    const rows = getDb_().getSheetByName(SHEET_REQUESTS).getDataRange().getValues();
    const row = rows.find(function(row, i) { return i > 0 && String(row[0]) === 'REQ-' + clientId; });
    if (!row) return jsonResponse_({ok: true, found: false});
    if (!sameExpenseSubmission_(row, checked.value, auth)) return jsonResponse_({ok: false, error: 'REQUEST_CONFLICT'});
    return jsonResponse_({ok: true, found: true, requestId: row[0]});
  } finally { lock.releaseLock(); }
}


function completeExpenseQueue_(requestId, fromUser) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return false; }
  try {
    const sheet = getDb_().getSheetByName(SHEET_QUEUE), rows = sheet.getDataRange().getValues();
    rows.forEach(function(row, i) {
      if (i > 0 && row[COL_Q_REQ_ID - 1] === requestId && row[COL_Q_FROM_USER - 1] === fromUser &&
          row[COL_Q_STATUS - 1] === '待機中') sheet.getRange(i + 1, COL_Q_STATUS).setValue('完了');
    });
    return true;
  } finally { lock.releaseLock(); }
}

// 通知の本文や認証情報はPropertiesへ複製せず、申請行のIDと配送状態だけを保持する。
function expenseNoticeKey_(requestId, kind) { return 'expense_notice_v1_' + kind + '_' + requestId; }
function getExpenseNotice_(requestId, kind) {
  const raw = PropertiesService.getScriptProperties().getProperty(expenseNoticeKey_(requestId, kind));
  return raw ? JSON.parse(raw) : null;
}
function ensureExpenseNotice_(requestId, kind) {
  const existing = getExpenseNotice_(requestId, kind);
  if (existing && ['pending','sending','done'].includes(existing.state)) return;
  PropertiesService.getScriptProperties().setProperty(expenseNoticeKey_(requestId, kind), JSON.stringify({
    requestId: requestId, kind: kind, state: 'pending', attempts: 0, createdAt: Date.now(), leasedUntil: 0
  }));
}
function deliverExpenseNotice_(requestId, kind) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return false; }
  const props = PropertiesService.getScriptProperties(), key = expenseNoticeKey_(requestId, kind);
  let entry, row;
  try {
    entry = getExpenseNotice_(requestId, kind);
    if (!entry || entry.state === 'done') return true;
    if (entry.leasedUntil > Date.now() || entry.attempts >= 3) return false;
    row = getDb_().getSheetByName(SHEET_REQUESTS).getDataRange().getValues()
      .find(function(r, i) { return i > 0 && r[COL_REQ_ID - 1] === requestId; });
    if (!row || (kind === 'approval' && row[COL_STATUS - 1] !== '申請中')) {
      entry.state = 'cancelled'; props.setProperty(key, JSON.stringify(entry)); return true;
    }
    if (kind === 'result' && !['承認','却下'].includes(row[COL_STATUS - 1])) return false;
    entry.attempts++; entry.state = 'sending'; entry.leasedUntil = Date.now() + 120000;
    props.setProperty(key, JSON.stringify(entry));
  } finally { lock.releaseLock(); }
  // 通信中は申請・承認のロックを持たない。結果不明時の通知だけは重複配達の可能性がある。
  let delivered = false;
  try {
    const token = getLwAccessToken_();
    if (token && kind === 'approval') delivered = sendApprovalRequest_(token, requestId, row[COL_APR_LW_ID - 1],
      row[COL_REQ_NAME - 1], row[COL_REQ_TYPE - 1], row[COL_REQ_PURPOSE - 1],
      formatDate_(row[COL_REQ_USE_DATE - 1]), row[COL_REQ_AMOUNT - 1]) !== false;
    else if (token) delivered = sendResultNotice_(token, row[COL_REQ_LW_ID - 1], row[COL_REQ_NAME - 1],
      row[COL_REQ_TYPE - 1], row[COL_REQ_PURPOSE - 1], formatDate_(row[COL_REQ_USE_DATE - 1]),
      row[COL_REQ_AMOUNT - 1], row[COL_STATUS - 1] === '承認', row[COL_COMMENT - 1]) !== false;
  } catch (e) { Logger.log('申請状態は保存済み・通知を後で再確認'); }
  entry.state = delivered ? 'done' : entry.attempts >= 3 ? 'failed' : 'pending'; entry.leasedUntil = 0;
  try { props.setProperty(key, JSON.stringify(entry)); }
  catch (e) { Logger.log('通知結果の保存失敗。通知は重複配達の可能性あり。'); return false; }
  return delivered;
}
function retryExpenseNotices_() {
  const props = PropertiesService.getScriptProperties(), all = props.getProperties();
  const prefix = 'expense_notice_v1_'; let count = 0, failed = false;
  Object.keys(all).filter(function(k) { return k.indexOf(prefix) === 0; }).forEach(function(key) {
    const entry = JSON.parse(all[key]);
    if (['done','cancelled'].includes(entry.state)) {
      if (entry.createdAt < Date.now() - 7 * 86400000) props.deleteProperty(key);
      return;
    }
    if (entry.state === 'failed') { failed = true; return; }
    if (entry.leasedUntil > Date.now()) return;
    if (count < 10) {
      count++; deliverExpenseNotice_(entry.requestId, entry.kind);
      const latest = getExpenseNotice_(entry.requestId, entry.kind);
      if (latest && latest.state === 'failed') failed = true;
    }
  });
  // 上限到達はGoogleの既存エラー通知で分かるようにする。外部の新通知先は作らない。
  return failed;
}
