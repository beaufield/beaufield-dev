// =========================================
// 経費承認フロー GAS バックエンド
// Version: 1.0.0
//
// --- スクリプトプロパティに設定する値 ---
// LW_CLIENT_ID       : LINE WORKS Client ID
// LW_CLIENT_SECRET   : Client Secret
// LW_SA_ID           : Service Account ID
// LW_PRIVATE_KEY     : RSA秘密鍵 PEM全文
// LW_BOT_ID          : Bot ID
// WEBHOOK_SECRET     : LINE WORKS Callback URL に付与するランダム文字列
// DB_SHEET_ID        : Google スプレッドシートのID
// =========================================

const VERSION = '1.0.0';

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

function doPost(e) {
  try {
    const type = e.parameter.type;

    if (type === 'form') {
      // フォーム送信：secret不要（GAS URLは非公開、user_idはマスタで検証）
      const body = e.postData ? JSON.parse(e.postData.contents) : {};
      return handleFormSubmit_(body);
    } else {
      // LINE WORKS コールバック：WEBHOOK_SECRET で検証
      const secret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET');
      if (e.parameter.secret !== secret) {
        return jsonResponse_({ ok: false, error: 'unauthorized' });
      }
      const body = e.postData ? JSON.parse(e.postData.contents) : {};
      return handleLwCallback_(body);
    }
  } catch (err) {
    Logger.log('doPost エラー: ' + err.message);
    return jsonResponse_({ ok: false, error: err.message });
  }
}

// =========================================
// 申請フォーム受信
// =========================================

function handleFormSubmit_(data) {
  const userId        = data.user_id;
  const applicantName = data.name;
  const expenseType   = data.expense_type;
  const purpose       = data.purpose;
  const useDate       = data.use_date;
  const amount        = data.amount;

  // 承認者マスタから情報を取得
  const masterInfo = lookupApprover_(userId);
  if (!masterInfo) {
    return jsonResponse_({ ok: false, error: '承認者マスタに登録がありません: ' + userId });
  }

  const applicantLwId = masterInfo.applicantLwId;
  const approverName  = masterInfo.approverName;
  const approverLwId  = masterInfo.approverLwId;

  // RequestId 生成
  const requestId = 'REQ-' + new Date().getTime();

  // 申請一覧に登録
  getDb_().getSheetByName(SHEET_REQUESTS).appendRow([
    requestId,
    new Date(),
    userId,
    applicantName,
    applicantLwId,
    expenseType,
    purpose,
    useDate,
    amount,
    approverName,
    approverLwId,
    '申請中',
    '',
    ''
  ]);

  // LW に承認依頼を送信
  const token = getLwAccessToken_();
  if (token) {
    sendApprovalRequest_(token, requestId, approverLwId,
                         applicantName, expenseType, purpose, useDate, amount);
  }

  return jsonResponse_({ ok: true, requestId: requestId });
}

// =========================================
// LINE WORKS コールバック受信
// =========================================

function handleLwCallback_(body) {
  const content  = body.content || {};
  const fromUser = (body.source || {}).userId;

  if (content.type === 'postback') {
    return handlePostback_(content.data, fromUser);
  } else if (content.type === 'text') {
    return handleTextMessage_(content.text, fromUser);
  }

  return jsonResponse_({ ok: true });
}

// --- Postback（ボタン押下）処理 ---
function handlePostback_(data, fromUser) {
  // data 形式: "action|requestId"
  const parts     = (data || '').split('|');
  const action    = parts[0];
  const requestId = parts[1];

  if (action === 'approve' || action === 'reject') {
    // コメントなし → 即確定
    finalizeRequest_(requestId, action === 'approve', '', fromUser);
  } else if (action === 'approve_comment' || action === 'reject_comment') {
    // コメントあり → キューに積んでコメント促し
    enqueueComment_(requestId, action.replace('_comment', ''), fromUser);
    const token = getLwAccessToken_();
    if (token) {
      sendLwMessage_(token, fromUser, 'コメントを入力して返信してください。');
    }
  }

  return jsonResponse_({ ok: true });
}

// --- テキストメッセージ（コメント返信）処理 ---
function handleTextMessage_(text, fromUser) {
  const qSheet = getDb_().getSheetByName(SHEET_QUEUE);
  const rows   = qSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[COL_Q_FROM_USER - 1] === fromUser && row[COL_Q_STATUS - 1] === '待機中') {
      const requestId = row[COL_Q_REQ_ID - 1];
      const action    = row[COL_Q_ACTION - 1];

      qSheet.getRange(i + 1, COL_Q_STATUS).setValue('完了');
      finalizeRequest_(requestId, action === 'approve', text, fromUser);
      return jsonResponse_({ ok: true });
    }
  }

  return jsonResponse_({ ok: true });
}

// =========================================
// 申請確定処理
// =========================================

function finalizeRequest_(requestId, approved, comment, fromUser) {
  const reqSheet = getDb_().getSheetByName(SHEET_REQUESTS);
  const rows     = reqSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL_REQ_ID - 1] !== requestId) continue;

    const status = approved ? '承認' : '却下';
    reqSheet.getRange(i + 1, COL_STATUS).setValue(status);
    reqSheet.getRange(i + 1, COL_COMMENT).setValue(comment);
    reqSheet.getRange(i + 1, COL_DONE_DATE).setValue(new Date());

    const applicantLwId = rows[i][COL_REQ_LW_ID - 1];
    const applicantName = rows[i][COL_REQ_NAME - 1];
    const expenseType   = rows[i][COL_REQ_TYPE - 1];
    const purpose       = rows[i][COL_REQ_PURPOSE - 1];
    const useDate       = rows[i][COL_REQ_USE_DATE - 1];
    const amount        = rows[i][COL_REQ_AMOUNT - 1];

    const token = getLwAccessToken_();
    if (token) {
      sendResultNotice_(token, applicantLwId, applicantName,
                        expenseType, purpose, useDate, amount, approved, comment);
    }
    return;
  }
}

// =========================================
// タイムアウト監視（5分毎トリガーで実行）
// =========================================

function checkTimeouts() {
  const qSheet = getDb_().getSheetByName(SHEET_QUEUE);
  const rows   = qSheet.getDataRange().getValues();
  const now    = new Date();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[COL_Q_STATUS - 1] !== '待機中') continue;

    const createdAt  = new Date(row[COL_Q_CREATED_AT - 1]);
    const elapsedMin = (now - createdAt) / 1000 / 60;

    if (elapsedMin >= 10) {
      const requestId = row[COL_Q_REQ_ID - 1];
      const action    = row[COL_Q_ACTION - 1];
      const fromUser  = row[COL_Q_FROM_USER - 1];

      qSheet.getRange(i + 1, COL_Q_STATUS).setValue('完了');
      finalizeRequest_(requestId, action === 'approve', '（コメントなし・タイムアウト）', fromUser);
    }
  }
}

// =========================================
// LINE WORKS メッセージ送信
// =========================================

// 承認依頼（ボタン4つ）
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
        { type: 'postback', label: '承認',               data: 'approve|' + requestId },
        { type: 'postback', label: '承認（コメントあり）', data: 'approve_comment|' + requestId },
        { type: 'postback', label: '却下',               data: 'reject|' + requestId },
        { type: 'postback', label: '却下（コメントあり）', data: 'reject_comment|' + requestId }
      ]
    }
  };

  sendLwMessage_(token, approverLwId, null, message);
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

  // 申請者に通知
  sendLwMessage_(token, applicantLwId, text);

  // 承認時のみ経理担当者にも通知
  if (approved) {
    const accountingLwId = getSetting_('経理担当者LWユーザーID');
    if (accountingLwId) {
      sendLwMessage_(token, accountingLwId, text);
    }
  }
}

// LW メッセージ送信（テキストまたは任意メッセージオブジェクト）
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

  if (res.getResponseCode() !== 200) {
    Logger.log('LW送信失敗 userId=' + userId + ' : ' + res.getContentText());
  }
}

// =========================================
// LINE WORKS アクセストークン取得
// =========================================

function getLwAccessToken_() {
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
  const rows = getDb_().getSheetByName(SHEET_APPROVERS).getDataRange().getValues();
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
  getDb_().getSheetByName(SHEET_QUEUE).appendRow([
    requestId, action, fromUser, '待機中', new Date()
  ]);
}

function getSetting_(key) {
  const rows = getDb_().getSheetByName(SHEET_SETTINGS).getDataRange().getValues();
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

function base64urlEncode_(str) {
  return Utilities.base64EncodeWebSafe(str).replace(/=+$/, '');
}

function base64urlEncodeBytes_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}
