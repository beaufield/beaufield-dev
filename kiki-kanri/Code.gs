// ============================================================
// ビューフィールド 貸出管理アプリ — バックエンド
// VERSION: GAS 1.9.0
// 更新日: 2026-04-25
// ============================================================

const VERSION  = 'GAS 1.11.0';
const SHEET_ID      = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
const AUTH_SHEET_ID = PropertiesService.getScriptProperties().getProperty('AUTH_SHEET_ID');
// SHEET_ID / AUTH_SHEET_ID / LINEWORKS_WEBHOOK はスクリプトプロパティで管理
const APP_NAME           = 'lending';
const SESSION_HOURS      = 12;
const CACHE_TTL_SESSION  = 120; // 権限剥奪・ログアウト反映を遅らせない
const LOGIN_MAX_ATTEMPTS = 5;
const LOGIN_LOCK_MINUTES = 10;
const ss = SpreadsheetApp.openById(SHEET_ID);

// ─── セッション検証 ──────────────────────────────────────────
// beaufield-auth の sessions シートでトークンを照合する
// CacheService で 15 分間キャッシュしてシート読み込みを削減
// 戻り値: { valid: true, user_id, role } または { valid: false }
function validateSession(token) {
  if (!token) return { valid: false };

  const cache    = CacheService.getScriptCache();
  // v2: アプリ権限を含む結果だけをキャッシュする。旧キャッシュはキーを分けて無効化する。
  const cacheKey = 'sess_lending_v2_' + token.slice(-32);
  const cached   = cache.get(cacheKey);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  try {
    const authSs = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const sh     = authSs.getSheetByName('sessions');
    if (!sh) return { valid: false };

    const data = sh.getDataRange().getValues();
    const now  = Date.now();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === token) {
        if (Number(data[i][2]) < now) return { valid: false };

        const userId = String(data[i][1]);
        const users = authSs.getSheetByName('users').getDataRange().getValues();
        const userRow = users.find(function(row, idx) {
          return idx > 0 && String(row[0]) === userId;
        });
        if (!userRow || !(userRow[3] === true || userRow[3] === 'TRUE')) {
          return { valid: false };
        }

        const roles = authSs.getSheetByName('user_app_roles').getDataRange().getValues();
        let role = '';
        for (let j = 1; j < roles.length; j++) {
          if (String(roles[j][0]) === userId && String(roles[j][1]) === APP_NAME) {
            role = String(roles[j][2] || '');
            break;
          }
        }
        if (!role || role === 'none') return { valid: false };

        const isAdmin = userRow[5] === true || userRow[5] === 'TRUE';
        const r = {
          valid: true,
          user_id: userId,
          name: String(userRow[1] || userId),
          role: role,
          is_admin: isAdmin
        };
        cache.put(cacheKey, JSON.stringify(r), CACHE_TTL_SESSION);
        return r;
      }
    }
  } catch(e) {
    Logger.log('validateSession error: ' + e);
  }

  // Google側の一時障害を「無効なセッション」として負キャッシュしない。
  return { valid: false };
}

// ─── セッション保存（beaufield-auth の sessions シートに書き込む） ─
function _saveSession(authSs, token, userId, expiresAt) {
  let sh = authSs.getSheetByName('sessions');
  if (!sh) {
    sh = authSs.insertSheet('sessions');
    sh.appendRow(['token', 'user_id', 'expires_at']);
  }

  // 期限切れセッションを遅延クリーンアップ
  const data = sh.getDataRange().getValues();
  const now  = Date.now();
  for (let i = data.length - 1; i >= 1; i--) {
    if (Number(data[i][2]) < now) sh.deleteRow(i + 1);
  }

  sh.appendRow([token, userId, expiresAt]);
}

// ─── レスポンスヘルパー ─────────────────────────────────────
function _respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── エントリーポイント（POST） ───────────────────────────────
function doPost(e) {
  try {
    const action = e.parameter.action;
    const data   = JSON.parse(e.parameter.data);
    const token  = e.parameter.session_token || '';

    // login / getAuthUsers は認証不要（ログイン画面用）
    const PUBLIC_ACTIONS = ['login', 'getAuthUsers'];
    let auth = null;
    if (!PUBLIC_ACTIONS.includes(action)) {
      auth = validateSession(token);
      if (!auth.valid) {
        return _respond({ status: 'error', error: 'SESSION_INVALID' });
      }
      const adminActions = [
        'saveDevice', 'registerDevice', 'saveSalesRep', 'deleteSalesRep',
        'uploadImage', 'saveMaker', 'deleteMaker', 'issueLabel',
        'updatePrintStatus', 'assignLabel'
      ];
      const role = String(auth.role || '').toLowerCase();
      const canAdmin = auth.is_admin === true || ['admin', 'manager'].includes(role);
      if (adminActions.includes(action) && !canAdmin) {
        return _respond({ status: 'error', error: 'FORBIDDEN' });
      }

      // 操作者名はクライアント入力を信用せず、検証済みセッションから確定する。
      const actorName = auth.name || auth.user_id;
      if (action === 'saveLoan' && data) data.registeredBy = actorName;
      if (action === 'saveLoanTransaction' && data && data.loan) data.loan.registeredBy = actorName;
      if (action === 'extendDueDate' && data) data.registeredBy = actorName;
      if (action === 'changePin' && String(data.user_id || '') !== String(auth.user_id)) {
        return _respond({ status: 'error', error: 'FORBIDDEN' });
      }
    }

    let result;
    if      (action === 'saveDevice')          result = saveDevice(data);
    else if (action === 'getAllData')          result = getAllData();
    else if (action === 'saveLoan')            result = saveLoan(data);
    else if (action === 'registerDevice')      result = registerDevice(data);
    else if (action === 'saveLoanTransaction') result = saveLoanTransaction(data);
    else if (action === 'saveSalesRep')        result = saveSalesRep(data);
    else if (action === 'deleteSalesRep')      result = deleteSalesRep(data.id);
    else if (action === 'uploadImage')         result = uploadImage(data);
    else if (action === 'saveMaker')           result = saveMaker(data);
    else if (action === 'deleteMaker')         result = deleteMaker(data.id);
    else if (action === 'issueLabel')          result = issueLabel(data);
    else if (action === 'updatePrintStatus')   result = updatePrintStatus(data);
    else if (action === 'assignLabel')         result = assignLabel(data);
    else if (action === 'extendDueDate')       result = extendDueDate(data);
    else if (action === 'notify')              result = notify(data);
    else if (action === 'login')               result = login(data);
    else if (action === 'getAuthUsers')        result = getAuthUsers();
    else if (action === 'changePin')           result = changePin(data);
    else result = { error: 'Unknown action: ' + action };

    return _respond({ status: 'ok', result });
  } catch(err) {
    return _respond({ status: 'error', error: err.toString() });
  }
}

// ─── エントリーポイント（GET） ───────────────────────────────
function doGet(e) {
  return _respond({ status: 'error', error: 'USE_POST' });
}

// ─── 全データ取得（起動時に呼ばれる） ────────────────────────
function getAllData() {
  const devices  = getSheet('DeviceMaster');
  const loans    = getSheet('LoanLog');
  const reps     = getSheet('SalesRep');
  const makers   = getSheet('MakerMaster');
  const labels   = getSheet('LabelPool');   // ← LabelPool追加
  return { devices, loans, salesReps: reps, makers, labels };
}

// ─── 汎用シート取得（ヘッダー付き配列をオブジェクト配列に変換） ─
// ヘッダー名はトリムして使用（余分なスペースによるキー不一致を防止）
// Date オブジェクトは YYYY-MM-DD 文字列に変換（日本時間基準）
function getSheet(name) {
  const sheet = ss.getSheetByName(name);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const val = row[i];
      obj[h] = val instanceof Date
        ? Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd')
        : val;
    });
    return obj;
  });
}

// ─── 商品登録トランザクション（saveDevice + assignLabel を1回で処理） ─
// フロントエンドから1回のリクエストで完結させ、通信往復を削減する
function registerDevice(data) {
  const savedResult = saveDevice(data.device);
  const savedDevice = savedResult.device;
  if (data.labelId) {
    assignLabel({ labelId: data.labelId, deviceId: savedDevice.id });
    savedDevice.labelId = data.labelId; // labelIdを確実に返す
  }
  return { success: true, device: savedDevice };
}

// ─── 貸出/返却トランザクション（saveDevice + saveLoan を1回で処理） ─
// フロントエンドから1回のリクエストで完結させ、通信往復を削減する
function saveLoanTransaction(data) {
  saveDevice(data.device);
  const loanResult = saveLoan(data.loan);
  return { success: true, notifyText: loanResult.notifyText };
}

// ─── 商品マスタ保存 ──────────────────────────────────────────
function saveDevice(device) {
  const sheet = ss.getSheetByName('DeviceMaster');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  if (!device.id) {
    device.id = Date.now();
    device.createdAt = new Date().toISOString().split('T')[0];
  }
  device.updatedAt = new Date().toISOString().split('T')[0];

  const existingRow = data.findIndex((row, i) => i > 0 && row[0] == device.id);
  if (existingRow > 0) {
    const rowData = headers.map(h => device[h] !== undefined ? device[h] : '');
    sheet.getRange(existingRow + 1, 1, 1, headers.length).setValues([rowData]);
  } else {
    const rowData = headers.map(h => device[h] !== undefined ? device[h] : '');
    sheet.appendRow(rowData);
  }
  return { success: true, device };
}

// ─── 貸出ログ保存 ────────────────────────────────────────────
// LINE WORKS通知はここでは送信せず、通知文言だけ組み立てて返す。
// フロントから応答後に notify アクションで後追い送信することで、登録応答を高速化する。
function saveLoan(loan) {
  const sheet = ss.getSheetByName('LoanLog');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!loan.id) loan.id = Date.now();
  const rowData = headers.map(h => loan[h] !== undefined ? loan[h] : '');
  sheet.appendRow(rowData);

  let notifyText = null;

  // 貸出登録の通知文言
  if (loan.type === '貸出') {
    const due = loan.returnDueDate || '未設定';
    notifyText = [
      '【貸出登録】',
      '商品ID: ' + (loan.labelId || ''),
      '商品名: ' + (loan.deviceName || ''),
      '貸出先: ' + (loan.loanTo || ''),
      '返却予定日: ' + due,
      '営業担当: ' + (loan.salesRep || ''),
      '操作者: ' + (loan.registeredBy || '')
    ].join('\n');
  }

  // 返却登録の通知文言
  if (loan.type === '返却') {
    notifyText = [
      '【返却登録】',
      '商品ID: ' + (loan.labelId || ''),
      '商品名: ' + (loan.deviceName || ''),
      '返却日: ' + (loan.date || ''),
      '操作者: ' + (loan.registeredBy || '')
    ].join('\n');
  }

  return { success: true, loan, notifyText };
}

// ─── 営業担当マスタ保存 ──────────────────────────────────────
function saveSalesRep(rep) {
  const sheet = ss.getSheetByName('SalesRep');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  if (!rep.id) rep.id = Date.now();
  const existingRow = data.findIndex((row, i) => i > 0 && row[0] == rep.id);
  if (existingRow > 0) {
    const rowData = headers.map(h => rep[h] !== undefined ? rep[h] : '');
    sheet.getRange(existingRow + 1, 1, 1, headers.length).setValues([rowData]);
  } else {
    const rowData = headers.map(h => rep[h] !== undefined ? rep[h] : '');
    sheet.appendRow(rowData);
  }
  return { success: true, rep };
}

// ─── 営業担当削除 ────────────────────────────────────────────
function deleteSalesRep(id) {
  const sheet = ss.getSheetByName('SalesRep');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, i) => i > 0 && row[0] == id);
  if (rowIndex > 0) sheet.deleteRow(rowIndex + 1);
  return { success: true };
}

// ─── 画像アップロード（Google Drive） ───────────────────────
function uploadImage(data) {
  const folder = getOrCreateFolder('貸出管理_画像');
  const blob = Utilities.newBlob(
    Utilities.base64Decode(data.base64),
    data.mimeType,
    data.fileName
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();
  const imageUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
  return { success: true, imageUrl };
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

// ─── メーカーマスタ保存 ──────────────────────────────────────
function saveMaker(maker) {
  const sheet = ss.getSheetByName('MakerMaster');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  if (!maker.id) maker.id = Date.now();
  const existingRow = data.findIndex((row, i) => i > 0 && row[0] == maker.id);
  if (existingRow > 0) {
    const rowData = headers.map(h => maker[h] !== undefined ? maker[h] : '');
    sheet.getRange(existingRow + 1, 1, 1, headers.length).setValues([rowData]);
  } else {
    const rowData = headers.map(h => maker[h] !== undefined ? maker[h] : '');
    sheet.appendRow(rowData);
  }
  return { success: true, maker };
}

// ─── メーカー削除 ────────────────────────────────────────────
function deleteMaker(id) {
  const sheet = ss.getSheetByName('MakerMaster');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, i) => i > 0 && row[0] == id);
  if (rowIndex > 0) sheet.deleteRow(rowIndex + 1);
  return { success: true };
}

// ============================================================
// ─── LabelPool（ラベル管理）────────────────────────────────
// ============================================================

// ラベル発行
// data: { count: 発行枚数（数値） }
// LabelPoolシートに新しいラベルを「未印刷」ステータスで追加する
// labelIdは既存の最大番号から連番で採番する（形式: BF-00001）
function issueLabel(data) {
  const sheet = ss.getSheetByName('LabelPool');
  const count = parseInt(data.count, 10) || 50;
  const today = new Date().toISOString().split('T')[0];

  // 既存ラベルの最大番号を取得して連番の開始値を決める
  const existing = sheet.getDataRange().getValues();
  let maxNum = 0;
  if (existing.length > 1) {
    for (let i = 1; i < existing.length; i++) {
      // labelIdはB列（index 1）を想定
      const labelId = existing[i][1] || '';
      const match = String(labelId).match(/BF-(\d+)/);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
  }

  // ヘッダー確認（シートが空の場合はヘッダーを追加）
  const headers = ['id', 'labelId', 'status', 'issuedAt', 'printedAt', 'assignedAt', 'deviceId'];
  if (existing.length === 0) {
    sheet.appendRow(headers);
  }

  // 新規ラベル行を追加
  const newRows = [];
  for (let i = 0; i < count; i++) {
    const num = maxNum + i + 1;
    const labelId = 'BF-' + String(num).padStart(5, '0');
    const id = Date.now() + i; // 一意なID
    newRows.push([id, labelId, '未印刷', today, '', '', '']);
  }

  // まとめて書き込み（1行ずつより高速）
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }

  return { success: true, issued: count, startLabelId: 'BF-' + String(maxNum + 1).padStart(5, '0') };
}

// 印刷済みにする
// data: { labelIds: ['BF-00001', 'BF-00002', ...] }
// 対象ラベルのstatusを「未印刷」→「印刷済」に更新し、printedAtに本日日付を記録する
function updatePrintStatus(data) {
  const sheet = ss.getSheetByName('LabelPool');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const labelIdCol = headers.indexOf('labelId');
  const statusCol  = headers.indexOf('status');
  const printedCol = headers.indexOf('printedAt');

  const targetIds = new Set(data.labelIds || []);
  const today = new Date().toISOString().split('T')[0];
  let updatedCount = 0;

  for (let i = 1; i < rows.length; i++) {
    if (targetIds.has(rows[i][labelIdCol]) && rows[i][statusCol] === '未印刷') {
      sheet.getRange(i + 1, statusCol + 1).setValue('印刷済');
      sheet.getRange(i + 1, printedCol + 1).setValue(today);
      updatedCount++;
    }
  }

  return { success: true, updated: updatedCount };
}

// ラベル割当済みにする
// data: { labelId: 'BF-00001', deviceId: 機器ID }
// 商品登録成功後に呼ばれ、ラベルのstatusを「印刷済」→「割当済」に更新する
function assignLabel(data) {
  const sheet = ss.getSheetByName('LabelPool');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const labelIdCol  = headers.indexOf('labelId');
  const statusCol   = headers.indexOf('status');
  const assignedCol = headers.indexOf('assignedAt');
  const deviceIdCol = headers.indexOf('deviceId');

  const today = new Date().toISOString().split('T')[0];

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][labelIdCol] === data.labelId) {
      sheet.getRange(i + 1, statusCol + 1).setValue('割当済');
      sheet.getRange(i + 1, assignedCol + 1).setValue(today);
      if (data.deviceId) {
        sheet.getRange(i + 1, deviceIdCol + 1).setValue(data.deviceId);
      }
      return { success: true, labelId: data.labelId };
    }
  }

  // 対象ラベルが見つからない場合はエラーを返す
  return { success: false, error: 'ラベルが見つかりません: ' + data.labelId };
}

// ─── 返却予定日延長 ──────────────────────────────────────────
// DeviceMasterのreturnDueDateを更新し、LoanLogに延長記録を残す
// LINE WORKS通知はnotifyTextとして返し、フロントから後追い送信する
function extendDueDate(data) {
  const today = new Date().toISOString().split('T')[0];

  // DeviceMasterのreturnDueDateを直接更新
  const devSheet  = ss.getSheetByName('DeviceMaster');
  const devData   = devSheet.getDataRange().getValues();
  const devHeaders = devData[0];
  const devIdCol     = devHeaders.indexOf('id');
  const dueDateCol   = devHeaders.indexOf('returnDueDate');
  const updatedAtCol = devHeaders.indexOf('updatedAt');

  for (let i = 1; i < devData.length; i++) {
    if (String(devData[i][devIdCol]) === String(data.deviceId)) {
      devSheet.getRange(i + 1, dueDateCol + 1).setValue(data.newDueDate);
      if (updatedAtCol >= 0) devSheet.getRange(i + 1, updatedAtCol + 1).setValue(today);
      break;
    }
  }

  // LoanLogに延長記録を追加（saveLoanを使うと通知が不要に走るためシート直書き）
  const loanSheet  = ss.getSheetByName('LoanLog');
  const loanHeaders = loanSheet.getRange(1, 1, 1, loanSheet.getLastColumn()).getValues()[0];
  const loanEntry  = {
    id: Date.now(),
    type: '延長',
    labelId: data.labelId || '',
    deviceName: data.deviceName || '',
    loanTo: data.loanTo || '',
    date: today,
    returnDueDate: data.newDueDate || '',
    salesRep: data.salesRep || '',
    notes: data.oldDueDate ? '変更前: ' + data.oldDueDate : '新規設定',
    registeredBy: data.registeredBy || '',
    deviceInfo: ''
  };
  loanSheet.appendRow(loanHeaders.map(h => loanEntry[h] !== undefined ? loanEntry[h] : ''));

  // LINE WORKS通知文言（送信はフロントからの notify アクションで後追い実行する）
  const notifyText = [
    '【返却期限延長】',
    '商品ID: '    + (data.labelId   || ''),
    '商品名: '    + (data.deviceName || ''),
    '貸出先: '    + (data.loanTo     || ''),
    '変更前: '    + (data.oldDueDate || '未設定'),
    '変更後: '    + (data.newDueDate || ''),
    '操作者: '    + (data.registeredBy || '')
  ].join('\n');

  return { success: true, notifyText };
}

// ─── LINE WORKS通知の後追い送信（フロントから登録応答後に呼ばれる） ─
function notify(data) {
  const text = data && String(data.text || '');
  if (!text) return { success: false };
  const allowedPrefixes = ['【貸出登録】', '【返却登録】', '【返却期限延長】'];
  if (text.length > 2000 || !allowedPrefixes.some(function(prefix) { return text.indexOf(prefix) === 0; })) {
    throw new Error('通知内容が許可されていません');
  }
  sendLineWorksMessage(text);
  return { success: true };
}

// ─── LINE WORKS通知 ────────────────────────────────────────
function sendLineWorksMessage(text) {
  try {
    const webhookUrl = PropertiesService.getScriptProperties().getProperty('LINEWORKS_WEBHOOK');
    if (!webhookUrl) { console.error('LINEWORKS_WEBHOOKが未設定です（スクリプトプロパティを確認）'); return; }
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ body: { text: text } })
    });
  } catch(e) {
    // 通知失敗してもメイン処理は継続する
    console.error('LINE WORKS通知エラー:', e.toString());
  }
}

// ─── 週次レポート（毎週火曜9:00にトリガー登録） ─────────────
// ① 返却期限超過 ② 7日以内に期限 ③ 返却期限未設定 ④ メーカー返却期限超過 をLINE WORKSに通知する
function sendWeeklyReport() {
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const limitDate = new Date();
  limitDate.setDate(limitDate.getDate() + 7);
  const limitStr = Utilities.formatDate(limitDate, 'Asia/Tokyo', 'yyyy-MM-dd');

  // DeviceMasterを取得し、各フィールドを正規化する
  // - status: 余分なスペースを除去（手動編集で混入しやすい）
  // - returnDueDate / makerReturnDueDate: DateオブジェクトはYYYY-MM-DD変換、スラッシュ区切りはハイフンに統一
  //   （'2026/03/01' < '2026-04-21' が文字コード順でfalseになるため比較が狂う）
  const devices = getSheet('DeviceMaster').map(function(d) {
    if (d.status !== undefined) d.status = String(d.status).trim();
    if (d.returnDueDate instanceof Date) {
      d.returnDueDate = Utilities.formatDate(d.returnDueDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else if (d.returnDueDate) {
      d.returnDueDate = String(d.returnDueDate).trim().replace(/\//g, '-');
    }
    if (d.makerReturnDueDate instanceof Date) {
      d.makerReturnDueDate = Utilities.formatDate(d.makerReturnDueDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else if (d.makerReturnDueDate) {
      d.makerReturnDueDate = String(d.makerReturnDueDate).trim().replace(/\//g, '-');
    }
    return d;
  });

  // ① 期限超過（返却期限が今日より前）
  const overdue = devices.filter(function(d) {
    return d.status === '貸出中' && d.returnDueDate && d.returnDueDate < todayStr;
  }).sort(function(a, b) { return a.returnDueDate.localeCompare(b.returnDueDate); });

  // ② 7日以内に期限（今日〜7日以内）
  const soon = devices.filter(function(d) {
    return d.status === '貸出中' && d.returnDueDate && d.returnDueDate >= todayStr && d.returnDueDate <= limitStr;
  }).sort(function(a, b) { return a.returnDueDate.localeCompare(b.returnDueDate); });

  // ③ 返却期限未設定
  const noDate = devices.filter(function(d) {
    return d.status === '貸出中' && !d.returnDueDate;
  });

  // ④ メーカー返却期限超過（メーカー借入商品で、廃棄・返却済以外）
  const makerOverdue = devices.filter(function(d) {
    return d.type === 'メーカー借入'
      && d.makerReturnDueDate
      && d.makerReturnDueDate < todayStr
      && d.status !== '廃棄'
      && d.status !== 'メーカー返却済';
  }).sort(function(a, b) { return a.makerReturnDueDate.localeCompare(b.makerReturnDueDate); });

  // すべて該当なしなら簡易通知
  if (overdue.length === 0 && soon.length === 0 && noDate.length === 0 && makerOverdue.length === 0) {
    sendLineWorksMessage('【週次レポート】要確認の貸出商品はありません。');
    return;
  }

  var msg = '【週次レポート】貸出状況確認（' + todayStr + '）';

  if (overdue.length > 0) {
    msg += '\n\n■ 返却期限超過（' + overdue.length + '件）';
    overdue.forEach(function(d, i) {
      msg += '\n' + (i + 1) + '. [' + (d.labelId || d.id) + '] ' + d.name;
      msg += '\n　貸出先: ' + (d.loanTo || '未設定') + ' / 期限: ' + d.returnDueDate;
      if (d.salesRep) msg += ' / ' + d.salesRep;
    });
  }

  if (soon.length > 0) {
    msg += '\n\n■ もうすぐ期限・7日以内（' + soon.length + '件）';
    soon.forEach(function(d, i) {
      msg += '\n' + (i + 1) + '. [' + (d.labelId || d.id) + '] ' + d.name;
      msg += '\n　貸出先: ' + (d.loanTo || '未設定') + ' / 期限: ' + d.returnDueDate;
      if (d.salesRep) msg += ' / ' + d.salesRep;
    });
  }

  if (noDate.length > 0) {
    msg += '\n\n■ 返却期限未設定（' + noDate.length + '件）';
    noDate.forEach(function(d, i) {
      msg += '\n' + (i + 1) + '. [' + (d.labelId || d.id) + '] ' + d.name;
      msg += '\n　貸出先: ' + (d.loanTo || '未設定');
      if (d.salesRep) msg += ' / ' + d.salesRep;
    });
  }

  if (makerOverdue.length > 0) {
    msg += '\n\n■ メーカー返却期限超過（' + makerOverdue.length + '件）';
    makerOverdue.forEach(function(d, i) {
      msg += '\n' + (i + 1) + '. [' + (d.labelId || d.id) + '] ' + d.name;
      msg += '\n　期限: ' + d.makerReturnDueDate;
      if (d.salesRep) msg += ' / ' + d.salesRep;
    });
  }

  sendLineWorksMessage(msg);
}

// ─── beaufield-auth 認証 ────────────────────────────────────

/**
 * getAuthUsers: lendingアプリにアクセス権があるユーザー一覧を返す（ログイン画面の名前グリッド用）
 * 出力: { users: [{ user_id, name }] }
 */
function getAuthUsers() {
  var authSs    = SpreadsheetApp.openById(AUTH_SHEET_ID);
  var roleRows  = authSs.getSheetByName('user_app_roles').getDataRange().getValues();
  var authRows  = authSs.getSheetByName('users').getDataRange().getValues();

  // lendingのアクセス権があるuser_idを収集
  var accessIds = new Set();
  for (var i = 1; i < roleRows.length; i++) {
    if (roleRows[i][1] === APP_NAME && roleRows[i][2] !== 'none') {
      accessIds.add(String(roleRows[i][0]));
    }
  }

  // アクティブなユーザーのみ返す
  var users = [];
  for (var i = 1; i < authRows.length; i++) {
    var row = authRows[i];
    if (accessIds.has(String(row[0])) && (row[3] === true || row[3] === 'TRUE')) {
      users.push({ user_id: String(row[0]), name: String(row[1]) });
    }
  }
  return { users: users };
}

/**
 * login: 名前選択 + PIN で認証する（beaufield-auth を使用）
 * 入力: { user_id, pin }
 * 出力: { user_id, name, role }
 */
function login(data) {
  var userId = data.user_id;
  var pin    = data.pin;
  if (!userId || pin === undefined || pin === null || pin === '') {
    throw new Error('user_idとpinは必須です');
  }

  var props = PropertiesService.getScriptProperties();
  var lockKey = 'login_lockout_' + userId;
  var lockData = JSON.parse(props.getProperty(lockKey) || '{"count":0,"until":0}');
  var now = Date.now();
  if (Number(lockData.until || 0) > now) {
    var remaining = Math.ceil((Number(lockData.until) - now) / 60000);
    throw new Error('PIN入力がロックされています。' + remaining + '分後に再試行してください');
  }

  var authSs  = SpreadsheetApp.openById(AUTH_SHEET_ID);
  var pinStr  = String(pin).padStart(4, '0');

  // Step 1: user_id + PIN を照合
  var authRows = authSs.getSheetByName('users').getDataRange().getValues();
  var authUser = null;
  for (var i = 1; i < authRows.length; i++) {
    var row = authRows[i];
    if (String(row[0]) === userId && (row[3] === true || row[3] === 'TRUE')) {
      if (String(row[2]).padStart(4, '0') !== pinStr) {
        lockData.count = Number(lockData.count || 0) + 1;
        if (lockData.count >= LOGIN_MAX_ATTEMPTS) {
          lockData.count = 0;
          lockData.until = now + LOGIN_LOCK_MINUTES * 60 * 1000;
        }
        props.setProperty(lockKey, JSON.stringify(lockData));
        throw new Error('PINが正しくありません');
      }
      authUser = { user_id: String(row[0]), name: String(row[1]) };
      props.deleteProperty(lockKey);
      break;
    }
  }
  if (!authUser) throw new Error('ユーザーが見つかりません');

  // Step 2: アクセス権確認
  var roleRows = authSs.getSheetByName('user_app_roles').getDataRange().getValues();
  var role = null;
  for (var i = 1; i < roleRows.length; i++) {
    if (String(roleRows[i][0]) === userId && roleRows[i][1] === APP_NAME && roleRows[i][2] !== 'none') {
      role = roleRows[i][2];
      break;
    }
  }
  if (!role) throw new Error('このアプリへのアクセス権限がありません');

  // Step 3: サーバー側セッショントークン発行（beaufield-auth sessions シートに保存）
  var token     = Utilities.getUuid();
  var expiresAt = now + SESSION_HOURS * 60 * 60 * 1000;
  _saveSession(authSs, token, authUser.user_id, expiresAt);

  return { user_id: authUser.user_id, name: authUser.name, role: role, session_token: token, expires: expiresAt };
}

/**
 * changePin: 本人によるPIN変更（beaufield-auth共通PIN）
 * 入力: { user_id, current_pin, new_pin }
 * 出力: { user_id }
 */
function changePin(data) {
  var userId     = data.user_id;
  var currentPin = data.current_pin;
  var newPin     = data.new_pin;

  if (!userId || currentPin === undefined || currentPin === null || !newPin) {
    throw new Error('user_id、current_pin、new_pinは必須です');
  }
  if (!/^\d{4}$/.test(String(newPin))) throw new Error('新しいPINは4桁の数字を指定してください');

  var authSs = SpreadsheetApp.openById(AUTH_SHEET_ID);
  var sheet  = authSs.getSheetByName('users');
  var rows   = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(userId)) {
      if (String(rows[i][2]).padStart(4, '0') !== String(currentPin).padStart(4, '0')) {
        throw new Error('現在のPINが正しくありません');
      }
      sheet.getRange(i + 1, 3).setValue(String(newPin));
      return { user_id: String(userId) };
    }
  }
  throw new Error('ユーザーが見つかりません');
}

// ─── テスト用関数（GASエディタから手動実行） ────────────────
function testNotify() {
  sendLineWorksMessage('【テスト】LINE WORKS通知の動作確認です。');
}
