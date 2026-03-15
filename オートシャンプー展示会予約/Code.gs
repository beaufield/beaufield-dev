// ============================================================
// オートシャンプー展示会 予約システム - Google Apps Script
// ============================================================
// ⚠️ デプロイ前に下記3つを必ず設定してください
// ============================================================

const SPREADSHEET_ID      = 'YOUR_SPREADSHEET_ID_HERE';       // スプレッドシートのID
const LINEWORKS_WEBHOOK   = 'YOUR_LINEWORKS_WEBHOOK_URL_HERE'; // LINE WORKS Webhook URL
const ADMIN_PASSWORD      = 'beaufield2026';                   // 管理画面パスワード（変更推奨）
const SHEET_NAME          = '予約データ';

// ============================================================
// エントリーポイント
// ============================================================
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';

  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('管理画面 | オートシャンプー展示会')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ご予約フォーム | オートシャンプー展示会')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// 予約済みスロット取得（予約ページから呼び出し）
// ============================================================
function getBookedSlots() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, slots: [] };

    const data = sheet.getRange(2, 2, lastRow - 1, 2).getValues(); // B列(日付), C列(時刻)
    const slots = data.map(row => ({
      date: _toDateStr(row[0]),
      time: _toTimeStr(row[1])
    }));
    return { success: true, slots: slots };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// 予約保存（予約ページから呼び出し）
// ============================================================
function saveBooking(bookingData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: 'サーバーが混み合っています。しばらくしてからお試しください。' };
  }

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();

    // 重複チェック
    const bookedSet = new Set();
    if (lastRow > 1) {
      const existing = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
      existing.forEach(row => bookedSet.add(_toDateStr(row[0]) + '_' + _toTimeStr(row[1])));
    }

    for (const slot of bookingData.slots) {
      const key = slot.date + '_' + slot.time;
      if (bookedSet.has(key)) {
        return {
          success: false,
          error: slot.date + ' ' + slot.time + ' の枠はすでに予約済みです。別の時間をお選びください。'
        };
      }
    }

    // 保存
    const bookingId = 'BF' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MMddHHmmss');
    const bookedAt  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    bookingData.slots.forEach(slot => {
      sheet.appendRow([
        bookingId,
        slot.date,
        slot.time,
        bookingData.name,
        bookingData.salon,
        bookingData.staff,
        parseInt(bookingData.count),
        bookingData.purpose,
        bookingData.notes || '',
        bookedAt
      ]);
    });

    // LINE WORKS 通知（失敗しても予約は成功扱い）
    try { _sendNotification(bookingData, bookingId); } catch (ne) {
      console.error('LINE WORKS通知エラー:', ne);
    }

    return { success: true, bookingId: bookingId };

  } catch (e) {
    return { success: false, error: '保存中にエラーが発生しました: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 全予約取得（管理画面から呼び出し）
// ============================================================
function getAllBookings(password) {
  if (password !== ADMIN_PASSWORD) {
    return { success: false, error: 'パスワードが正しくありません' };
  }

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, bookings: [] };

    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const bookings = data.map(row => ({
      id:       row[0],
      date:     _toDateStr(row[1]),
      time:     _toTimeStr(row[2]),
      name:     row[3],
      salon:    row[4],
      staff:    row[5],
      count:    row[6],
      purpose:  row[7],
      notes:    row[8],
      bookedAt: row[9]
    }));

    return { success: true, bookings: bookings };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// LINE WORKS 通知
// ============================================================
function _sendNotification(data, bookingId) {
  const slotsText = data.slots.map(s => '  ・' + s.date + '  ' + s.time + '〜').join('\n');
  const text = [
    '📋 新規予約が入りました',
    '',
    '予約ID: ' + bookingId,
    '━━━━━━━━━━━━━',
    '氏名　: ' + data.name,
    'サロン: ' + data.salon,
    '担当者: ' + data.staff,
    '人数　: ' + data.count + '名',
    '目的　: ' + data.purpose,
    '備考　: ' + (data.notes || 'なし'),
    '',
    '【予約枠】',
    slotsText,
    '━━━━━━━━━━━━━'
  ].join('\n');

  UrlFetchApp.fetch(LINEWORKS_WEBHOOK, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify({ content: { type: 'text', text: text } }),
    muteHttpExceptions: true
  });
}

// ============================================================
// ユーティリティ（日付・時刻の型ゆれを吸収）
// ============================================================
function _toDateStr(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(val);
}

function _toTimeStr(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'HH:mm');
  return String(val);
}

// ============================================================
// 初回セットアップ（一度だけ手動実行してください）
// ============================================================
function setupSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    const headers = ['予約ID','日付','時刻','氏名','サロン名','担当者','人数','目的','備考','予約日時'];
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold').setBackground('#c9daf8');
    sheet.setFrozenRows(1);

    // B列（日付）・C列（時刻）をテキスト形式に設定（型変換防止）
    sheet.getRange('B:B').setNumberFormat('@');
    sheet.getRange('C:C').setNumberFormat('@');
  }

  SpreadsheetApp.flush();
  Logger.log('✅ セットアップ完了');
}
