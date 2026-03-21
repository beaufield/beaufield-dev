// ============================================================
// オートシャンプー展示会 予約システム - Google Apps Script
// ============================================================
const SPREADSHEET_ID      = '1vr6KE6mQXtasSGK0gKlggNE1M379Ip7C4qgHqNuVuaU';
const LINEWORKS_WEBHOOK   = 'https://webhook.worksmobile.com/message/36bb8a6b-1912-4d60-b0b1-76f20ebf9124';
const SHEET_NAME          = '予約データ';
const APP_VERSION         = 'v2.3';

// ============================================================
// エントリーポイント
// ============================================================
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
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
// 予約済みスロット取得
// ============================================================
function getBookedSlots() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, slots: [] };
    var data = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
    var slots = data.map(function(row) {
      return { date: _toDateStr(row[0]), time: _toTimeStr(row[1]) };
    });
    return { success: true, slots: slots };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// 予約保存
// ============================================================
function saveBooking(bookingData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: 'サーバーが混み合っています。しばらくしてからお試しください。' };
  }
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var lastRow = sheet.getLastRow();
    var bookedSet = {};
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
      existing.forEach(function(row) {
        bookedSet[_toDateStr(row[0]) + '_' + _toTimeStr(row[1])] = true;
      });
    }
    for (var i = 0; i < bookingData.slots.length; i++) {
      var slot = bookingData.slots[i];
      var key = slot.date + '_' + slot.time;
      if (bookedSet[key]) {
        return { success: false, error: slot.date + ' ' + slot.time + ' の枠はすでに予約済みです。別の時間をお選びください。' };
      }
    }
    var bookingId = 'BF' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MMddHHmmss');
    var bookedAt  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    bookingData.slots.forEach(function(slot) {
      sheet.appendRow([
        bookingId, slot.date, slot.time,
        bookingData.name, bookingData.salon, bookingData.staff,
        parseInt(bookingData.count), bookingData.purpose,
        bookingData.notes || '', bookedAt
      ]);
    });
    if (LINEWORKS_WEBHOOK && LINEWORKS_WEBHOOK !== 'YOUR_LINEWORKS_WEBHOOK_URL_HERE') {
      try { _sendNotification(bookingData, bookingId); } catch (ne) {
        console.error('LINE WORKS通知エラー:', ne);
      }
    }
    return { success: true, bookingId: bookingId };
  } catch (e) {
    return { success: false, error: '保存中にエラーが発生しました: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 全予約取得
// ============================================================
function getAllBookings() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, bookings: [], version: APP_VERSION };
    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var bookings = data.map(function(row) {
      return {
        id:       String(row[0] || ''),
        date:     _toDateStr(row[1]),
        time:     _toTimeStr(row[2]),
        name:     String(row[3] || ''),
        salon:    String(row[4] || ''),
        staff:    String(row[5] || ''),
        count:    Number(row[6]) || 0,
        purpose:  String(row[7] || ''),
        notes:    String(row[8] || ''),
        bookedAt: row[9] instanceof Date
          ? Utilities.formatDate(row[9], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
          : String(row[9] || '')
      };
    });
    return { success: true, bookings: bookings, version: APP_VERSION };
  } catch (e) {
    return { success: false, error: e.toString(), version: APP_VERSION };
  }
}

// ============================================================
// 予約更新（管理画面から呼び出し）
// 更新対象: 氏名・サロン名・担当者・人数・目的・備考
// ============================================================
function updateBooking(bookingId, data) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: '予約データが見つかりません' };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var updated = 0;

    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(bookingId)) {
        // D列(4)〜I列(9): 氏名・サロン名・担当者・人数・目的・備考
        sheet.getRange(i + 2, 4, 1, 6).setValues([[
          data.name,
          data.salon,
          data.staff,
          parseInt(data.count),
          data.purpose,
          data.notes || ''
        ]]);
        updated++;
      }
    }

    if (updated === 0) return { success: false, error: '対象の予約が見つかりません' };
    return { success: true };
  } catch (e) {
    return { success: false, error: '更新中にエラーが発生しました: ' + e.toString() };
  }
}

// ============================================================
// 予約削除（管理画面から呼び出し）
// 同一予約IDのすべての行を削除する
// ============================================================
function deleteBooking(bookingId) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: '予約データが見つかりません' };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var rowsToDelete = [];

    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(bookingId)) {
        rowsToDelete.push(i + 2); // 1-indexed + ヘッダー行分
      }
    }

    if (rowsToDelete.length === 0) return { success: false, error: '対象の予約が見つかりません' };

    // 下から順に削除（行インデックスのずれを防ぐ）
    rowsToDelete.reverse().forEach(function(row) { sheet.deleteRow(row); });

    return { success: true };
  } catch (e) {
    return { success: false, error: '削除中にエラーが発生しました: ' + e.toString() };
  }
}

// ============================================================
// LINE WORKS 通知（正式フォーマット: body.text）
// ============================================================
function _sendNotification(data, bookingId) {
  var slotsText = data.slots.map(function(s) {
    return '  ・' + s.date + '  ' + s.time + '〜';
  }).join('\n');
  var text = [
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
    payload: JSON.stringify({ body: { text: text } }),
    muteHttpExceptions: true
  });
}

// ============================================================
// ユーティリティ
// ============================================================
function _toDateStr(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(val);
}
function _toTimeStr(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'HH:mm');
  var s = String(val);
  var parts = s.split(':');
  if (parts.length === 2) return parts[0].padStart(2, '0') + ':' + parts[1];
  return s;
}

// ============================================================
// 初回セットアップ（一度だけ手動実行してください）
// ============================================================
function setupSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    var headers = ['予約ID','日付','時刻','氏名','サロン名','担当者','人数','目的','備考','予約日時'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#c9daf8');
    sheet.setFrozenRows(1);
    sheet.getRange('B:B').setNumberFormat('@');
    sheet.getRange('C:C').setNumberFormat('@');
  }
  SpreadsheetApp.flush();
  Logger.log('✅ セットアップ完了');
}
