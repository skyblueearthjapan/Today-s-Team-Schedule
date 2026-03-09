/**
 * DirectWriteService.js - メンバーのAIカレンダーに直接書き込み
 * 書き込み後、自身のDB_TeamScheduleも即座に同期
 */

// AIカレンダー DB_Events カラム定義（19列）
var DEST_EVENT_COLS = {
  EVENT_ID: 0, CREATED_AT: 1, UPDATED_AT: 2, SOURCE: 3, RAW_TEXT: 4,
  TITLE: 5, START_DATE: 6, END_DATE: 7, START_TIME: 8, END_TIME: 9,
  ALL_DAY: 10, MEMO: 11, STATUS: 12, LAST_AI_MODEL: 13, COLOR_KEY: 14,
  GOOGLE_EVENT_ID: 15, GCAL_SYNC_STATUS: 16, GCAL_SYNCED_AT: 17, GCAL_ERROR: 18
};

/**
 * メンバーのAIカレンダーに予定を直接作成
 */
function directCreateEvent_(memberName, eventData) {
  var member = resolveMember_(memberName);
  if (!member) return { success: false, error: memberName + 'のカレンダーSSIDが設定されていません' };

  try {
    var ss = SpreadsheetApp.openById(member.ssid);
    var sheet = ss.getSheetByName('DB_Events');
    if (!sheet) return { success: false, error: 'DB_Eventsシートが見つかりません' };

    var now = getCurrentDateTime_();
    var eventId = 'evt_' + now.replace(/[-: ]/g, '').substring(0, 15) + '_' + randomStr_(4);

    var newRow = new Array(19).fill('');
    newRow[DEST_EVENT_COLS.EVENT_ID] = eventId;
    newRow[DEST_EVENT_COLS.CREATED_AT] = now;
    newRow[DEST_EVENT_COLS.SOURCE] = 'team_schedule';
    newRow[DEST_EVENT_COLS.RAW_TEXT] = '';
    newRow[DEST_EVENT_COLS.TITLE] = String(eventData.title || '').trim();
    newRow[DEST_EVENT_COLS.START_DATE] = eventData.start_date || '';
    newRow[DEST_EVENT_COLS.END_DATE] = eventData.end_date || eventData.start_date || '';
    newRow[DEST_EVENT_COLS.START_TIME] = eventData.start_time || '';
    newRow[DEST_EVENT_COLS.END_TIME] = eventData.end_time || '';
    newRow[DEST_EVENT_COLS.ALL_DAY] = eventData.all_day ? 'TRUE' : 'FALSE';
    newRow[DEST_EVENT_COLS.MEMO] = eventData.memo || '';
    newRow[DEST_EVENT_COLS.STATUS] = 'active';
    newRow[DEST_EVENT_COLS.COLOR_KEY] = eventData.color_key || 'other';
    newRow[DEST_EVENT_COLS.GCAL_SYNC_STATUS] = 'pending';

    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 19).setValues([newRow]);

    // DB_TeamScheduleにも即反映
    syncMemberToLocal_(memberName, member.ssid);

    return {
      success: true,
      event_id: eventId,
      message: memberName + 'の予定「' + eventData.title + '」を登録しました'
    };
  } catch (e) {
    Logger.log('directCreateEvent_ error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * メンバーのAIカレンダーの予定を直接更新
 */
function directUpdateEvent_(memberName, eventData) {
  var member = resolveMember_(memberName);
  if (!member) return { success: false, error: memberName + 'のカレンダーSSIDが設定されていません' };
  if (!eventData.event_id) return { success: false, error: 'event_idが指定されていません' };

  try {
    var ss = SpreadsheetApp.openById(member.ssid);
    var sheet = ss.getSheetByName('DB_Events');
    if (!sheet) return { success: false, error: 'DB_Eventsシートが見つかりません' };

    var data = sheet.getDataRange().getValues();
    var now = getCurrentDateTime_();

    for (var i = 2; i < data.length; i++) {
      if (String(data[i][DEST_EVENT_COLS.EVENT_ID]) === eventData.event_id) {
        var rowNum = i + 1;
        if (eventData.title !== undefined) sheet.getRange(rowNum, DEST_EVENT_COLS.TITLE + 1).setValue(eventData.title);
        if (eventData.start_date !== undefined) sheet.getRange(rowNum, DEST_EVENT_COLS.START_DATE + 1).setValue(eventData.start_date);
        if (eventData.end_date !== undefined) sheet.getRange(rowNum, DEST_EVENT_COLS.END_DATE + 1).setValue(eventData.end_date);
        if (eventData.start_time !== undefined) sheet.getRange(rowNum, DEST_EVENT_COLS.START_TIME + 1).setValue(eventData.start_time);
        if (eventData.end_time !== undefined) sheet.getRange(rowNum, DEST_EVENT_COLS.END_TIME + 1).setValue(eventData.end_time);
        if (eventData.all_day !== undefined) sheet.getRange(rowNum, DEST_EVENT_COLS.ALL_DAY + 1).setValue(eventData.all_day ? 'TRUE' : 'FALSE');
        if (eventData.memo !== undefined) sheet.getRange(rowNum, DEST_EVENT_COLS.MEMO + 1).setValue(eventData.memo);
        if (eventData.color_key !== undefined) sheet.getRange(rowNum, DEST_EVENT_COLS.COLOR_KEY + 1).setValue(eventData.color_key);
        sheet.getRange(rowNum, DEST_EVENT_COLS.UPDATED_AT + 1).setValue(now);
        sheet.getRange(rowNum, DEST_EVENT_COLS.GCAL_SYNC_STATUS + 1).setValue('pending');

        // DB_TeamScheduleにも即反映
        syncMemberToLocal_(memberName, member.ssid);

        return { success: true, message: memberName + 'の予定を更新しました' };
      }
    }

    return { success: false, error: '予定が見つかりませんでした' };
  } catch (e) {
    Logger.log('directUpdateEvent_ error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * メンバーのAIカレンダーの予定を直接削除（論理削除）
 */
function directDeleteEvent_(memberName, eventId) {
  var member = resolveMember_(memberName);
  if (!member) return { success: false, error: memberName + 'のカレンダーSSIDが設定されていません' };
  if (!eventId) return { success: false, error: 'event_idが指定されていません' };

  try {
    var ss = SpreadsheetApp.openById(member.ssid);
    var sheet = ss.getSheetByName('DB_Events');
    if (!sheet) return { success: false, error: 'DB_Eventsシートが見つかりません' };

    var data = sheet.getDataRange().getValues();
    var now = getCurrentDateTime_();

    for (var i = 2; i < data.length; i++) {
      if (String(data[i][DEST_EVENT_COLS.EVENT_ID]) === eventId) {
        var rowNum = i + 1;
        sheet.getRange(rowNum, DEST_EVENT_COLS.STATUS + 1).setValue('deleted');
        sheet.getRange(rowNum, DEST_EVENT_COLS.UPDATED_AT + 1).setValue(now);

        // DB_TeamScheduleにも即反映
        syncMemberToLocal_(memberName, member.ssid);

        return { success: true, message: '予定を削除しました' };
      }
    }

    return { success: false, error: '予定が見つかりませんでした' };
  } catch (e) {
    Logger.log('directDeleteEvent_ error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 指定メンバーのAIカレンダーからDB_TeamScheduleに即同期
 * （該当メンバー分のみ差し替え）
 */
function syncMemberToLocal_(memberName, ssid) {
  try {
    var settings = getSettings_();
    var dbEventsSheetName = settings.dbEventsSheet || 'DB_Events';

    // メンバーの最新イベントを読み取り
    var events = readMemberEvents_(memberName, ssid, dbEventsSheetName);
    var now = getCurrentDateTime_();
    events.forEach(function(e) { e.synced_at = now; });

    // DB_TeamScheduleから該当メンバーの行を削除して書き直す
    var tsSheet = getSheet_(SHEET_NAMES.DB_TEAM_SCHEDULE);
    var tsLastRow = tsSheet.getLastRow();

    if (tsLastRow >= 3) {
      var existingData = tsSheet.getRange(3, 1, tsLastRow - 2, 11).getValues();
      // 下から削除（行番号ずれ防止）
      for (var i = existingData.length - 1; i >= 0; i--) {
        if (String(existingData[i][0]).trim() === memberName) {
          tsSheet.deleteRow(i + 3);
        }
      }
    }

    // 新しいデータを追加
    if (events.length > 0) {
      var rows = events.map(function(e) {
        return [
          e.member_name, e.event_id, e.title,
          e.start_date, e.end_date, e.start_time, e.end_time,
          e.all_day ? 'TRUE' : 'FALSE', e.memo, e.color_key, e.synced_at
        ];
      });
      var newLastRow = tsSheet.getLastRow();
      tsSheet.getRange(newLastRow + 1, 1, rows.length, 11).setValues(rows);
    }
  } catch (e) {
    Logger.log('syncMemberToLocal_ error: ' + e.message);
  }
}

// ヘルパー
function resolveMember_(memberName) {
  var settings = getSettings_();
  var found = null;
  settings.members.forEach(function(m) {
    if (m.name === memberName && m.ssid) found = m;
  });
  return found;
}

function randomStr_(len) {
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var result = '';
  for (var i = 0; i < len; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}
