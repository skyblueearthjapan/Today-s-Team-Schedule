/**
 * TeamEventService.js - チーム全体予定CRUD
 */

/**
 * チーム全体予定を日付で取得
 */
function getTeamEventsForDate_TE_(dateStr) {
  try {
    var sheet = getSheet_(SHEET_NAMES.DB_TEAM_EVENTS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];

    var data = sheet.getRange(3, 1, lastRow - 2, 13).getValues();
    var events = [];

    data.forEach(function(row) {
      if (String(row[TE_COLS.STATUS] || '').trim() !== 'active') return;

      var startDate = formatDateVal_(row[TE_COLS.START_DATE]);
      var endDate = formatDateVal_(row[TE_COLS.END_DATE]) || startDate;

      if (!isDateInRange_(dateStr, startDate, endDate)) return;

      events.push({
        team_event_id: String(row[TE_COLS.TEAM_EVENT_ID] || ''),
        title: String(row[TE_COLS.TITLE] || ''),
        start_date: startDate,
        end_date: endDate,
        start_time: formatTimeVal_(row[TE_COLS.START_TIME]),
        end_time: formatTimeVal_(row[TE_COLS.END_TIME]),
        all_day: row[TE_COLS.ALL_DAY] === 'TRUE' || row[TE_COLS.ALL_DAY] === true,
        memo: String(row[TE_COLS.MEMO] || ''),
        color_key: String(row[TE_COLS.COLOR_KEY] || 'other'),
        is_team_event: true
      });
    });

    // Sort: all-day first, then by start_time
    events.sort(function(a, b) {
      if (a.all_day && !b.all_day) return -1;
      if (!a.all_day && b.all_day) return 1;
      return (a.start_time || '').localeCompare(b.start_time || '');
    });

    return events;
  } catch (e) {
    Logger.log('getTeamEventsForDate_TE_ error: ' + e.message);
    return [];
  }
}

/**
 * チーム全体予定を日付範囲で取得
 */
function getTeamEventsForRange_TE_(startDate, endDate) {
  try {
    var sheet = getSheet_(SHEET_NAMES.DB_TEAM_EVENTS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return {};

    var data = sheet.getRange(3, 1, lastRow - 2, 13).getValues();
    var result = {};

    data.forEach(function(row) {
      if (String(row[TE_COLS.STATUS] || '').trim() !== 'active') return;

      var evStartDate = formatDateVal_(row[TE_COLS.START_DATE]);
      var evEndDate = formatDateVal_(row[TE_COLS.END_DATE]) || evStartDate;

      if (evStartDate > endDate || evEndDate < startDate) return;

      var event = {
        team_event_id: String(row[TE_COLS.TEAM_EVENT_ID] || ''),
        title: String(row[TE_COLS.TITLE] || ''),
        start_date: evStartDate,
        end_date: evEndDate,
        start_time: formatTimeVal_(row[TE_COLS.START_TIME]),
        end_time: formatTimeVal_(row[TE_COLS.END_TIME]),
        all_day: row[TE_COLS.ALL_DAY] === 'TRUE' || row[TE_COLS.ALL_DAY] === true,
        memo: String(row[TE_COLS.MEMO] || ''),
        color_key: String(row[TE_COLS.COLOR_KEY] || 'other'),
        is_team_event: true
      };

      var evCurrent = parseLocalDate_(evStartDate < startDate ? startDate : evStartDate);
      var evEnd = parseLocalDate_(evEndDate > endDate ? endDate : evEndDate);
      while (evCurrent <= evEnd) {
        var dk = Utilities.formatDate(evCurrent, TIMEZONE, 'yyyy-MM-dd');
        if (!result[dk]) result[dk] = [];
        result[dk].push(event);
        evCurrent.setDate(evCurrent.getDate() + 1);
      }
    });

    return result;
  } catch (e) {
    Logger.log('getTeamEventsForRange_TE_ error: ' + e.message);
    return {};
  }
}

/**
 * チーム全体予定を作成
 */
function createTeamEvent_(payload) {
  var validation = Validation.validateEventPayload(payload);
  if (!validation.valid) {
    return { success: false, error: validation.message };
  }

  var sheet = getSheet_(SHEET_NAMES.DB_TEAM_EVENTS);
  var id = newTeamEventId_();
  var now = getCurrentDateTime_();

  var newRow = [
    id,
    now,
    '',
    String(payload.title || '').trim(),
    payload.start_date,
    payload.end_date || payload.start_date,
    payload.start_time || '',
    payload.end_time || '',
    payload.all_day ? 'TRUE' : 'FALSE',
    payload.memo || '',
    payload.color_key || 'other',
    'active',
    payload.created_by || ''
  ];

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);

  return { success: true, team_event_id: id, message: 'チーム予定を作成しました' };
}

/**
 * チーム全体予定を更新
 */
function updateTeamEvent_(payload) {
  if (!payload.team_event_id) {
    return { success: false, error: 'team_event_idが指定されていません' };
  }

  var validation = Validation.validateEventPayload(payload);
  if (!validation.valid) {
    return { success: false, error: validation.message };
  }

  var sheet = getSheet_(SHEET_NAMES.DB_TEAM_EVENTS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { success: false, error: '予定が見つかりません' };

  var data = sheet.getRange(3, 1, lastRow - 2, 13).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][TE_COLS.TEAM_EVENT_ID]) === payload.team_event_id) {
      var rowIdx = i + 3;
      if (payload.title !== undefined) sheet.getRange(rowIdx, TE_COLS.TITLE + 1).setValue(payload.title);
      if (payload.start_date !== undefined) sheet.getRange(rowIdx, TE_COLS.START_DATE + 1).setValue(payload.start_date);
      if (payload.end_date !== undefined) sheet.getRange(rowIdx, TE_COLS.END_DATE + 1).setValue(payload.end_date);
      if (payload.start_time !== undefined) sheet.getRange(rowIdx, TE_COLS.START_TIME + 1).setValue(payload.start_time || '');
      if (payload.end_time !== undefined) sheet.getRange(rowIdx, TE_COLS.END_TIME + 1).setValue(payload.end_time || '');
      if (payload.all_day !== undefined) sheet.getRange(rowIdx, TE_COLS.ALL_DAY + 1).setValue(payload.all_day ? 'TRUE' : 'FALSE');
      if (payload.memo !== undefined) sheet.getRange(rowIdx, TE_COLS.MEMO + 1).setValue(payload.memo || '');
      if (payload.color_key !== undefined) sheet.getRange(rowIdx, TE_COLS.COLOR_KEY + 1).setValue(payload.color_key || 'other');
      sheet.getRange(rowIdx, TE_COLS.UPDATED_AT + 1).setValue(getCurrentDateTime_());
      return { success: true, message: 'チーム予定を更新しました' };
    }
  }

  return { success: false, error: '指定された予定が見つかりません' };
}

/**
 * チーム全体予定を削除（論理削除）
 */
function deleteTeamEvent_(teamEventId) {
  if (!teamEventId) return { success: false, error: 'IDが指定されていません' };

  var sheet = getSheet_(SHEET_NAMES.DB_TEAM_EVENTS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { success: false, error: '予定が見つかりません' };

  var data = sheet.getRange(3, 1, lastRow - 2, 13).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][TE_COLS.TEAM_EVENT_ID]) === teamEventId) {
      var rowIdx = i + 3;
      sheet.getRange(rowIdx, TE_COLS.STATUS + 1).setValue('deleted');
      sheet.getRange(rowIdx, TE_COLS.UPDATED_AT + 1).setValue(getCurrentDateTime_());
      return { success: true, message: 'チーム予定を削除しました' };
    }
  }

  return { success: false, error: '指定された予定が見つかりません' };
}

/**
 * チーム予定をIDで取得
 */
function getTeamEventById_(teamEventId) {
  try {
    var sheet = getSheet_(SHEET_NAMES.DB_TEAM_EVENTS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return null;

    var data = sheet.getRange(3, 1, lastRow - 2, 13).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (String(row[TE_COLS.TEAM_EVENT_ID]) === teamEventId && String(row[TE_COLS.STATUS]).trim() === 'active') {
        return {
          team_event_id: String(row[TE_COLS.TEAM_EVENT_ID]),
          title: String(row[TE_COLS.TITLE] || ''),
          start_date: formatDateVal_(row[TE_COLS.START_DATE]),
          end_date: formatDateVal_(row[TE_COLS.END_DATE]),
          start_time: formatTimeVal_(row[TE_COLS.START_TIME]),
          end_time: formatTimeVal_(row[TE_COLS.END_TIME]),
          all_day: row[TE_COLS.ALL_DAY] === 'TRUE' || row[TE_COLS.ALL_DAY] === true,
          memo: String(row[TE_COLS.MEMO] || ''),
          color_key: String(row[TE_COLS.COLOR_KEY] || 'other'),
          is_team_event: true
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log('getTeamEventById_ error: ' + e.message);
    return null;
  }
}
