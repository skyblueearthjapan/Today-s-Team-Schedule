/**
 * SyncService.js - 外部AIカレンダーからの予定集約
 */

/**
 * 全メンバーの予定を同期（6AMトリガー＋手動実行）
 */
function syncAllMembers() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('他の同期処理が実行中です。');
  }

  try {
    var settings = getSettings_();
    var members = settings.members;
    var dbEventsSheetName = settings.dbEventsSheet || 'DB_Events';

    if (members.length === 0) {
      return { success: false, error: 'メンバーが設定されていません。SETTINGSシートを確認してください。' };
    }

    var allEvents = [];
    var errors = [];
    var syncedAt = getCurrentDateTime_();

    members.forEach(function(member) {
      if (!member.ssid) {
        // カレンダーが未設定のメンバーはスキップ
        return;
      }

      try {
        var events = readMemberEvents_(member.name, member.ssid, dbEventsSheetName);
        allEvents = allEvents.concat(events.map(function(e) {
          e.synced_at = syncedAt;
          return e;
        }));
      } catch (e) {
        Logger.log('Error reading calendar for ' + member.name + ': ' + e.message);
        errors.push({ member: member.name, error: e.message });
      }
    });

    // DB_TeamScheduleを再構築
    writeTeamSchedule_(allEvents);

    // LastSyncAtを更新
    updateLastSyncAt_(syncedAt);

    return {
      success: true,
      memberCount: members.filter(function(m) { return m.ssid; }).length,
      eventCount: allEvents.length,
      errors: errors,
      timestamp: syncedAt
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * メンバーのAIカレンダーからイベントを読み取り
 */
function readMemberEvents_(memberName, ssid, dbEventsSheetName) {
  var ss = SpreadsheetApp.openById(ssid);
  var sheet = ss.getSheetByName(dbEventsSheetName);

  if (!sheet) {
    throw new Error('シート「' + dbEventsSheetName + '」が見つかりません。');
  }

  var data = sheet.getDataRange().getValues();
  var events = [];

  // Row 0 = title, Row 1 = headers, Data starts at row 2
  for (var i = 2; i < data.length; i++) {
    var row = data[i];

    // statusがactiveのみ
    var status = String(row[SRC_EVENT_COLS.STATUS] || '').trim().toLowerCase();
    if (status !== 'active') continue;

    events.push({
      member_name: memberName,
      event_id: String(row[SRC_EVENT_COLS.EVENT_ID] || ''),
      title: String(row[SRC_EVENT_COLS.TITLE] || ''),
      start_date: formatDateVal_(row[SRC_EVENT_COLS.START_DATE]),
      end_date: formatDateVal_(row[SRC_EVENT_COLS.END_DATE]),
      start_time: formatTimeVal_(row[SRC_EVENT_COLS.START_TIME]),
      end_time: formatTimeVal_(row[SRC_EVENT_COLS.END_TIME]),
      all_day: row[SRC_EVENT_COLS.ALL_DAY] === 'TRUE' || row[SRC_EVENT_COLS.ALL_DAY] === true,
      memo: String(row[SRC_EVENT_COLS.MEMO] || ''),
      color_key: String(row[SRC_EVENT_COLS.COLOR_KEY] || 'other')
    });
  }

  return events;
}

/**
 * DB_TeamScheduleシートに書き込み（全置換）
 */
function writeTeamSchedule_(events) {
  var sheet = getSheet_(SHEET_NAMES.DB_TEAM_SCHEDULE);

  // データ行をクリア（行3以降）
  var lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    sheet.getRange(3, 1, lastRow - 2, 11).clearContent();
  }

  if (events.length === 0) return;

  // 書き込みデータを構築
  var rows = events.map(function(e) {
    return [
      e.member_name,
      e.event_id,
      e.title,
      e.start_date,
      e.end_date,
      e.start_time,
      e.end_time,
      e.all_day ? 'TRUE' : 'FALSE',
      e.memo,
      e.color_key,
      e.synced_at || ''
    ];
  });

  sheet.getRange(3, 1, rows.length, 11).setValues(rows);
}

/**
 * LastSyncAtを更新
 */
function updateLastSyncAt_(timestamp) {
  var sheet = getSheet_(SHEET_NAMES.SETTINGS);
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'LastSyncAt') {
      sheet.getRange(i + 1, 2).setValue(timestamp);
      return;
    }
  }
}

/**
 * 指定日のチーム予定を取得
 */
function getTeamEventsForDate_(dateStr) {
  var sheet = getSheet_(SHEET_NAMES.DB_TEAM_SCHEDULE);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};

  var data = sheet.getRange(3, 1, lastRow - 2, 11).getValues();
  var result = {};

  data.forEach(function(row) {
    var memberName = String(row[TS_COLS.MEMBER_NAME] || '').trim();
    if (!memberName) return;

    var startDate = formatDateVal_(row[TS_COLS.START_DATE]);
    var endDate = formatDateVal_(row[TS_COLS.END_DATE]);

    if (!isDateInRange_(dateStr, startDate, endDate)) return;

    if (!result[memberName]) result[memberName] = [];

    result[memberName].push({
      event_id: String(row[TS_COLS.EVENT_ID] || ''),
      title: String(row[TS_COLS.TITLE] || ''),
      start_date: startDate,
      end_date: endDate,
      start_time: formatTimeVal_(row[TS_COLS.START_TIME]),
      end_time: formatTimeVal_(row[TS_COLS.END_TIME]),
      all_day: row[TS_COLS.ALL_DAY] === 'TRUE' || row[TS_COLS.ALL_DAY] === true,
      memo: String(row[TS_COLS.MEMO] || ''),
      color_key: String(row[TS_COLS.COLOR_KEY] || 'other')
    });
  });

  // Sort events by start_time within each member
  Object.keys(result).forEach(function(name) {
    result[name].sort(function(a, b) {
      if (a.all_day && !b.all_day) return -1;
      if (!a.all_day && b.all_day) return 1;
      return (a.start_time || '').localeCompare(b.start_time || '');
    });
  });

  // Merge team-wide events
  try {
    var teamEvents = getTeamEventsForDate_TE_(dateStr);
    if (teamEvents.length > 0) {
      result[TEAM_MEMBER_NAME] = teamEvents;
    }
  } catch (e) {
    Logger.log('Team events merge error: ' + e.message);
  }

  return result;
}

/**
 * 日付範囲のチーム予定を取得
 */
function getTeamEventsForRange_(startDate, endDate) {
  var sheet = getSheet_(SHEET_NAMES.DB_TEAM_SCHEDULE);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};

  var data = sheet.getRange(3, 1, lastRow - 2, 11).getValues();
  var result = {}; // {date: {memberName: [events]}}

  // Generate all dates in range
  var current = parseLocalDate_(startDate);
  var end = parseLocalDate_(endDate);
  while (current <= end) {
    var dateKey = Utilities.formatDate(current, TIMEZONE, 'yyyy-MM-dd');
    result[dateKey] = {};
    current.setDate(current.getDate() + 1);
  }

  data.forEach(function(row) {
    var memberName = String(row[TS_COLS.MEMBER_NAME] || '').trim();
    if (!memberName) return;

    var evStartDate = formatDateVal_(row[TS_COLS.START_DATE]);
    var evEndDate = formatDateVal_(row[TS_COLS.END_DATE]);

    // Check if event overlaps with our range
    if (evStartDate > endDate || evEndDate < startDate) return;

    var event = {
      event_id: String(row[TS_COLS.EVENT_ID] || ''),
      title: String(row[TS_COLS.TITLE] || ''),
      start_date: evStartDate,
      end_date: evEndDate,
      start_time: formatTimeVal_(row[TS_COLS.START_TIME]),
      end_time: formatTimeVal_(row[TS_COLS.END_TIME]),
      all_day: row[TS_COLS.ALL_DAY] === 'TRUE' || row[TS_COLS.ALL_DAY] === true,
      memo: String(row[TS_COLS.MEMO] || ''),
      color_key: String(row[TS_COLS.COLOR_KEY] || 'other')
    };

    // Add event to each day it spans within our range
    var evCurrent = parseLocalDate_(evStartDate < startDate ? startDate : evStartDate);
    var evEnd = parseLocalDate_(evEndDate > endDate ? endDate : evEndDate);
    while (evCurrent <= evEnd) {
      var dk = Utilities.formatDate(evCurrent, TIMEZONE, 'yyyy-MM-dd');
      if (result[dk]) {
        if (!result[dk][memberName]) result[dk][memberName] = [];
        result[dk][memberName].push(event);
      }
      evCurrent.setDate(evCurrent.getDate() + 1);
    }
  });

  // Merge team-wide events
  try {
    var teamEventsMap = getTeamEventsForRange_TE_(startDate, endDate);
    Object.keys(teamEventsMap).forEach(function(dk) {
      if (!result[dk]) result[dk] = {};
      result[dk][TEAM_MEMBER_NAME] = teamEventsMap[dk];
    });
  } catch (e) {
    Logger.log('Team events range merge error: ' + e.message);
  }

  return result;
}

// parseLocalDate_ は Config.js に移動済み
