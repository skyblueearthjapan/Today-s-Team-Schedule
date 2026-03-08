/**
 * Config.js - 定数・ヘルパー
 */

var SHEET_NAMES = {
  SETTINGS: 'SETTINGS',
  DB_TEAM_SCHEDULE: 'DB_TeamSchedule',
  BOARD: 'BOARD'
};

// DB_TeamSchedule column indices (0-indexed)
var TS_COLS = {
  MEMBER_NAME: 0, EVENT_ID: 1, TITLE: 2, START_DATE: 3, END_DATE: 4,
  START_TIME: 5, END_TIME: 6, ALL_DAY: 7, MEMO: 8, COLOR_KEY: 9, SYNCED_AT: 10
};

// BOARD column indices (0-indexed)
var BOARD_COLS = {
  DATE: 0, MEMBER_NAME: 1, LOCATION: 2, RETURN_TIME: 3, NOTES: 4, UPDATED_AT: 5
};

// Source AI Calendar DB_Events column indices (0-indexed)
var SRC_EVENT_COLS = {
  EVENT_ID: 0, TITLE: 5, START_DATE: 6, END_DATE: 7, START_TIME: 8,
  END_TIME: 9, ALL_DAY: 10, MEMO: 11, STATUS: 12, COLOR_KEY: 14
};

var TIMEZONE = 'Asia/Tokyo';

// SKD Portal URL
var PORTAL_URL = 'https://script.google.com/macros/s/AKfycbwSeRFVWYgi-AzVfc0hzKiKHRl0pZBFS26s7BQl6m5R_nffWA72x-ZjpGOZb-VzrBs/exec';

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('シート「' + name + '」が見つかりません。setupTeamSchedule()を実行してください。');
  return sheet;
}

function getSettings_() {
  var sheet = getSheet_(SHEET_NAMES.SETTINGS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { members: [], syncTime: '06:00', lastSyncAt: '', companyName: '桜井電装' };

  var data = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
  var map = {};
  data.forEach(function(row) {
    var key = String(row[0] || '').trim();
    var val = String(row[1] || '').trim();
    if (key) map[key] = val;
  });

  return {
    members: getMembers_(map),
    syncTime: map['SyncTime'] || '06:00',
    lastSyncAt: map['LastSyncAt'] || '',
    companyName: map['CompanyName'] || '桜井電装',
    dbEventsSheet: map['DB_EventsSheet'] || 'DB_Events'
  };
}

function getMembers_(map) {
  var members = [];
  for (var i = 1; i <= 5; i++) {
    var name = map['Member' + i + '_Name'] || '';
    var ssid = map['Member' + i + '_SSID'] || '';
    if (name) {
      members.push({ name: name, ssid: ssid, hasCalendar: !!ssid });
    }
  }
  return members;
}

function formatDateVal_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, TIMEZONE, 'yyyy-MM-dd');
  }
  return String(value || '').trim();
}

function formatTimeVal_(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'GMT', 'HH:mm');
  }
  var s = String(value).trim();
  if (!s) return '';
  var m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return String(m[1]).padStart(2, '0') + ':' + m[2];
  return s;
}

function isDateInRange_(target, start, end) {
  return target >= start && target <= end;
}

function getCurrentDateTime_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}
