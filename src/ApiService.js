/**
 * ApiService.js - フロントエンド向けAPI
 */

/**
 * 初期データ取得
 */
function api_getInitialData() {
  var settings = getSettings_();
  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');

  return {
    today: today,
    members: settings.members.map(function(m) {
      return { name: m.name, hasCalendar: m.hasCalendar };
    }),
    companyName: settings.companyName,
    lastSyncAt: settings.lastSyncAt
  };
}

/**
 * 今日ビューのデータ取得
 */
function api_getTodayView(dateStr) {
  var settings = getSettings_();
  var events = getTeamEventsForDate_(dateStr);
  var board = getBoardDataForDate_(dateStr);

  var members = settings.members.map(function(m) {
    return {
      name: m.name,
      hasCalendar: m.hasCalendar,
      events: events[m.name] || [],
      board: board[m.name] || { location: '', returnTime: '', notes: '' }
    };
  });

  return {
    date: dateStr,
    members: members,
    lastSyncAt: settings.lastSyncAt
  };
}

/**
 * 週間ビューのデータ取得
 */
function api_getWeekView(dateStr) {
  // dateStrを含む週の月曜〜日曜を計算
  var d = parseLocalDate_(dateStr);
  var dayOfWeek = d.getDay(); // 0=Sun, 1=Mon, ...
  var mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;

  var monday = new Date(d);
  monday.setDate(d.getDate() + mondayOffset);

  var sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  var startDate = Utilities.formatDate(monday, TIMEZONE, 'yyyy-MM-dd');
  var endDate = Utilities.formatDate(sunday, TIMEZONE, 'yyyy-MM-dd');

  var events = getTeamEventsForRange_(startDate, endDate);
  var board = getBoardDataForRange_(startDate, endDate);
  var settings = getSettings_();

  var days = [];
  var current = new Date(monday);
  for (var i = 0; i < 7; i++) {
    var dk = Utilities.formatDate(current, TIMEZONE, 'yyyy-MM-dd');
    days.push({
      date: dk,
      dayOfWeek: current.getDay(),
      events: events[dk] || {},
      board: board[dk] || {}
    });
    current.setDate(current.getDate() + 1);
  }

  return {
    startDate: startDate,
    endDate: endDate,
    days: days,
    members: settings.members.map(function(m) { return { name: m.name, hasCalendar: m.hasCalendar }; }),
    lastSyncAt: settings.lastSyncAt
  };
}

/**
 * 月間ビューのデータ取得
 */
function api_getMonthView(year, month) {
  var startDate = year + '-' + String(month).padStart(2, '0') + '-01';
  var lastDay = new Date(year, month, 0).getDate();
  var endDate = year + '-' + String(month).padStart(2, '0') + '-' + String(lastDay).padStart(2, '0');

  var events = getTeamEventsForRange_(startDate, endDate);
  var settings = getSettings_();

  return {
    year: year,
    month: month,
    events: events,
    members: settings.members.map(function(m) { return { name: m.name, hasCalendar: m.hasCalendar }; }),
    lastSyncAt: settings.lastSyncAt
  };
}

/**
 * ボードデータ一括更新（A+C+BのB層）
 */
function api_applyBoardPatch(dateStr, changes) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他のユーザーがデータを更新中です。');
  }

  try {
    return applyBoardPatch_(dateStr, changes);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 手動同期
 */
function api_syncNow() {
  return syncAllMembers();
}
