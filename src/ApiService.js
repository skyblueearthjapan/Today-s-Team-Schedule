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

  // Team events (from _TEAM_ pseudo-member or dedicated function)
  var teamEvents = [];
  try {
    teamEvents = getTeamEventsForDate_TE_(dateStr) || [];
  } catch (e) {
    Logger.log('getTeamEventsForDate_TE_ error: ' + e.message);
    teamEvents = events[TEAM_MEMBER_NAME] || [];
  }

  return {
    date: dateStr,
    members: members,
    teamEvents: teamEvents,
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

  // チーム予定を明示的にマージ（SyncService内でのマージに加えて確実に反映）
  try {
    var teamEventsMap = getTeamEventsForRange_TE_(startDate, endDate);
    Object.keys(teamEventsMap).forEach(function(dk) {
      if (!events[dk]) events[dk] = {};
      events[dk][TEAM_MEMBER_NAME] = teamEventsMap[dk];
    });
  } catch (e) {
    Logger.log('api_getWeekView team events merge: ' + e.message);
  }

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

  // チーム予定を明示的にマージ
  try {
    var teamEventsMap = getTeamEventsForRange_TE_(startDate, endDate);
    Object.keys(teamEventsMap).forEach(function(dk) {
      if (!events[dk]) events[dk] = {};
      events[dk][TEAM_MEMBER_NAME] = teamEventsMap[dk];
    });
  } catch (e) {
    Logger.log('api_getMonthView team events merge: ' + e.message);
  }

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

// ========================================
// チーム全体予定API
// ========================================

function api_createTeamEvent(payload) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('他のユーザーが更新中です');
  try {
    return createTeamEvent_(payload);
  } finally {
    lock.releaseLock();
  }
}

function api_updateTeamEvent(payload) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('他のユーザーが更新中です');
  try {
    return updateTeamEvent_(payload);
  } finally {
    lock.releaseLock();
  }
}

function api_deleteTeamEvent(teamEventId) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('他のユーザーが更新中です');
  try {
    return deleteTeamEvent_(teamEventId);
  } finally {
    lock.releaseLock();
  }
}

// ========================================
// 個人予定 直接書き込みAPI
// ========================================

function api_createPersonalEvent(memberName, eventData) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('他のユーザーが更新中です');
  try {
    return directCreateEvent_(memberName, eventData);
  } finally {
    lock.releaseLock();
  }
}

function api_updatePersonalEvent(memberName, eventData) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('他のユーザーが更新中です');
  try {
    return directUpdateEvent_(memberName, eventData);
  } finally {
    lock.releaseLock();
  }
}

function api_deletePersonalEvent(memberName, eventId) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('他のユーザーが更新中です');
  try {
    return directDeleteEvent_(memberName, eventId);
  } finally {
    lock.releaseLock();
  }
}

// ========================================
// モーダル用詳細取得API
// ========================================

function api_getDayDetail(dateStr, memberName) {
  var result = {
    date: dateStr,
    personalEvents: [],
    teamEvents: [],
    board: { location: '', returnTime: '', notes: '' }
  };

  // Personal events for this member
  if (memberName) {
    var allEvents = getTeamEventsForDate_(dateStr);
    result.personalEvents = allEvents[memberName] || [];

    // Board data
    var boardData = getBoardDataForDate_(dateStr);
    result.board = boardData[memberName] || { location: '', returnTime: '', notes: '' };
  }

  // Team events
  result.teamEvents = getTeamEventsForDate_TE_(dateStr);

  return result;
}

function api_getTeamEventDetail(teamEventId) {
  return getTeamEventById_(teamEventId);
}
