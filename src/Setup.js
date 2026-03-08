/**
 * Setup.js - 初期設定（シート自動作成・トリガー設定）
 */

function setupTeamSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. SETTINGS sheet
  createSettingsSheet_(ss);

  // 2. DB_TeamSchedule sheet
  createDbTeamScheduleSheet_(ss);

  // 3. BOARD sheet
  createBoardSheet_(ss);

  // 4. Daily trigger
  setupDailyTrigger();

  ss.toast('初期設定が完了しました。SETTINGSシートにメンバー情報を入力してください。', 'セットアップ完了', 10);
}

function createSettingsSheet_(ss) {
  var existing = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (existing) {
    Logger.log('SETTINGS sheet already exists, skipping.');
    return;
  }

  var sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);

  // Title
  sheet.getRange('A1').setValue('SETTINGS').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A1:B1').mergeAcross();

  // Headers
  sheet.getRange('A2').setValue('Key').setFontWeight('bold');
  sheet.getRange('B2').setValue('Value').setFontWeight('bold');

  // Default values
  var defaults = [
    ['Member1_Name', ''],
    ['Member1_SSID', ''],
    ['Member2_Name', ''],
    ['Member2_SSID', ''],
    ['Member3_Name', ''],
    ['Member3_SSID', ''],
    ['Member4_Name', ''],
    ['Member4_SSID', ''],
    ['Member5_Name', ''],
    ['Member5_SSID', ''],
    ['DB_EventsSheet', 'DB_Events'],
    ['SyncTime', '06:00'],
    ['LastSyncAt', ''],
    ['CompanyName', '桜井電装']
  ];

  sheet.getRange(3, 1, defaults.length, 2).setValues(defaults);

  // Column widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 400);

  // Styling
  sheet.getRange('A2:B2').setBackground('#2c4a7c').setFontColor('#ffffff');
  sheet.getRange('A1').setBackground('#f6e27a');
}

function createDbTeamScheduleSheet_(ss) {
  var existing = ss.getSheetByName(SHEET_NAMES.DB_TEAM_SCHEDULE);
  if (existing) {
    Logger.log('DB_TeamSchedule sheet already exists, skipping.');
    return;
  }

  var sheet = ss.insertSheet(SHEET_NAMES.DB_TEAM_SCHEDULE);

  // Title
  sheet.getRange('A1').setValue('DB_TeamSchedule').setFontWeight('bold').setFontSize(14);

  // Headers
  var headers = ['member_name', 'event_id', 'title', 'start_date', 'end_date',
                 'start_time', 'end_time', 'all_day', 'memo', 'color_key', 'synced_at'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, 1, headers.length).setFontWeight('bold').setBackground('#2c4a7c').setFontColor('#ffffff');

  // Freeze header rows
  sheet.setFrozenRows(2);
}

function createBoardSheet_(ss) {
  var existing = ss.getSheetByName(SHEET_NAMES.BOARD);
  if (existing) {
    Logger.log('BOARD sheet already exists, skipping.');
    return;
  }

  var sheet = ss.insertSheet(SHEET_NAMES.BOARD);

  // Title
  sheet.getRange('A1').setValue('BOARD').setFontWeight('bold').setFontSize(14);

  // Headers
  var headers = ['date', 'member_name', 'location', 'return_time', 'notes', 'updated_at'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, 1, headers.length).setFontWeight('bold').setBackground('#2c4a7c').setFontColor('#ffffff');

  // Freeze header rows
  sheet.setFrozenRows(2);
}

function setupDailyTrigger() {
  // Remove existing triggers for syncAllMembers
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'syncAllMembers') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new 6AM trigger
  ScriptApp.newTrigger('syncAllMembers')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(0)
    .create();

  Logger.log('Daily sync trigger set for 6:00 AM');
}

function removeDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'syncAllMembers') {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  Logger.log('Removed ' + count + ' trigger(s).');
}
