/**
 * ReservationService.js - 予約キュー管理
 */

/**
 * 予約を作成
 * @param {string} memberName - メンバー名
 * @param {string} action - create/update/delete
 * @param {Object} eventData - イベントデータ
 * @returns {Object} 結果
 */
function createReservation_(memberName, action, eventData) {
  if (!memberName) return { success: false, error: 'メンバー名が指定されていません' };
  if (['create', 'update', 'delete'].indexOf(action) === -1) {
    return { success: false, error: '不正なアクションです' };
  }

  // Validate event data
  if (action === 'delete') {
    if (!eventData || !eventData.event_id) {
      return { success: false, error: '削除にはevent_idが必要です' };
    }
  } else {
    var validation = Validation.validateEventPayload(eventData || {});
    if (!validation.valid) {
      return { success: false, error: validation.message };
    }
  }

  // Resolve member SSID
  var settings = getSettings_();
  var member = null;
  settings.members.forEach(function(m) {
    if (m.name === memberName) member = m;
  });

  if (!member || !member.ssid) {
    return { success: false, error: memberName + 'のカレンダーSSIDが設定されていません' };
  }

  var sheet = getSheet_(SHEET_NAMES.DB_RESERVATION_QUEUE);
  var id = newReservationId_();
  var now = getCurrentDateTime_();
  var ed = eventData || {};

  var newRow = [
    id,                              // reservation_id
    now,                             // created_at
    memberName,                      // member_name
    member.ssid,                     // member_ssid
    action,                          // action
    ed.event_id || '',               // event_id
    ed.title || '',                  // title
    ed.start_date || '',             // start_date
    ed.end_date || ed.start_date || '', // end_date
    ed.start_time || '',             // start_time
    ed.end_time || '',               // end_time
    ed.all_day ? 'TRUE' : 'FALSE',  // all_day
    ed.memo || '',                   // memo
    ed.color_key || 'other',         // color_key
    'pending',                       // status
    '',                              // applied_at
    '',                              // error_message
    ed.requested_by || ''            // requested_by
  ];

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);

  var actionLabel = action === 'create' ? '新規追加' : action === 'update' ? '更新' : '削除';
  return {
    success: true,
    reservation_id: id,
    message: memberName + 'の予定(' + actionLabel + ')を予約しました。本人のAIカレンダーで承認後に反映されます。'
  };
}

/**
 * メンバーSSIDで未適用予約を取得
 */
function getReservationsForMemberSsid_(ssid) {
  try {
    var sheet = getSheet_(SHEET_NAMES.DB_RESERVATION_QUEUE);
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];

    var data = sheet.getRange(3, 1, lastRow - 2, 18).getValues();
    var results = [];

    data.forEach(function(row) {
      if (String(row[RQ_COLS.MEMBER_SSID]).trim() === ssid && String(row[RQ_COLS.STATUS]).trim() === 'pending') {
        results.push({
          reservation_id: String(row[RQ_COLS.RESERVATION_ID]),
          created_at: String(row[RQ_COLS.CREATED_AT]),
          member_name: String(row[RQ_COLS.MEMBER_NAME]),
          action: String(row[RQ_COLS.ACTION]),
          event_id: String(row[RQ_COLS.EVENT_ID]),
          title: String(row[RQ_COLS.TITLE]),
          start_date: formatDateVal_(row[RQ_COLS.START_DATE]),
          end_date: formatDateVal_(row[RQ_COLS.END_DATE]),
          start_time: formatTimeVal_(row[RQ_COLS.START_TIME]),
          end_time: formatTimeVal_(row[RQ_COLS.END_TIME]),
          all_day: row[RQ_COLS.ALL_DAY] === 'TRUE' || row[RQ_COLS.ALL_DAY] === true,
          memo: String(row[RQ_COLS.MEMO] || ''),
          color_key: String(row[RQ_COLS.COLOR_KEY] || 'other')
        });
      }
    });

    return results;
  } catch (e) {
    Logger.log('getReservationsForMemberSsid_ error: ' + e.message);
    return [];
  }
}

/**
 * 予約ステータスを更新
 */
function updateReservationStatus_(reservationId, status, errorMsg) {
  try {
    var sheet = getSheet_(SHEET_NAMES.DB_RESERVATION_QUEUE);
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return false;

    var data = sheet.getRange(3, 1, lastRow - 2, 18).getValues();

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][RQ_COLS.RESERVATION_ID]) === reservationId) {
        var rowIdx = i + 3;
        sheet.getRange(rowIdx, RQ_COLS.STATUS + 1).setValue(status);
        if (status === 'applied') {
          sheet.getRange(rowIdx, RQ_COLS.APPLIED_AT + 1).setValue(getCurrentDateTime_());
        }
        if (errorMsg) {
          sheet.getRange(rowIdx, RQ_COLS.ERROR_MESSAGE + 1).setValue(errorMsg);
        }
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log('updateReservationStatus_ error: ' + e.message);
    return false;
  }
}
