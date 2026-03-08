/**
 * BoardService.js - BOARD シート CRUD（手動入力：場所・帰社予定・備考）
 */

/**
 * 指定日の全メンバーのボードデータを取得
 */
function getBoardDataForDate_(dateStr) {
  var sheet = getSheet_(SHEET_NAMES.BOARD);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};

  var data = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
  var result = {};

  data.forEach(function(row) {
    var date = formatDateVal_(row[BOARD_COLS.DATE]);
    if (date !== dateStr) return;

    var memberName = String(row[BOARD_COLS.MEMBER_NAME] || '').trim();
    if (!memberName) return;

    result[memberName] = {
      location: String(row[BOARD_COLS.LOCATION] || ''),
      returnTime: formatTimeVal_(row[BOARD_COLS.RETURN_TIME]),
      notes: String(row[BOARD_COLS.NOTES] || '')
    };
  });

  return result;
}

/**
 * 日付範囲のボードデータを取得
 */
function getBoardDataForRange_(startDate, endDate) {
  var sheet = getSheet_(SHEET_NAMES.BOARD);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};

  var data = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
  var result = {}; // {date: {memberName: {location, returnTime, notes}}}

  data.forEach(function(row) {
    var date = formatDateVal_(row[BOARD_COLS.DATE]);
    if (date < startDate || date > endDate) return;

    var memberName = String(row[BOARD_COLS.MEMBER_NAME] || '').trim();
    if (!memberName) return;

    if (!result[date]) result[date] = {};
    result[date][memberName] = {
      location: String(row[BOARD_COLS.LOCATION] || ''),
      returnTime: formatTimeVal_(row[BOARD_COLS.RETURN_TIME]),
      notes: String(row[BOARD_COLS.NOTES] || '')
    };
  });

  return result;
}

/**
 * ボードデータの一括更新（A+C+BパターンのB層）
 */
function applyBoardPatch_(dateStr, changes) {
  var sheet = getSheet_(SHEET_NAMES.BOARD);
  var lastRow = sheet.getLastRow();
  var existingData = lastRow >= 3 ? sheet.getRange(3, 1, lastRow - 2, 6).getValues() : [];

  var results = [];

  changes.forEach(function(change) {
    try {
      // バリデーション
      var validation = Validation.validateBoardPayload(change);
      if (!validation.valid) {
        results.push({ success: false, memberName: change.memberName || '', error: validation.message });
        return;
      }

      var memberName = change.memberName;

      // 既存行を探す
      var foundRowIndex = -1;
      for (var i = 0; i < existingData.length; i++) {
        var rowDate = formatDateVal_(existingData[i][BOARD_COLS.DATE]);
        var rowMember = String(existingData[i][BOARD_COLS.MEMBER_NAME] || '').trim();
        if (rowDate === dateStr && rowMember === memberName) {
          foundRowIndex = i + 3; // Sheet row (1-based, data starts at row 3)
          break;
        }
      }

      var now = getCurrentDateTime_();

      if (foundRowIndex === -1) {
        // 新規行を追加
        var newRow = [dateStr, memberName, '', '', '', now];
        if (change.location !== undefined) newRow[BOARD_COLS.LOCATION] = change.location;
        if (change.returnTime !== undefined) newRow[BOARD_COLS.RETURN_TIME] = change.returnTime;
        if (change.notes !== undefined) newRow[BOARD_COLS.NOTES] = change.notes;

        var appendRow = sheet.getLastRow() + 1;
        sheet.getRange(appendRow, 1, 1, 6).setValues([newRow]);

        // existingDataに追加（同バッチ内の後続処理用）
        existingData.push(newRow);

        results.push({
          success: true,
          memberName: memberName,
          data: { location: newRow[BOARD_COLS.LOCATION], returnTime: newRow[BOARD_COLS.RETURN_TIME], notes: newRow[BOARD_COLS.NOTES] }
        });
      } else {
        // 既存行を更新
        if (change.location !== undefined) {
          sheet.getRange(foundRowIndex, BOARD_COLS.LOCATION + 1).setValue(change.location);
        }
        if (change.returnTime !== undefined) {
          var rtCell = sheet.getRange(foundRowIndex, BOARD_COLS.RETURN_TIME + 1);
          rtCell.setNumberFormat('@');
          rtCell.setValue(change.returnTime);
        }
        if (change.notes !== undefined) {
          sheet.getRange(foundRowIndex, BOARD_COLS.NOTES + 1).setValue(change.notes);
        }
        sheet.getRange(foundRowIndex, BOARD_COLS.UPDATED_AT + 1).setValue(now);

        // 更新後のデータを読み直し
        var updated = sheet.getRange(foundRowIndex, 1, 1, 6).getValues()[0];
        results.push({
          success: true,
          memberName: memberName,
          data: {
            location: String(updated[BOARD_COLS.LOCATION] || ''),
            returnTime: formatTimeVal_(updated[BOARD_COLS.RETURN_TIME]),
            notes: String(updated[BOARD_COLS.NOTES] || '')
          }
        });
      }
    } catch (e) {
      results.push({ success: false, memberName: change.memberName || '', error: e.message });
    }
  });

  SpreadsheetApp.flush();
  return results;
}

/**
 * メンバーのボードデータをクリア
 */
function clearBoardData_(dateStr, memberName) {
  return applyBoardPatch_(dateStr, [{
    memberName: memberName,
    location: '',
    returnTime: '',
    notes: ''
  }]);
}
