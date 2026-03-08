/**
 * Validation.js - バリデーション
 */
var Validation = (function() {

  function isValidDateFormat(dateStr) {
    if (!dateStr || dateStr.trim() === '') return true;
    var pattern = /^\d{4}-\d{2}-\d{2}$/;
    if (!pattern.test(dateStr.trim())) return false;
    var date = new Date(dateStr);
    return !isNaN(date.getTime());
  }

  function isValidTimeFormat(timeStr) {
    if (!timeStr || timeStr.trim() === '') return true;
    var pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9])$/;
    return pattern.test(timeStr.trim());
  }

  function normalizeTime(timeStr) {
    if (!timeStr || timeStr.trim() === '') return '';
    var match = timeStr.trim().match(/^(\d{1,2}):(\d{2})$/);
    if (!match) return timeStr;
    return match[1].padStart(2, '0') + ':' + match[2];
  }

  function validateBoardPayload(payload) {
    var errors = [];

    if (!payload.memberName || String(payload.memberName).trim() === '') {
      errors.push('メンバー名がありません');
    }

    if (payload.returnTime && !isValidTimeFormat(payload.returnTime)) {
      errors.push('帰社予定の形式が不正です（例: 17:30）');
    }

    if (errors.length > 0) {
      return { valid: false, message: errors.join('\n') };
    }
    return { valid: true, message: '' };
  }

  return {
    isValidDateFormat: isValidDateFormat,
    isValidTimeFormat: isValidTimeFormat,
    normalizeTime: normalizeTime,
    validateBoardPayload: validateBoardPayload
  };
})();
