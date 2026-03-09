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

  function validateEventPayload(payload) {
    var errors = [];

    if (!payload.title || String(payload.title).trim() === '') {
      errors.push('タイトルは必須です');
    } else if (String(payload.title).length > 100) {
      errors.push('タイトルは100文字以内にしてください');
    }

    if (!payload.start_date || !isValidDateFormat(payload.start_date)) {
      errors.push('開始日が不正です（例: 2026-03-09）');
    }

    if (payload.end_date && !isValidDateFormat(payload.end_date)) {
      errors.push('終了日が不正です');
    }

    if (payload.end_date && payload.start_date && payload.end_date < payload.start_date) {
      errors.push('終了日は開始日以降にしてください');
    }

    if (payload.start_time && !isValidTimeFormat(payload.start_time)) {
      errors.push('開始時刻が不正です（例: 14:00）');
    }

    if (payload.end_time && !isValidTimeFormat(payload.end_time)) {
      errors.push('終了時刻が不正です');
    }

    if (payload.memo && String(payload.memo).length > 500) {
      errors.push('メモは500文字以内にしてください');
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
    validateBoardPayload: validateBoardPayload,
    validateEventPayload: validateEventPayload
  };
})();
