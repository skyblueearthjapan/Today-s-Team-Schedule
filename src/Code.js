/**
 * Code.js - Webアプリ エントリポイント
 */

function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('index');

    var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
    var initialTab = (e && e.parameter && e.parameter.tab) || 'today';
    var initialDate = (e && e.parameter && e.parameter.date) || today;

    template.initialDate = initialDate;
    template.initialTab = initialTab;
    template.PORTAL_URL = PORTAL_URL;

    return template.evaluate()
      .setTitle('チームスケジュール - 桜井電装')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;">' +
      '<h1 style="color:#c00;">エラーが発生しました</h1>' +
      '<p>' + String(error.message).replace(/</g, '&lt;') + '</p>' +
      '<p>setupTeamSchedule() を実行してシートを作成してください。</p>' +
      '</body></html>'
    ).setTitle('エラー');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
