/**
 * Code.gs
 * メインエントリーポイント・APIルーティング
 */

/**
 * Webアプリのエントリーポイント
 * @param {Object} e - リクエストパラメータ
 * @returns {HtmlOutput}
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('生産工程表')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTMLファイルをインクルード
 * @param {string} filename - ファイル名
 * @returns {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * APIエンドポイント（POST）
 * @param {Object} e - リクエスト
 * @returns {TextOutput}
 */
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const { action, params } = request;

    let result;
    switch (action) {
      // ========== 取得系 ==========
      case 'getBootstrapData':
        result = getBootstrapData(params.rangeStart, params.days);
        break;

      case 'getSchedules':
        result = getSchedules(params.rangeStart, params.rangeEnd, params.filters);
        break;

      case 'getJobs':
        result = getAllJobs();
        break;

      case 'getProcesses':
        result = getAllProcesses(params.includeInactive);
        break;

      case 'getPeople':
        result = getAllPeople(params.includeInactive);
        break;

      case 'getAttachments':
        result = getAttachments(params.jobId);
        break;

      case 'getTrips':
        result = getTrips(params.rangeStart, params.rangeEnd, params.filters);
        break;

      // ========== Schedule CRUD ==========
      case 'createSchedule':
        result = createSchedule(params.payload);
        break;

      case 'updateSchedule':
        result = updateSchedule(params.scheduleId, params.patch, params.expectedUpdatedAt);
        break;

      case 'deleteSchedule':
        result = deleteSchedule(params.scheduleId, params.expectedUpdatedAt);
        break;

      default:
        throw new Error(`Unknown action: ${action}`);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    const statusCode = error.code || 500;
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.message,
        code: statusCode
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * クライアントサイドから呼び出し可能なAPI（google.script.run用）
 */

function api_getBootstrapData(rangeStart, days) {
  return getBootstrapData(rangeStart, days);
}

function api_getSchedules(rangeStart, rangeEnd, filters) {
  return getSchedules(rangeStart, rangeEnd, filters);
}

function api_createSchedule(payload) {
  return createSchedule(payload);
}

function api_updateSchedule(scheduleId, patch, expectedUpdatedAt) {
  return updateSchedule(scheduleId, patch, expectedUpdatedAt);
}

function api_deleteSchedule(scheduleId, expectedUpdatedAt) {
  return deleteSchedule(scheduleId, expectedUpdatedAt);
}

function api_getAttachments(jobId) {
  return getAttachments(jobId);
}

function api_getTrips(rangeStart, rangeEnd, filters) {
  return getTrips(rangeStart, rangeEnd, filters);
}

/**
 * テスト用：スプレッドシート接続確認
 */
function testConnection() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets().map(s => s.getName());
    Logger.log('接続成功: ' + ss.getName());
    Logger.log('シート一覧: ' + sheets.join(', '));
    return { success: true, name: ss.getName(), sheets };
  } catch (error) {
    Logger.log('接続失敗: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * テスト用：Bootstrap データ取得
 */
function testGetBootstrapData() {
  const data = getBootstrapData();
  Logger.log('Jobs: ' + data.jobs.length + '件');
  Logger.log('Processes: ' + data.processes.length + '件');
  Logger.log('People: ' + data.people.length + '件');
  Logger.log('Schedules: ' + data.schedules.length + '件');
  return data;
}
