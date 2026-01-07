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

/**
 * DateオブジェクトをISO文字列に変換（google.script.run対応）
 * @param {*} obj
 * @returns {*}
 */
function sanitizeForClient(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function api_ping() {
  return { ok: true, at: formatDateTime(new Date()), version: '2025-12-31-v2' };
}

function api_getBootstrapData(rangeStart, days) {
  Logger.log('[api_getBootstrapData] called: rangeStart=%s, days=%s', rangeStart, days);
  try {
    const result = getBootstrapData(rangeStart, days);
    Logger.log('[api_getBootstrapData] success: jobs=%s, schedules=%s',
      result?.jobs?.length || 0, result?.schedules?.length || 0);
    // Dateオブジェクトをシリアライズしてからクライアントへ返す
    return sanitizeForClient(result);
  } catch (e) {
    Logger.log('[api_getBootstrapData] error: %s', e.message);
    throw e;
  }
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

function api_createJob(payload) {
  return createJob(payload);
}

function api_updateJob(jobId, patch, expectedUpdatedAt) {
  return updateJob(jobId, patch, expectedUpdatedAt);
}

function api_getJob(jobId) {
  return sanitizeForClient(getJobById(jobId));
}

function api_syncExternalJobMaster() {
  return syncExternalJobMaster();
}

function api_searchExternalJobMaster(query, limit) {
  return searchExternalJobMaster(query, limit);
}

function api_createPerson(payload) {
  return sanitizeForClient(createPerson(payload));
}

function api_updatePerson(personId, patch, expectedUpdatedAt) {
  return sanitizeForClient(updatePerson(personId, patch, expectedUpdatedAt));
}

function api_getPeople(includeInactive) {
  return sanitizeForClient(getAllPeople(includeInactive));
}

// ========== Trip API ==========
function api_createTrip(payload) {
  return sanitizeForClient(createTrip(payload));
}

function api_updateTrip(tripId, patch, expectedUpdatedAt) {
  return sanitizeForClient(updateTrip(tripId, patch, expectedUpdatedAt));
}

function api_deleteTrip(tripId) {
  return deleteTrip(tripId);
}

function api_generateTravelPlan(rangeStart, rangeEnd) {
  return sanitizeForClient(generateTravelPlan(rangeStart, rangeEnd));
}

function api_lockTrip(tripId, isLocked) {
  return sanitizeForClient(lockTrip(tripId, isLocked !== false));
}

function api_createProcess(payload) {
  return sanitizeForClient(createProcess(payload));
}

// TopMemo API
function api_upsertTopMemo(payload) {
  return sanitizeForClient(upsertTopMemo(payload));
}

function api_reorderTopMemos(orderPayload) {
  return sanitizeForClient(reorderTopMemos(orderPayload));
}

function api_setTopMemoActive(memoId, isActive) {
  return sanitizeForClient(setTopMemoActive(memoId, isActive));
}

// ========== PDF管理 API ==========
function api_uploadJobPdf(jobNo, pdfType, fileName, base64Data) {
  return sanitizeForClient(uploadJobPdf(jobNo, pdfType, fileName, base64Data));
}

function api_deleteJobPdf(jobNo, pdfType) {
  return sanitizeForClient(deleteJobPdf(jobNo, pdfType));
}

function api_getJobPdfInfo(jobNo) {
  return sanitizeForClient(getJobPdfInfo(jobNo));
}

function api_setDaySetting(dateISO, type, memo) {
  return sanitizeForClient(setDaySetting(dateISO, type, memo));
}

function api_getDaySettings(fromISO, toISO) {
  return sanitizeForClient(getDaySettingsMap(fromISO, toISO));
}

/**
 * テスト用：スプレッドシート接続確認
 * GASエディタで実行して確認してください
 */
function testConnection() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets().map(s => s.getName());
    Logger.log('✓ 接続成功: ' + ss.getName());
    Logger.log('✓ シート一覧: ' + sheets.join(', '));

    // 必要なシートをチェック
    const requiredSheets = ['Jobs', 'ProcessMaster', 'People', 'Schedule', 'Attachments', 'Trips'];
    const missingSheets = requiredSheets.filter(name => !sheets.includes(name));

    if (missingSheets.length > 0) {
      Logger.log('✗ 不足シート: ' + missingSheets.join(', '));
      Logger.log('→ これらのシートをスプレッドシートに作成してください');
    } else {
      Logger.log('✓ 全シート存在');
    }

    return { success: true, name: ss.getName(), sheets, missingSheets };
  } catch (error) {
    Logger.log('✗ 接続失敗: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * テスト用：Bootstrap データ取得
 * GASエディタで実行して確認してください
 */
function testGetBootstrapData() {
  try {
    Logger.log('Bootstrap データ取得開始...');
    const data = getBootstrapData();
    Logger.log('✓ Jobs: ' + data.jobs.length + '件');
    Logger.log('✓ Processes: ' + data.processes.length + '件');
    Logger.log('✓ People: ' + data.people.length + '件');
    Logger.log('✓ Schedules: ' + data.schedules.length + '件');
    Logger.log('✓ 取得成功');
    return data;
  } catch (error) {
    Logger.log('✗ エラー: ' + error.message);
    throw error;
  }
}
