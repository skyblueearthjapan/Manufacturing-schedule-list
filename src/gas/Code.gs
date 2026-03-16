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
  var mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : '';

  // mode未指定 → ランディング（起動分岐）画面
  if (!mode) {
    var landingTemplate = HtmlService.createTemplateFromFile('landing');
    landingTemplate.BASE_URL = ScriptApp.getService().getUrl();
    return landingTemplate.evaluate()
      .setTitle('生産工程表')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // mode=viewer → タブレット用ビュアー（閲覧専用）
  if (mode === 'viewer') {
    var viewerTemplate = HtmlService.createTemplateFromFile('viewer');
    viewerTemplate.PORTAL_URL = 'https://script.google.com/a/macros/lineworks-local.info/s/AKfycbx2eyJMOYP9o--GPBuhY-pj071IIR6Kqb_0xALwwNzdLQZux0dIAlL3P9EoCucnzXA/exec';
    var viewerEmail = getCurrentUserEmail_();
    viewerTemplate.USER_EMAIL = viewerEmail;
    viewerTemplate.CAN_EDIT = false; // ビュアーは常に閲覧専用
    return viewerTemplate.evaluate()
      .setTitle('生産工程表 - ビュアー')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // mode=pc（またはその他） → 既存のPC版 index.html
  const template = HtmlService.createTemplateFromFile('index');
  template.PORTAL_URL = 'https://script.google.com/a/macros/lineworks-local.info/s/AKfycbx2eyJMOYP9o--GPBuhY-pj071IIR6Kqb_0xALwwNzdLQZux0dIAlL3P9EoCucnzXA/exec';
  // 権限情報を注入
  const userEmail = getCurrentUserEmail_();
  template.USER_EMAIL = userEmail;
  template.CAN_EDIT = isEditorEmail(userEmail);
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
  requireEditor();
  return sanitizeForClient(uploadJobPdf(jobNo, pdfType, fileName, base64Data));
}

function api_deleteJobPdf(jobNo, pdfType) {
  requireEditor();
  return sanitizeForClient(deleteJobPdf(jobNo, pdfType));
}

function api_getJobPdfInfo(jobNo) {
  return sanitizeForClient(getJobPdfInfo(jobNo));
}

function api_setDaySetting(dateISO, type, memo) {
  requireEditor();
  return sanitizeForClient(setDaySetting(dateISO, type, memo));
}

function api_getDaySettings(fromISO, toISO) {
  return sanitizeForClient(getDaySettingsMap(fromISO, toISO));
}

// ========== Vehicle API ==========
// api_checkVehicleAvailability(startDate, endDate, excludeTripId) は
// VehicleSyncService.gs で直接定義済み（google.script.run から呼び出し可能）

// ========== Calendar API ==========
function api_getCalendarEvents(calendarId, startDate, endDate) {
  return sanitizeForClient(getCalendarEvents(calendarId, startDate, endDate));
}

function api_getTSCCalendarEvents(startDate, endDate) {
  return sanitizeForClient(getTSCCalendarEvents(startDate, endDate));
}

function api_getElectricalCalendarEvents(startDate, endDate) {
  return sanitizeForClient(getElectricalCalendarEvents(startDate, endDate));
}

function api_testCalendarConnection() {
  return sanitizeForClient(testCalendarConnection());
}

/**
 * TSCカレンダーを出張計画に同期
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} 同期結果
 */
function api_syncTSCCalendarToTrips(startDate, endDate) {
  requireEditor();
  return sanitizeForClient(syncTSCCalendarToTrips(startDate, endDate));
}

/**
 * TSCカレンダー同期テスト（ドライラン）
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} テスト結果
 */
function api_testSyncTSCCalendar(startDate, endDate) {
  return sanitizeForClient(testSyncTSCCalendar(startDate, endDate));
}

/**
 * TSC部署メンバー一覧を取得
 * @returns {Object} メンバー情報
 */
function api_getTSCMembers() {
  const memberIds = getTSCMemberIds();
  const members = memberIds.map(id => {
    const person = getPersonById(id);
    return person ? { personId: id, name: person['表示名'] || person.name || id } : { personId: id, name: id };
  });
  return sanitizeForClient({ memberIds, members });
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

// ========== 権限管理 ==========

/**
 * Permissionsシートからeditor権限のメール一覧を取得
 * @returns {string[]} editorメールアドレスの配列
 */
function getEditorEmails() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PERMISSIONS);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const emailIdx = headers.indexOf('email');
    const roleIdx = headers.indexOf('role');
    const activeIdx = headers.indexOf('isActive');
    if (emailIdx === -1 || roleIdx === -1) return [];
    return data.slice(1)
      .filter(row => {
        const role = String(row[roleIdx]).toLowerCase();
        const active = activeIdx === -1 ? true : row[activeIdx] === true || String(row[activeIdx]).toUpperCase() === 'TRUE';
        return role === 'editor' && active;
      })
      .map(row => String(row[emailIdx]).toLowerCase().trim());
  } catch (e) {
    Logger.log('[getEditorEmails] error: ' + e.message);
    return [];
  }
}

/**
 * 指定メールがeditorか判定
 * @param {string} email
 * @returns {boolean}
 */
function isEditorEmail(email) {
  if (!email) return false;
  const editors = getEditorEmails();
  return editors.includes(email.toLowerCase().trim());
}

/**
 * 現在のユーザーのメールを取得（GASウェブアプリ対応）
 * google.script.run経由ではgetActiveUser()が空を返す場合があるため、
 * getEffectiveUser()もフォールバックとして使用
 * @returns {string}
 */
function getCurrentUserEmail_() {
  // 1. getActiveUser（doGetでは正しく返る）
  let email = '';
  try {
    email = Session.getActiveUser().getEmail() || '';
  } catch (e) {
    Logger.log('[getCurrentUserEmail_] getActiveUser failed: ' + e.message);
  }
  if (email) return email;

  // 2. getEffectiveUser（google.script.run経由のフォールバック）
  try {
    email = Session.getEffectiveUser().getEmail() || '';
  } catch (e) {
    Logger.log('[getCurrentUserEmail_] getEffectiveUser failed: ' + e.message);
  }
  return email;
}

/**
 * 現在のユーザーがeditorか判定
 * @returns {boolean}
 */
function isCurrentUserEditor() {
  const email = getCurrentUserEmail_();
  return isEditorEmail(email);
}

/**
 * editor権限がなければエラーをスロー
 * GASウェブアプリの「自分として実行」設定では、google.script.run経由で
 * ユーザーメールが取得できない場合がある。その場合はクライアント側の
 * EditGuardに委ねてサーバー側はパスする。
 */
function requireEditor() {
  const email = getCurrentUserEmail_();
  // メールが取得できない場合（google.script.runコンテキスト）は
  // クライアント側EditGuardに委ねる
  if (!email) return;
  if (!isEditorEmail(email)) {
    const err = new Error('編集権限がありません。管理者にお問い合わせください。');
    err.code = 403;
    throw err;
  }
}

/**
 * クライアント用: ユーザー権限情報取得API
 * @returns {Object} { email, canEdit }
 */
function api_getUserPermission() {
  const email = getCurrentUserEmail_();
  return { email: email, canEdit: isEditorEmail(email) };
}
