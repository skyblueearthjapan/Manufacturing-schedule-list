/**
 * DataService.gs
 * スプレッドシートとのデータ入出力を担当
 */

/**
 * シートデータを配列からオブジェクト配列に変換
 * @param {Array[]} data - 2次元配列（ヘッダー含む）
 * @returns {Object[]}
 */
function sheetDataToObjects(data) {
  if (!data || data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

/**
 * 日付をYYYY-MM-DD形式に変換
 * @param {Date|string} date
 * @returns {string}
 */
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 日時をISO形式に変換
 * @param {Date|string} datetime
 * @returns {string}
 */
function formatDateTime(datetime) {
  if (!datetime) return '';
  if (typeof datetime === 'string') return datetime;
  return new Date(datetime).toISOString();
}

// ============================================
// Jobs（工番マスタ）
// ============================================

/**
 * 全Jobsを取得
 * @returns {Object[]}
 */
function getAllJobs() {
  const sheet = getSheet(CONFIG.SHEETS.JOBS);
  const data = sheet.getDataRange().getValues();
  const jobs = sheetDataToObjects(data);

  return jobs.map(job => ({
    ...job,
    出荷予定日: formatDate(job['出荷予定日']),
    出荷実績日: formatDate(job['出荷実績日']),
    出図予定日: formatDate(job['出図予定日']),
    出図実績日: formatDate(job['出図実績日'])
  }));
}

/**
 * JobIdで1件取得
 * @param {string} jobId
 * @returns {Object|null}
 */
function getJobById(jobId) {
  const jobs = getAllJobs();
  return jobs.find(job => job.jobId === jobId) || null;
}

/**
 * Job新規作成
 * @param {Object} payload
 * @returns {Object}
 */
function createJob(payload) {
  const sheet = getSheet(CONFIG.SHEETS.JOBS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // ID採番
  const jobId = Utilities.getUuid();
  const now = new Date();
  const currentUser = Session.getActiveUser().getEmail() || 'system';

  // 新規行データ作成
  const newRow = headers.map(header => {
    switch(header) {
      case 'jobId': return jobId;
      case '工番': return payload['工番'] || '';
      case '顧客名': return payload['顧客名'] || '';
      case '設備/製品名': return payload['設備/製品名'] || '';
      case '台数': return payload['台数'] || '';
      case '納入先': return payload['納入先'] || '';
      case '出荷予定日': return payload['出荷予定日'] || '';
      case '出荷実績日': return payload['出荷実績日'] || '';
      case '出図予定日': return payload['出図予定日'] || '';
      case '出図実績日': return payload['出図実績日'] || '';
      case '状態': return payload['状態'] || '未着手';
      case '重要メモ': return payload['重要メモ'] || '';
      case 'mainPersonId': return payload['mainPersonId'] || '';
      case 'subPersonId': return payload['subPersonId'] || '';
      case 'updatedAt': return now;
      case 'updatedBy': return currentUser;
      default: return '';
    }
  });

  sheet.appendRow(newRow);

  return {
    jobId,
    ...payload,
    出荷予定日: formatDate(payload['出荷予定日']),
    出図予定日: formatDate(payload['出図予定日']),
    updatedAt: formatDateTime(now),
    updatedBy: currentUser
  };
}

/**
 * Job更新
 * @param {string} jobId
 * @param {Object} patch
 * @param {string} expectedUpdatedAt - 競合検知用（任意）
 * @returns {Object}
 */
function updateJob(jobId, patch, expectedUpdatedAt) {
  const sheet = getSheet(CONFIG.SHEETS.JOBS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 対象行を検索
  const jobIdIndex = headers.indexOf('jobId');
  const updatedAtIndex = headers.indexOf('updatedAt');

  let targetRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][jobIdIndex] === jobId) {
      targetRowIndex = i;
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error('指定された工番が見つかりません');
  }

  // 競合検知（updatedAt列がある場合のみ）
  if (updatedAtIndex !== -1 && expectedUpdatedAt) {
    const currentUpdatedAt = formatDateTime(data[targetRowIndex][updatedAtIndex]);
    if (currentUpdatedAt && currentUpdatedAt !== expectedUpdatedAt) {
      const error = new Error('他のユーザーが更新しました。最新データを取得してください。');
      error.code = 409;
      throw error;
    }
  }

  // 更新
  const now = new Date();
  const currentUser = Session.getActiveUser().getEmail() || 'system';

  headers.forEach((header, colIndex) => {
    if (patch.hasOwnProperty(header)) {
      sheet.getRange(targetRowIndex + 1, colIndex + 1).setValue(patch[header]);
    }
  });

  // updatedAt/By更新（列があれば）
  if (updatedAtIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, updatedAtIndex + 1).setValue(now);
  }
  const updatedByIndex = headers.indexOf('updatedBy');
  if (updatedByIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, updatedByIndex + 1).setValue(currentUser);
  }

  // 更新後データを返す
  const updatedData = sheet.getRange(targetRowIndex + 1, 1, 1, headers.length).getValues()[0];
  const result = {};
  headers.forEach((header, index) => {
    result[header] = updatedData[index];
  });
  result.updatedAt = formatDateTime(now);
  result.出荷予定日 = formatDate(result['出荷予定日']);
  result.出図予定日 = formatDate(result['出図予定日']);

  return result;
}

// ============================================
// ProcessMaster（工程マスタ）
// ============================================

/**
 * 全ProcessMasterを取得（有効のみ）
 * @param {boolean} includeInactive - 無効も含むか
 * @returns {Object[]}
 */
function getAllProcesses(includeInactive = false) {
  const sheet = getSheet(CONFIG.SHEETS.PROCESS_MASTER);
  const data = sheet.getDataRange().getValues();
  let processes = sheetDataToObjects(data);

  if (!includeInactive) {
    processes = processes.filter(p => p['有効(isActive)'] === true || p['有効(isActive)'] === 'TRUE');
  }

  // 表示順でソート
  return processes.sort((a, b) => (a['表示順(order)'] || 0) - (b['表示順(order)'] || 0));
}

// ============================================
// People（作業者マスタ）
// ============================================

/**
 * 全Peopleを取得（有効のみ）
 * @param {boolean} includeInactive - 無効も含むか
 * @returns {Object[]}
 */
function getAllPeople(includeInactive = false) {
  const sheet = getSheet(CONFIG.SHEETS.PEOPLE);
  const data = sheet.getDataRange().getValues();
  let people = sheetDataToObjects(data);

  if (!includeInactive) {
    people = people.filter(p => p['有効(isActive)'] === true || p['有効(isActive)'] === 'TRUE');
  }

  return people;
}

/**
 * PersonIdで1件取得
 * @param {string} personId
 * @returns {Object|null}
 */
function getPersonById(personId) {
  const people = getAllPeople(true);
  return people.find(p => p.personId === personId) || null;
}

// ============================================
// Schedule（工程期間）
// ============================================

/**
 * Scheduleを期間で取得
 * @param {string} rangeStart - 開始日（YYYY-MM-DD）
 * @param {string} rangeEnd - 終了日（YYYY-MM-DD）
 * @param {Object} filters - フィルタ条件
 * @returns {Object[]}
 */
function getSchedules(rangeStart, rangeEnd, filters = {}) {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  let schedules = sheetDataToObjects(data);

  // 日付フォーマット
  schedules = schedules.map(s => ({
    ...s,
    start: formatDate(s.start),
    end: formatDate(s.end),
    updatedAt: formatDateTime(s.updatedAt)
  }));

  // 期間フィルタ（期間が重なるもの）
  if (rangeStart && rangeEnd) {
    schedules = schedules.filter(s => {
      return s.end >= rangeStart && s.start <= rangeEnd;
    });
  }

  // 追加フィルタ
  if (filters.jobIds && filters.jobIds.length > 0) {
    schedules = schedules.filter(s => filters.jobIds.includes(s.jobId));
  }
  if (filters.personIds && filters.personIds.length > 0) {
    schedules = schedules.filter(s =>
      filters.personIds.includes(s.mainPersonId) ||
      filters.personIds.includes(s.subPersonId)
    );
  }
  if (filters.processIds && filters.processIds.length > 0) {
    schedules = schedules.filter(s => filters.processIds.includes(s.processId));
  }
  if (filters.status) {
    schedules = schedules.filter(s => s.status === filters.status);
  }

  return schedules;
}

/**
 * Schedule新規作成
 * @param {Object} payload
 * @returns {Object}
 */
function createSchedule(payload) {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // ID採番
  const scheduleId = Utilities.getUuid();
  const now = new Date();
  const currentUser = Session.getActiveUser().getEmail() || 'system';

  // 新規行データ作成
  const newRow = headers.map(header => {
    switch(header) {
      case 'scheduleId': return scheduleId;
      case 'jobId': return payload.jobId;
      case 'processId': return payload.processId;
      case 'start': return payload.start;
      case 'end': return payload.end;
      case 'mainPersonId': return payload.mainPersonId;
      case 'subPersonId': return payload.subPersonId || '';
      case 'label': return payload.label || '';
      case 'status': return payload.status || '予定';
      case 'memo': return payload.memo || '';
      case 'updatedAt': return now;
      case 'updatedBy': return currentUser;
      default: return '';
    }
  });

  sheet.appendRow(newRow);

  return {
    scheduleId,
    ...payload,
    updatedAt: formatDateTime(now),
    updatedBy: currentUser
  };
}

/**
 * Schedule更新
 * @param {string} scheduleId
 * @param {Object} patch
 * @param {string} expectedUpdatedAt - 競合検知用
 * @returns {Object}
 */
function updateSchedule(scheduleId, patch, expectedUpdatedAt) {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 対象行を検索
  const scheduleIdIndex = headers.indexOf('scheduleId');
  const updatedAtIndex = headers.indexOf('updatedAt');

  let targetRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][scheduleIdIndex] === scheduleId) {
      targetRowIndex = i;
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error('指定されたスケジュールが見つかりません');
  }

  // 競合検知
  const currentUpdatedAt = formatDateTime(data[targetRowIndex][updatedAtIndex]);
  if (expectedUpdatedAt && currentUpdatedAt !== expectedUpdatedAt) {
    const error = new Error('他のユーザーが更新しました。最新データを取得してください。');
    error.code = 409;
    throw error;
  }

  // 更新
  const now = new Date();
  const currentUser = Session.getActiveUser().getEmail() || 'system';

  headers.forEach((header, colIndex) => {
    if (patch.hasOwnProperty(header)) {
      sheet.getRange(targetRowIndex + 1, colIndex + 1).setValue(patch[header]);
    }
  });

  // updatedAt/By更新
  sheet.getRange(targetRowIndex + 1, updatedAtIndex + 1).setValue(now);
  sheet.getRange(targetRowIndex + 1, headers.indexOf('updatedBy') + 1).setValue(currentUser);

  // 更新後データを返す
  const updatedData = sheet.getRange(targetRowIndex + 1, 1, 1, headers.length).getValues()[0];
  const result = {};
  headers.forEach((header, index) => {
    result[header] = updatedData[index];
  });
  result.updatedAt = formatDateTime(now);

  return result;
}

/**
 * Schedule削除
 * @param {string} scheduleId
 * @param {string} expectedUpdatedAt
 * @returns {Object}
 */
function deleteSchedule(scheduleId, expectedUpdatedAt) {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const scheduleIdIndex = headers.indexOf('scheduleId');
  const updatedAtIndex = headers.indexOf('updatedAt');

  let targetRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][scheduleIdIndex] === scheduleId) {
      targetRowIndex = i;
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error('指定されたスケジュールが見つかりません');
  }

  // 競合検知
  const currentUpdatedAt = formatDateTime(data[targetRowIndex][updatedAtIndex]);
  if (expectedUpdatedAt && currentUpdatedAt !== expectedUpdatedAt) {
    const error = new Error('他のユーザーが更新しました。最新データを取得してください。');
    error.code = 409;
    throw error;
  }

  // 物理削除
  sheet.deleteRow(targetRowIndex + 1);

  return { success: true, scheduleId };
}

// ============================================
// Attachments（PDF添付）
// ============================================

/**
 * 工番のAttachmentsを取得
 * @param {string} jobId
 * @returns {Object[]}
 */
function getAttachments(jobId) {
  const sheet = getSheet(CONFIG.SHEETS.ATTACHMENTS);
  const data = sheet.getDataRange().getValues();
  let attachments = sheetDataToObjects(data);

  if (jobId) {
    attachments = attachments.filter(a => a.jobId === jobId);
  }

  return attachments.map(a => ({
    ...a,
    uploadedAt: formatDateTime(a.uploadedAt)
  }));
}

// ============================================
// Trips（出張・移動）
// ============================================

/**
 * Tripsを期間で取得
 * @param {string} rangeStart
 * @param {string} rangeEnd
 * @param {Object} filters
 * @returns {Object[]}
 */
function getTrips(rangeStart, rangeEnd, filters = {}) {
  const sheet = getSheet(CONFIG.SHEETS.TRIPS);
  const data = sheet.getDataRange().getValues();
  let trips = sheetDataToObjects(data);

  trips = trips.map(t => ({
    ...t,
    start: formatDate(t.start),
    end: formatDate(t.end)
  }));

  // 期間フィルタ
  if (rangeStart && rangeEnd) {
    trips = trips.filter(t => {
      return t.end >= rangeStart && t.start <= rangeEnd;
    });
  }

  // personIdフィルタ
  if (filters.personIds && filters.personIds.length > 0) {
    trips = trips.filter(t => filters.personIds.includes(t.personId));
  }

  return trips;
}

// ============================================
// Bootstrap（初期データ一括取得）
// ============================================

/**
 * 初期表示に必要なデータを一括取得
 * @param {string} rangeStart - 表示開始日
 * @param {number} days - 表示日数
 * @returns {Object}
 */
function getBootstrapData(rangeStart, days = CONFIG.DEFAULT_DISPLAY_DAYS) {
  const start = rangeStart || formatDate(new Date());
  const endDate = new Date(start);
  endDate.setDate(endDate.getDate() + days);
  const end = formatDate(endDate);

  return {
    jobs: getAllJobs(),
    processes: getAllProcesses(),
    people: getAllPeople(),
    schedules: getSchedules(start, end),
    attachments: [], // 初期は空、必要時に取得
    trips: getTrips(start, end),
    jobMaster: getExternalJobMaster(), // 外部工番マスター
    meta: {
      rangeStart: start,
      rangeEnd: end,
      days: days,
      fetchedAt: new Date().toISOString()
    }
  };
}

// ============================================
// External Job Master（外部工番マスター）
// ============================================

/**
 * 外部スプレッドシートから工番マスターを同期
 * 本アプリ側の「外部_工番マスター」シートに丸ごとコピー
 * @returns {Object} 同期結果
 */
function syncExternalJobMaster() {
  try {
    // 外部スプレッドシートを開く
    const externalSs = SpreadsheetApp.openById(CONFIG.EXTERNAL_MASTER_SPREADSHEET_ID);
    const externalSheet = externalSs.getSheetByName(CONFIG.EXTERNAL_MASTER_SHEET_NAME);

    if (!externalSheet) {
      throw new Error(`外部シート "${CONFIG.EXTERNAL_MASTER_SHEET_NAME}" が見つかりません`);
    }

    // 外部データを取得
    const externalData = externalSheet.getDataRange().getValues();
    if (externalData.length < 1) {
      throw new Error('外部工番マスターにデータがありません');
    }

    // 本アプリ側のシートを取得（なければ作成）
    const ss = getSpreadsheet();
    let localSheet = ss.getSheetByName(CONFIG.SHEETS.EXTERNAL_JOB_MASTER);

    if (!localSheet) {
      localSheet = ss.insertSheet(CONFIG.SHEETS.EXTERNAL_JOB_MASTER);
    }

    // 既存データをクリアして新データを書き込み
    localSheet.clearContents();
    localSheet.getRange(1, 1, externalData.length, externalData[0].length).setValues(externalData);

    const syncResult = {
      success: true,
      rowCount: externalData.length - 1, // ヘッダー除く
      syncedAt: new Date().toISOString()
    };

    Logger.log('外部工番マスター同期完了: ' + syncResult.rowCount + '件');
    return syncResult;

  } catch (error) {
    Logger.log('外部工番マスター同期エラー: ' + error.message);
    throw error;
  }
}

/**
 * 外部工番マスターを取得
 * @returns {Object[]}
 */
function getExternalJobMaster() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EXTERNAL_JOB_MASTER);

    if (!sheet) {
      // シートがない場合は空配列を返す（初回同期前）
      return [];
    }

    const data = sheet.getDataRange().getValues();
    return sheetDataToObjects(data);
  } catch (error) {
    Logger.log('外部工番マスター取得エラー: ' + error.message);
    return [];
  }
}

/**
 * 外部工番マスターを検索
 * @param {string} query - 検索クエリ（工番/客先名/製品名に部分一致）
 * @param {number} limit - 最大件数
 * @returns {Object[]}
 */
function searchExternalJobMaster(query, limit = 20) {
  const allData = getExternalJobMaster();

  if (!query || query.trim() === '') {
    return allData.slice(0, limit);
  }

  const lowerQuery = query.toLowerCase();
  const results = allData.filter(item => {
    const jobNo = String(item['工番'] || '').toLowerCase();
    const customer = String(item['客先名'] || '').toLowerCase();
    const product = String(item['製品名'] || '').toLowerCase();

    return jobNo.includes(lowerQuery) ||
           customer.includes(lowerQuery) ||
           product.includes(lowerQuery);
  });

  return results.slice(0, limit);
}
