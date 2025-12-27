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
    meta: {
      rangeStart: start,
      rangeEnd: end,
      days: days,
      fetchedAt: new Date().toISOString()
    }
  };
}
