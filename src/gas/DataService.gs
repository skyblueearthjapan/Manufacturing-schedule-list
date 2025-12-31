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

  // isWeekly を boolean に正規化
  processes = processes.map(p => ({
    ...p,
    isWeekly: p['重要（isWeekly）'] === true || p['重要（isWeekly）'] === 'TRUE' || p['重要（isWeekly）'] === 'true'
  }));

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

/**
 * 色を正規化（サーバー側）
 * @param {string} input
 * @returns {string}
 */
function normalizeColorServer(input) {
  if (!input) return '#6B7280';

  const str = String(input).trim();

  // 既に#RRGGBB形式
  if (/^#[0-9A-Fa-f]{6}$/i.test(str)) {
    return str.toUpperCase();
  }

  // 日本語色名マッピング
  const colorMap = {
    '青': '#2563EB',
    '赤': '#EF4444',
    '緑': '#22C55E',
    '黄': '#F59E0B',
    '黄色': '#F59E0B',
    '紫': '#A855F7',
    '橙': '#F97316',
    'オレンジ': '#F97316',
    '黒': '#111827',
    '灰': '#6B7280',
    '灰色': '#6B7280',
    'グレー': '#6B7280',
    '白': '#FFFFFF',
    'ピンク': '#EC4899',
    '水色': '#06B6D4',
    '茶': '#92400E',
    '茶色': '#92400E'
  };

  return colorMap[str] || colorMap[str.toLowerCase()] || '#6B7280';
}

/**
 * 担当者を新規作成
 * @param {Object} payload
 * @returns {Object}
 */
function createPerson(payload) {
  const sheet = getSheet(CONFIG.SHEETS.PEOPLE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 新しいpersonIdを生成（P.XX形式）
  const existing = sheetDataToObjects(data);
  let maxNum = 0;
  existing.forEach(p => {
    const match = String(p.personId).match(/^P\.(\d+)$/);
    if (match) {
      maxNum = Math.max(maxNum, parseInt(match[1]));
    }
  });
  const newPersonId = `P.${String(maxNum + 1).padStart(2, '0')}`;

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'system';

  // 色を正規化
  const colorHex = normalizeColorServer(payload['作業者色(colorHex)']);

  const newPerson = {
    personId: newPersonId,
    '氏名': payload['氏名'] || '',
    '部署/区分': payload['部署/区分'] || '',
    '作業者色(colorHex)': colorHex,
    '有効(isActive)': payload['有効(isActive)'] !== false,
    '備考': payload['備考'] || '',
    updatedAt: now,
    updatedBy: user
  };

  // 行データを作成
  const rowData = headers.map(header => newPerson[header] ?? '');
  sheet.appendRow(rowData);

  return {
    ...newPerson,
    updatedAt: formatDateTime(now)
  };
}

/**
 * 担当者を更新
 * @param {string} personId
 * @param {Object} patch
 * @param {string} expectedUpdatedAt - 楽観的ロック用
 * @returns {Object}
 */
function updatePerson(personId, patch, expectedUpdatedAt) {
  const sheet = getSheet(CONFIG.SHEETS.PEOPLE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 対象行を検索
  const personIdIndex = headers.indexOf('personId');
  const updatedAtIndex = headers.indexOf('updatedAt');

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][personIdIndex] === personId) {
      targetRow = i + 1; // 1-based
      break;
    }
  }

  if (targetRow === -1) {
    const error = new Error('担当者が見つかりません');
    error.code = 404;
    throw error;
  }

  // 楽観的ロック確認
  if (expectedUpdatedAt && updatedAtIndex !== -1) {
    const currentUpdatedAt = formatDateTime(data[targetRow - 1][updatedAtIndex]);
    if (currentUpdatedAt !== expectedUpdatedAt) {
      const error = new Error('他のユーザーによって更新されています。ページを再読み込みしてください。');
      error.code = 409;
      throw error;
    }
  }

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'system';

  // 色を正規化
  if (patch['作業者色(colorHex)']) {
    patch['作業者色(colorHex)'] = normalizeColorServer(patch['作業者色(colorHex)']);
  }

  // 更新データをマージ
  const currentRow = data[targetRow - 1];
  const updated = {};
  headers.forEach((header, index) => {
    if (patch.hasOwnProperty(header)) {
      updated[header] = patch[header];
    } else {
      updated[header] = currentRow[index];
    }
  });
  updated.updatedAt = now;
  updated.updatedBy = user;

  // 行を更新
  const rowData = headers.map(header => updated[header] ?? '');
  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

  return {
    ...updated,
    updatedAt: formatDateTime(now)
  };
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
    end: formatDate(t.end),
    updatedAt: formatDateTime(t.updatedAt),
    // isLockedをboolean化
    isLocked: t.isLocked === true || t.isLocked === 'TRUE' || t.isLocked === 'true'
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

/**
 * genKey（同一性キー）を生成
 * 形式: jobId|processId|start|end|personId|kind
 * @param {Object} trip
 * @returns {string}
 */
function generateGenKey(trip) {
  const jobId = trip['jobId(任意)'] || trip.jobId || '';
  const processId = trip.processId || '';
  const start = trip.start || '';
  const end = trip.end || '';
  const personId = trip.personId || '';
  const kind = trip.kind || '';
  return `${jobId}|${processId}|${start}|${end}|${personId}|${kind}`;
}

/**
 * Trip新規作成
 * @param {Object} payload
 * @returns {Object}
 */
function createTrip(payload) {
  const sheet = getSheet(CONFIG.SHEETS.TRIPS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // ID採番
  const tripId = Utilities.getUuid();
  const now = new Date();

  // genKeyを生成
  const genKey = payload.genKey || generateGenKey(payload);

  // ヘッダーベースで行データ作成
  const newRow = headers.map(header => {
    switch(header) {
      case 'tripId': return tripId;
      case 'personId': return payload.personId || '';
      case 'start': return payload.start || '';
      case 'end': return payload.end || '';
      case '行先': return payload['行先'] || '';
      case '用件': return payload['用件'] || '';
      case 'jobId(任意)': return payload['jobId(任意)'] || payload.jobId || '';
      case '備考': return payload['備考'] || '';
      case 'kind': return payload.kind || '';
      case 'processId': return payload.processId || '';
      case 'source': return payload.source || 'manual';
      case 'isLocked': return payload.isLocked || false;
      case 'genKey': return genKey;
      case 'updatedAt': return now;
      default: return payload[header] || '';
    }
  });

  sheet.appendRow(newRow);

  return {
    tripId,
    ...payload,
    genKey,
    start: formatDate(payload.start),
    end: formatDate(payload.end),
    updatedAt: formatDateTime(now)
  };
}

/**
 * Trip更新（ロック済みは拒否）
 * @param {string} tripId
 * @param {Object} patch
 * @param {string} expectedUpdatedAt - 競合検知用
 * @returns {Object}
 */
function updateTrip(tripId, patch, expectedUpdatedAt) {
  const sheet = getSheet(CONFIG.SHEETS.TRIPS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 対象行を検索（ヘッダーベース）
  const tripIdIndex = headers.indexOf('tripId');
  const updatedAtIndex = headers.indexOf('updatedAt');
  const isLockedIndex = headers.indexOf('isLocked');
  const sourceIndex = headers.indexOf('source');

  if (tripIdIndex === -1) {
    throw new Error('tripId列が見つかりません');
  }

  let targetRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][tripIdIndex] === tripId) {
      targetRowIndex = i;
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error('指定された出張予定が見つかりません');
  }

  // ロック済みチェック（isLocked更新以外は拒否）
  if (isLockedIndex !== -1) {
    const currentIsLocked = data[targetRowIndex][isLockedIndex];
    const isLocked = currentIsLocked === true || currentIsLocked === 'TRUE' || currentIsLocked === 'true';
    // isLockedの更新リクエストでなければ、ロック済みを拒否
    if (isLocked && !patch.hasOwnProperty('isLocked')) {
      const error = new Error('この予定は確定済みのため編集できません。');
      error.code = 403;
      throw error;
    }
  }

  // 競合検知
  if (expectedUpdatedAt && updatedAtIndex !== -1) {
    const currentUpdatedAt = formatDateTime(data[targetRowIndex][updatedAtIndex]);
    if (currentUpdatedAt && currentUpdatedAt !== expectedUpdatedAt) {
      const error = new Error('他のユーザーが更新しました。最新データを取得してください。');
      error.code = 409;
      throw error;
    }
  }

  // 更新
  const now = new Date();

  // AUTO→MANUALへの変更（isLocked更新以外の編集の場合）
  if (sourceIndex !== -1 && !patch.hasOwnProperty('isLocked')) {
    const currentSource = data[targetRowIndex][sourceIndex];
    if (currentSource === 'auto' && !patch.hasOwnProperty('source')) {
      patch.source = 'manual';
    }
  }

  headers.forEach((header, colIndex) => {
    if (patch.hasOwnProperty(header)) {
      sheet.getRange(targetRowIndex + 1, colIndex + 1).setValue(patch[header]);
    }
  });

  // updatedAt更新
  if (updatedAtIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, updatedAtIndex + 1).setValue(now);
  }

  // 更新後データを返す
  const updatedData = sheet.getRange(targetRowIndex + 1, 1, 1, headers.length).getValues()[0];
  const result = {};
  headers.forEach((header, index) => {
    result[header] = updatedData[index];
  });
  result.start = formatDate(result.start);
  result.end = formatDate(result.end);
  result.updatedAt = formatDateTime(now);
  result.isLocked = result.isLocked === true || result.isLocked === 'TRUE' || result.isLocked === 'true';

  return result;
}

/**
 * Trip確定（ロック）
 * @param {string} tripId
 * @param {boolean} isLocked - true=ロック, false=解除
 * @returns {Object}
 */
function lockTrip(tripId, isLocked = true) {
  const sheet = getSheet(CONFIG.SHEETS.TRIPS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const tripIdIndex = headers.indexOf('tripId');
  const isLockedIndex = headers.indexOf('isLocked');
  const updatedAtIndex = headers.indexOf('updatedAt');

  if (tripIdIndex === -1) {
    throw new Error('tripId列が見つかりません');
  }
  if (isLockedIndex === -1) {
    throw new Error('isLocked列が見つかりません');
  }

  let targetRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][tripIdIndex] === tripId) {
      targetRowIndex = i;
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error('指定された出張予定が見つかりません');
  }

  const now = new Date();

  // isLocked更新
  sheet.getRange(targetRowIndex + 1, isLockedIndex + 1).setValue(isLocked);

  // updatedAt更新
  if (updatedAtIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, updatedAtIndex + 1).setValue(now);
  }

  // 更新後データを返す
  const updatedData = sheet.getRange(targetRowIndex + 1, 1, 1, headers.length).getValues()[0];
  const result = {};
  headers.forEach((header, index) => {
    result[header] = updatedData[index];
  });
  result.start = formatDate(result.start);
  result.end = formatDate(result.end);
  result.updatedAt = formatDateTime(now);
  result.isLocked = isLocked;

  return result;
}

/**
 * Trip削除
 * @param {string} tripId
 * @returns {Object}
 */
function deleteTrip(tripId) {
  const sheet = getSheet(CONFIG.SHEETS.TRIPS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const tripIdIndex = headers.indexOf('tripId');
  if (tripIdIndex === -1) {
    throw new Error('tripId列が見つかりません');
  }

  let targetRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][tripIdIndex] === tripId) {
      targetRowIndex = i;
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error('指定された出張予定が見つかりません');
  }

  // 物理削除
  sheet.deleteRow(targetRowIndex + 1);

  return { success: true, tripId };
}

/**
 * 出張計画自動生成（P150スケジュールから）
 * @param {string} rangeStart - 対象期間開始
 * @param {string} rangeEnd - 対象期間終了
 * @returns {Object} 生成結果
 */
function generateTravelPlan(rangeStart, rangeEnd) {
  const PROCESS_ID_SITE_WORK = 'P150'; // 据付/現地工事の工程ID

  // 既存Tripsを取得
  const existingTrips = getTrips(null, null, {}); // 全件取得

  // P150のスケジュールを取得
  const schedules = getSchedules(rangeStart, rangeEnd, { processIds: [PROCESS_ID_SITE_WORK] });

  // Jobsを取得（mainPersonId/subPersonId取得用）
  const jobs = getAllJobs();
  const jobMap = {};
  jobs.forEach(j => { jobMap[j.jobId] = j; });

  const created = [];
  const updated = [];
  const skipped = [];

  schedules.forEach(schedule => {
    const job = jobMap[schedule.jobId];
    if (!job) {
      skipped.push({ scheduleId: schedule.scheduleId, reason: 'Job not found' });
      return;
    }

    // 対象者リスト（mainPersonId + subPersonId）
    const personIds = [];
    if (job.mainPersonId) personIds.push(job.mainPersonId);
    if (job.subPersonId && job.subPersonId !== job.mainPersonId) personIds.push(job.subPersonId);

    if (personIds.length === 0) {
      skipped.push({ scheduleId: schedule.scheduleId, reason: 'No person assigned' });
      return;
    }

    // スケジュールの日付
    const siteStart = schedule.start;
    const siteEnd = schedule.end;
    const moveBeforeDate = addDaysServer(siteStart, -1);
    const moveAfterDate = addDaysServer(siteEnd, 1);

    // 用件テキスト生成（3行形式）
    const jobNo = job['工番'] || '';
    const customerName = job['顧客名'] || '';
    const productName = job['設備/製品名'] || '';
    const destination = job['納入先'] || '';
    const siteDescription = `${jobNo} ${customerName}\n${productName}\n${destination}`;
    const moveDescription = '移動';

    personIds.forEach(personId => {
      // 1. 前日移動
      const moveBefore = {
        personId,
        start: moveBeforeDate,
        end: moveBeforeDate,
        '行先': destination,
        '用件': moveDescription,
        'jobId(任意)': schedule.jobId,
        kind: 'move',
        processId: PROCESS_ID_SITE_WORK,
        source: 'auto',
        isLocked: false
      };
      const moveBeforeResult = upsertTrip(existingTrips, moveBefore);
      if (moveBeforeResult.action === 'created') created.push(moveBeforeResult.trip);
      else if (moveBeforeResult.action === 'updated') updated.push(moveBeforeResult.trip);
      else skipped.push({ ...moveBefore, reason: moveBeforeResult.reason });

      // 2. 現場（site）
      const site = {
        personId,
        start: siteStart,
        end: siteEnd,
        '行先': destination,
        '用件': siteDescription,
        'jobId(任意)': schedule.jobId,
        kind: 'site',
        processId: PROCESS_ID_SITE_WORK,
        source: 'auto',
        isLocked: false
      };
      const siteResult = upsertTrip(existingTrips, site);
      if (siteResult.action === 'created') created.push(siteResult.trip);
      else if (siteResult.action === 'updated') updated.push(siteResult.trip);
      else skipped.push({ ...site, reason: siteResult.reason });

      // 3. 翌日移動
      const moveAfter = {
        personId,
        start: moveAfterDate,
        end: moveAfterDate,
        '行先': destination,
        '用件': moveDescription,
        'jobId(任意)': schedule.jobId,
        kind: 'move',
        processId: PROCESS_ID_SITE_WORK,
        source: 'auto',
        isLocked: false
      };
      const moveAfterResult = upsertTrip(existingTrips, moveAfter);
      if (moveAfterResult.action === 'created') created.push(moveAfterResult.trip);
      else if (moveAfterResult.action === 'updated') updated.push(moveAfterResult.trip);
      else skipped.push({ ...moveAfter, reason: moveAfterResult.reason });
    });
  });

  return {
    success: true,
    created: created.length,
    updated: updated.length,
    skipped: skipped.length,
    details: { created, updated, skipped }
  };
}

/**
 * Trip upsert（genKeyで同一性判定）
 * genKey形式: jobId|processId|start|end|personId|kind
 * @param {Object[]} existingTrips - 既存Trips配列（参照用）
 * @param {Object} newTrip - 新規Trip
 * @returns {Object} { action: 'created'|'updated'|'skipped', trip?, reason? }
 */
function upsertTrip(existingTrips, newTrip) {
  // genKeyを生成
  const newGenKey = generateGenKey(newTrip);
  newTrip.genKey = newGenKey;

  // genKeyでマッチ（既存にgenKeyがある場合はそれを使用、なければ動的生成）
  const existing = existingTrips.find(t => {
    const existingGenKey = t.genKey || generateGenKey(t);
    return existingGenKey === newGenKey;
  });

  if (existing) {
    // 既存がロック済みなら何もしない（最優先）
    if (existing.isLocked) {
      return { action: 'skipped', reason: 'isLocked=true' };
    }
    // 既存がmanualなら上書きしない
    if (existing.source === 'manual') {
      return { action: 'skipped', reason: 'source=manual' };
    }
    // 既存がautoなら更新（内容のみ、sourceはautoのまま）
    if (existing.source === 'auto') {
      // updateTripはロック済みチェックがあるが、AUTO→AUTO更新なのでisLockedは渡さない
      const updated = updateTrip(existing.tripId, {
        '行先': newTrip['行先'],
        '用件': newTrip['用件'],
        '備考': newTrip['備考'] || '',
        source: 'auto' // 明示的にautoを維持
      });
      return { action: 'updated', trip: updated };
    }
    // その他（sourceが未設定など）は上書きしない
    return { action: 'skipped', reason: 'source unknown' };
  }

  // 新規作成
  const created = createTrip(newTrip);
  // existingTripsに追加（後続の重複判定用）
  existingTrips.push(created);
  return { action: 'created', trip: created };
}

/**
 * サーバー側で日付に日数を加算
 * @param {string} dateStr - YYYY-MM-DD形式
 * @param {number} days - 加算日数
 * @returns {string} YYYY-MM-DD形式
 */
function addDaysServer(dateStr, days) {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + days);
  return formatDate(d);
}

// ============================================
// WorkerJobAssign（作業者別担当工番）
// ============================================

/**
 * 全WorkerJobAssignを取得
 * @param {boolean} activeOnly - 有効のみ（デフォルトtrue）
 * @returns {Object[]}
 */
function getAllWorkerJobAssigns(activeOnly = true) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.WORKER_JOB_ASSIGN);
    const data = sheet.getDataRange().getValues();
    let assigns = sheetDataToObjects(data);

    // isActiveをboolean化
    assigns = assigns.map(a => ({
      ...a,
      isActive: a.isActive === true || a.isActive === 'TRUE' || a.isActive === 'true' || a.isActive === ''
    }));

    // 有効のみフィルタ
    if (activeOnly) {
      assigns = assigns.filter(a => a.isActive !== false);
    }

    // priority（小さい順）→ jobCode 順でソート
    assigns.sort((a, b) => {
      const priorityA = Number(a.priority) || 9999;
      const priorityB = Number(b.priority) || 9999;
      if (priorityA !== priorityB) return priorityA - priorityB;
      return String(a.jobCode || '').localeCompare(String(b.jobCode || ''));
    });

    return assigns;
  } catch (e) {
    // シートがない場合は空配列を返す
    Logger.log('WorkerJobAssign取得エラー（シートなしの可能性）: ' + e.message);
    return [];
  }
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

  // 外部工番マスターは取得失敗しても続行
  let jobMaster = [];
  try {
    jobMaster = getExternalJobMaster();
  } catch (e) {
    Logger.log('外部工番マスター取得をスキップ: ' + e.message);
  }

  return {
    jobs: getAllJobs(),
    processes: getAllProcesses(),
    people: getAllPeople(),
    schedules: getSchedules(start, end),
    attachments: [], // 初期は空、必要時に取得
    trips: getTrips(start, end),
    jobMaster: jobMaster,
    workerJobAssign: getAllWorkerJobAssigns(),
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
