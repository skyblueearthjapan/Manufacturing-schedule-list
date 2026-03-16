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
 * 日付をYYYY-MM-DD形式に変換（JST）
 * @param {Date|string} date
 * @returns {string}
 */
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  // Utilities.formatDateを使用してタイムゾーンを明示的にJSTに設定
  return Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy-MM-dd');
}

/**
 * 日時をISO形式に変換（JST、タイムゾーンオフセット付き）
 * @param {Date|string} datetime
 * @returns {string}
 */
function formatDateTime(datetime) {
  if (!datetime) return '';
  if (typeof datetime === 'string') return datetime;
  // Utilities.formatDateを使用してJSTで出力（ISO 8601形式、+09:00オフセット付き）
  return Utilities.formatDate(new Date(datetime), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss'+09:00'");
}

/**
 * 日時をミリ秒（エポック）に変換
 * Date型でも文字列でも対応
 * @param {Date|string} datetime
 * @returns {number}
 */
function toMillis(datetime) {
  if (!datetime) return 0;
  return new Date(datetime).getTime();
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
    出図実績日: formatDate(job['出図実績日']),
    '出荷予定日変更後': formatDate(job['出荷予定日変更後']),
    '出図予定日変更後': formatDate(job['出図予定日変更後'])
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
      case '納入先住所': return payload['納入先住所'] || '';
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

  // 競合検知（updatedAt列がある場合のみ）- ミリ秒で比較
  if (updatedAtIndex !== -1 && expectedUpdatedAt) {
    const currentMs = toMillis(data[targetRowIndex][updatedAtIndex]);
    const expectedMs = toMillis(expectedUpdatedAt);
    if (currentMs && expectedMs && currentMs !== expectedMs) {
      const error = new Error('他のユーザーが更新しました。最新データを取得してください。');
      error.code = 409;
      throw error;
    }
  }

  // 更新
  const now = new Date();
  const currentUser = Session.getActiveUser().getEmail() || 'system';

  // patchに含まれるキーで、headersに存在しないものは新しいカラムとして追加
  const newColumns = [];
  Object.keys(patch).forEach(key => {
    if (!headers.includes(key)) {
      newColumns.push(key);
    }
  });

  // 新しいカラムをヘッダーに追加
  if (newColumns.length > 0) {
    const lastCol = headers.length;
    newColumns.forEach((colName, idx) => {
      const newColIndex = lastCol + idx + 1;
      sheet.getRange(1, newColIndex).setValue(colName);
      headers.push(colName); // ローカルのheadersも更新
      console.log('[updateJob] 新しいカラムを追加:', colName, 'at column', newColIndex);
    });
  }

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
  // 変更後日付もフォーマット
  if (result['出荷予定日変更後']) {
    result['出荷予定日変更後'] = formatDate(result['出荷予定日変更後']);
  }
  if (result['出図予定日変更後']) {
    result['出図予定日変更後'] = formatDate(result['出図予定日変更後']);
  }

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

/**
 * 新規工程を作成
 * @param {Object} payload - { processName, type, order, colorHex, isWeekly, isActive, memo }
 * @returns {Object} - 作成された工程レコード
 */
function createProcess(payload) {
  const sheet = getSheet(CONFIG.SHEETS.PROCESS_MASTER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const processes = sheetDataToObjects(data);

  // 工程名の重複チェック
  const trimmedName = (payload.processName || '').trim();
  if (!trimmedName) {
    throw new Error('工程名を入力してください');
  }
  const duplicate = processes.find(p => (p['工程名'] || '').trim() === trimmedName);
  if (duplicate) {
    throw new Error('同じ工程名が既に存在します');
  }

  // processId自動生成（P + 3桁数字）
  const existingIds = processes
    .map(p => p.processId)
    .filter(id => /^P\d{3}$/.test(id))
    .map(id => parseInt(id.slice(1), 10));
  const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 0;
  let nextId = maxId + 10;
  // 既存に同IDがあれば+10で空きを探す
  while (processes.some(p => p.processId === `P${String(nextId).padStart(3, '0')}`)) {
    nextId += 10;
  }
  const processId = `P${String(nextId).padStart(3, '0')}`;

  // デフォルト値設定
  const newProcess = {
    'processId': processId,
    '工程名': trimmedName,
    '種別(type)': payload.type || 'range',
    '表示順(order)': payload.order || (maxId + 10),
    '工程色(colorHex)': payload.colorHex || '#4A90E2',
    '主担当必須': 'TRUE',
    '副担当許可': 'TRUE',
    '副担当必須': 'FALSE',
    '有効(isActive)': payload.isActive !== false ? 'TRUE' : 'FALSE',
    '標準(isStandard)': 'TRUE',
    '重要（isWeekly）': payload.isWeekly ? 'TRUE' : 'FALSE',
    '備考': payload.memo || ''
  };

  // ヘッダー順に行データを作成
  const rowData = headers.map(header => newProcess[header] ?? '');
  sheet.appendRow(rowData);

  // boolean正規化して返却
  return {
    ...newProcess,
    isWeekly: payload.isWeekly === true,
    '有効(isActive)': payload.isActive !== false
  };
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

  // --- 車両予約同期 ---
  let vehicleSyncStatus = '';
  let vehicleReservationId = '';
  if (payload.vehicleId) {
    try {
      const syncResult = syncTripToVehicleReservation(tripId, {
        ...payload,
        tripId: tripId,
        start: formatDate(payload.start),
        end: formatDate(payload.end)
      }, 'create');

      const vehResIdIdx = headers.indexOf('vehicleReservationId');
      const vehSyncIdx = headers.indexOf('vehicleSyncStatus');
      const vehErrIdx = headers.indexOf('vehicleSyncError');
      const lastRow = sheet.getLastRow();

      if (syncResult.ok) {
        vehicleReservationId = syncResult.reservationId || '';
        vehicleSyncStatus = 'synced';
        if (vehResIdIdx !== -1) sheet.getRange(lastRow, vehResIdIdx + 1).setValue(syncResult.reservationId);
        if (vehSyncIdx !== -1) sheet.getRange(lastRow, vehSyncIdx + 1).setValue('synced');
        if (vehErrIdx !== -1) sheet.getRange(lastRow, vehErrIdx + 1).setValue('');
      } else {
        vehicleSyncStatus = syncResult.status || 'error';
        if (vehSyncIdx !== -1) sheet.getRange(lastRow, vehSyncIdx + 1).setValue(syncResult.status || 'error');
        if (vehErrIdx !== -1) sheet.getRange(lastRow, vehErrIdx + 1).setValue(syncResult.error || '');
      }
    } catch (e) {
      Logger.log('Vehicle sync error in createTrip: ' + e.message);
      vehicleSyncStatus = 'error';
      const vehSyncIdx = headers.indexOf('vehicleSyncStatus');
      const vehErrIdx = headers.indexOf('vehicleSyncError');
      if (vehSyncIdx !== -1) sheet.getRange(sheet.getLastRow(), vehSyncIdx + 1).setValue('error');
      if (vehErrIdx !== -1) sheet.getRange(sheet.getLastRow(), vehErrIdx + 1).setValue(e.message);
    }
  }

  // --- 関連移動予定への車両自動適用 ---
  if (payload.vehicleId && payload.kind !== 'move') {
    try {
      propagateVehicleToRelatedMoves(tripId, {
        personId: payload.personId,
        start: formatDate(payload.start),
        end: formatDate(payload.end),
        vehicleId: payload.vehicleId,
        kind: payload.kind || ''
      });
    } catch (e) {
      Logger.log('propagateVehicleToRelatedMoves error: ' + e.message);
    }
  }

  return {
    tripId,
    ...payload,
    genKey,
    start: formatDate(payload.start),
    end: formatDate(payload.end),
    updatedAt: formatDateTime(now),
    vehicleId: payload.vehicleId || '',
    vehicleReservationId: vehicleReservationId,
    vehicleSyncStatus: vehicleSyncStatus
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

  // --- 車両予約同期（更新）---
  const needsVehicleSync = (patch.hasOwnProperty('vehicleId') ||
                            patch.hasOwnProperty('start') ||
                            patch.hasOwnProperty('end') ||
                            patch.hasOwnProperty('personId')) &&
                           (result.vehicleId || patch.vehicleId);
  if (needsVehicleSync) {
    try {
      const syncResult = syncTripToVehicleReservation(tripId, result, 'update');

      const vehResIdIdx = headers.indexOf('vehicleReservationId');
      const vehSyncIdx = headers.indexOf('vehicleSyncStatus');
      const vehErrIdx = headers.indexOf('vehicleSyncError');
      const rowNum = targetRowIndex + 1;

      if (syncResult.ok) {
        result.vehicleReservationId = syncResult.reservationId || result.vehicleReservationId || '';
        result.vehicleSyncStatus = 'synced';
        if (vehResIdIdx !== -1) sheet.getRange(rowNum, vehResIdIdx + 1).setValue(syncResult.reservationId);
        if (vehSyncIdx !== -1) sheet.getRange(rowNum, vehSyncIdx + 1).setValue('synced');
        if (vehErrIdx !== -1) sheet.getRange(rowNum, vehErrIdx + 1).setValue('');
      } else {
        result.vehicleSyncStatus = syncResult.status || 'error';
        if (vehSyncIdx !== -1) sheet.getRange(rowNum, vehSyncIdx + 1).setValue(syncResult.status || 'error');
        if (vehErrIdx !== -1) sheet.getRange(rowNum, vehErrIdx + 1).setValue(syncResult.error || '');
      }
    } catch (e) {
      Logger.log('Vehicle sync error in updateTrip: ' + e.message);
      result.vehicleSyncStatus = 'error';
      const vehSyncIdx = headers.indexOf('vehicleSyncStatus');
      const vehErrIdx = headers.indexOf('vehicleSyncError');
      const rowNum = targetRowIndex + 1;
      if (vehSyncIdx !== -1) sheet.getRange(rowNum, vehSyncIdx + 1).setValue('error');
      if (vehErrIdx !== -1) sheet.getRange(rowNum, vehErrIdx + 1).setValue(e.message);
    }
  }

  // --- 関連移動予定への車両自動適用（更新時） ---
  if (result.vehicleId && result.kind !== 'move' && patch.hasOwnProperty('vehicleId')) {
    try {
      propagateVehicleToRelatedMoves(tripId, {
        personId: result.personId,
        start: result.start,
        end: result.end,
        vehicleId: result.vehicleId,
        kind: result.kind || ''
      });
    } catch (e) {
      Logger.log('propagateVehicleToRelatedMoves error in updateTrip: ' + e.message);
    }
  }

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

  // --- 車両予約同期（削除）---
  const vehResIdIdx = headers.indexOf('vehicleReservationId');
  const vehicleReservationId = (vehResIdIdx !== -1) ? data[targetRowIndex][vehResIdIdx] : '';

  // 物理削除
  sheet.deleteRow(targetRowIndex + 1);

  // 車両予約をキャンセル
  if (vehicleReservationId) {
    try {
      syncTripToVehicleReservation(tripId, {}, 'delete');
    } catch (e) {
      Logger.log('Vehicle sync error in deleteTrip: ' + e.message);
    }
  }

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
// TopMemo（上部メモ）- scope=global固定／複数メモ対応
// ============================================

/**
 * TopMemo用ヘッダーを正規化（日本語説明文を除去）
 * 例: "memoId（または固定で自動生成）" → "memoId"
 * @param {string} header - 元のヘッダー名
 * @returns {string} - 正規化されたヘッダー名
 */
function normalizeTopMemoHeader(header) {
  if (!header || typeof header !== 'string') return header;
  // 括弧以降を削除、または最初の英単語を抽出
  const match = header.match(/^([a-zA-Z]+)/);
  return match ? match[1] : header;
}

/**
 * TopMemoシートからヘッダー行のインデックスを検出
 * "memoId"で始まる行を探す
 * @param {Array[]} data - シートデータ
 * @returns {number} - ヘッダー行インデックス（0始まり）、見つからない場合は-1
 */
function findTopMemoHeaderRowIndex(data) {
  for (let i = 0; i < data.length; i++) {
    const firstCell = normalizeTopMemoHeader(data[i][0]);
    if (firstCell === 'memoId') {
      return i;
    }
  }
  return -1;
}

/**
 * TopMemoシートデータをオブジェクト配列に変換（ヘッダー行自動検出対応）
 * @param {Array[]} data - シートデータ
 * @returns {{ headerRowIndex: number, objects: Object[] }}
 */
function topMemoSheetDataToObjects(data) {
  if (!data || data.length < 2) return { headerRowIndex: -1, objects: [] };

  // ヘッダー行を自動検出
  const headerRowIndex = findTopMemoHeaderRowIndex(data);
  if (headerRowIndex === -1) {
    Logger.log('TopMemo: ヘッダー行が見つかりません');
    return { headerRowIndex: -1, objects: [] };
  }

  const rawHeaders = data[headerRowIndex];
  const headers = rawHeaders.map(h => normalizeTopMemoHeader(h));

  // ヘッダー行の次の行からデータ開始（説明行をスキップ）
  // 説明行は英字で始まらない行なのでスキップ
  let dataStartIndex = headerRowIndex + 1;

  // 説明行（日本語で始まる行）をスキップ
  while (dataStartIndex < data.length) {
    const firstCell = data[dataStartIndex][0];
    // memoIdカラムの値がUUIDっぽいか、または空でないデータ行かを判定
    if (firstCell && typeof firstCell === 'string' &&
        (firstCell.match(/^[a-f0-9-]{36}$/i) || // UUID
         firstCell.match(/^[A-Z0-9]+$/) ||       // 英数字ID（例: TOP）
         !firstCell.match(/^[ぁ-んァ-ン一-龥]/))) { // 日本語で始まらない
      break;
    }
    dataStartIndex++;
  }

  const objects = data.slice(dataStartIndex).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  }).filter(obj => obj.memoId); // memoIdが空の行は除外

  return { headerRowIndex, dataStartIndex, objects };
}

/**
 * TopMemoを取得（scope=globalに正規化）
 * @param {string} scope - 入力されてもglobalに強制
 * @returns {Object[]}
 */
function getTopMemos(scope = 'global') {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TOP_MEMO);
    const data = sheet.getDataRange().getValues();
    // ヘッダー行自動検出対応
    const { objects } = topMemoSheetDataToObjects(data);
    let memos = objects;

    // boolean変換・正規化
    memos = memos.map(m => ({
      ...m,
      scope: 'global', // 常にglobalに正規化
      isActive: m.isActive === true || m.isActive === 'TRUE' || m.isActive === 'true',
      isOpenDefault: m.isOpenDefault === true || m.isOpenDefault === 'TRUE' || m.isOpenDefault === 'true',
      sortOrder: Number(m.sortOrder) || 0,
      updatedAt: formatDateTime(m.updatedAt)
    }));

    // isActive=true のものをフィルタ（scopeは全て'global'なのでフィルタ不要）
    memos = memos.filter(m => m.isActive === true);

    // sortOrder昇順でソート
    memos.sort((a, b) => a.sortOrder - b.sortOrder);

    return memos;
  } catch (e) {
    Logger.log('TopMemo取得エラー: ' + e.message);
    return [];
  }
}

/**
 * TopMemo用ヘッダーからカラムインデックスを取得（正規化対応）
 * @param {string[]} headers - 元のヘッダー配列
 * @param {string} fieldName - 検索するフィールド名
 * @returns {number} - カラムインデックス（-1: 見つからない）
 */
function findTopMemoColumnIndex(headers, fieldName) {
  for (let i = 0; i < headers.length; i++) {
    const normalized = normalizeTopMemoHeader(headers[i]);
    if (normalized === fieldName) {
      return i;
    }
  }
  return -1;
}

/**
 * TopMemoを作成または更新（Upsert）
 * @param {Object} payload - { memoId?, title, body, isActive?, sortOrder? }
 * @returns {Object}
 */
function upsertTopMemo(payload) {
  const sheet = getSheet(CONFIG.SHEETS.TOP_MEMO);
  const data = sheet.getDataRange().getValues();

  // ヘッダー行を自動検出
  const headerRowIndex = findTopMemoHeaderRowIndex(data);
  if (headerRowIndex === -1) {
    throw new Error('TopMemoシートのヘッダー行が見つかりません');
  }
  const headers = data[headerRowIndex];

  const now = new Date();
  const currentUser = Session.getActiveUser().getEmail() || 'system';

  // memoIdがない場合は新規作成
  if (!payload.memoId) {
    const memoId = Utilities.getUuid();

    // デフォルトのsortOrderを算出（既存最大+10）
    let maxOrder = 0;
    const sortOrderCol = findTopMemoColumnIndex(headers, 'sortOrder');
    if (sortOrderCol !== -1) {
      for (let i = headerRowIndex + 1; i < data.length; i++) {
        const order = Number(data[i][sortOrderCol]) || 0;
        if (order > maxOrder) maxOrder = order;
      }
    }
    const sortOrder = payload.sortOrder !== undefined ? payload.sortOrder : maxOrder + 10;

    // 新規行データ作成（正規化されたヘッダーでマッチング）
    const newRow = headers.map(header => {
      const normalizedHeader = normalizeTopMemoHeader(header);
      switch (normalizedHeader) {
        case 'memoId': return memoId;
        case 'scope': return 'global'; // 常にglobal
        case 'title': return payload.title || '';
        case 'body': return payload.body || '';
        case 'isOpenDefault': return payload.isOpenDefault !== undefined ? payload.isOpenDefault : false;
        case 'isActive': return payload.isActive !== undefined ? payload.isActive : true;
        case 'sortOrder': return sortOrder;
        case 'updatedAt': return now;
        case 'updatedBy': return currentUser;
        default: return '';
      }
    });

    sheet.appendRow(newRow);

    return {
      memoId,
      scope: 'global',
      title: payload.title || '',
      body: payload.body || '',
      isOpenDefault: payload.isOpenDefault || false,
      isActive: payload.isActive !== undefined ? payload.isActive : true,
      sortOrder,
      updatedAt: formatDateTime(now),
      updatedBy: currentUser
    };
  }

  // 既存メモの更新
  const memoIdCol = findTopMemoColumnIndex(headers, 'memoId');
  if (memoIdCol === -1) {
    throw new Error('memoIdカラムが見つかりません');
  }

  let rowIndex = -1;
  for (let i = headerRowIndex + 1; i < data.length; i++) {
    if (data[i][memoIdCol] === payload.memoId) {
      rowIndex = i + 1; // シートは1始まり
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error(`メモが見つかりません: ${payload.memoId}`);
  }

  // 更新対象カラムを書き込み
  const updatableFields = ['title', 'body', 'isOpenDefault', 'isActive', 'sortOrder'];

  updatableFields.forEach(field => {
    if (payload.hasOwnProperty(field)) {
      const colIndex = findTopMemoColumnIndex(headers, field);
      if (colIndex !== -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(payload[field]);
      }
    }
  });

  // updatedAt, updatedBy を更新
  const updatedAtCol = findTopMemoColumnIndex(headers, 'updatedAt');
  const updatedByCol = findTopMemoColumnIndex(headers, 'updatedBy');
  if (updatedAtCol !== -1) {
    sheet.getRange(rowIndex, updatedAtCol + 1).setValue(now);
  }
  if (updatedByCol !== -1) {
    sheet.getRange(rowIndex, updatedByCol + 1).setValue(currentUser);
  }

  // 更新後のデータを返す
  return {
    memoId: payload.memoId,
    scope: 'global',
    title: payload.title,
    body: payload.body,
    isOpenDefault: payload.isOpenDefault,
    isActive: payload.isActive,
    sortOrder: payload.sortOrder,
    updatedAt: formatDateTime(now),
    updatedBy: currentUser
  };
}

/**
 * TopMemoの並び順を一括更新
 * @param {Array} orderPayload - [{ memoId, sortOrder }, ...]
 * @returns {Object}
 */
function reorderTopMemos(orderPayload) {
  const sheet = getSheet(CONFIG.SHEETS.TOP_MEMO);
  const data = sheet.getDataRange().getValues();

  // ヘッダー行を自動検出
  const headerRowIndex = findTopMemoHeaderRowIndex(data);
  if (headerRowIndex === -1) {
    throw new Error('TopMemoシートのヘッダー行が見つかりません');
  }
  const headers = data[headerRowIndex];

  const memoIdCol = findTopMemoColumnIndex(headers, 'memoId');
  const sortOrderCol = findTopMemoColumnIndex(headers, 'sortOrder');

  if (memoIdCol === -1 || sortOrderCol === -1) {
    throw new Error('必要なカラムが見つかりません');
  }

  // 各メモのsortOrderを更新
  orderPayload.forEach(item => {
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      if (data[i][memoIdCol] === item.memoId) {
        sheet.getRange(i + 1, sortOrderCol + 1).setValue(item.sortOrder);
        break;
      }
    }
  });

  return { success: true, updated: orderPayload.length };
}

/**
 * TopMemoのisActiveを更新（論理削除/復活）
 * @param {string} memoId
 * @param {boolean} isActive
 * @returns {Object}
 */
function setTopMemoActive(memoId, isActive) {
  const sheet = getSheet(CONFIG.SHEETS.TOP_MEMO);
  const data = sheet.getDataRange().getValues();

  // ヘッダー行を自動検出
  const headerRowIndex = findTopMemoHeaderRowIndex(data);
  if (headerRowIndex === -1) {
    throw new Error('TopMemoシートのヘッダー行が見つかりません');
  }
  const headers = data[headerRowIndex];

  const memoIdCol = findTopMemoColumnIndex(headers, 'memoId');
  const isActiveCol = findTopMemoColumnIndex(headers, 'isActive');

  if (memoIdCol === -1 || isActiveCol === -1) {
    throw new Error('必要なカラムが見つかりません');
  }

  let rowIndex = -1;
  for (let i = headerRowIndex + 1; i < data.length; i++) {
    if (data[i][memoIdCol] === memoId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error(`メモが見つかりません: ${memoId}`);
  }

  sheet.getRange(rowIndex, isActiveCol + 1).setValue(isActive);

  // updatedAt, updatedBy も更新
  const now = new Date();
  const currentUser = Session.getActiveUser().getEmail() || 'system';
  const updatedAtCol = findTopMemoColumnIndex(headers, 'updatedAt');
  const updatedByCol = findTopMemoColumnIndex(headers, 'updatedBy');
  if (updatedAtCol !== -1) {
    sheet.getRange(rowIndex, updatedAtCol + 1).setValue(now);
  }
  if (updatedByCol !== -1) {
    sheet.getRange(rowIndex, updatedByCol + 1).setValue(currentUser);
  }

  return { memoId, isActive, updatedAt: formatDateTime(now), updatedBy: currentUser };
}

// ============================================
// DaySettings（日付設定：出勤日/休日）
// ============================================

/**
 * 04_DaySettingsシートを確保（なければ作成）
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureDaySettingsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DAY_SETTINGS);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DAY_SETTINGS);
    sheet.getRange(1, 1, 1, 5).setValues([['日付', '種別', 'メモ', '更新日時', '更新者']]);
    sheet.setFrozenRows(1);
    Logger.log('Created new sheet: ' + CONFIG.SHEETS.DAY_SETTINGS);
  }

  return sheet;
}

/**
 * 指定期間の日付設定を取得
 * @param {string} fromISO - 開始日 (YYYY-MM-DD)
 * @param {string} toISO - 終了日 (YYYY-MM-DD)
 * @returns {Object} { 'YYYY-MM-DD': 'WORKDAY'|'HOLIDAY' }
 */
function getDaySettingsMap(fromISO, toISO) {
  const sheet = ensureDaySettingsSheet();
  const lastRow = sheet.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A:date, B:type

  for (const [dateVal, type] of values) {
    if (!dateVal) continue;

    let key;
    if (dateVal instanceof Date) {
      // Date型の場合はタイムゾーンを考慮して変換
      key = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else {
      // 文字列の場合はそのまま使用
      key = String(dateVal).trim();
    }

    const typeStr = String(type || '').trim();
    // YYYY-MM-DD形式かつ有効な種別のみ追加
    if (typeStr && key.match(/^\d{4}-\d{2}-\d{2}$/)) {
      // 期間フィルタ（指定があれば）
      if (fromISO && toISO) {
        if (key >= fromISO && key <= toISO) {
          map[key] = typeStr;  // 'WORKDAY' or 'HOLIDAY'（文字列のみ）
        }
      } else {
        map[key] = typeStr;
      }
    }
  }

  return map;
}

/**
 * 日付設定を追加/更新/削除
 * @param {string} dateISO - YYYY-MM-DD
 * @param {string|null} type - 'WORKDAY', 'HOLIDAY', または null（解除）
 * @param {string} memo - メモ（任意）
 * @returns {Object} { ok: boolean, action?: string, error?: string }
 */
function setDaySetting(dateISO, type, memo) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, error: 'LOCK_TIMEOUT' };
  }

  try {
    const currentUser = Session.getActiveUser().getEmail() || 'anonymous';
    const now = new Date();

    const sheet = ensureDaySettingsSheet();
    const lastRow = sheet.getLastRow();

    // 既存行を探索
    let existingRow = null;
    if (lastRow >= 2) {
      const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        let key;
        if (dates[i][0] instanceof Date) {
          key = Utilities.formatDate(dates[i][0], 'Asia/Tokyo', 'yyyy-MM-dd');
        } else {
          key = String(dates[i][0]).trim();
        }
        if (key === dateISO) {
          existingRow = 2 + i;
          break;
        }
      }
    }

    if (type === null || type === '') {
      // 解除: 行を削除
      if (existingRow) {
        sheet.deleteRow(existingRow);
        Logger.log('DaySetting removed: ' + dateISO);
      }
      return { ok: true, action: 'removed' };
    }

    const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    if (existingRow) {
      // 更新
      sheet.getRange(existingRow, 2).setValue(type);
      sheet.getRange(existingRow, 3).setValue(memo || '');
      sheet.getRange(existingRow, 4).setValue(timestamp);
      sheet.getRange(existingRow, 5).setValue(currentUser);
      Logger.log('DaySetting updated: ' + dateISO + ' -> ' + type);
      return { ok: true, action: 'updated' };
    }

    // 新規追加
    sheet.appendRow([
      dateISO,
      type,
      memo || '',
      timestamp,
      currentUser
    ]);
    Logger.log('DaySetting created: ' + dateISO + ' -> ' + type);

    return { ok: true, action: 'created' };
  } catch (e) {
    Logger.log('setDaySetting error: ' + e.message);
    return { ok: false, error: e.message || 'SAVE_FAILED' };
  } finally {
    lock.releaseLock();
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

  // TopMemoは取得失敗しても続行（scope=global固定）
  let topMemos = [];
  try {
    topMemos = getTopMemos('global');
  } catch (e) {
    Logger.log('TopMemo取得をスキップ: ' + e.message);
  }

  // 日付設定を取得（失敗しても続行）
  let daySettings = {};
  try {
    daySettings = getDaySettingsMap(start, end);
  } catch (e) {
    Logger.log('DaySettings取得をスキップ: ' + e.message);
  }

  // 車両マスタはBootstrapDataには含めない（GAS同時実行数制限対策）
  // TripModal初回オープン時にapi_getVehicleMaster()で遅延ロードする

  return {
    jobs: getAllJobs(),
    processes: getAllProcesses(),
    people: getAllPeople(),
    schedules: getSchedules(start, end),
    attachments: [], // 初期は空、必要時に取得
    trips: getTrips(start, end),
    jobMaster: jobMaster,
    workerJobAssign: getAllWorkerJobAssigns(),
    topMemos: topMemos,
    daySettings: daySettings,
    meta: {
      rangeStart: start,
      rangeEnd: end,
      days: days,
      fetchedAt: formatDateTime(new Date())
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

    // デバッグ: 実際のシート名一覧を出力
    const allSheets = externalSs.getSheets();
    const sheetNames = allSheets.map(s => s.getName());
    Logger.log('外部スプレッドシートのシート一覧: ' + JSON.stringify(sheetNames));
    Logger.log('検索対象シート名: "' + CONFIG.EXTERNAL_MASTER_SHEET_NAME + '"');

    const externalSheet = externalSs.getSheetByName(CONFIG.EXTERNAL_MASTER_SHEET_NAME);

    if (!externalSheet) {
      throw new Error(`外部シート "${CONFIG.EXTERNAL_MASTER_SHEET_NAME}" が見つかりません。存在するシート: ${sheetNames.join(', ')}`);
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
      syncedAt: formatDateTime(new Date())
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
 * 検索対象: 工番 / 受注先 / 品名
 * @param {string} query - 検索クエリ（部分一致）
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
    // 外部シートの列名に合わせる（工番/受注先/品名で検索）
    const jobNo = String(item['工番'] || '').toLowerCase();
    const customer = String(item['受注先'] || '').toLowerCase();
    const product = String(item['品名'] || '').toLowerCase();

    return jobNo.includes(lowerQuery) ||
           customer.includes(lowerQuery) ||
           product.includes(lowerQuery);
  });

  return results.slice(0, limit);
}

// ============================================
// Batch Save API（一括保存）
// ============================================

/**
 * 複数の変更を一括保存
 * @param {Object} payload - { changes: [...], clientRevision: string, user: string }
 * @returns {Object} - { ok: boolean, results: { schedule: {...}, trip: {...}, job: {...}, topMemo: {...} } }
 */
function api_saveBatch(payload) {
  requireEditor();
  const { changes, clientRevision, user } = payload;

  if (!changes || !Array.isArray(changes) || changes.length === 0) {
    return { ok: true, results: { schedule: {}, trip: {}, job: {}, topMemo: {}, jobProcessLayout: {}, person: {}, process: {} } };
  }

  const results = {
    schedule: { upserted: [], deleted: [] },
    trip: { upserted: [], deleted: [], locked: [], unlocked: [] },
    job: { upserted: [] },
    topMemo: { upserted: [], deleted: [] },
    jobProcessLayout: { reordered: [] },
    person: { upserted: [], deleted: [] },
    process: { upserted: [] }
  };

  const errors = [];

  // temp_*で始まるIDは新規作成として扱う
  const isNewRecord = (id) => !id || id.startsWith('temp_');

  // 各変更を処理
  for (const change of changes) {
    try {
      const { entityType, op, id, payload: changePayload } = change;

      switch (entityType) {
        case 'schedule':
          if (op === 'upsert') {
            let saved;
            if (isNewRecord(id)) {
              // 新規作成
              saved = createSchedule(changePayload);
              saved._tempId = id; // 仮IDを返却（クライアント側でマッピング用）
            } else {
              // 既存レコード更新
              saved = updateSchedule(id, changePayload);
            }
            results.schedule.upserted.push(saved);
          } else if (op === 'delete' && !isNewRecord(id)) {
            deleteSchedule(id);
            results.schedule.deleted.push(id);
          }
          break;

        case 'trip':
          if (op === 'upsert') {
            let saved;
            if (isNewRecord(id)) {
              // 新規作成
              saved = createTrip(changePayload);
              saved._tempId = id;
            } else {
              // 既存レコード更新
              saved = updateTrip(id, changePayload);
            }
            results.trip.upserted.push(saved);
          } else if (op === 'delete' && !isNewRecord(id)) {
            deleteTrip(id);
            results.trip.deleted.push(id);
          } else if (op === 'lock' && !isNewRecord(id)) {
            const locked = lockTrip(id, true);
            results.trip.locked.push(locked);
          } else if (op === 'unlock' && !isNewRecord(id)) {
            const unlocked = lockTrip(id, false);
            results.trip.unlocked.push(unlocked);
          }
          break;

        case 'job':
          if (op === 'upsert') {
            let saved;
            if (isNewRecord(id)) {
              // 新規作成
              saved = createJob(changePayload);
              saved._tempId = id;
            } else {
              // 既存レコード更新
              saved = updateJob(id, changePayload);
            }
            results.job.upserted.push(saved);
          } else if (op === 'close' && !isNewRecord(id)) {
            // 工番を完了（非表示）
            const closed = closeJob(id);
            results.job.closed = results.job.closed || [];
            results.job.closed.push(closed);
          } else if (op === 'reopen' && !isNewRecord(id)) {
            // 工番の完了を解除
            const reopened = reopenJob(id);
            results.job.reopened = results.job.reopened || [];
            results.job.reopened.push(reopened);
          } else if (op === 'delete' && !isNewRecord(id)) {
            // 工番を完全削除（カスケード）
            const deleted = deleteJobCascade(id);
            results.job.deleted = results.job.deleted || [];
            results.job.deleted.push(deleted);
          }
          break;

        case 'topMemo':
          if (op === 'upsert') {
            // TopMemoは upsertTopMemo を使用（memoIdがあれば更新、なければ新規）
            const memoPayload = { ...changePayload };
            if (!isNewRecord(id)) {
              memoPayload.memoId = id;
            }
            let saved = upsertTopMemo(memoPayload);
            if (isNewRecord(id)) {
              saved._tempId = id;
            }
            results.topMemo.upserted.push(saved);
          } else if (op === 'delete' || op === 'setActive') {
            // setActive=false は論理削除
            const isActive = op === 'setActive' ? (changePayload.isActive !== false) : false;
            const updated = setTopMemoActive(id, isActive);
            if (!isActive) {
              results.topMemo.deleted.push(id);
            } else {
              results.topMemo.upserted.push(updated);
            }
          }
          break;

        case 'jobProcessLayout':
          if (op === 'reorder') {
            // 工番別の工程並び順を保存
            // id = jobId, changePayload = { orderedProcessIds: [...] }
            const result = saveJobProcessLayout(id, changePayload.orderedProcessIds);
            results.jobProcessLayout.reordered.push({
              jobId: id,
              layouts: result.layouts
            });
          }
          break;

        case 'person':
          if (op === 'upsert') {
            let saved;
            if (isNewRecord(id)) {
              // 新規作成
              saved = createPerson(changePayload);
              saved._tempId = id;
            } else {
              // 既存レコード更新
              saved = updatePerson(id, changePayload, changePayload._updatedAt);
            }
            results.person.upserted.push(saved);
          } else if (op === 'delete' && !isNewRecord(id)) {
            // 担当者削除（論理削除：isActive = false）
            const deleted = updatePerson(id, { '有効(isActive)': false });
            results.person.deleted.push(deleted);
          }
          break;

        case 'process':
          if (op === 'upsert') {
            let saved;
            if (isNewRecord(id)) {
              // 新規作成
              saved = createProcess(changePayload);
              saved._tempId = id;
            } else {
              // 工程の更新は現在サポートしていない（将来拡張用）
              errors.push({ entityType, op, id, error: 'Process update not supported' });
            }
            if (saved) {
              results.process.upserted.push(saved);
            }
          }
          break;

        default:
          errors.push({ entityType, op, id, error: 'Unknown entityType' });
      }
    } catch (error) {
      errors.push({
        entityType: change.entityType,
        op: change.op,
        id: change.id,
        error: error.message,
        code: error.code
      });
    }
  }

  // エラーがあった場合も部分的な成功を返す
  if (errors.length > 0) {
    return {
      ok: errors.length < changes.length, // 一部成功ならtrue
      results,
      errors,
      partialSuccess: true
    };
  }

  return { ok: true, results };
}

/**
 * 工程のデフォルト並び順を更新
 * ProcessMasterの表示順(order)を一括更新
 * @param {string[]} orderedProcessIds - 新しい順序のprocessId配列
 * @returns {Object} - { success: boolean, error?: string }
 */
function api_updateProcessDefaultOrder(orderedProcessIds) {
  try {
    if (!Array.isArray(orderedProcessIds) || orderedProcessIds.length === 0) {
      return { success: false, error: '無効なデータです' };
    }

    const sheet = getSheet(CONFIG.SHEETS.PROCESS_MASTER);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const processIdCol = headers.indexOf('processId');
    const orderCol = headers.indexOf('表示順(order)');

    if (processIdCol === -1 || orderCol === -1) {
      return { success: false, error: 'シートの構造が不正です' };
    }

    // processId -> 新しいorder値のマップを作成
    const orderMap = {};
    orderedProcessIds.forEach((processId, idx) => {
      orderMap[processId] = (idx + 1) * 10;
    });

    // シートを更新
    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      const processId = data[i][processIdCol];
      if (orderMap.hasOwnProperty(processId)) {
        const newOrder = orderMap[processId];
        if (data[i][orderCol] !== newOrder) {
          sheet.getRange(i + 1, orderCol + 1).setValue(newOrder);
          updated++;
        }
      }
    }

    console.log(`ProcessMaster: ${updated}件の表示順を更新しました`);
    return { success: true, updated: updated };

  } catch (error) {
    console.error('api_updateProcessDefaultOrder error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================
// 工番別工程レイアウト（JobProcessLayout）
// ============================================

/**
 * 工番別の工程表示順を取得
 * @param {string} jobId - 工番ID
 * @returns {Array} - [{layoutId, jobId, processId, orderIndex, isHidden, updatedAt}]
 */
function getJobProcessLayout(jobId) {
  if (!jobId) return [];

  const sheet = getSheet(CONFIG.SHEETS.JOB_PROCESS_LAYOUT);
  if (!sheet) return []; // シートがない場合は空配列

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // ヘッダーのみ

  const headers = data[0];
  const jobIdCol = headers.indexOf('jobId');
  const processIdCol = headers.indexOf('processId');
  const orderIndexCol = headers.indexOf('orderIndex');
  const isHiddenCol = headers.indexOf('isHidden');
  const layoutIdCol = headers.indexOf('layoutId');
  const updatedAtCol = headers.indexOf('updatedAt');

  const layouts = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][jobIdCol] === jobId) {
      layouts.push({
        layoutId: data[i][layoutIdCol] || '',
        jobId: data[i][jobIdCol],
        processId: data[i][processIdCol],
        orderIndex: data[i][orderIndexCol] || 0,
        isHidden: data[i][isHiddenCol] === true,
        updatedAt: data[i][updatedAtCol] || ''
      });
    }
  }

  // orderIndex順にソート
  layouts.sort((a, b) => a.orderIndex - b.orderIndex);
  return layouts;
}

/**
 * 工番別の工程表示順を保存（並び替え）
 * @param {string} jobId - 工番ID
 * @param {Array<string>} orderedProcessIds - 並び順のprocessId配列
 * @returns {Object} - { success: boolean, layouts: Array }
 */
function saveJobProcessLayout(jobId, orderedProcessIds) {
  if (!jobId || !orderedProcessIds || !Array.isArray(orderedProcessIds)) {
    throw new Error('Invalid parameters for saveJobProcessLayout');
  }

  const sheet = getOrCreateJobProcessLayoutSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // カラムインデックス取得
  const layoutIdCol = headers.indexOf('layoutId');
  const jobIdCol = headers.indexOf('jobId');
  const processIdCol = headers.indexOf('processId');
  const orderIndexCol = headers.indexOf('orderIndex');
  const isHiddenCol = headers.indexOf('isHidden');
  const updatedAtCol = headers.indexOf('updatedAt');
  const updatedByCol = headers.indexOf('updatedBy');

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'system';

  // 既存レコードをマップ化（processId -> rowIndex）
  const existingMap = new Map();
  for (let i = 1; i < data.length; i++) {
    if (data[i][jobIdCol] === jobId) {
      existingMap.set(data[i][processIdCol], i);
    }
  }

  // 更新するレコード
  const updatedLayouts = [];

  // orderedProcessIdsを順に処理
  orderedProcessIds.forEach((processId, idx) => {
    const orderIndex = (idx + 1) * 10; // 10, 20, 30...

    if (existingMap.has(processId)) {
      // 既存レコードを更新
      const rowIndex = existingMap.get(processId);
      sheet.getRange(rowIndex + 1, orderIndexCol + 1).setValue(orderIndex);
      sheet.getRange(rowIndex + 1, updatedAtCol + 1).setValue(now);
      sheet.getRange(rowIndex + 1, updatedByCol + 1).setValue(user);

      updatedLayouts.push({
        layoutId: data[rowIndex][layoutIdCol],
        jobId: jobId,
        processId: processId,
        orderIndex: orderIndex,
        isHidden: data[rowIndex][isHiddenCol] === true,
        updatedAt: formatDateTime(now)
      });

      existingMap.delete(processId); // 処理済みとしてマーク
    } else {
      // 新規レコードを追加
      const layoutId = Utilities.getUuid();
      sheet.appendRow([layoutId, jobId, processId, orderIndex, false, now, user]);

      updatedLayouts.push({
        layoutId: layoutId,
        jobId: jobId,
        processId: processId,
        orderIndex: orderIndex,
        isHidden: false,
        updatedAt: formatDateTime(now)
      });
    }
  });

  // 残った既存レコード（orderedProcessIdsに含まれない）はそのまま保持
  // （未使用工程を隠した状態で並べ替えた場合の整合性を保つ）

  return {
    success: true,
    layouts: updatedLayouts
  };
}

/**
 * JobProcessLayoutシートを取得または作成
 */
function getOrCreateJobProcessLayoutSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.JOB_PROCESS_LAYOUT);

  if (!sheet) {
    // シートを新規作成
    sheet = ss.insertSheet(CONFIG.SHEETS.JOB_PROCESS_LAYOUT);
    // ヘッダー行を追加
    sheet.getRange(1, 1, 1, 7).setValues([[
      'layoutId', 'jobId', 'processId', 'orderIndex', 'isHidden', 'updatedAt', 'updatedBy'
    ]]);
    sheet.setFrozenRows(1);
    console.log('Created JobProcessLayout sheet');
  }

  return sheet;
}

// ============================================
// 工程表シート出力
// ============================================

/**
 * 工番別工程表を新規シートにエクスポート（色付き）
 * @param {Object} exportData - エクスポートデータ
 * @returns {Object} - { success: boolean, sheetName: string }
 */
function exportJobDetailToSheet(exportData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const job = exportData.job;
  const processes = exportData.processes;
  const schedules = exportData.schedules;
  const dates = exportData.dates;

  // シート名を生成（工番_製品名_日時）
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MMdd_HHmm');
  const sheetName = `工程表_${job.jobNo}_${timestamp}`.substring(0, 31); // シート名は31文字制限

  // 新規シートを作成
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet); // 既存シートがあれば削除
  }
  sheet = ss.insertSheet(sheetName);

  // ヘッダー情報を書き込み
  sheet.getRange('A1').setValue('工番別工程表');
  sheet.getRange('A1').setFontWeight('bold').setFontSize(14);

  sheet.getRange('A2').setValue(`工番: ${job.jobNo}`);
  sheet.getRange('B2').setValue(`顧客: ${job.customer}`);
  sheet.getRange('C2').setValue(`製品: ${job.product}`);
  sheet.getRange('D2').setValue(`出荷予定: ${job.shipDate}`);
  sheet.getRange('E2').setValue(`出図予定: ${job.drawDate}`);

  // 日付ヘッダー行を作成（4行目）
  const headerRow = 4;
  sheet.getRange(headerRow, 1).setValue('工程');
  sheet.getRange(headerRow, 1).setBackground('#f3f4f6').setFontWeight('bold');

  // 日付を書き込み
  dates.forEach((date, i) => {
    const cell = sheet.getRange(headerRow, i + 2);
    const d = new Date(date);
    const month = d.getMonth() + 1;
    const day = d.getDate();
    const dow = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
    cell.setValue(`${month}/${day}\n${dow}`);
    cell.setHorizontalAlignment('center');
    cell.setVerticalAlignment('middle');
    cell.setFontSize(8);
    cell.setWrap(true);

    // 土日は背景色を変える
    if (d.getDay() === 0) {
      cell.setBackground('#fee2e2'); // 日曜
    } else if (d.getDay() === 6) {
      cell.setBackground('#fef3c7'); // 土曜
    } else {
      cell.setBackground('#f3f4f6');
    }
  });

  // 工程行を書き込み
  processes.forEach((process, pIdx) => {
    const row = headerRow + 1 + pIdx;

    // 工程名セル
    const nameCell = sheet.getRange(row, 1);
    const displayName = process.isMilestone ? `★${process.name}` : process.name;
    nameCell.setValue(displayName);
    nameCell.setBackground(hexToRgbLight(process.color));
    nameCell.setFontWeight('bold');

    // 日付セルを初期化
    dates.forEach((date, dIdx) => {
      const cell = sheet.getRange(row, dIdx + 2);
      const d = new Date(date);
      if (d.getDay() === 0) {
        cell.setBackground('#fef2f2');
      } else if (d.getDay() === 6) {
        cell.setBackground('#fffbeb');
      }
    });

    // スケジュールバーを描画
    const processSchedules = schedules.filter(s => s.processId === process.processId);
    processSchedules.forEach(schedule => {
      const startIdx = dates.indexOf(schedule.start);
      const endIdx = dates.indexOf(schedule.end);

      if (startIdx >= 0 && endIdx >= 0) {
        for (let i = startIdx; i <= endIdx; i++) {
          const cell = sheet.getRange(row, i + 2);
          cell.setBackground(process.color);
          if (i === startIdx && schedule.label) {
            cell.setValue(schedule.label);
            cell.setFontColor('#ffffff');
            cell.setFontSize(8);
          }
        }
      } else if (process.isMilestone && startIdx >= 0) {
        // マイルストーン
        const cell = sheet.getRange(row, startIdx + 2);
        cell.setValue('◆');
        cell.setBackground('#fef3c7');
        cell.setFontColor('#f59e0b');
        cell.setHorizontalAlignment('center');
      }
    });
  });

  // 列幅を調整
  sheet.setColumnWidth(1, 120); // 工程名列
  for (let i = 2; i <= dates.length + 1; i++) {
    sheet.setColumnWidth(i, 35); // 日付列
  }

  // 行高さを調整
  sheet.setRowHeight(headerRow, 40);
  for (let i = headerRow + 1; i <= headerRow + processes.length; i++) {
    sheet.setRowHeight(i, 25);
  }

  // 罫線を設定
  const dataRange = sheet.getRange(headerRow, 1, processes.length + 1, dates.length + 1);
  dataRange.setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);

  return { success: true, sheetName: sheetName };
}

/**
 * 紺一色テーマ用パレット（4段階濃淡）
 */
const NAVY_PALETTE = ['#0B1F4B', '#123A7A', '#1E55B3', '#4B7BD8'];

/**
 * processIdから紺パレットの色を決定（安定的に同じ色になる）
 */
function pickNavyShade(processId) {
  // 簡易ハッシュ: processIdの文字コードを合計して4で割った余り
  let hash = 0;
  const str = String(processId);
  for (let i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash = hash & hash; // Convert to 32bit integer
  }
  const idx = Math.abs(hash) % NAVY_PALETTE.length;
  return NAVY_PALETTE[idx];
}

/**
 * テーマに応じた工程色を取得
 * mono_navy (単色): 全て水色
 */
function resolveProcessColor(theme, process) {
  if (theme === 'mono_navy') {
    return '#4A90D9'; // 水色
  }
  return process.color || '#6B7280';
}

/**
 * 工番別工程表をGoogleスプレッドシートとしてDriveに保存
 * @param {Object} exportData - エクスポートデータ
 * @returns {Object} - { success: boolean, url: string, fileName: string }
 */
function exportJobDetailToSpreadsheet(exportData) {
  const job = exportData.job;
  const processes = exportData.processes;
  const schedules = exportData.schedules;
  const dates = exportData.dates;
  const theme = exportData.theme || 'color'; // 'color' or 'mono_navy'
  const paperSize = exportData.paperSize || 'a4'; // 'a4' or 'a3'

  // 用紙サイズに応じた1段あたりの最大日付列数
  // A4横: 30日、A3横: 45日（列幅22pxで収まるよう調整）
  const MAX_COLS_A4 = 30;
  const MAX_COLS_A3 = 45;
  const maxColsPerBlock = paperSize === 'a3' ? MAX_COLS_A3 : MAX_COLS_A4;

  // 日付を複数ブロックに分割
  const dateBlocks = [];
  for (let i = 0; i < dates.length; i += maxColsPerBlock) {
    dateBlocks.push(dates.slice(i, i + maxColsPerBlock));
  }
  const blockCount = dateBlocks.length;

  // スプレッドシートを作成
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
  const themeLabel = theme === 'mono_navy' ? '_青' : '';
  const fileName = `工番別工程表_${job.jobNo}_${job.product}${themeLabel}_${timestamp}`;
  const ss = SpreadsheetApp.create(fileName);
  const sheet = ss.getActiveSheet();
  sheet.setName('工番別工程表');

  // 各ブロックを縦に積み上げる
  let currentRow = 1;

  for (let blockIdx = 0; blockIdx < blockCount; blockIdx++) {
    const blockDates = dateBlocks[blockIdx];
    const blockStartRow = currentRow;

    // ヘッダー情報（各ブロックの先頭に配置）
    // A1: タイトル、B1: 顧客
    sheet.getRange(currentRow, 1).setValue('工番別工程表').setFontWeight('bold').setFontSize(12);
    sheet.getRange(currentRow, 2).setValue(`顧客: ${job.customer}`).setFontSize(10);
    currentRow++;

    // A2: 工番、B2: 製品
    sheet.getRange(currentRow, 1).setValue(`工番: ${job.jobNo}`).setFontSize(10);
    sheet.getRange(currentRow, 2).setValue(`製品: ${job.product}`).setFontSize(10);
    currentRow++;

    // A3: ブロック番号表示（2段以上の場合）
    if (blockCount > 1) {
      const blockLabel = `（${blockIdx + 1}/${blockCount}）`;
      sheet.getRange(currentRow, 1).setValue(blockLabel).setFontColor('#6B7280').setFontSize(9);
    }
    currentRow++;

    // 日付ヘッダー行
    const headerRow = currentRow;
    sheet.getRange(headerRow, 1).setValue('工程').setBackground('#f3f4f6').setFontWeight('bold').setFontSize(9);

    // 日付を書き込み
    blockDates.forEach((date, i) => {
      const cell = sheet.getRange(headerRow, i + 2);
      const d = new Date(date);
      const month = d.getMonth() + 1;
      const day = d.getDate();
      const dow = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
      cell.setValue(`${month}/${day}\n${dow}`);
      cell.setHorizontalAlignment('center');
      cell.setVerticalAlignment('middle');
      cell.setFontSize(8);
      cell.setWrap(true);

      // 休日背景（両テーマ共通: 土曜=薄黄、日曜=ピンク）
      if (d.getDay() === 0) {
        cell.setBackground('#fee2e2'); // 日曜: ピンク
      } else if (d.getDay() === 6) {
        cell.setBackground('#fef3c7'); // 土曜: 薄黄
      } else {
        cell.setBackground('#f3f4f6');
      }
    });
    currentRow++;

    // 工程行を書き込み
    processes.forEach((process, pIdx) => {
      const row = currentRow + pIdx;
      const nameCell = sheet.getRange(row, 1);
      const displayName = process.isMilestone ? `★${process.name}` : process.name;
      nameCell.setValue(displayName);
      nameCell.setFontWeight('bold');
      nameCell.setFontSize(9);

      // 工程名セルの背景色
      if (theme !== 'mono_navy') {
        const processColor = resolveProcessColor(theme, process);
        nameCell.setBackground(hexToRgbLight(processColor));
      }

      // 休日背景を設定
      blockDates.forEach((date, dIdx) => {
        const cell = sheet.getRange(row, dIdx + 2);
        const d = new Date(date);
        if (d.getDay() === 0) {
          cell.setBackground('#fef2f2'); // 日曜: 薄ピンク
        } else if (d.getDay() === 6) {
          cell.setBackground('#fffbeb'); // 土曜: 薄黄
        }
      });

      // スケジュールバーを描画（このブロックの日付範囲内のみ）
      const processSchedules = schedules.filter(s => s.processId === process.processId);
      processSchedules.forEach(schedule => {
        // ブロック内での相対インデックスを計算
        const startIdx = blockDates.indexOf(schedule.start);
        const endIdx = blockDates.indexOf(schedule.end);

        // スケジュールがこのブロックに含まれる場合
        if (startIdx >= 0 || endIdx >= 0) {
          const barColor = resolveProcessColor(theme, process);

          // ブロック内での開始・終了を計算
          let blockStartIdx = startIdx >= 0 ? startIdx : 0;
          let blockEndIdx = endIdx >= 0 ? endIdx : blockDates.length - 1;

          // 開始日がこのブロックより前の場合
          if (startIdx < 0 && dates.indexOf(schedule.start) < blockIdx * maxColsPerBlock) {
            blockStartIdx = 0;
          }
          // 終了日がこのブロックより後の場合
          if (endIdx < 0 && dates.indexOf(schedule.end) >= (blockIdx + 1) * maxColsPerBlock) {
            blockEndIdx = blockDates.length - 1;
          }
          // スケジュールがこのブロックと重ならない場合はスキップ
          const scheduleStartGlobal = dates.indexOf(schedule.start);
          const scheduleEndGlobal = dates.indexOf(schedule.end);
          const blockStartGlobal = blockIdx * maxColsPerBlock;
          const blockEndGlobal = blockStartGlobal + blockDates.length - 1;

          if (scheduleEndGlobal < blockStartGlobal || scheduleStartGlobal > blockEndGlobal) {
            return; // このブロックには含まれない
          }

          // 実際の描画範囲を計算
          const drawStart = Math.max(0, scheduleStartGlobal - blockStartGlobal);
          const drawEnd = Math.min(blockDates.length - 1, scheduleEndGlobal - blockStartGlobal);

          for (let i = drawStart; i <= drawEnd; i++) {
            const cell = sheet.getRange(row, i + 2);
            cell.setBackground(barColor);
            // ラベルは最初のセルのみ（かつブロック内の開始位置）
            if (i === drawStart && scheduleStartGlobal === blockStartGlobal + drawStart && schedule.label) {
              cell.setValue(schedule.label);
              cell.setFontColor('#ffffff');
              cell.setFontSize(8);
            }
          }
        } else if (process.isMilestone && startIdx >= 0) {
          const cell = sheet.getRange(row, startIdx + 2);
          cell.setValue('◆');
          if (theme === 'mono_navy') {
            cell.setBackground('#E5E7EB');
            cell.setFontColor('#1E3A8A');
          } else {
            cell.setBackground('#fef3c7');
            cell.setFontColor('#f59e0b');
          }
          cell.setHorizontalAlignment('center');
        }
      });
    });

    // 罫線を設定
    const dataRange = sheet.getRange(headerRow, 1, processes.length + 1, blockDates.length + 1);
    dataRange.setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);

    // 行高さ設定（コンパクト化）
    sheet.setRowHeight(headerRow, 35);
    for (let i = 0; i < processes.length; i++) {
      sheet.setRowHeight(headerRow + 1 + i, 22);
    }

    // 次のブロックの開始位置（空行2行分）
    currentRow = headerRow + processes.length + 3;
  }

  // 列幅設定（コンパクト化: 工程列140px、日付列22px）
  sheet.setColumnWidth(1, 140);
  const maxBlockCols = Math.max(...dateBlocks.map(b => b.length));
  for (let i = 2; i <= maxBlockCols + 1; i++) {
    sheet.setColumnWidth(i, 22);
  }

  // スプレッドシートをフラッシュして確定
  SpreadsheetApp.flush();

  const spreadsheetId = ss.getId();
  console.log('[exportJobDetailToSpreadsheet] createdSpreadsheetId:', spreadsheetId);

  // 「工番別工程表」フォルダに移動
  try {
    moveFileToOutputFolder_(spreadsheetId);
    console.log('[exportJobDetailToSpreadsheet] ファイルをフォルダに移動完了');
  } catch (e) {
    // フォルダ移動に失敗してもルートにあるので続行
    console.error('[exportJobDetailToSpreadsheet] フォルダ移動失敗:', e, e.stack);
  }

  return {
    success: true,
    url: ss.getUrl(),
    fileName: fileName,
    spreadsheetId: spreadsheetId,
    blockCount: blockCount // 分割数を返す
  };
}

// 出力先フォルダID（共有ドライブ内の「工番別工程表」フォルダ）
const OUTPUT_FOLDER_ID = '1DH0vLRhrTYwBxctPEPshe9VZtSud8xeg';

/**
 * 作成したスプレッドシート(ファイル)を出力フォルダへ移動
 * Drive API を使用（共有ドライブ対応）
 * @param {string} fileId - ファイルID（SpreadsheetのID）
 */
function moveFileToOutputFolder_(fileId) {
  console.log('[moveFileToOutputFolder_] 開始:', fileId);

  try {
    // 現在の親フォルダを取得
    const file = Drive.Files.get(fileId, { supportsAllDrives: true });
    const parents = (file.parents || []).map(p => p.id);
    const removeParents = parents.join(',');

    console.log('[moveFileToOutputFolder_] 現在の親:', removeParents);

    // 親フォルダを差し替え（MyDrive → 共有ドライブフォルダへ移動）
    Drive.Files.update(
      {},  // 更新するメタデータなし
      fileId,
      null,
      {
        addParents: OUTPUT_FOLDER_ID,
        removeParents: removeParents,
        supportsAllDrives: true
      }
    );

    console.log('[moveFileToOutputFolder_] 移動完了:', OUTPUT_FOLDER_ID);
  } catch (e) {
    console.error('[moveFileToOutputFolder_] 移動失敗:', e);
    throw e;
  }
}

/**
 * HEXカラーを薄い背景色に変換
 * @param {string} hex - HEXカラー (#RRGGBB)
 * @returns {string}
 */
function hexToRgbLight(hex) {
  if (!hex || !hex.startsWith('#')) return '#f3f4f6';

  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);

  // 薄くする（白に近づける）
  const lightR = Math.round(r + (255 - r) * 0.7);
  const lightG = Math.round(g + (255 - g) * 0.7);
  const lightB = Math.round(b + (255 - b) * 0.7);

  return `#${lightR.toString(16).padStart(2, '0')}${lightG.toString(16).padStart(2, '0')}${lightB.toString(16).padStart(2, '0')}`;
}

// ============================================
// 工番 完了（非表示）/ 削除
// ============================================

/**
 * 工番を完了（非表示）にする
 * @param {string} jobId
 * @returns {Object} - { jobId, status, isHidden, closedAt, closedBy }
 */
function closeJob(jobId) {
  const sheet = getSheet(CONFIG.SHEETS.JOBS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const jobIdIndex = headers.indexOf('jobId');
  const statusIndex = headers.indexOf('status');
  const isHiddenIndex = headers.indexOf('isHidden');
  const closedAtIndex = headers.indexOf('closedAt');
  const closedByIndex = headers.indexOf('closedBy');

  if (jobIdIndex === -1) {
    throw new Error('jobId列が見つかりません');
  }

  // 対象行を検索
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

  const now = new Date();
  const currentUser = Session.getActiveUser().getEmail() || 'system';

  // 列が存在する場合のみ更新
  if (statusIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, statusIndex + 1).setValue('closed');
  }
  if (isHiddenIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, isHiddenIndex + 1).setValue(true);
  }
  if (closedAtIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, closedAtIndex + 1).setValue(now);
  }
  if (closedByIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, closedByIndex + 1).setValue(currentUser);
  }

  return {
    jobId,
    status: 'closed',
    isHidden: true,
    closedAt: formatDateTime(now),
    closedBy: currentUser
  };
}

/**
 * 工番の完了を解除（再開）する
 * @param {string} jobId
 * @returns {Object}
 */
function reopenJob(jobId) {
  const sheet = getSheet(CONFIG.SHEETS.JOBS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const jobIdIndex = headers.indexOf('jobId');
  const statusIndex = headers.indexOf('status');
  const isHiddenIndex = headers.indexOf('isHidden');

  if (jobIdIndex === -1) {
    throw new Error('jobId列が見つかりません');
  }

  // 対象行を検索
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

  // 列が存在する場合のみ更新
  if (statusIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, statusIndex + 1).setValue('active');
  }
  if (isHiddenIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, isHiddenIndex + 1).setValue(false);
  }

  return {
    jobId,
    status: 'active',
    isHidden: false
  };
}

/**
 * 工番を完全削除（関連データも削除）
 * @param {string} jobId
 * @returns {Object} - { jobId, deleted: true, deletedCounts: {...} }
 */
function deleteJobCascade(jobId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error('ロックの取得に失敗しました。しばらく待ってから再試行してください。');
  }

  try {
    const deletedCounts = {
      jobs: 0,
      schedules: 0,
      attachments: 0,
      trips: 0,
      jobProcessLayouts: 0
    };

    // 1. Scheduleから削除
    deletedCounts.schedules = deleteRowsByJobId(CONFIG.SHEETS.SCHEDULE, 'jobId', jobId);

    // 2. Attachmentsから削除
    deletedCounts.attachments = deleteRowsByJobId(CONFIG.SHEETS.ATTACHMENTS, 'jobId', jobId);

    // 3. JobProcessLayoutから削除
    try {
      deletedCounts.jobProcessLayouts = deleteRowsByJobId(CONFIG.SHEETS.JOB_PROCESS_LAYOUT, 'jobId', jobId);
    } catch (e) {
      // シートがない場合は無視
      console.log('JobProcessLayout削除スキップ: ' + e.message);
    }

    // 4. Tripsから削除（jobId(任意)列）
    deletedCounts.trips = deleteRowsByJobId(CONFIG.SHEETS.TRIPS, 'jobId(任意)', jobId);

    // 5. Jobsから削除
    deletedCounts.jobs = deleteRowsByJobId(CONFIG.SHEETS.JOBS, 'jobId', jobId);

    return {
      jobId,
      deleted: true,
      deletedCounts
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 指定シートから指定列の値が一致する行を削除
 * @param {string} sheetName
 * @param {string} columnName
 * @param {string} value
 * @returns {number} - 削除した行数
 */
function deleteRowsByJobId(sheetName, columnName, value) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const columnIndex = headers.indexOf(columnName);
    if (columnIndex === -1) {
      return 0; // 列がない場合は0件
    }

    // 削除対象行を逆順で収集（下から削除するため）
    const rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][columnIndex] === value) {
        rowsToDelete.push(i + 1); // 1-indexed
      }
    }

    // 逆順で削除
    rowsToDelete.forEach(rowNum => {
      sheet.deleteRow(rowNum);
    });

    return rowsToDelete.length;
  } catch (e) {
    console.log(`${sheetName}からの削除エラー: ${e.message}`);
    return 0;
  }
}

// ============================================
// PDF ファイル管理（共有ドライブ）
// ============================================

// PDF保存先フォルダID（共有ドライブ内のフォルダ）
const PDF_FOLDER_ID = '1Om_Rq22kyVIrRLi0yyvJD8lrBvPD3uGm';

// PDFタイプの列名マッピング
const PDF_TYPE_COLUMNS = {
  'order': 'orderPdfFileId',        // 受注表
  'instruction': 'instructionPdfFileId'  // 工番別指示書
};

/**
 * 工番用のPDFサブフォルダを取得または作成
 * @param {string} jobNo - 工番
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateJobPdfFolder_(jobNo) {
  console.log('[getOrCreateJobPdfFolder_] 工番:', jobNo);

  const parentFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
  const folderName = jobNo;

  // 既存フォルダを検索
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    const existingFolder = folders.next();
    console.log('[getOrCreateJobPdfFolder_] 既存フォルダ使用:', existingFolder.getId());
    return existingFolder;
  }

  // 新規作成
  const newFolder = parentFolder.createFolder(folderName);
  console.log('[getOrCreateJobPdfFolder_] 新規フォルダ作成:', newFolder.getId());
  return newFolder;
}

/**
 * PDFをアップロード
 * @param {string} jobNo - 工番
 * @param {string} pdfType - PDFタイプ ('order' | 'instruction')
 * @param {string} fileName - ファイル名
 * @param {string} base64Data - Base64エンコードされたPDFデータ
 * @returns {Object} - { fileId, fileName, url }
 */
function uploadJobPdf(jobNo, pdfType, fileName, base64Data) {
  console.log('[uploadJobPdf] 開始:', jobNo, pdfType, fileName);

  // PDFタイプの検証
  if (!PDF_TYPE_COLUMNS[pdfType]) {
    throw new Error('無効なPDFタイプ: ' + pdfType);
  }

  // 工番の存在確認
  const job = findJobByJobNo_(jobNo);
  if (!job) {
    throw new Error('工番が見つかりません: ' + jobNo);
  }

  // 既存ファイルがあれば削除
  const columnName = PDF_TYPE_COLUMNS[pdfType];
  const existingFileId = job[columnName];
  if (existingFileId) {
    try {
      DriveApp.getFileById(existingFileId).setTrashed(true);
      console.log('[uploadJobPdf] 既存ファイル削除:', existingFileId);
    } catch (e) {
      console.log('[uploadJobPdf] 既存ファイル削除スキップ:', e.message);
    }
  }

  // フォルダ取得/作成
  const folder = getOrCreateJobPdfFolder_(jobNo);

  // Base64デコードしてBlobを作成
  const decodedData = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decodedData, 'application/pdf', fileName);

  // ファイル作成
  const file = folder.createFile(blob);
  const fileId = file.getId();
  const url = file.getUrl();

  console.log('[uploadJobPdf] ファイル作成:', fileId, url);

  // Jobsシートにファイル IDを保存
  updateJobPdfFileId_(job.jobId, columnName, fileId);

  return {
    fileId: fileId,
    fileName: fileName,
    url: url
  };
}

/**
 * PDFを削除
 * @param {string} jobNo - 工番
 * @param {string} pdfType - PDFタイプ ('order' | 'instruction')
 * @returns {Object} - { success: true }
 */
function deleteJobPdf(jobNo, pdfType) {
  console.log('[deleteJobPdf] 開始:', jobNo, pdfType);

  // PDFタイプの検証
  if (!PDF_TYPE_COLUMNS[pdfType]) {
    throw new Error('無効なPDFタイプ: ' + pdfType);
  }

  // 工番の存在確認
  const job = findJobByJobNo_(jobNo);
  if (!job) {
    throw new Error('工番が見つかりません: ' + jobNo);
  }

  const columnName = PDF_TYPE_COLUMNS[pdfType];
  const fileId = job[columnName];

  if (!fileId) {
    console.log('[deleteJobPdf] ファイルIDなし、スキップ');
    return { success: true };
  }

  // ファイル削除（ゴミ箱へ移動）
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    console.log('[deleteJobPdf] ファイル削除:', fileId);
  } catch (e) {
    console.log('[deleteJobPdf] ファイル削除エラー:', e.message);
  }

  // JobsシートからファイルIDをクリア
  updateJobPdfFileId_(job.jobId, columnName, '');

  return { success: true };
}

/**
 * 工番のPDF情報を取得
 * @param {string} jobNo - 工番
 * @returns {Object} - { order: { fileId, url } | null, instruction: { fileId, url } | null }
 */
function getJobPdfInfo(jobNo) {
  console.log('[getJobPdfInfo] 工番:', jobNo);

  const job = findJobByJobNo_(jobNo);
  if (!job) {
    return { order: null, instruction: null };
  }

  const result = {
    order: null,
    instruction: null
  };

  // 受注表
  if (job.orderPdfFileId) {
    try {
      const file = DriveApp.getFileById(job.orderPdfFileId);
      result.order = {
        fileId: job.orderPdfFileId,
        fileName: file.getName(),
        url: file.getUrl()
      };
    } catch (e) {
      console.log('[getJobPdfInfo] 受注表ファイル取得エラー:', e.message);
    }
  }

  // 工番別指示書
  if (job.instructionPdfFileId) {
    try {
      const file = DriveApp.getFileById(job.instructionPdfFileId);
      result.instruction = {
        fileId: job.instructionPdfFileId,
        fileName: file.getName(),
        url: file.getUrl()
      };
    } catch (e) {
      console.log('[getJobPdfInfo] 指示書ファイル取得エラー:', e.message);
    }
  }

  return result;
}

/**
 * 工番（工番文字列）でJobを検索
 * @param {string} jobNo - 工番
 * @returns {Object|null}
 */
function findJobByJobNo_(jobNo) {
  const sheet = getSheet(CONFIG.SHEETS.JOBS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const jobNoIndex = headers.indexOf('工番');
  if (jobNoIndex === -1) return null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][jobNoIndex] === jobNo) {
      const job = {};
      headers.forEach((header, index) => {
        job[header] = data[i][index];
      });
      return job;
    }
  }
  return null;
}

/**
 * JobsシートのPDFファイルIDを更新
 * @param {string} jobId - jobId
 * @param {string} columnName - 列名 (orderPdfFileId | instructionPdfFileId)
 * @param {string} fileId - ファイルID（削除時は空文字）
 */
function updateJobPdfFileId_(jobId, columnName, fileId) {
  const sheet = getSheet(CONFIG.SHEETS.JOBS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const jobIdIndex = headers.indexOf('jobId');
  let columnIndex = headers.indexOf(columnName);

  // 列が存在しない場合は追加
  if (columnIndex === -1) {
    const lastCol = headers.length;
    sheet.getRange(1, lastCol + 1).setValue(columnName);
    columnIndex = lastCol;
    console.log('[updateJobPdfFileId_] 列追加:', columnName, '位置:', columnIndex + 1);
  }

  // 対象行を検索して更新
  for (let i = 1; i < data.length; i++) {
    if (data[i][jobIdIndex] === jobId) {
      sheet.getRange(i + 1, columnIndex + 1).setValue(fileId);
      console.log('[updateJobPdfFileId_] 更新完了: row=', i + 1, 'col=', columnIndex + 1, 'value=', fileId);
      return;
    }
  }

  throw new Error('Job更新対象が見つかりません: ' + jobId);
}

