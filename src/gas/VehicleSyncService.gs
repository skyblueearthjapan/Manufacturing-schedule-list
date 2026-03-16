/**
 * VehicleSyncService.gs
 * 車両管理DBとの同期サービス
 * 出張計画(Trips)のCRUDに連動して車両予約を車両管理DBに同期する
 */

// ============================================
// ヘルパー関数
// ============================================

/**
 * 車両予約IDを生成（R + YYYYMMDD-HHmmss + 6桁乱数）
 * @returns {string}
 */
function genVehicleReservationId_() {
  const now = new Date();
  const y = now.getFullYear();
  const mo = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  const h = String(now.getHours()).padStart(2, '0');
  const mi = String(now.getMinutes()).padStart(2, '0');
  const s = String(now.getSeconds()).padStart(2, '0');
  const rand = String(Math.floor(Math.random() * 1000000)).padStart(6, '0');
  return `R${y}${mo}${d}-${h}${mi}${s}${rand}`;
}

/**
 * PeopleシートからpersonIdで氏名と部署を取得
 * @param {string} personId
 * @returns {{name: string, dept: string}}
 */
function getPersonInfo_(personId) {
  const person = getPersonById(personId);
  if (!person) {
    return { name: '', dept: '' };
  }
  const name = person['氏名'] || person.name || '';
  const dept = person['部署'] || person.department || '';
  return { name, dept };
}

/**
 * 車両管理DBの指定シートを開く
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getVehicleDbSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.VEHICLE_DB.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`車両管理DB: シート "${sheetName}" が見つかりません`);
  }
  return sheet;
}

/**
 * 開始日〜終了日の日付配列を生成
 * @param {string|Date} startDate
 * @param {string|Date} endDate
 * @returns {Date[]}
 */
function expandDates_(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);

  const dates = [];
  const current = new Date(start);
  while (current <= end) {
    dates.push(new Date(current));
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

/**
 * 現在日時のISO文字列を返す
 * @returns {string}
 */
function nowIso_() {
  return new Date().toISOString();
}

/**
 * 日付をYYYY-MM-DD形式にフォーマット
 * @param {Date} date
 * @returns {string}
 */
function formatDateForVehicle_(date) {
  if (!date) return '';
  const d = new Date(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/**
 * 車両管理DBのシートデータからヘッダー行を検索してデータを構造化
 * 車両管理DBはタイトル行+説明行の後にヘッダー行がある構造
 * @param {Array[]} data - getDataRange().getValues()の結果
 * @param {string} keyColumn - ヘッダー行を特定するキーカラム名
 * @returns {{headerRowIdx: number, headers: string[], rows: Object[]}}
 */
function parseVehicleDbSheet_(data, keyColumn) {
  let headerRowIdx = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i].indexOf(keyColumn) !== -1) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx === -1) {
    return { headerRowIdx: -1, headers: [], rows: [] };
  }

  const headers = data[headerRowIdx];
  const rows = [];
  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, idx) => {
      if (h) row[h] = data[i][idx];
    });
    rows.push(row);
  }
  return { headerRowIdx, headers, rows };
}

// ============================================
// メイン関数
// ============================================

/**
 * Tripsシートに車両管理用カラムを追加（冪等）
 * vehicleId, vehicleReservationId, vehicleSyncStatus, vehicleSyncError
 */
function ensureTripsVehicleColumns() {
  const sheet = getSheet(CONFIG.SHEETS.TRIPS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const requiredColumns = [
    'vehicleId',
    'vehicleReservationId',
    'vehicleSyncStatus',
    'vehicleSyncError'
  ];

  const missingColumns = requiredColumns.filter(col => headers.indexOf(col) === -1);

  if (missingColumns.length === 0) {
    Logger.log('車両管理カラムは既に存在します');
    return;
  }

  // 既存の最終列の後ろに追加
  const startCol = headers.length + 1;
  missingColumns.forEach((colName, idx) => {
    sheet.getRange(1, startCol + idx).setValue(colName);
  });

  Logger.log(`車両管理カラムを追加しました: ${missingColumns.join(', ')}`);
}

/**
 * 車両マスタ取得API（フロントエンドから遅延ロード用）
 * google.script.run.api_getVehicleMaster() で呼び出される
 * @returns {Object[]}
 */
function api_getVehicleMaster() {
  return getVehicleMaster();
}

/**
 * 車両管理DBのVehiclesシートからactive車両一覧を取得
 * @returns {Object[]} [{vehicle_id, name, category, display_order, plate_no, ui_style, active}]
 */
function getVehicleMaster() {
  try {
    const sheet = getVehicleDbSheet_(CONFIG.VEHICLE_DB.SHEETS.VEHICLES);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const parsed = parseVehicleDbSheet_(data, 'vehicle_id');
    if (parsed.headerRowIdx === -1) {
      Logger.log('車両マスタ: vehicle_idヘッダーが見つかりません');
      return [];
    }

    const vehicles = [];

    for (let i = 0; i < parsed.rows.length; i++) {
      const row = parsed.rows[i];

      // 空行スキップ
      if (!row.vehicle_id) continue;

      // activeな車両のみ
      const active = row.active;
      if (active === true || active === 'TRUE' || active === 'true' || active === 1) {
        vehicles.push({
          vehicle_id: row.vehicle_id || '',
          name: row.name || '',
          category: row.category || '',
          display_order: row.display_order || 0,
          plate_no: row.plate_no || '',
          ui_style: row.ui_style || '',
          active: true
        });
      }
    }

    return vehicles;
  } catch (e) {
    Logger.log('車両マスタ取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 指定期間の車両空き状況を返す
 * @param {string} startDate - YYYY-MM-DD
 * @param {string} endDate - YYYY-MM-DD
 * @param {string} [excludeTripId] - 除外するtripId（更新時の自分自身除外）
 * @returns {Object} { vehicleId: { available: true/false, conflicts: [...] } }
 */
function api_checkVehicleAvailability(startDate, endDate, excludeTripId) {
  try {
    const vehicles = getVehicleMaster();
    const result = {};

    // 全車両を初期化
    vehicles.forEach(v => {
      result[v.vehicle_id] = { available: true, conflicts: [] };
    });

    // ReservationDaysシートから期間内の予約日を取得
    const daysSheet = getVehicleDbSheet_(CONFIG.VEHICLE_DB.SHEETS.RESERVATION_DAYS);
    const daysData = daysSheet.getDataRange().getValues();
    const daysParsed = parseVehicleDbSheet_(daysData, 'reservation_id');
    if (daysParsed.headerRowIdx === -1) return result;

    const daysHeaders = daysParsed.headers;
    const daysResIdIdx = daysHeaders.indexOf('reservation_id');
    const daysVehicleIdx = daysHeaders.indexOf('vehicle_id');
    const daysDateIdx = daysHeaders.indexOf('date');

    // 期間内のReservationDayからreservation_idを集める
    const relevantReservationIds = new Set();
    const allDaysRows = daysData.slice(daysParsed.headerRowIdx + 1);
    for (let i = 0; i < allDaysRows.length; i++) {
      const dayDate = formatDateForVehicle_(allDaysRows[i][daysDateIdx]);
      if (dayDate >= startDate && dayDate <= endDate) {
        relevantReservationIds.add(allDaysRows[i][daysResIdIdx]);
      }
    }

    if (relevantReservationIds.size === 0) return result;

    // Reservationsシートから予約詳細を取得
    const resSheet = getVehicleDbSheet_(CONFIG.VEHICLE_DB.SHEETS.RESERVATIONS);
    const resData = resSheet.getDataRange().getValues();
    const resParsed = parseVehicleDbSheet_(resData, 'reservation_id');
    if (resParsed.headerRowIdx === -1) return result;

    const resHeaders = resParsed.headers;
    const resIdIdx = resHeaders.indexOf('reservation_id');
    const resVehicleIdx = resHeaders.indexOf('vehicle_id');
    const resStatusIdx = resHeaders.indexOf('status');
    const resWorkerIdx = resHeaders.indexOf('worker_name');
    const resDeptIdx = resHeaders.indexOf('dept_name');
    const resStartIdx = resHeaders.indexOf('start_date');
    const resEndIdx = resHeaders.indexOf('end_date');
    const resSourceSystemIdx = resHeaders.indexOf('source_system');
    const resSourceIdIdx = resHeaders.indexOf('source_id');

    const allResRows = resData.slice(resParsed.headerRowIdx + 1);
    for (let i = 0; i < allResRows.length; i++) {
      const resId = allResRows[i][resIdIdx];
      if (!relevantReservationIds.has(resId)) continue;

      const status = allResRows[i][resStatusIdx];
      if (status !== 'active') continue;

      // excludeTripId指定時はsource_id=excludeTripIdの予約を除外
      if (excludeTripId) {
        const sourceId = allResRows[i][resSourceIdIdx];
        if (sourceId === excludeTripId) continue;
      }

      const vehicleId = allResRows[i][resVehicleIdx];
      if (!result[vehicleId]) {
        result[vehicleId] = { available: true, conflicts: [] };
      }

      result[vehicleId].available = false;
      result[vehicleId].conflicts.push({
        worker_name: allResRows[i][resWorkerIdx] || '',
        dept_name: allResRows[i][resDeptIdx] || '',
        start_date: formatDateForVehicle_(allResRows[i][resStartIdx]),
        end_date: formatDateForVehicle_(allResRows[i][resEndIdx])
      });
    }

    return result;
  } catch (e) {
    Logger.log('車両空き状況チェックエラー: ' + e.message);
    return {};
  }
}

/**
 * Trip CRUDに連動して車両管理DBに予約を同期する
 * @param {string} tripId
 * @param {Object} tripData - Tripデータ
 * @param {string} op - 'create' | 'update' | 'delete'
 * @returns {Object} {ok: true/false, reservationId?, status?, error?}
 */
function syncTripToVehicleReservation(tripId, tripData, op) {
  // 車両IDが指定されていない場合（deleteは除く）
  if (op !== 'delete' && (!tripData || !tripData.vehicleId)) {
    return { ok: true, status: 'skipped', message: '車両未指定のためスキップ' };
  }

  let lock = null;
  try {
    lock = LockService.getScriptLock();
    lock.waitLock(30000);

    switch (op) {
      case 'create':
        return syncCreate_(tripId, tripData, lock);
      case 'update':
        return syncUpdate_(tripId, tripData, lock);
      case 'delete':
        return syncDelete_(tripId, lock);
      default:
        return { ok: false, status: 'error', error: '不明な操作: ' + op };
    }
  } catch (e) {
    return { ok: false, status: 'error', error: e.message };
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (ignore) {}
    }
  }
}

/**
 * 新規予約を同期（create）
 * @private
 */
function syncCreate_(tripId, tripData, lock) {
  const resSheet = getVehicleDbSheet_(CONFIG.VEHICLE_DB.SHEETS.RESERVATIONS);
  const daysSheet = getVehicleDbSheet_(CONFIG.VEHICLE_DB.SHEETS.RESERVATION_DAYS);

  // 競合チェック
  const conflictResult = checkConflicts_(resSheet, daysSheet, tripData.vehicleId, tripData.start, tripData.end, null);
  if (conflictResult.hasConflict) {
    if (conflictResult.seisanConflicts.length > 0) {
      // 生産工程表からの予約と競合 → conflict返却
      return {
        ok: false,
        status: 'conflict',
        error: '生産工程表の他の出張計画と車両が競合しています',
        conflicts: conflictResult.seisanConflicts
      };
    }
    // 一般予約との競合 → 一般予約をcancelledに更新
    cancelGeneralReservations_(resSheet, conflictResult.generalConflictRows);
  }

  // 担当者情報を取得
  const personInfo = getPersonInfo_(tripData.personId);
  const now = nowIso_();
  const user = Session.getActiveUser().getEmail() || '';
  const reservationId = genVehicleReservationId_();

  // Reservationsに1行追加
  const resData = resSheet.getDataRange().getValues();
  const resParsed = parseVehicleDbSheet_(resData, 'reservation_id');
  const resHeaders = resParsed.headerRowIdx !== -1 ? resParsed.headers : resData[0];
  const newResRow = resHeaders.map(h => {
    switch (h) {
      case 'reservation_id': return reservationId;
      case 'vehicle_id': return tripData.vehicleId;
      case 'start_date': return tripData.start;
      case 'end_date': return tripData.end;
      case 'slot': return CONFIG.VEHICLE_DB.DEFAULT_SLOT;
      case 'dept_name': return personInfo.dept;
      case 'worker_code': return tripData.personId;
      case 'worker_name': return personInfo.name;
      case 'purpose': return '[出張計画] ' + (tripData['用件'] || '');
      case 'destination': return tripData['行先'] || '';
      case 'memo': return tripData['備考'] || '';
      case 'status': return 'active';
      case 'source_system': return CONFIG.VEHICLE_DB.SOURCE_SYSTEM;
      case 'source_id': return tripId;
      case 'source_url': return '';
      case 'source_last_sync_at': return now;
      case 'created_at': return now;
      case 'created_by': return user;
      case 'updated_at': return now;
      case 'updated_by': return user;
      default: return '';
    }
  });
  resSheet.appendRow(newResRow);

  // ReservationDaysに各日の行を追加
  appendReservationDays_(daysSheet, reservationId, tripData.vehicleId, tripData.start, tripData.end);

  return { ok: true, reservationId: reservationId };
}

/**
 * 予約を更新（update）
 * @private
 */
function syncUpdate_(tripId, tripData, lock) {
  const resSheet = getVehicleDbSheet_(CONFIG.VEHICLE_DB.SHEETS.RESERVATIONS);
  const daysSheet = getVehicleDbSheet_(CONFIG.VEHICLE_DB.SHEETS.RESERVATION_DAYS);

  // 既存予約を検索
  const existing = findReservationBySourceId_(resSheet, tripId);

  if (!existing) {
    // 既存予約なし → vehicleIdがあれば新規作成
    if (tripData.vehicleId) {
      return syncCreate_(tripId, tripData, lock);
    }
    return { ok: true, status: 'skipped', message: '既存予約なし・車両未指定' };
  }

  // vehicleIdが空 → 旧予約をcancelled
  if (!tripData.vehicleId) {
    cancelReservation_(resSheet, existing.rowIndex);
    return { ok: true, status: 'cancelled', message: '車両指定が解除されたため予約をキャンセル' };
  }

  const now = nowIso_();
  const user = Session.getActiveUser().getEmail() || '';
  const resData = resSheet.getDataRange().getValues();
  const resParsed = parseVehicleDbSheet_(resData, 'reservation_id');
  const resHeaders = resParsed.headerRowIdx !== -1 ? resParsed.headers : resData[0];

  // 車両変更チェック
  const vehicleChanged = existing.data.vehicle_id !== tripData.vehicleId;

  if (vehicleChanged) {
    // 旧予約をcancelled → 新予約を作成
    cancelReservation_(resSheet, existing.rowIndex);
    return syncCreate_(tripId, tripData, lock);
  }

  // 日程変更チェック
  const startChanged = formatDateForVehicle_(existing.data.start_date) !== formatDateForVehicle_(tripData.start);
  const endChanged = formatDateForVehicle_(existing.data.end_date) !== formatDateForVehicle_(tripData.end);

  if (startChanged || endChanged) {
    // 競合チェック（自分自身を除外）
    const conflictResult = checkConflicts_(resSheet, daysSheet, tripData.vehicleId, tripData.start, tripData.end, existing.data.reservation_id);
    if (conflictResult.hasConflict && conflictResult.seisanConflicts.length > 0) {
      return {
        ok: false,
        status: 'conflict',
        error: '日程変更先で車両が競合しています',
        conflicts: conflictResult.seisanConflicts
      };
    }
    if (conflictResult.hasConflict) {
      cancelGeneralReservations_(resSheet, conflictResult.generalConflictRows);
    }

    // ReservationDaysを全削除→再生成
    deleteReservationDays_(daysSheet, existing.data.reservation_id);
    appendReservationDays_(daysSheet, existing.data.reservation_id, tripData.vehicleId, tripData.start, tripData.end);

    // Reservationsの日程を更新
    const startDateIdx = resHeaders.indexOf('start_date');
    const endDateIdx = resHeaders.indexOf('end_date');
    const updatedAtIdx = resHeaders.indexOf('updated_at');
    const updatedByIdx = resHeaders.indexOf('updated_by');
    const syncAtIdx = resHeaders.indexOf('source_last_sync_at');

    if (startDateIdx !== -1) resSheet.getRange(existing.rowIndex, startDateIdx + 1).setValue(tripData.start);
    if (endDateIdx !== -1) resSheet.getRange(existing.rowIndex, endDateIdx + 1).setValue(tripData.end);
    if (updatedAtIdx !== -1) resSheet.getRange(existing.rowIndex, updatedAtIdx + 1).setValue(now);
    if (updatedByIdx !== -1) resSheet.getRange(existing.rowIndex, updatedByIdx + 1).setValue(user);
    if (syncAtIdx !== -1) resSheet.getRange(existing.rowIndex, syncAtIdx + 1).setValue(now);
  }

  // 担当者変更チェック
  const personInfo = getPersonInfo_(tripData.personId);
  const workerCodeIdx = resHeaders.indexOf('worker_code');
  const workerNameIdx = resHeaders.indexOf('worker_name');
  const deptIdx = resHeaders.indexOf('dept_name');
  const purposeIdx = resHeaders.indexOf('purpose');
  const destIdx = resHeaders.indexOf('destination');
  const memoIdx = resHeaders.indexOf('memo');
  const updatedAtIdx2 = resHeaders.indexOf('updated_at');
  const updatedByIdx2 = resHeaders.indexOf('updated_by');
  const syncAtIdx2 = resHeaders.indexOf('source_last_sync_at');

  // 常に最新情報で上書き（担当者・用件・行先・備考）
  if (workerCodeIdx !== -1) resSheet.getRange(existing.rowIndex, workerCodeIdx + 1).setValue(tripData.personId);
  if (workerNameIdx !== -1) resSheet.getRange(existing.rowIndex, workerNameIdx + 1).setValue(personInfo.name);
  if (deptIdx !== -1) resSheet.getRange(existing.rowIndex, deptIdx + 1).setValue(personInfo.dept);
  if (purposeIdx !== -1) resSheet.getRange(existing.rowIndex, purposeIdx + 1).setValue('[出張計画] ' + (tripData['用件'] || ''));
  if (destIdx !== -1) resSheet.getRange(existing.rowIndex, destIdx + 1).setValue(tripData['行先'] || '');
  if (memoIdx !== -1) resSheet.getRange(existing.rowIndex, memoIdx + 1).setValue(tripData['備考'] || '');
  if (updatedAtIdx2 !== -1) resSheet.getRange(existing.rowIndex, updatedAtIdx2 + 1).setValue(now);
  if (updatedByIdx2 !== -1) resSheet.getRange(existing.rowIndex, updatedByIdx2 + 1).setValue(user);
  if (syncAtIdx2 !== -1) resSheet.getRange(existing.rowIndex, syncAtIdx2 + 1).setValue(now);

  return { ok: true, reservationId: existing.data.reservation_id };
}

/**
 * 予約を削除（delete）→ statusをcancelledに更新
 * @private
 */
function syncDelete_(tripId, lock) {
  const resSheet = getVehicleDbSheet_(CONFIG.VEHICLE_DB.SHEETS.RESERVATIONS);
  const existing = findReservationBySourceId_(resSheet, tripId);

  if (!existing) {
    return { ok: true, status: 'skipped', message: '対応する予約が見つかりません' };
  }

  cancelReservation_(resSheet, existing.rowIndex);
  return { ok: true, status: 'cancelled', reservationId: existing.data.reservation_id };
}

// ============================================
// 内部ヘルパー関数
// ============================================

/**
 * source_idで予約を検索
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} resSheet
 * @param {string} tripId
 * @returns {{rowIndex: number, data: Object}|null}
 */
function findReservationBySourceId_(resSheet, tripId) {
  const data = resSheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const parsed = parseVehicleDbSheet_(data, 'reservation_id');
  if (parsed.headerRowIdx === -1) return null;

  const headers = parsed.headers;
  const sourceIdIdx = headers.indexOf('source_id');
  const statusIdx = headers.indexOf('status');
  if (sourceIdIdx === -1) return null;

  for (let i = parsed.headerRowIdx + 1; i < data.length; i++) {
    if (data[i][sourceIdIdx] === tripId && data[i][statusIdx] === 'active') {
      const rowData = {};
      headers.forEach((h, idx) => {
        rowData[h] = data[i][idx];
      });
      return { rowIndex: i + 1, data: rowData }; // 1-based row index for Sheet
    }
  }
  return null;
}

/**
 * 競合チェック
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} resSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} daysSheet
 * @param {string} vehicleId
 * @param {string} startDate
 * @param {string} endDate
 * @param {string|null} excludeReservationId - 除外する予約ID（自分自身）
 * @returns {{hasConflict: boolean, seisanConflicts: Object[], generalConflictRows: number[]}}
 */
function checkConflicts_(resSheet, daysSheet, vehicleId, startDate, endDate, excludeReservationId) {
  const result = { hasConflict: false, seisanConflicts: [], generalConflictRows: [] };

  // ReservationDaysから期間内・同車両の予約IDを集める
  const daysData = daysSheet.getDataRange().getValues();
  if (daysData.length < 2) return result;

  const daysParsed = parseVehicleDbSheet_(daysData, 'reservation_id');
  if (daysParsed.headerRowIdx === -1) return result;

  const daysHeaders = daysParsed.headers;
  const dResIdIdx = daysHeaders.indexOf('reservation_id');
  const dVehicleIdx = daysHeaders.indexOf('vehicle_id');
  const dDateIdx = daysHeaders.indexOf('date');

  const conflictResIds = new Set();
  for (let i = daysParsed.headerRowIdx + 1; i < daysData.length; i++) {
    if (daysData[i][dVehicleIdx] !== vehicleId) continue;
    const dayDate = formatDateForVehicle_(daysData[i][dDateIdx]);
    if (dayDate >= startDate && dayDate <= endDate) {
      const resId = daysData[i][dResIdIdx];
      if (resId !== excludeReservationId) {
        conflictResIds.add(resId);
      }
    }
  }

  if (conflictResIds.size === 0) return result;

  // Reservationsシートで予約詳細を確認
  const resData = resSheet.getDataRange().getValues();
  const resParsed = parseVehicleDbSheet_(resData, 'reservation_id');
  if (resParsed.headerRowIdx === -1) return result;

  const resHeaders = resParsed.headers;
  const rIdIdx = resHeaders.indexOf('reservation_id');
  const rStatusIdx = resHeaders.indexOf('status');
  const rSourceSystemIdx = resHeaders.indexOf('source_system');
  const rWorkerIdx = resHeaders.indexOf('worker_name');
  const rDeptIdx = resHeaders.indexOf('dept_name');
  const rStartIdx = resHeaders.indexOf('start_date');
  const rEndIdx = resHeaders.indexOf('end_date');

  for (let i = resParsed.headerRowIdx + 1; i < resData.length; i++) {
    const resId = resData[i][rIdIdx];
    if (!conflictResIds.has(resId)) continue;
    if (resData[i][rStatusIdx] !== 'active') continue;

    result.hasConflict = true;

    if (resData[i][rSourceSystemIdx] === CONFIG.VEHICLE_DB.SOURCE_SYSTEM) {
      // 生産工程表からの予約と競合
      result.seisanConflicts.push({
        reservation_id: resId,
        worker_name: resData[i][rWorkerIdx] || '',
        dept_name: resData[i][rDeptIdx] || '',
        start_date: formatDateForVehicle_(resData[i][rStartIdx]),
        end_date: formatDateForVehicle_(resData[i][rEndIdx])
      });
    } else {
      // 一般予約との競合
      result.generalConflictRows.push(i + 1); // 1-based
    }
  }

  return result;
}

/**
 * 一般予約をcancelledに更新
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} resSheet
 * @param {number[]} rowIndices - 1-based row indices
 */
function cancelGeneralReservations_(resSheet, rowIndices) {
  if (rowIndices.length === 0) return;

  const data = resSheet.getDataRange().getValues();
  const parsed = parseVehicleDbSheet_(data, 'reservation_id');
  const headers = parsed.headerRowIdx !== -1 ? parsed.headers : data[0];
  const statusIdx = headers.indexOf('status');
  const updatedAtIdx = headers.indexOf('updated_at');
  const updatedByIdx = headers.indexOf('updated_by');
  const now = nowIso_();
  const user = Session.getActiveUser().getEmail() || '';

  rowIndices.forEach(rowIdx => {
    if (statusIdx !== -1) resSheet.getRange(rowIdx, statusIdx + 1).setValue('cancelled');
    if (updatedAtIdx !== -1) resSheet.getRange(rowIdx, updatedAtIdx + 1).setValue(now);
    if (updatedByIdx !== -1) resSheet.getRange(rowIdx, updatedByIdx + 1).setValue(user);
  });
}

/**
 * 予約をcancelledに更新
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} resSheet
 * @param {number} rowIndex - 1-based
 */
function cancelReservation_(resSheet, rowIndex) {
  const data = resSheet.getDataRange().getValues();
  const parsed = parseVehicleDbSheet_(data, 'reservation_id');
  const headers = parsed.headerRowIdx !== -1 ? parsed.headers : data[0];
  const statusIdx = headers.indexOf('status');
  const updatedAtIdx = headers.indexOf('updated_at');
  const updatedByIdx = headers.indexOf('updated_by');
  const now = nowIso_();
  const user = Session.getActiveUser().getEmail() || '';

  if (statusIdx !== -1) resSheet.getRange(rowIndex, statusIdx + 1).setValue('cancelled');
  if (updatedAtIdx !== -1) resSheet.getRange(rowIndex, updatedAtIdx + 1).setValue(now);
  if (updatedByIdx !== -1) resSheet.getRange(rowIndex, updatedByIdx + 1).setValue(user);
}

/**
 * ReservationDaysに日付行を追加
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} daysSheet
 * @param {string} reservationId
 * @param {string} vehicleId
 * @param {string} startDate
 * @param {string} endDate
 */
function appendReservationDays_(daysSheet, reservationId, vehicleId, startDate, endDate) {
  const dates = expandDates_(startDate, endDate);
  const now = nowIso_();

  const daysData = daysSheet.getDataRange().getValues();
  const daysParsed = parseVehicleDbSheet_(daysData, 'reservation_id');
  const daysHeaders = daysParsed.headerRowIdx !== -1 ? daysParsed.headers : daysData[0];

  dates.forEach(date => {
    const dateStr = formatDateForVehicle_(date);
    const conflictKey = vehicleId + '_' + dateStr + '_' + CONFIG.VEHICLE_DB.DEFAULT_SLOT;

    const newRow = daysHeaders.map(h => {
      switch (h) {
        case 'reservation_id': return reservationId;
        case 'vehicle_id': return vehicleId;
        case 'date': return dateStr;
        case 'slot': return CONFIG.VEHICLE_DB.DEFAULT_SLOT;
        case 'start_time': return '';
        case 'end_time': return '';
        case 'effective_start': return '08:00';
        case 'effective_end': return '18:00';
        case 'conflict_key': return conflictKey;
        case 'created_at': return now;
        case 'updated_at': return now;
        case 'notes': return '';
        default: return '';
      }
    });
    daysSheet.appendRow(newRow);
  });
}

/**
 * ReservationDaysから指定reservation_idの行を全削除
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} daysSheet
 * @param {string} reservationId
 */
function deleteReservationDays_(daysSheet, reservationId) {
  const data = daysSheet.getDataRange().getValues();
  if (data.length < 2) return;

  const parsed = parseVehicleDbSheet_(data, 'reservation_id');
  if (parsed.headerRowIdx === -1) return;

  const headers = parsed.headers;
  const resIdIdx = headers.indexOf('reservation_id');
  if (resIdIdx === -1) return;

  // 下から上に向かって削除（行番号のズレを防ぐ）
  for (let i = data.length - 1; i >= parsed.headerRowIdx + 1; i--) {
    if (data[i][resIdIdx] === reservationId) {
      daysSheet.deleteRow(i + 1); // 1-based
    }
  }
}

/**
 * 出張予定(site)に車両を設定した際、関連する移動(move)にも同じ車両を自動適用する
 * 関連移動 = 同じpersonId + kind='move' + 日付が出張の前日or翌日
 * @param {string} tripId - 対象の出張予定tripId
 * @param {Object} tripData - {personId, start, end, vehicleId, kind}
 */
function propagateVehicleToRelatedMoves(tripId, tripData) {
  // siteトリップ以外、または車両未指定の場合はスキップ
  if (!tripData.vehicleId) return;
  if (tripData.kind === 'move') return; // 移動自体は伝播しない

  try {
    const sheet = getSheet(CONFIG.SHEETS.TRIPS);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => String(h).trim());
    const tripIdIdx = headers.indexOf('tripId');
    const personIdIdx = headers.indexOf('personId');
    const startIdx = headers.indexOf('start');
    const endIdx = headers.indexOf('end');
    const kindIdx = headers.indexOf('kind');
    const vehicleIdIdx = headers.indexOf('vehicleId');

    if (tripIdIdx === -1 || personIdIdx === -1 || kindIdx === -1 || vehicleIdIdx === -1) return;

    const personId = tripData.personId;
    const siteStart = tripData.start; // YYYY-MM-DD
    const siteEnd = tripData.end;

    // 前日・翌日を計算
    const dayBefore = Utils_addDays_(siteStart, -1);
    const dayAfter = Utils_addDays_(siteEnd, 1);

    let updatedCount = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][tripIdIdx] === tripId) continue; // 自分自身はスキップ
      if (data[i][personIdIdx] !== personId) continue;
      if (data[i][kindIdx] !== 'move') continue;

      const moveStart = formatDate(data[i][startIdx]);
      const moveEnd = formatDate(data[i][endIdx]);

      // 前日移動 or 翌日移動に該当するか
      const isRelated = (moveStart === dayBefore || moveEnd === dayBefore ||
                         moveStart === dayAfter || moveEnd === dayAfter ||
                         (moveStart >= siteStart && moveEnd <= siteEnd));

      if (isRelated) {
        const currentVehicleId = data[i][vehicleIdIdx];
        if (currentVehicleId !== tripData.vehicleId) {
          // 車両IDを更新
          sheet.getRange(i + 1, vehicleIdIdx + 1).setValue(tripData.vehicleId);
          updatedCount++;

          // 車両予約も同期
          const moveTripId = data[i][tripIdIdx];
          const moveTripData = {};
          headers.forEach((h, idx) => { moveTripData[h] = data[i][idx]; });
          moveTripData.vehicleId = tripData.vehicleId;
          moveTripData.start = formatDate(moveTripData.start);
          moveTripData.end = formatDate(moveTripData.end);

          try {
            const syncResult = syncTripToVehicleReservation(moveTripId, moveTripData, 'create');
            if (syncResult.ok) {
              const vehResIdIdx = headers.indexOf('vehicleReservationId');
              const vehSyncIdx = headers.indexOf('vehicleSyncStatus');
              if (vehResIdIdx !== -1) sheet.getRange(i + 1, vehResIdIdx + 1).setValue(syncResult.reservationId || '');
              if (vehSyncIdx !== -1) sheet.getRange(i + 1, vehSyncIdx + 1).setValue('synced');
            }
          } catch (e) {
            Logger.log('移動予定の車両同期エラー (tripId=' + moveTripId + '): ' + e.message);
          }
        }
      }
    }

    if (updatedCount > 0) {
      Logger.log('関連移動予定に車両を自動適用: ' + updatedCount + '件');
    }
  } catch (e) {
    Logger.log('propagateVehicleToRelatedMoves error: ' + e.message);
  }
}

/**
 * 日付文字列に日数を加算
 * @param {string} dateStr - YYYY-MM-DD
 * @param {number} days
 * @returns {string} YYYY-MM-DD
 */
function Utils_addDays_(dateStr, days) {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + days);
  return formatDateForVehicle_(d);
}

/**
 * 車両統合セットアップ（GASエディタから手動実行用）
 * Tripsシートに車両管理用カラムを追加する
 */
function setupVehicleIntegration() {
  ensureTripsVehicleColumns();
  Logger.log('車両統合セットアップ完了');
}

/**
 * デバッグ用: 車両マスタ取得テスト
 * GASエディタから実行してログを確認
 */
function debugGetVehicleMaster() {
  try {
    const sheet = getVehicleDbSheet_(CONFIG.VEHICLE_DB.SHEETS.VEHICLES);
    const data = sheet.getDataRange().getValues();
    console.log('シート行数: ' + data.length);

    const parsed = parseVehicleDbSheet_(data, 'vehicle_id');
    console.log('ヘッダー行インデックス: ' + parsed.headerRowIdx);
    console.log('ヘッダー: ' + JSON.stringify(parsed.headers));

    if (parsed.headerRowIdx !== -1 && parsed.rows.length > 0) {
      console.log('1行目データ: ' + JSON.stringify(parsed.rows[0]));
      // active列の型を確認
      const headers = parsed.headers;
      const activeIdx = headers.indexOf('active');
      console.log('active列インデックス: ' + activeIdx);
      if (activeIdx >= 0) {
        for (var i = parsed.headerRowIdx + 1; i < Math.min(data.length, parsed.headerRowIdx + 4); i++) {
          var val = data[i][activeIdx];
          console.log('行' + i + ' active値: ' + JSON.stringify(val) + ' 型: ' + typeof val);
        }
      }
    }

    var vehicles = getVehicleMaster();
    console.log('取得車両数: ' + vehicles.length);
    console.log('車両一覧: ' + JSON.stringify(vehicles));
  } catch (e) {
    console.log('エラー: ' + e.message);
    console.log('スタック: ' + e.stack);
  }
}
