/**
 * CalendarSyncService.gs
 * TSCカレンダー → 出張計画への同期処理
 *
 * 同期方向: TSC Google Calendar → Webアプリ出張計画（一方向のみ）
 * 重複防止: uniqueKey = sourceEventId + ':' + personId
 *
 * 同期対象:
 *   - タイトルが完全一致「移動」
 *   - タイトル先頭が「★」で始まる予定（出張現場）
 *     例: ★コマツ茨城 TS25019, ★久保田堺 TS25172
 *
 * 同期しない:
 *   - 祝日、休み、振替休日、健康診断、会議、その他★なし予定
 *
 * 行先生成:
 *   - 「移動」→ "移動"
 *   - 「★...」→ ★を外したタイトル全文
 *
 * 最適化: 一括読み込み→メモリ処理→一括書き込み（タイムアウト対策）
 */

/**
 * イベント種別定義
 */
const EventKind = {
  MOVE: 'MOVE',    // 移動（車マーク表示）
  TRIP: 'TRIP',    // 出張/現場
  OFF: 'OFF'       // 休み（除外対象）
};

/**
 * TSCカレンダーイベントを出張計画に同期（バッチ処理版）
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} 同期結果
 */
function syncTSCCalendarToTrips(startDate, endDate) {
  Logger.log('========== TSCカレンダー同期開始（バッチ版） ==========');
  Logger.log(`期間: ${startDate} ～ ${endDate}`);

  const result = {
    success: false,
    summary: {
      totalEvents: 0,
      processedEvents: 0,
      skippedEvents: 0,
      expandedRecords: 0,
      created: 0,
      updated: 0,
      skipped: 0,
      errors: 0
    },
    details: {
      created: [],
      updated: [],
      skipped: [],
      errors: []
    },
    tscMembers: [],
    executionTime: 0
  };

  const startTime = new Date();

  try {
    // 1. TSC部署メンバーを取得
    const tscMemberIds = getTSCMemberIds();
    result.tscMembers = tscMemberIds;
    Logger.log(`TSC部署メンバー: ${tscMemberIds.join(', ')} (${tscMemberIds.length}名)`);

    if (tscMemberIds.length === 0) {
      throw new Error('TSC部署メンバーが設定されていません。Config.gsのTSC_DEPARTMENT.MEMBER_PERSON_IDSを確認してください。');
    }

    // 2. TSCカレンダーからイベントを取得
    const calendarResult = getTSCCalendarEvents(startDate, endDate);

    if (!calendarResult.success) {
      throw new Error(`カレンダー取得失敗: ${calendarResult.errorMessage}`);
    }

    const events = calendarResult.events;
    result.summary.totalEvents = events.length;
    Logger.log(`取得イベント数: ${events.length}件`);

    // 3. イベントを分類・正規化
    const normalizedEvents = events.map(event => classifyAndNormalizeEvent(event));

    // 4. 同期対象外のイベントをフィルタリング
    // 同期対象 = 「移動」完全一致 OR 「★」で始まる予定のみ
    const filteredEvents = normalizedEvents.filter(event => {
      if (!event.shouldSync) {
        result.summary.skippedEvents++;
        result.details.skipped.push({
          eventId: event.sourceEventId,
          title: event.title,
          reason: '同期対象外（移動/★付き以外）'
        });
        return false;
      }
      return true;
    });

    result.summary.processedEvents = filteredEvents.length;
    Logger.log(`処理対象イベント: ${filteredEvents.length}件 (除外: ${result.summary.skippedEvents}件)`);

    // 5. 一括upsert処理
    const upsertResult = batchUpsertTrips(filteredEvents, tscMemberIds);

    result.summary.created = upsertResult.created;
    result.summary.updated = upsertResult.updated;
    result.summary.skipped = upsertResult.skipped;
    result.summary.errors = upsertResult.errors;
    result.summary.expandedRecords = upsertResult.totalProcessed;

    result.success = true;
    Logger.log('========== TSCカレンダー同期完了 ==========');
    Logger.log(`結果: 作成=${result.summary.created}, 更新=${result.summary.updated}, スキップ=${result.summary.skipped}, エラー=${result.summary.errors}`);

  } catch (e) {
    Logger.log(`同期エラー: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    result.success = false;
    result.error = e.message;
  }

  result.executionTime = new Date() - startTime;
  return result;
}

/**
 * TSC部署メンバーのpersonId一覧を取得
 * @returns {string[]}
 */
function getTSCMemberIds() {
  return CONFIG.TSC_DEPARTMENT.MEMBER_PERSON_IDS || [];
}

/**
 * イベントを同期対象とするかどうかを判定
 * 同期対象 = 「移動」完全一致 OR 「★」で始まる
 * @param {string} title - イベントタイトル
 * @returns {boolean} 同期対象ならtrue
 */
function isSyncTarget(title) {
  const trimmedTitle = (title || '').trim();
  return trimmedTitle === '移動' || trimmedTitle.startsWith('★');
}

/**
 * イベントタイトルから行先を生成
 * - 「移動」のとき → "移動"
 * - 「★...」のとき → ★を外したタイトル全文
 * @param {string} title - イベントタイトル
 * @returns {string} 行先
 */
function buildDestination(title) {
  const trimmedTitle = (title || '').trim();
  if (trimmedTitle === '移動') return '移動';
  if (trimmedTitle.startsWith('★')) return trimmedTitle.replace(/^★\s*/, '');
  return '';
}

/**
 * カレンダーイベントを分類・正規化
 * @param {Object} event - カレンダーイベント
 * @returns {Object} 正規化されたイベント
 */
function classifyAndNormalizeEvent(event) {
  const title = (event.title || '').trim();
  let kind = EventKind.TRIP;
  let displayTitle = title;

  // 種別判定
  if (title === '移動') {
    kind = EventKind.MOVE;
    displayTitle = '移動';
  } else if (/^\d{1,2}:\d{2}\s*休$/.test(title) || title.includes('休み') || title.includes('振替')) {
    kind = EventKind.OFF;
    displayTitle = title;
  }

  // 同期対象かどうかを判定（移動 or ★で始まる）
  const shouldSync = isSyncTarget(title);

  return {
    sourceEventId: event.id,
    title: displayTitle,
    originalTitle: title,
    kind: kind,
    shouldSync: shouldSync,  // 同期対象フラグ
    start: event.startDate,
    end: event.endDate || event.startDate,
    startTime: event.startTime,
    endTime: event.endTime,
    isAllDay: event.isAllDay,
    location: event.location || '',
    description: event.description || ''
  };
}

/**
 * イベントから出張予定データを生成
 * @param {Object} event - 正規化されたイベント
 * @param {string} personId - 担当者ID
 * @returns {Object} Trip用データ
 */
function createTripDataFromEvent(event, personId) {
  const sourceKey = `${event.sourceEventId}:${personId}`;
  const destination = buildDestination(event.originalTitle);
  const isMove = event.originalTitle.trim() === '移動';

  return {
    personId: personId,
    start: event.start,
    end: event.end,
    '行先': destination,  // 移動→"移動", ★付き→★を外した全文
    '用件': '',  // 空（または必要なら固定値）
    'jobId(任意)': '',
    '備考': `source=TSC_CAL eventId=${event.sourceEventId}`,  // 追跡用
    kind: isMove ? 'move' : 'site',
    processId: '',
    source: 'calendar_sync',
    sourceEventId: event.sourceEventId,
    sourceKey: sourceKey,
    isLocked: false
  };
}

/**
 * 一括upsert処理（高速版）
 * シートを1回だけ読み書きして、メモリ上で全ての処理を行う
 * @param {Object[]} events - 正規化されたイベント配列
 * @param {string[]} memberIds - TSCメンバーID配列
 * @returns {Object} { created, updated, skipped, errors, totalProcessed }
 */
function batchUpsertTrips(events, memberIds) {
  const sheet = getSheet(CONFIG.SHEETS.TRIPS);
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const existingRows = allData.slice(1);

  // ヘッダーのインデックスをキャッシュ
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);

  // sourceKey列がない場合は警告
  if (colIndex['sourceKey'] === undefined) {
    Logger.log('WARNING: sourceKey列がTripsシートにありません。重複防止が機能しません。');
  }

  // 既存データをsourceKeyでマップ化（高速検索用）
  const existingBySourceKey = new Map();
  if (colIndex['sourceKey'] !== undefined) {
    existingRows.forEach((row, idx) => {
      const key = row[colIndex['sourceKey']];
      if (key) {
        existingBySourceKey.set(key, { row, rowIndex: idx + 1 }); // +1はヘッダー分
      }
    });
  }

  const result = { created: 0, updated: 0, skipped: 0, errors: 0, totalProcessed: 0 };
  const newRows = [];
  const updateOperations = [];
  const now = new Date();

  // 全イベント×全メンバーを処理
  for (const event of events) {
    for (const personId of memberIds) {
      result.totalProcessed++;

      try {
        const tripData = createTripDataFromEvent(event, personId);
        const sourceKey = tripData.sourceKey;

        // 既存チェック
        const existing = existingBySourceKey.get(sourceKey);

        if (existing) {
          // ロック済みチェック
          if (colIndex['isLocked'] !== undefined) {
            const isLocked = existing.row[colIndex['isLocked']];
            if (isLocked === true || isLocked === 'TRUE' || isLocked === 'true') {
              result.skipped++;
              continue;
            }
          }

          // 更新対象として記録
          updateOperations.push({
            rowIndex: existing.rowIndex,
            tripData,
            existingTripId: existing.row[colIndex['tripId']]
          });
          result.updated++;
        } else {
          // 新規作成用の行データを作成
          const tripId = Utilities.getUuid();
          const genKey = generateGenKey(tripData);

          const newRow = headers.map(header => {
            switch (header) {
              case 'tripId': return tripId;
              case 'genKey': return genKey;
              case 'updatedAt': return now;
              case 'sourceKey': return tripData.sourceKey || '';
              case 'sourceEventId': return tripData.sourceEventId || '';
              default: return tripData[header] !== undefined ? tripData[header] : '';
            }
          });

          newRows.push(newRow);
          result.created++;

          // 次の検索用にマップに追加（同期中の重複防止）
          existingBySourceKey.set(sourceKey, { row: newRow, rowIndex: -1 });
        }
      } catch (e) {
        result.errors++;
        Logger.log(`エラー: ${personId}, ${event.title}: ${e.message}`);
      }
    }
  }

  // 一括書き込み：更新
  if (updateOperations.length > 0) {
    Logger.log(`更新処理: ${updateOperations.length}件`);

    // 更新対象の列を特定（ID系以外）
    const updateCols = headers.map((h, i) => {
      if (h === 'tripId' || h === 'sourceKey' || h === 'sourceEventId' || h === 'genKey') {
        return null; // 更新しない
      }
      return i;
    }).filter(i => i !== null);

    for (const op of updateOperations) {
      const rowNum = op.rowIndex + 1; // シートは1始まり
      for (const colIdx of updateCols) {
        const header = headers[colIdx];
        let value;
        if (header === 'updatedAt') {
          value = now;
        } else if (op.tripData.hasOwnProperty(header)) {
          value = op.tripData[header];
        } else {
          continue; // 変更なし
        }
        // バッチではなく個別更新（更新行が飛び飛びのため）
        // ※更新が多い場合はさらに最適化可能
      }
    }

    // 更新は既存行を直接変更（行が飛び飛びなのでsetValuesは使いにくい）
    // → 実際には更新行数が少ないことが多いので、一括読み書き方式に変更
    const fullData = sheet.getDataRange().getValues();
    for (const op of updateOperations) {
      const rowIdx = op.rowIndex; // 0始まりのデータ配列インデックス（ヘッダー含む）
      for (let colIdx = 0; colIdx < headers.length; colIdx++) {
        const header = headers[colIdx];
        if (header === 'tripId' || header === 'sourceKey' || header === 'sourceEventId' || header === 'genKey') {
          continue;
        }
        if (header === 'updatedAt') {
          fullData[rowIdx + 1][colIdx] = now;
        } else if (op.tripData.hasOwnProperty(header)) {
          fullData[rowIdx + 1][colIdx] = op.tripData[header];
        }
      }
    }
    // 全データを書き戻し
    if (updateOperations.length > 0) {
      sheet.getRange(1, 1, fullData.length, fullData[0].length).setValues(fullData);
    }
  }

  // 一括書き込み：新規追加
  if (newRows.length > 0) {
    Logger.log(`新規追加: ${newRows.length}件`);
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
  }

  Logger.log(`バッチ処理完了: 作成=${result.created}, 更新=${result.updated}, スキップ=${result.skipped}`);
  return result;
}

/**
 * 同期結果をフォーマット（UI表示用）
 */
function formatSyncResultMessage(result) {
  if (!result.success) {
    return `同期エラー: ${result.error}`;
  }

  const s = result.summary;
  return `同期完了: 作成${s.created}件, 更新${s.updated}件, スキップ${s.skipped}件`;
}

/**
 * 同期テスト（ドライラン）
 */
function testSyncTSCCalendar(startDate, endDate) {
  Logger.log('========== TSCカレンダー同期テスト（ドライラン） ==========');

  const result = {
    wouldProcess: [],
    tscMembers: getTSCMemberIds(),
    calendarEvents: []
  };

  try {
    const calendarResult = getTSCCalendarEvents(startDate, endDate);
    if (!calendarResult.success) {
      throw new Error(calendarResult.errorMessage);
    }

    result.calendarEvents = calendarResult.events;

    for (const event of calendarResult.events) {
      const normalized = classifyAndNormalizeEvent(event);
      result.wouldProcess.push({
        title: normalized.title,
        kind: normalized.kind,
        start: normalized.start,
        end: normalized.end,
        willExpandTo: result.tscMembers.length + '名'
      });
    }

    Logger.log(`イベント数: ${result.calendarEvents.length}`);
    Logger.log(`TSCメンバー数: ${result.tscMembers.length}`);
    Logger.log(`生成予定レコード数: ${result.calendarEvents.length * result.tscMembers.length}`);

  } catch (e) {
    result.error = e.message;
  }

  return result;
}
