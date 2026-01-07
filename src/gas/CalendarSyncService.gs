/**
 * CalendarSyncService.gs
 * TSCカレンダー → 出張計画への同期処理
 *
 * 同期方向: TSC Google Calendar → Webアプリ出張計画（一方向のみ）
 * 重複防止: uniqueKey = sourceEventId + ':' + personId
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
 * TSCカレンダーイベントを出張計画に同期
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} 同期結果
 */
function syncTSCCalendarToTrips(startDate, endDate) {
  Logger.log('========== TSCカレンダー同期開始 ==========');
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

    // 4. 除外対象（OFF）をフィルタリング
    const includeOff = CONFIG.TSC_DEPARTMENT.SYNC_SETTINGS.INCLUDE_OFF_EVENTS;
    const filteredEvents = normalizedEvents.filter(event => {
      if (event.kind === EventKind.OFF && !includeOff) {
        result.summary.skippedEvents++;
        result.details.skipped.push({
          eventId: event.sourceEventId,
          title: event.title,
          reason: '休みイベントは除外設定'
        });
        return false;
      }
      return true;
    });

    result.summary.processedEvents = filteredEvents.length;
    Logger.log(`処理対象イベント: ${filteredEvents.length}件 (除外: ${result.summary.skippedEvents}件)`);

    // 5. 各イベントをTSC部署メンバー全員に展開してupsert
    for (const event of filteredEvents) {
      for (const personId of tscMemberIds) {
        result.summary.expandedRecords++;

        try {
          const tripData = createTripDataFromEvent(event, personId);
          const upsertResult = upsertTripBySourceEvent(tripData);

          if (upsertResult.action === 'created') {
            result.summary.created++;
            result.details.created.push({
              tripId: upsertResult.trip.tripId,
              personId,
              title: event.title,
              start: event.start,
              end: event.end
            });
          } else if (upsertResult.action === 'updated') {
            result.summary.updated++;
            result.details.updated.push({
              tripId: upsertResult.trip.tripId,
              personId,
              title: event.title,
              start: event.start,
              end: event.end
            });
          } else if (upsertResult.action === 'skipped') {
            result.summary.skipped++;
            result.details.skipped.push({
              personId,
              title: event.title,
              reason: upsertResult.reason
            });
          }
        } catch (e) {
          result.summary.errors++;
          result.details.errors.push({
            personId,
            eventId: event.sourceEventId,
            title: event.title,
            error: e.message
          });
          Logger.log(`エラー: personId=${personId}, event=${event.title}: ${e.message}`);
        }
      }
    }

    result.success = true;
    Logger.log('========== TSCカレンダー同期完了 ==========');
    Logger.log(`結果: 作成=${result.summary.created}, 更新=${result.summary.updated}, スキップ=${result.summary.skipped}, エラー=${result.summary.errors}`);

  } catch (e) {
    Logger.log(`同期エラー: ${e.message}`);
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
  // CONFIGから取得（将来的にはシートから動的取得に変更可能）
  return CONFIG.TSC_DEPARTMENT.MEMBER_PERSON_IDS || [];
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
    // 完全一致で「移動」の場合のみMOVE
    kind = EventKind.MOVE;
    displayTitle = '移動';
  } else if (/^\d{1,2}:\d{2}\s*休$/.test(title) || title.includes('休み') || title.includes('振替')) {
    // 「08:00 休」「休み」「振替」などはOFF
    kind = EventKind.OFF;
    displayTitle = title;
  }

  return {
    sourceEventId: event.id,
    title: displayTitle,
    originalTitle: title,
    kind: kind,
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
  // uniqueKeyを生成（重複防止用）
  const sourceKey = `${event.sourceEventId}:${personId}`;

  return {
    personId: personId,
    start: event.start,
    end: event.end,
    '行先': event.kind === EventKind.MOVE ? '' : event.title,
    '用件': event.kind === EventKind.MOVE ? '移動' : '',
    'jobId(任意)': '',
    '備考': event.location ? `場所: ${event.location}` : '',
    kind: event.kind === EventKind.MOVE ? 'move' : 'site',
    processId: '',
    source: 'calendar_sync',
    sourceEventId: event.sourceEventId,
    sourceKey: sourceKey,
    isLocked: false
  };
}

/**
 * sourceEventIdをキーにTripをupsert（重複防止）
 * @param {Object} tripData - Trip用データ
 * @returns {Object} { action: 'created'|'updated'|'skipped', trip?, reason? }
 */
function upsertTripBySourceEvent(tripData) {
  const sheet = getSheet(CONFIG.SHEETS.TRIPS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // sourceKey列のインデックスを取得（なければ追加が必要）
  let sourceKeyIndex = headers.indexOf('sourceKey');
  let sourceEventIdIndex = headers.indexOf('sourceEventId');

  // 既存レコードを検索（sourceKeyで一意検索）
  const uniqueKey = tripData.sourceKey;
  let existingRowIndex = -1;

  if (sourceKeyIndex !== -1) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][sourceKeyIndex] === uniqueKey) {
        existingRowIndex = i;
        break;
      }
    }
  }

  // 既存レコードがある場合は更新
  if (existingRowIndex !== -1) {
    const tripIdIndex = headers.indexOf('tripId');
    const existingTripId = data[existingRowIndex][tripIdIndex];

    // ロック済みチェック
    const isLockedIndex = headers.indexOf('isLocked');
    if (isLockedIndex !== -1) {
      const isLocked = data[existingRowIndex][isLockedIndex];
      if (isLocked === true || isLocked === 'TRUE' || isLocked === 'true') {
        return {
          action: 'skipped',
          reason: '確定済み（ロック）のため更新をスキップ'
        };
      }
    }

    // 更新
    const now = new Date();
    const updatedAtIndex = headers.indexOf('updatedAt');

    for (let col = 0; col < headers.length; col++) {
      const header = headers[col];
      if (header === 'tripId' || header === 'sourceKey' || header === 'sourceEventId') {
        continue; // ID系は更新しない
      }
      if (header === 'updatedAt') {
        sheet.getRange(existingRowIndex + 1, col + 1).setValue(now);
      } else if (tripData.hasOwnProperty(header)) {
        sheet.getRange(existingRowIndex + 1, col + 1).setValue(tripData[header]);
      }
    }

    return {
      action: 'updated',
      trip: {
        tripId: existingTripId,
        ...tripData,
        updatedAt: formatDateTime(now)
      }
    };
  }

  // 新規作成
  const tripId = Utilities.getUuid();
  const now = new Date();
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

  sheet.appendRow(newRow);

  return {
    action: 'created',
    trip: {
      tripId,
      ...tripData,
      genKey,
      updatedAt: formatDateTime(now)
    }
  };
}

/**
 * 同期結果をフォーマット（UI表示用）
 * @param {Object} result - 同期結果
 * @returns {string} フォーマットされたメッセージ
 */
function formatSyncResultMessage(result) {
  if (!result.success) {
    return `同期エラー: ${result.error}`;
  }

  const s = result.summary;
  let message = `同期完了\n`;
  message += `━━━━━━━━━━━━━━━━━━━\n`;
  message += `対象期間のイベント: ${s.totalEvents}件\n`;
  message += `処理対象: ${s.processedEvents}件\n`;
  message += `TSC部署メンバー: ${result.tscMembers.length}名\n`;
  message += `━━━━━━━━━━━━━━━━━━━\n`;
  message += `新規作成: ${s.created}件\n`;
  message += `更新: ${s.updated}件\n`;
  message += `スキップ: ${s.skipped}件\n`;
  if (s.errors > 0) {
    message += `エラー: ${s.errors}件\n`;
  }
  message += `━━━━━━━━━━━━━━━━━━━\n`;
  message += `実行時間: ${result.executionTime}ms`;

  return message;
}

/**
 * 同期テスト（ドライラン）
 * 実際のデータ更新は行わず、処理内容のみ確認
 * @param {string} startDate
 * @param {string} endDate
 * @returns {Object}
 */
function testSyncTSCCalendar(startDate, endDate) {
  Logger.log('========== TSCカレンダー同期テスト（ドライラン） ==========');

  const result = {
    wouldProcess: [],
    tscMembers: getTSCMemberIds(),
    calendarEvents: []
  };

  try {
    // カレンダーイベント取得
    const calendarResult = getTSCCalendarEvents(startDate, endDate);
    if (!calendarResult.success) {
      throw new Error(calendarResult.errorMessage);
    }

    result.calendarEvents = calendarResult.events;

    // 分類結果を表示
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
