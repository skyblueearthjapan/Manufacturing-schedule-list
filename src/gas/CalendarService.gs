/**
 * CalendarService.gs
 * Google Calendar APIとの連携を担当
 *
 * 重要: スコープ追加後は必ず「新しいバージョン」として再デプロイが必要
 *
 * appsscript.json に必要なスコープ:
 * - https://www.googleapis.com/auth/calendar.readonly （カレンダー読み取り）
 * - https://www.googleapis.com/auth/script.external_request （ICS取得用）
 * - https://www.googleapis.com/auth/userinfo.email （ユーザー情報取得用）
 *
 * デプロイ設定:
 * - Execute as: USER_ACCESSING（アクセスしているユーザー）
 * - Who has access: 必要に応じて
 *
 * カレンダー共有設定:
 * - アクセスユーザーがカレンダーの「閲覧者」以上の権限を持っていること
 * - または「すべてのイベントの詳細を見る」権限が付与されていること
 */

// 電気カレンダーの公開ICS URL（フォールバック用）
const ELECTRICAL_ICS_URL = 'https://calendar.google.com/calendar/ical/electrical-01%40lineworks-local.info/public/basic.ics';

/**
 * エラータイプ定義
 */
const CalendarErrorType = {
  PERMISSION_DENIED: 'PERMISSION_DENIED',
  CALENDAR_NOT_FOUND: 'CALENDAR_NOT_FOUND',
  NETWORK_ERROR: 'NETWORK_ERROR',
  PARSE_ERROR: 'PARSE_ERROR',
  UNKNOWN: 'UNKNOWN'
};

/**
 * 実行コンテキストのデバッグ情報を取得
 * これで「誰として実行されているか」を確認できる
 */
function getExecutionContext() {
  let activeUser = 'unknown';
  let effectiveUser = 'unknown';

  try {
    activeUser = Session.getActiveUser().getEmail() || '(empty)';
  } catch (e) {
    activeUser = '(error: ' + e.message + ')';
  }

  try {
    effectiveUser = Session.getEffectiveUser().getEmail() || '(empty)';
  } catch (e) {
    effectiveUser = '(error: ' + e.message + ')';
  }

  return {
    activeUser: activeUser,
    effectiveUser: effectiveUser,
    timestamp: new Date().toISOString()
  };
}

/**
 * カレンダーエラーを分類
 * @param {Error} e - エラーオブジェクト
 * @returns {string} エラータイプ
 */
function classifyCalendarError(e) {
  const msg = (e.message || '').toLowerCase();
  if (msg.includes('permission') || msg.includes('forbidden') || msg.includes('403') || msg.includes('calendar.readonly')) {
    return CalendarErrorType.PERMISSION_DENIED;
  }
  if (msg.includes('not found') || msg.includes('404')) {
    return CalendarErrorType.CALENDAR_NOT_FOUND;
  }
  if (msg.includes('network') || msg.includes('timeout') || msg.includes('fetch')) {
    return CalendarErrorType.NETWORK_ERROR;
  }
  return CalendarErrorType.UNKNOWN;
}

/**
 * ユーザー向けエラーメッセージを生成（デバッグ情報付き）
 * @param {string} errorType - エラータイプ
 * @param {string} calendarId - カレンダーID
 * @param {Object} debugInfo - デバッグ情報
 * @param {string} method - 取得方式 ('CalendarApp' または 'ICS')
 */
function getCalendarErrorMessage(errorType, calendarId, debugInfo, method) {
  const context = debugInfo || getExecutionContext();
  const methodLabel = method || 'CalendarApp';

  let baseMessage = '';
  let requiredScope = '';

  switch (errorType) {
    case CalendarErrorType.PERMISSION_DENIED:
      baseMessage = `カレンダーへのアクセス権限がありません。`;
      requiredScope = 'calendar.readonly';
      break;
    case CalendarErrorType.CALENDAR_NOT_FOUND:
      baseMessage = `カレンダーが見つかりません。`;
      requiredScope = 'calendar.readonly';
      break;
    case CalendarErrorType.NETWORK_ERROR:
      baseMessage = `ネットワークエラーが発生しました。`;
      requiredScope = 'script.external_request';
      break;
    default:
      baseMessage = `カレンダーの読み込みに失敗しました。`;
      requiredScope = 'calendar.readonly';
  }

  return `${baseMessage}\n\n` +
    `【デバッグ情報】\n` +
    `取得方式: ${methodLabel}\n` +
    `対象カレンダー: ${calendarId}\n` +
    `実行ユーザー: ${context.activeUser}\n` +
    `必要スコープ: ${requiredScope}\n\n` +
    `【対処方法】\n` +
    `1. 「${context.activeUser}」がカレンダーにアクセス権を持っているか確認\n` +
    `2. Webアプリを「新しいバージョン」として再デプロイ\n` +
    `3. 初回アクセス時の承認ダイアログで許可`;
}

/**
 * 指定カレンダーのイベントを取得（CalendarApp使用）
 * @param {string} calendarId - カレンダーID
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} { success, events, error, errorType, errorMessage, debug }
 */
function getCalendarEvents(calendarId, startDate, endDate) {
  // デバッグ情報を最初に取得
  const debugInfo = getExecutionContext();
  Logger.log('[getCalendarEvents] ========== START ==========');
  Logger.log('[getCalendarEvents] Calendar ID: ' + calendarId);
  Logger.log('[getCalendarEvents] Date range: ' + startDate + ' to ' + endDate);
  Logger.log('[getCalendarEvents] Effective User: ' + debugInfo.effectiveUser);
  Logger.log('[getCalendarEvents] Active User: ' + debugInfo.activeUser);

  try {
    // CalendarApp でカレンダーを取得
    const calendar = CalendarApp.getCalendarById(calendarId);

    if (!calendar) {
      Logger.log('[getCalendarEvents] FAILED: CalendarApp.getCalendarById returned null');
      Logger.log('[getCalendarEvents] → カレンダーが見つからないか、権限がありません');
      return {
        success: false,
        events: [],
        errorType: CalendarErrorType.CALENDAR_NOT_FOUND,
        errorMessage: getCalendarErrorMessage(CalendarErrorType.CALENDAR_NOT_FOUND, calendarId, debugInfo, 'CalendarApp'),
        debug: debugInfo
      };
    }

    // カレンダー名を取得して権限確認
    let calendarName = '';
    try {
      calendarName = calendar.getName();
      Logger.log('[getCalendarEvents] SUCCESS: Calendar found - Name: ' + calendarName);
    } catch (nameError) {
      Logger.log('[getCalendarEvents] WARNING: Could not get calendar name: ' + nameError.message);
    }

    const start = new Date(startDate + 'T00:00:00');
    const end = new Date(endDate + 'T23:59:59');

    const events = calendar.getEvents(start, end);
    Logger.log('[getCalendarEvents] SUCCESS: Found ' + events.length + ' events');

    const mappedEvents = events.map(event => ({
      id: event.getId(),
      title: event.getTitle(),
      description: event.getDescription() || '',
      location: event.getLocation() || '',
      startTime: formatDateTime(event.getStartTime()),
      endTime: formatDateTime(event.getEndTime()),
      isAllDay: event.isAllDayEvent(),
      startDate: formatDate(event.getStartTime()),
      endDate: formatDate(event.getEndTime()),
      color: event.getColor() || '',
      source: 'CalendarApp'
    }));

    Logger.log('[getCalendarEvents] ========== END (SUCCESS) ==========');
    return {
      success: true,
      events: mappedEvents,
      calendarName: calendarName,
      errorType: null,
      errorMessage: null,
      debug: debugInfo
    };
  } catch (e) {
    Logger.log('[getCalendarEvents] EXCEPTION: ' + e.message);
    Logger.log('[getCalendarEvents] Stack: ' + (e.stack || 'N/A'));
    Logger.log('[getCalendarEvents] ========== END (FAILED) ==========');

    const errorType = classifyCalendarError(e);
    return {
      success: false,
      events: [],
      errorType: errorType,
      errorMessage: getCalendarErrorMessage(errorType, calendarId, debugInfo, 'CalendarApp'),
      rawError: e.message,
      debug: debugInfo
    };
  }
}

/**
 * ICSファイルからイベントを取得（フォールバック用）
 * @param {string} icsUrl - ICSファイルのURL
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} { success, events, error }
 */
function getEventsFromICS(icsUrl, startDate, endDate) {
  try {
    Logger.log('[getEventsFromICS] Fetching ICS from: ' + icsUrl);

    const response = UrlFetchApp.fetch(icsUrl, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log('[getEventsFromICS] HTTP error: ' + responseCode);
      return {
        success: false,
        events: [],
        errorType: CalendarErrorType.NETWORK_ERROR,
        errorMessage: `ICS取得失敗\n\n【デバッグ情報】\n取得方式: ICS\nHTTPステータス: ${responseCode}\nURL: ${icsUrl}\n\n【対処方法】\nカレンダーが「一般公開」されているか確認してください`
      };
    }

    const icsContent = response.getContentText();
    const events = parseICS(icsContent, startDate, endDate);

    Logger.log('[getEventsFromICS] Parsed ' + events.length + ' events');

    return {
      success: true,
      events: events,
      errorType: null,
      errorMessage: null
    };
  } catch (e) {
    Logger.log('[getEventsFromICS] Error: ' + e.message);
    return {
      success: false,
      events: [],
      errorType: CalendarErrorType.NETWORK_ERROR,
      errorMessage: `ICS取得失敗\n\n【デバッグ情報】\n取得方式: ICS\nエラー: ${e.message}\nURL: ${icsUrl}\n必要スコープ: script.external_request\n\n【対処方法】\n1. Webアプリを「新しいバージョン」として再デプロイ\n2. カレンダーが「一般公開」されているか確認`
    };
  }
}

/**
 * ICSファイルをパース
 * @param {string} icsContent - ICSファイルの内容
 * @param {string} startDate - フィルタ開始日 (YYYY-MM-DD)
 * @param {string} endDate - フィルタ終了日 (YYYY-MM-DD)
 * @returns {Object[]} イベント配列
 */
function parseICS(icsContent, startDate, endDate) {
  const events = [];
  const lines = icsContent.replace(/\r\n /g, '').split(/\r?\n/);

  let currentEvent = null;
  const startFilter = new Date(startDate + 'T00:00:00');
  const endFilter = new Date(endDate + 'T23:59:59');

  for (const line of lines) {
    if (line === 'BEGIN:VEVENT') {
      currentEvent = {};
    } else if (line === 'END:VEVENT') {
      if (currentEvent && currentEvent.startDate) {
        // 日付フィルタリング
        const eventStart = new Date(currentEvent.startDate);
        const eventEnd = currentEvent.endDate ? new Date(currentEvent.endDate) : eventStart;

        if (eventStart <= endFilter && eventEnd >= startFilter) {
          events.push({
            id: currentEvent.uid || '',
            title: currentEvent.summary || '(タイトルなし)',
            description: currentEvent.description || '',
            location: currentEvent.location || '',
            startTime: currentEvent.startTime || '',
            endTime: currentEvent.endTime || '',
            isAllDay: currentEvent.isAllDay || false,
            startDate: currentEvent.startDate,
            endDate: currentEvent.endDate || currentEvent.startDate,
            color: '',
            source: 'ICS'
          });
        }
      }
      currentEvent = null;
    } else if (currentEvent) {
      const colonIndex = line.indexOf(':');
      if (colonIndex > 0) {
        const key = line.substring(0, colonIndex);
        const value = line.substring(colonIndex + 1);

        if (key === 'SUMMARY') {
          currentEvent.summary = value;
        } else if (key === 'DESCRIPTION') {
          currentEvent.description = value.replace(/\\n/g, '\n').replace(/\\,/g, ',');
        } else if (key === 'LOCATION') {
          currentEvent.location = value;
        } else if (key === 'UID') {
          currentEvent.uid = value;
        } else if (key.startsWith('DTSTART')) {
          const parsed = parseICSDateTime(key, value);
          currentEvent.startDate = parsed.date;
          currentEvent.startTime = parsed.dateTime;
          currentEvent.isAllDay = parsed.isAllDay;
        } else if (key.startsWith('DTEND')) {
          const parsed = parseICSDateTime(key, value);
          currentEvent.endDate = parsed.date;
          currentEvent.endTime = parsed.dateTime;
        }
      }
    }
  }

  return events;
}

/**
 * ICSの日時をパース
 * @param {string} key - DTSTARTまたはDTEND（パラメータ付き）
 * @param {string} value - 日時値
 * @returns {Object} { date, dateTime, isAllDay }
 */
function parseICSDateTime(key, value) {
  const isAllDay = key.includes('VALUE=DATE') || value.length === 8;

  let date, dateTime;

  if (isAllDay) {
    // YYYYMMDD形式
    date = value.substring(0, 4) + '-' + value.substring(4, 6) + '-' + value.substring(6, 8);
    dateTime = '';
  } else {
    // YYYYMMDDTHHMMSS or YYYYMMDDTHHMMSSZ形式
    const cleanValue = value.replace('Z', '');
    date = cleanValue.substring(0, 4) + '-' + cleanValue.substring(4, 6) + '-' + cleanValue.substring(6, 8);
    if (cleanValue.length >= 15) {
      const hours = cleanValue.substring(9, 11);
      const minutes = cleanValue.substring(11, 13);
      const seconds = cleanValue.substring(13, 15);
      dateTime = date + 'T' + hours + ':' + minutes + ':' + seconds + '+09:00';
    } else {
      dateTime = date + 'T00:00:00+09:00';
    }
  }

  return { date, dateTime, isAllDay };
}

/**
 * TSCカレンダーのイベントを取得
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} { success, events, errorType, errorMessage }
 */
function getTSCCalendarEvents(startDate, endDate) {
  Logger.log('[getTSCCalendarEvents] Called: ' + startDate + ' to ' + endDate);
  const result = getCalendarEvents(CONFIG.CALENDARS.TSC, startDate, endDate);
  result.source = 'CalendarApp';

  // TSCはCalendarAppでのみ取得可能（公開ICSなし）
  return result;
}

/**
 * 電気カレンダーのイベントを取得（ICSフォールバック付き）
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} { success, events, errorType, errorMessage, source }
 */
function getElectricalCalendarEvents(startDate, endDate) {
  Logger.log('[getElectricalCalendarEvents] Called: ' + startDate + ' to ' + endDate);

  // 優先1: CalendarAppで取得を試みる
  const calendarResult = getCalendarEvents(CONFIG.CALENDARS.ELECTRICAL, startDate, endDate);

  if (calendarResult.success) {
    Logger.log('[getElectricalCalendarEvents] Success via CalendarApp');
    return {
      ...calendarResult,
      source: 'CalendarApp'
    };
  }

  Logger.log('[getElectricalCalendarEvents] CalendarApp failed, trying ICS fallback...');

  // 優先2: ICSフォールバック
  const icsResult = getEventsFromICS(ELECTRICAL_ICS_URL, startDate, endDate);

  if (icsResult.success) {
    Logger.log('[getElectricalCalendarEvents] Success via ICS fallback');
    return {
      ...icsResult,
      source: 'ICS',
      // ICSで成功したことをログ（元のエラーも記録）
      originalError: calendarResult.errorMessage
    };
  }

  // 両方失敗
  Logger.log('[getElectricalCalendarEvents] Both methods failed');
  return {
    success: false,
    events: [],
    errorType: CalendarErrorType.PERMISSION_DENIED,
    errorMessage: `電気カレンダーの取得に失敗しました。\n\n` +
      `CalendarApp: ${calendarResult.errorMessage || '不明なエラー'}\n` +
      `ICS: ${icsResult.errorMessage || '不明なエラー'}`,
    source: 'none'
  };
}

/**
 * 複数カレンダーのイベントを一括取得
 * @param {string[]} calendarIds - カレンダーID配列
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} カレンダーIDをキーとした結果のマップ
 */
function getMultipleCalendarEvents(calendarIds, startDate, endDate) {
  const result = {};
  for (const calendarId of calendarIds) {
    result[calendarId] = getCalendarEvents(calendarId, startDate, endDate);
  }
  return result;
}

/**
 * イベントを日付でグループ化
 * @param {Object[]} events - イベント配列
 * @returns {Object} 日付をキーとしたイベント配列のマップ
 */
function groupEventsByDate(events) {
  const grouped = {};
  for (const event of events) {
    const date = event.startDate;
    if (!grouped[date]) {
      grouped[date] = [];
    }
    grouped[date].push(event);
  }
  return grouped;
}

/**
 * カレンダー接続テスト（デバッグ用）
 * GASエディタで直接実行するか、api_testCalendarConnection()経由で呼び出す
 * @returns {Object} テスト結果
 */
function testCalendarConnection() {
  Logger.log('========== カレンダー接続テスト開始 ==========');
  Logger.log('※ executeAs: USER_ACCESSING モードで動作');

  // 実行コンテキスト
  const context = getExecutionContext();
  Logger.log('実行ユーザー: ' + context.activeUser);
  Logger.log('（USER_ACCESSINGモードでは、このユーザーの権限でカレンダーにアクセス）');

  const results = {
    executionContext: context,
    calendars: {}
  };

  // TSCカレンダー
  Logger.log('--- TSCカレンダー テスト ---');
  try {
    const tscCal = CalendarApp.getCalendarById(CONFIG.CALENDARS.TSC);
    if (tscCal) {
      const name = tscCal.getName();
      Logger.log('✓ TSC: 成功 - カレンダー名: ' + name);
      results.calendars.TSC = {
        success: true,
        name: name,
        id: CONFIG.CALENDARS.TSC,
        method: 'CalendarApp'
      };
    } else {
      Logger.log('✗ TSC: 失敗 - getCalendarById が null を返しました');
      Logger.log('  → カレンダーが見つからないか、「' + context.activeUser + '」にアクセス権がありません');
      results.calendars.TSC = {
        success: false,
        error: 'Calendar not found or no permission',
        id: CONFIG.CALENDARS.TSC,
        method: 'CalendarApp',
        hint: '「' + context.activeUser + '」がカレンダーにアクセス権を持っているか確認してください'
      };
    }
  } catch (e) {
    Logger.log('✗ TSC: 例外発生 - ' + e.message);
    results.calendars.TSC = {
      success: false,
      error: e.message,
      errorType: classifyCalendarError(e),
      id: CONFIG.CALENDARS.TSC,
      method: 'CalendarApp'
    };
  }

  // 電気カレンダー（CalendarApp）
  Logger.log('--- 電気カレンダー（CalendarApp） テスト ---');
  try {
    const elecCal = CalendarApp.getCalendarById(CONFIG.CALENDARS.ELECTRICAL);
    if (elecCal) {
      const name = elecCal.getName();
      Logger.log('✓ 電気(CalendarApp): 成功 - カレンダー名: ' + name);
      results.calendars.ELECTRICAL_CalendarApp = {
        success: true,
        name: name,
        id: CONFIG.CALENDARS.ELECTRICAL,
        method: 'CalendarApp'
      };
    } else {
      Logger.log('✗ 電気(CalendarApp): 失敗 - getCalendarById が null を返しました');
      results.calendars.ELECTRICAL_CalendarApp = {
        success: false,
        error: 'Calendar not found or no permission',
        id: CONFIG.CALENDARS.ELECTRICAL,
        method: 'CalendarApp',
        hint: '「' + context.activeUser + '」がカレンダーにアクセス権を持っているか確認してください'
      };
    }
  } catch (e) {
    Logger.log('✗ 電気(CalendarApp): 例外発生 - ' + e.message);
    results.calendars.ELECTRICAL_CalendarApp = {
      success: false,
      error: e.message,
      errorType: classifyCalendarError(e),
      id: CONFIG.CALENDARS.ELECTRICAL,
      method: 'CalendarApp'
    };
  }

  // 電気カレンダー（ICSフォールバック）
  Logger.log('--- 電気カレンダー（ICS） テスト ---');
  try {
    const response = UrlFetchApp.fetch(ELECTRICAL_ICS_URL, { muteHttpExceptions: true });
    const code = response.getResponseCode();
    if (code === 200) {
      Logger.log('✓ 電気(ICS): 成功 - HTTP 200');
      results.calendars.ELECTRICAL_ICS = {
        success: true,
        httpStatus: code,
        url: ELECTRICAL_ICS_URL,
        method: 'ICS'
      };
    } else {
      Logger.log('✗ 電気(ICS): 失敗 - HTTP ' + code);
      results.calendars.ELECTRICAL_ICS = {
        success: false,
        httpStatus: code,
        url: ELECTRICAL_ICS_URL,
        method: 'ICS'
      };
    }
  } catch (e) {
    Logger.log('✗ 電気(ICS): 例外発生 - ' + e.message);
    results.calendars.ELECTRICAL_ICS = {
      success: false,
      error: e.message,
      url: ELECTRICAL_ICS_URL,
      method: 'ICS'
    };
  }

  // サマリー
  Logger.log('========== テスト結果サマリー ==========');
  Logger.log('TSC(CalendarApp): ' + (results.calendars.TSC?.success ? '✓ 成功' : '✗ 失敗'));
  Logger.log('電気(CalendarApp): ' + (results.calendars.ELECTRICAL_CalendarApp?.success ? '✓ 成功' : '✗ 失敗'));
  Logger.log('電気(ICS): ' + (results.calendars.ELECTRICAL_ICS?.success ? '✓ 成功' : '✗ 失敗'));
  Logger.log('========================================');

  return results;
}
