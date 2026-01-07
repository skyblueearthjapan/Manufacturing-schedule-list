/**
 * CalendarService.gs
 * Google Calendar APIとの連携を担当
 *
 * 重要: スコープ追加後は必ず再デプロイが必要
 * appsscript.json に以下のスコープが必要:
 * - https://www.googleapis.com/auth/calendar.readonly
 * - https://www.googleapis.com/auth/calendar.events.readonly
 * - https://www.googleapis.com/auth/script.external_request
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
 * ユーザー向けエラーメッセージを生成
 * @param {string} errorType - エラータイプ
 * @param {string} calendarId - カレンダーID
 * @returns {string} ユーザー向けメッセージ
 */
function getCalendarErrorMessage(errorType, calendarId) {
  switch (errorType) {
    case CalendarErrorType.PERMISSION_DENIED:
      return `カレンダーへのアクセス権限がありません。\n` +
             `対象: ${calendarId}\n\n` +
             `【対処方法】\n` +
             `1. Webアプリを再デプロイしてカレンダー権限を認可してください\n` +
             `2. または、カレンダーの所有者に共有設定を依頼してください`;
    case CalendarErrorType.CALENDAR_NOT_FOUND:
      return `カレンダーが見つかりません: ${calendarId}`;
    case CalendarErrorType.NETWORK_ERROR:
      return `ネットワークエラーが発生しました。しばらく待ってから再試行してください。`;
    default:
      return `カレンダーの読み込みに失敗しました。`;
  }
}

/**
 * 指定カレンダーのイベントを取得（CalendarApp使用）
 * @param {string} calendarId - カレンダーID
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} { success, events, error, errorType, errorMessage }
 */
function getCalendarEvents(calendarId, startDate, endDate) {
  try {
    Logger.log('[getCalendarEvents] Fetching: ' + calendarId + ' from ' + startDate + ' to ' + endDate);

    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log('[getCalendarEvents] Calendar not found: ' + calendarId);
      return {
        success: false,
        events: [],
        errorType: CalendarErrorType.CALENDAR_NOT_FOUND,
        errorMessage: getCalendarErrorMessage(CalendarErrorType.CALENDAR_NOT_FOUND, calendarId)
      };
    }

    const start = new Date(startDate + 'T00:00:00');
    const end = new Date(endDate + 'T23:59:59');

    const events = calendar.getEvents(start, end);
    Logger.log('[getCalendarEvents] Found ' + events.length + ' events');

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

    return {
      success: true,
      events: mappedEvents,
      errorType: null,
      errorMessage: null
    };
  } catch (e) {
    Logger.log('[getCalendarEvents] Error: ' + e.message);
    const errorType = classifyCalendarError(e);
    return {
      success: false,
      events: [],
      errorType: errorType,
      errorMessage: getCalendarErrorMessage(errorType, calendarId),
      rawError: e.message
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
        errorMessage: `ICS取得失敗 (HTTP ${responseCode})`
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
      errorMessage: 'ICSファイルの取得に失敗しました: ' + e.message
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

  // TSCはCalendarAppでのみ取得可能（公開ICSなし）
  // 失敗時は詳細なエラーメッセージを返す
  if (!result.success) {
    result.errorMessage = `TSCカレンダー (${CONFIG.CALENDARS.TSC}) へのアクセスに失敗しました。\n\n` +
      `【対処方法】\n` +
      `1. Webアプリを再デプロイして、カレンダー権限を認可してください\n` +
      `2. デプロイしたユーザーがTSCカレンダーにアクセス権を持っている必要があります\n` +
      `3. カレンダー所有者に「閲覧者」として共有を依頼してください`;
  }

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
 * カレンダー接続テスト
 * @returns {Object} テスト結果
 */
function testCalendarConnection() {
  const results = {};

  // TSCカレンダー
  try {
    const tscCal = CalendarApp.getCalendarById(CONFIG.CALENDARS.TSC);
    results.TSC = {
      success: !!tscCal,
      name: tscCal ? tscCal.getName() : null,
      id: CONFIG.CALENDARS.TSC,
      method: 'CalendarApp'
    };
  } catch (e) {
    results.TSC = {
      success: false,
      error: e.message,
      errorType: classifyCalendarError(e),
      id: CONFIG.CALENDARS.TSC,
      method: 'CalendarApp'
    };
  }

  // 電気カレンダー（CalendarApp）
  try {
    const elecCal = CalendarApp.getCalendarById(CONFIG.CALENDARS.ELECTRICAL);
    results.ELECTRICAL_CalendarApp = {
      success: !!elecCal,
      name: elecCal ? elecCal.getName() : null,
      id: CONFIG.CALENDARS.ELECTRICAL,
      method: 'CalendarApp'
    };
  } catch (e) {
    results.ELECTRICAL_CalendarApp = {
      success: false,
      error: e.message,
      errorType: classifyCalendarError(e),
      id: CONFIG.CALENDARS.ELECTRICAL,
      method: 'CalendarApp'
    };
  }

  // 電気カレンダー（ICSフォールバック）
  try {
    const response = UrlFetchApp.fetch(ELECTRICAL_ICS_URL, { muteHttpExceptions: true });
    const code = response.getResponseCode();
    results.ELECTRICAL_ICS = {
      success: code === 200,
      httpStatus: code,
      url: ELECTRICAL_ICS_URL,
      method: 'ICS'
    };
  } catch (e) {
    results.ELECTRICAL_ICS = {
      success: false,
      error: e.message,
      url: ELECTRICAL_ICS_URL,
      method: 'ICS'
    };
  }

  Logger.log('Calendar connection test results: ' + JSON.stringify(results, null, 2));
  return results;
}
