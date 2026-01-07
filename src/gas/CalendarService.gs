/**
 * CalendarService.gs
 * Google Calendar APIとの連携を担当
 */

/**
 * 指定カレンダーのイベントを取得
 * @param {string} calendarId - カレンダーID
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object[]} イベント配列
 */
function getCalendarEvents(calendarId, startDate, endDate) {
  try {
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log('[getCalendarEvents] Calendar not found: ' + calendarId);
      return [];
    }

    const start = new Date(startDate + 'T00:00:00');
    const end = new Date(endDate + 'T23:59:59');

    const events = calendar.getEvents(start, end);

    return events.map(event => ({
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
      creators: event.getCreators() || [],
      guestList: event.getGuestList().map(g => ({
        email: g.getEmail(),
        name: g.getName(),
        status: g.getGuestStatus().toString()
      }))
    }));
  } catch (e) {
    Logger.log('[getCalendarEvents] Error: ' + e.message);
    throw new Error('カレンダーイベントの取得に失敗しました: ' + e.message);
  }
}

/**
 * TSCカレンダーのイベントを取得
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object[]} イベント配列
 */
function getTSCCalendarEvents(startDate, endDate) {
  return getCalendarEvents(CONFIG.CALENDARS.TSC, startDate, endDate);
}

/**
 * 電気カレンダーのイベントを取得
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object[]} イベント配列
 */
function getElectricalCalendarEvents(startDate, endDate) {
  return getCalendarEvents(CONFIG.CALENDARS.ELECTRICAL, startDate, endDate);
}

/**
 * 複数カレンダーのイベントを一括取得
 * @param {string[]} calendarIds - カレンダーID配列
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 * @returns {Object} カレンダーIDをキーとしたイベント配列のマップ
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
      id: CONFIG.CALENDARS.TSC
    };
  } catch (e) {
    results.TSC = { success: false, error: e.message, id: CONFIG.CALENDARS.TSC };
  }

  // 電気カレンダー
  try {
    const elecCal = CalendarApp.getCalendarById(CONFIG.CALENDARS.ELECTRICAL);
    results.ELECTRICAL = {
      success: !!elecCal,
      name: elecCal ? elecCal.getName() : null,
      id: CONFIG.CALENDARS.ELECTRICAL
    };
  } catch (e) {
    results.ELECTRICAL = { success: false, error: e.message, id: CONFIG.CALENDARS.ELECTRICAL };
  }

  Logger.log('Calendar connection test results: ' + JSON.stringify(results, null, 2));
  return results;
}
