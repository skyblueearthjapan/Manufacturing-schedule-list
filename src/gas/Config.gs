/**
 * 設定ファイル
 * スプレッドシートのシート名やカラム定義を管理
 */

const CONFIG = {
  // スプレッドシートID（デプロイ時に設定）
  SPREADSHEET_ID: '',

  // シート名
  SHEETS: {
    JOBS: 'Jobs',
    PROCESS_MASTER: 'ProcessMaster',
    PEOPLE: 'People',
    SCHEDULE: 'Schedule',
    ATTACHMENTS: 'Attachments',
    TRIPS: 'Trips'
  },

  // デフォルト表示日数
  DEFAULT_DISPLAY_DAYS: 70,

  // ステータス定義
  STATUS: {
    SCHEDULE: ['予定', '進行中', '完了', '遅延'],
    JOB: ['未着手', '進行中', '完了', '保留']
  },

  // 工程種別
  PROCESS_TYPE: {
    RANGE: 'range',
    MILESTONE: 'milestone'
  }
};

/**
 * 現在のスプレッドシートを取得
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSpreadsheet() {
  let ss = null;

  if (CONFIG.SPREADSHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } catch (e) {
      throw new Error(`スプレッドシートを開けません (ID: ${CONFIG.SPREADSHEET_ID}): ${e.message}`);
    }
  } else {
    // バインドされたスクリプトの場合
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error('スプレッドシートが見つかりません。Config.gsのSPREADSHEET_IDを設定するか、スプレッドシートにバインドしてください。');
  }

  return ss;
}

/**
 * 指定シートを取得
 * @param {string} sheetName - シート名
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート "${sheetName}" が見つかりません`);
  }
  return sheet;
}
