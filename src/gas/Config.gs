/**
 * 設定ファイル
 * スプレッドシートのシート名やカラム定義を管理
 */

const CONFIG = {
  // スプレッドシートID（デプロイ時に設定）
  SPREADSHEET_ID: '1tBwMSYpWtt9ozLh8bd7CE68HM1mITSpzTvr6Y4uOHRM',

  // 外部スプレッドシートID（工番マスターの参照元）
  EXTERNAL_MASTER_SPREADSHEET_ID: '1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ',
  EXTERNAL_MASTER_SHEET_NAME: 'LW／作業日報_全従業員用',

  // 外部工番マスターの列定義（A〜E列、順序固定）
  EXTERNAL_JOB_MASTER_COLUMNS: {
    JOB_NO: 0,      // A列: 工番
    CUSTOMER: 1,    // B列: 受注先
    DESTINATION: 2, // C列: 納入先
    PRODUCT: 3,     // D列: 品名
    QUANTITY: 4     // E列: 数量
  },

  // シート名
  SHEETS: {
    JOBS: 'Jobs',
    PROCESS_MASTER: 'ProcessMaster',
    PEOPLE: 'People',
    SCHEDULE: 'Schedule',
    ATTACHMENTS: 'Attachments',
    TRIPS: 'Trips',
    EXTERNAL_JOB_MASTER: '外部_工番マスター',
    WORKER_JOB_ASSIGN: 'WorkerJobAssign',
    TOP_MEMO: 'TopMemo',
    JOB_PROCESS_LAYOUT: 'JobProcessLayout',
    DAY_SETTINGS: '04_DaySettings'
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
