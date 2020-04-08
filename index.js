/**
 * get google calendar id
 */
function getCalendarId () {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let val = sheet.getRange('C2').getValue();

  if (! val.length) {
    Browser.msgBox('カレンダ ID が指定されていません。');
    return;
  }

  return val;
}

/**
 * create and register schedules into google calendar
 */
function createSchedule() {
  // credential 取得
  const GOOGLE_CALENDAR_ID = getCalendarId();

  // sheet からの取得データは配列. index を指定.
  const COL_DATE        = 0;
  const COL_DAY         = 1;
  const COL_START_TIME  = 3;
  const COL_END_TIME    = 5;
  const COL_TITLE       = 8;
  const COL_DESCRIPTION = 9;

  // get schedule list from shee
  const FIRST_ROW = 6;
  const FIRST_COLUMN = 2;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let valueList = sheet
    .getRange(FIRST_ROW, FIRST_COLUMN, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();

  // register schedules to google calendar
  let calendar = CalendarApp.getCalendarById(GOOGLE_CALENDAR_ID);
  valueList.forEach(value => {
    // 空行はスキップ
    let cntEmptyCellList = value.filter(cell => {
      return cell == '';
    }).length;
    if (value.length == cntEmptyCellList) {
      return true;
    }

    // 休み 登録
    let day = new Date(value[COL_DATE]);
    if (value[COL_TITLE] == '休み') {
      let title   = value[COL_TITLE];
      let options = {
        description: value[COL_DESCRIPTION]
      };
      try {
        let event = calendar.createAllDayEvent(
          title,
          day,
          options
        );
        event.setColor(CalendarApp.EventColor.RED);
      } catch(e) {
        Logger.log(e);
        Browser.msgBox('カレンダへデータ更新中にエラーが発生しました。例のあの人に相談して下さい。');
      }

      return true;
    }

    // 通常勤務日 登録
    let startTime = value[COL_START_TIME];
    let endTime   = value[COL_END_TIME];

    let startDate = new Date(day);
    startDate.setHours(startTime.getHours());
    startDate.setMinutes(startTime.getMinutes());

    let endDate = new Date(day);
    endDate.setHours(endTime.getHours())
    endDate.setMinutes(endTime.getMinutes());

    let title   = value[COL_TITLE];
    let options = {
      description: value[COL_DESCRIPTION]
    };
    try {
      calendar.createEvent(
        title,
        startDate,
        endDate,
        options
      );
    } catch(e) {
      Logger.log(e);
      Browser.msgBox('カレンダへデータ更新中にエラーが発生しました。例のあの人に相談して下さい。');
    }

  });

  // ブラウザへ完了通知
  Browser.msgBox("カレンダへ予定を登録しました!\\n共有したカレンダを開いて確認して下さい。\\nバグがあるかもしれないから登録されたデータが正しいかちゃんと確認してね。");
}

/**
 * create calendar title
 */
function createTitle () {
  // sheet から取得したデータは配列なので, 対象列を 0 始まりで指定
  const COL_DATE        = 0;
  const COL_DAY         = 1;
  const COL_WORK_STYLE  = 2;
  const COL_START_TIME  = 3;
  const COL_END_TIME    = 5;
  const COL_TITLE       = 8;
  const COL_DESCRIPTION = 9;

  //予定の一覧を配列で取得
  const FIRST_ROW = 6;
  const FIRST_COLUMN = 2;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let valueList = sheet
    .getRange(FIRST_ROW, FIRST_COLUMN, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();

  // 生成
  const TITLE_HOLIDAY = '休み';
  let titleList = [];
  valueList.forEach(value => {
    let workStyle = value[COL_WORK_STYLE];
    let startTime = value[COL_START_TIME];
    let endTime   = value[COL_END_TIME];

    // 空行 は '' をセット
    if (!workStyle && !startTime && !endTime) {
      titleList.push(['']);
      return true;
    }

    // 休日セット
    if (workStyle == '休' || (!startTime && !endTime)) {
      titleList.push([TITLE_HOLIDAY]);
      return true;
    }

    // 勤務日シフト セット
    // 早・中・遅 判定
    let prefix = '';
    if (startTime.getHours() < 10) {
      prefix = '早';
    } else if (startTime.getHours() >= 12) {
      prefix = '遅';
    } else {
      prefix = '中';
    }
    // short/long セット
    let workingTime = (endTime.getTime() - startTime.getTime()) / (1000 * 60 * 60);
    let postfix = '';
    if (workingTime < 9) {
      postfix = 'short';
    } else if (workingTime > 9) {
      postfix = 'long';
    } else {
      postfix = '';
    }
    titleList.push([prefix + postfix]);
  });

  // sheet に書き出し
  const TITLE_COLUMN = 10;
  sheet.getRange(FIRST_ROW, TITLE_COLUMN, titleList.length, 1).setValues(titleList);

  // ブラウザへ完了通知
  Browser.msgBox("タイトルを作成しました!");
}

/**
 * delete registered calendar data
 */
function deleteRegisteredScheduleList () {
  const GOOGLE_CALENDAR_ID = getCalendarId();

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dt = sheet.getRange('B6').getValue();
  let startTime = new Date(dt);
  let endTime = new Date(startTime.getFullYear(), startTime.getMonth() + 1, 0); // 月末
  endTime.setDate(endTime.getDate() + 1);
  let calendar = CalendarApp.getCalendarById(GOOGLE_CALENDAR_ID);
  let eventList = calendar.getEvents(startTime, endTime);
  eventList.forEach(event => {
    event.deleteEvent();
  });
}

/**
 * initialized
 */
function init () {
  // カレンダ登録済みデータ 削除
  deleteRegisteredScheduleList();
}

/**
 * entry point
 */
function main() {
  init();
  createSchedule();
}

