var entrylist_SheetID = 'set_google_spreadsheet_id_to_this_value_____';
var listSheetName = 'データの整形';
var dest_sheets = SpreadsheetApp.openById(entrylist_SheetID);
var destSheet = dest_sheets.getSheetByName(listSheetName);
var last_row = 0;
var destSheetRange;
//GoogleカレンダーのID
var cal_id = 'set_calendar_id_to_this_value_set_calendar_id_to_this_value_1234@group.calendar.google.com';

function modifyEvent(e) {
  let calendar = CalendarApp.getCalendarById(cal_id);

  let [timestamp, date, time, title, input_title] = e.values;
  
  let y = Number(date.split('/')[0]);
  let m = Number(date.split('/')[1]) - 1;
  let d = Number(date.split('/')[2]);
  let start_hour = Number(time.split('〜')[0].split(':')[0]);
  let start_minute = Number(time.split('〜')[0].split(':')[1]);
  let finish_hour = Number(time.split('〜')[1].split(':')[0]);
  let finish_minute = Number(time.split('〜')[1].split(':')[1]);
  
  let startTime = new Date(y, m, d, start_hour, start_minute);
  let endTime = new Date(y, m, d, finish_hour, finish_minute);
  if(title == "** 手入力 **"){
    title = input_title;
  }
  last_row = destSheet.getLastRow();
  Logger.log("last_row:" + last_row);
  destSheet.getRange(last_row + 1, 1).setValue(startTime);
  destSheet.getRange(last_row + 1, 2).setValue(endTime);
  destSheet.getRange(last_row + 1, 3).setValue(title);
  let description = `入力日時: ${timestamp}`;
  let cal_location = '施設名サンプル';

  let options = {
    description: description,
    location: cal_location,
  };
  calendar.createEvent(title, startTime, endTime, options);
}
