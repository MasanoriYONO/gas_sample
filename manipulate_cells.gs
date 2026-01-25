/* ここのIDはURL欄
https://docs.google.com/spreadsheets/d/abcdefghijk123lmnopqrstuVwXyz456zymkiosample/edit
の /d/ から/edit の間を指定する。
*/
var entrylist_SheetID = 'set_google_spreadsheet_id_to_this_value_____';
// これでもIDを取得可能だけれども、複数のシートを扱う場合は
// アクティブなシートを切り替える必要があるため、他のシートも操作する場合にはお勧めできない。
//var entrylist_SheetID = SpreadsheetApp.getActiveSpreadsheet().getId();

var listSheetName = 'order_list';
var dest_sheets = SpreadsheetApp.openById(entrylist_SheetID);
var destSheet = dest_sheets.getSheetByName(listSheetName);
var last_row = 0;
var destSheetRange;

var faceSheetName = 'order_face';
var faceSheet = dest_sheets.getSheetByName(faceSheetName);

function makeOrder() {

  last_row = destSheet.getLastRow();
  Logger.log("last_row:" + last_row);
  destSheet.getRange(last_row + 1, 1).setFormula("=Form_Responses[タイムスタンプ]");
  destSheet.getRange(last_row + 1, 2).setFormula("=Form_Responses[書類名]");
  destSheet.getRange(last_row + 1, 3).setFormula("=Form_Responses[在庫数]");
  destSheet.getRange(last_row + 1, 4).setFormula("=xlookup(B" +  (last_row+1) + 
        ",needs!$A$2:$A$47,needs!$B$2:$B$47,\"not found.\",0,1)" );
  destSheet.getRange(last_row + 1, 5).setFormula("=C" + (last_row+1) 
        + "-D" + (last_row+1));
  destSheet.getRange(last_row + 1, 6).setFormula("=if(E" + (last_row+1) + "<0, ceiling(abs(E"
        + (last_row+1) +"),10),\"\")");
}

function setOrderFace() {
  last_row = destSheet.getLastRow();
  Logger.log("last_row:" + last_row);

  var order_range = destSheet.getRange(2, 1, last_row-1, 6);
  var face_row = 7;
  var face_column = 1;
  //一覧の内容のセルをクリア
  faceSheet.getRange(face_row, face_column,16,5).clearContent();

  var stock_check_date = Utilities.formatDate(faceSheet.getRange(2,5,1,1).getValue(),"Asia/Tokyo","yyyy-MM-dd");
  Logger.log("stock_check_date:" + stock_check_date);

  for(var row = 1; row <= order_range.getNumRows(); row++) {
    var input_date = new Date(order_range.getCell(row, 1).getValue());
    var input_date_modify = Utilities.formatDate(input_date,"Asia/Tokyo","yyyy-MM-dd");
      
    Logger.log("input_date_modify:" + input_date_modify);
    Logger.log("true or false:" + (input_date_modify == stock_check_date));

    if(input_date_modify == stock_check_date){
      var order_value = order_range.getCell(row, 6).getValue();
      if(order_value > 0){
        faceSheet.getRange(face_row, face_column).setValue(order_range.getCell(row, 2).getValue());
        faceSheet.getRange(face_row, face_column+2).setValue(order_range.getCell(row, 6).getValue());
        face_row++;
      }
    }else{
      continue;
    }
  }
}
