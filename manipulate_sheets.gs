var listSheetName = 'detail_1';
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var templateSheet = spreadsheet.getSheetByName(listSheetName);
const faceSheetURL = "https://docs.google.com/spreadsheets/d/set_google_spreadsheet_id_to_this_value_____/edit";
// 他のシートを参照する場合。
function copySheets() {
  var faceSheetRow = 6;
  for(var i=2; i <= 16; i++){
    var newSheetName = "detail_" + i.toString();
    var newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);
    newSheet.getRange(6,2).setFormula("=IMPORTRANGE(\"" + faceSheetURL + "\",\"order_face!a" 
      + (faceSheetRow+i).toString() + "\")");
    newSheet.getRange(7,2).setFormula("=IMPORTRANGE(\"" + faceSheetURL + "\",\"order_face!c" 
      + (faceSheetRow+i).toString() + "\")");
  }
  
}
// アクティブなシートに元データをコピーしたシートが存在する場合。
function copySheetsLookAtLocal() {
  var faceSheetRow = 6;
  for(var i=2; i <= 16; i++){
    var newSheetName = "detail_" + i.toString();
    var newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);
    newSheet.getRange(14,2).setFormula("=order_face!a" + (faceSheetRow+i).toString());
    newSheet.getRange(15,2).setFormula("=order_face!c" + (faceSheetRow+i).toString());
    newSheet.getRange(16,2).setFormula("=order_face!f" + (faceSheetRow+i).toString());
    newSheet.getRange(17,2).setFormula("=order_face!d" + (faceSheetRow+i).toString());
  }
  
}

// テンプレートのシート以外を一括削除する。
function removeSheets() {
  for(var i=2; i <= 16; i++){
    let delSheetName = "detail_" + i.toString();
    let delSheet = spreadsheet.getSheetByName(delSheetName);
    //deleteSheetメソッドでシート削除を実行
    spreadsheet.deleteSheet(delSheet);
  }
}
