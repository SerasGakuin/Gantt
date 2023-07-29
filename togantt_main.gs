function test() {
  var ganttChart = new GanttChart(SpreadsheetApp.openById("1TMhJ9CTq0zXY1IaOfxraM_cAAMvwZn42TuUtwEB4KTw"));
  ganttChart.toGantt();
}

function fix() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var idArray = ss.getSheetByName("移管").getRange(1, 1, 19, 7).getValues();
  var spreadsheetIds = transpose(idArray)[1];
  var fromIds = transpose(idArray)[6].reverse();

  var folder = DriveApp.getFolderById("1wJXnHda8183QMPqm058mkxFcHbq-9oER");

  for(let i=0;i<spreadsheetIds.length;i++){
    ssTo = SpreadsheetApp.openById(spreadsheetIds[i])
    ssFrom = SpreadsheetApp.openById(fromIds[i])
    if(!ssFrom.getSheetByName("月初")){
      console.log(ssTo.getName())
      continue
    }
    

    ssTo.getSheetByName("月初").getRange("BK2:BO19").setValues(ssFrom.getSheetByName("月初").getRange("BK2:BO19").getValues())
  }

}

//任意のフォルダにファイルを移動させるサンプルコード
function moveFile(fileId) {
  
  let folder = DriveApp.getFolderById("1Xz-dUPc3iNU7BSTF80Z6AsZ42iwv3eOj"); //任意のフォルダを指定してください
  let file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
}

function copySpreadsheets() {
  var newSpreadsheetIds = [];  // 新しいスプレッドシートのIDを格納するための配列を初期化します。

  var folder = DriveApp.getFolderById("1wJXnHda8183QMPqm058mkxFcHbq-9oER");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var idArray = ss.getSheetByName("移管").getRange(1, 1, 20, 3).getValues();
  var spreadsheetIds = transpose(idArray)[1];
  // すべてのスプレッドシートIDに対してループを回します。
  for(var i = 0; i < spreadsheetIds.length; i++) {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetIds[i]);  // スプレッドシートを開きます。
    var copy = DriveApp.getFileById(spreadsheet.getId()).makeCopy();  // スプレッドシートのコピーを作成します。
    DriveApp.getFileById(copy.getId()).moveTo(folder);
    
    //newSpreadsheetIds.push(copy.getId());  // 新しく作成したスプレッドシートのIDを配列に追加します。
  }
  
  //ss.getSheetByName("移管").getRange(1, 4, 20, 1).setValues(transpose([newSpreadsheetIds]))
  // 新しく作成したスプレッドシートのIDの配列を返します。
}



function main() {
  let startTime = new Date(); // ①実行開始時点の日時

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var idArray = ss.getSheetByName("移管").getRange(1, 1, 19, 5).getValues();


  let startIndex = Number(PropertiesService.getScriptProperties().getProperty('nextIndex'));
  if (!startIndex) startIndex = 0; // もしstartIndexがnullの場合は1を代入

  for (let i=startIndex; i < idArray.length; i++) {

    let currentTime = new Date(); // ②ループx周目時点の日時
    let seconds = (currentTime - startTime)/1000; // 経過秒数を計算(①と②の差分)

    if(seconds > 240){
      // 300秒経過したら、スクリプトプロパティを設定し、トリガーをセットして、returnする
      PropertiesService.getScriptProperties().setProperty('nextIndex', i);
      setTrigger();
      return;
    }

    var ssId = idArray[i][1];
    //var ssId = idArray[i][3];
    var excelId = idArray[i][2];

    console.log(idArray[i])

    if (excelId && !SpreadsheetApp.openById(ssId).getSheetByName("今月プラン")) {
      copySheetsFromExcelToSpreadsheet(ssId, excelId)
      reDoFormulas(ssId)
    }
    var ganttChart = new GanttChart(SpreadsheetApp.openById(ssId));
    ganttChart.toGantt();

    ss.getSheetByName("移管").getRange(i+1, 5, 1, 1).setValue("done");
  }
  // 500周し終えたらトリガーを削除
  let triggers = ScriptApp.getScriptTriggers();
  for(let trigger of triggers){
    if(trigger.getHandlerFunction() == 'main'){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 500周し終えたらスクリプトプロパティを削除
  PropertiesService.getScriptProperties().deleteProperty('nextIndex');  
}

function setTrigger() {
  let triggers = ScriptApp.getScriptTriggers();
  for(let trigger of triggers){
    if(trigger.getHandlerFunction() == 'main'){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 1分後にトリガーをセット(1分 = 60秒 = 1秒*60 = 1000ミリ秒 * 60)
  ScriptApp.newTrigger('main').timeBased().after(1000 * 60).create();
}

function listFilesInFolder() {

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.insertSheet("移管")

  var folder = DriveApp.getFolderById('1Rhsh7-yKrI6hP39lNGo_WjPbUiWeHf1M');
  var files = folder.getFiles();

  var i = 1;
  while (files.hasNext()) {
    var file = files.next();
    sheet.getRange(i, 1, 1, 2).setValues([[file.getName(), file.getId()]])
    i++
  }
  var folder = DriveApp.getFolderById("1dZwkgL4kP1sZu3VdrV2R-iXk2iKjkE5O");
  var files = folder.getFiles();
  i = 1
  while (files.hasNext()) {
    var file = files.next();
    sheet.getRange(i, 3, 1, 2).setValues([[file.getName(), file.getId()]])
    i++
  }
}


function copySheetsFromExcelToSpreadsheet(ssId, excelId) {
  // 操作するExcelファイルとスプレッドシートのIDを指定
  var excelFileId = excelId;  // ExcelファイルのID
  var targetSpreadsheetId = ssId;  // 対象のスプレッドシートのID

  // ExcelファイルをGoogleスプレッドシートに変換
  var excelFile = DriveApp.getFileById(excelFileId);
  var convertedSheetID = Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS}, excelFile.getBlob()).id;
  var convertedSpreadsheet = SpreadsheetApp.openById(convertedSheetID);


  // 変換したスプレッドシートの各シートを対象のスプレッドシートにコピー
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

  if(targetSpreadsheet.getSheetByName("今月プラン")) { 
    // 変換したスプレッドシートを削除
    DriveApp.getFileById(convertedSpreadsheet.getId()).setTrashed(true);
    return 
  }
  var sheets = convertedSpreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var newSheet = sheets[i].copyTo(targetSpreadsheet);
    newSheet.setName(sheets[i].getName());
  }

  // 変換したスプレッドシートを削除
  DriveApp.getFileById(convertedSpreadsheet.getId()).setTrashed(true);
}


function reDoFormulas(ssId) {

  var ss = SpreadsheetApp.openById(ssId);
  // すべてのシートオブジェクトを取得する
  const sheets = ss.getSheets();

  // 1シートずつループする
  for (let i = 0; i < sheets.length; i++) {

    // シートにデータがない場合はスキップする
    if(sheets[i].getLastRow() == 0){
      continue;
    }

    // シート内のデータが存在する範囲を取得
    let range = sheets[i].getRange(1, 1, sheets[i].getLastRow(), sheets[i].getLastColumn());

    // 範囲内の数式を取得
    let formulas = range.getFormulas();

    // 1セルずつ数式を再セットしていく
    // setFormulasでまとめてセットすると数式エラーが出た場合に該当セルがわからないので1セルずつセットする
    for (let j = 0; j < formulas.length; j++) { // 行のループ

      for (let k = 0; k < formulas[j].length; k++) { // 列のループ

        if (formulas[j][k] != '') { // 数式としてなにか文字列が入っていれば

          range = sheets[i].getRange(j + 1, k + 1); // 該当セルの範囲オブジェクトを取得
          range.setFormula(formulas[j][k]); // 数式をセットする

        }
      }
    }
  }
}