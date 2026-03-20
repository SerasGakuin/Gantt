
function onOpen() {
  // スプレッドシートのメニューバーに関数を配置
  var ui = SpreadsheetApp.getUi()
  
  //※※※※※ 独自メニューその1 ※※※※※
  //メニュー名を決定
  var menu = ui.createMenu("週間計画を作成");
  
  //メニューに実行ボタン名と関数を割り当て
  menu.addItem("第2週","week2Plan");
  menu.addItem("第3週","week3Plan");
  menu.addItem("第4週","week4Plan");
  menu.addItem("第5週","week5Plan");
  
  //スプレッドシートに反映
  menu.addToUi();

  //※※※※※ 独自メニューその2 ※※※※※
  //メニュー名を決定
  var menu = ui.createMenu("ヒアリング項目作成");
  
  //メニューに実行ボタン名と関数を割り当て
  menu.addItem("生成(GPT3.5)","chatGpt3");
  menu.addItem("生成(GPT4)","chatGpt4");
  //スプレッドシートに反映
  menu.addToUi();

}

function week2Plan() {
  var columnLetter = "P" // 「目標・実績」列
  addNumbersList("週間管理", columnLetter)
}
function week3Plan() {
  var columnLetter = "X"
  addNumbersList("週間管理", columnLetter)
}
function week4Plan() {
  var columnLetter = "AF"
  addNumbersList("週間管理", columnLetter)
}
function week5Plan() {
  var columnLetter = "AN"
  addNumbersList("週間管理", columnLetter)
}

// ======================================================================================================================================

function addNumbersList(sheetName, columnLetter) {
  /**
   * 週間計画を自動で作成する関数
   * sheetName: 「週間計画」シート
   * 
   */

  // 「習慣管理」シート
  let weeklySheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  
  // "AF"列を番号に変換し、左右の列番号も取得
  var columnNumber = ExcelUtils.convertAlphabetColumnToNumeric(columnLetter); // 「目標・実績」列
  var leftColumnNumber = columnNumber - 1; // 「目安処理量」列 
  var rightColumnNumber = columnNumber + 1; // 「最新の学習実績」列

  // 参考書が書いてある行の範囲
  var startRow = 4;
  var endRow = 30;

  // 目安処理量を取得
  var addRange = weeklySheet.getRange(startRow, leftColumnNumber, endRow - startRow + 1, 1);
  let valuesToAdd = addRange.getDisplayValues();
  valuesToAdd = transpose(valuesToAdd)[0];
  console.log(valuesToAdd)

  // 先週の実績を取得
  var inputRange = weeklySheet.getRange(startRow, rightColumnNumber, endRow - startRow + 1, 1);
  let weeklyLog = inputRange.getDisplayValues();
  weeklyLog = transpose(weeklyLog)[0];

  // 参考書IDを取得 (1列目)
  let referenceBookIDs = weeklySheet.getRange(startRow, 1, endRow - startRow + 1, 1).getDisplayValues();
  referenceBookIDs = transpose(referenceBookIDs)[0];

  // 最新の計画を出力範囲として設定
  var outputRange = weeklySheet.getRange(startRow, columnNumber, endRow - startRow + 1, 1);
  let weeklyPlan = outputRange.getDisplayValues();

  // 既存データがある場合の確認処理
  if (!isArrayEmpty(weeklyPlan)) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Continue?',
      'この週にはすでに入力された文字列があります。続けますか?',
      ui.ButtonSet.YES_NO
    );

    if (response == ui.Button.NO) {
      return;
    }
  }

  // 「参考書マスター」シートから、該当の参考書の情報を取得する
  var referenceBookAllData = getReferenceBookAllData()
  
  // メイン処理
  resultArray = []
  for (var i = 0; i<referenceBookIDs.length; i++) {

    // 当該参考書のデータをフィルタして取得
    var referenceBookID = String(referenceBookIDs[i]);
    referenceBookID = extractLettersAndNumbers(referenceBookID)
    console.log(referenceBookID)
    var referenceBookData = getReferenceBookData(referenceBookAllData, referenceBookID);

    if (!referenceBookData) {referenceBookData = "参考書IDが存在しません"}

    // 週間計画を作成
    var result = addNumbers(weeklyLog[i], valuesToAdd[i], referenceBookID, referenceBookData)
    console.log(result)
    resultArray.push(result)
  }

  // 出力
  resultArray = transpose([resultArray]);
  outputRange.setValues(resultArray);
}
