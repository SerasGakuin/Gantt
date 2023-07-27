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
  let inputRange = "Q4:Q19"
  let addRange = "O4:O19"
  let ourputRange = "P4:P19"
  addNumbersList("週間管理", inputRange, addRange, ourputRange)
}
function week3Plan() {
  let inputRange = "Y4:Y19"
  let addRange = "W4:W19"
  let ourputRange = "X4:X19"
  addNumbersList("週間管理", inputRange, addRange, ourputRange)
}
function week4Plan() {
  let inputRange = "AG4:AG19"
  let addRange = "AE4:AE19"
  let ourputRange = "AF4:AF19"
  addNumbersList("週間管理", inputRange, addRange, ourputRange)
}
function week5Plan() {
  let inputRange = "AO4:AO19"
  let addRange = "AM4:AM19"
  let ourputRange = "AN4:AN19"
  addNumbersList("週間管理", inputRange, addRange, ourputRange)
}

// UTILITY FUNCTIONS ===========================================================================================================================================

const transpose = a=> a[0].map((_, c) => a.map(r => r[c]));

function isArrayEmpty(array) {
  // 2次元配列の全ての要素をチェックし、空の配列であることを確かめる
  for(var i = 0; i < array.length; i++) {
    for(var j = 0; j < array[i].length; j++) {
      // 要素が空でない場合、falseを返す
      if(array[i][j]) {
        return false;
      }
    }
  }
  // 全ての要素が空の場合、trueを返す
  return true;
}

function detectDoubleTilde(inputString) {
  // 文字列の中に、〜が2つ以上含まれているかをチェック
  var pattern = /~/g;
  var matches = inputString.match(pattern);
  if (matches && matches.length >= 2) {
    return true;
  } else {
    return false;
  }
}

function convertNumbers(inputString) {
  // 全角数字（０-９）を半角数字（0-9）に変換
  var outputString = inputString.replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  
  return outputString;
}

function doesNotContainNumbers(inputString) {
  // テキスト中に数字が含まれていないことを検出する
  var regex = /\d/;
  
  return !regex.test(inputString);
}

function convertTildas(inputString) {
  // 全角チルダ（～）、全角波ダッシュ（〜）、半角波ダッシュ（∼）を半角チルダ（~）に変換
  var outputString = inputString.replace(/[～〜∼]/g, '~');
  
  return outputString;
}

// ======================================================================================================================================

function addNumbers(inputString, numToAdd) {

  numToAdd = Math.round(Number(numToAdd))

  // 今週やらない場合は空欄を返す
  if (numToAdd==0) {
    return ""
  }

  // 先週やっていなかったら★を返す
  if (!inputString) {
    return "★";
  }

  // 全角の数字を半角の数字になおす
  inputString = convertNumbers(inputString);

  // 数字が一切含まれていなかったら★を返す
  if (doesNotContainNumbers(inputString)) {
    return "★";
  }

  // 全角の〜を半角の~に変換する
  inputString = convertTildas(inputString);

  // 複数の~が含まれる場合は★を返す
  if (detectDoubleTilde(inputString)) {
    return "★"
  }

  console.log(inputString)

  // ~で囲まれた数字を検出
  var pattern = /(\d+)~(\d+)/;
  var match = inputString.match(pattern);
  if (!match) {
    
    // ~がないのに複数の箇所に数字が含まれる場合は、★を返す
    var matches = inputString.match(/\d+/g);
    if (matches && matches.length >= 2) {
      return "★";
    }

    var pattern = /(\d+)/;
    var match = inputString.match(pattern);
    var endNumber = parseInt(match[1]);
  }else{
    var endNumber = parseInt(match[2]);
  }
  if (match) {
    var newStartNumber = endNumber + 1;

    if (numToAdd > 1) {
      var newEndNumber = newStartNumber + numToAdd - 1;
      var outputString = inputString.replace(pattern, newStartNumber + "~" + newEndNumber);
    } 
    else {
      var outputString = inputString.replace(pattern, newStartNumber)
    }
    return outputString;

  } else {
    return "★";;
  }
}

function addNumbersList(sheetName, inputRange, addRange, outputRange) {
  // 週間管理シートの情報を取得
  let weeklySheet = SpreadsheetApp.getActive().getSheetByName(sheetName)
  let weeklyLog = weeklySheet.getRange(inputRange).getDisplayValues()
  let valuesToAdd = weeklySheet.getRange(addRange).getDisplayValues()

  // すでに入力がある場合は、確認のダイアログを表示
  let weeklyPlan = weeklySheet.getRange(outputRange).getDisplayValues();

  if (!isArrayEmpty(weeklyPlan)) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Continue?',  // ダイアログのタイトル
      'この週にはすでに入力された文字列があります。続けますか?',  // ダイアログのメッセージ
      ui.ButtonSet.YES_NO  // ボタンのセット
    );
    
    // ユーザーが「いいえ」を選択した場合、関数を終了
    if (response == ui.Button.NO) {
      return;
    }
  }
  // 1次元配列に変換
  weeklyLog = transpose(weeklyLog)[0]
  valuesToAdd = transpose(valuesToAdd)[0]

  var result = weeklyLog.map(function(num, index) {
    return addNumbers(num, valuesToAdd[index]);
  });

  result = transpose([result]);
  console.log(result);
  weeklySheet.getRange(outputRange).setValues(result)
}


function test(){
  console.log(addNumbers("問題21∼24", 3));
  console.log(addNumbers("No.612~651", 1));
  console.log(addNumbers("第2回", 3));
  console.log(addNumbers("第2回", 1));

  let inputString = "１．速度と加速度～５．エネルギーの問題1~65のうち、*1個の問題"
  console.log(addNumbers(inputString, 18))
}


