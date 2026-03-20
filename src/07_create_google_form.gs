//このファイルの用途調査中の関数は、recordFunctionUsage()から使用報告が2027/06/01ぐらいまでなければ、コメントアウトされ、アーカイブされます。
// 一部の関数について、recordFunctionUsage()の出力から、使用時のログが確認されました。

function blockRunning_createGoogleForm() { }//誤操作での実行を防止

//用途調査中の関数。
function week1form() { creategoogleform("1") }
function week2form() { creategoogleform("2") }
function week3form() { creategoogleform("3") }
function week4form() { creategoogleform("4") }
function week5form() { creategoogleform("5") }

//用途調査中の関数。
function creategoogleform(weeknumber) {
  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 2027/06/01ぐらいまで使用が確認されなければ削除予定

  if (!weeknumber) return;//誤実行防止のために追加しました

  deleteFormSubmitTriggers();
  var spreadsheetid = SpreadsheetApp.getActiveSpreadsheet().getId();

  var sankousyo = SpreadsheetApp.openById(spreadsheetid).getSheetByName("週間管理").getRange(4, 1, 16, 40).getDisplayValues()
  var sankousyocontents = []
  for (i = 0; i < sankousyo.length; i++) {
    if (sankousyo[i][0] != "") {
      if (weeknumber == "1" && sankousyo[i][7] != "") {
        sankousyocontents.push(sankousyo[i])
      }
      if (weeknumber == "2" && sankousyo[i][15] != "") {
        sankousyocontents.push(sankousyo[i])
      }
      if (weeknumber == "3" && sankousyo[i][23] != "") {
        sankousyocontents.push(sankousyo[i])
      }
      if (weeknumber == "4" && sankousyo[i][31] != "") {
        sankousyocontents.push(sankousyo[i])
      }
      if (weeknumber == "5" && sankousyo[i][39] != "") {
        sankousyocontents.push(sankousyo[i])
      }
    }
  }
  var oldstudentname = new StudentMasterLib.studentMaster().getStudentNameFromLinkOrId(spreadsheetid)
  var studentname = oldstudentname.replace(/[\s\u3000]+/g, '');
  Logger.log(studentname)
  var seitomaster = new StudentMasterLib.studentMaster().data
  Logger.log(seitomaster)
  for (i = 0; i < seitomaster.length; i++) {
    if (seitomaster[i].名前 == studentname) {
      var lineid = seitomaster[i].生徒LINEID
    }
  }
  Logger.log(lineid)
  var monthnumber = extractNumbersFromString(SpreadsheetApp.openById(spreadsheetid).getSheetByName("今月プラン").getRange(1, 4).getValue())[0]
  var copiedfile = DriveApp.getFileById("1C44f6LvKgcsu9KduqgpuXdLrrqxCV-ib7HB5snlSIXg").makeCopy()
  copiedfile.setName(studentname + "さんの" + monthnumber + "月の第" + weeknumber + "週の週間計画について")
  DriveApp.getFolderById("1glcwo_o6FHSgdmx-ioBejDX1hfkm7tCA").addFile(copiedfile)
  var copiedfileid = copiedfile.getId()
  const form = FormApp.openById(copiedfileid)
  form.setTitle(studentname + "さんの" + monthnumber + "月の第" + weeknumber + "週の週間計画について")
  form.getItems().forEach(function (item) {
    form.deleteItem(item);
  })
  for (j = 0; j < sankousyocontents.length; j++) {
    form.addMultipleChoiceItem().setTitle("参考書：" + sankousyocontents[j][2] + "の理解度は？")
      .setHelpText('ラジオボタンの質問項目に対する説明文')
      .setChoiceValues(['○', '△', '✕'])
      .setRequired(true);
  }
  ScriptApp.newTrigger("onFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create()
  var formUrl = form.getPublishedUrl();
  linesend.sendform(formUrl, lineid)
  // ヒアリング入れる？
}

//用途調査中の関数。
function extractNumbersFromString(inputString) {
  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 
  // 正規表現を使用して文字列内の全ての数字をマッチさせる
  var matches = inputString.match(/\d+/g);
  // マッチする数字があればそれを返し、なければ空の配列を返す
  return matches || [];
}


/**
 * recordFunctionUsage()の出力から、使用報告ログが確認されました。
 * "
 * 2026/01/26 1:34:29:
 * エラーではなく関数の用途調査記録です。
 * もしこのログを見かけましたら、このログをコピーしてどこかに張り付けて記録しておいていただけると幸いです。:
 * at recordFunctionUsage (01_LogUtils:22:24)
 * at deleteFormSubmitTriggers (?_create_google_form:104:3)
 * at __GS_INTERNAL_top_function_call__.gs:1:8
 * "
 */
function deleteFormSubmitTriggers() {
  // 現在設定されている全てのトリガーを取得
  const triggers = ScriptApp.getProjectTriggers();

  // 各トリガーをチェック
  triggers.forEach(function (trigger) {
    Logger.log(trigger.getEventType())
    // フォーム送信時のトリガーかどうかを確認
    if (trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
      // トリガーを削除
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('フォーム送信時のトリガーをすべて削除しました。');
}


/**
 * recordFunctionUsage()の出力から、使用報告ログが確認されました。
 * "
 * 2026/01/26 2:38:04:
 * ！！！！！！重要！！！！！！
 * エラーではなく関数の用途調査記録です。
 * もしこのログを見かけましたら、このログをコピーしてどこかに張り付けて記録しておいていただけると幸いです。:
 *  at recordFunctionUsage (01_LogUtils:22:24)
 *  at deleteallForms (?_create_google_form:125:3)
 *  at __GS_INTERNAL_top_function_call__.gs:1:8
 * "
 */
function deleteallForms() {
  var folderId = '1glcwo_o6FHSgdmx-ioBejDX1hfkm7tCA'; // 対象フォルダのIDを入力
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    Logger.log('Deleting file: ' + file.getName());
    file.setTrashed(true);
  }
  Logger.log('All files moved to trash.');
}

//用途調査中の関数。
function onFormSubmit(e) {
  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 2027/06/01ぐらいまで使用が確認されなければ削除予定

  if (!weeknumber) return;//誤実行防止のために追加しました

  const form = FormApp.openById(e.source.getId()); // フォームIDでフォームを取得
  const title = form.getTitle(); // フォームのタイトルを取得
  var name = extractSubstring(title)
  var number = extractNumbersFromString(title)
  var monthlynumber = Number(number[0])
  var weeklynumber = Number(number[1])
  const studentMaster = new StudentMasterLib.studentMaster();
  var studentsheeturl = studentMaster.getStudentDataByName(name.replace(/\s+/g, ''))?.スプレッドシート ?? '';
  var studentsheet = SpreadsheetApp.openByUrl(studentsheeturl);
  var nowmonth = Number(extractNumbersFromString(studentsheet.getSheetByName("今月プラン").getRange(1, 4).getValue())[0])
  var studentschedule = studentsheet.getSheetByName("週間管理").getRange(4, 1, 16, 40).getDisplayValues()
  var weektable = { "week1": 7, "week2": 15, "week3": 23, "week4": 31, "week5": 39 }
  var weekstring = "week" + weeklynumber.toString()
  if (monthlynumber == nowmonth) {
    var plansheet = studentsheet.getSheetByName("今月プラン");
    var plantime = plansheet.getRange(4, 6, 16, 5).getValues();
  } else {
    var plansheet = studentsheet.getSheetByName("今月実績");
    var plantime = plansheet.getRange(4, 5, 16, 5).getValues();
  }
  var responses = e.response.getItemResponses();
  var questions = []
  var notemptime = []
  for (var i = 0; i < responses.length; i++) {
    var itemResponse = responses[i];
    questions.push(itemResponse.getItem().getTitle())
  }
  for (var i = 0; i < plantime.length; i++) {
    if (plantime[i][weeklynumber - 1] != "") {
      notemptime.push(i)
    }
  }
  var changeschedule = []
  for (i = 0; i < notemptime.length; i++) {
    if (questions[i].includes("参考書") && responses[i].getResponse() == "○") {
      studentschedule[notemptime[i]][weektable[weekstring]] = "○" + studentschedule[notemptime[i]][weektable[weekstring]]
    }
    if (questions[i].includes("参考書") && responses[i].getResponse() == "△") {
      studentschedule[notemptime[i]][weektable[weekstring]] = "△" + studentschedule[notemptime[i]][weektable[weekstring]]
      plantime[notemptime[i]][weeklynumber - 1] = Number(plantime[notemptime[i]][weeklynumber - 1]) / 2;
    }
    if (questions[i].includes("参考書") && responses[i].getResponse() == "✕") {
      studentschedule[notemptime[i]][weektable[weekstring]] = "✕" + studentschedule[notemptime[i]][weektable[weekstring]]
      plantime[notemptime[i]][weeklynumber - 1] = 0
    }
  }
  for (i = 0; i < studentschedule.length; i++) {
    changeschedule.push([studentschedule[i][weektable[weekstring]]])
  }
  Logger.log(changeschedule)
  studentsheet.getSheetByName("週間管理").getRange(4, weektable[weekstring] + 1, 16, 1).setValues(changeschedule)
  if (monthlynumber == nowmonth) {
    studentsheet.getSheetByName("今月プラン").getRange(4, 6, 16, 5).setValues(plantime)
  } else {
    studentsheet.getSheetByName("今月実績").getRange(4, 5, 16, 5).setValues(plantime)
  }
}

//用途調査中の関数。
function extractSubstring(str) {
  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 2027/06/01ぐらいまで使用が確認されなければ削除予定

  // 「さんの」が最初に現れる位置を見つける
  var index = str.indexOf('さんの');

  // 「さんの」より前の部分を切り出す
  if (index !== -1) {
    var substring = str.substring(0, index);
    Logger.log(substring); // ログに出力する（Google Apps Scriptの場合）
    return substring; // 戻り値として返す
  } else {
    Logger.log("パターンが見つかりませんでした。");
    return null; // パターンが見つからない場合はnullを返す
  }
}

//用途調査中の関数。
function extractNumbersFromString(inputString) {
  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 2027/06/01ぐらいまで使用が確認されなければ削除予定
  // 正規表現を使用して文字列内の全ての数字をマッチさせる
  var matches = inputString.match(/\d+/g);
  // マッチする数字があればそれを返し、なければ空の配列を返す
  return matches || [];
}