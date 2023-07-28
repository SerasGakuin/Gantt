function getStartRow(ss, alreadyDeletedColumns) {
  /**
   * 「英単語・英熟語」と書いてある行を検索し、行数を返す関数。
   */

  //ssが指定されていないとき：現在のシート
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  gantt = ss.getSheetByName("ガントチャート");
  // ID列が１列か３列かによって、検索する列を変更
  if (alreadyDeletedColumns) {
    var range = "B:B"
  } else {
    var range = "D:D"
  }

  var values = gantt.getRange(range).getValues();
  var rownumber = null

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == "英単語・英熟語") {
      rownumber = i + 1
      break
    }
  }
  return rownumber;
}

function getRowsWithEnglishTime(ss, alreadyDeletedColumns) {
  /**
   * 英語(時間)と書いてある行を検索し、行数を返す関数。
   */

  //ssが指定されていないとき：現在のシート
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  gantt = ss.getSheetByName("ガントチャート");
  // ID列が１列か３列かによって、検索する列を変更
  if (alreadyDeletedColumns) {
    var range = "B:B"
  } else {
    var range = "D:D"
  }

  var values = gantt.getRange(range).getValues();
  var rownumber = null

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == "英語(時間)") {
      rownumber = i + 1
      break
    }
    if (values[i][0] == "英語(時間）") {
      rownumber = i + 1
      break 
    }
  }
  
  // rowsWithEnglishTime を必要に応じて他の処理に使用することができます
  return rownumber;
}