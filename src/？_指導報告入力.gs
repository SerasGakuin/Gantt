//このファイルの用途調査中の関数は、recordFunctionUsage()から使用報告が2027/06/01ぐらいまでなければ、コメントアウトされ、アーカイブされます。
/**
 * @deprecated 用途調査中の関数。
 */
function setreport() {
  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 2027/06/01ぐらいまで使用が確認されなければ削除予定
  var id = SpreadsheetApp.getActiveSpreadsheet().getId()
  var comiru = SpreadsheetApp.openById("1boXl7xnE0RBUHNyYDlWOzmOvnYgmMhlv8w3q3lcuRSQ")
  var comiruid = comiru.getSheetByName("生徒番号対スプシID対応表").getRange(2, 1, comiru.getSheetByName("生徒番号対スプシID対応表").getLastRow() - 1, 3).getValues()
  var sidoucontent = comiru.getSheetByName("最新指導報告ログ").getRange(2, 1, comiru.getSheetByName("最新指導報告ログ").getLastRow() - 1, 3).getValues()
  for (i = 0; i < comiruid.length; i++) {
    if (comiruid[i][2] == id) {
      for (j = 0; j < sidoucontent.length; j++) {
        if (comiruid[i][0] == sidoucontent[j][0]) {
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ヒアリング").getRange(1, 2).setValue(sidoucontent[j][2])
          break
        }
      }
      break
    }
  }
}
