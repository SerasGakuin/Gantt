//このファイルの用途調査中の関数は、recordFunctionUsage()から使用報告が2027/06/01ぐらいまでなければ、コメントアウトされ、アーカイブされます。
/**
 * @deprecated 用途調査中の関数。
 */
function englishtime() {
  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 2027/06/01ぐらいまで使用が確認されなければ削除予定
  var gantttt = SpreadsheetApp.getActive().getSheetByName("ガントチャート")
  var ganttttarow = gantttt.getRange("B:B").getValues()
  var englize;
  for (var t = 0; t < ganttttarow.length; t++) {
    if (ganttttarow[t][0] == "英語(時間）") {
      englize = t + 1
      break
    }

  }

  Logger.log(englize)
  return englize
}

/**
 * @deprecated 用途調査中の関数。
 */
function dlow() {
  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 2027/06/01ぐらいまで使用が確認されなければ削除予定
  var imaster = SpreadsheetApp.getActive().getSheetByName("個人マスター")
  var sankousho = imaster.getRange("D:D").getDisplayValues()
  var id = imaster.getRange("A:A").getDisplayValues()
  var ganted = SpreadsheetApp.getActive().getSheetByName("ガントチャート")
  var gantedvalues = ganted.getRange("C:C").getValues()
  var gyou = [];
  var newgyou = [];
  for (var i = 1; i < sankousho.length; i++) {
    for (var j = 0; j < gantedvalues.length; j++) {
      if (sankousho[i][0] == gantedvalues[j][0]) {
        gyou.push(j + 1)
        newgyou.push(i + 1)
      }
    }
  }
  for (var k = 0; k < gyou.length; k++) {
    ganted.getRange(gyou[k], 1).setValue(id[newgyou[k] - 1][0])
  }
}

/**
 * @deprecated 用途調査中の関数。
 */
function makenumber() {
  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 2027/06/01ぐらいまで使用が確認されなければ削除予定
  const transpose = a => a[0].map((_, c) => a.map(r => r[c]));
  var ganttsheet = SpreadsheetApp.getActive().getSheetByName("ガントチャート")
  var getupon = ganttsheet.getRange(1, 1, 200, 3).getValues()
  var transget = transpose(getupon)
  var getA = transpose([transget[0]])
  var getB = transpose([transget[1]])
  var getC = transpose([transget[2]])
  var individual = SpreadsheetApp.getActive().getSheetByName("個人マスター")
  var individuallast = individual.getLastRow()
  var du = []
  var er = []
  var es = []
  var numbercolumn = []
  var english = englishtime()
  for (i = 1; i < english - 1; i++) {
    if (getB[i][0] != null && getB[i][0] != "") {
      numbercolumn.push(i + 1)
    }
  }
  numbercolumn.splice(-4, 4)
  var newNumberk = null
  var newStringk = null
  var numberrow = transpose([numbercolumn])
  for (var j = 0; j < numberrow.length - 1; j++) {
    if (numberrow[j + 1][0] - numberrow[j][0] != 1) {
      var array = getA.slice(numberrow[j][0] - 1, numberrow[j + 1][0])
      var arrays = getC.slice(numberrow[j][0] - 1, numberrow[j + 1][0])
      var newarray = array.flat()
      Logger.log(array)
      var maxnumber = Math.max.apply(null, newarray.map(function (element) {
        return Number(element.slice(-2));
      }));
      var maxString = newarray.find(function (element) {
        return Number(element.slice(-2)) === maxnumber;
      });
      Logger.log(newarray)
      Logger.log(maxnumber)
      for (var k = 0; k < array.length; k++) {
        Logger.log(array[k][0])
        Logger.log(arrays[k][0])

        if (array[k][0] == null || array[k][0] == "") {
          if (arrays[k][0] != null && arrays[k][0] != "") {
            newNumberk = maxnumber + k
            newStringk = maxString.slice(0, -2) + newNumberk.toString().padStart(2, '0')
            ganttsheet.getRange(numberrow[j][0] + k, 1).setValue(newStringk)
            Logger.log(numberrow[j][0] + k)
            du.push(newStringk)
            er.push(getC[numberrow[j][0] + k - 1][0])
            es.push("")
            Logger.log(newNumberk)
            Logger.log(newStringk)
          }
        }
      }
    }
  }
  var setrow = [du, es, es, er]
  var transetrow = transpose(setrow)
  individual.getRange(individuallast + 1, 1, transetrow.length, 4).setValues(transetrow)

}