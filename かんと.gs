function gesho() {
  const transpose = a=> a[0].map((_, c) => a.map(r => r[c]));
  // 週間管理シートの情報を取得
  var weeklySheet = SpreadsheetApp.getActive().getSheetByName("週間管理")
  
  // 月間管理シートの情報を取得
  var monthlySheet = SpreadsheetApp.getActive().getSheetByName("月間管理");
  var monthlyValues = monthlySheet.getRange("A:A").getValues(); // シート内の適切な範囲を指定
  var monthlyfirst = SpreadsheetApp.getActive().getSheetByName("月初");

  const value = weeklySheet.getRange("AR4:AV19").getDisplayValues()
  var originrow = monthlySheet.getRange("A:A").getDisplayValues()
  var number = weeklySheet.getRange("A4").getValue()
  monthlyfirst.getRange("BK4:BO19").setValues(value)

  // 週間管理シートの各行について処理
  for (var i = 0; i < monthlyValues.length; i++) {
    var row = originrow[i] 
    
    if (number == row) {
      monthlySheet.getRange(i + 1 ,14,16,5).setValues(value)
      break;
    }
  }
  
  weeklySheet.getRange("H4:H19").clearContent()
  weeklySheet.getRange("P4:P19").clearContent()
  weeklySheet.getRange("X4:X19").clearContent()
  weeklySheet.getRange("AF4:AF19").clearContent()
  weeklySheet.getRange("AN4:AN19").clearContent()
  
  var monthplan = SpreadsheetApp.getActive().getSheetByName("今月プラン")
  var monthachievement = SpreadsheetApp.getActive().getSheetByName("今月実績")

  let lastmonth = monthplan.getRange("A4:J19").getDisplayValues()
  monthachievement.getRange("D1").setValue(monthplan.getRange("D1").getDisplayValue())
  let lastarry = transpose(lastmonth)
  const lastID = transpose([lastarry[0]])
  const lasttime = transpose([lastarry[5],lastarry[6],lastarry[7],lastarry[8],lastarry[9]])

  monthachievement.getRange("A4:A19").setValues(lastID)
  monthachievement.getRange("E4:I19").setValues(lasttime)
  monthplan.getRange("A4:A19").clearContent()
  monthplan.getRange("F4:J19").clearContent()

  
  // 月間実績を”ガントチャート”から”月間実績”へ表示する
  const macroSheet = SpreadsheetApp.getActive().getSheetByName("macro");
  const yearAndMonthArray  = macroSheet.getRange(12, 2, 1, 2).getDisplayValues(); // B12, C12セル
  const year = yearAndMonthArray[0][0];
  const month = yearAndMonthArray[0][1];
  displayMonthlyDataFromGantt("20" + year, month); //function.gsに定義されている関数を呼び出しています！
  const montext = month + "月度分"
  monthplan.getRange("D1").setValue(montext)

  const schemeSheet = SpreadsheetApp.getActive().getSheetByName("月間実績")
  const schemeSheetLast = schemeSheet.getRange("A:A").getValues().filter(String).length
  const dayrow = monthlySheet.getRange(1,2,1000,2).getDisplayValues()
  const scheme = schemeSheet.getRange(5,1,schemeSheetLast - 1,8).getDisplayValues()
  const schemevertical = transpose(scheme)
  const scheme1 = transpose([schemevertical[0]])
  const scheme2 = transpose([schemevertical[1]])
  const scheme3 = transpose([schemevertical[2],schemevertical[3],schemevertical[4],schemevertical[5],schemevertical[6]])
  const scheme4 = transpose([schemevertical[7]])
  
  for (var i = 0; i < dayrow.length; i++){
    var yearrow1 = dayrow[i][0]
    var monthrow1 = dayrow[i][1]
    var yearrow2 = dayrow[i + 1][0]
    var monthrow2 = dayrow[i + 1][1]
    if (month != 1) {
      if (yearrow1 == year && monthrow1 == month - 1 && yearrow2 != year && monthrow2 != month - 1) {
       monthlySheet.getRange(i + 4,7,schemeSheetLast - 1, 1).setValues(scheme1)
       monthlySheet.getRange(i + 4,9,schemeSheetLast - 1,1).setValues(scheme2)
       monthlySheet.getRange(i + 4,12,schemeSheetLast - 1,1).setValues(scheme4)
       monthlySheet.getRange(i + 4,2,schemeSheetLast - 1,1).setValue(year)
       monthlySheet.getRange(i + 4,3,schemeSheetLast - 1,1).setValue(month)
       if ( scheme1.length <= 15) {
       monthplan.getRange(4,6,schemeSheetLast - 1,5).setValues(scheme3)
       monthlySheet.getRange(i + 4,1,schemeSheetLast - 1,1).copyTo(monthplan.getRange(4,1,19))
       }  
       break
      }
    }else{
      if (yearrow1 == year - 1 && monthrow1 == 12 && yearrow2 != year && monthrow2 != 12) {
       monthlySheet.getRange(i + 4, 7, schemeSheetLast - 1, 1).setValues(scheme1)
       monthlySheet.getRange(i + 4,9,schemeSheetLast - 1,1).setValues(scheme2)
       monthlySheet.getRamge(i + 4,12,schemeSheetLast - 1,1).setValues(scheme4)
       monthlySheet.getRange(i + 4,2,schemeSheetLast - 1,1).setValue(year)
       monthlySheet.getRange(i + 4,3,schemeSheetLast - 1,1).setValue(month)
       if ( scheme1.length <= 15) {
       monthplan.getRange(4,6,schemeSheetLast - 1,5).setValues(scheme3)
       monthlySheet.getRange(i + 4,1,schemeSheetLast - 1,1).copyTo(monthplan.getRange(4,1,19))
       }  
       break
      }
    }
  }
}


