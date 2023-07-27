function getRowsWithEnglishTime() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("ガントチャート");
  var range = sheet.getRange("D:D");
  var values = range.getValues();
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

  Logger.log(rownumber)

  // rowsWithEnglishTime を必要に応じて他の処理に使用することができます

  return rownumber;
}

function togantt() {
  const transpose = a=> a[0].map((_, c) => a.map(r => r[c]))
  var englishrow = getRowsWithEnglishTime()
  SpreadsheetApp.getActiveSpreadsheet().insertSheet().setName("個人マスター")
  var gantt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ガントチャート")
  gantt.deleteColumns(2,2)
  var ganttscheme = gantt.getRange(11, 3, englishrow - 10, 1).getDisplayValues()
  var newscheme = ganttscheme[0]
  for (var i = 1; i < ganttscheme.length; i++) {
    if (ganttscheme[i][0] === null || ganttscheme[i][0] === ""){
      
    }else{
      newscheme.push(ganttscheme[i][0])
    }
  }
  Logger.log(newscheme)
  var newnewscheme = transpose([newscheme])
  Logger.log(newnewscheme)
  var master = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("個人マスター")
  master.getRange(2,4,newnewscheme.length,1).setValues(newnewscheme)
  master.getRange(1,1,1,6).setValues([["ID","カテゴリー","教科","教材・学習内容","月間目標","単位当たり処理量"]])
  gantt.getRange(1,8).clearContent()
  gantt.getRange(1,8).setValue("ID")
  var achievement = SpreadsheetApp.getActive().getSheetByName("月間実績")
  achievement.insertColumnBefore(1)
  achievement.getRange(1,4).setValue("ID")
  var kanri = SpreadsheetApp.getActive().getSheetByName("月間管理")
  for(var i = 0 ; i < 5; i++){
    kanri.insertColumnBefore(15)
  }
  kanri.getRange(1,14).clearContent()
  kanri.getRange(1,14,1,5).setValues([["実績  1週目","実績  2週目","実績  3週目","実績  4週目","実績  5週目"]])
  var thisplan = SpreadsheetApp.getActive().getSheetByName("今月プラン")
  thisplan.hideColumn(thisplan.getRange("A1"))
  thisplan.setRowHeights(4,16,32)
  thisplan.setColumnWidth(2,40)
  thisplan.setColumnWidths(3,2,224)
  thisplan.setColumnWidths(5,7,49)
  thisplan.setColumnWidth(12,230)
  thisplan.setColumnWidth(15,174)
  thisplan.setColumnWidths(18,2,204)
  thisplan.setColumnWidths(21,2,224)
  thisplan.setColumnWidths(24,10,33)
  thisplan.setColumnWidth(34,168)
  thisplan.setColumnWidth(35,380)
  thisplan.setColumnWidth(39,380)
  thisplan.setColumnWidths(36,2,60)
  thisplan.setColumnWidths(40,2,60)
  var monthfirst = SpreadsheetApp.getActive().getSheetByName("月初")
  monthfirst.hideColumn(monthfirst.getRange("A1"))
  monthfirst.setRowHeights(4,16,32)
  monthfirst.setColumnWidth(2,40)
  monthfirst.setColumnWidths(3,2,224)
  monthfirst.setColumnWidths(5,7,49)
  monthfirst.setColumnWidth(12,230)
  monthfirst.setColumnWidth(15,174)
  monthfirst.setColumnWidths(18,2,204)
  monthfirst.setColumnWidth(20,40)
  monthfirst.setColumnWidths(21,2,224)
  monthfirst.setColumnWidths(23,7,49)
  monthfirst.setColumnWidth(30,230)
  monthfirst.setColumnWidths(32,7,140)
  monthfirst.setColumnWidths(41,2,224)
  monthfirst.setColumnWidths(43,10,33)
  monthfirst.setColumnWidth(54,168)
  monthfirst.setColumnWidth(55,380)
  monthfirst.setColumnWidth(59,380)
  monthfirst.setColumnWidths(56,2,60)
  monthfirst.setColumnWidths(60,2,60)
}