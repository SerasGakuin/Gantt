function sendweeklyplan(ssurl){
  var studentdata = new StudentMasterLib.studentMaster().data
  for(i=0;i<studentdata.length;i++){
    var seitourl = studentdata[i].スプレッドシート
    if(ssurl.includes(seitourl) || seitourl.includes(ssurl)){
      var lineid = studentdata[i].生徒LINEID
      break;
    }
  }
  Logger.log(lineid)
  var studentsheet = SpreadsheetApp.openByUrl(ssurl)
  var weeklyplan = studentsheet.getSheetByName("週間管理").getRange(4,1,16,40).getDisplayValues()
  var weeknumber = extractNumbersFromString(studentsheet.getSheetByName("今月プラン").getRange(1,1).getDisplayValue())[0]
  Logger.log(weeknumber)
  var thismonthplan = studentsheet.getSheetByName("今月プラン").getRange(4,6,16,5).getValues();
  var weektable = {"week1":7,"week2":15,"week3":23,"week4":31,"week5":39}
  var week = "week" + weeknumber
  for(i=0;i<weeklyplan.length;i++){
    Logger.log(weeklyplan[i][weektable[week]])
    if(weeklyplan[i][weektable[week]] == ""){
      thismonthplan[i][Number(weeknumber)-1] = ""

    }
  }
  studentsheet.getSheetByName("今月プラン").getRange(4,6,16,5).setValues(thismonthplan)
  linesend.sendWeeklyPlan(lineid)
}