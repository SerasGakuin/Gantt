function test() {


  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ganttSheet = new GanttChart(ss)
  console.log(ganttSheet.endRow)
}
