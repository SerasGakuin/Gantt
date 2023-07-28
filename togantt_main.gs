function test() {


  var ss = SpreadsheetApp.openById("1yJYH_0SG_DmEbRAbZrqjwnDP6yUQxp5LCoWfEQ5gLtQ");
  var ganttChart = new GanttChart(ss)
  ganttChart.toGantt();
}
