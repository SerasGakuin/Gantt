function bokumetsu() {
  /**
   * この関数ではシート「今月プラン」から週間計画表の部分を抜き出し、新たにシート「週間計画表」としてシートの作成をしています。
   */
  var spreadsheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('今月プラン')

  var formulas = [];
  for (var i = 0; i < 16; i++) {
    formulas.push(['=今月プラン!B' + (4 + i)]);
  }

  spreadsheet1.getRange('T4:T19').setFormulas(formulas);  
  var spreadsheet = SpreadsheetApp.getActive()
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('今月プラン'), true);
  spreadsheet.getRange('U4').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(\'今月プラン\'!C4=0,"",\'今月プラン\'!C4)');

  spreadsheet.getRange('U5').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(\'今月プラン\'!C5=0,"",\'今月プラン\'!C5)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('U5:U18'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  spreadsheet.getRange('U19').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(\'今月プラン\'!C19=0,"",\'今月プラン\'!C19)');
  spreadsheet.getRange('V4').activate();
  spreadsheet.getCurrentCell().setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IF(IFS(\'今月プラン\'!$A$1="1週",\'週間管理\'!H4,\'今月プラン\'!$A$1="2週",\'週間管理\'!P4,\'今月プラン\'!$A$1="3週",\'週間管理\'!X4,\'今月プラン\'!$A$1="4週",\'週間管理\'!AF4,\'今月プラン\'!$A$1="5週",\'週間管理\'!AN4)=0,"",IFS(\'今月プラン\'!$A$1="1週",\'週間管理\'!H4,\'今月プラン\'!$A$1="2週",\'週間管理\'!P4,\'今月プラン\'!$A$1="3週",\'週間管理\'!X4,\'今月プラン\'!$A$1="4週",\'週間管理\'!AF4,\'今月プラン\'!$A$1="5週",\'週間管理\'!AN4))), 1, 1)');
  spreadsheet.getRange('V5').activate();
  spreadsheet.getCurrentCell().setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IF(IFS(\'今月プラン\'!$A$1="1週",\'週間管理\'!H5,\'今月プラン\'!$A$1="2週",\'週間管理\'!P5,\'今月プラン\'!$A$1="3週",\'週間管理\'!X5,\'今月プラン\'!$A$1="4週",\'週間管理\'!AF5,\'今月プラン\'!$A$1="5週",\'週間管理\'!AN5)=0,"",IFS(\'今月プラン\'!$A$1="1週",\'週間管理\'!H5,\'今月プラン\'!$A$1="2週",\'週間管理\'!P5,\'今月プラン\'!$A$1="3週",\'週間管理\'!X5,\'今月プラン\'!$A$1="4週",\'週間管理\'!AF5,\'今月プラン\'!$A$1="5週",\'週間管理\'!AN5))), 1, 1)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('V5:V18'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('V19').activate();
  spreadsheet.getCurrentCell().setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IF(IFS(\'今月プラン\'!$A$1="1週",\'週間管理\'!H19,\'今月プラン\'!$A$1="2週",\'週間管理\'!P19,\'今月プラン\'!$A$1="3週",\'週間管理\'!X19,\'今月プラン\'!$A$1="4週",\'週間管理\'!AF19,\'今月プラン\'!$A$1="5週",\'週間管理\'!AN19)=0,"",IFS(\'今月プラン\'!$A$1="1週",\'週間管理\'!H19,\'今月プラン\'!$A$1="2週",\'週間管理\'!P19,\'今月プラン\'!$A$1="3週",\'週間管理\'!X19,\'今月プラン\'!$A$1="4週",\'週間管理\'!AF19,\'今月プラン\'!$A$1="5週",\'週間管理\'!AN19))), 1, 1)');
  spreadsheet.getRange('W4').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,2,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,2,FALSE))');
  spreadsheet.getRange('W5').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,3,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,3,FALSE))');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('W5:W18'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('W19').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,17,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,17,FALSE))');
  spreadsheet.getRange('T1:AO20').activate();
  spreadsheet.insertSheet(3);
  spreadsheet.getActiveSheet().setName('週間計画表');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('今月プラン'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('週間計画表'), true);
  spreadsheet.getRange('T1').activate();
  spreadsheet.getRange('\'今月プラン\'!T1:AO20').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  spreadsheet.getRange('\'今月プラン\'!T1:AO20').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('D4').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('週間計画表'), true);
  spreadsheet.getRange('V8').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('今月プラン'), true);
  spreadsheet.getRange('1:1').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('P1'));
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('週間計画表'), true);
  spreadsheet.getRange('1:1').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('R1'));
  spreadsheet.getActiveSheet().setRowHeight(1, 43);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('今月プラン'), true);
  spreadsheet.getRange('2:2').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('P2'));
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('週間計画表'), true);
  spreadsheet.getRange('2:2').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('R2'));
  spreadsheet.getActiveSheet().setRowHeight(2, 23);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('今月プラン'), true);
  spreadsheet.getRange('4:4').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('P4'));
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('週間計画表'), true);
  spreadsheet.getRange('3:3').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('R3'));
  spreadsheet.getActiveSheet().setRowHeight(3, 23);
  spreadsheet.getRange('4:19').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('R4'));
  spreadsheet.getActiveSheet().setRowHeights(4, 16, 32);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('今月プラン'), true);
  spreadsheet.getRange('20:20').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('Q20'));
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('週間計画表'), true);
  spreadsheet.getRange('20:20').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('R20'));
  spreadsheet.getActiveSheet().setRowHeight(20, 33);
  spreadsheet.getRange('V1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('今月プラン'), true);
  spreadsheet.getRange('T5').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('週間計画表'), true);
  spreadsheet.getCurrentCell().setFormula('=LEFT(\'今月プラン\'!D1,4)&" 第"&\'今月プラン\'!A1');
  spreadsheet.getRange('A:S').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('S1'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('T:T').activate();
};

function syuzen() {
  /**
   * この関数ではbokumetsu()で生じてしまった週間時間の関数のミスをシート「今月プラン」とシート「週間計画表」において修正しています。
   */
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('今月プラン'), true);
  spreadsheet.getRange('W6').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,4,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,4,FALSE))');
  spreadsheet.getRange('W7').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,4,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,4,FALSE))')
  .setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,5,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,5,FALSE))');
  spreadsheet.getRange('V1:W8').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('W8'))
  .setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,6,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,6,FALSE))');
  spreadsheet.getRange('W9').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,7,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,7,FALSE))');
  spreadsheet.getRange('W10').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,8,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,8,FALSE))');
  spreadsheet.getRange('W11').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,9,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,9,FALSE))');
  spreadsheet.getRange('W12').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,10,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,10,FALSE))');
  spreadsheet.getRange('W13').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,11,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,11,FALSE))');
  spreadsheet.getRange('W14').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,12,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,12,FALSE))');
  spreadsheet.getRange('W15').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,13,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,13,FALSE))');
  spreadsheet.getRange('W16').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,14,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,14,FALSE))');
  spreadsheet.getRange('W17').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,15,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,15,FALSE))');
  spreadsheet.getRange('W18').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,16,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,16,FALSE))');
  spreadsheet.getRange('W19').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('週間計画表'), true);
  spreadsheet.getRange('V1:W6').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('W6'))
  .setFormula('=IF(HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,4,FALSE)=0,"",HLOOKUP(\'今月プラン\'!$A$1,\'今月プラン\'!$F$3:$J$19,4,FALSE))');
  spreadsheet.getRange('V1:W6').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('今月プラン'), true);
  spreadsheet.getRange('W4:W19').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('週間計画表'), true);
  spreadsheet.getRange('W4').activate();
  spreadsheet.getRange('\'今月プラン\'!W4:W19').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('Y7').activate();
};
