var ss = SpreadsheetApp.getActiveSpreadsheet();

class GanttChart {

  constructor(ss) {
    this.ss = ss;
    this.gantt = this.ss.getSheetByName("ガントチャート")

    // 「ガントチャート」が存在するか
    this.existGantt = !!this.gantt

    // ガントチャートのID列は3列から1列に変更済みか
    this.alreadyDeletedColumns = this.existGantt ? this.gantt.getRange("C3").getValue() === "■参考書ガントチャート" : false ;
    
    // 「ガントチャート」の「英単語・英熟語」は何行目か
    this.startRow = this.existGantt ? getStartRow(this.ss, this.alreadyDeletedColumns) : null ;

    // 「ガントチャート」の「英語（時間）」は何行目か
    this.endRow = existGantt ? getRowsWithEnglishTime(this.ss, alreadyDeletedColumns) : null ;

  }

}

function toGantt(ss) {
  /**
   * ssで指定されたシートに、一連のガントチャート移管作業を適用する関数
   */
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  //ssを受け取り、その個体差をチェックする
  var gantt = ss.getSheetByName("ガントチャート")

  // 「ガントチャート」が存在するか
  var existGantt = !!gantt

  // ガントチャートのID列は3列から1列に変更済みか
  var alreadyDeletedColumns = existGantt ? gantt.getRange("C3").getValue() === "■参考書ガントチャート" : false ;
  
  // 「ガントチャート」の「英単語・英熟語」は何行目か
  var startRow = existGantt ? getStartRow(ss, alreadyDeletedColumns) : null ;

  // 「ガントチャート」の「英語（時間）」は何行目か
  var endRow = existGantt ? getRowsWithEnglishTime(ss, alreadyDeletedColumns) : null ;

  insertIndividualMasterSheet(ss, existGantt, alreadyDeletedColumns, startRow, endRow);
  updatePercentageFormulas(ss, startRow, endRow);
  updateGanttChartIdView(ss, alreadyDeletedColumns,startRow);
  updateMonthlyAchievement();
  updateMonthlyFirst();
  updateMonthlyManagement();
  updateMonthlyPlan();
}

function test() {

  var gantt = new GanttChart(SpreadsheetApp.getActiveSpreadsheet())
  console.log(gantt)
}





// TODO: 右から2列目に移動
function insertIndividualMasterSheet(ss, alreadyDeletedColumns, startRow, endRow) {
  /**
   * ①新しいシート「個人マスター」を作成する
   * ②A1に「ID」、B1に「カテゴリー」、C1に「教科」、D1に「教材・学習内容」、E1に「月間目標」、F1に「単位当たり処理量」と入力する。(デザインは未定。G1の標準時間は廃止)
   * ③D列に「ガントチャート」シートを参照しながら、該当生徒の使用する参考書を全て記入する。
   * ④それぞれの参考書に対応する教科名をC列、カテゴリーをD列に入力(教科名、カテゴリー名はシート「マスター」に準拠)
   * ⑤シート「マスター」のそれぞれの参考書に対応する「CAT」をID欄に入力。またシート「マスター」にあるそれぞれの参考書に対応する「月間目標」、「単位当たり処理量」をそれぞれシート「個人マスター」のE列、F列、G列に入力する。
   */

  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!alreadyDeletedColumns) {
    alreadyDeletedColumns = true;
  }
  if (!startRow) {
    startRow = getStartRow(ss, alreadyDeletedColumns);
  }
  if (!endRow) {
    endRow = getRowsWithEnglishTime(ss, alreadyDeletedColumns);
  }

  var individualMasterSheet = ss.getSheetByName("個人マスター")
  if (!individualMasterSheet) {
    individualMasterSheet = ss.insertSheet().setName("個人マスター")
  }

  // 「個人マスター」を右から２番目に移動 
  ss.setActiveSheet(individualMasterSheet);
  ss.moveActiveSheet(ss.getNumSheets() - 1);

  // 「個人マスター」のID列に何も値がなかったら、シート全体を初期化して空白にする 
  var data = individualMasterSheet.getRange('A2:A' + individualMasterSheet.getLastRow()).getValues();
  // A列の二行目以降が全て空かどうかを判定
  var isEmpty = data.every(function(row) { return row[0] === ""; });

  // すでに個人マスターに情報がある場合、処理を終了
  if (!isEmpty) { return }

  individualMasterSheet.clear();

  // 列名を追加
  individualMasterSheet.getRange(1,1,1,6).setValues([["ID","カテゴリー","教科","教材・学習内容","月間目標","単位当たり処理量"]])

  // ガントチャートから参考書名を取得
  var [from, to] = alreadyDeletedColumns ? ["B", "C"] : ["D", "E"]
  var catAndBookArray = ss.getSheetByName("ガントチャート").getRange(from+startRow.toString()+":"+to+(endRow-1).toString()).getDisplayValues();

  // 「マスター」から参考書の情報を取得
  var masterSheet = ss.getSheetByName("マスター");
  var masterArray = masterSheet.getRange(3, 4, masterSheet.getLastRow(), 7).getValues();

  // 「個人マスター」に貼り付ける配列を作成 
  var newSheme = createIndividualMasterArray(catAndBookArray, masterArray);
  individualMasterSheet.getRange(2,4,newSheme.length,1).setValues(newSheme)
}


function updateGanttChartIdView(ss, alreadyDeletedColumns, startRow) {
  /**
   * ⑥続いてシート「ガントチャート」に関する変更を行う。「列B-Cを削除」をクリックして、B、C列を削除する。
   * ⑦セルA9の「Code」を「ID」に変える。
   * ⑧参考書のある行のA列に、それぞれの参考書に対応するIDをシート「個人マスター」からコピペする。
   */

  // ID列が3行になっている時だけ、BC列を削除
  var gantt = ss.getSheetByName("ガントチャート")
  if (!alreadyDeletedColumns) {
    gantt.deleteColumns(2,2);
  }
  gantt.getRange(1,8).clearContent()
  gantt.getRange(1,8).setValue("ID")

  // A列にIDを表示する関数を挿入
  var formula = '=ARRAYFORMULA(IFNA(XLOOKUP(C11:C, \'個人マスター\'!D2:D, \'個人マスター\'!A2:A, ""), ""))';
  gantt.getRange('A11').setFormula(formula);

}


function updatePercentageFormulas(ss, endRow) {
  /**
   * ⑨スプレッドシート「上堅さん検証用_0331」のシート「ガントチャート」のセルF114からF150を各生徒の英語の必要時間セルから社会の必要時間比のセルまでコピーする。
   * その後、コピペした後のセルを同じ行のG列、MからBQまでオートフィル。参考書の時間が書いてある最終列にコピー。
   */

  //コピー元とコピー先のシートを指定
  var gantt = ss.getSheetByName("ガントチャート");
  var ganttUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("ガントチャート");
  var toRange = function (from, to) { return  from + endRow.toString() + ":" + to + (endRow + 36).toString() }

  //F列に関数をコピー
  var formulas = ganttUekata.getRange("F114:F150").getFormulas();
  var sourceRange = gantt.getRange(toRange("F", "F"))
  sourceRange.setFormulas(formulas)

  // FからBQ列まで関数をオートフィル、その後HからL列の関数を消去
  sourceRange.autoFill(gantt.getRange(toRange("F", "BQ")), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  gantt.getRange(toRange("H", "L")).clearContent();
}


function updateMonthlyAchievement() {
  /**
   * ⑩次に、シート「月間実績」に関する変更を行う。A列を選択し、右クリックした後、「左に一列挿入」をクリックし、新たなA列を作成する。
   * ⑪A4セルにIDと入力。
   */

  var monthlyAchievementSheet = ss.getSheetByName("月間実績")
  
  monthlyAchievementSheet.insertColumnBefore(1)
  monthlyAchievementSheet.setColumnWidth(1,120)
  monthlyAchievementSheet.getRange("A4").setValue("ID")
}


function updateMonthlyManagement() {
  /**
   * ⑫次に、シート「月間管理」に関する変更を行う。その後、スプレッドシート「上堅さん検証用_0331」のシート「月間管理」のF2からK2までを該当生徒のシート「月間管理」のF(lastrow)からK(lastrow)までコピーする。ただし、vlookupの参照元の範囲がズレるので、Lastrowを取得してそれに合わせる。
   * ⑬⑩と同様の操作を行いO列の左に4列追加する。そして、N列、新しくできたO、P、Q、R列にそれぞれ「実績　1週目」、「実績　2週目」、「実績　3週目」、「実績　4週目」、「実績　5週目」と入力する。
   */

  var manageSheet = ss.getSheetByName("月間管理")
  var manageUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("月間管理");

  var formulas = manageUekata.getRange("F2:K2").getFormulas();

  var rowToPaste = manageSheet.getRange(sheet.getMaxRows(), 9).getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 2;

  //貼り付ける関数内で使用しているG列の範囲を、貼り付ける行に合わせる
  for (var i = 0; i < formulas[0].length; i++) {
    if (formulas[0][i] !== '') {
      formulas[0][i] = formulas[0][i].replace('G2', 'G' + rowToPaste);
    }
  }

  manageSheet.getRange("F"+rowToPaste+":K"+rowToPaste).setFormulas(formulas);
  
  for(var i = 0 ; i < 5; i++){
    manageSheet.insertColumnBefore(15)
  }
  manageSheet.getRange(1,14).clearContent()
  manageSheet.getRange(1,14,1,5).setValues([["実績  1週目","実績  2週目","実績  3週目","実績  4週目","実績  5週目"]])
  

}


function updateMonthlyPlan() {
  /**
   * ⑭次に、シート「今月プラン」に関する変更を行う。まずA列を選択し、右クリックしたのち、「列を非表示」を選択して、A列を非表示にする。
   * ⑮次にスプレッドシート「上堅さん検証用_0331」のシート「今月プラン」のB4からD19までを該当生徒のシート「今月プラン」の対応箇所にコピー。
   * ⑯以下行や列のサイズ変更を行う。その手順を4〜19行を例にとって説明する。まず、4〜19行を選択し、右クリックした後、「行4 - 19のサイズを変更」を選択する。その後、「行の高さを指定する」に「32」ピクセルと入力する。
   * ⑰以下、列とそれに対応する列の幅を書いている。
   * ⑱その他移行時にテーマが崩れた部分を直す: 日付のフォーマットをDDDにする
   */

  var monthlyPlanSheet = ss.getSheetByName("今月プラン")

  // 関数をコピペ
  var monthlyPlanUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("今月プラン");
  monthlyPlanSheet.getRange("B4:D19").setFormulas(monthlyPlanUekata.getRange("B4:D19").getFormulas());

  // 行や列のサイズを変更
  monthlyPlanSheet.hideColumn(monthlyPlanSheet.getRange("A1"))
  monthlyPlanSheet.setRowHeights(4,16,32)
  monthlyPlanSheet.setColumnWidth(2,40)
  monthlyPlanSheet.setColumnWidths(3,2,224)
  monthlyPlanSheet.setColumnWidths(5,7,49)
  monthlyPlanSheet.setColumnWidth(12,230)
  monthlyPlanSheet.setColumnWidth(15,174)
  monthlyPlanSheet.setColumnWidths(18,2,204)
  monthlyPlanSheet.setColumnWidths(21,2,224)
  monthlyPlanSheet.setColumnWidths(24,10,33)
  monthlyPlanSheet.setColumnWidth(34,168)
  monthlyPlanSheet.setColumnWidth(35,380)
  monthlyPlanSheet.setColumnWidth(39,380)
  monthlyPlanSheet.setColumnWidths(36,2,60)
  monthlyPlanSheet.setColumnWidths(40,2,60)

  // 曜日の表示フォーマットを"ddd"に統一
  monthlyPlanSheet.getRange("X3:AG3").setFormulas(monthlyPlanUekata.getRange("X3:AG3").getFormulas());

}


function updateMonthlyFirst() {
  /**
   * ⑲次にシート「月初」に関する変更を行う。⑭、⑮と同じ操作を行う。
   * ⑳⑯、⑰同様、行と列それぞれの大きさを変える。
   */

  var monthlyFirstSheet = ss.getSheetByName("月初")

  // 関数をコピペ
  var monthlyFirstUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("今月プラン");
  monthlyFirstSheet.getRange("B4:D19").setFormulas(monthlyFirstUekata.getRange("B4:D19").getFormulas());
  
  // 行や列のサイズを変更
  monthlyFirstSheet.hideColumn(monthlyFirstSheet.getRange("A1"))
  monthlyFirstSheet.setRowHeights(4,16,32)
  monthlyFirstSheet.setColumnWidth(2,40)
  monthlyFirstSheet.setColumnWidths(3,2,224)
  monthlyFirstSheet.setColumnWidths(5,7,49)
  monthlyFirstSheet.setColumnWidth(12,230)
  monthlyFirstSheet.setColumnWidth(15,174)
  monthlyFirstSheet.setColumnWidths(18,2,204)
  monthlyFirstSheet.setColumnWidth(20,40)
  monthlyFirstSheet.setColumnWidths(21,2,224)
  monthlyFirstSheet.setColumnWidths(23,7,49)
  monthlyFirstSheet.setColumnWidth(30,230)
  monthlyFirstSheet.setColumnWidths(32,7,140)
  monthlyFirstSheet.setColumnWidths(41,2,224)
  monthlyFirstSheet.setColumnWidths(43,10,33)
  monthlyFirstSheet.setColumnWidth(54,168)
  monthlyFirstSheet.setColumnWidth(55,380)
  monthlyFirstSheet.setColumnWidth(59,380)
  monthlyFirstSheet.setColumnWidths(56,2,60)
  monthlyFirstSheet.setColumnWidths(60,2,60)
}


