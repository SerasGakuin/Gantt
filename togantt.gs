class GanttChart {

  constructor(ss) {

    this.ss = ss;
    this.ganttSheet              = this.ss.getSheetByName("ガントチャート")
    this.individualMasterSheet   = this.ss.getSheetByName("個人マスター")
    this.masterSheet             = this.ss.getSheetByName("マスター");
    this.monthlyAchievementSheet = this.ss.getSheetByName("月間実績");
    this.manageSheet             = this.ss.getSheetByName("月間管理");
    this.monthlyPlanSheet        = this.ss.getSheetByName("今月プラン");
    this.monthlyFirstSheet       = this.ss.getSheetByName("月初")

    // 「ガントチャート」が存在するか
    this.existGantt = !!this.ganttSheet

    // ガントチャートのID列は3列から1列に変更済みか
    this.alreadyDeletedColumns = this.existGantt ? this.ganttSheet.getRange("C3").getValue() === "■参考書ガントチャート" : false ;
    
    // 「ガントチャート」の「英単語・英熟語」は何行目か
    this.startRow = this.existGantt ? getStartRow(this.ss, this.alreadyDeletedColumns) : null ;

    // 「ガントチャート」の「英語（時間）」は何行目か
    this.endRow = this.existGantt ? getRowsWithEnglishTime(this.ss, this.alreadyDeletedColumns) : null ;

  }

  toGantt() {
    /**
     * ssで指定されたシートに、一連のガントチャート移管作業を適用する関数
     */
    insertIndividualMasterSheet();
    if (this.existGantt) {
      updateGanttChartIdView();
      updatePercentageFormulas();
    }
    updateMonthlyAchievement();
    updateMonthlyFirst();
    updateMonthlyManagement();
    updateMonthlyPlan();

  }
  

  insertIndividualMasterSheet() {
    /**
     * ①新しいシート「個人マスター」を作成する
     * ②A1に「ID」、B1に「カテゴリー」、C1に「教科」、D1に「教材・学習内容」、E1に「月間目標」、F1に「単位当たり処理量」と入力する。(デザインは未定。G1の標準時間は廃止)
     * ③D列に「ガントチャート」シートを参照しながら、該当生徒の使用する参考書を全て記入する。
     * ④それぞれの参考書に対応する教科名をC列、カテゴリーをD列に入力(教科名、カテゴリー名はシート「マスター」に準拠)
     * ⑤シート「マスター」のそれぞれの参考書に対応する「CAT」をID欄に入力。またシート「マスター」にあるそれぞれの参考書に対応する「月間目標」、「単位当たり処理量」をそれぞれシート「個人マスター」のE列、F列、G列に入力する。
     */

    if (!this.individualMasterSheet) {
      this.individualMasterSheet = this.ss.insertSheet().setName("個人マスター")
    }

    // 「個人マスター」を右から２番目に移動 
    this.ss.setActiveSheet(this.individualMasterSheet);
    this.ss.moveActiveSheet(this.ss.getNumSheets() - 1);

    // 「個人マスター」のID列に何も値がなかったら、シート全体を初期化して空白にする 
    var data = this.individualMasterSheet.getRange('A2:A' + this.individualMasterSheet.getLastRow()).getValues();
    // A列の二行目以降が全て空かどうかを判定
    var isEmpty = data.every(function(row) { return row[0] === ""; });

    // すでに個人マスターに情報がある場合、処理を終了
    if (!isEmpty) { return }

    this.individualMasterSheet.clear();

    // 列名を追加
    this.individualMasterSheet.getRange(1,1,1,6).setValues([["ID","カテゴリー","教科","教材・学習内容","月間目標","単位当たり処理量"]])

    if (!this.existGantt) { return }

    // ガントチャートから参考書名を取得
    var [from, to] = this.alreadyDeletedColumns ? ["B", "C"] : ["D", "E"]
    var catAndBookArray = this.ganttSheet.getRange(from+this.startRow.toString()+":"+to+(this.endRow-1).toString()).getDisplayValues();

    // 「マスター」から参考書の情報を取得
    var masterArray = this.masterSheet.getRange(3, 4, this.masterSheet.getLastRow(), 7).getValues();

    // 「個人マスター」に貼り付ける配列を作成 
    var newSheme = createIndividualMasterArray(catAndBookArray, masterArray);
    this.individualMasterSheet.getRange(2,4,newSheme.length,1).setValues(newSheme)
  }


  updateGanttChartIdView() {
    /**
     * ⑥続いてシート「ガントチャート」に関する変更を行う。「列B-Cを削除」をクリックして、B、C列を削除する。
     * ⑦セルA9の「Code」を「ID」に変える。
     * ⑧参考書のある行のA列に、それぞれの参考書に対応するIDをシート「個人マスター」からコピペする。
     */

    // ID列が3行になっている時だけ、BC列を削除
    if (!this.alreadyDeletedColumns) {
      this.ganttSheet.deleteColumns(2,2);
      this.alreadyDeletedColumns = true;
    }
    this.ganttSheet.getRange(1,8).clearContent()
    this.ganttSheet.getRange(1,8).setValue("ID")

    // A列にIDを表示する関数を挿入
    var formula = '=ARRAYFORMULA(IFNA(XLOOKUP(C'+ (this.startRow).toString() +':C, \'個人マスター\'!D2:D, \'個人マスター\'!A2:A, ""), ""))';
    this.ganttSheet.getRange('A'+(this.startRow-1).toString()).setFormula(formula);

  }


  updatePercentageFormulas() {
    /**
     * ⑨スプレッドシート「上堅さん検証用_0331」のシート「ガントチャート」のセルF114からF150を各生徒の英語の必要時間セルから社会の必要時間比のセルまでコピーする。
     * その後、コピペした後のセルを同じ行のG列、MからBQまでオートフィル。参考書の時間が書いてある最終列にコピー。
     */

    //コピー元とコピー先のシートを指定
    var ganttUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("ガントチャート");
    var toRange = function (from, to) { return  from + this.endRow.toString() + ":" + to + (this.endRow + 36).toString() }

    //F列に関数をコピー
    var formulas = ganttUekata.getRange("F114:F150").getFormulas();
    var sourceRange = this.ganttSheet.getRange(toRange("F", "F"));
    sourceRange.setFormulas(formulas);

    // FからBQ列まで関数をオートフィル、その後HからL列の関数を消去
    sourceRange.autoFill(this.ganttSheet.getRange(toRange("F", "BQ")), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    this.ganttSheet.getRange(toRange("H", "L")).clearContent();
  }


  updateMonthlyAchievement() {
    /**
     * ⑩次に、シート「月間実績」に関する変更を行う。A列を選択し、右クリックした後、「左に一列挿入」をクリックし、新たなA列を作成する。
     * ⑪A4セルにIDと入力。
     */

    if (this.monthlyAchievementSheet.getRange("B4").getValue() === "参考書") {
      this.monthlyAchievementSheet.insertColumnBefore(1)
    }
    this.monthlyAchievementSheet.setColumnWidth(1,120)
    this.monthlyAchievementSheet.getRange("A4").setValue("ID")
  }


  updateMonthlyManagement() {
    /**
     * ⑫次に、シート「月間管理」に関する変更を行う。その後、スプレッドシート「上堅さん検証用_0331」のシート「月間管理」のF2からK2までを該当生徒のシート「月間管理」のF(lastrow)からK(lastrow)までコピーする。ただし、vlookupの参照元の範囲がズレるので、Lastrowを取得してそれに合わせる。
     * ⑬⑩と同様の操作を行いO列の左に4列追加する。そして、N列、新しくできたO、P、Q、R列にそれぞれ「実績　1週目」、「実績　2週目」、「実績　3週目」、「実績　4週目」、「実績　5週目」と入力する。
     */

    var manageUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("月間管理");
    var formulas = manageUekata.getRange("F2:K2").getFormulas();

    // 参考書名が入力されている最後の行の2行下に、新たな関数をペーストする
    var rowToPaste = this.manageSheet.getRange(sheet.getMaxRows(), 9).getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 2;

    //貼り付ける関数内で使用しているG列の範囲を、貼り付ける行に合わせる
    for (var i = 0; i < formulas[0].length; i++) {
      if (formulas[0][i] !== '') {
        formulas[0][i] = formulas[0][i].replace('G2', 'G' + rowToPaste);
      }
    }

    this.manageSheet.getRange("F"+rowToPaste+":K"+rowToPaste).setFormulas(formulas);
    
    for(var i = 0 ; i < 5; i++){
      this.manageSheet.insertColumnBefore(15)
    }
    this.manageSheet.getRange(1,14).clearContent()
    this.manageSheet.getRange(1,14,1,5).setValues([["実績  1週目","実績  2週目","実績  3週目","実績  4週目","実績  5週目"]])

  }


  updateMonthlyPlan() {
    /**
     * ⑭次に、シート「今月プラン」に関する変更を行う。まずA列を選択し、右クリックしたのち、「列を非表示」を選択して、A列を非表示にする。
     * ⑮次にスプレッドシート「上堅さん検証用_0331」のシート「今月プラン」のB4からD19までを該当生徒のシート「今月プラン」の対応箇所にコピー。
     * ⑯以下行や列のサイズ変更を行う。その手順を4〜19行を例にとって説明する。まず、4〜19行を選択し、右クリックした後、「行4 - 19のサイズを変更」を選択する。その後、「行の高さを指定する」に「32」ピクセルと入力する。
     * ⑰以下、列とそれに対応する列の幅を書いている。
     * ⑱その他移行時にテーマが崩れた部分を直す: 日付のフォーマットをDDDにする
     */
    

    // 関数をコピペ
    var monthlyPlanUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("今月プラン");
    this.monthlyPlanSheet.getRange("B4:D19").setFormulas(monthlyPlanUekata.getRange("B4:D19").getFormulas());

    // 行や列のサイズを変更
    this.monthlyPlanSheet.hideColumn(this.monthlyPlanSheet.getRange("A1"))
    this.monthlyPlanSheet.setRowHeights(4,16,32)
    this.monthlyPlanSheet.setColumnWidth(2,40)
    this.monthlyPlanSheet.setColumnWidths(3,2,224)
    this.monthlyPlanSheet.setColumnWidths(5,7,49)
    this.monthlyPlanSheet.setColumnWidth(12,230)
    this.monthlyPlanSheet.setColumnWidth(15,174)
    this.monthlyPlanSheet.setColumnWidths(18,2,204)
    this.monthlyPlanSheet.setColumnWidths(21,2,224)
    this.monthlyPlanSheet.setColumnWidths(24,10,33)
    this.monthlyPlanSheet.setColumnWidth(34,168)
    this.monthlyPlanSheet.setColumnWidth(35,380)
    this.monthlyPlanSheet.setColumnWidth(39,380)
    this.monthlyPlanSheet.setColumnWidths(36,2,60)
    this.monthlyPlanSheet.setColumnWidths(40,2,60)

    // 曜日の表示フォーマットを"ddd"に統一
    this.monthlyPlanSheet.getRange("X3:AG3").setFormulas(monthlyPlanUekata.getRange("X3:AG3").getFormulas());

  }


  updateMonthlyFirst() {
    /**
     * ⑲次にシート「月初」に関する変更を行う。⑭、⑮と同じ操作を行う。
     * ⑳⑯、⑰同様、行と列それぞれの大きさを変える。
     */


    // 関数をコピペ
    var monthlyFirstUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("今月プラン");
    this.monthlyFirstSheet.getRange("B4:D19").setFormulas(monthlyFirstUekata.getRange("B4:D19").getFormulas());
    
    // 行や列のサイズを変更
    this.monthlyFirstSheet.hideColumn(this.monthlyFirstSheet.getRange("A1"))
    this.monthlyFirstSheet.setRowHeights(4,16,32)
    this.monthlyFirstSheet.setColumnWidth(2,40)
    this.monthlyFirstSheet.setColumnWidths(3,2,224)
    this.monthlyFirstSheet.setColumnWidths(5,7,49)
    this.monthlyFirstSheet.setColumnWidth(12,230)
    this.monthlyFirstSheet.setColumnWidth(15,174)
    this.monthlyFirstSheet.setColumnWidths(18,2,204)
    this.monthlyFirstSheet.setColumnWidth(20,40)
    this.monthlyFirstSheet.setColumnWidths(21,2,224)
    this.monthlyFirstSheet.setColumnWidths(23,7,49)
    this.monthlyFirstSheet.setColumnWidth(30,230)
    this.monthlyFirstSheet.setColumnWidths(32,7,140)
    this.monthlyFirstSheet.setColumnWidths(41,2,224)
    this.monthlyFirstSheet.setColumnWidths(43,10,33)
    this.monthlyFirstSheet.setColumnWidth(54,168)
    this.monthlyFirstSheet.setColumnWidth(55,380)
    this.monthlyFirstSheet.setColumnWidth(59,380)
    this.monthlyFirstSheet.setColumnWidths(56,2,60)
    this.monthlyFirstSheet.setColumnWidths(60,2,60)
  }

}




