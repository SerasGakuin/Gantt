/* 01_minor_projects.gs : このファイルの関数は、このプロジェクト内のどのスクリプトからでも、呼び出すことができます。

  - recordFunctionUsage          : 使用しているのかわからない関数の用途の特定に利用する関数。
  - transpose                    : 配列の転置用の関数 @deprecated
  - AdjustMonthView              : ガントチャートの開始日が変更された際に，2行目の月の表示のずれを自動で修正する関数
  - MergeHorizontally            : セル結合を解除し、連続している値が存在するときにそのセルを結合する関数
  - concat                       : 与えられた２つの配列を結合する関数
  - getMonthlyData               : ガントチャートから、year, month月のデータをフィルターして取得する関数
  - displayMonthlyDataFromGantt  : 指定された年月の月間実績をガントチャートから取得して、”月間実績”シートに表示する関数
  - updateMonthlyRecord          : ガントチャートを見にいって、"月間実績"表示を更新する関数
  
*/

/**
 * 使用しているのかわからない関数の用途の特定に利用する関数。
 * 関数の用途をstackで確認し、ログシートに記録。
 *
 * 2026/01/10:
 * レガシーの整理目的で設計。
 * ?_から名前が始まるファイルの関数全部の冒頭でこの関数を呼ぶようにした。
 * 2027/06/01ぐらいまで使用が確認されなければ、
 * あれらはまとめて相談の上、コメントアウトしようと思います。(今後この関数を仕込む場合にはそのたびに、いつ頃まで使用されなければコメントアウトするかここにメモしておいてください。)
 *
 */
function recordFunctionUsage() {
  try {
    const stack = String(new Error().stack);
    const msg =
      "！！！！！！重要！！！！！！\nエラーではなく関数の用途調査記録です。\nもしこのログを見かけましたら、このログをコピーしてどこかに張り付けて記録しておいていただけると幸いです。:\n" +
      stack;
    console.info(msg);
    GASRefferenceSheetLogService.info(msg);
  } catch (e) {
    //何かあっても処理本体を止めない
    console.error(e);
  }
}

/**
 * @deprecated
 * 01_MatrixUtils.gsのMatrixUtils.transposeに移行することをおすすめします。これは後方互換性のために残しています。
 * @type {(a: any) => any}
 */
function transpose(target) {
  //var transpose = a => a[0].map((_, c) => a.map(r => r[c])) // もとの定義
  return MatrixUtils.transpose(target, false);
}

/**
 * @deprecated
 */
function AdjustMonthView() {
  recordFunctionUsage(); //使用を記録 2027/06/01ぐらいまで使用が確認されなければ削除予定
  /**
   * ガントチャートの開始日が変更された際に，2行目の月の表示のずれを自動で修正する関数
   */

  var objCell = SpreadsheetApp.getActiveSheet().getActiveCell();

  // J2セル(ガントチャート開始日)が編集されたときのみ，結合しなおす関数を実行
  if (objCell.getRow() == 2 && objCell.getColumn() == 10) {
    Browser.msgBox(
      "ガントチャート開始日が変更されました．月表示を修正します．",
    );
    MergeHorizontally("O", 3, "BS", 3); //O3:BS3の範囲
  }
}

/**
 * @deprecated
 */
function MergeHorizontally(cell1, n1, cell2, n2) {
  recordFunctionUsage(); //使用を記録 2027/06/01ぐらいまで使用が確認されなければ削除予定
  /**
   * function taken from https://sei-simple.com/it/gas-merge-v/
   * 与えられたセルの範囲の中で、一度セル結合を解除し、連続している値が存在する時に，そのセルを結合する関数
   *
   * @param {str} cell1  - 範囲の始まりの列名。”A”などのアルファベット。
   * @param {num} n1     - 範囲の始まりの行数。
   * @param {str} cell2  - 範囲の終わりの列名。”A”などのアルファベット。
   * @param {num} n２     - 範囲の終わりの行数。
   */
  //

  //一番左上のセル(cell1, n1)にarrayformulaで何らかの関数が入っていることが前提

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getRange(cell1 + n1 + ":" + cell2 + n2);

  range.breakApart(); //まずセル結合解除

  var values = range.getValues();

  for (var row = 0; row < range.getNumRows(); row++) {
    //行数分処理
    var rowA1 = range.getRow() + row; //行Noに1(row)ずつ足していくことで現在の行Noを代入
    var mCount = 1;

    for (var col = 0; col < range.getNumColumns(); col++) {
      //列数分処理
      var colA1 = range.getColumn() + col; //列Noに1(col)ずつ足していくことで現在の列Noを代入
      //次の値のnullチェック && 空白チェック && 結合済みチェック
      if (
        values[row][col + 1] &&
        !sheet.getRange(rowA1, colA1 + 1).isBlank() &&
        !sheet.getRange(rowA1, colA1 + 1).isPartOfMerge()
      ) {
        //次の値がnull/空白/結合済みじゃなければ、一致しているかチェック
        if (values[row][col] == values[row][col + 1]) {
          mCount++;
          // 一致していなくて mCountが1より大きい場合結合してmCountリセット
        } else if (mCount > 1) {
          //2行一致の時にmCount=2、３行一致の時に mCount=3、セルは値が一致している一番下のセルになる
          sheet.getRange(rowA1, colA1 - mCount + 1, 1, mCount).mergeAcross(); //（値が一致している一番上の行No,列No,結合する行数(1),結合する列数）を横方向に結合
          var mCount = 1; //結合セルカウントリセット
        }
        //次の値がnull/空白/結合済みで mCountが1より大きい場合は結合してmCountリセット
      } else if (mCount > 1) {
        //2行一致の時にmCount=2、３行一致の時に mCount=3、セルは値が一致している一番下のセルになる
        sheet.getRange(rowA1, colA1 - mCount + 1, 1, mCount).mergeAcross(); //（値が一致している一番上の行No,列No,結合する行数(1),結合する列数）を横方向に結合
        var mCount = 1; //結合セルカウントリセット
        //次の値がnull/空白/結合済みで mCountが1なら何もしない
      } else {
      }
    }
  }
}

/**
 * ２つの配列を結合する関数
 * @param {array} array1  - １つ目の配列
 * @param {array} array2  - ２つ目の配列
 * @param {num}   axis    - 結合する軸、０なら縦方向、１なら横方向
 *
 * @returns {array} 結合された配列
 */
function concat(array1, array2, axis) {
  recordFunctionUsage(); //使用を記録
  if (axis != 1) axis = 0;
  var array3 = [];
  if (axis == 0) {
    //　縦方向の結合
    array3 = array1.slice();
    for (var i = 0; i < array2.length; i++) {
      array3.push(array2[i]);
    }
  } else {
    //　横方向の結合
    for (var i = 0; i < array1.length; i++) {
      array3[i] = array1[i].concat(array2[i]);
    }
  }
  return array3;
}

/**
 * 指定された年月のデータ（月間実績）を、ガントチャートをフィルターして取得する関数。「ID」「参考書名」「その月の週ごとの計画時間」が入った配列を返す
 *
 * 単純に言えば、「特定年月の学習予定」を二次元配列で取得する関数です。行は教材に、列は「ID」「参考書名」...などの情報の種類に対応します。
 * ガントチャートシートから「勉強予定時間が0でない」「教材行」を抜き出しています。
 * 前者の条件は自明ですが、後者は、教材ではなく勉強時間の集計用の行が存在するので、それを除外するためのものです。
 *
 * 2026_01_31:
 * TODO: この機能はSheetIOに移行したい。そうでないと一貫性がなく、後任が混乱しそう
 * NOTE: この機能だけ、シートとのIOであるのにも関わらずここに定義されていますが、設計上の理由はありません。改修予定のまま、停止しています。
 * 確認している限りでは、geshoとchangeMaterialsで使用されています。
 *
 * @deprecated : SheetIOに移行したい。代わりはまだないので注意
 */
function getMonthlyData(year, month) {
  const yearNum = Number(year); //高速判定のために変換。
  const monthNum = Number(month);

  // 渡された要素が指定年月のものであるかを判定する補助関数
  function isThisYearMonth(element) {
    const date = element instanceof Date ? element : new Date(element);
    return date.getFullYear() === yearNum && date.getMonth() + 1 === monthNum;
  }

  //指定年月に対応する列のindexを取得する補助関数。
  const DATE_START_COL_INDEX = 12; //これで省かないと「ガントチャート開始日」というテキストでエラーになる
  function collectYearMonthCols(dateRow) {
    const cols = [];
    for (let i = DATE_START_COL_INDEX; i < dateRow.length; i++) {
      if (isThisYearMonth(dateRow[i])) {
        //指定年月を発見
        cols.push(i);
      } else if (cols.length !== 0) {
        //連続範囲である前提なので、すでに指定年月の列を見つけた後に違う年月に出会ったら終了
        break;
      }
    }
    return cols;
  }

  // データ読み込み
  const thisBook = SpreadsheetApp.getActive();
  const spIOManager = SheetIO.getSpeedPlannerIOManager(thisBook);
  const ganttMatrix = spIOManager.ganttChartSheet_getAll();

  // 指定年月の列の範囲を取得
  const DATE_ROW_IDX = 1;
  const dateRow = ganttMatrix[DATE_ROW_IDX]; // 日付の入っている行を取得
  const studyTimeIndexArray = collectYearMonthCols(dateRow);
  const yearMonthStudyTimeStartCol = studyTimeIndexArray[0]; //開始行
  const yearMonthStudyTimeEndCol =
    studyTimeIndexArray[studyTimeIndexArray.length - 1]; //終了行

  //列と内容の対応。0始まりの配列idxであって、1始まりの列番号ではない
  const MATERIALS_IDS_COL_IDX = 0;
  const MATERIALS_NAMES_COL_IDX = 2;

  //参考書の行であり、かつ勉強時間が0でない行を絞り込む。
  const ID_MIN_LENGTH = 3; //idとして認められる最短のながさ
  const filteredRows = ganttMatrix.filter((row) => {
    //参考書の行であるか判定
    const hasMaterialId = row[MATERIALS_IDS_COL_IDX].length >= ID_MIN_LENGTH;
    const hasMaterialName =
      row[MATERIALS_NAMES_COL_IDX]?.toString().trim() !== "";
    const isMaterialRow = hasMaterialId && hasMaterialName;
    if (!isMaterialRow) return false; //参考書の行でないなら除外
    //勉強時間が0でないか判定
    const studyTimes = row.slice(
      yearMonthStudyTimeStartCol,
      yearMonthStudyTimeEndCol + 1,
    );
    const totalStudyTime = studyTimes.reduce(
      (sum, value) => sum + (Number(value) || 0),
      0,
    );
    if (totalStudyTime == 0) return false; //勉強時間が0なら除外
    //ここまで除外されなければ含める
    return true;
  });

  //勉強時間が0時間でない参考書の行たちから、目的のデータを抜き出して最終行列に格納
  const yearMonthData = filteredRows.map((row) => {
    const id = row[MATERIALS_IDS_COL_IDX];
    const materialsName = row[MATERIALS_NAMES_COL_IDX];
    const studyTimes = row.slice(
      yearMonthStudyTimeStartCol,
      yearMonthStudyTimeEndCol + 1,
    );
    return [id, materialsName, ...studyTimes]; //行を組み立てる
  });
  //Logger.log(yearMonthData);

  return yearMonthData;
}

/**
 * 指定された年月の月間実績をガントチャートから取得して、”月間実績”シートに表示する関数
 * 「ID」「参考書名」「その月の週ごとの計画時間」が入った二次元配列を返しつつ、月間実績にデータセット
 * @deprecated : 数式版に移行予定
 */
function displayMonthlyDataFromGantt(year, month) {
  const yearNum = Number(year);
  const monthNum = Number(month);
  if (Number.isNaN(yearNum) || Number.isNaN(monthNum)) {
    throw new Error(
      `正常な年月の数字を渡してください。\nyear:${year}, month:${month}`,
    );
  }

  const targetSheet = SpreadsheetApp.getActive().getSheetByName("月間実績");

  const maxrow = targetSheet.getMaxRows();
  targetSheet.getRange(5, 1, Math.max(maxrow - 5, 1), 6).clearContent();

  //指定された年月の月間実績をガントチャートから取得
  const monthData = getMonthlyData(yearNum, monthNum);

  // データが存在しなければ終了
  if (monthData.length == 0) {
    const stackTrace = new Error().stack.split("\n").slice(1).join("\n");
    Logger.log(
      `警告: 指定された年月のデータが存在しませんでした。空データが月間実績に代入されます。\nyear:${year}, month:${month}\n${stackTrace}`,
    );
    return;
  }

  //月間実績のシートに貼り付け。
  targetSheet.getRange(1, 3).setValue(yearNum);
  targetSheet.getRange(2, 3).setValue(monthNum);

  targetRange = targetSheet.getRange(
    5,
    1,
    monthData.length,
    monthData[0].length,
  );
  targetRange.setValues(monthData);

  return monthData;
}

/**
 * @deprecated : 代わりはまだないので注意 数式で同様の機能を実装するように変更議論中
 */
function updateMonthlyRecord() {
  /**
   * ガントチャートを見にいって、"月間実績"表示を更新する関数。
   * year, monthは引数を取ってくるのではなく、「月間実績」のセルに入力されているものを取得する。
   */

  console.log("Hi");

  const targetSheet = SpreadsheetApp.getActive().getSheetByName("月間実績");

  // 引数を読み込む
  const year = targetSheet.getRange(1, 3).getValue();
  const month = targetSheet.getRange(2, 3).getValue();

  const lastRow = targetSheet.getLastRow();
  targetSheet.getRange(5, 1, Math.max(lastRow - 4, 0), 7).clearContent();

  const monthData = getMonthlyData(year, month);

  console.log(monthData);

  // データが存在しなければ終了
  if (monthData.length == 0) return;

  targetRange = targetSheet.getRange(
    5,
    1,
    monthData.length,
    monthData[0].length,
  );
  targetRange.setValues(monthData);
}
