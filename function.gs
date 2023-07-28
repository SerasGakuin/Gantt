/* function.gs
このスクリプトには、５つの関数が定義されています。これらの関数は、このプロジェクト内のどのスクリプトからでも、呼び出すことができます。
  - AdjustMonthView             : ガントチャートの開始日が変更された際に，2行目の月の表示のずれを自動で修正する関数
  - MergeHorizontally           : セル結合を解除し、連続している値が存在するときにそのセルを結合する関数
  - concat                      : 与えられた２つの配列を結合する関数
  - getMonthlyData              : ガントチャートから、year, month月のデータをフィルターして取得する関数
  - displayMonthlyDataFromGantt : 指定された年月の月間実績をガントチャートから取得して、”月間実績”シートに表示する関数
*/

 /**
* @type {(a: any) => any}
*/
var transpose = a=> a[0].map((_, c) => a.map(r => r[c]))


function AdjustMonthView(){
  /**
   * ガントチャートの開始日が変更された際に，2行目の月の表示のずれを自動で修正する関数
   */

  var objCell = SpreadsheetApp.getActiveSheet().getActiveCell();

  // J2セル(ガントチャート開始日)が編集されたときのみ，結合しなおす関数を実行
  if (objCell.getRow() == 2 && objCell.getColumn() == 10) {
    Browser.msgBox("ガントチャート開始日が変更されました．月表示を修正します．");
    MergeHorizontally('O',3,'BS',3);//O3:BS3の範囲
  }
}

function MergeHorizontally(cell1,n1,cell2,n2){
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
  var range = sheet.getRange(cell1+n1+':'+cell2+n2);

  range.breakApart(); //まずセル結合解除

  var values = range.getValues();

  for (var row = 0; row < range.getNumRows(); row++) {//行数分処理
  var rowA1= range.getRow() + row;//行Noに1(row)ずつ足していくことで現在の行Noを代入
  var mCount = 1;

    for (var col = 0; col < range.getNumColumns(); col++) {//列数分処理
    var colA1 = range.getColumn()+col;//列Noに1(col)ずつ足していくことで現在の列Noを代入
      //次の値のnullチェック && 空白チェック && 結合済みチェック 
      if(values[row][col + 1] && !sheet.getRange(rowA1,colA1+1).isBlank() && !sheet.getRange(rowA1,colA1+1).isPartOfMerge() ){     
        //次の値がnull/空白/結合済みじゃなければ、一致しているかチェック
        if (values[row][col]==values[row][col + 1] ) {
        mCount ++;
        // 一致していなくて mCountが1より大きい場合結合してmCountリセット
        }else if(mCount>1){//2行一致の時にmCount=2、３行一致の時に mCount=3、セルは値が一致している一番下のセルになる
          sheet.getRange(rowA1,colA1-mCount+1,1,mCount).mergeAcross();//（値が一致している一番上の行No,列No,結合する行数(1),結合する列数）を横方向に結合
          var mCount = 1;//結合セルカウントリセット
        }
        //次の値がnull/空白/結合済みで mCountが1より大きい場合は結合してmCountリセット
      }else if(mCount>1){//2行一致の時にmCount=2、３行一致の時に mCount=3、セルは値が一致している一番下のセルになる
        sheet.getRange(rowA1,colA1-mCount+1,1,mCount).mergeAcross();//（値が一致している一番上の行No,列No,結合する行数(1),結合する列数）を横方向に結合
        var mCount = 1;//結合セルカウントリセット
      //次の値がnull/空白/結合済みで mCountが1なら何もしない
      }else{}
    }
  }
}

function concat(array1, array2, axis){
  /**
  　　* ２つの配列を結合する関数
　　　　 * @param {array} array1  - １つ目の配列
　　　　　　* @param {array} array2  - ２つ目の配列
   * @param {num}   axis    - 結合する軸、０なら縦方向、１なら横方向
   * 
　　　　 * @return {array} 結合された配列
　　　　 */
  if(axis != 1) axis = 0;
  var array3 = [];
  if(axis == 0){  //　縦方向の結合
    array3 = array1.slice();
    for(var i = 0; i < array2.length; i++){
      array3.push(array2[i]);
    }
  }
  else{  //　横方向の結合
    for(var i = 0; i < array1.length; i++){
      array3[i] = array1[i].concat(array2[i]);
    }
  }
  return array3;
}


function getMonthlyData(year, month){
  /**
  　　* 指定された年月のデータ（月間実績）を、ガントチャートをフィルターして取得する関数
　　　　 * @param {number} year  - 取得する年
　　　　　　* @param {number} month - 取得する月
　　　　　　*
　　　　 * @return {array} 月間実績のデータ
　　　　 */

  // データ読み込み
  const dataSheet = SpreadsheetApp.getActive().getSheetByName("ガントチャート")
  const dataArray = dataSheet.getDataRange().getValues()

  // 参考書の行を絞り込む
  const rowOfBook = 2;

  const bookArray = dataArray.filter(
    record => record[0].length >= 3 & record[rowOfBook] != ""
  )

  //Logger.log(bookArray)

  // 時刻データを取得
  const timeArray = dataArray[1]
  //Logger.log(timeArray)

  function isThisMonth(element, i){
    // timeArrayの要素が，Date型であり，かつyear/monthのものであるかを判定する関数
    if (Object.prototype.toString.call(element) == '[object Date]' & i != 9){ //9はガントチャート開始日のセルの位置なので除外する
      if (element.getFullYear() == year & element.getMonth()+1 == month){
        return i;
      }
    }
    return [];
  }

  // month月の計画時間を絞り込む
  const monthIndex = timeArray.flatMap(isThisMonth) //今月である列のindexを取得
  var monthData = bookArray.map(data => data.slice(monthIndex[0], monthIndex.slice(-1)[0] + 1));
  //Logger.log(monthData)

  // 参考書名をくっつける
  const bookName = bookArray.map(data => [data[rowOfBook]]);
  monthData = concat(bookName, monthData, 1);
  //Logger.log(bookName)

  const idData = bookArray.map(data => [data[0]])
  //Logger.log(idData)
  monthData = concat(idData, monthData, 1)

  
  // 今月の計画時間が0時間である参考書を除外する
  monthData = monthData.filter(
    data => data.slice(2, data.length).reduce((sum, element) => sum + element, 0) > 0
  )
  //Logger.log(monthData)

  return monthData
}

function displayMonthlyDataFromGantt(year, month) {
  /**
  　　* 指定された年月を”macro”シートから読み取り、その年月の月間実績をガントチャートから取得して、”月間実績”シートに表示する関数
　　　　 * @param {number} year  - 取得する年
　　　　　　* @param {number} month - 取得する月
　　　　 */
  const targetSheet = SpreadsheetApp.getActive().getSheetByName("月間実績")

  const maxrow = targetSheet.getMaxRows()
  targetSheet.getRange(5, 1, Math.max(maxrow-5, 1), 6).clearContent()

  // 月間実績を取得
  const monthData = getMonthlyData(year, month)
  //Logger.log(monthData)

  // データが存在しなければ終了
  if (monthData.length == 0){
    return;
  }

  //月間実績のシートに貼り付け。
  targetSheet.getRange(1, 3).setValue(year)
  targetSheet.getRange(2, 3).setValue(month)

  targetRange = targetSheet.getRange(5, 1, monthData.length, monthData[0].length)
  targetRange.setValues(monthData)
}


