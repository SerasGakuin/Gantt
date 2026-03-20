

function concat(array1, array2, axis){
  // 2つの配列を結合する関数
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
    if (Object.prototype.toString.call(element) == '[object Date]' & i != 7){ //7はガントチャート開始日のセルの位置なので除外する
      if (element.getFullYear() == year & element.getMonth()+1 == month){
        return i;
      }
    }
    return [];
  }

  // month月の計画時間を絞り込む
  const monthIndex = timeArray.flatMap(isThisMonth) //今月である列のindexを取得
  var monthData = bookArray.map(data => data.slice(monthIndex[0], monthIndex.slice(-1)[0] + 1));

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

function main() {

  console.log("Hi")

  const targetSheet = SpreadsheetApp.getActive().getSheetByName("月間実績")
  
  // 引数を読み込む
  const year = targetSheet.getRange(1, 3).getValue();
  const month = targetSheet.getRange(2, 3).getValue();

  const maxrow = targetSheet.getMaxRows()
  targetSheet.getRange(5, 2, Math.max(maxrow-5, 1), 6).clearContent()

  const monthData = getMonthlyData(year, month)

  console.log(monthData)

  // データが存在しなければ終了
  if (monthData.length == 0){
    return;
  }

  targetRange = targetSheet.getRange(5, 1, monthData.length, monthData[0].length)
  targetRange.setValues(monthData)
}