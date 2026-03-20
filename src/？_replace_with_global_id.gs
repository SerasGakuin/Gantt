//このファイルの用途調査中の関数は、recordFunctionUsage()から使用報告が2027/06/01ぐらいまでなければ、コメントアウトされ、アーカイブされます。
/**
 * @deprecated 用途調査中の関数。
 */
function replaceWithGlobalIds() {

  recordFunctionUsage();//この関数が使用されたらそれを検知して記録 2027/06/01ぐらいまで使用が確認されなければ削除予定

  // データを取得
  const sheet = SpreadsheetApp.getActive().getSheetByName("ガントチャート");
  var data = sheet.getRange("A1:B1000").getValues();

  // ガントチャート部分をスライス
  var subjectNames = transpose(data)[1]
  const topRow = subjectNames.indexOf('英単語・英熟語');
  const bottomRow = subjectNames.indexOf('英語(時間）');
  data = data.slice(topRow, bottomRow)

  // 英数字のみの正規表現
  const alphanumericRegex = /^[a-zA-Z0-9]+$/;

  // ２列目にグローバル参考書IDがある場合、１列目に移動する
  var transformedData = data.map(function (row) {
    // もしrow[1]が空ではない、かつ英数字のみで構成されているなら、
    if (row[1] != "" & alphanumericRegex.test(row[1])) {
      return [row[1], ""];
    }
    else {
      return [row[0], row[1]];
    }
  })

  //console.log(transformedData)

  // 出力の貼り付け
  sheet.getRange(topRow + 1, 1, bottomRow - topRow, 2).setValues(transformedData);
}
