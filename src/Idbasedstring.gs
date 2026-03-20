function insertIdBasedOnString() {
  // ガントチャートシートの取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ガントチャート');
  var range = sheet.getDataRange();
  var data = range.getValues();
  
  // 特定の文字列とそれに対応するIDをマッピングしたオブジェクト
  var idMap = {
    "英単語・英熟語": "11ET00",
    "英文法": "12EB00",
    "英文解釈": "13EK00",
    "英語長文": "14EC00",
    "リスニング": "15EL00",
    "英作文": "16EW00",
    "現代文": "21JG00",
    "古文漢文": "22JO00",
    "記述": "23JA00",
    "数学": "31MB00",
    "物理基礎": "4PHB00",
    "化学基礎": "5CHB00",
    "生物基礎": "6BIB00",
    "地学基礎": "7ESB00",
    "物理": "45PH00",
    "化学": "46CH00",
    "生物": "47BI00",
    "地学": "48ES00",
    "世界史": "51WH00",
    "日本史": "52JH00",
    "地理": "53GG00",
    "現代社会": "54MS00",
    "倫理": "55MR00",
    "政治経済": "56GE00"
  };
  
  // B列で特定の文字列を検索し、対応するIDをA列に入力
  for (var i = 0; i < data.length; i++) {
    var cellValue = data[i][1];  // B列の値
    if (idMap[cellValue]) {  // 値がマッピングオブジェクトに存在する場合
      data[i][0] = idMap[cellValue];  // 対応するIDをA列に入力
    }
  }
  
  // 更新されたデータをシートに反映
  range.setValues(data);
}