function getStartRow(ss, alreadyDeletedColumns) {
  /**
   * 「英単語・英熟語」と書いてある行を検索し、行数を返す関数。
   */

  //ssが指定されていないとき：現在のシート
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  gantt = ss.getSheetByName("ガントチャート");
  // ID列が１列か３列かによって、検索する列を変更
  if (alreadyDeletedColumns) {
    var range = "B:B"
  } else {
    var range = "D:D"
  }

  var values = gantt.getRange(range).getValues();
  var rownumber = null

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == "英単語・英熟語") {
      rownumber = i + 1
      break
    }
  }
  return rownumber;
}

function getRowsWithEnglishTime(ss, alreadyDeletedColumns) {
  /**
   * 英語(時間)と書いてある行を検索し、行数を返す関数。
   */

  //ssが指定されていないとき：現在のシート
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  gantt = ss.getSheetByName("ガントチャート");
  // ID列が１列か３列かによって、検索する列を変更
  if (alreadyDeletedColumns) {
    var range = "B:B"
  } else {
    var range = "D:D"
  }

  var values = gantt.getRange(range).getValues();
  var rownumber = null

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == "英語(時間)") {
      rownumber = i + 1
      break
    }
    if (values[i][0] == "英語(時間）") {
      rownumber = i + 1
      break 
    }
  }
  
  // rowsWithEnglishTime を必要に応じて他の処理に使用することができます
  return rownumber;
}

function vlookupPartial(array, searchKey, searchInd, returnInd) {
  /**
   * VLOOKUPのように、値を検索して別の行の値を返す関数
   */
  var partialMatchCount = 0;
  var partialMatchValue = '';
  
  for (var i = 0; i < array.length; i++) {
    if (array[i][searchInd] === searchKey) { // 完全一致
      return array[i][returnInd];
    } else if (array[i][searchInd].indexOf(searchKey) !== -1 || searchKey.indexOf(array[i][searchInd]) !== -1) { // 双方向部分一致

      if (array[i][searchInd] !== "") {
        partialMatchCount++;
        partialMatchValue = array[i][returnInd];
      }
      
    }
  }
  
  if (partialMatchCount === 1) { // 部分一致が1つだけの場合
    return partialMatchValue;
  } else {
    return "";
  }
}

var cat2sub = {
  '英単語・英熟語': '英語',
  '英文法': '英語',
  '英文解釈': '英語',
  '英語長文': '英語',
  'リスニング': '英語',
  '英作文': '英語',
  '英語': '英語',
  '現代文': '国語',
  '古文漢文': '国語',
  '記述': '国語',
  '国語': '国語',
  '数学': '数学',
  '物理基礎': '理科',
  '化学基礎': '理科',
  '生物基礎': '理科',
  '地学基礎': '理科',
  '物理': '理科',
  '化学': '理科',
  '生物': '理科',
  '地学': '理科',
  '理科': '理科',
  '世界史': '社会',
  '日本史': '社会',
  '地理': '社会',
  '現代社会': '社会',
  '倫理': '社会',
  '政治経済': '社会',
  '社会': '社会',
  '小論文': 'その他',
  '面接': 'その他',
  '推薦書類': 'その他',
  '課外活動': 'その他'
}

var cat2id = {
  '英単語・英熟語': '11ET',
  '英文法': '12EB',
  '英文解釈': '13EK',
  '英語長文': '14EC',
  'リスニング': '15EL',
  '英作文': '16EW',
  '現代文': '21JG',
  '古文漢文': '22JO',
  '記述': '23JW',
  '数学': '31MM',
  '物理基礎': '41PHB',
  '化学基礎': '42BIB',
  '生物基礎': '43CHB',
  '地学基礎': '44ESB',
  '物理': '45PH',
  '化学': '46BI',
  '生物': '47CH',
  '地学': '48ES',
  '世界史': '51WH',
  '日本史': '52JH',
  '地理': '53GG',
  '現代社会': '54MS',
  '倫理': '55MR',
  '政治経済': '56GE'
};

function createIndividualMasterArray(catAndBookArray, masterArray) {

  var pairsArray = [];  // カテゴリー名と参考書名のペアを格納する配列
  var currentCategory = '';  // 現在のカテゴリー名を格納する変数
  var idArray = []; // IDを格納する配列
  var inarow = 0 // 同じカテゴリーがどれだけ連続しているか

  // 2次元配列を行ごとにループ処理する。
  for (var i = 0; i < catAndBookArray.length; i++) {
    var category = catAndBookArray[i][0];  // 1列目のデータ（カテゴリー名）を取得する。
    var referenceBook = catAndBookArray[i][1];  // 2列目のデータ（参考書名）を取得する。

    // カテゴリー名が空でなければ、現在のカテゴリー名を更新する。
    if (category !== '') {
      currentCategory = category;
      inarow = 0
    }

    // 参考書名が空でなければ、ペアの配列に現在のカテゴリー名と参考書名を追加する。IDも追加する。
    if (referenceBook !== '') {
      pairsArray.push([currentCategory, referenceBook]);
      inarow ++;
      idArray.push(cat2id[currentCategory] + Utilities.formatString("%02d", inarow))
    }
  }

  var categoryArray = transpose(pairsArray)[0];
  var subjectArray = categoryArray.map(function(cat) { return cat2sub[cat] });
  var bookArray = transpose(pairsArray)[1];
  var goalArray = bookArray.map(function(book) { return vlookupPartial(masterArray, book, 3, 4) });
  var throughputArray = bookArray.map(function(book) { return vlookupPartial(masterArray, book, 3, 5) });

  var newScheme = transpose([idArray, categoryArray, subjectArray, bookArray, goalArray, throughputArray]);

  return newScheme;
}
