var ExcelUtils = (function() {
  const RADIX = 26; //アルファベットの文字数
  const A = 'A'.charCodeAt(0);

  return {
    /**
     * アルファベット表記の列番号を、始まりを1とした数字にる列番号に変換する。
     * @param {string} str
     * @return {number} 
     */
    convertAlphabetColumnToNumeric: function(str) {
      var s = str.toUpperCase();
      var n = 0;
      for (var i = 0, len = s.length; i < len; i++) {
        n = (n * RADIX) + (s.charCodeAt(i) - A + 1);
      }
      return n;
    },

    /**
     * 始まりを1とした数字にる列番号を、アルファベット表記の列番号に変換する。
     * @param {number} num
     * @return {string} 
     */
    convertNumericColumnToAlphabet: function(num) {
      var n = num;
      var s = "";
      while (n >= 1) {
        n--;
        s = String.fromCharCode(A + (n % RADIX)) + s;
        n = Math.floor(n / RADIX);
      }
      return s;
    }
  };
})();

// UTILITY FUNCTIONS ===========================================================================================================================================

function isArrayEmpty(array) {
  // 2次元配列の全ての要素をチェックし、空の配列であることを確かめる
  for(var i = 0; i < array.length; i++) {
    for(var j = 0; j < array[i].length; j++) {
      // 要素が空でない場合、falseを返す
      if(array[i][j]) {
        return false;
      }
    }
  }
  // 全ての要素が空の場合、trueを返す
  return true;
}

function detectDoubleTilde(inputString) {
  // 文字列の中に、〜が2つ以上含まれているかをチェック
  var pattern = /~/g;
  var matches = inputString.match(pattern);
  if (matches && matches.length >= 2) {
    return true;
  } else {
    return false;
  }
}

function convertNumbers(inputString) {
  // 全角数字（０-９）を半角数字（0-9）に変換
  var outputString = inputString.replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  
  return outputString;
}

function doesNotContainNumbers(inputString) {
  // テキスト中に数字が含まれていないことを検出する
  var regex = /\d/;
  
  return !regex.test(inputString);
}

function convertTildas(inputString) {
  // 全角チルダ（～）、全角波ダッシュ（〜）、半角波ダッシュ（∼）を半角チルダ（~）に変換
  var outputString = inputString.replace(/[～〜∼]/g, '~');
  
  return outputString;
}

function extractLettersAndNumbers(input) {
  // 正規表現を使用して英文字から始まる部分を検索し、その後に続く数字も含める
  var pattern = /([A-Za-z]+\d*)/;
  var match = input.match(pattern);
  
  // マッチした場合はその結果を返し、しない場合は空文字列を返す
  return match ? match[1] : "";
}

// 対応する範囲の目次上の名前を返す関数
function findSectionName(number, data) {
  for (let row of data) {
    let start = parseInt(row[7]);
    let end = parseInt(row[8]);
    if (number >= start && number <= end) {
      return row[6];  // 7行目（インデックス6）の文字列を返す
    }
  }
  return "";  // 該当する範囲が見つからない場合
}