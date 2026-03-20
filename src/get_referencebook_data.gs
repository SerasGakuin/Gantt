function getReferenceBookData(referenceBookAllData, referenceBookID) {
  let dataArray = referenceBookAllData;
  let result = {
    "参考書ID": "",
    "参考書名": "",
    "教科": "",
    "月間目標": "",
    "単位当たり処理量": "",
    "章立て": [],
    "章の名前": [],
    "章のはじめ": [],
    "章の終わり": [],
    "番号の数え方": []
  };
  
  // Find the row with the matching referenceBookID
  let mainRowIndex = dataArray.findIndex(row => row[0] === referenceBookID);
  
  if (mainRowIndex === -1) {
    return null; // Return null if the ID is not found
  }
  
  let mainRow = dataArray[mainRowIndex];
  
  // Set main info
  result["参考書ID"] = mainRow[0];
  result["参考書名"] = mainRow[1];
  result["教科"] = mainRow[2];
  result["月間目標"] = mainRow[3];
  result["単位当たり処理量"] = mainRow[4];
  
  function addChapterData(row) {
    if (row[5] !== "") {
      result["章立て"].push(row[5]);
      result["章の名前"].push(row[6]);
      result["章のはじめ"].push(row[7]);
      result["章の終わり"].push(row[8]);
      result["番号の数え方"].push(row[9]);
    }
  }
  
  // Add chapter data from the main row
  addChapterData(mainRow);
  
  // Add chapter data from subsequent rows
  for (let i = mainRowIndex + 1; i < dataArray.length; i++) {
    let row = dataArray[i];
    if (row[0] !== "") break; // Stop if we reach a new main entry
    
    addChapterData(row);
  }
  
  return result;
}


function getReferenceBookAllData() {
  // 「参考書マスター」スプレッドシート
  var referenceSheet = SpreadsheetApp.openById("1Z0mMUUchd9BT6r5dB6krHjPWETxOJo7pJuf2VrQ_Pvs").getSheetByName("参考書マスター")
  // 「参考書マスター」シートから、該当の参考書の情報を取得する
  var referenceBookAllData = referenceSheet.getRange("A2:J" + referenceSheet.getLastRow()).getValues();
  return referenceBookAllData
}

function testGetReferenceBookData() {
  var allData = getReferenceBookAllData();
  var bookData = getReferenceBookData(allData, "gEB004"); // Replace with the desired ID
  Logger.log(JSON.stringify(bookData, null, 2));
}