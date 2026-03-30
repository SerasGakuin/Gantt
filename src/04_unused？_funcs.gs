// 使われて無さそうな関数を置く場所。

/**
 * 今週の課題に過去問が含まれている場合、その情報を表示する関数
 * @returns {string} 過去問・模試の情報
 */
function displayKakomon() {
  recordFunctionUsage();// 2026-03-27から設置
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const weeklySheet = sheet.getSheetByName("週間管理");
  const thisWeek = Number(
    sheet.getSheetByName("今月プラン").getRange("A1").getValue()[0],
  );

  const weekToRange = {
    1: "H4:H19",
    2: "P4:P19",
    3: "X4:X19",
    4: "AF4:AF19",
    5: "AN4:AN19",
  };
  const weeklyLog = weeklySheet
    .getRange(weekToRange[thisWeek])
    .getValues()
    .flat();

  const keywords = ["大学", "過去問", "模試"];
  const filteredStrings = weeklyLog.filter((str) =>
    keywords.some((keyword) => str.includes(keyword)),
  );

  const res = filteredStrings.length ? filteredStrings.join("\n") : "なし";
  return "【今週の過去問・模試】\n\n" + res;
}
