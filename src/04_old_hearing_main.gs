//SYSTEM_PROMPTとAPI_URLは hearingConstants.gsに

//呼び出し用ラッパー
function generateHearingEntries_old() {
  chatGpt4();
}

/**
 * ChatGPT 4を使用してリクエストを送信する関数
 */
function chatGpt4() {
  chatGptRequest("gpt-4o");
}

/**
 * ChatGPTにヒアリング項目を生成してもらう関数
 * @param {string} model - 使用するGPTモデル
 */
function chatGptRequest(model) {
  const apiKey = ScriptProperties.getProperty("APIKEY");
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ヒアリング");

  clearPreviousResponses(sheet);
  const messages = prepareMessages(sheet);

  const requestBody = {
    model: model,
    temperature: 1,
    max_completion_tokens: 2048,
    messages: messages,
  };

  const options = {
    method: "POST",
    muteHttpExceptions: true,
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + apiKey,
    },
    payload: JSON.stringify(requestBody),
  };

  try {
    const response = JSON.parse(
      UrlFetchApp.fetch(API_URL, options).getContentText(),
    );
    processResponse(response, sheet);
  } catch (e) {
    console.error("Error in chatGptRequest:", e);
    sheet.getRange("B3").setValue("Error: " + e.message);
  }
}

/**
 * 前回の回答をクリアする
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - ヒアリングシート
 */
function clearPreviousResponses(sheet) {
  sheet.getRange("B3:B16").clearContent();
  sheet.getRange("B3").setValue("generating...");
}

/**
 * メッセージを準備する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - ヒアリングシート
 * @returns {Array} 準備されたメッセージの配列
 */
function prepareMessages(sheet) {
  let messages = [{ role: "system", content: SYSTEM_PROMPT }];

  const previousWeek = sheet.getRange("A3:A16").getValues();
  const previousPrompt = preparePreviousWeekPrompt(previousWeek);
  messages.push({ role: "user", content: previousPrompt });

  const currentPrompt = sheet.getRange("B1").getValue();
  messages.push({ role: "user", content: currentPrompt + "\n\n## 出力" });

  return messages;
}

/**
 * 前週のプロンプトを準備する
 * @param {Array} previousWeek - 前週の質問データ
 * @returns {string} 整形された前週のプロンプト
 */
function preparePreviousWeekPrompt(previousWeek) {
  let prompt = "ちなみに、先週の質問は以下のようなものでした： \n\n";
  for (let i = 0; i < previousWeek.length; i++) {
    if (i % 2 === 0) {
      prompt += `${i / 2 + 1}. ${previousWeek[i][0]}\n`;
    }
  }
  return prompt;
}

/**
 * APIレスポンスを処理する
 * @param {Object} response - APIからのレスポンス
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - ヒアリングシート
 */
function processResponse(response, sheet) {
  const responseCell = sheet.getRange("B3:B16");
  if (response.choices && response.choices[0].message) {
    const sentences = extractSentences(response.choices[0].message.content);
    responseCell.setValues(sentences);
  } else {
    console.error("Unexpected API response:", response);
    responseCell.setValue(JSON.stringify(response));
  }
}

/**
 * テキストから文を抽出し、番号とピリオドを削除する関数
 * @param {string} text - 入力テキスト
 * @returns {Array} 抽出された文の2次元配列
 */
function extractSentences(text) {
  const lines = text.split("\n");
  const sentences = lines.map((line) => line.replace(/^\d+\.\s/, ""));
  return sentences.flatMap((sentence) => [[sentence], [""]]);
}

/**
 * テスト用関数
 */
function test() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("週間管理");
  const range = sheet.getRange("P4:P19");
  console.log(range.getValues());
}
