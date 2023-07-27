function chatGpt3() {
  chatGptRequest("gpt-3.5-turbo-0613")
}
function chatGpt4() {
  chatGptRequest("gpt-4-0613")
}

/**
 * ChatGPTに、ヒアリング項目を生成してもらう関数です。
 * 先週のヒアリング項目と、スプレッドシートのB1セルに記入された追加情報をもとに、質問を7つ生成します。
 * 
 */
function chatGptRequest(model) {
  // OpenAIのAPIキーを取得
  const apiKey = ScriptProperties.getProperty('APIKEY');
  // OpenAIのエンドポイント
  const apiUrl = 'https://api.openai.com/v1/chat/completions';
  // System パラメータ
  const systemPrompt = `
  あなたは優秀な大学受験予備校の講師です。コーチング手法に基づく適切な質問文を作ってもらいます。毎回、質問を７つ作成してもらいます。質問の字数は、最大で30文字までです。これから付加される生徒の情報を考慮して、質問を出力してください。質問は、7つで生徒の幅広い情報を得ることができるようにしてください。
  `
  // 入出力シート
  const sheetPrompt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ヒアリング');
  sheetPrompt.getRange("B3:B16").clearContent();
  sheetPrompt.getRange("B3").setValue("generating...");
　
  // 送信メッセージを定義
  let messages = [{'role': 'system', 'content': systemPrompt},];

  // 先週のヒアリング項目を取得
  let previousWeek = sheetPrompt.getRange("A3:A16").getValues();
  console.log(previousWeek)

  var previous_prompt = "ちなみに、先週の質問は以下のようなものでした： \n\n";
  
  for (var i = 0; i < previousWeek.length; i++) {
    if (i % 2 == 0) {
      previous_prompt += (i/2 + 1) + ". " + previousWeek[i][0] + "\n";
    }
  }

  console.log(previous_prompt)

  // 質問内容を追加
  messages.push({'role': 'user', 'content': previous_prompt});

  // 入出力シートからの入力情報を取得
  let promptCell = sheetPrompt.getRange("B1");
  let prompt = promptCell.getValue();

  // 質問内容を追加
  messages.push({'role': 'user', 'content': prompt + "\n\n## 出力"});

  // パラメータ設定
  const requestBody = {
    'model': model,
    'temperature': 0.7,
    'max_tokens': 2048,
    'messages': messages
  }
　
  // 送信内容を設定
  const request = {
    method: "POST",
    muteHttpExceptions : true,
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + apiKey,
    },
    payload: JSON.stringify(requestBody),
  }

  // 回答先のセルを指定
  responseCell = sheetPrompt.getRange("B3:B16");
  try {
    //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
    const response = JSON.parse(UrlFetchApp.fetch(apiUrl, request).getContentText());
    console.log(response.choices[0].message.content)
    sentences = extractSentences(response.choices[0].message.content);
    console.log(sentences);

    // ChatGPTのAPIレスポンスをセルに記載
    if (response.choices) {
      responseCell.setValues(sentences);
    } else {
      // レスポンスが正常に返ってこなかった場合の処理
      console.log(response);
      responseCell.setValue(response);
    }
  } catch(e) {
    // 例外エラー処理
    console.log(e);
    responseCell.setValue(e);
  }
}


function displayKakomon() {
  /**
   * 今週の課題に過去問が含まれている場合、その情報を「ヒアリング」シートのC１セルに表示する関数。
   * 3つのキーワード「大学」「過去問」「模試」のいずれかが含まれていたら、その情報を抽出して表示します。
   */
  const sheetPrompt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ヒアリング');
  const sheetWeekly = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('週間管理');
  let thisWeek = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('今月プラン').getRange("A1").getValue()[0];
  thisWeek = Number(thisWeek)

  let weekToRange = {1: "H4:H19", 2: "P4:P19", 3: "X4:X19", 4: "AF4:AF19", 5: "AN4:AN19"};
  let weeklyLog = sheetWeekly.getRange(weekToRange[thisWeek]).getValues()
  weeklyLog = transpose(weeklyLog)[0]

  console.log(weeklyLog)

  var keywords = ["大学", "過去問", "模試"];
  
  var filteredStrings = weeklyLog.filter(function(str) {
    return keywords.some(function(keyword) {
      return str.indexOf(keyword) >= 0;
    });
  });

  if (filteredStrings.length == 0) {
    res = "なし";
  }else{
    res = filteredStrings.join("\n")
  }

  console.log(res)
  
  return "【今週の過去問・模試】\n\n" + res
}

// UTILITY FUNCTIONS ============================================================================================================================

function extractSentences(text) {
  // Split the text into lines
  var lines = text.split("\n");
  
  // Initialize an empty array to hold the sentences
  var sentences = [];
  
  // Iterate over each line
  for (var i = 0; i < lines.length; i++) {
    // Remove the initial number and period from each line
    var sentence = lines[i].replace(/^\d+\.\s/, "");
    
    // Add the sentence to the array
    sentences.push(sentence);
  }

  // Initialize an empty 2D array
  var values = [];
  
  // Iterate over the sentences
  for (var i = 0; i < sentences.length; i++) {
    // Add the sentence and an empty string to the 2D array
    values.push([sentences[i]]);
    values.push([""]);
  }
  
  // Return the array of sentences
  return values;
}


function test(){

  const sheetPrompt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('週間管理');
  let promptCell = sheetPrompt.getRange("P4:P19");

  console.log(promptCell)
}
