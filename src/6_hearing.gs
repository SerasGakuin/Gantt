// Constants
const API_URL = 'https://api.openai.com/v1/chat/completions';
const SYSTEM_PROMPT = `
あなたは優秀な大学受験予備校の講師です。コーチング手法に基づく適切な質問文を作ってもらいます。毎回、オープンクエスチョン形式の質問を７つ作成してもらいます。質問の字数は、最小で20文字、最大で40文字までです。これから付加される生徒の情報を考慮して、質問を出力してください。質問は、7つで生徒の幅広い情報を得ることができるようにしてください。

注意：字数の制限は必ず守ってください。あまり具体的すぎる勉強内容の質問は控えて、それをより一般化したものにしてください。

以下に、テンプレートの質問をいくつか列挙します。これらのうちどれかを使うことが適切だと思われた場合は、それを使ってくれても構いません。その場合、これらのテンプレートのものを一字一句間違えずに使ってください。⚫️になっている部分は、適切に埋めて使ってください。

## テンプレート
志望校・進路
志望校について変更はありませんか？

志望校について、別添のアンケートに記入ください。

進路について、最近考えたり、行動したことは？

進路について、学校で考える機会は最近ありますか？

進路について、今後調べたいor調べて欲しいことは何ですか？

併願校の受験科目・配点・出題傾向で気になることはありますか？

年間計画/参考書の使い方・勉強法
模試の成績で弱点だと思う科目・分野は何ですか？

現在使用している参考書で難易度が合わないものは？

学校では●(科目)●は今、何を習っていますか？

英単語帳は1日に何回触れるようにしていますか？

数学の問題演習は、解きなおし（2周目）をしていますか？

英語長文では、SVOCの構造分析の解説をチェックしていますか？

●月の目標は？

定期テスト後（テスト返却後）
勉強のモチベーションは何%ですか？
（　　　　％）⇒理由：
【定期考査】英語はどうでしたか？反省点は？

【定期考査】数学はどうでしたか？反省点は？

【定期考査】国語はどうでしたか？反省点は？

【定期考査】理科はどうでしたか？反省点は？

【定期考査】社会はどうでしたか？反省点は？

今週の良かったこと、新たに気づいたことは何ですか？

Template5
定期テストの目標点/順位はどのくらいですか？

模試の目標はどのくらいですか？（偏差値、得点、判定）

受験に向けて不安なことはありますか？

ライバルや目標にしている同級生・先輩は？その人の良いところは？

出来るようになりたいことは？（勉強、部活、趣味など）

今週出来るようになったことは？

最近自分が頑張ったこと、自分を褒めたいことはありますか？

Template6
前回の個別特訓では、どんなことを教わりましたか？

最近おすすめのテレビ番組かYouTube動画を教えてください。

〇〇の参考書の中で、重要だと思ったのは何ですか？

最近勉強や進路について同級生や先輩と話しましたか？

来週の予定で何か憂鬱なものはありますか？

参考書の難易度・理解度
●●の難易度はどうでしたか？
（難しい・少し難しい・ちょうど・少し易しい・易しい）
全レベル問題集古文〇で目標点は取れましたか？

●●の解説は理解できそうですか？

●●の問題の初見での正答率はどれくらい？
（　　　　％）減点される理由：

学校生活/学習量の調整/モチベーションコントロール
勉強のモチベーションは何%ですか？
（　　　　％）⇒理由：
今週をもう一度やり直せるなら、どうしたいですか？

今週の良かったこと、新たに気づいたことは何ですか？

部活の状況は？

文化祭の準備は何％くらい完了した？

体育祭の準備は何％くらい完了した？

通塾する予定の曜日に〇をしましょう！
（ 火・水・木・金・土・日 ）

過去問演習・弱点分析
●(過去問)●は何点とれましたか？

●(過去問)●をやってみて、弱点は何だと思いましたか？

数学で復習が必要な単元はどこですか？
単元名：
国語で復習が必要なのは？
（現代文単語・古文単語・古文文法・古文読解・漢文句法）

個別特訓・質問対応の希望・来月の予定
次の個別特訓の希望は？（講師・曜日・内容）

●●の勉強で分からないことはありませんか？

●●の勉強で講師に質問したいことはありませんか？

定期テスト対策で講師の質問対応の希望は？

△次の定期考査のスケジュール（部活オフ、開始日～終了日）

△次の模試のスケジュール（模試名と日程）

来月の部活の予定は？

## それでは、質問の生成をお願いします：
`;

/**
 * ChatGPT 3.5を使用してリクエストを送信する関数
 */
function chatGpt3() {
  chatGptRequest("gpt-3.5-turbo-0613");
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
  const apiKey = ScriptProperties.getProperty('APIKEY');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ヒアリング');
  
  clearPreviousResponses(sheet);
  const messages = prepareMessages(sheet);
  
  const requestBody = {
    'model': model,
    'temperature': 0.7,
    'max_tokens': 2048,
    'messages': messages
  };

  const options = {
    method: "POST",
    muteHttpExceptions: true,
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + apiKey,
    },
    payload: JSON.stringify(requestBody),
  };

  try {
    const response = JSON.parse(UrlFetchApp.fetch(API_URL, options).getContentText());
    processResponse(response, sheet);
  } catch(e) {
    console.error('Error in chatGptRequest:', e);
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
 * @return {Array} 準備されたメッセージの配列
 */
function prepareMessages(sheet) {
  let messages = [{'role': 'system', 'content': SYSTEM_PROMPT}];
  
  const previousWeek = sheet.getRange("A3:A16").getValues();
  const previousPrompt = preparePreviousWeekPrompt(previousWeek);
  messages.push({'role': 'user', 'content': previousPrompt});

  const currentPrompt = sheet.getRange("B1").getValue();
  messages.push({'role': 'user', 'content': currentPrompt + "\n\n## 出力"});

  return messages;
}

/**
 * 前週のプロンプトを準備する
 * @param {Array} previousWeek - 前週の質問データ
 * @return {string} 整形された前週のプロンプト
 */
function preparePreviousWeekPrompt(previousWeek) {
  let prompt = "ちなみに、先週の質問は以下のようなものでした： \n\n";
  for (let i = 0; i < previousWeek.length; i++) {
    if (i % 2 === 0) {
      prompt += `${i/2 + 1}. ${previousWeek[i][0]}\n`;
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
    console.error('Unexpected API response:', response);
    responseCell.setValue(JSON.stringify(response));
  }
}

/**
 * 今週の課題に過去問が含まれている場合、その情報を表示する関数
 * @return {string} 過去問・模試の情報
 */
function displayKakomon() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const weeklySheet = sheet.getSheetByName('週間管理');
  const thisWeek = Number(sheet.getSheetByName('今月プラン').getRange("A1").getValue()[0]);

  const weekToRange = {1: "H4:H19", 2: "P4:P19", 3: "X4:X19", 4: "AF4:AF19", 5: "AN4:AN19"};
  const weeklyLog = weeklySheet.getRange(weekToRange[thisWeek]).getValues().flat();

  const keywords = ["大学", "過去問", "模試"];
  const filteredStrings = weeklyLog.filter(str => 
    keywords.some(keyword => str.includes(keyword))
  );

  const res = filteredStrings.length ? filteredStrings.join("\n") : "なし";
  return "【今週の過去問・模試】\n\n" + res;
}

/**
 * テキストから文を抽出し、番号とピリオドを削除する関数
 * @param {string} text - 入力テキスト
 * @return {Array} 抽出された文の2次元配列
 */
function extractSentences(text) {
  const lines = text.split("\n");
  const sentences = lines.map(line => line.replace(/^\d+\.\s/, ""));
  return sentences.flatMap(sentence => [[sentence], [""]]);
}

/**
 * テスト用関数
 */
function test() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('週間管理');
  const range = sheet.getRange("P4:P19");
  console.log(range.getValues());
}