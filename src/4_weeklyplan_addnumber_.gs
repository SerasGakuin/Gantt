function addNumbers(inputString, numToAdd, referenceBookID, referenceBookData) {
  /**
   * 週間計画の範囲を更新する関数
   * 
   * inputString: 入力の文字列
   * numToAdd: 今週追加する問題数
   * referenceBookID: 参考書ID
   * referenceBookData: 参考書マスターの情報が入った配列
   */

  // 今週追加する問題数の丸め
  numToAdd = Math.round(Number(numToAdd));

  // ◎参考書IDが空欄、または今週やらない場合は空欄を返す
  if (referenceBookID == "" | numToAdd == 0) {
    return "";
  }
  // ◎先週やっていなかったら★を返す
  if (!inputString) {
    return "★";
  }

  // 以下、参考書マスターが存在しない場合の従来処理
  // 全角の数字を半角の数字に、全角の〜を半角の~に変換する
  inputString = convertNumbers(inputString);
  inputString = convertTildas(inputString);
  // tildeの前後の空白を消す
  inputString = inputString.replace(/\s*~\s*/g, "~");

  // ◎数字が一切含まれていなかったら文字列をそのまま返す
  if (doesNotContainNumbers(inputString)) {
    return inputString;
  }

  // ◎複数の~が含まれる場合は★を返す
  if (detectDoubleTilde(inputString)) {
    return "★";
  }

  // =============================範囲を更新=============================================================
  // 先頭の評価記号（◯×△）を取得
  var lastEvaluation = getFirstCharacter(inputString);

  // チルダで囲まれた数字を検出する正規表現
  var pattern = /(\d+)~(\d+)/;
  var result = getStartAndEndNumbers(inputString, numToAdd, pattern);

  if (lastEvaluation) {
    // 先頭の評価記号を削除
    inputString = inputString.slice(1);

    // ◯の場合はそのまま
    // △の場合はnewEndNumberを空にし、★を付与
    if (lastEvaluation == "△") {
      result.newStartNumber = result.startNumber;
      result.newEndNumber = "";
      inputString = "★" + inputString;
    }
    // ✕の場合はnewStartNumberとnewEndNumberを元の値に戻す
    if (lastEvaluation == "✕") {
      result.newStartNumber = result.startNumber;
      result.newEndNumber = result.endNumber;
    }
  }

  // ============================表示の更新==============================================================
  // ★★【新仕様】参考書マスターが存在する場合は、inputStringからの数字抽出などの処理をOpenAI APIに委譲する
  if (referenceBookData != "参考書IDが存在しません") {
    // 引数は inputString, numToAdd, referenceBookID, referenceBookData として渡す
    return updateWithRefernceBookMaster(inputString, numToAdd, result.startNumber, result.endNumber, referenceBookID, referenceBookData);
  }

  return updateWithoutRefernceBookMaster(inputString, numToAdd, result.newStartNumber, result.newEndNumber, pattern, result.match, result.singleMatch);
}


function testchar() {
  var s1 = "○Section 44~45"
  var s2 = "△数3: 例題52~61"
  var s3 = "✕模試の過去問"
  var s4 = "1900:No.1~1000の暗記\n熟語1000：No.1~400　始点固定"

  console.log(getFirstCharacter(s1))
  console.log(getFirstCharacter(s2))
  console.log(getFirstCharacter(s3))
  console.log(getFirstCharacter(s4))
}

function getFirstCharacter(inputString) {
  /**
   * 入力の文字列に対して、先頭に○△✕があればそれを返す関数
   */
  var firstChar = inputString[0]

  if ( ["○", "△", "✕"].includes(firstChar) ) {
    return firstChar
  }

  return 
}

function getStartAndEndNumbers(inputString, numToAdd, pattern) {
  /**
   * 文字列を解析して、開始や終了の数字、始点固定の文字を抽出する
   */

  var match = inputString.match(pattern);

  var startNumber;
  var endNumber;

  if (match) {
    // ~で囲まれた数字が検出されたらそれらを取得
    startNumber = parseInt(match[1]);
    endNumber = parseInt(match[2]);

  } else {
    // ~で囲まれた数字が検出されなかった場合、数字を探して1つ目を取得
    var singleMatch = inputString.match(/\d+/);

    if (singleMatch) {
      startNumber = parseInt(singleMatch[0]);
      endNumber = startNumber;
    } 
  }

  var newEndNumber = endNumber + numToAdd;

  // 始点固定の場合_はstartNumberを変更しない
  var fixtedStart = inputString.includes('始点固定');
  if (fixtedStart) {
    var newStartNumber = startNumber;
  } else {
    var newStartNumber = endNumber + 1;
  }

  return {
    match: match,
    singleMatch: singleMatch,
    newStartNumber: newStartNumber,
    newEndNumber: newEndNumber, 
    fixtedStart: fixtedStart, 
    startNumber: startNumber, 
    endNumber: endNumber
  };
}

function updateWithRefernceBookMaster(inputString, numToAdd, startNumber, endNumber, referenceBookID, referenceBookData) {
  // APIキーの取得
  const apiKey = ScriptProperties.getProperty('APIKEY');

  // function calling 用の関数定義（JSON スキーマ）
  const functionDefinition = {
    name: "generateOutputString",
    description: "inputString、numToAdd、参考書マスターデータに基づき、学習計画の出力文字列 outputString を生成する。返り値は outputString のみで、余計な説明やコメントは含めない。",
    parameters: {
      type: "object",
      properties: {
        outputString: {
          type: "string",
          description: "更新された学習計画の文字列。余計な説明やコメントは含まない。",
        }
      },
      required: ["outputString"]
    }
  };

  // 詳細なルールを含むプロンプト
  const prompt = `あなたは、与えられた inputString を元に、学習計画の出力文字列 outputString を生成する補助役です。
  
入力と出力は基本的に同じ形式を保つ必要がありますが、数値部分および単元名（章名）の部分は以下のルールに従って更新してください。

【ルール】

1. startNumber, endNumberがある、範囲指定の場合（連続番号の場合）
   - 例：「No.612~651」、numToAdd = 40 → 出力：「No.652~692」
   - 数値範囲の終了値 endNumber に対して  
     ・新しい開始値 = endNumber + 1  
     ・新しい終了値 = endNumber + numToAdd  
   ※ 接頭辞、区切り文字、接尾辞はそのまま保持する

2. startNumber, endNumberがなく、単一の数値の場合
   - 例：「第2回」、numToAdd = 3 → 出力：「第3~5回」
   - 新しい開始値 = 入力の数値 + 1  
   - 新しい終了値 = 入力の数値 + numToAdd  
   ※ 接頭辞・接尾辞はそのまま保持する

3. 複数の数値または数値範囲が存在する場合  
   - 各部分に上記ルールを個別に適用する

4. 参考書マスターデータに基づく単元名（章名）の更新  
   - 入力文字列内の単元名（章名）として推測される部分は、参考書マスターデータの「章立て」「章の名前」「章のはじめ」「章の終わり」「番号の数え方」を利用して更新する  
   - 入力文字列に単元名がなくても、参考書マスターデータに単元名があるなら、それを必ず追記してください。
   - 例：入力「5章　酸と塩基　ABC問題、すべて」, numToAdd = 2 と参考書マスターデータが与えられた場合、出力は「6章　酸化還元反応 ~ 7章　電池と電気分解　ABC問題、すべて」となる

5. 参考書ごとの番号の連続性の違い  
   - 通し番号の場合はそのまま更新（例: gET002）  
   - 章ごとに番号がリセットされる場合は、各章内での番号として更新する  
     ・例：入力「Chapter 3 単元3 ~ Chapter 4 単元2」、numToAdd = 4、参考書マスターデータの「番号の数え方」が「単元」の場合  
       → 現在の終了値は Chapter 4 単元2  
       → 新しい開始値 = Chapter 4 単元2 + 1 = Chapter 4 単元3  
       → 新しい終了値 = Chapter 4 単元2 + 4 = Chapter 4 単元6  
       → 出力：「Chapter 4 単元3 ~ Chapter 4 単元6」

7. 終わりの範囲の再チェックを必ずしてください。
  - 参考書マスターデータを見て、更新後の範囲が参考書の最後の範囲を超えないようにしてください。
  - 参考書の範囲の最後まで到達したら、参考書の最後の問題を範囲の最後尾にしてください。そして、「★完了！」という文字列を最後尾に加えてください。
  - 入力の方の範囲ですでに参考書の最後に到達済みなら、全ての文字列を「★完了！2周目か次の参考書に進みましょう」に置き換えてください。

6. その他のテキストについては、評価記号、プレフィックス、コロン、空白、その他の記号は可能な限りそのまま保持する。ただし、単元名に該当する部分は更新する。

【指示】
- 上記ルールに従い、与えられた inputString、numToAdd、参考書マスターデータ（referenceBookData）を元に、更新された学習計画の出力文字列 outputString を生成してください。
- 出力は必ず JSON 形式で、余計な説明やコメントは一切含めず、以下の形式に従ってください:

{
  "outputString": "<更新された学習計画の文字列>"
}

入力:
inputString: ${inputString}
startNumber: ${startNumber}
endNumber: ${endNumber}
numToAdd: ${numToAdd}
参考書ID: ${referenceBookID}
参考書マスターデータ: ${JSON.stringify(referenceBookData)}
`;

  // API 呼び出し用ペイロードの作成
  const payload = {
    model: "gpt-4.1", // または適宜選択
    messages: [
      { role: "system", content: "You are a helpful assistant." },
      { role: "user", content: prompt }
    ],
    functions: [functionDefinition],
    function_call: { name: "generateOutputString" },
    max_tokens: 300,
    temperature: 0.0
  };

  

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + apiKey
    },
    payload: JSON.stringify(payload)
  };

  // APIを呼び出して結果を取得
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);

  console.log(response)
  const responseData = JSON.parse(response.getContentText());

  // function_call の結果から arguments を抽出し、JSONをパースする
  const args = JSON.parse(responseData.choices[0].message.function_call.arguments);
  return args.outputString;
}



function updateWithRefernceBookMasterOld(referenceBookData, newStartNumber, newEndNumber, fixtedStart){
    /**
     * 参考書マスターデータが存在する場合の処理
     * referenceBookData: 参考書マスターから取得した参考書の情報
     * newStartNumber
     */
    // 出力の最後に追加する追加の文字列
    var stringToAddFirst = "";
    var stringToAddLast = "";
    var prefix = "No.";

    // 章が参考書の終わりを超えた場合、範囲の最後を、参考書の終わりの問題にする
    var endChapters = referenceBookData.map(row => parseInt(row[8]));
    var maxEndChapter = Math.max(...endChapters);
    
    if (newEndNumber > maxEndChapter) {
      newEndNumber = maxEndChapter;
      stringToAddLast = stringToAddLast + " 完了!"
    }

    // 範囲の最後の問題が含まれる参考書の単元名を取得
    var sectionName = findSectionName(newEndNumber, referenceBookData);
    if (sectionName != ""){
      stringToAddLast = " " + sectionName + " " + stringToAddLast
    }

    // 問題の数え方(「No.」や「問」)を取得
    var tmp = referenceBookData[0][9]
    if (tmp != ""){
      // 後ろにつけるか前につけるか判断
      if (tmp == "章" | tmp == "回"){
        stringToAddLast = tmp + "" + stringToAddLast
      } else {
        prefix = tmp
      }
    }

    // ◎範囲が1問だけの場合は、範囲の初めの部分を返す
    if (newStartNumber >= newEndNumber) {
      return stringToAddFirst + prefix + newStartNumber.toString() + stringToAddLast;
    }

    // 「始点固定」の文字列が含まれる場合、範囲の初めを前回の範囲の始めと同じにする
    if (fixtedStart){
      stringToAddFirst = stringToAddFirst + "始点固定\n"
    }

    // ◎新しい範囲を返す
    //console.log(`新しい範囲: No.${newStartNumber}~${newEndNumber}`);
    return stringToAddFirst + prefix + newStartNumber + "~" + newEndNumber + stringToAddLast;
  }


function updateWithoutRefernceBookMaster(inputString, numToAdd, newStartNumber, newEndNumber, pattern, match, singleMatch) {

  // 前回の範囲が取得できていない場合
  if (!match & !singleMatch) {
    return "★";
  }

  // 前回の範囲が~囲みなのか単一の数字なのかでパターンを分岐
  if (match) {
    pattern = pattern;
  } else if (singleMatch) {
    pattern = /\d+/;
  }

  // 
  if (numToAdd > 1) {
    var outputString = inputString.replace(pattern, newStartNumber + "~" + newEndNumber);
  } 
  else {
    var outputString = inputString.replace(pattern, newStartNumber)
  }
  return outputString;
}


// テスト関数
function testAddNumber(){
  console.log(addNumbers("△問題21∼24", 3, null, "参考書IDが存在しません"));
  console.log(addNumbers("No.612~651", 1, null, "参考書IDが存在しません"));
  console.log(addNumbers("第2回", 3, null, "参考書IDが存在しません"));
  console.log(addNumbers("5章　酸と塩基　ABC問題、すべて", 5, null, "参考書IDが存在しません"));
  let inputString = "１．速度と加速度～５．エネルギーの問題1~65のうち、*1個の問題*";
  console.log(addNumbers(inputString, 18, null, "参考書IDが存在しません"));

  console.log("No. 81 ~ 110".replace(/\s*~\s*/g, "~"))
}

function testAddNumberWithReferenceBookData() {

  // 「参考書マスター」シートから、該当の参考書の情報を取得する
  var referenceBookAllData = getReferenceBookAllData();

  var referenceBookIds = [ 
    'gET002',
    'gEK001',
    'gCHB003',
    'gEK001', 
    'gET003',
    'gET002',
    'gEB004',
    'gEK005',
    'gEC010',
    'kakomon',
    'gEK006',
    'gMB003',
    'gMB023',  
    'gEC005',
    'gJG003',
    'gJO008',
    'gJO004',
    'gJH001' 
  ];

  // サンプルのテストデータ（inputStringと今週追加する問題数）
  var testInputString = "△問題21∼24";
  var testNumToAdd = 3;

  for (var i = 0; i < referenceBookIds.length; i++) {
    var refBookData = getReferenceBookData(referenceBookAllData, referenceBookIds[i]);
    console.log("Reference Book Data for " + referenceBookIds[i] + ": " + JSON.stringify(refBookData));
    
    // addNumbers関数を呼び出して結果を取得
    var result = addNumbers(testInputString, testNumToAdd, referenceBookIds[i], refBookData);
    console.log("Result for " + referenceBookIds[i] + ": " + result);
  }
}
