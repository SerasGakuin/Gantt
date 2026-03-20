//03_handlers_withAI.gs ハンドラー置き場
/**
 * 計画生成ハンドラーのうち、特にAIを利用するものを格納するオブジェクト。
 * ファイル外から中身にアクセスする場合は、テスト関数での一時的な目的でない場合は03_handlers.gsのgetterを経由してください。
 * 
 * ハンドラーごとの独立性を最優先してoptionsなどは共通化していないが、一部の設定はdependenciesに格納するようにしてもいいかもしれない。検討中。
 */
//TODO: AIと教材マスターを使用するものについては、AIを最新の学習実績から課題範囲をパースしてもらって、結果の計算は厳密なコードで行うのがよいかもしれない。AIに文字列を解釈させ、数字を抜粋させ、加算させてテキスト化させる、のは責務が重すぎるように感じる。
/** @type {GenWeeklyPlanHandlersContainer} */
const GenWeeklyPlan_handlers_withAI = Object.freeze({
  /**
   * 参考書マスターの情報をもとにAIに頼んで処理してもらうパターン。
   * AI生成はカバーできるパターンがハンドラーの中で最多だが、api料金がかかり、生成時間も長めなので最終手段。
   * @type {GenWeeklyPlanGenerationHandler}
   */
  askingAIWithBookInfoHandler: {
    name: 'askingAIWithBookInfoHandler',
    run(context) {
      if (GenWeeklyPlan_handlers_flags.isAIBanned) return '★AI生成キャンセル';

      if (!context.hasValidMetadata) return null;

      try {
        const result = this._generate(context);
        return result;
      } catch (e) {
        return context.lastAchievement + GenWeeklyPlan_handlers_defaultTexts.onAIError;
      }
    },

    _generate(context) {

      // 詳細なルールを含むプロンプト
      const prompt = this._getPrompt(context.lastAchievement, context.valueToAdd, context.materialInfo);

      const options = this._getOptionsObj(prompt, context);

      // APIを呼び出して結果を取得
      const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);

      //console.log(response)
      const responseData = JSON.parse(response.getContentText());

      // function_call の結果から arguments を抽出し、JSONをパースする
      const rawArgs = responseData.choices[0].message.function_call.arguments;
      const outputStr = this._sanitizeAIOutput(rawArgs);

      return outputStr;
    },

    //AIに渡すoptionオブジェクトを取得する関数。
    _getOptionsObj(prompt, context) {
      // APIキーを取得（セキュリティのために、ベタ打ちを回避してスクリプトプロパティに格納）
      const apiKey = context.dependencies.getChatGPTAPIKEY();
      const modelName = context.dependencies.getChatGPTModelName();

      // function calling 用の関数定義（JSON スキーマ）
      const functionDefinition = {
        name: "generatePlan",
        description: "生徒の最新の学習実績と参考書マスターデータに基づき、学習計画の出力文字列 plan を生成する。返り値は plan のみで、余計な説明やコメントは含めない。",
        parameters: {
          type: "object",
          properties: {
            plan: {
              type: "string",
              description: "更新された学習計画の文字列。余計な説明やコメントは含まない。",
            }
          },
          required: ["plan"]
        }
      };

      // API 呼び出し用ペイロードの作成
      const payload = {
        model: modelName, // または適宜選択。
        messages: [
          { role: "system", content: "あなたは生徒の学習計画管理をする真面目な塾講師です。" },
          { role: "user", content: prompt }
        ],
        functions: [functionDefinition],
        function_call: { name: "generatePlan" },
        max_tokens: 100,
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

      return options;
    },

    //プロンプトを取得する関数。
    _getPrompt(studyLog, valueToAdd, referenceBookData) {

      const materialDescription = this._formMaterialDescription(referenceBookData);
      const prompt = `
      あなたは生徒の学習計画管理をする真面目な塾講師です。
      あなたに計画を作成していただく教材の情報は以下です:
      ${materialDescription}

      最近の生徒の学習記録は"${studyLog}"です。
      次回の学習計画の課題量は厳密に"${valueToAdd}"です。この数は範囲が教材の限界に達しない限りは原則遵守すること。
      次の学習計画を学習実績を基準に、その範囲を次に進めることで作成してください。入力の学習記録と同一の学習計画を出力してはいけません。着実に先へ進む必要があります。
      このとき、今回の学習記録に含まれる名称・接頭語・単位は意味を保つ範囲で維持し、学習記録と同じ程度の文字数になるように簡潔にしてください。
      注意点として、学習実績には章やChapterの数字と、それより下位の問題番号や単語ナンバーの数字という複数種類の情報が入る可能性があります。
      
      出力の形式は { "plan": "<次の学習計画の文字列>" } です。
      出力は必ず JSON 形式で、余計な説明やコメントは一切含めない最小限の内容にしてください。

      <次の学習計画の文字列>の形式は、最終的に以下の5通りか、その末尾に最低限の学習実績に付随していた小さな補足や${GenWeeklyPlan_handlers_defaultTexts.onFinished}を追加したもののみです:
      1."課題範囲の開始 課題範囲開始の章の名前 ~ 課題範囲の終了 課題範囲終了の章の名前"
      2."課題内容 課題内容の名前"
      3."課題範囲の開始 ~ 課題範囲の終了"
      4."課題内容(3章など)"
      5."例外テキスト"

      例1: 
      生徒の学習実績が"5章 酸と塩基 ABC問題、すべて"で課題量が"2"で、教材が章を最小単位として進むならば、出力は{ "plan": "6章 酸化還元反応 ~ 7章 電池と電気分解 ABC問題、すべて" }となります。
      つまり、学習実績が5章なので、次に課題量だけ範囲をすすめると次の章である6が範囲の最初で、最後は6+2-1で7となり、それをもとの学習実績と同じように整形すると、"6章 酸化還元反応 ~ 7章 電池と電気分解 ABC問題"となります。
      このように、提示した教材の情報に基づく単元名（章名）の更新をしてください。
      入力文字列内の単元名（章名）として推測される部分は、教材の情報を利用して更新します。
      入力文字列に単元名がなくても、教材に単元名があるなら、それを必ず追記してください。

      例2:
      生徒の学習実績が"Chapter 3 単元3 ~ Chapter 4 単元2"で、課題量が"4"で、参考書の"番号の数え方"が"単元"の場合、出力は、{ "plan": "Chapter 4 単元3 ~ Chapter 4 単元6" }となります。
      参考書ごとの番号の連続性の違いがあります。章を通じて問題番号が一系統になっている場合はそのまま更新し、章ごとに番号がリセットされる場合は、各章内での番号として更新します。
      上記の例では現在の終了値は"Chapter 4 単元2"だから、新しい開始値はChapter 4 単元2 +1で、"Chapter 4 単元3"となります。
      そして、新しい終了値は、新しい開始値に課題量を足して、Chapter 4 単元3 +4 -1で"Chapter 4 単元6"となり、最終的に出力は、{ "plan": "Chapter 4 単元3 ~ Chapter 4 単元6" }となります。
      章をまたぐケースでも、累計問題数が課題量と厳密に一致するようにしてください。

      例3:
      例えば、生徒の学習実績が"例題:201~300"で、課題量が"150"で、参考書の最後の問題が"3章"で終わりの番号"400"の場合、出力は、{ "plan": "例題:301~400 ${GenWeeklyPlan_handlers_defaultTexts.onFinished}" }となります。
      上記の例のように、参考書マスターデータを見て、更新後の範囲が参考書の最後の範囲を超えないようにキャップしてください。
      参考書の範囲の最後まで到達したら、参考書の最後の問題を範囲の最後尾にしてください。

      例4:
      生徒の学習実績が"例題:101~350"で、課題量が"150"で、参考書の最後の問題が"7章"で終わりの番号"350"の場合、出力は、{ "plan": "${GenWeeklyPlan_handlers_defaultTexts.onFinished}2周目か次の参考書に進みましょう" }となります。
      このように、入力の方の範囲ですでに参考書の最後に到達済みなら、{ "plan": "${GenWeeklyPlan_handlers_defaultTexts.onFinished}2周目か次の参考書に進みましょう" }に置き換えてください。
      
      例5:
      例えば、生徒の学習実績が"3章~4章"で、課題量が"150"で、参考書の"3章"が初めの番号が"301"終わりの番号"400"の場合、出力は、{ "plan": "5章~6章" }となります。
      上記の例のように、提示した教材の情報を見て参考書は問題番号単位で進む構成なのに学習実績が章単位で指定されている場合は、前回の課題量を参考にしてください。このようなケースでは課題量はあくまで問題数ベースの目安であって、章の数ではない可能性が高いです。
      上記の例なら、"3章~4章"なので、4-3+1で前回の課題量は2です。
      このような指定された課題量は信用せず、前回と同じ量になるように指定してください。5から始まって、5+2-1で6までであり、そして章だから、"5章~6章"、場合によっては章の単元名も付与する、となります。
      ただし、このパターンは学習実績が"n章""Chapter n"のように、その教材でのもっとも細かい基準での範囲指定を伴わず、章などの大きな単位のみで指定されている場合に限り適用してください。例えば、"例題 101~200""No.5~20"のように数値範囲が明示されている場合は、必ず課題量を厳密に適用してください。
      `

      return prompt;
    },

    //参考書マスターの情報をAI向けに整形
    _formMaterialDescription(referenceBookData) {
      const chapPairCount = referenceBookData.chapters?.length || 0;
      const chapterPairNotice = chapPairCount >= 2 ?
        `(全${chapPairCount}章で対応あり)` :
        '(1章のみ)';

      const description = `
      - 基本情報
      教材名:${referenceBookData.name}
      教科:${referenceBookData.subject}
      - 構成${chapterPairNotice}
      章立て:[${referenceBookData.chapters}]
      章の名前:[${referenceBookData.chapterNames}]
      章のはじめ:[${referenceBookData.chapterStarts}]
      章の終わり:[${referenceBookData.chapterEnds}]
      番号の数え方:[${referenceBookData.numberingType}]
      `
      return description;
    },

    //AIの出力を正規化して出力用のstringに変換する関数。
    _sanitizeAIOutput(rawArgs) {
      /*
        中身を安全に取り出す関数（ネスト何階層でもOK）
        本来{"plan": "出力"}の形式になるが、たまに{"plan": {"plan": "出力"}}になってしまうようなので、一応複数回取りだせるかこころみる。
        さらに悪い時には、{"plan": \"{"plan": "出力"}\"}という風に中のオブジェクトを文字列にして出力されてしまうようなので、そのケースにも対応。
      */
      function unwrapObject(obj) {
        let target = obj;
        while (target && typeof target === 'object' && !Array.isArray(target)) {
          const values = Object.values(target);
          if (values.length === 1) {
            target = values[0];//中身を取り出す
            //取り出した中身が "{"plan" : "text"}"みたいなパースされていないJSONstringだった時にはパースする。
            if (typeof target === 'string' && target.trim().startsWith('{')) {
              try {
                target = JSON.parse(target);
              } catch (e) {
                return target;//パースできなかった場合はそのまま返す
              }
            }
          } else {
            break
          };
        }
        return target;
      }

      const argsObj = JSON.parse(rawArgs);//AIの出力をパースしてオブジェクトに
      const unwrappedStr = String(unwrapObject(argsObj));//オブジェクトをアンラップしてテキスト取り出し
      const trimmedStr = unwrappedStr.trim();//取り出したテキストをトリミング
      return trimmedStr;
    }


  },


  /**
   * 参考書マスターの情報なしでAIに頼んで処理してもらうパターン。
   * AI生成はカバーできるパターンがハンドラーの中で最多だが、api料金がかかり、生成時間も長めなので最終手段。
   */
  askingAIWithoutBookInfoHandler: {
    name: 'askingAIWithoutBookInfoHandler',
    run(context) {
      if (GenWeeklyPlan_handlers_flags.isAIBanned) return '★AI生成キャンセル';

      if (context.hasValidMetadata || !context.lastAchievement) return null;

      try {
        const result = this._generate(context);
        return result + GenWeeklyPlan_handlers_defaultTexts.onNoMasterData;
      } catch (e) {
        return context.lastAchievement + GenWeeklyPlan_handlers_defaultTexts.onAIError;
      }
    },

    _generate(context) {

      // 詳細なルールを含むプロンプト
      const prompt = this._getPrompt(context.lastAchievement, context.valueToAdd);

      const options = this._getOptionsObj(prompt, context);

      // APIを呼び出して結果を取得
      const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);

      //console.log(response)
      const responseData = JSON.parse(response.getContentText());

      // function_call の結果から arguments を抽出し、JSONをパースする
      const rawArgs = responseData.choices[0].message.function_call.arguments;
      const outputStr = this._sanitizeAIOutput(rawArgs);

      return outputStr;
    },

    //AIに渡すoptionオブジェクトを取得する関数。
    _getOptionsObj(prompt, context) {
      // APIキーを取得
      const apiKey = context.dependencies.getChatGPTAPIKEY();
      const modelName = context.dependencies.getChatGPTModelName();

      // function calling 用の関数定義（JSON スキーマ）
      const functionDefinition = {
        name: "generatePlan",
        description: "生徒の最新の学習実績に基づき、学習計画の出力文字列 plan を生成する。返り値は plan のみで、余計な説明やコメントは含めない。",
        parameters: {
          type: "object",
          properties: {
            plan: {
              type: "string",
              description: "更新された学習計画の文字列。余計な説明やコメントは含まない。",
            }
          },
          required: ["plan"]
        }
      };

      // API 呼び出し用ペイロードの作成
      const payload = {
        model: modelName, // または適宜選択。
        messages: [
          { role: "system", content: "あなたは生徒の学習計画管理をする真面目な塾講師です。" },
          { role: "user", content: prompt }
        ],
        functions: [functionDefinition],
        function_call: { name: "generatePlan" },
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

      return options;
    },

    //プロンプトを取得する関数。
    _getPrompt(studyLog, valueToAdd) {
      const prompt = `
      あなたは生徒の学習計画管理をする真面目な塾講師です。

      最近の生徒の学習記録は"${studyLog}"です。
      次回の学習計画の課題量は厳密に"${valueToAdd}"でなければなりません。
      次の学習計画を範囲を次に進めることで作成してください。
      このとき、今回の学習記録に含まれる名称・接頭語・単位は可能な限り維持し、学習記録と同じ程度の文字数になるようにしてください。

      単一の問題や項目ならば単純にその項目の呼び名を出力し、範囲を持っているものならば"開始範囲~終了範囲"の形式にしてください。
      開始点や進め方を推測できない場合は、必ず ${GenWeeklyPlan_handlers_defaultTexts.onConsultationNeeded} を出力してください。
      
      出力の形式は { "plan": "<次の学習計画の文字列>" } です。
      出力は必ず JSON 形式で、余計な説明やコメントは一切含めない最小限の内容にしてください。

      例1:
      学習記録"例題101~200"で課題量"150"ならば、次の問題番号は201で、課題量が150より範囲の最後は201+150-1=350なので、出力は{ "plan": "例題201~350" }

      例2:
      学習記録"1章"で課題量"2"ならば、次の章は2章であり、課題量は2より終わりの章は2+2-1=3なので、出力は{ "plan": "2章~3章" }

      例3:
      学習記録"1日3テーマ"で課題量"2"ならば、前回の学習実績が範囲ではなく課題量で定義されているので、計画の出力は、課題量で定義して、{ "plan": "1日2テーマ" }

      例4:
      学習記録"2024年 数学"で課題量"7"ならば、過去問演習であり進め方が不明なので、引き続き計画の出力は{ "plan": "★" } 

      例5:
      学習記録"No.1~100"で課題量"1"ならば、次の問題番号は101で課題量は1だけなので、計画の出力は{ "plan": "No.101" } 
      `;
      return prompt;
    },

    //AIの出力を正規化して出力用のstringに変換する関数。
    _sanitizeAIOutput(rawArgs) {
      /*
        中身を安全に取り出す関数（ネスト何階層でもOK）
        本来{"plan": "出力"}の形式になるが、たまに{"plan": {"plan": "出力"}}になってしまうようなので、一応複数回取りだせるかこころみる。
        さらに悪い時には、{"plan": \"{"plan": "出力"}\"}という風に中のオブジェクトを文字列にして出力されてしまうようなので、そのケースにも対応。
      */
      function unwrapObject(obj) {
        let target = obj;
        while (target && typeof target === 'object' && !Array.isArray(target)) {
          const values = Object.values(target);
          if (values.length === 1) {
            target = values[0];//中身を取り出す
            //取り出した中身が "{"plan" : "text"}"みたいなパースされていないJSONstringだった時にはパースする。
            if (typeof target === 'string' && target.trim().startsWith('{')) {
              try {
                target = JSON.parse(target);
              } catch (e) {
                return target;//パースできなかった場合はそのまま返す
              }
            }
          } else {
            break
          };
        }
        return target;
      }

      const argsObj = JSON.parse(rawArgs);//AIの出力をパースしてオブジェクトに
      const unwrappedStr = String(unwrapObject(argsObj));//オブジェクトをアンラップしてテキスト取り出し
      const trimmedStr = unwrappedStr.trim();//取り出したテキストをトリミング
      return trimmedStr;
    }


  }
})
