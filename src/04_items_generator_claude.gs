class GenHearingItems_ItemsGeneratorWithClaude {

  static get _apiKey() { return PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY'); }
  static get _apiUrl() { return 'https://api.anthropic.com/v1/messages'; }
  static get _apiVersion() { return '2023-06-01'; }
  static get _modelName() { return 'claude-sonnet-4-6'; }

  // _basePromptだけは非常に長いのでクラスの末尾にgetterを定義しています。

  static generate(context) {
    if (!context || typeof context !== 'object') {
      throw new Error(`contextが不正です。 typeof=${typeof context}, 内容=${String(context)}`);
    }
    const prevItemsStr = String(context.prevItemsStr ?? '');
    const lastTeachingMemo = String(context.lastTeachingMemo ?? '');
    // 既存データから内部コンテキスト生成
    const userContent = this._constructUserContent(prevItemsStr, lastTeachingMemo);
    //　AIの返答を文字列化してゲット
    const responseText = this._askAndReceive(userContent);
    // AIの返答の文字列を項目ごとに分割して配列にして返却
    return this._splitAIResponseIntoItemsArr(responseText);
  }

  static _constructUserContent(prevItemsStr, lastTeachingMemo) {
    return [
      `# 前回の生成内容(全文です)\n\n${prevItemsStr}`,
      `# 面談メモ\n\n${lastTeachingMemo}\n\n## 出力`
    ].join('\n\n');
  }

  static _askAndReceive(userContent) {
    const requestBody = {
      model: this._modelName,
      max_tokens: 2048,
      system: this._basePrompt,
      messages: [
        { role: 'user', content: userContent }
      ]
    };

    const options = {
      method: 'POST',
      muteHttpExceptions: true,
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': this._apiKey,
        'anthropic-version': this._apiVersion,
      },
      payload: JSON.stringify(requestBody),
    };

    const httpRes = UrlFetchApp.fetch(this._apiUrl, options);
    const status = httpRes.getResponseCode();
    const text = httpRes.getContentText();

    let json;
    try {
      json = JSON.parse(text);
    } catch (e) {
      throw new Error(`Claude response is not JSON. status=${status}, body=${text}`);
    }
    if (status < 200 || status >= 300) {
      throw new Error(`Claude API error. status=${status}, body=${JSON.stringify(json)}`);
    }

    // content[].text を結合して返す
    const chunks = (Array.isArray(json.content) ? json.content : [])
      .filter(c => c?.type === 'text' && typeof c.text === 'string')
      .map(c => c.text);

    const joined = chunks.join('\n').trim();
    if (!joined) throw new Error(`Unexpected Claude API payload: ${JSON.stringify(json)}`);
    return joined;
  }

  static _splitAIResponseIntoItemsArr(response) {
    if (typeof response !== 'string') throw new Error(`不正な型：${typeof response}`);
    return response
      .split('\n')
      .map(line => line
        .replace(/^\s*(?:[-*•]\s*|\d+[.)]\s*|\(\d+\)\s*)/, '')
        .trim()
      )
      .filter(v => v.length >= 6);// 5文字以下はさすがにエラー
  }


  static get _basePrompt() {
    return `
あなたは優秀な大学受験予備校の講師です。コーチング手法に基づく適切な質問文を作ってもらいます。毎回、オープンクエスチョン形式の質問を7つ作成してもらいます。質問の字数は、最小で20文字、最大で40文字までです。これから付加される生徒の情報を考慮して、質問を出力してください。質問は、7つで生徒の幅広い情報を得ることができるようにしてください。

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

## 「生成内容はこちらです」というような案内は生成しないでください。機械のように箇条書きで項目だけを生成してください。それでは、質問の生成をお願いします。：
`;
  }
}