//00[interface]_main.gs : 直接呼ばれる関数を置いておく場所

/**
 * @typedef {Object} RunFunctionRequest
 * @property {string} id - 実行したい関数を指定するid。
 * @property {Object} [context] - 実行されるときに関数に渡したい引数を詰め込んだオブジェクト。
 */

/**
 * 任意の関数を実行できるようにするための抽象関数。
 * オブジェクト idObjを受け取り、idとcontextを取り出してFunctionRunnerServiceに渡す。
 *
 * 呼び出せる関数のレパートリーを変更する方法はDocumentを参照して下さい。
 * @param {RunFunctionRequest} idObj - id{string} と context{object} (contextはなくてもよい)が入った、関数実行用のオブジェクト。
 */
function abstractFunction(idObj) {
  //idObjのバリデーション
  if (!idObj || typeof idObj !== "object") {
    throw new Error(
      "abstractFunctionの引数はオブジェクトである必要があります。\n詳しい実装については00[Document].gsをご覧ください。",
    );
  }

  //idObjの中身のバリデーション
  const id = idObj.id; //関数特定用のid
  const context = idObj.context; //関数に渡す用のコンテキストオブジェクト

  if (!id || typeof id !== "string") {
    throw new Error(
      `渡されたオブジェクトに適切なid(string)が入っていません。:\n${JSON.stringify(idObj)}`,
    );
  }

  if (context !== undefined && typeof context !== "object") {
    throw new Error(
      `渡されたオブジェクトに不適切なcontext{object}が入っていました。contextはオブジェクトにするか、そもそも入れないかのいずれかにしてください。:\n${JSON.stringify(idObj)}`,
    );
  }

  const runner = new FunctionRunnerService(); //関数実行用のクラスを初期化して、コールバックをセットしてから実行。
  runner.setOnFinishedListener(() => SidebarMenuService.show());
  runner.setOnStartedListener(() =>
    ToastNotificationService.send("スクリプトが実行開始しました。"),
  );
  runner.setOnSucceededListener(() =>
    ToastNotificationService.send("スクリプトが実行終了しました。"),
  );

  runner.run(String(id).trim(), context);
}

//abstractFunction(idObj)のテスト用関数
function test_abstractFunction() {
  const idObj = {
    id: "geshoWithGantt",
  };
  abstractFunction(idObj);
}

/**
 * 生徒のスピードプランナーのonOpenで呼ばれる関数。
 * ここで呼ばれる関数はこのライブラリ内のものなので違和感があるかもしれませんが、呼び出し元の目線で定義する必要があるので、
 * きちんと「Gantt.」接頭辞をつけて関数名を登録する必要があります。これは各生徒のスピードプランナーのスクリプト内でのこのライブラリのインポート名です。
 * このインポート名はどの呼び出しもとでも共通しているはずです。複数種類への対応は考慮する必要はありません。
 */
function onOpenAction() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu1 = ui.createMenu("マクロ一覧");
    menu1.addItem("表示", "Gantt.showSidebarMenu");
    menu1.addToUi();
    const menu2 = ui.createMenu("週間計画作成");
    menu2.addItem("第一週", "Gantt.week1Plan");
    menu2.addItem("第二週", "Gantt.week2Plan");
    menu2.addItem("第三週", "Gantt.week3Plan");
    menu2.addItem("第四週", "Gantt.week4Plan");
    menu2.addItem("第五週", "Gantt.week5Plan");
    menu2.addToUi();
    const menu3 = ui.createMenu("ヒアリング項目作成");
    menu3.addItem("実行", "Gantt.genHearingItems");
    menu3.addToUi();
  } catch (e) {
    const message = e.message || String(e); // ユーザー向け（簡潔）
    const detail = e.stack || "No stack trace"; // 開発者向け（詳細）
    // ユーザーにはトーストで簡潔に（エラーがあることだけは伝える）
    ToastNotificationService.send(
      "メニューの読み込みに失敗しました: " + message,
    );
    // 開発者はログ（コンソール）で詳細を確認
    console.error("onOpenAction Error Details:\n" + detail);
  }
}
