//03_handlers.gs
/**
 * # 次回の計画テキストを生成するメソッドとして働くハンドラー群を管理するファイル。
 * 
 * 実際のハンドラーは、03_handlers_....gsという名前のファイルに定義されている。
 * このファイルに定義される内容は、それらが共有する定数やフラグ、あるいはハンドラーにアクセスするための関数。
 * 
 * ハンドラーの追加/削除をする場合には以下も把握してください。
 * 
 * ### 各ハンドラーが共有すべき特徴
 * 
 * - 各ハンドラーはnameプロパティを持ち、これには自身の識別名を文字列として登録しておく。デバッグ時の確認用。
 * 
 * - 各ハンドラーはrunメソッドを持ち、runの引数はcontextという各種データがパッキングされたオブジェクト一つのみ。
 * 
 * - runは、その中身のうち必要なものを見て、自分が扱えるタイプのデータか判定する。
 *  このとき、なるべく判定ロジックは自分単体で完結させること。
 *  つまり、自身より前のハンドラーが、自分自身では扱うことのできないパターンを扱ってくれるからという理由でエラー防止を怠ってはならない。
 *  ただし生成エラーに繋がらない条件についてはハンドラー内で完結させなくてよい。それはハンドラー間の優先順位で振り分ける。
 * 
 * - 扱えないならnullを返却する。扱えるなら次回の計画を自分で生成して、その計画のテキストをstring型で返す。
 * 
 * 
 * ### 技術的負債・予定
 * 
 * - なぜ自分が生成できなかったか、の理由も返却するようにするとよりデバッグしやすくなるので、そうするかもしれない。
 *  {result, wasGenerationSucceded, logs} のような。今はデバッグ機能をonにしても、結果と最終的に生成したハンドラの名前の特定までしかできない。
 * 
 * - 問題番号が章ごとにリセットされるタイプの問題については問題番号がかならず1から始まる想定でハンドラーを設計してしまっている。(章は0スタートでも問題ない)今後、この仕様のせいでバグが発生した場合には、該当ハンドラーを修正する必要がある。AIにハンドラーをすべて見せればどれかはすぐに特定できると思います。
 * 
 */

/**
 * @typedef {Object} GenWeeklyPlanHandler
 * @property {string} name - ハンドラーの識別名。理由がなければ全体オブジェクト名と同一
 * @property {(context : GenWeeklyPlanContext) => string} run - 生成実行用の関数。contextという一つのオブジェクトのみを受け取る。
 */
/**
 * @typedef {Object<GenWeeklyPlanHandler>} GenWeeklyPlanHandlersContainer
 */


/**
 * 特殊な生成時に設定するフラグをまとめるオブジェクト。
 * 
 * isAIBanned : AIを使用停止するフラグテスト時にAIを無駄に利用してapi使用料を浪費したくないときにtrueに。
 * 
 * pushLogAboutWhatHandlerGenerated : 生成時にどのハンドラーで生成された結果なのかを結果に結合して出力するかどうかを出力するようになるか。デバッグ用。
 */
const GenWeeklyPlan_handlers_flags = {
  isAIBanned: false,
  pushLogAboutWhatHandlerGenerated: false
}

/**
 * 種々の状況で使用する規定テキストを格納するオブジェクト。
 */
const GenWeeklyPlan_handlers_defaultTexts = Object.freeze({
  onError: '★',//失敗したとき全般
  onConsultationNeeded: '★',//相談が必要な時
  onNoMasterData: '★参考書マスターを更新してください',//参考書マスターにデータなし
  onCanceled: '',//エラーではなく意図的に生成をしないとき(次回の学習予定時間が0など)
  onFinished: '★完了！',//教材の範囲が終了しているので入力できないとき
  onAIError: '★', //特にAIで（通信障害など）問題が発生したため断念したとき
})

/**
 * ハンドラーたちを利用して次回の計画を作成するクラス。
 * 各ハンドラーの優先順位も管理する。
 */
class GenWeeklyPlan_handlers_generator {
  /**
   * contextを受け取って、次回の計画を実際に生成するメソッド。
   * @param {GenWeeklyPlanContext} context - 各ハンドラーにとって十分な情報を詰め込んだオブジェクト。
   * 
   * @returns {string} - 生成結果。
   */
  static generate(context) {
    
    const handlers = this._getHandlersArray();
    this._validateHandlers(handlers);
    const willTellWhatHandler = GenWeeklyPlan_handlers_flags.pushLogAboutWhatHandlerGenerated;

    for (const handler of handlers) {//ハンドラーを一個ずつ実行して、結果がnullでなくなったら終了。
      const result = handler.run(context);
      if (typeof result === 'string') {
        if (willTellWhatHandler) {
          const handlerName = handler.name;
          this._tryPushLog(context, result, handlerName)
          Logger.log(`使用されたハンドラー: ${handlerName}`);
          return result + ':' + handlerName;
        }
        return result; // ここで関数全体から結果を返して終了する
      }
    }

    return GenWeeklyPlan_handlers_defaultTexts.onError;//ここまで全部失敗した場合
  }

  static _tryPushLog(context, result, handlerName){
    try{
      console.info(`週間計画作成\ncontext:\n`,context,'\n\nresult: ',result,'\n\nhandler name: ',handlerName);
    }catch(e){
      console.error(e);
    }
  }


  /**
   * ハンドラーをすべて適用優先度の順番に取得する関数。
   * 優先順位を厳密に管理するため、ここで配列リテラルで明示的に定義します。
   * 
   * ### 補足
   * 
   * 各種ハンドラーは自身が扱えるなら結果を出し、扱えないならnullをだすということしかしないので、
   * 他の、より適任のハンドラーに処理を譲るという判断はしません。
   * そこで、ここで配列内での順によって適用の優先順位を設定しています。
   * 最初のハンドラーから順に仕事をし、最初に扱えるハンドラーに出会った時点でそのハンドラーの生成結果を最終出力として採用する。
   * 
   * @returns {GenWeeklyPlanHandler[]} ハンドラーの配列
   * 
   * @private
   */
  static _getHandlersArray() {
    if (!!this._handlersArray) return this._handlersArray;//キャッシュにあるならそれを返す
    const normalHanders = GenWeeklyPlan_handlers_general;
    const aiHandlers = GenWeeklyPlan_handlers_withAI

    // 優先順位が高い順（特殊ケース -> 一般ケース -> AI）に並べています
    const handlersArray = [
      normalHanders.noStudyHandler,
      normalHanders.emptyHandler,
      normalHanders.noNumberToAddHandler,
      normalHanders.multipleTildasHandler,
      normalHanders.pastExamsHandler,
      normalHanders.rangeOfAnotherTextsHandler,

      normalHanders.probNumberRangeHandler,
      normalHanders.probSingleNumberHandler,
      normalHanders.chapProbRangeHandler,
      aiHandlers.askingAIWithBookInfoHandler,
      aiHandlers.askingAIWithoutBookInfoHandler
    ];
    this._handlersArray = Object.freeze(handlersArray);//キャッシュに格納
    return this._handlersArray;
  }

  /**
   * generateメソッドで使用される各ハンドラーが、満たすべき構造の条件を満たしているかどうかの簡易的なチェックを行う関数。
   * 計画の生成結果までは確認しない。異常があればエラーを出す。検証にかかる時間は5ms未満。
   * 
   * @param {GenWeeklyPlanHandler[]} targetHandlers - 検証対象のハンドラーの配列
   * 
   * @throws {Error} ハンドラーの設定に異常があればエラーを出す。
   * 
   * @private
   */
  static _validateHandlers(targetHandlers) {

    const seenHandlersSet = new Set();//ハンドラーの重複チェックに使う、すでにみた名前のリスト。
    const seenNamesSet = new Set();//名前の重複チェックに使う、すでにみた名前のリスト。
    const errors = [];//エラーをためていく配列

    function checkName(itemName) {//ハンドラーのnameプロパティのチェック用関数
      if (itemName === undefined) {
        errors.push(`nameプロパティが未定義なハンドラーが検出されました。:${JSON.stringify(item)}`);
        return;
      }
      if (typeof itemName !== 'string') {
        errors.push(`nameプロパティがstringでないハンドラーが検出されました。:${JSON.stringify(item)}`);
        return;
      }
      if (seenNamesSet.has(itemName)) {
        errors.push(`nameプロパティが他と重複したハンドラーが検出されました。ハンドラー複製時の名前変更ミスなどが疑われます:${JSON.stringify(item)}`);
      }
      seenNamesSet.add(itemName);
    }

    function checkRunFunc(runFunc) {//ハンドラーのrunというfunctionのチェック用関数
      if (runFunc === undefined) {
        errors.push(`runFuncプロパティが未定義なハンドラーが検出されました。:${JSON.stringify(item)}`);
        return;
      }
      if (typeof runFunc !== 'function') {
        errors.push(`runプロパティがfunctionでないハンドラーが検出されました。:${JSON.stringify(item)}`);
        return;
      }
      if (runFunc.length !== 1) {//これで引数の数を検出できる
        errors.push(`runプロパティのlengthが1でない(つまり取る引数がcontextの一つのみでない)ハンドラーが検出されました。:${JSON.stringify(item)}`);
      }
    }

    //それぞれのハンドラーについて調査
    targetHandlers.forEach(item => {
      if (!item || typeof item !== 'object') {
        errors.push(`オブジェクトでないハンドラーが検出されました。:${JSON.stringify(item)}`);
        return;
      }
      if (seenHandlersSet.has(item)) {
        errors.push(`複数回_getHandlersArray()内に登録されているハンドラーが検出されました。:${JSON.stringify(item)}`)
      }
      seenHandlersSet.add(item);
      checkName(item.name);
      checkRunFunc(item.run);
    })

    if (errors.length !== 0) throw new Error(`エラー\n${errors.join('\n')}`);

  }
}

