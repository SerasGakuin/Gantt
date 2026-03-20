//00_FunctionRunnerService.gs

function test_FunctionRunnerService() {
  const runner = new FunctionRunnerService();
  runner.setOnFinishedListener(() => {
    Logger.log("すべての処理が正常、または例外処理を経て終了しました。");
  });
  runner.run("reportError", { message: "テスト用メッセージです。" });
}
/**
 * # FunctionRunnerService
 *
 * このライブラリの関数をid(関数の識別用文字列)とcontext(関数に渡したい引数を詰め込んだオブジェクト)で実行するクラス。newで初期化して使用する。
 * 同時実行防止と、関数の実際の実行開始と、ユーザーへのエラーなどの通知を責務とする。
 * contextが不要ならここで消滅させる役割も果たしている。
 *
 *
 * ### 構成
 *
 * - FunctionRunnerService
 *
 * api用のクラス。インスタンス化して、コールバックを設定してgenerateメソッドにfunctionIdとcontextを渡すことで実行できる。contextがnullでよいか、などは00_FunctionConfigs.gsの定義に依存し、関数によって、つまり関数と一対一対応するfunctionIdごとに違う。
 * また、これだけがこのファイル外の00_番台からアクセスされうるapi用のクラス。他は補助である。
 *
 *
 * - FunctionRunnerService_IdToFuncMapFactory
 *
 * 00_FunctionConfigs.gsの定義に依存してfunctionId(string) to function(function) のマップオブジェクトを生成するクラス。
 * マップの生成時には、00_FunctionConfigs.gsで、FunctionConfigsやContextFunctionConfigsが正しく設定されていることを仮定しておこなう。
 *
 *
 * - FunctionRunnerService_IdToFuncMapReader
 *
 * functionId(string) to function(function)のマップオブジェクトを読解して、対応するfunctionを返却するクラス。
 * もし対応するfunctionがマップに見つからなかった時にはデバッグ時に役立つエラーを出す。
 *
 *
 *
 * ### 使用例
 * ========================================================================================================
 * const runner = new FunctionRunnerService();
 * runner.setOnFinishedListener(() => SidebarMenuService.show());
 * runner.setOnStartedListener(() => ToastNotificationService.send('スクリプトが実行開始しました。'));
 * runner.setOnSucceededListener(() => ToastNotificationService.send('スクリプトが実行終了しました。'));
 *
 * runner.run(functionId, context);
 * ========================================================================================================
 *
 *
 *
 * ### 技術的負債・注意点
 *
 * - contextが必要な関数に対して関数を実行するときに、contextの渡し忘れを自動検出する方法がない。context不要な関数のほうが実務上圧倒的多数ではあるが...
 *
 * 1. ContextFunctionConfigsで必要なcontextの中身を指定して、それを満たさないものが来たらエラーをだすように仕様変更
 * 2. contextが必要なケースは本当に少ないので、後任の方がcontextが必要な関数を追加したい場合には実装を理解してからやってもらえばよいとする(AIも使える)
 *
 * といった方針がありうるように思う。
 *
 *
 */
class FunctionRunnerService {
  /**
   * 各イベントリスナーの設定。関数を実行し、実行開始時、実行失敗時などにやってほしい処理があれば、これらの関数に引数のないfunctionを渡してからrunを実行すればそれぞれに与えられたタイミングで実行してくれる。
   */
  /** @param {function} func - 引数なしの関数 */
  setOnStartedListener(func) {
    this._validateCallBackFunc(func);
    this._onStartedListener = func;
  } //関数実行開始時
  /** @param {function} func - 引数なしの関数 */
  setOnSucceededListener(func) {
    this._validateCallBackFunc(func);
    this._onSucceededListener = func;
  } //関数実効成功時
  /** @param {function} func - 引数なしの関数 */
  setOnFailedListener(func) {
    this._validateCallBackFunc(func);
    this._onFailedListener = func;
  } //関数実行失敗時
  /** @param {function} func - 引数なしの関数 */
  setOnFinishedListener(func) {
    this._validateCallBackFunc(func);
    this._onFinishedListener = func;
  } //関数実行終了時（成功/失敗に関係なし）

  /**
   * 指定された関数 ID に基づき、対応する処理を実行する関数。
   * @param {string} functionId - functionsMap に定義された実行対象関数のキー。
   * @param {Object} [context] - 実行関数に渡す任意のコンテキストデータ。
   * @throws {Error} ロックの取得に失敗した場合、または不正な functionId の場合にエラーを投げます。
   * @public
   */
  run(functionId, context = {}) {
    const lock = LockService.getScriptLock();

    try {
      if (!lock.tryLock(30000))
        throw new Error(
          "FunctionRunnerService:他の処理が実行中です。時間を置いて再度お試しください。",
        );

      this._safeExecuteCallback(this.onStartedListener); //開始コールバック

      //NOTE: この二つのクラスはこのファイルにある
      const functionMap = FunctionRunnerService_IdToFuncMapFactory.generate();
      const targetFn = FunctionRunnerService_IdToFuncMapReader.findFunctionFrom(
        functionMap,
        functionId,
      );

      try {
        this._runFunction(targetFn, context);
      } catch (e) {
        this._sendErrorMessage(e, functionId);
        throw e;
      }

      this._safeExecuteCallback(this.onSucceededListener); //成功時コールバック
    } catch (e) {
      this._safeExecuteCallback(this.onFailedListener); //失敗時コールバック
      throw e;
    } finally {
      this._forceUnlock(lock);
      this._safeExecuteCallback(this.onFinishedListener); //終了時コールバック
    }
  }

  /**
   * コールバックを実行しつつ、関数本体の実行は妨げないようにする関数。
   * @param {function} func - 実行するコールバック関数
   * @private
   */
  _safeExecuteCallback(func) {
    try {
      func();
    } catch (e) {
      console.error("Callback Error:", e);
    }
  }

  /**
   * 渡されたロックを解除する関数。
   * @param {ScriptLock} lock - アンロックしたいもの
   * @private
   */
  _forceUnlock(lock) {
    if (!!lock) {
      try {
        lock.releaseLock();
      } catch (e) {
        console.error(`アンロックに失敗:\n\n${e}`);
      }
    }
  }

  /**
   * 関数実行する関数
   * @param {function|null} functionId - 実行したい関数
   * @param {object|null|undefined} context - 実行したい関数に渡したい引数をつめこんだオブジェクト。
   * @private
   */
  _runFunction(targetFn, context) {
    //関数実行
    FunctionRunnerService._runCount =
      (FunctionRunnerService._runCount || 0) + 1; //実行回数の初期化またはインクリメント
    if (FunctionRunnerService._runCount > 100) {
      throw new Error(
        `FunctionRunnerServiceが今回の実行時間中に100回以上呼ばれています！再帰ループまたは過剰なループによる無限実行が疑われます！FunctionRunnerServiceによって実行される関数の中でFunctionRunnerService.runを実行しないでください！`,
      );
    }
    if (!!targetFn && typeof targetFn === "function") targetFn(context); //実行
  }

  /** @returns {number} */
  static get _runCount() {
    return FunctionRunnerService.__runCount;
  }
  /** @param {number} countNum */
  static set _runCount(countNum) {
    FunctionRunnerService.__runCount = countNum;
  }

  /**
   * エラー送信用関数
   * @param {Error} e - エラーオブジェクト
   * @param {string} functionId - それを使って関数を実行したかったid
   * @private
   */
  _sendErrorMessage(e, functionId) {
    const msg = (e && e.stack) || (e && e.message) || e.toString();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const name = spreadsheet.getName();
    const splitter =
      "==================================================================";
    const finalMsg = `${msg}\n${splitter}\n対象:${name}\nfunctionId:${functionId}`;
    GASRefferenceSheetLogService.error(finalMsg);
    PopupAlertService.send(finalMsg);
  }

  /**
   * コールバック用の関数をバリデーションする関数。不適ならエラーを出すだけ。
   * @private
   */
  _validateCallBackFunc(func) {
    if (typeof func !== "function" || func.length !== 0)
      throw new Error("引数のない関数を渡してください。");
  }

  // コールバックの関数たちのgetter
  get onStartedListener() {
    return this._onStartedListener || (() => {});
  }
  get onFinishedListener() {
    return this._onFinishedListener || (() => {});
  }
  get onSucceededListener() {
    return this._onSucceededListener || (() => {});
  }
  get onFailedListener() {
    return this._onFailedListener || (() => {});
  }

  set onStartedListener(func) {
    this._validateCallBackFunc(func);
    this._onStartedListener = func;
  } //関数実行開始時
  set onSucceededListener(func) {
    this._validateCallBackFunc(func);
    this._onSucceededListener = func;
  } //関数実効成功時
  set onFailedListener(func) {
    this._validateCallBackFunc(func);
    this._onFailedListener = func;
  } //関数実行失敗時
  set onFinishedListener(func) {
    this._validateCallBackFunc(func);
    this._onFinishedListener = func;
  } //関数実行終了時（成功/失敗に関係なし）
}

//========================================== IdToFuncMapFactory ====================================================

function test_FunctionRunnerService_IdToFuncMapFactory() {
  const map = FunctionRunnerService_IdToFuncMapFactory.generate();
  console.log("生成されるfunctionId to functionマップ\n", map);
}
/**
 * functionId(string) to function(function) のマップオブジェクトを生成するクラス。
 */
class FunctionRunnerService_IdToFuncMapFactory {
  /**
   * マップを生成する関数。
   * @returns {object} mapObj - functionId(string) to function(function) のマップオブジェクト
   */
  static generate() {
    if (!!this._mapObj) return this._mapObj;
    try {
      const mapObj = Object.freeze(this._generateFunctionsMap()); //設定からマップオブジェクトを生成
      this._mapObj = mapObj; //キャッシュ
      Object.freeze(this); //この工場はグローバル初期化時点で生成物が決まっているので、マップをキャッシュした今、安定性のために凍結する
      return mapObj;
    } catch (e) {
      throw new Error(`FunctionRunnerService_IdToFuncMapFactory:\n${e}`);
    }
  }

  /**
   * ハードコードされた設定から、自動的にfunctionId => 関数 のマップを生成する関数。
   * @private
   */
  static _generateFunctionsMap() {
    const map = {};

    //関数設定から自動的にid=>関数のマップを作成 (context不要版)
    FunctionConfigs.forEach((obj) => {
      const func = obj.func;
      const functionId = obj.id;
      map[functionId] = (context) => func();
    });

    //関数設定から自動的にid=>関数のマップを作成 (context必要版)
    ContextFunctionConfigs.forEach((obj) => {
      const func = obj.func;
      const functionId = obj.id;
      map[functionId] = (context) => func(context);
    });

    return map;
  }
}

//========================================== IdToFuncMapReader ====================================================

function test_FunctionRunnerService_IdToFuncMapReader() {
  const functionId = "genHearingEntries"; //ここを変えてテスト

  const map = FunctionRunnerService_IdToFuncMapFactory.generate();
  console.log("生成されたマップ\n", map);
  const fn = FunctionRunnerService_IdToFuncMapReader.findFunctionFrom(
    map,
    functionId,
  );
  console.log("ヒットした関数のnameプロパティ\n", fn?.name);
}
/**
 * functionId(string) to function(function)のマップオブジェクトを読解して、対応するfunctionを返却、
 * もし対応するfunctionが見つからなかった時にはデバッグ時に役立つエラーを出すクラス。
 */
class FunctionRunnerService_IdToFuncMapReader {
  /**
   * @param {object} mapObj - 読解対象のマップ。
   * @param {string} functionId - 対応する関数を見つけたいid。
   * @returns {function} 対応する関数。
   */
  static findFunctionFrom(mapObj, functionId) {
    try {
      if (typeof mapObj !== "object")
        throw new Error(
          `findFunctionFrom(mapObj, functionId)に渡されたmapObjの型が不正です。\n渡されたもの:${mapObj}, 型:${typeof mapObj}`,
        );
      if (typeof functionId !== "string")
        throw new Error(
          `findFunctionFrom(mapObj, functionId)に渡されたfunctionIdの型が不正です。\n渡されたもの:${functionId}, 型:${typeof functionId}`,
        );

      const fn = mapObj[functionId];
      if (typeof fn !== "function")
        this._pushDetailLogOnNotFoundFunc(functionId, mapObj); //idが非対応の時

      return fn;
    } catch (e) {
      throw new Error(`FunctionRunnerService_IdToFuncMapReader: \n${e}`);
    }
  }

  /**
   * 対応する関数が見つからなかった時のデバッグのために、エラーログを出す関数。
   * @param {string} target - もっとも近いidを検証したい対象。
   * @param {object} functionMap - idと関数の対応を定義したオブジェクト。
   * @throws {Error} この関数は実行されると、デバッグ用の情報を載せたエラーを出します。
   * @private
   */
  static _pushDetailLogOnNotFoundFunc(functionId, mapObj) {
    const closestIds = this._getClosestIdsList(functionId, mapObj);
    const similarGlobalFuncs = this._getSimilarNameGlobaleFuncs(functionId);

    const invalidIdView = `\n無効なfunctionId:\n'${functionId}'\n`;
    const closestCorrectIdsView = !!closestIds
      ? `\n有効なidのうち最も近いもの(簡易的なサジェストです):\n${closestIds}\n`
      : "";
    const similarGlobalFuncsView = !!similarGlobalFuncs
      ? `\nわかりやすさのために設定オブジェクトのidは対応する関数の定義名と一致させることが推奨されます。\n打ち間違えていませんか？\nあるいは、設定を登録し忘れていませんか？\n\nfunctionIdと名前の似ているまたは一致するグローバル関数(この検出は不安定なのでご注意ください。):\n${similarGlobalFuncs}`
      : "";

    throw new Error(
      invalidIdView + closestCorrectIdsView + similarGlobalFuncsView,
    );
  }

  /**
   * _functionsMapに登録されたidのうち、渡された文字列にレーベンシュタイン距離のもっとも近いものたちを配列としてまとめて、結合して返却する関数。
   * デバッグのときのtypo検出用
   * @param {string} target - もっとも近いidを検証したい対象。
   * @param {object} functionMap - idと関数の対応を定義したオブジェクト。
   * @returns {string | null} 渡されたfunctionIdに最も近い、有効なidの一覧文字列
   * @private
   */
  static _getClosestIdsList(target, functionMap) {
    try {
      const targetStr = String(target);
      const correctIdsArr = Object.keys(functionMap);
      if (correctIdsArr.length === 0) return null;

      //idの距離を配列に格納
      const idsDistancesArr = correctIdsArr.map((v) =>
        this._getLevenshteinDistance(v, targetStr),
      );

      //最短距離を記録
      const minDist = Math.min(...idsDistancesArr);

      //遠すぎるならnull
      const dynamicThreshold = Math.min(
        Math.max(3, Math.floor(targetStr.length / 2)),
        7,
      );
      if (minDist > dynamicThreshold) return null;

      //最短のidたち
      const closestIds = correctIdsArr.filter(
        (_, i) => idsDistancesArr[i] === minDist,
      );
      if (closestIds.length === 0) return null;

      return `[ ${closestIds.join(", ")} ]`;
    } catch (e) {
      //この機能はあくまで開発補助なので、ここで問題が起きた場合は記録だけ
      console.error(
        `渡されたidに近い正しいidをサジェストする関数「_getSimilarNameGlobaleFuncs」でエラーが発生しました。\n${e}`,
      );
    }
  }

  /**
   * グローバルに渡されたidと名前の似た関数があった場合、それの存在を提示する関数。
   * NOTE: globalThisの仕様は不安定なのであまり信頼しないように！functionId to functionのマップをglobalThisから自動生成しないのもその不安定さが理由です
   *
   * @param {string} target - もっとも近いidを検証したい対象。
   * @returns {string | null} 渡されたfunctionIdに最も近い関数名(例外が発生したらnullを返す)
   * @private
   */
  static _getSimilarNameGlobaleFuncs(target) {
    try {
      const targetStr = String(target);
      const globalObj = globalThis;
      if (!globalObj) return null;

      // GASのグローバル関数を抽出
      const globalFuncsNamesArr = Object.keys(globalObj).filter(
        (key) => typeof globalObj[key] === "function",
      );

      if (globalFuncsNamesArr.length === 0) return null;

      const distances = globalFuncsNamesArr.map((v) =>
        this._getLevenshteinDistance(v, targetStr),
      );
      const minDist = Math.min(...distances);

      // 距離が遠すぎる（全く似ていない）ものはサジェストしない
      const dynamicThreshold = Math.min(
        Math.max(3, Math.floor(targetStr.length / 2)),
        7,
      );
      if (minDist > dynamicThreshold) return null;

      const closestNames = globalFuncsNamesArr.filter(
        (_, i) => distances[i] === minDist,
      );
      return `[ ${closestNames.join(", ")} ]`;
    } catch (e) {
      //この機能はあくまで開発補助なので、ここで問題が起きた場合は記録だけ
      console.error(
        `グローバルに渡されたidと名前の似た関数があった場合、それの存在を提示する関数「_getSimilarNameGlobaleFuncs」でエラーが発生しました。\n${e}`,
      );
    }
  }

  /**
   * 二つの文字列間のレーベンシュタイン距離（編集距離）を測定します。
   * @param {string} s1Input - 比較対象の文字列1
   * @param {string} s2Input - 比較対象の文字列2
   * @returns {number} 二つの文字列間の最小編集回数
   * @private
   */
  static _getLevenshteinDistance(s1Input, s2Input) {
    let s1 = String(s1Input);
    let s2 = String(s2Input);

    if (s1 === s2) return 0;
    if (s1.length === 0) return s2.length;
    if (s2.length === 0) return s1.length;
    // s2を短い方の文字列にすることでメモリ消費を抑える
    if (s1.length < s2.length) [s1, s2] = [s2, s1];
    const v0 = new Array(s2.length + 1);
    // 初期のコストを設定
    for (let i = 0; i <= s2.length; i++) v0[i] = i;

    for (let i = 0; i < s1.length; i++) {
      let prevDistance = i + 1;
      for (let j = 0; j < s2.length; j++) {
        // 置換コストの計算（文字が同じなら0、違うなら1）
        const cost = s1[i] === s2[j] ? 0 : 1;
        const currentDistance = Math.min(
          v0[j + 1] + 1, // 削除
          prevDistance + 1, // 挿入
          v0[j] + cost, // 置換
        );
        v0[j] = prevDistance;
        prevDistance = currentDistance;
      }
      v0[s2.length] = prevDistance;
    }

    return v0[s2.length];
  }
}
