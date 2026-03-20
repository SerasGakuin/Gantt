//03[GenWeeklyPlan]_main.gs
/**
 * 「週間計画作成マクロ」の実行開始用の関数や、処理パイプラインの実装をするファイルです。
 */

/**
 * 呼び出し用関数。
 */
function week1Plan() {
  _GenWeeklyPlanProdFunc(1);
}
function week2Plan() {
  _GenWeeklyPlanProdFunc(2);
}
function week3Plan() {
  _GenWeeklyPlanProdFunc(3);
}
function week4Plan() {
  _GenWeeklyPlanProdFunc(4);
}
function week5Plan() {
  _GenWeeklyPlanProdFunc(5);
}

/**
 * 本番実行用の関数。
 * @param {number} weekNum - 計画を生成したい週の番号
 * @private
 */
function _GenWeeklyPlanProdFunc(weekNum) {
  const dependencies = getGenWeeklyPlanProdDependencies(); //本番用の依存関係
  GenWeeklyPlan.start(weekNum, dependencies);
}

/**
 * 週間計画作成マクロの処理のパイプラインをつかさどるクラス。
 * TODO: データ収集を別のクラスに分離したい。
 */
class GenWeeklyPlan {
  /**
   * 生成開始用関数。
   *
   * @param {number} weekNum - 週の番号
   * @param {GenWeeklyPlanDependencies} dependencies - 依存
   *
   * @public
   */
  static start(weekNum, dependencies) {
    const spIOManager = dependencies.getSpeedPlannerIOManager();
    const existingAchievements = spIOManager.getActiveMaterialsAchievements(); //既存の全週の計画

    const weekIndex = weekNum - 1; //週の番号を対応する配列インデックスに変換

    //実行チェック。データがある場合は上書きしてよいか確認して、拒否されたら終了。
    const isRunAllowed = GenWeeklyPlan_processGuard.isRunAllowed(
      existingAchievements,
      weekIndex,
    );
    if (!isRunAllowed) return;

    //現在の教材の関連データをcontextオブジェクトに詰め込んでcontextたちの配列として取得。
    const materialsContextsArr = this._collectInfoAsContexts(
      spIOManager,
      dependencies,
      weekNum,
    );

    // 教材ごとに処理する。
    for (let i = 0; i < materialsContextsArr.length; i++) {
      const context = materialsContextsArr[i];
      if (!context.materialInfo) {
        //デバッグ用に参考書マスターに情報がない場合は
        console.warn(
          `参考書マスターに情報なし。\nID：${String(context.materialId)}\n最新の学習実績：${String(context.lastAchievement)}`,
        );
      }
      const nextPlan = GenWeeklyPlan_handlers_generator.generate(context);
      existingAchievements[i][weekIndex] = nextPlan; //指定週の計画データ行列の対応する位置にセット
    }

    //TODO: 現状、すべての週の実績を一気に取得したり、一気に張り付ける方法しかないのでsheetIOManagerの最適化が望まれる。
    // 処理完了。シートに出力
    spIOManager.setActiveMaterialsAchievements(existingAchievements); //シートに新しい実績を送信して終了
  }

  /**
   * 必要なデータを、教材ごとにcontextオブジェクトに詰めて、そのオブジェクトを配列に格納して返却する関数。
   *
   * @param {SpeedPlannerIOManager} spIOManager - シートとの通信役
   * @param {GenWeeklyPlanDependencies} dependencies
   * @param {number} weekNum
   *
   * @returns {GenWeeklyPlanContext[]}
   *
   * @private
   */
  static _collectInfoAsContexts(spIOManager, dependencies, weekNum) {
    //参考書のidを取得してフラット配列化。
    const materialsIds = spIOManager.getActiveMaterialsIds().map((v) => v[0]);
    //参考書の諸データをまとめて配列として取得。
    const materialsData = dependencies.getMaterialsInfoByIds(materialsIds);

    /**
     * 最新の学習実績を正規化する補助関数。
     * @param {string} value - 正規化前の学習実績
     * @returns {string} 正規化された学習実績
     */
    function sanitizeLastAchievement(value) {
      if (!value) return "";
      let output = String(value).trim();
      output = output.replace(/[０-９]/g, function (s) {
        //全角の数字を半角に
        return String.fromCharCode(s.charCodeAt(0) - 0xfee0);
      });
      output = output
        .replace(/\s*~\s*/g, "~") //チルダの前後の空白を削除
        .replace(/[～〜∼]/g, "~") //半角チルダの偽物を半角チルダに変換
        .replace(/　/g, " ")
        .replace(/；/g, ";")
        .replace(/：/g, ":");
      return output;
    }

    //先週のデータを取得して、正規化したうえでフラット配列化。
    const lastAchievements = spIOManager
      .getActiveMaterialsAchievementsUpTo(weekNum)
      .map((v) => sanitizeLastAchievement(v[0]));

    //「目安処理量」を取得して、整数に丸めたうえでフラット配列化。
    const valuesToAddArr = spIOManager
      .getActiveMaterialsProcessAmount(weekNum)
      .map((v) => {
        const num = Number(v[0]);
        return Number.isFinite(num) ? Math.max(0, Math.round(num)) : 0;
      });

    // 教材ごとにコンテクストに格納する。
    let contextsArr = [];
    for (let i = 0; i < materialsIds.length; i++) {
      const context = GenWeeklyPlan_contextFactory.create(
        dependencies,
        materialsIds[i],
        materialsData[materialsIds[i]],
        lastAchievements[i],
        valuesToAddArr[i],
      );
      contextsArr.push(context);
    }
    return Object.freeze(contextsArr);
  }
}
