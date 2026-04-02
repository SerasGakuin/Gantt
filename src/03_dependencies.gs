//03_dependencies.gs
/**
 * 本番用の依存関係を注入されたdependenciesオブジェクトを取得する関数。
 * ライブラリの新実装のテスト時には、一時的な関数を適当な箇所に定義して、
 * そこから直にgetGenWeeklyPlanDependenciesを使って新実装のライブラリの関数を注入して実行するようにしてください。
 * @param {string} spId - スピードプランナーのid。指定しなければいま開いてるブックに
 */
function getGenWeeklyPlanProdDependencies(spId = null) {
  const ss = spId
    ? SpreadsheetApp.openById(spId)
    : SpreadsheetApp.getActiveSpreadsheet();

  return new GenWeeklyPlanDependencies(
    () => SheetIO.getSpeedPlannerIOManager(ss), // ← ssを閉じ込める
    (arr, logFlag) => MaterialsMasterLib.getMaterialsInfoByIds(arr, logFlag),
    () => PropertiesService.getScriptProperties().getProperty("APIKEY"),
    () => {
      return "gpt-4.1";
    },
  );
}

/**
 * 適当なライブラリがきちんと使用できるかを試すための関数。中身は適当にいじってよいです。
 */
function genWeeklyPlan_test_dependencies() {
  const dependencies = getGenWeeklyPlanProdDependencies();
  const obj = dependencies.getMaterialsInfoByIds(["gEK014"], true);
  console.log(obj);
}

/**
 * genWeeklyPlan用の依存性の格納をするクラス。必要なライブラリのapi関数が詰まっている。
 *
 * @typedef GenWeeklyPlanDependencies
 * @property {function} getSpeedPlannerIOManager
 * @property {function} getMaterialsInfoByIds
 * @property {function} getChatGPTAPIKEY
 * @property {function} getChatGPTModelName
 */
class GenWeeklyPlanDependencies {
  constructor(
    getSpeedPlannerIOManager,
    getMaterialsInfoByIds,
    getChatGPTAPIKEY,
    getChatGPTModelName,
  ) {
    const isLackOfElements =
      !getSpeedPlannerIOManager ||
      !getMaterialsInfoByIds ||
      !getChatGPTAPIKEY ||
      !getChatGPTModelName;
    if (isLackOfElements) throw new Error(`依存の注入が不足しています。`);
    this.getSpeedPlannerIOManager = getSpeedPlannerIOManager;
    this.getMaterialsInfoByIds = getMaterialsInfoByIds;
    this.getChatGPTAPIKEY = getChatGPTAPIKEY;
    this.getChatGPTModelName = getChatGPTModelName;
    Object.freeze(this); //書き換え防止
  }
}

/**
 * 週間計画作成マクロ用のSpeedPlannerIOManagerクラスのヘルパー
 * NOTE: まだ使ってない。週間計画作成マクロで使用するものは、こっちに移行予定
 */
class GenWeeklyPlan_spIOManagerHelper {
  /**
   * NOTE: まだ使ってない。週間計画作成マクロで使用するものは、こっちに移行予定
   * @param {SpreadSheet} thisBook - 操作対象ブック
   * @param {number} weekNum - 操作したい週の番号
   */
  constructor(thisBook, weekNum) {
    this._spIOManager = SheetIO.getSpeedPlannerIOManager(thisBook);
    this._weekNum = weekNum; //週の番号。1から始める。
    this._weekIndexInArr = weekNum - 1; //週の配列内index。0から始まる。
    this._allAchievements = this._spIOManager.getActiveMaterialsAchievements();
    this._hasGottenAchievements = false;
    Object.freeze(this);
  }

  /**
   * NOTE: まだ使ってない。週間計画作成マクロで使用するものは、こっちに移行予定
   * 週の番号でアクセスして、その週の既存の実績をフラットな配列で取得する関数。
   * @param {number} weekNum - データが欲しい週の番号
   * @public
   */
  getWeekAchievementsAsFlatArray() {
    if (!!this._hasGottenAchievements)
      throw new Error(
        "「週間計画作成」で既存の実績の取得は一回限りにしてください。データの取得は重い操作であり、かつ思わぬ操作の副作用のリスクが大きい行為です。",
      );
    this._hasGottenAchievements = true;

    const achievements = this._allAchievements.map(
      (row) => row[this._weekIndexInArr],
    );
    return achievements;
  }

  /**
   * NOTE: まだ使ってない。週間計画作成マクロで使用するものは、こっちに移行予定
   * 渡された変数が1次元配列であるかどうか確かめる関数。
   * @param {Array} flatArray - フラットな配列であるかどうか調査される対象。
   * @private
   */
  _isFlatArray(flatArray) {
    const isFlatArray =
      Array.isArray(flatArray) && flatArray.every((v) => !Array.isArray(v));
    return isFlatArray;
  }

  /**
   * NOTE: まだ使ってない。週間計画作成マクロで使用するものは、こっちに移行予定
   * 週の番号でアクセスして、フラットな配列ですべての週の既存の実績をセットする関数。
   * @param {number} weekNum - データが欲しい週の番号
   * @public
   */
  setWeekAchievementsWithFlatArray(achievementsArray) {
    if (this._hasSetAchievement)
      throw new Error(
        "「週間計画作成」で既存の実績のセットは一回限りにしてください。データのセットは重い操作であり、かつ思わぬ操作の副作用のリスクが大きい行為です。",
      );
    this._hasSetAchievements = true;

    const isGivenAchievementsValid = this._isFlatArray(achievementsArray);
    if (!isGivenAchievementsValid)
      throw new Error(
        `このメソッドにわたす変数はフラットな配列にしてください。実績の新旧のindexの対応関係は保ってください。\nわたされた変数: ${JSON.stringify(achievementsArray)}`,
      );
  }
}
