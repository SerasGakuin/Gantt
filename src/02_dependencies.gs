//02_dependencies.gs
//TODO: spIOManagerを直接注入するよりも、間に入るViewModelを注入する方がメンテナンスが誰にでもやりやすくなると思います
/**
 * 本番用の依存関係を注入されたdependenciesオブジェクトを取得する関数。
 * ライブラリの新実装のテスト時には、test_geshoなどの一時的な関数を適当な箇所に定義して、
 * そこから直にgetGeshoDependenciesを使って新実装のライブラリの関数を注入して実行するようにしてください。
 * 
 * @returns {GeshoDependencies} - 月末月初の依存外部関数
 */
function getGeshoProdDependencies() {
  return new GeshoDependencies(
    () => SheetIO.getSpeedPlannerIOManager(SpreadsheetApp.getActiveSpreadsheet()),
    (ids, cols) => MaterialsMasterLib.getMaterialsMaster(ids, cols),
    (arr, logFlag) => MaterialsMasterLib.getMaterialsInfoByIds(arr, logFlag),
    (nameStr) => MaterialsMasterLib.guessSubjectsByName(nameStr),
    (spIOManager) => UpdateOnGesho.start(spIOManager),
    (yearAfterGeshoNum, monthAfterGeshoNum) => getMonthlyData(yearAfterGeshoNum, monthAfterGeshoNum)
  )
}

/**
 * 適当なライブラリがきちんと使用できるかを試すための関数。
 * @private
 */
function gesho_test_dependencies() {
  const dependencies = getGeshoProdDependencies();
  const obj = dependencies.getMaterialsInfoByIds(['gEK014'], true)
  console.log(obj);
}

/**
 * gesho用の依存先を持つクラス。
 * 
 * @param {function} getSpeedPlannerIOManager - SpeedPlannerIOMangerを取得する関数。
 * 
 * @param {function} getMaterialsMaster - 参考書マスターのnon-staticバージョンの関数をつかってMaterialsMasterインスタンスを取得する関数。 TODO: これはSheetIOとgeshoの間を中継するようなViewModelを注入するようにしたほうが独立性が上がる
 * 
 * @param {function} getMaterialsInfoByIds - 参考書マスターのstaticバージョンの関数をつかって参考書データを検索する関数。
 * 
 * @param {function} guessSubjectsByName - 参考書マスターの機能で参考書名をわたすと、そこから科目を推測する機能をもつ関数。
 * 
 * @param {function} updateOnGesho -  月末月初につけておこないたい軽微な更新を実行する関数。
 * 
 * @param {function} getMonthlyData - 01_minor_projects.gsの、ガントチャートを読む関数 TODO:これはSheetIOに統合したい。
 */
class GeshoDependencies {
  constructor(
    getSpeedPlannerIOManager,
    getMaterialsMaster,
    getMaterialsInfoByIds,
    guessSubjectsByName,
    updateOnGesho,
    getMonthlyData
  ) {
    const isLackOfElements = !getSpeedPlannerIOManager || !getMaterialsInfoByIds || !getMaterialsMaster || !guessSubjectsByName || !updateOnGesho || !getMonthlyData;
    if (isLackOfElements) throw new Error(`依存ライブラリの注入が不足しています。`);
    this.getSpeedPlannerIOManager = getSpeedPlannerIOManager;
    this.getMaterialsMaster = getMaterialsMaster;
    this.getMaterialsInfoByIds = getMaterialsInfoByIds;
    this.guessSubjectsByName = guessSubjectsByName;
    this.updateOnGesho = updateOnGesho;
    this.getMonthlyData = getMonthlyData;
    console.log('注入された依存ライブラリは:\n', this);
    Object.freeze(this);//書き換え防止
  }
}

