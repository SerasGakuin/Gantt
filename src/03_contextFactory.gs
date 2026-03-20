//03_contextFactory.gs
/**
 * 計画の生成にあたって必要な情報を詰め込むデータオブジェクト、contextを生成する疑似クラス。
 */
class GenWeeklyPlan_contextFactory {
  /**
   * IDE補完のためのここで返すcontextオブジェクトの型定義
   * @typedef {Object} GenWeeklyPlanContext
   * @property {Object} dependencies
   * @property {string} materialId
   * @property {MaterialInfo} materialInfo
   * @property {string} lastAchievement
   * @property {number} valueToAdd
   * @property {number} numbersCount
   * @property {number} tildasCount
   * @property {boolean} hasValidMetadata
   * @property {boolean} isCountedOnlyByChapter
   * @property {boolean} isCountingRepeating
   */

  /**
   * コンテキストオブジェクトにもろもろの情報を詰め込む関数。
   * contextは、原則として純粋なデータの箱であることが要求される。ただし、dependenciesの中身だけは例外。
   *
   * @param {object} dependencies - 週間計画作成マクロが依存している外部の関数を詰め込んだオブジェクト
   * @param {string} materialId - 参考書のID。参考書マスターのデータに従う。
   * @param {object} materialInfo - 参考書の情報。参考書マスターのapiに従う。
   * @param {string} lastAchievement - この参考書の最新の学習実績。シートの情報にしたがう（2026_01_25時点）
   * @param {number} valueToAdd - 次の週でこの参考書の課題を処理する量。単位は章だったり問題番号だったり、参考書による。
   *
   * @returns {GenWeeklyPlanContext} - 計画作成用の各handlerが必要とする情報をすべて詰め込んだ、参考書単位の情報のまとまったオブジェクト。contextと呼んでいる
   */
  static create(
    dependencies,
    materialId,
    materialInfo,
    lastAchievement,
    valueToAdd,
  ) {
    const context = {
      dependencies,

      materialId,
      materialInfo,
      lastAchievement,
      valueToAdd,

      numbersCount: (lastAchievement.match(/\d+/g) || []).length,
      tildasCount: (lastAchievement.match(/~/g) || []).length,

      hasValidMetadata: this._hasValidMetadata(materialInfo),
      isCountedOnlyByChapter: this._isCountedOnlyByChapter(materialInfo),
      isCountingRepeating: this._isCountingRepeating(materialInfo),
    };
    return Object.freeze(context);
  }

  /**
   * 参考書マスターに有効な情報のある参考書であるか判定。
   * ここでの「有効」とは、週間計画作成につかえる情報があるかどうかという意味。
   */
  static _hasValidMetadata(materialInfo) {
    if (!materialInfo) return false; //情報自体がundefined

    /*
      データ自体はあるが問題番号に関して何も情報がないケースが案外ある。これはデータなしと判定したいので以下のフラグで判定。
    */
    const starts = materialInfo.chapterStarts; //各章のはじめの問題番号
    const ends = materialInfo.chapterEnds; //各章の終わりの問題番号
    const hasStartsNums = starts?.length != 0 && !!starts[0]; //配列が存在して、かつ番号が有効
    const hasEndsNums = ends?.length != 0 && !!ends[0];

    const hasValidChapter = hasStartsNums && hasEndsNums;
    return hasValidChapter;
  }

  /**
   * 章立てのみで進み、、問題番号などの章より下位の区分が不要な参考書か判定。
   */
  static _isCountedOnlyByChapter(materialInfo) {
    if (!materialInfo) return false; //情報自体がundefined

    const chapterStartsNums = materialInfo.chapterStarts || []; //各章のはじめの問題番号
    const chapterEndsNums = materialInfo.chapterEnds || []; //各章の終わりの問題番号

    //章の数
    const chapterCount = Math.min(
      chapterStartsNums.length,
      chapterEndsNums.length,
    );
    if (chapterCount < 2) return false; //章が一つしかないならfalse

    //各章について、「すべての章の問題数が1」または「すべての章について、問題番号の情報がない」ならtrue。途中で例外を発見したらfalse
    for (let i = 0; i < chapterCount; i++) {
      const s = chapterStartsNums[i];
      const e = chapterEndsNums[i];

      const isLackOfStartOrEnd =
        s === undefined || s === "" || e === undefined || e === "";
      const isStartAndEndSame = s == e;

      const canBeChaperOnlyMaterial = isLackOfStartOrEnd || isStartAndEndSame; //NOTE: まだ確認していないが、これだと1問しかない章と問題番号が欠けている章が交互にくるタイプの参考書では誤検出する

      if (!canBeChaperOnlyMaterial) return false;
    }
    return true;
  }

  /**
   * 章ごとに問題番号がループするタイプであるかどうかを判定する。
   */
  static _isCountingRepeating(materialInfo) {
    if (!materialInfo) return false; //情報自体がundefined

    const chapterStartsNums = materialInfo.chapterStarts || []; //各章のはじめの問題番号
    if (chapterStartsNums.length <= 1) return false;

    const firstChapProbNum = chapterStartsNums[0];
    for (let i = 1; i < chapterStartsNums.length; i++) {
      if (chapterStartsNums[i] != firstChapProbNum) return false; //一個でも一致しないなら非該当
    }
    return true;
  }
}
