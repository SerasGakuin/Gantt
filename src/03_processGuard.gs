//03_processGuard.gs
/**
 * 週間計画作成マクロを動作させて良いか確認する動作の責務を負うクラス。
 */
class GenWeeklyPlan_processGuard {

  /**
   * この関数を実行して、返り値がtrueなら実行してもよい。
   * 
   * @param {(string|null|undefined)[][]} existingAchievements - 既存の学習実績の二次元配列。5列になっており、列index + 1 = 対応する週の番号 (第n週)になっている必要がある。
   * @param {number} weekIndex - 今から計画を生成しようとしている週の番号
   * 
   * @returns {boolean} 実行してよいならtrue、実行してはいけないのならfalse
   */
  static isRunAllowed(existingAchievements, weekIndex) {
    const planBeforeRun = existingAchievements.map(row => row[weekIndex]);//今埋めようとしている週の既存データ。
    const isAchievementsEmpty = this._checkEmptyOrNot(planBeforeRun);//今から計画を生成しようとしている週にすでにデータがあるか？
    if (!isAchievementsEmpty && !this._askRunOrNot()) return false;//データがあって、かつ実行許可をもとめても拒否されたら終了
    return true;
  }

  /**
   * 渡された配列が空であるかどうかチェックするメソッド。
   * 
   * @param {(string|null|undefined)[]} array - 対象の配列
   * 
   * @returns {boolean} 空ならtrueそうでないならfalse
   * 
   * @private
   */
  static _checkEmptyOrNot(array) {//既存の実績が空であるかチェックする
    let isPlanExisting = false;//一個でも空でない計画が入っていたらtrueになる
    for (let i = 0; i < array.length; i++) {
      const plan = array[i];
      if (!(!plan || plan == '' || plan == 0)) {
        isPlanExisting = true;
        break;
      }
    }
    return !isPlanExisting;
  }

  /**
   * ユーザーに処理を開始してよいか尋ねるメソッド。
   * 
   * @returns {boolean} 処理してよいといわれた場合はtrue
   * 
   * @private
   */
  static _askRunOrNot() {// 既存データがある場合の確認処理
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Continue?',
      'この週にはすでに入力された文字列があります。続けますか?',
      ui.ButtonSet.YES_NO
    );
    return response === ui.Button.YES;
  }
}
