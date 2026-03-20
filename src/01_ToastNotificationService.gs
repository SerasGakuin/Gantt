//01_ToastNotificationService.gs
/**
 * Toast通知の送信を担当するクラス。
 * Toast通知とは、画面に比較的小さく表示される、一定時間で自動的に表示がきえる通知方法のこと。軽微な内容の通知にどうぞ
 */
class ToastNotificationService {

  /** @returns {string} @private */
  static get LOG_TAG() { return 'ToastNotificationService:' }
  /** @returns {number} @private */
  static get DEFAULT_TIME() { return 16 }
  /** @returns {SpreadSheetApp.SpreadSheet.Sheet | null} @private */
  static get targetSS() {
    if (!this._targetSS) this._targetSS = SpreadsheetApp.getActiveSpreadsheet();
    if (!this._targetSS) console.error(`${this.LOG_TAG}\nActiveSpreadsheetが取得できません。送信しようとしたメッセージ内容:\n${message}`);
    return this._targetSS;
  }

  /**
   * Toast通知を送信する。
   * @param {string} message - 送信したい通知内容
   * @param {number} [time = 16] - toastの表示時間(秒単位。ミリ秒ではない)
   * @public
   */
  static send(message, time = 16) {
    try {
      //対象ブック
      const targetSS = this.targetSS;

      //初期化できなかったらあきらめる
      if (!targetSS) return;

      //入力の正規化
      const sanitizedMsg = !message ? '' : String(message);
      const sanitizedTime = this._sanitizeTimeInput(time);

      //toastとLoggerで両方で送信する。
      targetSS.toast(sanitizedMsg, "通知", sanitizedTime);
      Logger.log(`${this.LOG_TAG}\nToast通知として以下の内容を${sanitizedTime}秒間表示します:\n${sanitizedMsg}`);
    } catch (e) {
      //ただの通知用ユーティリティのためにthrowして動作を止めることがないようにconsoleに記録して終わりにする
      console.error(`${this.LOG_TAG}\ntoastの送信に失敗しました！`, e);
    }
  }

  /**
   * Toastの表示時間の入力を正規化する関数。
   * 
   * @param {any} baseTime - 生の入力
   * @returns {number} - 正規化された値
   * @private
   */
  static _sanitizeTimeInput(baseTime) {
    const timeNum = Math.ceil(Number(baseTime));

    if (!Number.isFinite(timeNum)) {//数字でない入力は禁止なのでデフォルト値を返す
      console.warn(`${this.LOG_TAG}\n渡されたtime:${baseTime}秒はToast表示時間として不正です。Toastの表示時間をデフォルト値${this.DEFAULT_TIME}秒に設定します。`);
      return this.DEFAULT_TIME;

    } else if (timeNum <= 0) {//負の値は禁止なのでデフォルト値を返す
      console.warn(`${this.LOG_TAG}\n渡されたtime:${baseTime}秒はToast表示時間として不正です。Toastの表示時間をデフォルト値${this.DEFAULT_TIME}秒に設定します。`);
      return this.DEFAULT_TIME;

    } else if (timeNum > 600) {//異常に長い場合は警告してそのまま返す
      console.warn(`${this.LOG_TAG}\nToastの表示時間が${baseTime}秒に指定されました。異常に長い値です。間違いでないか確認してください。`);
      return timeNum;

    } else {//特に問題なかった場合はそのまま返す
      return timeNum;
    }
  }
}

