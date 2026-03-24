//01_GASRefSheetLogService.gs
/**
 * @class
 * # GASRefferenceSheetLogService
 *
 * GAS参照用ブックのログ記録用シートに見逃したくないログを記録するクラス。
 * appendRowを使用するため、何度も通知を送信する場合には重いです。特に重要なログを永続させるために使用してください。
 * メッセージの内容さえ送れば、時刻は自動で記録される。
 *
 * ### 提供メソッド一覧
 *
 * - log(message) : 一般ログ。
 * - info(message) : 特に重要な実行情報。
 * - user(message) : ユーザーが送信したエラー報告など、システムによる自動入力ではないもの。
 * - warn(message) : 警告文。
 * - error(message) : エラーログ。
 *
 */
class GASRefferenceSheetLogService {
  /** @private */
  static get CONFIG() {
    return {
      BOOK_ID: "10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ",
      SHEET_NAME: "logSheet",
    };
  }

  /**
   * GAS参照用ブックのエラー記録用シートに任意のログを記録する。
   * @param {string} message - 記録したいログの内容。
   * ログレベルを明示的に追加することをお勧めします
   * @deprecated
   */
  static record(message) {
    GASRefferenceSheetLogService._recordWithLevel(message, "LOG");
  }

  /**
   * GAS参照用ブックのエラー記録用シートに任意のログを記録する。
   * ログレベル:'LOG' - 一般ログ。
   * @param {string} message - 記録したいログの内容。
   * @public
   */
  static log(message) {
    console.log(message);
    GASRefferenceSheetLogService._recordWithLevel(message, "LOG");
  }

  /**
   * GAS参照用ブックのエラー記録用シートに任意のログを記録する。
   * ログレベル:'USER' - ユーザーが送信したエラー報告など、システムによる自動入力ではないもの。
   * @param {string} message - 記録したいログの内容。
   * @public
   */
  static user(message) {
    console.info("[user input] "+message);
    GASRefferenceSheetLogService._recordWithLevel(message, "USER");
  }

  /**
   * GAS参照用ブックのエラー記録用シートに任意のログを記録する。
   * ログレベル:'INFO' - 特に重要な実行情報。
   * @param {string} message - 記録したいログの内容。
   * @public
   */
  static info(message) {
    console.info(message);
    GASRefferenceSheetLogService._recordWithLevel(message, "INFO");
  }

  /**
   * GAS参照用ブックのエラー記録用シートに任意のログを記録する。
   * ログレベル:'WARN' - 警告文
   * @param {string} message - 記録したいログの内容。
   * @public
   */
  static warn(message) {
    console.warn(message);
    GASRefferenceSheetLogService._recordWithLevel(message, "WARN");
  }

  /**
   * GAS参照用ブックのエラー記録用シートに任意のログを記録する。
   * ログレベル:'ERROR' - エラーログ
   * @param {string} message - 記録したいログの内容。
   * @public
   */
  static error(message) {
    console.error(message);
    GASRefferenceSheetLogService._recordWithLevel(message, "ERROR");
  }

  /**
   * GAS参照用ブックのエラー記録用シートに任意のログを記録する。
   * @param {string} message - メッセージ
   * @param {string} level - メッセージのレベルを指定する
   * @private
   */
  static _recordWithLevel(message, level) {
    try {
      const isTarget =
        typeof message === "string" && message.trim().length !== 0;
      if (!isTarget) return; //logとして残せるものでないならreturn

      const logSheet = this.logSheet;
      if (!logSheet) return;

      const timestamp = new Date();
      const msgSafe = message.trim();

      // メッセージ行を追加
      logSheet.appendRow([timestamp, level, msgSafe]);

      // 追記した行の高さ調整
      const lastRow = logSheet.getLastRow();
      if (lastRow > 1) logSheet.setRowHeight(lastRow, 48);
    } catch (e) {
      //ただのログ保存用ユーティリティのためにthrowして動作を止めないようにconsoleに記録して終わりにする
      console.error(
        `GASRefferenceSheetLogService:\nログの記録に失敗しました！\n送信しようとしていた内容:\n${message}`,
        e,
      );
    }
  }

  /**
   * logを出力するためのシートを取得するメソッド。
   *
   * @returns {SpreadSheetApp.SpreadSheet.Sheet} - logSheet
   * @private
   */
  static get logSheet() {
    if (!this._logSheet) {
      try {
        const gasRefBook = SpreadsheetApp.openById(this.CONFIG.BOOK_ID);
        this._logSheet = gasRefBook.getSheetByName(this.CONFIG.SHEET_NAME);
      } catch (e) {
        console.error("GASRefferenceSheetLogService: シート取得エラー", e);
      }
    }
    return this._logSheet;
  }
}
