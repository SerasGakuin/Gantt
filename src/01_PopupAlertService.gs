/**
 * 画面中央に、Toastとは違って永続する通知を出すクラス。
 * 2026/02/05現在は
 * SpreadsheetApp.getUi().alert(messege);
 * を利用しているが、将来的に別の方法に切り替えることができるようにラッパー化している。
 */
class PopupAlertService {
  /**
   * ui.alertでメッセージを表示する関数。ユーザーがOKを押すまで後続の処理は行われない。
   * 返り値で、アラートを表示するために消費した時間を返却する。実行時間制限対策で後続処理をキャンセルするかどうかの判断にご利用ください。
   * @param {string} message
   * @returns {number} アラートを表示するために消費した時間(ミリ秒単位)
   */
  static send(message) {
    const startTime = new Date().getTime();
    SpreadsheetApp.getUi().alert(message);
    const endTime = new Date().getTime();
    return endTime - startTime;
  }
}
