/**
 * # ヒアリング項目作成マクロ
 * 
 * 毎週の面談のために作成される「ヒアリング項目」を自動作成するマクロです。
 * AIを利用して項目を生成します。
 * 
 * @param {string} spId - ヒアリング項目を作成したいスピードプランナーのid(将来的にdbになった場合はこのidをレコードカラムにすれば一貫性は保たれる。)
 */
function genHearingItems(spId = null) {
  if(spId !== null && typeof spId !== 'string'){
    throw new Error("スピードプランナーのidを指定する場合は、文字列型で指定してください。");
  }
  const currSpSheet = !!spId? SpreadsheetApp.openById(spId): SpreadsheetApp.getActiveSpreadsheet();
  // 前回までの実行履歴などのコンテキストを取得するクラス
  const contextProvider = new GenHearingItems_HearingContextProvider(currSpSheet);
  GenHearingItems.run(contextProvider, currSpSheet);
}

// TODO: 具体的なセッターのSpeedPlannerIOManagerへの移行をしたい。
// シート構成を各マクロが知っていなくても良いようにリファクタしたい。そうでないとシート構成の変更時の影響範囲が予測不能になる。


class GenHearingItems {

  static run(contextProvider, currSpSheet) {
    if(typeof contextProvider.getContext !== 'function'){
      throw new Error("[内部バグ]不正なプロバイダオブジェクトです。");
    }
    try {
      // シートをクリアして生成中表示　TODO: 具体的なセッターのSpeedPlannerIOManagerへの移行
      this._clearPreviousResponses(currSpSheet);
      // AIに項目を配列で生成してもらう
      const aiResponcesArr = GenHearingItems_ItemsGeneratorWithChatGPT.generate(contextProvider.getContext());
      // 項目をセット　TODO: 具体的なセッターのSpeedPlannerIOManagerへの移行
      this._setAIOutputArrToSheet(currSpSheet, aiResponcesArr);
    } catch (e) {
      this._setAIOutputArrToSheet(currSpSheet, [`[${e.name || 'Error'}] ${e.message || e}`]);
      console.error(e);
      throw e;
    }
  }

  /**
   * ヒアリングシート取得
   * TODO: 具体的なシート取得のSpeedPlannerIOManagerへの移行
   * @deprecated - SpeedPlannerIOManagerへの移行
   */
  static _getSheet(currSpSheet) {
    const sheet = currSpSheet.getSheetByName('ヒアリング');
    if (!sheet) throw new Error("ヒアリングシートが見つかりません");
    return sheet;
  }

  /**
   * TODO: 具体的なセッターのSpeedPlannerIOManagerへの移行
   * @deprecated - SpeedPlannerIOManagerへの移行
   */
  static _clearPreviousResponses(currSpSheet) {
    const sheet = this._getSheet(currSpSheet);
    sheet.getRange("B3:B16").clearContent();
    sheet.getRange("B3").setValue("generating...");
  }

  /**
   * TODO: 具体的なセッターのSpeedPlannerIOManagerへの移行
   * @deprecated - SpeedPlannerIOManagerへの移行
   */
  static _setAIOutputArrToSheet(currSpSheet, contentArr) {
    const setTargetRange = this._getSheet(currSpSheet).getRange("B3:B16");

    //　配列そのままでは詰まっているので一個ごとに改行を入れる
    const setSrcMat = [];
    for (let i = 0; i < contentArr.length; i++) {
      setSrcMat.push([contentArr[i]]);
      setSrcMat.push(['']);
    }
    // はみ出した分を削除・足りない数を補完
    const neededRowCount = setTargetRange.getNumRows();//この数を超える分は不要
    const finalMat = setSrcMat.slice(0, neededRowCount);
    while (finalMat.length < neededRowCount) finalMat.push(['']);//足りないならpush
    // セット
    setTargetRange.setValues(finalMat);
  }


}








