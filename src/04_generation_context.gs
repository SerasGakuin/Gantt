/**
 * ChatGPT生成用のコンテキストをスプレッドシートから取得するプロバイダ
 * - 取得方法はクラス内にカプセル化
 * - 以前実装（A3:A16 と B1）を模倣
 */
class GenHearingItems_HearingContextProvider {
  constructor(spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) {
    if(!spreadsheet || (typeof spreadsheet.getRange !== 'function')){
      throw new Error("有効なSpreadsheetApp.Spreadsheetを渡してください！")
    } 
    this._ss = spreadsheet;
    this._hearingSheetName = 'ヒアリング';

    // 以前実装に合わせたレンジ
    this._prevItemsRangeA1 = 'A3:A16';
    this._currentMemoCellA1 = 'B1';
  }

  /**
   * @returns {{ prevItemsStr: string, lastTeachingMemo: string }}
   */
  getContext() {
    const sheet = this._getHearingSheet_();

    const prevItemsStr = this._buildPrevItemsStr_(sheet);
    const lastTeachingMemo = this._getLastTeachingMemo_(sheet);

    return { prevItemsStr, lastTeachingMemo };
  }

  // ---------- private ----------

  _getHearingSheet_() {
    const sheet = this._ss.getSheetByName(this._hearingSheetName);
    if (!sheet) throw new Error(`シートが見つかりません: ${this._hearingSheetName}`);
    return sheet;
  }

  /**
   * 以前の preparePreviousWeekPrompt() をほぼそのまま再現
   * A3:A16 を取り、偶数行（0,2,4...）だけを質問として列挙する
   */
  _buildPrevItemsStr_(sheet) {
    const values = sheet.getRange(this._prevItemsRangeA1).getValues(); // 2次元 [[...],[...]]
    let prompt = '';

    // i%2==0 の行だけ採用
    let qNum = 1;
    for (let i = 0; i < values.length; i++) {
      if (i % 2 !== 0) continue;

      const v = values[i][0];
      const s = (v == null) ? '' : String(v).trim();
      if (s === '') continue;

      prompt += `${qNum}. ${s}\n`;
      qNum++;
    }

    return prompt;
  }

  _getLastTeachingMemo_(sheet) {
    const v = sheet.getRange(this._currentMemoCellA1).getValue();
    const s = (v == null) ? '' : String(v).trim();
    return s;
  }
  
}