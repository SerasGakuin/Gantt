/**
 * 各種シートのフォーマットを修復する関数。
 * 詳しい実装はシート構成参照用スピードプランナーに紐づいているSheetIOライブラリの実装を参照してください。
 */
function repairFormats() {
  const thisBook = SpreadsheetApp.getActive();
  const spIOManager = SheetIO.getSpeedPlannerIOManager(thisBook);
  spIOManager.repairFormats(); //数式修復
}
