// 00_SidebarMenuService.gs
/**
 * ライブラリの呼び出し元でサイドバーに00_sidebar_interface.htmlのマクロメニューを表示するための関数。
 */
function showSidebarMenu() {
  SidebarMenuService.show();
}

/**
 * アクティブなスプレッドシートに、「ガントチャートテンプレート」ライブラリのhtmlで定義されているサイドバーメニューを表示するクラス。
 * 00_FunctionConfigs.gs内での関数についての設定の定義を参照する。
 * そして、contextオブジェクトを引数として必要としない、引数のない関数の実行用ボタンを自動でサイドバーに追加し、追加した状態で表示する。
 * 詳しい、GAS側からFunctionConfigsを受け取ってボタンを追加するアルゴリズムの実装はhtml内の特定のタグに定義されている。
 */
class SidebarMenuService {
  /**
   * サイドバーを表示するためのメソッド。
   * @throws {Error} 何かエラーがあれば、SidebarMenuService:というタグをつけつつ、再度throwする。
   * @public
   */
  static show() {
    try {
      const template = HtmlService.createTemplateFromFile(
        "00_sidebar_interface",
      ); //00_sidebar_interface.html
      template.FunctionConfigs = FunctionConfigs;

      const html = template
        .evaluate()
        .setTitle("スピードプランナー用マクロ選択");
      SpreadsheetApp.getUi().showSidebar(html);
    } catch (e) {
      throw new Error(
        `SidebarMenuService:\nサイドバーの表示をしようとしてエラーが発生しました。\n${e}`,
      );
    }
  }
}
