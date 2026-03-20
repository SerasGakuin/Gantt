//00[Document].gs
/**
 * # interface
 * 
 * 個別の生徒のスピードプランナーと、この「ガントチャートテンプレート」ライブラリとの接続部分のふるまいを定義します。
 * 
 * 
 * ### 関数の呼び出し追加手段
 * 
 * 1. サイドバーから実行できる「引数を受け取らない」関数を追加する手順(なるべくこちらで対応していただけると、実装が楽です)
 * 
 * - まず、グローバルに追加したい関数を定義する。(「function funcX(){...」という風に。この場合この関数の名前をfuncXとする。)
 * 
 * - 00_FunctionsConfigs.gsの、FunctionsConfigsという配列に、同ファイルにあるコメントの説明通りにfuncXの設定を追加すればokです。
 * この設定に準拠して自動でサイドバーのボタンが追加され、GAS側も対応されます。
 * 
 * 
 * 2. サイドバーから実行できる「引数を受け取る」関数を追加する手順(どうしても引数が必要な場合)
 * 
 * - 00_contextFunctions.gsに、contextを受け取る関数をグローバルに定義。この関数の名前を仮にfuncXとする。例は00_contextFunctions.gsにあります。
 * 
 * - 最初に、00_FunctionsConfigs.gsの、ContextFunctionsConfigsという配列に、同ファイルにあるコメントの説明通りにfuncXを追加
 * 
 * - 次に、htmlに手動でボタンを<div id="selection">ブロック内に、対応する関数を<script>タグ内に追加する。
 * GAS側のabstractFunctionに、idを最初にContextFunctionに追加したidとして、{id, context}という構造のオブジェクトを渡すように設定する。
 * 例:
 *  const idObj = {id: 'funcX', context: {message: message, time: new Date()}};
 *  google.script.run.abstractFunction(idObj);
 * 既存の引数あり関数も参考になると思います。
 * 
 * 
 * 3. スプレッドシートの上のバーから実行できる関数を追加する方法
 * 
 * 00[interface]_main.gs の onOpenAction() 内で生成するuiの内容を変えます。
 * 
 * 
 * ※サイドバーを表示する機構はサイドバー内のボタン以外に用意する必要があります。
 * onOpenAction() 内で生成するuiにサイドバー表示用のボタンを登録しておくのがよいと思います。
 * 
 * ※JSではグローバル関数もオブジェクト扱いです。
 * 
 * ※functionIdを関数名とする仕組みに、現状なっていますが、安全のためにはidを明示的に設定したほうが良いかもしれません。その場合は00_FunctionsRunnerService.gsのidからfunctionへのマップ生成クラスと、htmlでのidの取得方法を両方ともconfig内のfunctionのnameを得るやり方から、idを得るやり方にして、00_FunctionsConfigs.gsのコンフィグすべてにidプロパティを追加する必要があります。
 * 
 * 
 * ### 仕組み
 * 
 * 個別のスピードプランナーのスクリプトの内容は以下になっています。
 * 
 * =================================================================================
 * //Ganttライブラリに処理の中身の全てを投げる。abstactFunctionは、任意の関数を実行するための抽象ラッパー
 * function onOpen(){Gantt.onOpenAction();}
 * function onEdit(e){Gantt.onEditAction(e);}
 * function onInstall(e){Gantt.onInstallAction(e);}
 * function onChange(e){Gantt.onChangeAction(e);}
 * function onSelectionChange(e){Gantt.onSelectionChangeAction(e);}
 * function onFormSubmit(e){Gantt.onFormSubmitAction(e);}
 * function doGet(e){Gantt.onGetAction(e);}
 * function doPost(e){Gantt.onPostAction(e);}
 * function abstractFunction(functionId) { Gantt.abstractFunction(functionId); }
 * =================================================================================
 * 
 * 種々の方法で上記の関数が呼ばれ、そこからライブラリの対応する関数が呼ばれる、という仕組みになっています。
 * 
 * ※functionIdという引数名は誤解を招きますが、これは現状は{id,context}という構造のobjectです。
 *  過去にはstringを渡していましたが、それでは関数しか区別できず、その他の情報を渡しにくいという問題が発生したため、現状の仕組みになりました。
 *  これを修正するためにはすべての呼び出し元を変更する必要があるので、放置しています。動作に問題はないです。
 * 
 * 
 * 挙動は以下です。
 * 
 * 1. 個別のスピードプランナーが開かれる
 * 2. そのブックでonOpenが実行される
 * 3. Gantt.onOpenActionが実行される
 * 4. Gantt.onOpenActionが上部にサイドバーメニューを生成するボタンを含むuiを生成する
 * 
 * 5. もし、そのサイドバーのボタンが押されると、個別のスピードプランナーのスクリプトのabstractFunctionにボタンごとに異なるidの入ったオブジェクトが渡され、実行される
 * 6. 個別のスピードプランナーのスクリプトのabstractFunctionが、Gantt.abstractFunctionにidの入ったオブジェクトを横流しして実行する
 * 7. Gantt.abstractFunction(つまりこのライブラリの00[interface]_main.gsに定義されているもの)が渡されたオブジェクトからidとcontextを取り出し、FunctionRunnerServiceに渡して実行する。
 * 
 * 8. FunctionRunnerServiceは、00_FunctionsConfigs.gsに定義された設定から自動でfunctionIdに関数を対応させるマップを生成し、そのマップに渡されたfunctionIdに対応する関数があるかどうかチェック。
 * 9. あった場合は、対応する関数がcontextを受け取るよう設定されていたらcontextをわたし、そうでなければ何も渡さずに実行する。
 * 10. 実行開始がユーザーに通知され、実行開始。実行終了通知も出し、途中でエラーになればエラー文も通知
 * 
 * 引数がない関数については、configオブジェクトから自動でhtmlの設定と、GAS側のidからfunctionへのマップを両方とも自動生成します。
 * したがって、引数がない関数についてはidの対応関係を保証されます。
 * 
 * しかし、引数がある関数については引数の入力方法の多様性から設定で対応するとかえって複雑化すると判断し、htmlは手動で書き換えていただく方式になっています。
 * 
 * 
 * ### 仕組みの理由
 * 
 * 具体的な関数を各スピードプランナーのスクリプトに定義すると更新時（特に機能追加時）に毎回個別のスクリプトを更新する必要があり、手間です。
 * そこで、各スクリプトには抽象関数とonOpenなどの抽象的な関数だけを定義しておき、実際の処理はすべてライブラリに委託することで更新がライブラリのみで完結するようにしています。
 * 
 * ほとんどの機能は引数を必要としないので、それらについてはGAS側の設定オブジェクト配列から自動でhtmlもGASも設定をセットアップするようになっています。
 * 
 * 
 * ### abstractFunction(id)以外の...Action()関数について
 * 
 * abstractFunction(id)以外のonOpenなどのトリガーについては、
 * このライブラリ内のどこかのファイルに上記の呼び出し元の例で示されている、関数の中で呼ばれている関数を定義すれば、このライブラリ内のコードだけで振る舞いを決定できます。
 * (できれば仕分けしたいのでそれは名前が00_から始まるファイルにお願いします)
 * 
 * 例えば、onOpenで上部に生成されるメニューの内容を定義することも可能です。
 * 
 * 
 * ### abstractFunction(idObj)について
 * 
 * abstractFunctionは、渡されたidObjから、id(string)とcontext(object)をとりだし、FunctionRunnerServiceにわたして関数を実行させる関数です。
 * 
 * 当初はidObjではなく、functionIdという関数を識別するための単独の文字列だけ渡していました。
 * そのために各呼び出しもとでは変数の命名がfunctionIdとなっています。
 * しかし、現状、実際に渡されているのはidプロパティとcontextプロパティをもつオブジェクトであり、関数の呼び分けはこのオブジェクトから取り出されたidによって行い、選ばれた関数に必要な場合はcontextを渡しています。
 * contextは関数ごとになにか必要な引数があればここに詰め込むことを想定したオブジェクトで、利用しない関数に対しては破棄され、関数には渡されません。
 * 
 * idに対応する関数は、FunctionRunnerService内の_functionsMapに定義しています。
 * 
 * 個別のスピードプランナーのスクリプトのほうのabstractFunctionの呼び出し元は、このライブラリに定義されている00_sidebar_interface.html内のボタンに紐づくhtml内部のスクリプトです。
 * このhtml内部で、各ボタンが押されたときにabstractFunctionにわたす引数を定義しています。
 * 
 * 
 * ### ファイル構成
 * 
 * - 00[interface]_main.gs
 * 
 * 呼び出し用関数などのグローバル関数置き場。
 * 
 * 
 * -　00_FunctionsRunnerService.gs
 * 
 * 00_FunctionsConfigs.gsの設定を参照して、わたされたfunctionIdとcontextを解釈して適切な関数に処理をわたすクラスを定義している。
 * 基本的に、関数追加時に改修の必要はない。
 * 
 * 
 * - 00_FunctionsConfigs.gs
 * 
 * FunctionsRunnerServiceでの関数の設定の参照と、サイドバーメニューのhtmlのボタン用の設定につかう、ハードコード設定群を定義するファイル。
 * 設定が間違っていないかの実行時チェックが自動的に行われるが、それもここに定義されている。
 * 
 * 
 * - 00_SidebarMenuService.gs
 * 
 * 00_FunctionsConfigs.gsの設定を参照して、htmlに「引数のない関数」の呼び出し用のボタンを注入しつつ、サイドバーメニューを表示するクラスを定義するファイル。
 * 基本的に、関数追加時に改修の必要はない。
 * 
 * 
 * - 00_sidebar_interface.html
 * 
 * サイドバーの構造の定義をするhtmlファイル。引数の必要な関数の呼び出しボタンなどはここに適切なfunctionIdやcontextを定義し、ボタン自体も定義する必要がある。
 * 
 * 
 * - 00_contextFunctions.gs
 * 
 * htmlからcontextを渡されて実行される関数を置くファイル。contextから中身を取り出して本体の関数に渡すようなラッパー関数は雑多かつ特殊なので、
 * 一か所にあるほうがわかりやすいと思ってまとめています。
 * 
 * 
 * 
*/


