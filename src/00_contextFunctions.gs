/*
  00_contextFunctions.gs

  htmlからabstractFunctionを経由してcontextをわたして関数を実行するときに、その実行命令を直接受けるためのラッパー関数を定義するファイル。
  contextを受け取らない関数なら、ここでcontextから引数をとりだす必要はない。
  例えば月末月初マクロは02番台のファイルにそのまま呼び出し関数を登録している。
  contextから引数を取り出してほかの関数に実行させるだけの関数は、ここにまとめています。

  例:
  function funcX(context){
    //contextから必要な引数を取りだす
    const arg1 = context.arg1;
    const arg2 = context.arg2;
    //取り出した引数を別の場所に定義した処理本体の関数に渡す。
    funcXProcessor(arg1, arg2);
  }
*/
//エラー報告用の関数。
function reportError(context) {
  const message = context.message;
  if (!message) return;
  GASRefferenceSheetLogService.user(message);
}
