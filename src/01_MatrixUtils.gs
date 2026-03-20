// 01_MatrixUtils.gs : 配列操作系のユーティリティを定義するファイル。
/*
メソッド一覧
transpose: 配列の転置用のメソッド。転置できない引数を渡された場合に明示的にエラーをだす機能がある。
*/

class MatrixUtils {
  static get LOG_TAG(){return 'MatrixUtils'}

  /**
   * 行列(1×1以上の長方形の2次元配列)の転置をする関数。
   * 引数のフラグで要求されれば、「Cannot read properties of undefined (reading '0')」のような意味不明のエラーが起こらないように厳重にバリデーションを行う。
   * 
   * @param {any[][]} matrix - 転置したい1×1以上の長方形の2次元配列。
   * @param {boolean} [isValidationNeeded = true] - 転置できるかどうかの厳密なバリデーションを望むかどうか。本番実行時にはfalseにしたバージョンを依存性として注入するようにすれば本番時のコストはほぼない。trueだと、かかる時間はfalseのときと比べて2倍になります。(※1)とはいえ通常の30*10程度の行列を数回転置する用途であれば問題ないはずです。
   * @returns {any[][]} - 転置後の二次元配列。
   * @throws {Error} - validationNeededがtrueで、転置できない引数を渡された場合は詳細なエラーを出します。falseの場合はデフォルトのエラーメッセージとなります。
   * @desc
   * ※1:
   * 検証したところ、500x100行列の転置を100回実行すると、バリデーションありでは200ms程度、なしでは100ms程度かかっていました。
   * V8のコンディションに依存するので大きく変動しますが、ほぼ必ず、2倍程度の性能差は確認されました。
   */
  static transpose(matrix, isValidationNeeded = true) {
    if (isValidationNeeded) this._ensureRectMatrix(matrix);//必要ならバリデーション
    return matrix[0].map((_, colIndex) => matrix.map(row => row[colIndex]));//結果生成。
  }

  /**
   * 引数が長方形の（列数のそろった）二次元配列であるかどうか確認し、もしそうでなければエラーを出すメソッド。
   * 
   * @param {any} target - 調査対象
   * @throws {Error} 渡されたものが長方形の二次元配列でなければエラー
   * @private 
   */
  static _ensureRectMatrix(target) {
    function _stringify(target) {
      try {
        return JSON.stringify(target);
      } catch (e) {
        try {
          return String(target);
        } catch (e) {
          return '[引数ビュー作成失敗]'
        }
      }
    }
    if (!Array.isArray(target)) {
      throw new Error(`\n${this.LOG_TAG}: 引数が配列でありませんでした。引数は1×1以上の長方形の2次元配列である必要があります。:\n${_stringify(target)}`)
    }
    if (target.length === 0) {
      throw new Error(`\n${this.LOG_TAG}: 引数が空配列でした。引数は1×1以上の長方形の2次元配列である必要があります。\nなお、[]は1次元とみなされます。:\n${_stringify(target)}`);
    }
    //"行"(内部配列と期待される位置)に配列でないものがないかチェック
    const isThereNotArrayRow = target.some(row => !Array.isArray(row));
    if (isThereNotArrayRow) {
      throw new Error(`\n${this.LOG_TAG}: 引数が1次元配列でした(配列内に配列でない要素がありました)。引数は1×1以上の長方形の2次元配列である必要があります。\nなお、[]は1次元とみなされます。:\n${_stringify(target)}`);
    }
    //各行について中に配列がないかチェックしている。(3次元以上は禁止)
    const isThereArrayInRow = target.some(row => row.some(itemInRow => Array.isArray(itemInRow)));
    if (isThereArrayInRow) {
      throw new Error(`\n${this.LOG_TAG}: 引数が三次元以上でした(配列のネストが3重になっている部分がありました)。引数は1×1以上の長方形の2次元配列である必要があります。:\n${_stringify(target)}`);
    }
    //列数が不ぞろいの行を検出
    const colsCountSample = target[0].length;
    if (target.some(row => row.length !== colsCountSample)) {
      throw new Error(`\n${this.LOG_TAG}: 引数に列数が不ぞろいな行がありました。引数は1×1以上の長方形の2次元配列である必要があります。:\n${_stringify(target)}`);
    }
    // [[]] など、各行が空配列の場合は処理が非自明で複雑なので非対応。
    if (colsCountSample === 0) {
      throw new Error(`\n${this.LOG_TAG}: 引数の列数が0でした。引数は1×1以上の長方形の2次元配列である必要があります。:\n${_stringify(target)}`);
    }
  }

}
