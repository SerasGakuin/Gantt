// 01_MatrixUtils.gs : 配列操作系のユーティリティを定義するファイル。
/**
 * # MatrixUtils
 *
 * 矩形の二次元配列(行列)に対する特有の操作を行ってくれるユーティリティクラス。
 * 多くのメソッドはデフォルトでは厳格な引数チェック(渡されたものがそもそも行列でない、など)を行い、早期に詳細なエラーを起こすが、もし性能が欲しい特殊な状況であればチェックをバイパスすることもできる。詳しくは各種メソッドのJSDocを参照のこと。
 *
 * ## メソッド一覧
 *
 * - extractRowAsFlatArray(matrix, rowIndex, isValidationNeeded = true)
 *
 * 元の行列の特定の行を1次元配列として抽出するメソッド。内部で条件のチェックを行い、問題があれば詳細なエラーを出す。
 *
 * - extractColAsFlatArray(matrix, colIndex, isValidationNeeded = true)
 *
 * 元の行列の特定の列を1次元配列として抽出するメソッド。内部で条件のチェックを行い、問題があれば詳細なエラーを出す。
 *
 * - transpose(matrix, isValidationNeeded = true)
 *
 * 転置をするメソッド。転置できない引数を渡された場合に明示的にエラーをだす機能がある。
 */
class MatrixUtils {
  //================================================================================================================================

  static get LOG_TAG() {
    return "MatrixUtils";
  }

  //================================================================================================================================

  /**
   * 行列(1×1以上の長方形の2次元配列)の転置をする関数。
   * 引数のフラグで要求されれば、「Cannot read properties of undefined (reading '0')」のような意味不明のエラーが起こらないように厳重にバリデーションを行う。
   *
   * @param {any[][]} matrix - 転置したい1×1以上の長方形の2次元配列。
   * @param {boolean} [isValidationNeeded = true] - 引数の厳密な検証を望むかどうか。本番実行時にはfalseにしたバージョンを依存性として注入するようにすれば本番時のコストはほぼない。trueだと、かかる時間はfalseのときと比べて2倍になります。(※1)とはいえ通常の30*10程度の行列を数回転置する用途であれば問題ないはずです。
   *
   * @returns {any[][]} - 転置後の二次元配列。
   *
   * @throws {Error} - validationNeededがtrueで、転置できない引数を渡された場合は詳細なエラーを出します。falseの場合はデフォルトのエラーメッセージとなります。
   *
   *
   * @example
   *
   * // 適切な引数
   *
   * const matrix = [[1,2],[3,4]];
   *
   * const transposed = MatrixUtils.transpose(matrix);
   *
   * console.log(transposed); // 結果は[[1,3],[2,4]]
   *
   * // 偽物
   *
   * const fakeArg = 'I am not a matrix!'; // string型
   *
   * const transposed = MatrixUtils.transpose(fakeArg); // エラー！ 引数が配列でありませんでした。引数は1×1以上の長方形の2次元配列である必要があります。: I am not a matrix!
   *
   *
   * @desc
   * ※1:
   * 検証したところ、500x100行列の転置を100回実行すると、バリデーションありでは200ms程度、なしでは100ms程度かかっていました。
   * V8のコンディションに依存するので大きく変動しますが、ほぼ必ず、2倍程度の性能差は確認されました。
   *
   * @public
   */
  static transpose(matrix, isValidationNeeded = true) {
    if (!!isValidationNeeded) this._ensureRectMatrix(matrix); //必要ならバリデーション
    return matrix[0].map((_, colIndex) => matrix.map((row) => row[colIndex])); //結果生成。
  }

  //================================================================================================================================

  /**
   *  渡された二次元配列の特定の行を1次元配列として抽出し、返却する関数。内部の要素は浅いコピーなので注意。
   *
   * @param {any[][]} matrix - ソースとなる行列。
   * @param {number} rowIndex - 抽出したい列番号。内部でNumber()ラップして数字にしたのちにバリデーションするので'0'などでも通るが、思わぬ作用を防ぐためにnumberを推奨
   * @param {boolean} [isValidationNeeded = true] - 引数の厳密な検証を望むかどうか。テスト時にはtrueまたは何も渡さないのを推奨。本番時でも小規模な操作ならわざわざfalseにしないほうが、エッジケースでエラーが出た場合の対応が迅速になる。
   *
   * @returns {any[]} - 抽出された列の一次元配列。
   *
   * @throws {Error} - 不適切な引数を渡された場合はエラーを出します。matrixが矩形の二次元配列でない、rowIndexがの行インデックスとして不正な値、など。
   *
   * @public
   */
  static extractRowAsFlatArray(matrix, rowIndex, isValidationNeeded = true) {
    // 0. 列インデックスの引数をnumberに変換
    const rowIdxNum = Number(rowIndex);
    // 1. 必要なら検証
    if (!!isValidationNeeded) {
      // 1-1. 行列であるかどうかチェック
      this._ensureRectMatrix(matrix);
      // 1-2. 列インデックスが整数でないならアウト
      if (!Number.isInteger(rowIdxNum)) {
        throw new Error(
          `${this.LOG_TAG}: 指定された行インデックスが不正(数字でない)です！\n渡されたインデックス: ${rowIndex}`,
        );
      }
      // 1-3. 列インデックスが[0,行数-1]の範囲でないならアウト
      const maxRowIdx = matrix.length - 1; //行数-1
      if (rowIdxNum < 0 || maxRowIdx < rowIdxNum) {
        throw new Error(
          `${this.LOG_TAG}: 指定された行インデックスが範囲外です！[0,${maxRowIdx}]の範囲に収めてください。\n渡されたインデックス: ${rowIndex}`,
        );
      }
    }
    // 2. 引数が妥当なら、抽出して返却(破壊的操作対策にslice)
    return matrix[rowIdxNum].slice();
  }

  //================================================================================================================================

  /**
   *  渡された二次元配列の特定の列を1次元配列として抽出し、返却する関数。内部の要素は浅いコピーなので注意。
   *
   * @param {any[][]} matrix - ソースとなる行列。
   * @param {number} colIndex - 抽出したい列番号。内部でNumber()ラップして数字にしたのちにバリデーションするので'0'などでも通るが、思わぬ作用を防ぐためにnumberを推奨
   * @param {boolean} [isValidationNeeded = true] - 引数の厳密な検証を望むかどうか。テスト時にはtrueまたは何も渡さないのを推奨。本番時でも小規模な操作ならわざわざfalseにしないほうが、エッジケースでエラーが出た場合の対応が迅速になる。
   *
   * @returns {any[]} - 抽出された列の一次元配列。
   *
   * @throws {Error} - 不適切な引数を渡された場合はエラーを出します。matrixが矩形の二次元配列でない、colIndexが列のインデックスとして不正な値、など。
   *
   * @public
   */
  static extractColAsFlatArray(matrix, colIndex, isValidationNeeded = true) {
    // 0. 列インデックスの引数をnumberに変換
    const colIdxNum = Number(colIndex);
    // 1. 必要なら検証
    if (!!isValidationNeeded) {
      // 1-1. 行列であるかどうかチェック
      this._ensureRectMatrix(matrix);
      // 1-2. 列インデックスが整数でないならアウト
      if (!Number.isInteger(colIdxNum)) {
        throw new Error(
          `${this.LOG_TAG}: 指定された列インデックスが不正(数字でない)です！\n渡されたインデックス: ${colIndex}`,
        );
      }
      // 1-3. 列インデックスが[0,列数-1]の範囲でないならアウト
      const maxColIdx = matrix[0].length - 1; //列数-1
      if (colIdxNum < 0 || maxColIdx < colIdxNum) {
        throw new Error(
          `${this.LOG_TAG}: 指定された列インデックスが範囲外です！[0,${maxColIdx}]の範囲に収めてください。\n渡されたインデックス: ${colIndex}`,
        );
      }
    }
    // 2. 引数が妥当なら、抽出して返却
    return matrix.map((row) => row[colIdxNum]);
  }

  //================================================================================================================================

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
          return "[引数ビュー作成失敗]";
        }
      }
    }
    if (!Array.isArray(target)) {
      throw new Error(
        `\n${this.LOG_TAG}: 引数が配列でありませんでした。引数は1×1以上の長方形の2次元配列である必要があります。:\n${_stringify(target)}`,
      );
    }
    if (target.length === 0) {
      throw new Error(
        `\n${this.LOG_TAG}: 引数が空配列でした。引数は1×1以上の長方形の2次元配列である必要があります。\nなお、[]は1次元とみなされます。:\n${_stringify(target)}`,
      );
    }
    //"行"(内部配列と期待される位置)に配列でないものがないかチェック
    const isThereNotArrayRow = target.some((row) => !Array.isArray(row));
    if (isThereNotArrayRow) {
      throw new Error(
        `\n${this.LOG_TAG}: 引数が1次元配列でした(配列内に配列でない要素がありました)。引数は1×1以上の長方形の2次元配列である必要があります。\nなお、[]は1次元とみなされます。:\n${_stringify(target)}`,
      );
    }
    //各行について中に配列がないかチェックしている。(3次元以上は禁止)
    const isThereArrayInRow = target.some((row) =>
      row.some((itemInRow) => Array.isArray(itemInRow)),
    );
    if (isThereArrayInRow) {
      throw new Error(
        `\n${this.LOG_TAG}: 引数が三次元以上でした(配列のネストが3重になっている部分がありました)。引数は1×1以上の長方形の2次元配列である必要があります。:\n${_stringify(target)}`,
      );
    }
    //列数が不ぞろいの行を検出
    const colsCountSample = target[0].length;
    if (target.some((row) => row.length !== colsCountSample)) {
      throw new Error(
        `\n${this.LOG_TAG}: 引数に列数が不ぞろいな行がありました。引数は1×1以上の長方形の2次元配列である必要があります。:\n${_stringify(target)}`,
      );
    }
    // [[]] など、各行が空配列の場合は処理が非自明で複雑なので非対応。
    if (colsCountSample === 0) {
      throw new Error(
        `\n${this.LOG_TAG}: 引数の列数が0でした。引数は1×1以上の長方形の2次元配列である必要があります。:\n${_stringify(target)}`,
      );
    }
  }

  //================================================================================================================================
}
