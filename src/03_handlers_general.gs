//03_handlers_general.gs ハンドラー置き場
/**
 * 大体のハンドラーを定義するオブジェクト。
 * ファイル外から中身にアクセスする場合は、テスト関数での一時的な目的でない場合は03_handlers.gsのgetterを経由してください。
 */
/** @type {GenWeeklyPlanHandlersContainer} */
const GenWeeklyPlan_handlers_general = Object.freeze({

  /**
   * 勉強予定量が0だったとき
   */
  noStudyHandler: {
    name: 'noStudyHandler',
    run(context) {
      if (!context.valueToAdd) return GenWeeklyPlan_handlers_defaultTexts.onCanceled;
      else return null;
    }
  },

  /**
   *　参考書マスターの情報がなかったとき
   現在、AIに委託するので未使用
   */
  noMaterialInfoHandler: {
    name: 'noMaterialInfoHandler',
    run(context) {
      if (!context.hasValidMetadata) return GenWeeklyPlan_handlers_defaultTexts.onNoMasterData;
      else return null;
    }
  },

  /**
   * 前回の計画が空白だったとき
   */
  emptyHandler: {
    name: 'emptyHandler',
    run(context) {
      if (!context.lastAchievement) return GenWeeklyPlan_handlers_defaultTexts.onConsultationNeeded;
      else return null;
    }
  },

  /**
   * 最新の学習実績に数字がふくまれていないとき（数字を加算して進めることができない）
   */
  noNumberToAddHandler: {
    name: 'noNumberToAddHandler',
    run(context) {
      if (context.numbersCount === 0) return context.lastAchievement;
      return null;
    }
  },

  /**
   * 複数のチルダが含まれているとき。（複数範囲）
   */
  multipleTildasHandler: {
    name: 'multipleTildasHandler',
    run(context) {
      if (context.tildasCount >= 2) return GenWeeklyPlan_handlers_defaultTexts.onConsultationNeeded;
      else return null;
    }
  },

  /**
   * 過去問演習っぽいパターンのとき
   */
  pastExamsHandler: {
    name: 'pastExamsHandler',
    run(context) {
      const last = context.lastAchievement;
      const materialName = context.materialInfo?.name;
      const isPastBook = (typeof materialName === 'string') && (materialName.includes('過去問') || materialName.includes('共通テスト'));

      const hasKakomonPattern =
        (context.numbersCount ?? 0) >= 2 && /\d{4}.{0,4}\d+回/.test(last);

      if (hasKakomonPattern || isPastBook) {
        return last + GenWeeklyPlan_handlers_defaultTexts.onConsultationNeeded;
      }
      return null;
    }
  },
  /**
   *　「の該当範囲」パターンで、改行無し。（改行があると2教材が一つにまとめられているパターンの可能性が高い）
   */
  rangeOfAnotherTextsHandler: {
    name: 'rangeOfAnotherTextsHandler',
    run(context) {
      const last = context.lastAchievement;

      const isRangeOfIncluded = last.includes('の該当範囲');
      const isMultiRows = last.includes('\n');

      if (isRangeOfIncluded && !isMultiRows) return last;

      return null;
    }
  },

  /**
   * 章をまたいでも問題番号のカウントがリセットしない教材で、
   * 数字~数字という問題番号二つで範囲が示されているとき。
  */
  probNumberRangeHandler: {
    name: 'probNumberRangeHandler',
    run(context) {
      //まず教材の性質で判定
      const isNotTargetMaterial = !context.hasValidMetadata || context.isCountedOnlyByChapter || context.isCountingRepeating;
      if (isNotTargetMaterial) return null;//自分の担当外ならnull

      // [数字前テキスト, 最初のチルダ前の数字, チルダ後の二個目の数字, 二個目以降のテキスト]の配列に分解して返す
      const dividedLogTexts = this._getComponents(context.lastAchievement);
      if (!dividedLogTexts) return null;//マッチしなかった場合
      const firstStr = dividedLogTexts[0];
      const endNumber = Number(dividedLogTexts[2]);
      const endStr = dividedLogTexts[3];

      const endNums = context.materialInfo.chapterEnds;//各章の最後の問題の問題番号の配列
      const finalNum = Number(endNums[endNums.length - 1]);//最後の問題番号。連続する問題番号を期待しているのでこれを天井にする

      if (finalNum === 0) return GenWeeklyPlan_handlers_defaultTexts.onError;//最後の番号の入力忘れが疑われる

      //すでに終わってた場合
      const isAlreadyFinished = finalNum <= endNumber;
      if (isAlreadyFinished) return GenWeeklyPlan_handlers_defaultTexts.onFinished;

      //加算してあたらしい範囲を導出。
      const newStartNumber = endNumber + 1;
      const newEndNumber = Math.min(finalNum, (endNumber + context.valueToAdd));

      //次の計画の範囲。「数字~数字」または「数字」
      const range = (newStartNumber === newEndNumber) ?
        newStartNumber :
        newStartNumber + '~' + newEndNumber;

      //基本的なアウトプットテキスト。
      const baseText = firstStr + range + endStr;

      //次の範囲で終わりなら終了テキストを付加。
      const willFinish = newEndNumber === finalNum;
      if (willFinish) return baseText + GenWeeklyPlan_handlers_defaultTexts.onFinished;
      else return baseText;
    },

    /**
    * [数字前テキスト, 最初のチルダ前の数字, チルダ後の二個目の数字, 二個目以降のテキスト]の配列に分解して返す
    */
    _getComponents(inputString) {
      const rangeMatch = inputString.match(/^(.*?)(\d+)~(\d+)(.*)$/);
      if (rangeMatch) {//マッチした。
        return [
          rangeMatch[1],           // 数字前のテキスト
          parseInt(rangeMatch[2]), // 最初の数字
          parseInt(rangeMatch[3]), // 二個目の数字
          rangeMatch[4]            // 二個目の数字の後のテキスト
        ];
      } else {
        return null;//マッチしなかった。エラー
      }
    }
  },

  /**
   * 章をまたいでも問題番号のカウントがリセットしない教材で、
   * 数字という問題番号一つで問題一つだけが示されているとき。
   * 数字一つしかない場合はミスマッチの発生率が高いので、数字の直後に対応する参考書の単位がついている必要がある。
  */
  probSingleNumberHandler: {
    name: 'probSingleNumberHandler',
    run(context) {
      //まず教材の性質で判定
      const isNotTargetMaterial = !context.hasValidMetadata || context.isCountedOnlyByChapter || context.isCountingRepeating;
      if (isNotTargetMaterial) return null;//自分の担当外ならnull

      if (context.tildasCount !== 0) return null;//チルダがある場合はキャンセル

      // [数字前, マッチ数字, 数字後] に分解。単位が合わなければnull。
      const divided = this._getComponents(context.lastAchievement, context.materialInfo);
      Logger.log(divided);
      if (!divided) return null;

      const [firstStr, currentNumber, endStr] = divided;

      // マスターデータから最終番号を取得して天井にする
      const endNums = context.materialInfo.chapterEnds;
      const finalNum = Number(endNums[endNums.length - 1]);

      if (finalNum === 0) return GenWeeklyPlan_handlers_defaultTexts.onError;//最後の番号の入力忘れが疑われる

      // すでに完了している場合
      if (currentNumber >= finalNum) return GenWeeklyPlan_handlers_defaultTexts.onFinished;

      // 新しい範囲を計算
      const newStartNumber = currentNumber + 1;
      const newEndNumber = Math.min(finalNum, currentNumber + context.valueToAdd);

      // 範囲文字列を作成（1問だけなら単一数字、それ以外は範囲形式）
      const range = (newStartNumber === newEndNumber) ?
        newStartNumber :
        newStartNumber + '~' + newEndNumber;

      // 文字列を再構築
      const baseText = firstStr + range + endStr;

      // 終了判定を付加
      return (newEndNumber === finalNum) ?
        baseText + GenWeeklyPlan_handlers_defaultTexts.onFinished :
        baseText;
    },

    /**
     * 単位（番号の数え方）を頼りに文字列を3分割する。
     * 数字が一つだけの場合、「数3」などをひっかけるリスクが非常に高いため、3だけではマッチせず、3の直後にその教材の問題のカウント単位がついている場合のみマッチ。「3講」など。
     * 参考書マスターに「講」とあるのなら、「回」など別の数え方ではマッチしない。
     * [数字前のテキスト、数字、数字後のテキスト]
     */
    _getComponents(inputString, materialInfo) {
      const unit = materialInfo?.numberingType[0];
      if (!unit) return null;

      const escapedUnit = String(unit).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

      // 数字と単位の間の空白はマッチングには含めるが、返り値では消える。（残す必要がないので）
      let pattern = new RegExp(`^(.*?)(\\d+)\\s*${escapedUnit}(.*)$`);
      let match = inputString.match(pattern);

      if (match) {
        return [
          match[1],           // 数字前のテキスト（例：「第」）
          parseInt(match[2]), // 数字（例：4）
          unit + match[3]     // 単位以降（例：「回」）※空白はここに含まれない
        ];
      }
      //数字+単位でヒットしない場合は、単位+数字で再度（No.2みたいなケース）
      pattern = new RegExp(`^(.*?)${escapedUnit}\\s*(\\d+)(.*)$`);
      match = inputString.match(pattern);

      if (match) {
        return [
          match[1] + unit,
          parseInt(match[2]), // 数字（例：4）
          match[3]
        ];
      }

      return null;
    }
  },

  /**
   * TODO: 
   * Focus Gold Smart 数C2 gMB078 (数C：例題C1.25~C.29) に対応できていない。
   * isCountingRepeatingの判定が、全問題が同じ問題番号から開始することを前提としているため。
   * 章の表示同様、問題番号もstartNumsをみてindexと表示用途ラベルで変換し、章内部の相対indexで評価すべきかもしれない。
   * しかしこのパターンは少ないのでいったん保留。
   * 
   * 章ごとに問題番号のカウントアップがリセットされるタイプを想定。
   * 数字X数字~数字X数字 のパターンを処理する疑似クラス。
   * 例: 3-5~4-1,  Chapter n 単元m ~ Chapter N 単元M
  */
  chapProbRangeHandler: {
    name: 'chapProbRangeHandler',
    //処理の開始用メソッド。
    run(context) {

      const isNotTargetMaterial = !context.hasValidMetadata || context.isCountedOnlyByChapter || !context.isCountingRepeating || context.numbersCount < 4;
      if (isNotTargetMaterial) return null;

      const divided = this._getComponents(context.lastAchievement);
      if (!divided) return null;

      const [
        firstText, commonUnit, tildeText, lastText,
        chap1Num, problem1, chap2Num, problem2
      ] = divided;

      const chapterNames = context.materialInfo.chapters;
      const endNums = context.materialInfo.chapterEnds;

      if (!chapterNames || !endNums || !Array.isArray(chapterNames)) return null;//この方式はキャンセル

      const chap2Index = chapterNames.indexOf(chap2Num);//配列内indexに変換
      if (chap2Index === -1) return null;//ヒットしないならエラー

      const lastChapterIndex = endNums.length - 1;
      const lastProblemNum = endNums[lastChapterIndex];

      // すでに終了しているか
      if (
        chap2Index == lastChapterIndex &&
        problem2 == lastProblemNum
      ) {
        return GenWeeklyPlan_handlers_defaultTexts.onFinished;
      }

      // === 次の開始位置（indexベース） ===
      let startChapterIndex;
      let startProblemNum;

      const isEndOfChapter = problem2 == endNums[chap2Index];

      if (isEndOfChapter) {
        startChapterIndex = chap2Index + 1;
        startProblemNum = 1;
      } else {
        startChapterIndex = chap2Index;
        startProblemNum = problem2 + 1;
      }

      // === 次の終了位置を計算 ===
      const [endChapterIndex, endProblemNum] =
        this._getWhereProblemIsByIndex(
          startChapterIndex,
          startProblemNum,
          context.valueToAdd,
          endNums
        );

      const willFinish = endChapterIndex === lastChapterIndex && endProblemNum === lastProblemNum;

      // === 表示用に変換 ===
      const startChapterLabel = chapterNames[startChapterIndex];
      const endChapterLabel = chapterNames[endChapterIndex];

      const isChapterOutOfRange = startChapterLabel === undefined || endChapterLabel === undefined;
      if (isChapterOutOfRange) return null;//万が一、章が範囲外なら

      const baseResult = this._constructBaseResult(firstText, startChapterLabel, commonUnit, startProblemNum,
        tildeText, endChapterLabel, endProblemNum, lastText);

      return willFinish ?
        baseResult + GenWeeklyPlan_handlers_defaultTexts.onFinished :
        baseResult;
    },

    /**
     * 渡されたテキストを
    * [数字前テキスト, 二組の数字の間に共通するテキスト（'章第'など）,二組の数字の間のテキスト,数字後のテキスト,
         数字1,数字2,数字3,,数字4 ]の配列に分解して返す。
    * 数字の間に挟まっているテキストは完全一致する必要がある。
    * 例:
    *   3-4~4-5なら、「-」は完全一致。
    *   3章第4問~4章第2問なら、「章第」は完全一致。
    */
    _getComponents(inputString) {
      //数字で分割し、数字自体もぬきだす。数字が最初にある場合は最初の要素は""になる
      const parts = inputString.split(/(\d+)/);
      /*
      もしinputStringが適する形になっていれば、
      [..., 数字, テキストA, 数字, ~を含むテキスト, 数字, テキストA, 数字, ...]
      の形になっていると想定される。
      まず~を含むテキストの位置を発見して、その前後のパーツを見ることによって適切に数字を抜き出す。
       */

      //~がどこにあるかチェック。
      let tildePos = -1;
      for (let i = 4; i < parts.length - 2; i++) {//4番目からparts.length-2番目の間にしか~はあり得ない。
        if (parts[i].includes('~')) {
          tildePos = i;
          break;
        }
      }
      if (tildePos === -1) return null;//もしチルダが適切な範囲に発見されなければnull

      // セパレータ（章第、- など）の一致確認。ここが一致しない場合、実際には章の表示ではない可能性がある。ややきつい条件だが安全のため。
      const firstChapUnit = parts[tildePos - 2].trim();//最初と次の章のカウント部分。
      const secondChapUnit = parts[tildePos + 2].trim();
      if (firstChapUnit != secondChapUnit) return null;//もし一個目と二個目で章の数え方であるはずの部分が不一致ならnull

      //一致している場合でも、もともとが空白のみの場合はtrim後のものでは不格好なのでここで最終出力を確認。
      const commonUnit = !firstChapUnit ? parts[tildePos - 2] : firstChapUnit;

      // 外側のテキストを回収
      const firstText = parts.slice(0, tildePos - 3).join('');
      const finalText = parts.slice(tildePos + 4).join('');

      return [
        firstText, commonUnit, parts[tildePos], finalText,
        Number(parts[tildePos - 3]), Number(parts[tildePos - 1]),
        Number(parts[tildePos + 1]), Number(parts[tildePos + 3])
      ];

    },

    /**
     * 次の範囲の終わりの問題が、何章の問題番号何にあるかを計算するメソッド。[章番号,問題番号]で返す。 
     * 最初の問題の章と問題番号から章ごとの問題数を加算していき、最初に目標数を超えた章を返す。
     * endNumsの配列インデックスではなくて章番号（インデックス+1）を返す。
     */
    _getWhereProblemIsByIndex(startChapterIndex, startProblemNum, valueToAdd, endNums) {
      let remaining = valueToAdd;

      //console.log('remaining '+remaining);

      // 最初の章で残り何問あるか
      let firstRemain = Number(endNums[startChapterIndex]) - startProblemNum + 1;
      //console.log('firstRemain '+firstRemain);

      if (remaining <= firstRemain) {
        return [startChapterIndex, startProblemNum + remaining - 1];
      }

      remaining -= firstRemain;

      //console.log('remaining '+remaining);

      // 次の章以降
      for (let i = startChapterIndex + 1; i < endNums.length; i++) {
        const problemsInThisChapter = Number(endNums[i]); // 最初の問題番号が1なら endNums[i] が問題数
        if (remaining <= problemsInThisChapter) {
          return [i, remaining];
        }
        remaining -= problemsInThisChapter;
        //console.log('remaining '+remaining);
      }

      // すべて終わる場合
      const lastChapterIndex = endNums.length - 1;
      return [lastChapterIndex, endNums[lastChapterIndex]];
    },

    /**
     * 情報から適切に結果のテキストを組み立てる。
     */
    _constructBaseResult(firstText, startChapterLabel, commonUnit, startProblemNum,
      tildeText, endChapterLabel, endProblemNum, lastText) {
      const isSingleProb =
        startChapterLabel === endChapterLabel && startProblemNum === endProblemNum;

      const rangeText = isSingleProb ?
        startChapterLabel + commonUnit + startProblemNum :
        startChapterLabel + commonUnit + startProblemNum +
        tildeText +
        endChapterLabel + commonUnit + endProblemNum;

      return firstText + rangeText + lastText;
    }

  }

});





