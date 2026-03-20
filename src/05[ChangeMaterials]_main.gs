//05[changeMaterials]_main.gs 「教材中途変更マクロ」機能

//呼び出し用
function changeMaterials() {
  Object.freeze(changeMaterialsClass);
  changeMaterialsClass.start();
}

/**
 * 「教材中途変更マクロ」機能を実行するクラス。
 */
const changeMaterialsClass = {
  /**
   * 処理の開始用関数。唯一の窓口。
   */
  start() {
    // 1. スプレッドシートへのアクセスをするためのインスタンスを初期化
    const thisBook = SpreadsheetApp.getActiveSpreadsheet();
    const spIOManager = SheetIO.getSpeedPlannerIOManager(thisBook);
    // 2. 日付情報を取得
    const todayObj = spIOManager.getActiveDate();
    const thisYearNum = todayObj.getFullYear();
    const thisMonthNum = todayObj.getMonth() + 1; //jsの月は0はじまり
    // ログ
    console.log(
      `${changeMaterialsClass._LOGTAG}: シートは、${thisYearNum}/${thisMonthNum}のデータを格納しているものと認識しました。`,
    );
    console.log(
      `${changeMaterialsClass._LOGTAG}: 新しい教材計画の元データの作成を開始します。`,
    );
    // 3. 必要なデータを集める
    //「ID（年月のprefixなし）」「教科名」「参考書名」「月間目標」「週ごとの計画時間(五列)」
    // 「週ごとの実績(五列)」「単位当たり処理量」が入った二次元配列を取得。
    const updatedDataMatrix = this._dataGetter.getUpdatedDataMatrix(
      spIOManager,
      todayObj,
    );
    // 4. データの貼り付け前にシートのバックアップを取る
    spIOManager.monthlySheet_createBackUp(); //月間管理のバックアップをとる
    console.log(
      `${changeMaterialsClass._LOGTAG}: 新しい教材計画の元データ作成完了しました。\n貼りつけ開始します。`,
    );
    // 5. データを張り付ける
    this._dataSetter.setData(spIOManager, todayObj, updatedDataMatrix);
    console.log(`${changeMaterialsClass._LOGTAG}: 貼りつけ完了しました。`);
  },

  //====================================== データ構造定義 ===============================================

  /*
  これはデータを収集するときにどの列indexにどのデータが対応するかを定義するオブジェクト。
  
  以下のstudyTime1などの1,2,3,4,5というのは1週目、2週目、、、の意。
  非連続indexになっても対応できるように範囲ではなく別々に指定している。
  
  NOTE: 将来的に、ひと月が5週以外になることはないということを前提で作っているが、もしあやういなら調整（教材のオブジェクト化など）
  */
  _colToIndex: {
    id: 0,
    subject: 1,
    name: 2,
    monthGoal: 3,
    studyTime1: 4,
    studyTime2: 5,
    studyTime3: 6,
    studyTime4: 7,
    studyTime5: 8,
    achievement1: 9,
    achievement2: 10,
    achievement3: 11,
    achievement4: 12,
    achievement5: 13,
    numPerUnit: 14,
  },

  _colsCount: 15,

  _LOGTAG: "教材中途変更マクロ",

  //==================================== 必要なデータ取得疑似クラス ===============================================

  /**
   * 必要なデータの収集をするクラス。シート上の情報の更新はしない。
   *
   * @private
   */
  _dataGetter: {
    /**
     * 新しい計画としてセットすべき情報を二次元配列の形で返却する関数。
     *
     * @param {SpeedPlannerIOManager} spIOManager - シートとのIOをつかさどるクラスのインスタンス
     * @param {Date} todayObj - 現在のシートの日付の入ったDateオブジェクト
     *
     * @returns {(string|number)[][]} - 新しい計画としてセットすべき情報をまとめた二次元配列。列indexは_colToIndexの設定に従う。
     */
    getUpdatedDataMatrix(spIOManager, todayObj) {
      const thisYearNum = todayObj.getFullYear();
      const thisMonthNum = todayObj.getMonth() + 1; //jsの月は0始まり

      //ガントチャート参照で新しい教材レパートリーと、各種シート参照でいままでのレパートリーのデータをそれぞれ取得。
      //それぞれ「ID（prefixなし）」「科目」「参考書名」「月間目標」「週ごとの計画時間(五列)」「週ごとの実績(五列)」が入った二次元配列。
      const newMaterialDataMatrix = this._getNewMaterialDataMatrix(
        thisYearNum,
        thisMonthNum,
      );
      const oldMaterialDataMatrix = this._getOldMaterialDataMatrix(spIOManager);

      //教材数チェック
      const materialsCapacity = spIOManager.getMaterialsCapacity();
      const finalIdsArray = this._getCombinedIdArray(
        oldMaterialDataMatrix,
        newMaterialDataMatrix,
      );

      if (materialsCapacity < finalIdsArray.length) {
        throw new Error(
          `${changeMaterialsClass._LOGTAG}:教材数が限界値${materialsCapacity}を超えています。\n現在の教材数: ${finalIdsArray.length}`,
        );
      }

      //「ID（prefixなし）」「科目名」「参考書名」「月間目標」「週ごとの計画時間(五列)」「週ごとの実績(五列)」「単位当たり処理量」の二次元配列。
      const finalDataMatrix = this._getFinalMatrix(
        finalIdsArray,
        oldMaterialDataMatrix,
        newMaterialDataMatrix,
      );

      return finalDataMatrix;
    },

    /**
     * 新しい計画の設定を収集して返却する関数。
     *
     * @param {number} targetYearNum - 今データの欲しい年の数字
     * @param {number} targetMonthNum - 今データの欲しい月の数字
     *
     * @returns {(string|number)[][]} - 計画情報をまとめた二次元配列。列indexは_colToIndexの設定に従う。
     *
     * @private
     */
    _getNewMaterialDataMatrix(targetYearNum, targetMonthNum) {
      const colIdx = changeMaterialsClass._colToIndex;
      //minor_projects.gsの関数。今月の学習予定データを取得。これがあたらしい計画内容。
      //「ID」「参考書名」「その月の週ごとの計画時間」が入った二次元配列が返ってくる。
      //NOTE: 非自明な行列操作が続いています。これはgetMonthlyDataが、01_minor_projects.gsに定義されているグローバル関数であり、SheetIOによるシート構成の影響吸収の恩恵を受けられていないためです。詳しくはgetMonthlyData関数のコメントをご覧ください。
      const baseMat = getMonthlyData(
        String(targetYearNum),
        String(targetMonthNum),
      );
      const outputMat = [];
      for (let i = 0; i < baseMat.length; i++) {
        //データがないところは放置
        const row = this._getEmptyRow();
        row[colIdx.id] = baseMat[i][0];
        row[colIdx.name] = baseMat[i][1];
        row[colIdx.studyTime1] = baseMat[i][2];
        row[colIdx.studyTime2] = baseMat[i][3];
        row[colIdx.studyTime3] = baseMat[i][4];
        row[colIdx.studyTime4] = baseMat[i][5];
        row[colIdx.studyTime5] = baseMat[i][6];
        outputMat.push(row);
      }
      Object.freeze(outputMat);
      return outputMat;
    },

    /**
     * 古い元々の計画の設定を収集する関数。返却する二次元配列内の列index構成は_colToIndexにしたがう。
     *
     * @param {SpeedPlannerIOManager} spIOManager - シートとのIOをつかさどるクラスのインスタンス
     *
     * @returns {(string|number)[][]} - 計画情報をまとめた二次元配列。列indexは_colToIndexの設定に従う。
     *
     * @private
     */
    _getOldMaterialDataMatrix(spIOManager) {
      const colIdx = changeMaterialsClass._colToIndex;

      const pureIdsArray = spIOManager.getActiveMaterialsIds().map((v) => v[0]);

      const studyTimeMatrix = spIOManager.getActiveMaterialsStudyTimes();
      const namesMatrix = spIOManager.getActiveMaterialsNames();
      const subjectsMatrix = spIOManager.getActiveMaterialsSubjects();
      const monthGoalsMatrix = spIOManager.getActiveMaterialsMonthGoals();
      const achievementMatrix = spIOManager.getActiveMaterialsAchievements();

      const oldMaterialDataMatrix = [];
      //古いデータを結合する(実績や調整された月間目標などのデータ移行のため)
      //このとき、あとでデータを参照するので、やりやすいように新旧のデータの構成(列と属性の対応)を統一する。
      for (let i = 0; i < pureIdsArray.length; i++) {
        const pureId = pureIdsArray[i]; //prefixなしのid
        if (!pureId) continue;

        const row = this._getEmptyRow();
        row[colIdx.id] = pureId; //prefixなしのid
        row[colIdx.name] = namesMatrix[i][0]; //参考書名
        row[colIdx.subject] = subjectsMatrix[i][0]; //教科名
        row[colIdx.monthGoal] = monthGoalsMatrix[i][0]; //月間目標
        row[colIdx.studyTime1] = studyTimeMatrix[i][0]; //勉強時間 五週間分
        row[colIdx.studyTime2] = studyTimeMatrix[i][1];
        row[colIdx.studyTime3] = studyTimeMatrix[i][2];
        row[colIdx.studyTime4] = studyTimeMatrix[i][3];
        row[colIdx.studyTime5] = studyTimeMatrix[i][4];
        row[colIdx.achievement1] = achievementMatrix[i][0]; //学習実績 五週間分
        row[colIdx.achievement2] = achievementMatrix[i][1];
        row[colIdx.achievement3] = achievementMatrix[i][2];
        row[colIdx.achievement4] = achievementMatrix[i][3];
        row[colIdx.achievement5] = achievementMatrix[i][4];

        oldMaterialDataMatrix.push(row); //データが行に詰まったら行ごとpush
      }
      Object.freeze(oldMaterialDataMatrix);
      return oldMaterialDataMatrix;
    },

    /**
     * 新旧idを結合して最終的な教材のレパートリーの選択をしめすidの配列を返却する。
     * わたすidは年月prefixなしである必要がある(spIOManagerが勝手にprefixぬいてくれるので大丈夫)。
     *
     * @param {(string|number)[][]} oldMaterialDataMatrix - もともとの計画の情報の入った二次元配列
     * @param {(string|number)[][]} newMaterialDataMatrix - 新しい計画の情報の入った二次元配列
     *
     * @returns {string[]} - 最終的に採用されるidをまとめた配列。
     *
     * @private
     */
    _getCombinedIdArray(oldMaterialDataMatrix, newMaterialDataMatrix) {
      const colIdx = changeMaterialsClass._colToIndex;

      ///実績の列indexをSetにしておく。実績がすでに登録されている教材は削除しないようにするため。
      const achievementsColsSet = new Set([
        colIdx.achievement1,
        colIdx.achievement2,
        colIdx.achievement3,
        colIdx.achievement4,
        colIdx.achievement5,
      ]);

      let outputSet = new Set();
      newMaterialDataMatrix.forEach((row) => {
        //基本的には、新教材の順を尊重するので優先的に、新計画の中のidを追加していく
        const id = String(row[colIdx.id]).trim();
        outputSet.add(id);
      });

      oldMaterialDataMatrix.forEach((row) => {
        //実績があるのに新教材から消してしまった古い教材は最後尾に追加する。

        const id = String(row[colIdx.id]).trim();
        if (outputSet.has(id)) return; //すでに追加されてるなら何もしない

        //学習実績のみ抽出抽出して、実績が存在するなら保存する。行内を走査して、実績列について内容があったらtrue.
        const hasAchievement = row.some((value, colIdxNum) => {
          //実績の列でないなら関係ないのでfalse
          const isAchivementCol = achievementsColsSet.has(colIdxNum);
          if (!isAchivementCol) return false;
          //実績の列であり、かつ値がfalsyでないならtrue
          else return !!value;
        });
        if (hasAchievement) outputSet.add(id); //実績があるなら保存のため追加
      });

      const idsArr = Array.from(outputSet);
      Object.freeze(idsArr);
      return idsArr;
    },

    /**
     * 今まで収集した情報をつかって、最終的にデータソースとして出力する配列を返す関数。
     *
     * @param {string[]} finalPureIdsArray -  最終的に出力される計画に含まれるすべてのの参考書id一覧。
     * @param {(string|number)[][]} oldMaterialDataMatrix - もともとの計画の情報の入った二次元配列
     * @param {(string|number)[][]} newMaterialDataMatrix - 新しい計画の情報の入った二次元配列
     *
     * @returns {(string|number)[][]} - 計画情報をまとめた二次元配列。列indexは_colToIndexの設定に従う。
     *
     * TODO: srcRowを検索するマップの構築と、その使用が両方ともこの関数で行われているが、これはマップ構築関数として別離したほうがわかりやすい。この関数はそれを受け取るだけにするとよりよいだろう。
     * @private
     */
    _getFinalMatrix(
      finalPureIdsArray,
      oldMaterialDataMatrix,
      newMaterialDataMatrix,
    ) {
      const colIdx = changeMaterialsClass._colToIndex;

      // IDから「参考書マスター」を検索して、「教科」と「月間目標」と「単位当たり処理量」を取得し、行列に代入
      //TODO:ここは、あたらしいstaticなapiに変えたい。
      const materialsMaster = MaterialsMasterLib.getMaterialsMaster(
        finalPureIdsArray,
        ["教科", "月間目標", "単位当たり処理量"],
      );

      const subjectsArr = materialsMaster.getSubjectsArray();
      const monthlyGoalsArr = materialsMaster.getMonthGoalsArray();
      const numPerUnitArr = materialsMaster.getNumPerUnitArray();

      //データ格納先の二次元配列
      const self = this; //アロー関数内のthisのスコープ問題に対応
      const outputMatrix = Array.from(
        { length: finalPureIdsArray.length },
        () => self._getEmptyRow(),
      );

      //新旧データを参照するために、indexでアクセスする必要があるので、教材idと対応する行のindexの対応mapを作る。
      //getSrcRowFromMapsが参照するので、宣言はここで。
      const oldIdMap = new Map(
        oldMaterialDataMatrix.map((row, idx) => [
          String(row[colIdx.id]).trim(),
          idx,
        ]),
      );
      const newIdMap = new Map(
        newMaterialDataMatrix.map((row, idx) => [
          String(row[colIdx.id]).trim(),
          idx,
        ]),
      );
      /**
       * 特定のidの参考書のソースとなる行を検索し、なければundefinedを返却する補助関数。
       * 優先順位: 既存(old) > 新規(new)
       *
       * @param {string} id - 検索したいid
       *
       * @returns {(string|number)[]} idに対応する情報源
       */
      function getSrcRowFromMaps(id) {
        let srcRow = undefined;
        let rowIdx = oldIdMap.get(id);
        if (rowIdx !== undefined) {
          //oldにデータありならそれを採用
          srcRow = oldMaterialDataMatrix[rowIdx];
        } else {
          //oldにデータなし
          rowIdx = newIdMap.get(id);
          srcRow =
            rowIdx !== undefined ? newMaterialDataMatrix[rowIdx] : undefined;
        }
        return srcRow;
      }

      for (let i = 0; i < finalPureIdsArray.length; i++) {
        const pureId = finalPureIdsArray[i];
        outputMatrix[i][colIdx.id] = pureId || "";

        //参考書マスターからとってきたものを格納
        outputMatrix[i][colIdx.subject] = subjectsArr[i] || "";
        outputMatrix[i][colIdx.monthGoal] = monthlyGoalsArr[i] || "";
        outputMatrix[i][colIdx.numPerUnit] = numPerUnitArr[i] || "";

        //参考書マスターから直接取得したデータの不足をold/new DataMatrixから補う/修正する
        const srcRow = getSrcRowFromMaps(pureId); //まず情報源の行を取得

        //ソース行からデータを取得して格納(上で格納したやつを上書きすることがあるのは想定通りの挙動。)
        if (srcRow !== undefined) {
          outputMatrix[i][colIdx.name] = srcRow[colIdx.name]; //参考書名
          outputMatrix[i][colIdx.subject] = srcRow[colIdx.subject] || ""; //教科名
          outputMatrix[i][colIdx.monthGoal] = srcRow[colIdx.monthGoal] || ""; //月間目標

          outputMatrix[i][colIdx.studyTime1] = srcRow[colIdx.studyTime1]; //勉強時間
          outputMatrix[i][colIdx.studyTime2] = srcRow[colIdx.studyTime2];
          outputMatrix[i][colIdx.studyTime3] = srcRow[colIdx.studyTime3];
          outputMatrix[i][colIdx.studyTime4] = srcRow[colIdx.studyTime4];
          outputMatrix[i][colIdx.studyTime5] = srcRow[colIdx.studyTime5];

          outputMatrix[i][colIdx.achievement1] = srcRow[colIdx.achievement1]; //実績
          outputMatrix[i][colIdx.achievement2] = srcRow[colIdx.achievement2];
          outputMatrix[i][colIdx.achievement3] = srcRow[colIdx.achievement3];
          outputMatrix[i][colIdx.achievement4] = srcRow[colIdx.achievement4];
          outputMatrix[i][colIdx.achievement5] = srcRow[colIdx.achievement5];
        }

        const setMaterialName = outputMatrix[i][colIdx.name];
        const setSubject = outputMatrix[i][colIdx.subject];
        if (setMaterialName && !setSubject) {
          //科目名が空白なら参考書名から推測して補完
          outputMatrix[i][colIdx.subject] =
            MaterialsMasterLib.guessSubjectsByName(setMaterialName);
        }
      }

      /*
      もともとここで科目名に番号をわりふってその番号順にソートしていたが、ガントチャートで定義された順番を尊重すべきとの要件にあわせて、それはやめている。
      */
      Object.freeze(outputMatrix);
      return outputMatrix;
    },

    /**
     * データソースとして必要な行数を満たすために、その補填の空の行配列を作成する関数。
     *
     * @returns {null[]} 空の一次元行配列。
     *
     * @private
     */
    _getEmptyRow() {
      return Array(changeMaterialsClass._colsCount).fill(null);
    },
  },

  //==================================== 必要なデータを格納する疑似クラス ===============================================

  /**
   * ソースとなる配列から必要な情報をとりだし、シートにセットしていくクラス。
   * @private
   */
  _dataSetter: {
    /**
     * データセットを行うクラス。
     *
     * @param {SpeedPlannerIOManager} spIOManager - シートとのIOをつかさどるクラスのインスタンス
     * @param {Date} todayObj - 現在のシートの日付の入ったDateオブジェクト
     * @param {(string|number)[][]} updatedDataMatrix - 情報源となる配列。
     */
    setData(spIOManager, todayObj, updatedDataMatrix) {
      const colIdx = changeMaterialsClass._colToIndex;

      //種類ごとに行列を分割
      const pureIdsMatrix = updatedDataMatrix.map((row) => [row[colIdx.id]]);
      const studyTimeMatrix = updatedDataMatrix.map((row) => [
        row[colIdx.studyTime1],
        row[colIdx.studyTime2],
        row[colIdx.studyTime3],
        row[colIdx.studyTime4],
        row[colIdx.studyTime5],
      ]);
      const subjectsMatrix = updatedDataMatrix.map((row) => [
        row[colIdx.subject],
      ]); // 科目名
      const referenceBookNamesMatrix = updatedDataMatrix.map((row) => [
        row[colIdx.name],
      ]); // 参考書名
      const monthlyGoalsMatrix = updatedDataMatrix.map((row) => [
        row[colIdx.monthGoal],
      ]); // 月間目標

      const numPerUnitMatrix = updatedDataMatrix.map((row) => [
        row[colIdx.numPerUnit],
      ]); // 単位当たり処理量

      // 実績を週ごとに分割
      const achievementMatrix = updatedDataMatrix.map((row) => [
        row[colIdx.achievement1],
        row[colIdx.achievement2],
        row[colIdx.achievement3],
        row[colIdx.achievement4],
        row[colIdx.achievement5],
      ]);

      // 各教材の「月間合計学習時間」を計算
      const totalStudyTimeMatrix = studyTimeMatrix.map((weeks) => {
        const total = weeks.reduce((sum, val) => {
          const num = Number(val);
          return sum + (isNaN(num) ? 0 : num);
        }, 0);
        return [total];
      });

      //貼り付け実行
      spIOManager.setActiveMaterialsStudyTimes(studyTimeMatrix);
      spIOManager.setActiveMaterialsAchievements(achievementMatrix);

      this._updateMonthDataArchive(
        spIOManager,
        todayObj,
        pureIdsMatrix,
        subjectsMatrix,
        referenceBookNamesMatrix,
        monthlyGoalsMatrix,
        totalStudyTimeMatrix,
        numPerUnitMatrix,
      );
    },

    /**
     * アーカイブシート(月間管理シート)に情報をセットする補助関数。
     * NOTE: シート構成の単純化計画が完成されたあと、調整されたときにはここはarchiveではなく、setActive...になるはずである。
     * @private
     */
    _updateMonthDataArchive(
      spIOManager,
      todayObj,
      idsMatrix,
      subjectsMatrix,
      referenceBookNamesMatrix,
      monthlyGoalsMatrix,
      totalStudyTimeMatrix,
      numPerUnitMatrix,
    ) {
      const y = todayObj.getFullYear();
      const m = todayObj.getMonth() + 1;

      spIOManager.deleteMaterialsLogsArchive(y, m);
      spIOManager.registerMaterialsMetadataToArchive(
        y,
        m,
        idsMatrix,
        subjectsMatrix,
        referenceBookNamesMatrix,
        totalStudyTimeMatrix,
        monthlyGoalsMatrix,
        numPerUnitMatrix,
      );
    },
  },
};
