// 02[gesho]_main.gs
/**
 * 「月末月初マクロ」機能
 * 詳細な説明は02[Document].gsに
*/

/**
 * ガントチャートの「ある」ブックに対する月末月初
 */
function geshoWithGantt() {
  _doProdGesho(true);
}

/**
 * ガントチャートの「ない」ブックに対する月末月初
 */
function geshoWithoutGantt() {//旧 forexcel
  _doProdGesho(false);
}

/**
 * geshoの本番実行用関数
 * @param {boolean} isGanttChartUsed - ガントチャートを使用するか否か。
 */
function _doProdGesho(isGanttChartUsed){
  const dependencies = getGeshoProdDependencies();
  Gesho.start(!!isGanttChartUsed, dependencies);
}

/**
 * 月末月初の処理をするクラス。
 * 
 * NOTE: 
 * getRangeの範囲のハードコードは、じきに廃止される部分に限定されています。
 * 詳しくは: https://docs.google.com/presentation/d/1WpC_waQCcmiH029292JFEQT1Unb2i_QpyylXtDhnpzk/edit?slide=id.p#slide=id.p
 */
class Gesho {
  /**
   * 月末月初処理を開始するためのメソッド。
   * @param {boolean} isGanttChartUsed - ガントチャートを使用する生徒のスピードプランナーかどうか。
   * @param {GeshoDependencies} dependencies - 月末月初用の依存関数
   * @public
   */
  static start(isGanttChartUsed, dependencies) {//「ガントチャート」を使用しないならfalse、使用するならtrue

    if (isGanttChartUsed) Logger.log(`ガントチャートの「ある」ブックに対する月末月初処理を実行します`);
    else Logger.log(`ガントチャートの「ない」ブックに対する月末月初処理を実行します`)

    console.log(dependencies)

    const spIOManager = dependencies.getSpeedPlannerIOManager();

    this._validateGeshoState(isGanttChartUsed, spIOManager);//実行前に必須な構成要素が存在することをチェック。

    //======================= 日付データを確定 ==========================

    // 必要な日付データを取得。JavaScriptのDate型のオブジェクトでgetMonth()すると、0から始まる月が取得されるので注意。つまり一月は0になる。
    const theDateBeforeGesho = spIOManager.getActiveYearMonth();

    const yearBeforeGeshoNum = theDateBeforeGesho.getFullYear();// 例: 2025
    const monthBeforeGeshoNum = theDateBeforeGesho.getMonth() + 1;

    const yearAfterGeshoNum = (monthBeforeGeshoNum === 12) ? yearBeforeGeshoNum + 1 : yearBeforeGeshoNum;
    const monthAfterGeshoNum = (monthBeforeGeshoNum === 12) ? 1 : monthBeforeGeshoNum + 1;

    const theDateAfterGesho = new Date(yearAfterGeshoNum, (monthAfterGeshoNum - 1), 15);//年月が正しければいいので日は適当にちょっとずれても年月が変化ないように真ん中あたり

    Logger.log(`月末月初実行前の年月は、年:${yearBeforeGeshoNum}, 月:${monthBeforeGeshoNum}です。`);
    Logger.log(`月末月初実行後の年月は、年:${yearAfterGeshoNum}, 月:${monthAfterGeshoNum}です。`);

    //======================= 処理本体開始 ==========================

    //週間管理の古いデータを月間管理と月初に格納して、週間管理からは削除する。
    this._archiveWeeklyData(spIOManager, yearBeforeGeshoNum, monthBeforeGeshoNum);

    //今月プランの古いデータを今月実績や月初に移動
    this._archiveThisMonthPlanData(spIOManager, yearBeforeGeshoNum, monthBeforeGeshoNum);

    //ガントチャートの情報を利用するか否かで分岐。来月の教材セット
    if (isGanttChartUsed) {
      this._setNextMonthMaterialsWithGantt(spIOManager, yearAfterGeshoNum, monthAfterGeshoNum, dependencies);
    } else {
      this._setNextMonthMaterialsWithoutGantt(spIOManager, yearAfterGeshoNum, monthAfterGeshoNum, dependencies);
    }

    //日付関連の情報を更新。
    //ここまでに「今月のデータ」(アクティブなデータ)を取得して格納する場面があり、日付の設定に「今月のデータ」は追従するので、日付のセットはgeshoの最後に行う必要がある。
    //TODO: この必要性はシートの構成に由来する。たとえば週間管理は教材のメタデータを持たず、月間管理からその月のデータをただ表示しているので、日付によってここから取得される内容が変化してしまう。
    //将来的には週間管理は「今月のデータを一時的に作業用に保管する作業テーブル」にし、月間管理は過去のデータのアーカイブに専念させるので、それによってこの複雑さは解消される、と期待される
    this._setDateInfo(spIOManager, theDateAfterGesho);

    //なにか一定期間のみ行いたい更新があればこのクラス（02_updateOnGesho.gsのクラス）に定義しておけばここで実行される。更新がない時でも置いておいて下さい。
    dependencies.updateOnGesho(spIOManager);

  }

  /**
   * 実行前に必須なシートがそろっていることをチェックする補助関数
   * 
   * @param {boolean} isGanttChartUsed - ガントチャートを使用する生徒のスピードプランナーかどうか。
   * @param {SpeedPlannerIOManager} spIOManager - シートIO用のマネージャー
   * @throws {Error} 必須なシートが存在しない場合はエラーを出す
   * @private
   */
  static _validateGeshoState(isGanttChartUsed, spIOManager) {
    const book = spIOManager.book;
    const required = ["今月プラン", "週間管理", "月間管理", "月初", "今月実績", "月間実績"];
    const notFoundSheetNames = [];
    required.forEach(name => {
      const s = book.getSheetByName(name);
      if (!s) notFoundSheetNames.push(name);
    });
    if (isGanttChartUsed) {
      const s = book.getSheetByName("ガントチャート");
      if (!s) notFoundSheetNames.push("ガントチャート");
    }
    if (notFoundSheetNames.length !== 0) throw new Error(`以下の必須シートが見つかりません: ${notFoundSheetNames.join(', ')}`);
  }

  //週間管理の古いデータを月間管理と月初に格納して、週間管理からは削除する。
  static _archiveWeeklyData(spIOManager, yearBeforeGeshoNum, monthBeforeGeshoNum) {
    console.log('週間管理の古いデータを月間管理と月初に移動します。');
    //週間管理シートの末尾にまとまっている各週の学習実績を取得　5週間分。行が教材で列が週
    const achievements = spIOManager.getActiveMaterialsAchievements();

    //TODO: 月初に先月の実績を貼り付け(不都合がなければ廃止予定なのでBJ1のベタ打ちも放置)
    const colBJ = spIOManager.monthlyFirstSheet.getRange("BJ1").getColumn();
    //月初シートの右端の実績貼り付けエリアの左端
    spIOManager.monthlyFirstSheet.getRange(4, colBJ, achievements.length, achievements[0].length).setValues(achievements);

    // 週間管理の学習実績を月間管理に格納する
    spIOManager.archiveMaterialsAchievements(yearBeforeGeshoNum, monthBeforeGeshoNum, achievements);

    //週間管理の実績データをクリア。
    spIOManager.setActiveMaterialsAchievements(null);
  }

  //今月プランのデータを月初と今月実績にコピー TODO: この関数はシート構成の改良をすると不要になります。詳しくはDocumentを参照のこと
  static _archiveThisMonthPlanData(spIOManager, yearBeforeGeshoNum, monthBeforeGeshoNum) {
    Logger.log('今月プランの古いデータを今月実績や月初にコピーします。');
    const prefix = spIOManager.getYearMonthPrefix(yearBeforeGeshoNum, monthBeforeGeshoNum);

    const lastIDsMatrix = spIOManager.getActiveMaterialsIds().map(row => [prefix + row[0]]);
    const lastTimeMatrix = spIOManager.getActiveMaterialsStudyTimes();

    //TODO: この二行については今後廃止を検討している（現在も不要かもしれない）ので、sheetIOへの統合は行わないと思います
    spIOManager.thisMonthAchievementSheet.getRange(4, 1, lastIDsMatrix.length, 1).setValues(lastIDsMatrix)//先月の教材を今月プランから今月実績に貼りつけ。A4起点。
    spIOManager.thisMonthAchievementSheet.getRange(4, 4, lastTimeMatrix.length, 5).setValues(lastTimeMatrix)//今月実績に勉強時間貼りつけ。D4起点。
  }

  // 日付関連の情報を書き込み
  static _setDateInfo(spIOManager, theDateAfterGesho) {
    Logger.log('各シートの週の指定を更新します。');
    spIOManager.setActiveWeek(1); // 「1週」にリセット
    Logger.log('各シートの年月の指定を更新します。');
    spIOManager.setActiveYearMonth(theDateAfterGesho);//各シートの日付情報を特定の年月に適合させる。
  }

  //ガントチャートを参照して、来月の教材データを準備格納
  static _setNextMonthMaterialsWithGantt(spIOManager, yearAfterGeshoNum, monthAfterGeshoNum, dependencies) {
    Logger.log('ガントチャートを参照して必要なデータを取得します。');

    //TODO: getMonthlyDataは、01_minor_projects.gsに定義されています。この関数は責務としてはSheetIOに分離したい。indexでアクセスするのも危険。
    // この関数は、「ガントチャート」を参照して特定の年月の情報を取得しています。
    //「ID」「参考書名」「その月の週ごとの計画時間」が入った二次元配列
    const materialsMatrix = dependencies.getMonthlyData(yearAfterGeshoNum, monthAfterGeshoNum);
    if (!materialsMatrix || materialsMatrix.length === 0) {
      throw new Error("「ガントチャート」の来月のデータ（来月に使用する参考書のデータ、それぞれの週ごとの勉強時間、累計勉強時間）が空です。");
    }
    Logger.log('データを取得できました。');

    //setValues用の行列を作成 
    //NOTE: 非自明な行列操作が続いています。これはgetMonthlyDataが、01_minor_projects.gsに定義されているグローバル関数であり、SheetIOによるシート構成の影響吸収の恩恵を受けられていないためです。詳しくはgetMonthlyData関数のコメントをご覧ください。
    const refBookIdsFlatArray = materialsMatrix.map(row => row[0]);//一列目が年月prefixなしのid
    const referenceBookIds = refBookIdsFlatArray.map(value => [value]); // 参考書ID
    const referenceBookNames = materialsMatrix.map(row => [row[1]]); // 参考書名
    const studyTimes = materialsMatrix.map(row => [row[2], row[3], row[4], row[5], row[6]]);//各週の勉強時間
    const totalStudyTimes = studyTimes.map(row => {//合計勉強時間
      const timeSum = row.reduce((acc, v) => {
        const num = Number(v);
        return acc + (Number.isFinite(num) ? num : 0);
      }, 0);
      return [timeSum];
    });

    console.log('参考書マスターを参照して必要なデータを取得します。');
    // IDから「参考書マスター」を検索して、「月間目標」と「単位当たり処理量」を取得する。
    //TODO:ここは、あたらしいstaticなapiに変えたい
    const materialsMaster = dependencies.getMaterialsMaster(refBookIdsFlatArray, ["教科", "月間目標", "単位当たり処理量"]);
    console.log('データを取得できました。');
    const subjects = materialsMaster.getSubjectsArray().map(v => [v]); //各要素を配列に包んで一列の二次元化
    const monthlyGoals = materialsMaster.getMonthGoalsArray().map(v => [v]);
    const numPerUnit = materialsMaster.getNumPerUnitArray().map(v => [v]);

    //参考書マスターに科目データのない教材の科目名推測
    this._guessSubjects(subjects, referenceBookNames, dependencies);

    const materialsCapacity = spIOManager.getMaterialsCapacity();

    //条件;今月プランの限界行数を突破しなければ
    if (refBookIdsFlatArray.length <= materialsCapacity) {
      //来月に学習予定の教材のデータを「月間管理」にセット
      spIOManager.registerMaterialsMetadataToArchive(
        yearAfterGeshoNum, monthAfterGeshoNum,//年月prefixのついていない状態
        referenceBookIds, subjects, referenceBookNames, totalStudyTimes, monthlyGoals, numPerUnit
      );

      // 「今月プラン」に第1週から第5週までの勉強時間を貼り付け
      spIOManager.setActiveMaterialsStudyTimes(studyTimes);
    } else {
      console.log("教材数が限界を超えています。来月のデータの自動代入はできないので、先月データのクリアのみを行います。\n");
      spIOManager.thisMonthPlanSheet.getRange(4, 4).setValue("行数が今月プランの限界を超えていたのでデータのアーカイブのみを行いました。");
    }
  }


  //ガントチャートのないブックでも来月の教材をセットするための関数。
  //先月のものをそのまま移行するだけ。
  static _setNextMonthMaterialsWithoutGantt(spIOManager, yearAfterGeshoNum, monthAfterGeshoNum, dependencies) {

    const pureIdsMatrix = spIOManager.getActiveMaterialsIds();
    const refBookIdsFlatArray = pureIdsMatrix.map(row => row[0]);

    const studyTimesMatrix = spIOManager.getActiveMaterialsStudyTimes();
    const totalStudyTimeMatrix = studyTimesMatrix.map(row => {
      const totalNum = row.reduce((sum, val) => sum + val, 0);
      if (totalNum != 0) return [totalNum];
      else return [null];
    }
    );
    const namesMatrix = spIOManager.getActiveMaterialsNames();

    // IDから「参考書マスター」を検索して、「月間目標」と「単位当たり処理量」を取得する。
    //TODO:ここは、あたらしいstaticなapiに変えたい
    const materialsMaster = dependencies.getMaterialsMaster(refBookIdsFlatArray, ["教科", "月間目標", "単位当たり処理量"]);
    console.log('データを取得できました。');
    const subjects = materialsMaster.getSubjectsArray().map(v => [v]); //各要素を配列に包んで一列の二次元化
    const monthlyGoals = materialsMaster.getMonthGoalsArray().map(v => [v]);
    const numPerUnit = materialsMaster.getNumPerUnitArray().map(v => [v]);

    //参考書マスターに科目データのない教材の科目名推測
    this._guessSubjects(subjects, namesMatrix, dependencies);//科目補完

    //今月プランの勉強時間の更新ナシ(先月そのまま)

    //月間管理に教材登録
    spIOManager.registerMaterialsMetadataToArchive(
      yearAfterGeshoNum, monthAfterGeshoNum,
      pureIdsMatrix,
      subjects,
      namesMatrix,
      totalStudyTimeMatrix,
      monthlyGoals,
      numPerUnit
    );

  }


  //参考書マスターからでは科目名が取得できない参考書の科目名を推測・補完する。
  static _guessSubjects(subjects, referenceBookNames, dependencies) {
    subjects.forEach((row, index) => {
      const value = row[0];//既に入っている科目名
      if (value && value !== '') return;//すでに科目名が参考書マスターから得られていればスキップ
      const materialName = referenceBookNames[index][0];//参考書名
      row[0] = dependencies.guessSubjectsByName(materialName);//apiで推測
    })
  }




}