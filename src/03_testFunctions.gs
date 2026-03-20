//03_testFunctions.gs テスト関数置き場
/*
test_GenWeeklyPlan_prod_like_generate() : 実際に一つの教材について本番同様のgenerateメソッドを用いて生成し、結果を確認するテスト関数。

test_GenWeeklyPlan_generate_withAI_withRefBookInfo() : 参考書マスターデータありでのAI生成実験用関数

test_GenWeeklyPlan_generate_withAI_withoutRefBookInfo() : 参考書マスターデータなしでのAI生成実験用関数
*/


//======================================================================================================

/**
 * 実際に一つの教材について生成し、結果を確認するテスト関数。
 * 計画作成のアルゴリズムを変更したときに使用してみてください。
 * 
 * いろいろなパターンがカバーできるので、以下のパターンについて最低でもテストするのがおすすめです。
 * 
 * LEAP gET003、青チャート gMB017、入門英文法の核心 gEB008、やさしい高校数学 gMB023、はじめの英文読解ドリル gEK001、
 * Focus Gold Smart 数C2 gMB078 (数C：例題C1.25~C.29)、
 * （任意の）過去問演習、参考書マスターに情報の登録されていない適当な教材
 * 
 * 一つの章の中で完結するケース、章をまたぐケース、最後の問題に到達するのでキャップするケース、すでに最新の実績が最後の問題に到達しているケース、目安処理量が1のケース（n~nみたいにならないかどうか）
 * @private
 */
function test_GenWeeklyPlan_prod_like_generate() {
  //必要なデータのモックデータ。このオブジェクトの中身だけ書き換えればテストできる。
  const mockInput = {
    materialId: 'gMB017',               //参考書id
    lastAchievement: '例題101~120',    //最新の学習実績
    valueToAdd: 25                     //目安処理量
  }
  GenWeeklyPlan_handlers_flags.isAIBanned = true;    //AIを使う生成を禁止してテストする場合はこれをtrueに。
  GenWeeklyPlan_handlers_flags.pushLogAboutWhatHandlerGenerated = true;   //どのハンドラーが生成したのか見たい場合はこれをtrueに
  /*
  isAIBannedがtrueの場合は、本番であればAIに処理が委託されるケースではlogに「★AI生成キャンセル」というテキストが出力されます。
  AIのapiは使用料金がかかるので、テスト時に浪費したくない場合に使ってください。
  */

  const dependencies = getGenWeeklyPlanProdDependencies();//本番用の依存関係
  //参考書マスターから情報を検索。
  const materialInfo = dependencies.getMaterialsInfoByIds([mockInput.materialId])[mockInput.materialId];

  console.time("test_GenWeeklyPlan_prod_like_generate");
  //生成用コンテキスト生成
  const context = GenWeeklyPlan_contextFactory.create(dependencies,
    mockInput.materialId, materialInfo, mockInput.lastAchievement, mockInput.valueToAdd
  );
  Logger.log(`context:\n`, JSON.stringify(context))

  Logger.log('=== 単一教材の計画生成モックテスト結果 ===\n');

  //実際に生成。
  const result = GenWeeklyPlan_handlers_generator.generate(context);

  console.timeEnd("test_GenWeeklyPlan_prod_like_generate");
  //結果をlogにだす。
  Logger.log(`入力:\n${JSON.stringify(mockInput)}\n`);
  Logger.log(`参考書マスターから得た教材データ\n${JSON.stringify(materialInfo)}\n`);
  Logger.log(`出力:\n${result}\n`);
  Logger.log('=== 単一教材の計画生成モックテスト結果終わり ===\n');
}

//======================================================================================================

/**
 * 参考書マスターデータありでのAI生成実験用関数
 * @private
 */
function test_GenWeeklyPlan_generate_withAI_withRefBookInfo() {
  //必要なデータのモックデータ。このオブジェクトの中身だけ書き換えればテストできる。
  const mockInput = {
    materialId: 'gET002',               //参考書id
    lastAchievement: 'No.100~900',    //最新の学習実績
    valueToAdd: 150                     //目安処理量
  }
  const dependencies = getGenWeeklyPlanProdDependencies();//本番用の依存関係
  //参考書マスターから情報を検索。
  const materialInfo = dependencies.getMaterialsInfoByIds([mockInput.materialId])[mockInput.materialId];

  const context = GenWeeklyPlan_contextFactory.create(dependencies,
    mockInput.materialId, materialInfo, mockInput.lastAchievement, mockInput.valueToAdd
  );
  Logger.log(`context:\n`, context)

  Logger.log('=== 単一教材の計画生成モックテスト結果 ===\n');

  //実際に生成。
  const result = GenWeeklyPlan_handlers_withAI.askingAIWithBookInfoHandler.run(context);

  //結果をlogにだす。
  Logger.log(`入力:\n${JSON.stringify(mockInput)}\n`);
  Logger.log(`出力:\n${result}\n`);
  Logger.log('=== 単一教材の計画生成モックテスト結果終わり ===\n');
}

//======================================================================================================

/**
 * 参考書マスターデータなしでのAI生成実験用関数
 * @private
 */
function test_GenWeeklyPlan_generate_withAI_withoutRefBookInfo() {
  //必要なデータのモックデータ。このオブジェクトの中身だけ書き換えればテストできる。
  const mockInput = {
    materialId: 'gCH005',               //参考書id
    lastAchievement: '9章　酸化還元・電池電気分解',    //最新の学習実績
    valueToAdd: 3                     //目安処理量
  }

  const dependencies = getGenWeeklyPlanProdDependencies();//本番用の依存関係
  const context = GenWeeklyPlan_contextFactory.create(dependencies,
    mockInput.materialId, undefined, mockInput.lastAchievement, mockInput.valueToAdd
  );
  Logger.log(`context:\n`, context)

  Logger.log('=== 単一教材の計画生成モックテスト結果 ===\n');

  //実際に生成。
  const result = GenWeeklyPlan_handlers_withAI.askingAIWithoutBookInfoHandler.run(context);

  //結果をlogにだす。
  Logger.log(`入力:\n${JSON.stringify(mockInput)}\n`);
  Logger.log(`出力:\n${result}\n`);
  Logger.log('=== 単一教材の計画生成モックテスト結果終わり ===\n');
}

