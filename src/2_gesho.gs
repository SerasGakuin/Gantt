function arrayFromDict(data, col) {
    return Object.keys(data).map(key => {
      if (data[key] === null) return '';
      return data[key][col] || '';
    })
  }

function testGetInfoFromReferenceBookMaster(){

  const schemeSheet = SpreadsheetApp.getActive().getSheetByName("月間実績")
  const schemeSheetLast = schemeSheet.getRange("A:A").getValues().filter(String).length

  // 「月間実績」からデータを取得
  const scheme = schemeSheet.getRange(5,1,schemeSheetLast - 1,8).getDisplayValues()
  console.log(scheme)
  const schemevertical = transpose(scheme)
  const scheme1 = transpose([schemevertical[0]]) // 参考書ID

  var referenceBookIds = transpose(scheme1)[0]
  console.log(referenceBookIds)

  var infoFromMaster = getInfoFromReferenceBookMaster(referenceBookIds=referenceBookIds, cols=["月間目標", "単位当たり処理量"])
  var monthlyGoal = arrayFromDict(infoFromMaster, "月間目標");
  var numPerUnit  = arrayFromDict(infoFromMaster, "単位当たり処理量");
  
  monthlyGoal = transpose([monthlyGoal]);
  numPerUnit = transpose([numPerUnit]);

  console.log(monthlyGoal)
  console.log(numPerUnit)

  //const monthlyGoal = infoFromMaster["月間目標"];
  //const numPerUnit = infoFromMaster["単位あたり処理量"]

  //console.log(monthlyGoal)
  //console.log(numPerUnit)
}

function getInfoFromReferenceBookMaster(referenceBookIds, cols=["月間目標", "単位当たり処理量"]) {
  // 「参考書マスター」スプレッドシート
  var referenceSheet = SpreadsheetApp.openById("1Z0mMUUchd9BT6r5dB6krHjPWETxOJo7pJuf2VrQ_Pvs").getSheetByName("参考書マスター");

  // 「参考書マスター」シートから、該当の参考書の情報を取得する
  var referenceBookAllData = referenceSheet.getRange("A1:M" + referenceSheet.getLastRow()).getValues();

  // ヘッダー行から列名と列インデックスのマッピングを作成
  var columnMapping = {};
  referenceBookAllData[0].forEach((header, index) => {
    columnMapping[header] = index;
  });

  // 結果を格納するオブジェクト
  var result = {};

  // 各参考書IDについて処理
  referenceBookIds.forEach(function(id) {
    // データを検索
    var bookData = referenceBookAllData.find(function(row) {
      return row[columnMapping["参考書ID"]] === id;
    });

    if (bookData) {//bookDataが見つかっていれば
      result[id] = {};
      // 指定された列の情報を取得
      cols.forEach(function(col) {
        if (col in columnMapping) {
          result[id][col] = bookData[columnMapping[col]];
        }
      });
    } else {
      // 参考書が見つからない場合
      result[id] = null;
    }
  });

  return result;
}

//const transpose = a => a[0].map((_, c) => a.map(r => r[c]));

const limitOfRowsInMonthPlan = 27;//今月プランに張り付けする行ｎ数の限界。週間管理や月初の教材の限界数でもある。
const monthGetudobunPosition = "C1"//今月プランや月初の左上の方の「月度分」テキストの位置
const monthGetudoDaiWeekSyubunPosition = "S1"//今月プランの左上の方の「ｎ月度分第m週分」テキストの位置

function gesho() {
  /**
   * この関数はガントチャートがある生徒のスピードプランナーに関する月末月初の作業をしています。
   * 月末月初とは以下の一連の作業のことを指します
   　　　*シート「週間管理」からシート「月初」のBK4:BO19に先月の学習実績をコピペ
   　　　*シート「週間管理」からシート「月間管理」の該当月、該当参考書にあたる箇所にコピペ
   　　　*シート「週間管理」の前月の実績を削除
   　　　*先月の週間計画をシート「今月実績」にコピー
   　　　*シート「ガントチャート」から今月分の予定をシート「月間実績」に表示して、さらに、それを月間管理に追加。そしてさらにそれを週間管理に追加。
   */


  // 週間管理シートの情報を取得
  var weeklySheet = SpreadsheetApp.getActive().getSheetByName("週間管理")
 
  // 月間管理シートの情報を取得
  var monthlySheet = SpreadsheetApp.getActive().getSheetByName("月間管理");
  var monthlyValues = monthlySheet.getRange("A:A").getValues(); // 月間管理のA列全部取得
  var monthlyfirst = SpreadsheetApp.getActive().getSheetByName("月初");
  var monthAchievement = SpreadsheetApp.getActive().getSheetByName("今月実績")
  // idを判別

  const collumnAR = weeklySheet.getRange("AR1").getColumn();  // 週間管理の5週間の実績を格納しているところの最初の列の列番号

  const achievementOnWeeklySheetRightEdge = weeklySheet.getRange(4,collumnAR, limitOfRowsInMonthPlan, 5).getDisplayValues();//週間管理シートの末尾にまとまっている各週の学習実績取得　5週間分

  var originrow = monthlySheet.getRange("A:A").getDisplayValues()//月間管理シートに格納されている一意な「年+月+ID」を配列として一気に取得
  var number = weeklySheet.getRange("A4").getValue()//「マクロ動作前の今月」の教材1個目の年+月+ID

  var achievementRowCount = achievementOnWeeklySheetRightEdge.length;//週間管理シート右端の実績の行数と列数
  var achievementCollumnCount = achievementOnWeeklySheetRightEdge[0].length;
  const collumnBJ = monthlyfirst.getRange("BJ1").getColumn();//BK列の列番号　月初の実績貼り付け先の左端
  monthlyfirst.getRange(4,collumnBJ,achievementRowCount,achievementCollumnCount)
              .setValues(achievementOnWeeklySheetRightEdge)

  var materialToRecordAchievementFound = false
  // 週間管理シートの各行について処理
    //月間管理の上の行から順番に走査して、教材に振られたIDをもとに週間管理の実績を張るべき行の開始位置を特定している。
    //その位置を発見したら、そこに週間管理の右端由来の学習実績を張り付ける。
  for (var i = 0; i < monthlyValues.length; i++) {
    if (originrow[i]==number) {//条件;現在の行の一意な教材IDが、今回実績を格納すべき教材のIDと一致するなら
      monthlySheet.getRange(i + 1 ,14,achievementRowCount,achievementCollumnCount)
                                              .setValues(achievementOnWeeklySheetRightEdge)
      materialToRecordAchievementFound = true;//教材が見つかったらマーク
      break;
    }
  }
  //上の処理で実績の貼り付け位置を見つけられなかったらエラーをだす
  if(!materialToRecordAchievementFound){
    Logger.log("エラー！週間管理から学習実績をコピーすべき教材のデータ(一意なID)が、月間管理シートに見つかませんでした。 at gesho() at _gesho.js")
  }
  
  //週間管理シートの内容をクリア
  const collumnH = weeklySheet.getRange("H1").getColumn();
  const collumnP = weeklySheet.getRange("P1").getColumn();
  const collumnX = weeklySheet.getRange("X1").getColumn();
  const collumnAF = weeklySheet.getRange("AF1").getColumn();
  const collumnAN = weeklySheet.getRange("AN1").getColumn();
  weeklySheet.getRange(4,collumnH,achievementRowCount,1).clearContent()//各週ごとに一列ずつ
  weeklySheet.getRange(4, collumnP,achievementRowCount,1).clearContent()
  weeklySheet.getRange(4, collumnX,achievementRowCount,1).clearContent()
  weeklySheet.getRange(4,collumnAF,achievementRowCount,1).clearContent()
  weeklySheet.getRange(4,collumnAN,achievementRowCount,1).clearContent()
 
  var monthPlan = SpreadsheetApp.getActive().getSheetByName("今月プラン")

  let lastmonth = monthPlan.getRange(4,1,limitOfRowsInMonthPlan,8).getDisplayValues()//A列からH列まで
  monthAchievement.getRange(monthGetudobunPosition).setValue(monthPlan.getRange(monthGetudobunPosition).getDisplayValue())//先月の月のテキスト
  let lastarry = transpose(lastmonth)
  const lastID = transpose([lastarry[0]])
  const lastarryLastIndex = lastarry.length-1
  
  const lasttime = transpose(
    [
    lastarry[lastarryLastIndex-4],
    lastarry[lastarryLastIndex-3],
    lastarry[lastarryLastIndex-2],
    lastarry[lastarryLastIndex-1],
    lastarry[lastarryLastIndex]
    ]
    )

  monthAchievement.getRange(4,1,lastID.length,1).setValues(lastID)//先月の教材を今月プランから今月実績に貼りつけ。A4起点。
  monthAchievement.getRange(4,4,lasttime.length,5).setValues(lasttime)//勉強時間貼りつけ。D4起点。
 
  // 月間実績を”ガントチャート”から”月間実績”へ表示する
  //macroシート取得
  const macroSheet = SpreadsheetApp.getActive().getSheetByName("macro");
  //来月の日付データ取得
  const yearAndMonthArray  = macroSheet.getRange(12, 2, 1, 2).getDisplayValues(); // B12, C12セル
  const year = yearAndMonthArray[0][0];
  const month = yearAndMonthArray[0][1];
  
  displayMonthlyDataFromGantt("20" + year, month); 
  //function.gsに定義されている関数を呼び出しています！
  //ガントチャートの年数と月間実績の年数が不一致の場合、月間実績にからの表が代入される


  //表示するために日付テキストを整形
  const montext = month + "月度分"
  const monformula = '="' + month + '"&"月度 第"&A1&"分"';

  //今月プランに整形した日付テキストを代入
  monthPlan.getRange(monthGetudobunPosition).setValue(montext)//新しい月の名前を今月プランのC1にセット。
  monthPlan.getRange(monthGetudoDaiWeekSyubunPosition).setValue(monformula)//n月度第m週分というテキストを今月プランにセット。

  //先月(たったいま終えた月)の年と月を計算する。まずmonth（来月の月）から単純に-1するが、1月の場合に先月が0月になってしまうので、0になった場合は12にセットしなおす。
  var lastMonth = -1;//仮代入しないとエラー出るのでとりあえず-1を代入
  var theYearAsOfLastMonth = -1;//同様
  if (month == 1){//一月の時は単純に先月は-1で求められないので。
    lastMonth = 12;
    theYearAsOfLastMonth = year -1;
  }else{//その他は規則的な計算
    lastMonth = month - 1;
    theYearAsOfLastMonth = year;
  }


  //計算した先月の月をつかって、表示テキストを組む
  const lastMonthText = lastMonth + "月度分"
  //表示テキストを代入。
  monthlyfirst.getRange(monthGetudobunPosition).setValue(lastMonthText)//月初
  monthAchievement.getRange(monthGetudobunPosition).setValue(lastMonthText)//月間実績

  //今月プランのA1を月初のA1にコピー
  //monthlyfirst.getRange("A1").setValue(monthPlan.getRange("A1").getValue());

  //月間管理シートを参照する変数monthlySheetをつかって、月間管理シートの特定範囲から値を取得する。
  //範囲は、1行2列目(B1)から、下に1000行、右に2行。結果的に1000×2マス。ここには各教材の使用年と月が記録されている。
  //つまりdayrowには年も月も両方格納される。この値は、範囲内の値をすべて含む二次元配列の形で得られる。
  const dayrow = monthlySheet.getRange(1,2,1000,2).getDisplayValues()

  //月間実績（ガントチャートから教材ごとの目標勉強時間を取得しまとめる場所）のシートを取得
  const schemeSheet = SpreadsheetApp.getActive().getSheetByName("月間実績")
  //月間実績シートからA列全体のうち、文字列を含む範囲の長さを取得
  const schemeSheetLast = schemeSheet.getRange("A:A").getValues().filter(String).length  
  Logger.log('schemeSheetLast: ' + schemeSheetLast + "\nこの変数は、月間実績のA列の値のある行の数。教材の数は、これより1少ない。");

  // 「月間実績」からデータを取得（参考書データ、それぞれの週ごとの勉強時間、累計勉強時間。）
  const scheme = schemeSheet.getRange(5,1,schemeSheetLast-1,8).getDisplayValues()//schemeSheetLast-1←この-1は必須！
  //この後各列を別々の配列に分離したいが、列が二つ目の次元の番号であり、このまま列ごとに分離することは難しいので、「transpose」で行と列を入れ替える。
  const schemevertical = transpose(scheme)
  
  const scheme1 = transpose([schemevertical[0]]) // 参考書ID
  const scheme2 = transpose([schemevertical[1]]) // 参考書名
  const scheme3 = transpose([schemevertical[2],  // 第1週~第5週の勉強時間
                             schemevertical[3],
                             schemevertical[4],
                             schemevertical[5],
                             schemevertical[6]]) 
  const scheme4 = transpose([schemevertical[7]]) // 合計時間

  // IDから「参考書マスター」を検索して、「月間目標」と「単位当たり処理量」を取得する
  var referenceBookIds = transpose(scheme1)[0]
  var infoFromMaster = getInfoFromReferenceBookMaster(referenceBookIds=referenceBookIds, cols=["月間目標", "単位当たり処理量"])
  var monthlyGoal = arrayFromDict(infoFromMaster, "月間目標");
  var numPerUnit  = arrayFromDict(infoFromMaster, "単位当たり処理量");
  
  monthlyGoal = transpose([monthlyGoal]);
  numPerUnit = transpose([numPerUnit]);
  
  
  // データを順番に「今月プラン」に貼り付けていく。dayrowは月間管理シート由来の日付データ。
  for (var i = 0; i < dayrow.length - 1; i++){

    var yearrow1 = dayrow[i][0]
    var monthrow1 = dayrow[i][1]
    var yearrow2 = dayrow[i + 1][0]
    var monthrow2 = dayrow[i + 1][1]

    //上から順番に教材の情報を調べてゆき、先月の教材と今月の教材の境界を発見する。発見したら、その位置を上端としてデータを一気にはりつけて、for全体からbreakで抜け出している。
    if (yearrow1 == theYearAsOfLastMonth && monthrow1 == lastMonth && yearrow2 != year && monthrow2 != lastMonth) {
       
       //月間管理シートに貼りつけ
       monthlySheet.getRange(i + 4,7,scheme1.length, 1).setValues(scheme1) // G列に参考書ID
       monthlySheet.getRange(i + 4,9,scheme2.length,1).setValues(scheme2)  // I列に参考書名
       monthlySheet.getRange(i + 4,12,scheme4.length,1).setValues(scheme4) // L列に合計時間
       monthlySheet.getRange(i + 4,10,monthlyGoal.length,1).setValues(monthlyGoal) // J列に月間目標
       monthlySheet.getRange(i + 4,11,numPerUnit.length,1).setValues(numPerUnit) // K列に単位あたり処理量
       monthlySheet.getRange(i + 4,2,year.length,1).setValue(year)      // B列に年
       monthlySheet.getRange(i + 4,3,month.length,1).setValue(month)     // C列に月

      //今月プランにデータまたはエラーメッセージを貼りつけ
      if ( scheme1.length <= limitOfRowsInMonthPlan) {//条件;今月プランの限界行数を突破しなければ
            //　今月プランの勉強時間をクリア
            monthPlan.getRange(4,4,limitOfRowsInMonthPlan,5).clearContent()
            // 「今月プラン」に第1週から第5週までの勉強時間を貼り付け
            monthPlan.getRange(4,4,scheme3.length,5).setValues(scheme3)  
            //　今月プランのA列をクリア
            monthPlan.getRange(4,1,limitOfRowsInMonthPlan,1).clearContent()
            // A列の「年月＋ID」を、「今月プラン」に貼り付け
            var copytoscheme1 = monthlySheet.getRange(i + 4,1,schemeSheetLast-1,1).getValues() 
            monthPlan.getRange(4,1,copytoscheme1.length).setValues(copytoscheme1)
      }else{
            Logger.log("エラー！行数が今月プランの限界を超えています! at gesho() at _gesho.gs");
            //　今月プランの勉強時間をクリア
            monthPlan.getRange(4,4,limitOfRowsInMonthPlan,5).clearContent()
            //　今月プランのA列をクリア
            monthPlan.getRange(4,1,limitOfRowsInMonthPlan,1).clearContent()
            //  エラーメッセージ代入
            monthPlan.getRange(4,4).setValue("行数が今月プランの限界を超えています!");  
      }
      break
          
      }
      
    }//forループ終わり


    //週間管理シートの1行の週の番号の横の開始日の日付を、ガントチャートをもとに入力する。
    //monthに今月の月の数が格納されている。
    var ganttChartSheet = SpreadsheetApp.getActive().getSheetByName("ガントチャート");
    const ganttChartThirdRow = ganttChartSheet.getRange("3:3").getDisplayValues()[0];//ガントチャートの3行目の値たち(表示内容)　getDisplayValuesはかならず二次元配列を返すので、[0]をつけないと、行全体の一次元配列にならない。

    //forループでガントチャートの今月の週の日付のデータ位置を特定します。
    var startCollumnOfThisMonth = -1;//コンパイルエラー回避用初期値
    for(let i = 0 ;i< ganttChartThirdRow.length;i++){
      //monthは、一桁月のとき、十の位に0が代入されているが、ガントチャートのほうの月は十の位の0がないので変換している
      if(ganttChartThirdRow[i].includes( String(Number(month)) + "月" )){//該当列を発見したら記録してbreak
        startCollumnOfThisMonth=i+1;//i+1にしないと、列番号が1始まりなのでバグる　つまり、1列目が配列の0番に格納されている。
        break;
      }
    }

    if(startCollumnOfThisMonth==-1){//月が発見できなければエラー出して終了

      Logger.log("エラー！ガントチャートに来月のデータが発見できません。週間管理の日付自動入力をキャンセルします。 at gesho at _gesho.js")

    }else{//月が発見されていれば

    const thisMonthEachWeekFirstDates = ganttChartSheet.getRange(2,startCollumnOfThisMonth,1,5).getDisplayValues()[0];//今月の週の開始日たち ここも
    Logger.log("来月第一週開始日; "+thisMonthEachWeekFirstDates[0])
    weeklySheet.getRange("D1").setValue(thisMonthEachWeekFirstDates[0])//日付貼りつけ

    }





}
//function gesho終わり



// ガントチャートなしの場合の月末月初
function forexcel(){ //gesho_wo_gantt
  const transpose = a=> a[0].map((_, c) => a.map(r => r[c]));
  // 週間管理シートの情報を取得
  var weeklySheet = SpreadsheetApp.getActive().getSheetByName("週間管理")
  
  // 月間管理シートの情報を取得
  var monthlySheet = SpreadsheetApp.getActive().getSheetByName("月間管理");
  var monthlyValues = monthlySheet.getRange("A:A").getValues(); // シート内の適切な範囲を指定
  var monthlyfirst = SpreadsheetApp.getActive().getSheetByName("月初");

  const value = weeklySheet.getRange("AR4:AV19").getDisplayValues()
  var originrow = monthlySheet.getRange("A:A").getDisplayValues()
  var number = weeklySheet.getRange("A4").getValue()
  monthlyfirst.getRange("BK4:BO19").setValues(value)

  // 週間管理シートの各行について処理
  for (var i = 0; i < monthlyValues.length; i++) {
    var row = originrow[i] 
    
    if (number == row) {
      monthlySheet.getRange(i + 1 ,14,16,5).setValues(value)
      break;
    }
  }
  
  weeklySheet.getRange("H4:H19").clearContent()
  weeklySheet.getRange("P4:P19").clearContent()
  weeklySheet.getRange("X4:X19").clearContent()
  weeklySheet.getRange("AF4:AF19").clearContent()
  weeklySheet.getRange("AN4:AN19").clearContent()
  
  var monthplan = SpreadsheetApp.getActive().getSheetByName("今月プラン")
  var monthachievement = SpreadsheetApp.getActive().getSheetByName("今月実績")

  let lastmonth = monthplan.getRange("A4:J19").getDisplayValues()
  monthachievement.getRange("D1").setValue(monthplan.getRange("D1").getDisplayValue())
  let lastarry = transpose(lastmonth)
  const lastID = transpose([lastarry[0]])
  const lasttime = transpose([lastarry[5],lastarry[6],lastarry[7],lastarry[8],lastarry[9]])

  monthachievement.getRange("A4:A19").setValues(lastID)
  monthachievement.getRange("E4:I19").setValues(lasttime)
  monthplan.getRange("A4:A19").clearContent()
  monthplan.getRange("F4:J19").clearContent()

  
  const macroSheet = SpreadsheetApp.getActive().getSheetByName("macro");
  const yearAndMonthArray  = macroSheet.getRange(12, 2, 1, 2).getDisplayValues(); // B12, C12セル
  const year = yearAndMonthArray[0][0];
  const month = yearAndMonthArray[0][1];
  const montext = month + "月度分"
  const monformula = '="' + month + '"&"月度 第"&A1&"分"'
  monthplan.getRange("D1").setValue(montext)
  monthplan.getRange("V1").setValue(monformula)

  const dayrow = monthlySheet.getRange(1,2,1000,2).getDisplayValues()

  for (var i = 0; i< dayrow.length - 1; i++){    
    var yearrow1 = dayrow[i][0]
    var monthrow1 = dayrow[i][1]
    var yearrow2 = dayrow[i + 1][0]
    var monthrow2 = dayrow[i + 1][1]
    if (month != 1) {
      if (yearrow1 == year && monthrow1 == month - 1 && yearrow2 != year && monthrow2 != month - 1) {
       monthlySheet.getRange(i + 4,2,2,1).setValue(year)
       monthlySheet.getRange(i + 4,3,2,1).setValue(month) 
       break
      }
    }else{
      if (yearrow1 == year - 1 && monthrow1 == 12 && yearrow2 != year && monthrow2 != 12) {
       monthlySheet.getRange(i + 4,2,2,1).setValue(year)
       monthlySheet.getRange(i + 4,3,2,1).setValue(month) 
       break
      }
    }
  }


}
//function forexcel終わり




function magokoro () {//シートの整形をする関数
  
  var monthlyPlanUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("今月プラン");
  var monthlyPlanSheet = SpreadsheetApp.getActive().getSheetByName("今月プラン")
  monthlyPlanSheet.getRange("B4:D19").setFormulas(monthlyPlanUekata.getRange("B4:D19").getFormulas());
  monthlyPlanSheet.hideColumn(monthlyPlanSheet.getRange("A1"))
  monthlyPlanSheet.setRowHeights(4,16,32)
  monthlyPlanSheet.setColumnWidth(2,40)
  monthlyPlanSheet.setColumnWidths(3,2,224)
  monthlyPlanSheet.setColumnWidths(5,7,49)
  monthlyPlanSheet.setColumnWidth(12,230)
  monthlyPlanSheet.setColumnWidth(15,174)
  monthlyPlanSheet.setColumnWidths(18,2,204)
  monthlyPlanSheet.setColumnWidths(21,2,224)
  monthlyPlanSheet.setColumnWidths(24,10,33)
  monthlyPlanSheet.setColumnWidth(34,168)
  monthlyPlanSheet.setColumnWidth(35,380)
  monthlyPlanSheet.setColumnWidth(39,380)
  monthlyPlanSheet.setColumnWidths(36,2,60)
  monthlyPlanSheet.setColumnWidths(40,2,60)

  // 曜日の表示フォーマットを"ddd"に統一
  monthlyPlanSheet.getRange("X3:AG3").setFormulas(monthlyPlanUekata.getRange("X3:AG3").getFormulas());

  var monthlyFirstUekata = SpreadsheetApp.openById("10KzulodqrBj5EIGLIhr6n0hjZvMPLKI3N0mbHmrauoQ").getSheetByName("月初");
  var monthlyFirstSheet = SpreadsheetApp.getActive().getSheetByName("月初");
  monthlyFirstSheet.getRange("B4:D19").setFormulas(monthlyFirstUekata.getRange("B4:D19").getFormulas());
    
  // 行や列のサイズを変更
  monthlyFirstSheet.hideColumn(monthlyFirstSheet.getRange("A1"))
  monthlyFirstSheet.setRowHeights(4,16,32)
  monthlyFirstSheet.setColumnWidth(2,40)
  monthlyFirstSheet.setColumnWidths(3,2,224)
  monthlyFirstSheet.setColumnWidths(5,7,49)
  monthlyFirstSheet.setColumnWidth(12,230)
  monthlyFirstSheet.setColumnWidth(15,174)
  monthlyFirstSheet.setColumnWidths(18,2,204)
  monthlyFirstSheet.setColumnWidth(20,40)
  monthlyFirstSheet.setColumnWidths(21,2,224)
  monthlyFirstSheet.setColumnWidths(23,7,49)
  monthlyFirstSheet.setColumnWidth(30,230)
  monthlyFirstSheet.setColumnWidths(32,7,140)
  monthlyFirstSheet.setColumnWidths(41,2,224)
  monthlyFirstSheet.setColumnWidths(43,10,33)
  monthlyFirstSheet.setColumnWidth(54,168)
  monthlyFirstSheet.setColumnWidth(55,380)
  monthlyFirstSheet.setColumnWidth(59,380)
  monthlyFirstSheet.setColumnWidths(56,2,60)
  monthlyFirstSheet.setColumnWidths(60,2,60)

  // 曜日の表示フォーマットを"ddd"に統一
  monthlyFirstSheet.getRange("AR3:BA3").setFormulas(monthlyFirstUekata.getRange("AR3:BA3").getFormulas());

  var hearingSheet = SpreadsheetApp.getActive().getSheetByName("ヒアリング")
  hearingSheet.setRowHeights(1, 1, 120);
  hearingSheet.getRange("A1").setValue("生徒に関する情報を入れてください。→");
  hearingSheet.getRange("A1").setHorizontalAlignment("right");
  hearingSheet.getRange("B2").setValue("ChatGPTからの提案")
  hearingSheet.getRange("C1").setFormula("=displayKakomon()");
  hearingSheet.getRange("B3:B16").clearContent();
}

