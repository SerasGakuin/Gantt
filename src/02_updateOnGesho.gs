/**
 * 月末月初時に行う軽微な更新の内容を定義するクラス。
 * Geshoの最後でこのクラスのstartメソッドが呼ばれる。
 * 不可逆な更新を一回実行するために使用する想定です。毎回実行するようにしたい処理はGeshoのほうに定義してください。
 *
 * 以下のように期限ごとにクラス化して、UpdateOnGesho本体で順に実行するようにするとわかりやすいと思います:
 *
 * class UpdateOnGesho{
 *   static start(spIOManager){
 *     updateOnGesho_2026_04.start(spIOManager)
 *   }
 * }
 *
 * class updateOnGesho_2026_04 {
 *  static start(spIOManager){
 *    if(new Date().getTime() > new Date(2026,4-1,2)) return;//2026_04以降は不要であることをコードでも示す(Date型の仕様的に月は-1する)
 *    ...(処理)
 *  }
 * }
 *
 */

class UpdateOnGesho {
  /**
   * 更新開始用関数。
   * @param {SpeedPlannerIOManager} spIOManager - シートIOの責務を負うオブジェクト
   */
  static start(spIOManager) {
    if (!spIOManager) return;
    updateOnGesho_2026_04.start(spIOManager);
  }
}

class updateOnGesho_2026_04 {
  static start(spIOManager) {
    if (new Date().getTime() > new Date(2026, 4 - 1, 2)) return; //2026_04以降は不要

    this._fixMonthBlocksSpacing(spIOManager); //月間管理シートのデータの正規化

    {
      const thisMonthPlanSheet = spIOManager.thisMonthPlanSheet;
      thisMonthPlanSheet.getRange("J2").setValue("①");
      thisMonthPlanSheet.getRange("J5").setValue("②");
      thisMonthPlanSheet.getRange("J8").setValue("③");
      thisMonthPlanSheet.getRange("J11").setValue("④");
      thisMonthPlanSheet.getRange("J14").setValue("⑤");

      thisMonthPlanSheet
        .getRange("D4:H30")
        .setFontSize(12)
        .setHorizontalAlignment("center");

      thisMonthPlanSheet.getDataRange().setFontFamily("Arial");
      thisMonthPlanSheet.getRange("B4:C30").setWrap(true);
      thisMonthPlanSheet.getRange("I4:I30").setWrap(true);
      thisMonthPlanSheet.getRange("R4:T30").setWrap(true);

      thisMonthPlanSheet.getRange("R2:R3").setFontSize(9);
    }

    {
      const weeklySheet = spIOManager.weeklySheet;
      weeklySheet.getRange("AQ1").clearContent();
      weeklySheet.getRange("AP1").setValue("◆学習進捗（週次・実績）");
      weeklySheet.getRange("B2:D30").setWrap(true);
      weeklySheet.getDataRange().setFontFamily("Arial");
    }

    {
      const monthlyFirst = spIOManager.monthlyFirstSheet;
      monthlyFirst.getRange("AE1:AH1").clear();
      monthlyFirst
        .getRange("D4:H30")
        .setFontSize(12)
        .setHorizontalAlignment("center");

      monthlyFirst.getDataRange().setFontFamily("Arial");
      monthlyFirst.getRange("B4:C30").setWrap(true);
      monthlyFirst.getRange("R4:T30").setWrap(true);

      const formula = `=LET(
    parts, SPLIT($A$2, ","),
    startRow, VALUE(INDEX(parts,1)),
    endRow, VALUE(INDEX(parts,2)),
    
    memo1, "まず該当年月の実績の長方形範囲を切り出す。",
    achievements, INDIRECT("'月間管理'!N"&startRow&":"&"R"&endRow),
    
    memo2, "COLLUMNだと絶対的な列番号が帰るので、上で取得したachievements内の相対列番号に変換して返却している",
    memo3, "そして列ごとに文字列結合して、それが空でない場合にそこがデータのある列だとみなしている。",
    lastCol, MAX(FILTER(
        COLUMN(achievements)-COLUMN(INDEX(achievements,1,1))+1,
        BYCOL(achievements,LAMBDA(col,SUM(N(col<>""))>0))
    )),

   IFNA(lastCol&"週","0週")
)`;

      monthlyFirst.getRange("A1").setFormula(formula);

      const formula1 = `=LET(
    parts, SPLIT($A$2, ","),
      startRow, VALUE(INDEX(parts,1)),
      endRow, VALUE(INDEX(parts,2)),
      
      week, $A$1,

      colText, SWITCH(week,
          "1週","N",
          "2週","O",
          "3週","P",
          "4週","Q",
          "5週","R"
      ),
      
      data, INDIRECT("'月間管理'!"&colText&startRow&":"&colText&endRow),

      achievements, IFNA(data,"データなし"),
      
      {"学習実績の修正"; ""; achievements}
  )`;

      monthlyFirst.getRange("I2").setFormula(formula1);

      monthlyFirst.getRange("I4:I30").clearContent(); //最新の実績の数式の展開に邪魔
    }

    const monthlySheet = spIOManager.monthlySheet;
    {
      // 見出しをセット
      monthlySheet
        .getRange("N1:R1")
        .setValues([
          [
            "実績(第1週)",
            "実績(第2週)",
            "実績(第3週)",
            "実績(第4週)",
            "実績(第5週)",
          ],
        ]);
      monthlySheet.getRange("D1").setValue("印刷要否");
      monthlySheet.getRange("I1").setValue("教材・学習内容\n※メモ入力時は空欄");
      monthlySheet
        .getRange("J1")
        .setValue("月間目標・計画のヒント・メモのタイトル");
      monthlySheet.getRange("G1").setValue("ID");
      monthlySheet.getRange("K1").setValue("単位\n処理量");
      // --- ヘッダの中央揃え（縦横）、太字、フォント統一 ---
      monthlySheet
        .getRange("A1:U1")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setFontWeight("bold");

      monthlySheet.getDataRange().setFontFamily("Arial").setWrap(true);

      //列幅調整
      monthlySheet.setColumnWidth(1, 70);
      monthlySheet.setColumnWidth(2, 32);
      monthlySheet.setColumnWidths(5, 2, 25); // EF列は使ってないので細目に（一応隠すまではしない）
      monthlySheet.setColumnWidth(4, 50); // Dも幅を最適化
      monthlySheet.setColumnWidth(8, 45); // H
      monthlySheet.setColumnWidths(11, 3, 64); // K〜Mに同じ幅を設定
      monthlySheet.setColumnWidths(14, 5, 160); // N〜R（15〜18列目）に同じ幅を設定

      //フォントサイズ調整
      monthlySheet.getDataRange().setFontSize(8);
      monthlySheet.getRange("K2:M").setFontSize(15);

      //折り返し設定
      monthlySheet.getRange("I:J").setWrap(true);
      monthlySheet.getRange("N:R").setWrap(true);
    }

    {
      // C列の月修復
      // 01が勝手に1に直されるのを防ぐために文字列書式に変更（全行）
      monthlySheet.getRange("A:J").setNumberFormat("@");
      //データを正規化
      const lastRow = monthlySheet.getLastRow();
      const monthlySheetMonthCol = monthlySheet
        .getRange(1, 3, lastRow, 1)
        .getDisplayValues();
      // C列の値を書き換えるための配列を作成
      const updated = monthlySheetMonthCol.map((row) => {
        let value = row[0];
        // 空白はそのまま
        if (value === "" || value == null) return row;
        // trim して余分なスペースを削除する対応
        value = value.toString().trim();
        // すでに 01〜12 の2桁 → そのまま
        if (/^(0[1-9]|1[0-2])$/.test(value)) {
          return [value];
        }
        // 1〜9 の **1桁のみ** → 0埋めする
        if (/^[1-9]$/.test(value)) {
          return ["0" + value];
        }
        // それ以外（文字、日付、他形式）はそのまま
        return [value];
      });
      // C列へ書き戻し
      monthlySheet.getRange(1, 3, updated.length, 1).setValues(updated);
    }

    {
      // G列の数式は消す
      const lastRow = monthlySheet.getLastRow();
      const formulas = monthlySheet.getRange(1, 7, lastRow, 1).getFormulas();
      const values = monthlySheet.getRange(1, 7, lastRow, 1).getValues();
      const newValues = formulas.map((row, i) => {
        // 数式がある場合だけ空にし、値は残す
        return [row[0] && row[0].startsWith("=") ? "" : values[i][0]];
      });
      monthlySheet.getRange(1, 7, lastRow, 1).setValues(newValues);
    }

    {
      //下のほうの謎の数式のはいった二行を削除
      SpreadsheetApp.flush();
      const lastRow = monthlySheet.getLastRow();
      const yearCol = monthlySheet.getRange(1, 2, lastRow, 1).getFormulas();
      for (let i = yearCol.length - 1; i >= 0; i--) {
        //下から上へ走査
        const formula = yearCol[i][0]; //通常はここに数式はない。
        if (formula && formula.includes("macro")) {
          //数式がある時点で本来はアウトだが、慎重にポンポイントで削除
          monthlySheet.deleteRow(i + 1);
        }
      }
    }
  }

  //この関数はデータの不可逆な正規化をしているだけなのでいずれ不要になる。
  static _fixMonthBlocksSpacing(spIOManager) {
    const sheet = spIOManager.monthlySheet;
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return;

    // B列・C列を取得（1行目はヘッダー想定なので2行目から）
    const ym = sheet.getRange(2, 2, lastRow - 1, 2).getValues();

    // 「空行かどうか」の配列に変換
    // true = データあり（空でない）
    // false = 空行
    const isData = ym.map(([y, m]) => {
      const yEmpty =
        (typeof y === "string" && y.trim() === "") || y === "" || y == null;
      const mEmpty =
        (typeof m === "string" && m.trim() === "") || m === "" || m == null;
      return !(yEmpty && mEmpty); // 両方空なら空行とみなす
    });

    // データブロック（連続した true の区間）を抽出
    const blocks = [];
    let inside = false;
    let start = 0;

    for (let i = 0; i < isData.length; i++) {
      if (isData[i] && !inside) {
        inside = true;
        start = i;
      }
      if (!isData[i] && inside) {
        inside = false;
        blocks.push({ start, end: i - 1 });
      }
    }
    if (inside) blocks.push({ start, end: isData.length - 1 });

    // ブロックがない or 1つだけなら空行調整不要
    if (blocks.length < 2) return;

    // 下のブロックから順にチェック（行ずれに強い）
    for (let i = blocks.length - 1; i > 0; i--) {
      const prev = blocks[i - 1];
      const curr = blocks[i];

      // シート上の行番号に変換 (+2 = シートの2行目が ym[0])
      const prevEndRow = prev.end + 2;
      const currStartRow = curr.start + 2;

      // gap = ブロック間の空行数
      const gap = currStartRow - prevEndRow - 1;

      // 空行が1行しかないなら → 1行だけ挿入
      if (gap === 1) {
        sheet.insertRowsBefore(currStartRow, 1);

        // 以降のブロックのシート行番号を補正
        for (let j = i; j < blocks.length; j++) {
          blocks[j].start++;
          blocks[j].end++;
        }
      }
    }
  }
}
