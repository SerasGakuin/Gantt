// 00_FunctionsConfigs.gs
/*
  サイドバーからの関数の実行をはじめ、functionIdを使ってこのライブラリのグローバル関数を実行するためのインターフェースの設定ファイルです。
  GAS側のfunctionIdの自動設定や、引数がない関数についてはhtml側の呼び出し用functionIdとボタンの自動生成をするなどに参照されています。

  onOpenや、onOpenAction、abstractFunctionなど、interface用の関数は登録しないでください。再帰的に実行され続ける恐れがあるためです。
  （ただし実行が異常な回数に達した場合、実行時に検出され、エラーを出します）
  
  このファイルで扱う設定:

  FunctionsConfigs: 引数なし関数の設定配列
  ContextFunctionsConfigs: contextを引数として受け取る関数の配列

  両者をまたいで、id, func, labelの重複は許されません。データの整合性は、実行時にこのファイルの末尾の仕組みによって自動で検証されます。
  このとき設定オブジェクトを凍結しているので注意してください。
 */
//========================================== FunctionsConfigs ==========================================
/**
 * FunctionsConfigs: 引数なし関数の設定配列
 * 
 * 格納順が、サイドバーメニューでの表示順になります。GAS、html双方でfunctionIdと実行用ボタンと説明がつじつまを合わせながら自動で生成されます。
 * contextを無視する（受け取らない）関数については、こちらに登録するだけで充分です。こちらのほうが多いと思います。
 * また、サイドバーでのボタンの順序はここでの配列の順序と同じになります。
 * 
 * @typedef {Object} FunctionConfig
 * @property {string} id - 必須。識別子。重複があるとエラー。関数名にすることをお勧めします。functionから動的に生成せず明示的に設定することで動作の安定させている。
 * @property {Function} func - 必須。実行関数。（引数を受け取らない関数のみ）
 * @property {string} [label] - 任意。表示名設定を推奨します。(もしも未設定の場合は、idと同じになります。)
 * @property {string} [desc] - 任意。詳細説明。
 */
/** @type {FunctionConfig[]} */
const FunctionsConfigs = [
  {
    id: 'generateHearingEntries', func: generateHearingEntries, label: "ヒアリング項目作成",
  },
  {
    id: 'updateMonthlyRecord', func: updateMonthlyRecord, label: "月間実績に計画を表示",
    desc: "ガントチャート入力に従って、「月間実績」の内容を更新します。年月は「月間実績」シートのC1、C2セルを参照します。"
  },
  {
    id: 'repairFormats', func: repairFormats, label: "数式修復・フォーマット修復",
    desc: "各種シートの数式などのフォーマットを修復するマクロです。"
  },
  {
    id: 'changeMaterials', func: changeMaterials, label: "使用教材途中変更",
    desc: "ガントチャート入力に従って月の途中に使用教材のラインナップを変更するマクロです。"
  },

  {
    id: 'geshoWithGantt', func: geshoWithGantt, label: "gesho",
    desc: "「ガントチャート」シートを使用する生徒の月末月初処理です。"
  },
  {
    id: 'geshoWithoutGantt', func: geshoWithoutGantt, label: "forexcel",
    desc: "「ガントチャート」シートを使用しない生徒の月末月初処理です。"
  }


];


//================================================ ContextFunctionsConfigs =================================================
/**
 * ContextFunctionsConfigs: contextを引数として受け取る関数の配列
 * 
 * 予期せぬ動作の回避のため、オブジェクトやクラスの中の変数は避けて、グローバル関数を格納してください。
 * 呼び出し用の関数がない場合は、00_contextFunctions.gsにcontextを受け取って本体の関数にわたす、関数を定義してください。
 * 
 * 自動で設定されるもの:
 * GAS側のこの配列内の関数を実行するためのfunctionIdは、自動でidプロパティになります。
 * 
 * 手動で設定するもの:
 * htmlに、関数を実行するボタンと、GASのabstractFunctionに渡す{functionId, context}オブジェクトの内容。
 * htmlでのfunctionIdは、こちらに登録しているidプロパティにしてください。(GAS側と同じにする必要があるので)
 * 
 * @typedef {Object} ContextFunctionConfig
 * @property {string} id - 必須。識別子。内部で使用されていて、重複があるとエラーになる。
 * @property {Function} func - 必須。実行関数。（引数を一つだけ受け取る関数のみ）
 * @property {string} [label] - 任意。表示名。(もしも未設定の場合は、idと同じになります。)
 */
/** @type {ContextFunctionConfig[]} */
const ContextFunctionsConfigs = [
  { id: 'reportError', func: reportError, label: 'エラー報告機能' }
];




//==================================================================================================================================
//============================ 自動セットアップ・整合性チェック機構(関数追加するだけならいじる必要なし) ========================================
//==================================================================================================================================

/**
 * このファイルのグローバルに定義されている二種類の関数設定オブジェクト配列の「整合性チェック」と、補完を主とした「初期化」をすることを責務とするクラス。
 * 
 * 任意の機能を使用するときにはinitializeが必要である可能性が高い(※1)ので、性能的デメリットを許容し(※2)、
 * グローバル初期化時にこのファイルの末尾のIIFEによって即時実行します。したがって、個別の機能において明示的に呼び出す必要はありません。
 * 
 * 内部で、設定オブジェクトを凍結しているので注意してください。
 * 
 * 2026/01/20時点:
 * ここでのセットアップは、このファイルのグローバルオブジェクトのみを対象にしたものであり、00_番台のファイルに定義された機能にしか影響しません。
 * 「個別の生徒のスピードプランナーのサイドバーからこのライブラリの関数を実行するボタンを追加する機能」向けの設定を整えることのみが目的です。
 * 
 * ※1 サイドメニューから関数を実行するための機能で使用する設定なので、例外はあるかもしれないが、たいてい関数実行のたびに参照される。
 * 
 * ※2 バリデーションやセットアップにも処理の時間的コストが必要ですが、計測の結果、登録されている関数の数が累計10個程度の場合、1ms程度で終了するようです。
 */
class FunctionsConfigsInitializer {
  /**
   * バリデーションとセットアップをする関数。外からはこの関数を実行すればいい。
   * もし問題があればエラーを出す。
   */
  static initialize() {
    FunctionsConfigsInitializer._validate();   //バリデーション
    FunctionsConfigsInitializer._setup();      //セットアップ
  }

  /**
   * バリデーションする内部補助関数。
   * 1. id, funcの存在と、型チェック
   * 2. (登録されている場合は)labelの型チェック
   * 3. id, func, labelの重複チェック(labelは存在しない場合は対象外)
   * 4. FunctionsConfigs と ContextFunctionsConfigsに登録されているfuncの引数が、前者は0,後者は1になっていることを確認(contextの必要性の有無)
   * @private
   */
  static _validate() {
    const errors = [];//エラーをためて最後に排出するための配列

    const allFuncsConfigs = [...FunctionsConfigs, ...ContextFunctionsConfigs];//設定されたすべての関数の配列

    const seenIdsSet = new Set();//重複検知目的のSet群
    const seenFuncsSet = new Set();
    const seenLabelsSet = new Set();

    //すべての関数に共通のチェック
    allFuncsConfigs.forEach(config => {
      if (!config || typeof config !== 'object') {//configがオブジェクトかどうか確認
        errors.push(`関数の設定オブジェクトの配列にオブジェクトではないものが入っていました。検出されたもの:\n${config}`);
        return;
      }
      const id = config.id;
      const func = config.func;
      const label = config.label
      if (!id || typeof id !== 'string') {//idが適切かチェック。
        errors.push(`idが無効なものでした。各コンフィグごとに一意なstring型の変数にしてください。対応する関数名にすることを推奨します。\n該当id:${id}\n該当コンフィグ(function型のプロパティは表示されない):\n${JSON.stringify(config)}`);
      }
      if (!func || typeof func !== 'function') {//funcが適切かチェック。
        errors.push(`funcが無効なものでした。各コンフィグごとに一意なfunction型の変数にしてください。つまり同じ関数に対して複数のコンフィグは禁止です。また、挙動の安定のために、グローバル関数を設定することをおススメします。\n該当func:${func}\n該当コンフィグ(function型のプロパティは表示されない):\n${JSON.stringify(config)}`);
      }
      //重複チェック(labelはundefined可なのでundefinedっぽかったらスルー)
      if (seenIdsSet.has(id)) errors.push(`重複しているidがみつかりました！\n該当id:${id}\n該当コンフィグ(function型のプロパティは表示されない):\n${JSON.stringify(config)}`);
      if (seenFuncsSet.has(func)) errors.push(`重複しているfuncがみつかりました！\n該当func:${func}\n該当コンフィグ(function型のプロパティは表示されない):\n${JSON.stringify(config)}`);
      if (!!label && seenLabelsSet.has(label)) errors.push(`重複しているlabelがみつかりました！\n該当label:${label}\n該当コンフィグ(function型のプロパティは表示されない):\n${JSON.stringify(config)}`);
      //見たものを登録 undefinedなら登録する意味がないのでスルー
      if (id !== undefined) seenIdsSet.add(id);
      if (func !== undefined) seenFuncsSet.add(func);
      if (label !== undefined) seenLabelsSet.add(label);
    })

    //contextを受け取らない関数は引数を持たない必要がある
    FunctionsConfigs.forEach(config => {
      if (!config || typeof config !== 'object') return;
      const func = config.func;
      if (!!func && typeof func === 'function') {
        if (func.length !== 0) errors.push(`引数が0でない関数がFunctionsConfigsにみつかりました！\n該当func:${config.func}\n該当コンフィグ(function型のプロパティは表示されない):\n${JSON.stringify(config)}`);
      }
    })

    //contextを受け取る関数は引数を一つだけ持つ必要がある
    ContextFunctionsConfigs.forEach(config => {
      if (!config || typeof config !== 'object') return;
      const func = config.func;
      if (!!func && typeof func === 'function') {
        if (func.length !== 1) errors.push(`引数が1でない関数がContextFunctionsConfigsにみつかりました！(contextオブジェクトを一つだけ受け取るようにして、contextから必要なデータを取り出す仕組みにしてください)\n該当func:${config.func}\n該当コンフィグ(function型のプロパティは表示されない):\n${JSON.stringify(config)}`);
      }
    })

    if (errors.length !== 0) throw new Error("FunctionConfigs 設定エラー:\n" + errors.map(str => ` - ${str}`).join("\n"));
  }

  /**
   * セットアップする内部補助関数。
   * 1. labelが登録されていない場合はidをlabelとして自動登録
   * 2. 関数の設定オブジェクトを、全体配列と中身の関数ごとのオブジェクト、いずれもfreeze(再代入不可能に)させる。
   * @private
   */
  static _setup() {
    FunctionsConfigs.forEach(config => {
      if (!config.label) config.label = config.id;//label補完
      Object.freeze(config);
    })
    ContextFunctionsConfigs.forEach(config => {
      if (!config.label) config.label = config.id;//label補完
      Object.freeze(config);
    })
    //凍結
    Object.freeze(FunctionsConfigs);
    Object.freeze(ContextFunctionsConfigs);
  }
}
(function () { FunctionsConfigsInitializer.initialize() })();// IIFE



