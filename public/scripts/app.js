// 名前空間のエイリアス名
const spreadNS = GC.Spread.Sheets;

// ライセンスキーとカルチャの設定
// spreadNS.LicenseKey = 'ここにSpreadJSのライセンスキーを設定します';
GC.Spread.Common.CultureManager.culture("ja-jp");

document.addEventListener("DOMContentLoaded", () => {
  // SpreadJSの生成
  const spread = new spreadNS.Workbook("ss");
  spread.options.allowDynamicArray = true;
  injectAI(spread);
  initSheet(spread);

  // 数式エディタの生成
  const formulaEditor = new spreadNS.FormulaPanel.FormulaEditor("fe");
  formulaEditor.options.formatWidthLimit = -1;
  formulaEditor.options.tabSize = 4;
  formulaEditor.attach(spread);

  // 数式をセルに反映
  document.getElementById("apply").addEventListener('click', (e) => {
    formulaEditor.commandManager().execute({ cmd: 'commitContentToActiveCell' });
  });

  // クリップボードにテスト用文字列を設定
  document.getElementById("test1").addEventListener('click', (e) => {
    const test1 = "COUNTIF関数を使って最も商品の多い販売店を調べてください。";
    navigator.clipboard.writeText(test1);
  });

  document.getElementById("test2").addEventListener('click', (e) => {
    const test2 = "COUNTIF関数を使わないでMODE関数を使うように変更してください。";
    navigator.clipboard.writeText(test2);
  });
});

// ワークシートの初期化
const initSheet = (spread) => {
  // データの設定
  const data = [
    { カテゴリ: 'PC本体', 商品名: 'デスクトップPC1', 価格: 128000, '販売店': 'PCショップA' },
    { カテゴリ: 'PC本体', 商品名: 'ラップトップPC1', 価格: 249000, '販売店': 'PCショップA' },
    { カテゴリ: '周辺機器', 商品名: 'マウス1', 価格: 1980, '販売店': 'PCショップA' },
    { カテゴリ: '周辺機器', 商品名: 'キーボード2', 価格: 5680, '販売店': '家電量販店B' },
    { カテゴリ: '周辺機器', 商品名: 'プリンタ1', 価格: 14480, '販売店': '家電量販店B' },
    { カテゴリ: 'PCパーツ', 商品名: 'SSD1', 価格: 7980, '販売店': 'ネットショップC' },
    { カテゴリ: 'PC本体', 商品名: 'デスクトップPC2', 価格: 168000, '販売店': 'PCショップA' },
    { カテゴリ: 'PC本体', 商品名: 'ラップトップPC2', 価格: 326000, '販売店': 'PCショップA' },
    { カテゴリ: '周辺機器', 商品名: 'マウス2', 価格: 4980, '販売店': 'PCショップA' },
    { カテゴリ: '周辺機器', 商品名: 'キーボード1', 価格: 4680, '販売店': '家電量販店B' },
  ];

  // ワークシートの基本設定
  const sheet = spread.getSheet(0);
  sheet.defaults.colWidth = 40;
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);

  // テーブルの作成
  const tableStyle = spreadNS.Tables.TableThemes.medium7;
  const table = sheet.tables.addFromDataSource("table", 1, 1, data, tableStyle);

  // 項目ヘッダのスタイル
  const headerStyle = new spreadNS.Style();
  headerStyle.backColor = "royalblue";
  headerStyle.foreColor = "white";
  headerStyle.font = "11pt メイリオ";
  headerStyle.hAlign = spreadNS.HorizontalAlign.center;
  headerStyle.vAlign = spreadNS.VerticalAlign.center;

  // 項目のスタイル
  const cellStyle = new spreadNS.Style();
  cellStyle.foreColor = "blue";
  cellStyle.font = "bold 12pt メイリオ";
  cellStyle.hAlign = spreadNS.HorizontalAlign.center;
  cellStyle.vAlign = spreadNS.VerticalAlign.center;
  cellStyle.watermark = "AIアシスタントを使って生成した数式をこのセルに設定します";

  // AIアシスタントで生成した数式の設定
  sheet.addSpan(13, 1, 1, 4);
  sheet.setValue(13, 1, "AIアシスタントで生成した数式の設定");
  sheet.setStyle(13, 1, headerStyle);
  sheet.getRange(13, 1, 11, 4).setBorder(
    new GC.Spread.Sheets.LineBorder("royalblue", spreadNS.LineStyle.medium),
    { outline: true }
  );
  sheet.addSpan(14, 1, 10, 4);
  sheet.setStyle(14, 1, cellStyle);
  sheet.setActiveCell(14, 1);
}

// AIリクエストをサーバーを経由してAIサービスに送信
const serverCallback = async (requestBody) => {
  const response = await fetch("/api/queryAI", {
    method: 'POST',
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(requestBody),
  });
  if (!response.ok) {
    throw new Error(`HTTP error! status: ${response.status}`);
  }
  return response;
};

const injectAI = (spread) => {
  spread.injectAI(serverCallback);
}