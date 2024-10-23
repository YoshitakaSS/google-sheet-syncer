function transferCheckedRows() {
  const sourceSpreadsheetId = PropertiesService.getScriptProperties().getProperty("SOURCE_SPREAD_SHEET_ID");
  const sourceSheet = SpreadsheetApp.openById(sourceSpreadsheetId).getSheetByName(
    PropertiesService.getScriptProperties().getProperty("SOURCE_SPREAD_SHEET_NAME")
  );
  const targetSpreadsheetId = PropertiesService.getScriptProperties().getProperty("TARGET_SPREAD_SHEET_ID");
  const targetSheet = SpreadsheetApp.openById(targetSpreadsheetId).getSheetByName(
    PropertiesService.getScriptProperties().getProperty("TARGET_SPREAD_SHEET_NAME")
  );

  // 列のアルファベットとインデックスのマッピング
  const sourceColumnMap = {
    'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5,
    'G': 6, 'H': 7, 'I': 8, 'J': 9, 'K': 10, 'L': 11,
    'M': 12, 'N': 13, 'O': 14, 'P': 15, 'Q': 16, 'R': 17,
    'S': 18, 'T': 19, 'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24, 'Z': 25,
    'AA': 26, 'AB': 27, 'AC': 28, 'AD': 29, 'AE': 30, 'AF': 31
  };

  // 添え字が1つずれるため定義
  const targetSheetColumnMap = {
    'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6,
    'G': 7, 'H': 8, 'I': 9, 'J': 10, 'K': 11, 'L': 12,
    'M': 13, 'N': 14, 'O': 15, 'P': 16, 'Q': 17, 'R': 18,
    'S' : 19, 'T' : 20, 'U' : 21, 'V' : 22, 'W' : 23, 'X' : 24, 'Y' : 25, 'Z' : 26,
    'AA' : 27, 'AB' : 28, 'AC' : 29, 'AD' : 30, 'AE' : 31, 'AF' : 32,
  };
  
  // ソースシートとターゲットシートのデータを取得
  let sourceData = sourceSheet.getDataRange().getValues();
  let targetData = targetSheet.getDataRange().getValues();
  
  // 転記済みのF列の値を保持するセットとそのインデックスをマップに保存
  let existingKeys = new Map();
  
  // タスクキー列の最終行を取得して追記開始行を設定
  let lastRow = targetSheet.getRange(targetSheet.getLastRow(), targetSheetColumnMap['C']).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  let targetRowIndex = lastRow + 1;

  // ターゲットシートのF列（キー）とそのインデックスをマップに保存
  for (let j = 2; j < targetData.length; j++) { // 2からスタートしてヘッダーをスキップ
    let key = targetData[j][targetSheetColumnMap['C'] - 1]; // 添字がズレる
    if (key !== '') {
      existingKeys.set(key, j); // キーと行インデックスを保存
    }
  }

  // 転記処理
  for (let i = 2; i < sourceData.length; i++) { // 2からスタートしてヘッダーをスキップ
    if (sourceData[i][sourceColumnMap['A']] === true) { // A列にチェックが入っている場合
      let sourceTaskKey = sourceData[i][sourceColumnMap['F']]; // タスクキー
      
      // キーが既に存在するか確認
      if (existingKeys.has(sourceTaskKey)) {
        // 既存行を更新（例えば、C列の内容が異なれば更新）
        let existingRowIndex = existingKeys.get(sourceTaskKey);
        let existingTitle = targetData[existingRowIndex][targetSheetColumnMap['E'] - 1]; // E列（タスクタイトル）の値
        let newTitle = sourceData[i][4]; // 転記元のタスクタイトル
        
        const insertedRow = existingRowIndex + 1;

        // もし新しいタイトルが異なる場合、更新
        if (existingTitle !== newTitle) {
          targetSheet.getRange(insertedRow, targetSheetColumnMap['B']).setValue(newTitle); // タスクタイトルを更新
        }
        // 必要に応じて他の列も更新
        targetSheet.getRange(insertedRow, targetSheetColumnMap['D']).setValue(sourceData[i][sourceColumnMap['L']]); // 主要顧客を更新
        targetSheet.getRange(insertedRow, targetSheetColumnMap['E']).setValue(sourceData[i][sourceColumnMap['M']]); // 対象企業・対象店舗を更新
        targetSheet.getRange(insertedRow, targetSheetColumnMap['F']).setValue(sourceData[i][sourceColumnMap['N']].replace(/;/g, ',')); // 対象プロダクトを更新
        targetSheet.getRange(insertedRow, targetSheetColumnMap['G']).setValue(sourceData[i][sourceColumnMap['O']].replace(/;/g, ',')); // 対象業務を更新
        targetSheet.getRange(insertedRow, targetSheetColumnMap['H']).setValue(sourceData[i][sourceColumnMap['S']]); // 保守分類を更新
        let claim = sourceData[i][sourceColumnMap['T']] === 'クレーム懸念あり';
        targetSheet.getRange(insertedRow, targetSheetColumnMap['I']).setValue(claim); // クレーム判定を更新
        targetSheet.getRange(insertedRow, targetSheetColumnMap['X']).setValue(sourceData[i][sourceColumnMap['V']]); // 着手判定更新
        targetSheet.getRange(insertedRow, targetSheetColumnMap['Y']).setValue(sourceData[i][sourceColumnMap['W']]); // CO日判定更新
        targetSheet.getRange(insertedRow, targetSheetColumnMap['Z']).setValue(sourceData[i][sourceColumnMap['X']]); // サービスイン日判定更新
      } else {
        // 新規追加（重複しない場合）
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['B']).setValue(sourceData[i][sourceColumnMap['E']]); // タスクタイトル

        let taskKey = sourceData[i][sourceColumnMap['F']]; // タスクキー
        let linkUrl = 'https://linkprocessing.atlassian.net/browse/' + taskKey;
        let hyperlinkFormula = '=HYPERLINK("' + linkUrl + '", "' + taskKey + '")';

        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['C']).setFormula(hyperlinkFormula); // ハイパーリンクを設定
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['D']).setValue(sourceData[i][sourceColumnMap['L']]); // 主要顧客
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['E']).setValue(sourceData[i][sourceColumnMap['M']]); // 対象企業・対象店舗
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['F']).setValue(sourceData[i][sourceColumnMap['N']].replace(/;/g, ',')); // 対象プロダクト
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['G']).setValue(sourceData[i][sourceColumnMap['O']].replace(/;/g, ',')); // 対象業務
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['H']).setValue(sourceData[i][sourceColumnMap['S']]); // 保守分類を更新
        let claim = sourceData[i][sourceColumnMap['T']] === 'クレーム懸念あり';
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['I']).setValue(claim); // クレーム判定を更新
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['X']).setValue(sourceData[i][sourceColumnMap['V']]); // 着手判定更新
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['Y']).setValue(sourceData[i][sourceColumnMap['W']]); // CO日判定更新
        targetSheet.getRange(targetRowIndex, targetSheetColumnMap['Z']).setValue(sourceData[i][sourceColumnMap['X']]); // サービスイン日判定更新

        // 新しいキーをマップに追加
        existingKeys.set(sourceTaskKey, targetRowIndex - 1);

        // 次の行に移動
        targetRowIndex++;
      }
    }
  }
  
  Logger.log("Checked rows have been transferred and updated.");
}

function onOpen() {
  // カスタムメニューを作成
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('PB転記にチェック済のものをAnywherePBに転記する', 'transferCheckedRows')
    .addToUi();
}
