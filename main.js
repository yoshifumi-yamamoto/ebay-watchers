var SHEET_NAME = 'Active' // 出力するシート名
var RC_ROW = 4;     // 作成フォームのレコード開始行
var RC_COL = 2;      // 作成フォームのレコード開始列
var SETTINGS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定') // 設定シート情報
var RESEARCHER_GET_SHEET_ID = SETTINGS.getRange('D3').getDisplayValue() // リサーチ者を参照するシートのID
// var ITEMS = SETTINGS.getRange(3, 2, 4).getValues() // 必要な項目
// flatがうまくいかないので固定で対応（後日対応予定）
var ITEMS = ['Title', 'Custom label (SKU)', 'Start date','eBay category 1 name',  'Watchers']

var LABELS = SETTINGS.getRange(3, 3, 4).getValues() // 項目のラベル
var samples = ['Item number', 'Title', 'Variation details', 'Custom label (SKU)', 'Available quantity', 'Format', 'Currency', 'Start price', 'Auction Buy It Now price', 'Reserve price', 'Current price', 'Sold quantity', 'Views (future)', 'Watchers', 'Bids', 'Start date', 'End date', 'eBay category 1 name', 'eBay category 1 number', 'eBay category 2 name', 'eBay category 2 number', 'Condition', 'eBay Product ID(ePID)', 'Listing site', 'P:UPC', 'P:EAN', 'P:ISBN']


// var DATA = [['Item number', 'Title', 'Variation details', 'Custom label (SKU)', 'Available quantity', 'Format', 'Currency', 'Start price', 'Auction Buy It Now price', 'Reserve price', 'Current price', 'Sold quantity', 'Views (future)', 'Watchers', 'Bids', 'Start date', 'End date', 'eBay category 1 name', 'eBay category 1 number', 'eBay category 2 name', 'eBay category 2 number', 'Condition', 'eBay Product ID(ePID)', 'Listing site', 'P:UPC', 'P:EAN', 'P:ISBN'], [353847193474, 'JoJo,s Bizarre Adventure All Star Battle PS3 Japanese version Used', , 'B00BHAF688', 3, 'FIXED_PRICE', 'USD', 25.74, , , 25.74, 1, , 2, , 'Jan-07-22 19:09:09 PST', 'Oct-07-22 20:09:09 PDT', 'Video Games', 139973, , , 'VERY_GOOD', , 'US', , , ], [353864972609, 'Street Fighter Collection Playstation 1 PS1 Sony Japan Capcom 1997 Japanese used', , 'B000069TD7', 1, 'FIXED_PRICE', 'USD', '38.61', , , '38.61', 0, , 9, , 'Jan-19-22 00:57:23 PST', 'Oct-19-22 01:57:23 PDT', 'Video Games', 139973, , , 'GOOD', , 'US', , , ]]

// モーダルを開く
function showModal() {

  // 開いているスプレッドシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // HTMLファイルを取得
  const output = HtmlService.createTemplateFromFile('index');
  const data = spreadsheet.getSheetByName(SHEET_NAME);

  const projectsLastRow = data.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  output.projects = data.getRange(2, 1, projectsLastRow - 1).getValues();

  const html = output.evaluate();
  spreadsheet.show(html);
}

// アップロードボタン
function sendForm(formObject) {
  
  // フォームから受け取ったcsvデータ
  const blob = formObject.myFile;
  const csvText = blob.getDataAsString();
  const values = Utilities.parseCsv(csvText);

  // アップロードするファイル名を取得
  const fileName = blob.getName()

  // ファイル名にunsoldが含まれていたらUnsoldシートに出力する
  if(fileName.indexOf('unsold') !== -1) {
    SHEET_NAME = 'Unsold'
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);

  // リサーチ担当を取得するシートを取得
  const researcherGetSheet = SpreadsheetApp.openById(RESEARCHER_GET_SHEET_ID)

  //データがある最終列を取得（上手くいってない）
  const lastCol = researcherGetSheet.getLastColumn();
  console.log('lastCol')
  console.log(lastCol)

  // SKU列全取得
  const skus = researcherGetSheet.getSheetByName("出品 年月").getRange(1,4,4999,1).getValues();

  // リサーチ担当列全取得
  const researchers = researcherGetSheet.getSheetByName("出品 年月").getRange(1,30,4999,1).getValues();


  // 二次元配列を一次元配列に変換
  const formattedSkus = skus.reduce(function (acc, cur) {
    return acc.concat(cur);
  });
  
  const formattedResearchers = researchers.reduce(function (acc, cur) {
    return acc.concat(cur);
  });

  // 必要な項目のインデックスを取得
  const indexs = ITEMS.map(function (item) {
    return values[0].indexOf(item)
  })

  // 2次元配列に整形
  var addValues = []
  // １行目は項目名なのでsliceで排除
  values.slice(1).map(function (value){
    // 必要な項目の値のみ抽出
    const fiteredValues = indexs.map(function (index) {
      return value[index]
    })
    
    // watcherがある場合のみaddValuesに追加する
    if(fiteredValues[4] !== '0' && fiteredValues[4] !== '' && fiteredValues[4] !== undefined){
      const skuIndex = formattedSkus.indexOf(fiteredValues[1])
      fiteredValues.push(formattedResearchers[skuIndex])
      addValues.push(fiteredValues)
    }
  })

    // 現在日時を取得
  var today = new Date();
  // Date型データをフォーマット
  var todayStr = Utilities.formatDate(today, 'JST', 'yyyy-MM-dd HH:mm:ss');
  // 最終更新日を出力
  sheet.getRange('B1').setValue(todayStr);
  
  // 既存レコードをクリアし、CSVのレコードを貼り付け
  // clearRecords(RC_ROW, RC_COL, sheet);
  sheet.getRange(RC_ROW - 1, RC_COL, addValues.length, addValues[0].length).setValues(addValues.reverse());
}
