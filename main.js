var SHEET_NAME = 'Active' // 出力するシート名
var RC_ROW = 4;     // 作成フォームのレコード開始行
var RC_COL = 2;      // 作成フォームのレコード開始列
var SETTINGS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定') // 設定シート情報
var RESEARCHER_GET_SHEET_ID = SETTINGS.getRange('D3').getDisplayValue() // リサーチ者を参照するシートのID
var ITEMS = ['Title', 'Custom label (SKU)', 'Item number',  'Start date','eBay category 1 name',  'Watchers', 'Available quantity' ]


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

  // ebayURL列全取得
  const ebayURLs = researcherGetSheet.getSheetByName("出品 年月").getRange(2,12,20000,1).getValues();

  // リサーチ担当列全取得
  const researchers = researcherGetSheet.getSheetByName("出品 年月").getRange(2,30,20000,1).getValues();


  // 二次元配列を一次元配列に変換
  const formattedEbayURLs = ebayURLs.reduce(function (acc, cur, i) {
    return acc.concat(cur);
  });
  
  const formattedResearchers = researchers.reduce(function (acc, cur) {
    return acc.concat(cur);
  });

  // 必要な項目のインデックスを取得
  const headers = values[0]
  var indexes = [0]

  ITEMS.map(function(item){
    headers.map(function(header){
      if(item == header){
        indexes.push(headers.indexOf(header))
      }
    })
  })

  var addValues = []
  // １行目は項目名なのでsliceで排除
  values.slice(1).map(function (value){
    // 必要な項目の値のみ抽出
    const filteredValues = indexes.map(function (index) {
      // console.log(value[index])
      // if(headers[index] == 'Item number'){
      if(index == '0'){
        return 'https://www.ebay.com/itm/' + value[index]
      }
      else{
        return value[index]
      }
    })
    
    // watcherがある場合のみaddValuesに追加する
    if(filteredValues[5] !== '0' && filteredValues[5] !== '' && filteredValues[5] !== undefined){
      const ebayURLIndex = formattedEbayURLs.indexOf(filteredValues[0])
      filteredValues.push(formattedResearchers[ebayURLIndex])
      addValues.push(filteredValues)
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
