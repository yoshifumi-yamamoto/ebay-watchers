var SHEET_NAME = 'Active' // 出力するシート名
var RC_ROW = 4;     // 作成フォームのレコード開始行
var RC_COL = 2;      // 作成フォームのレコード開始列
var SETTINGS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定') // 設定シート情報
var RESEARCHER_GET_SHEET_ID = SETTINGS.getRange('D3').getDisplayValue() // リサーチ者を参照するシートのID
// var ITEMS = SETTINGS.getRange(3, 2, 4).getValues() // 必要な項目
// flatがうまくいかないので固定で対応（後日対応予定）
var ITEMS = ['Title', 'Custom label (SKU)', 'Start date', 'Watchers']

var LABELS = SETTINGS.getRange(3, 3, 4).getValues() // 項目のラベル
var samples = ['Item number', 'Title', 'Variation details', 'Custom label (SKU)', 'Available quantity', 'Format', 'Currency', 'Start price', 'Auction Buy It Now price', 'Reserve price', 'Current price', 'Sold quantity', 'Views (future)', 'Watchers', 'Bids', 'Start date', 'End date', 'eBay category 1 name', 'eBay category 1 number', 'eBay category 2 name', 'eBay category 2 number', 'Condition', 'eBay Product ID(ePID)', 'Listing site', 'P:UPC', 'P:EAN', 'P:ISBN']


// var DATA = [['Item number', 'Title', 'Variation details', 'Custom label (SKU)', 'Available quantity', 'Format', 'Currency', 'Start price', 'Auction Buy It Now price', 'Reserve price', 'Current price', 'Sold quantity', 'Views (future)', 'Watchers', 'Bids', 'Start date', 'End date', 'eBay category 1 name', 'eBay category 1 number', 'eBay category 2 name', 'eBay category 2 number', 'Condition', 'eBay Product ID(ePID)', 'Listing site', 'P:UPC', 'P:EAN', 'P:ISBN'], [353847193474, 'JoJo,s Bizarre Adventure All Star Battle PS3 Japanese version Used', , 'B00BHAF688', 3, 'FIXED_PRICE', 'USD', 25.74, , , 25.74, 1, , 2, , 'Jan-07-22 19:09:09 PST', 'Oct-07-22 20:09:09 PDT', 'Video Games', 139973, , , 'VERY_GOOD', , 'US', , , ], [353864972609, 'Street Fighter Collection Playstation 1 PS1 Sony Japan Capcom 1997 Japanese used', , 'B000069TD7', 1, 'FIXED_PRICE', 'USD', '38.61', , , '38.61', 0, , 9, , 'Jan-19-22 00:57:23 PST', 'Oct-19-22 01:57:23 PDT', 'Video Games', 139973, , , 'GOOD', , 'US', , , ]]


function showModal() {


  const researcherGetSheet = SpreadsheetApp.openById(RESEARCHER_GET_SHEET_ID)
  const cells =  researcherGetSheet.createTextFinder('B000069TD7').findAll()
  const cell = cells[0].getA1Notation() 
  console.log(cell)
  const r = cell.substr(1)
  console.log(r)
  const range = 'AD' + r
  const researcher = researcherGetSheet.getRange(range).getDisplayValue()
  console.log('researcher')
  console.log(researcher)

  //データがある最終列を取得
  const lastCol = researcherGetSheet.getLastColumn();

  console.log('lastCol')
  console.log(lastCol)

  // SKU列全取得
  const skus = researcherGetSheet.getSheetByName("出品 年月").getRange(1,4,10,1).getValues();
  const researchers = researcherGetSheet.getSheetByName("出品 年月").getRange(1,32,10,1).getValues();


  const formattedSkus = skus.reduce(function (acc, cur) {
    return acc.concat(cur);
  });
  
  // リサーチ担当の列を全取得
  const formattedResearchers = researchers.reduce(function (acc, cur) {
    return acc.concat(cur);
  });

  


  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const output = HtmlService.createTemplateFromFile('index');
  const data = spreadsheet.getSheetByName(SHEET_NAME);

  const projectsLastRow = data.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  output.projects = data.getRange(2, 1, projectsLastRow - 1).getValues();

  const html = output.evaluate();
  spreadsheet.show(html);
}

function sendForm(formObject) {
  const blob = formObject.myFile;
  console.log('blob')
  console.log(blob)
  // const csvText = blob.getDataAsString();
  // const values = Utilities.parseCsv(csvText)
  // console.log(values)
  const csvText = blob.getDataAsString("sjis");
  const values = Utilities.parseCsv(csvText);

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const researcherGetSheet = SpreadsheetApp.openById(RESEARCHER_GET_SHEET_ID)

    //データがある最終列を取得
  const lastCol = researcherGetSheet.getLastColumn();

  console.log('lastCol')
  console.log(lastCol)

  // SKU列全取得
  const skus = researcherGetSheet.getSheetByName("出品 年月").getRange(1,4,4500,1).getValues();
  const researchers = researcherGetSheet.getSheetByName("出品 年月").getRange(1,30,4500,1).getValues();


  const formattedSkus = skus.reduce(function (acc, cur) {
    return acc.concat(cur);
  });
  
  // リサーチ担当の列を全取得
  const formattedResearchers = researchers.reduce(function (acc, cur) {
    return acc.concat(cur);
  });

  // 必要な項目のインデックスを取得
  const indexs = ITEMS.map(function (item) {
    return values[0].indexOf(item)
})

// 2次元配列に整形
  var addValues = []
  values.slice(1).map(function (value){
    const fiteredValues = indexs.map(function (index) {
      return value[index]
    })
    console.log('fiteredValues')
    console.log(fiteredValues)
    
    if(fiteredValues[3] !== '0' && fiteredValues[3] !== ''){
      const skuIndex = formattedSkus.indexOf(fiteredValues[1])
      fiteredValues.push(formattedResearchers[skuIndex])
      addValues.push(fiteredValues)
    }
  })
  
  console.log(addValues)
  
  // 既存レコードをクリアし、CSVのレコードを貼り付け
  // clearRecords(RC_ROW, RC_COL, sheet);
  sheet.getRange(RC_ROW - 1, RC_COL, addValues.length, addValues[0].length).setValues(addValues);
}
