var SHEET_NAME = 'Active' // 出力するシート名
var RC_ROW = 4;     // 作成フォームのレコード開始行
var RC_COL = 2;      // 作成フォームのレコード開始列
// flatがうまくいかないので固定で対応（後日対応予定）
var ITEMS = ['Title', 'Custom label (SKU)', 'Start date','eBay category 1 name',  'Watchers']

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
