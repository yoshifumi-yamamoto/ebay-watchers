var SHEET_NAME = 'Active'
var RC_ROW = 3;     // 作成フォームのレコード開始行
var RC_COL = 2;      // 作成フォームのレコード開始列


function showModal() {
  
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
  console.log('values')
  console.log(values)
  
  // 既存レコードをクリアし、CSVのレコードを貼り付け
  // clearRecords(RC_ROW, RC_COL, sheet);
  // sheet.getRange(RC_ROW - 1, RC_COL, values.length, values[0].length).setValues(values);
}
