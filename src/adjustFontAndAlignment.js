// スプレッドシートのフォントと配置を調整する関数
function adjustFontAndAlignment() {
  // アクティブなスプレッドシートとシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  // シート内のデータ範囲を取得
  var range = sheet.getDataRange();

  // 各関数を呼び出して範囲のプロパティを設定
  setFontProperties(range);
  setAlignment(sheet);
  setNumberFormat(sheet);
  setBorders(range);
}

// フォントのプロパティを設定する関数
function setFontProperties(range) {
  // フォントの名前、サイズ、色を定義
  var fontName = "Arial";
  var fontSize = 10;
  var fontColor = "black";

  // フォントのプロパティを範囲に適用
  range.setFontFamily(fontName);
  range.setFontSize(fontSize);
  range.setFontColor(fontColor);
}

// テキストの配置を設定する関数
function setAlignment(sheet) {
  // ヘッダー行の範囲を取得し、中央揃えに設定
  var headerRange = sheet.getRange("A1:F1");
  headerRange.setHorizontalAlignment("center");

  // AからE列の範囲を取得し、左揃えに設定
  var columnBtoERange = sheet.getRange("A2:E" + sheet.getLastRow());
  columnBtoERange.setHorizontalAlignment("left");

  // F列の範囲を取得し、右揃えに設定
  var columnFRange = sheet.getRange("F2:F" + sheet.getLastRow());
  columnFRange.setHorizontalAlignment("right");
}

// 数値の形式を設定する関数
function setNumberFormat(sheet) {
  // F列の範囲を取得
  var columnFRange = sheet.getRange("F2:F" + sheet.getLastRow());
  // F列の数値の形式を6桁のゼロ埋めに設定
  columnFRange.setNumberFormat("000000");
}

// 罫線を設定する関数
function setBorders(range) {
  // 範囲のセルに罫線を追加
  range.setBorder(true, true, true, true, true, true);
}
