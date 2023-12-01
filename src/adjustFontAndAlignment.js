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
  setRowBackgroundColor(sheet);
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
  var headerRange = sheet.getRange("A1:G1");
  headerRange.setHorizontalAlignment("center");

  // A列の範囲を取得し、中央揃えに設定
  var columnARange = sheet.getRange("A2:A" + sheet.getLastRow());
  columnARange.setHorizontalAlignment("center");

  // BからF列の範囲を取得し、左揃えに設定
  var columnBtoFRange = sheet.getRange("B2:F" + sheet.getLastRow());
  columnBtoFRange.setHorizontalAlignment("left");

  // G列の範囲を取得し、右揃えに設定
  var columnGRange = sheet.getRange("G2:G" + sheet.getLastRow());
  columnGRange.setHorizontalAlignment("right");
}

// 行の背景色を設定する関数
function setRowBackgroundColor(sheet) {
  // A列の範囲を取得
  var columnARange = sheet.getRange("A2:A" + sheet.getLastRow());
  var values = columnARange.getValues();
  // A列の値が0の行の背景色を灰色に設定
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == 0) {
      sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).setBackground("#999999");
    }
  }
}

// 数値の形式を設定する関数
function setNumberFormat(sheet) {
  // G列の範囲を取得
  var columnGRange = sheet.getRange("G2:G" + sheet.getLastRow());
  // G列の数値の形式を6桁のゼロ埋めに設定
  columnGRange.setNumberFormat("000000");
}

// 罫線を設定する関数
function setBorders(range) {
  // 範囲のセルに罫線を追加
  range.setBorder(true, true, true, true, true, true);
}

