// カスタムメニューを作成し、メニュー項目を追加する
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('新規スペース登録', 'addNewRecord')
    .addItem('PJCD変更', 'replaceValues')
    .addItem('スペース削除', 'deleteSpace')
    .addToUi();
}

// 値を入力するためのプロンプトを表示し、入力された値を返す
function showPrompt(message) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
  return result.getSelectedButton() === ui.Button.OK ? result.getResponseText() : null;
}

// メッセージを表示するアラートを表示する
function showMessage(message) {
  SpreadsheetApp.getUi().alert(message);
}

// シートを取得する
function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ConfluenceProject');
}

// カラムの値を入力するためのプロンプトを表示し、入力を受け取る
function getInputValue(promptMessage) {
  var value = showPrompt(promptMessage);
  if (value === null) return null;
  return value.trim();
}

// キーの重複を確認し、新しいキーを取得する
function getUniqueKey() {
  var sheet = getSheet();
  var valuesA = sheet.getRange('A:A').getValues().flat();
  var newKey = getInputValue('新しいKey(A列)の値を入力してください');

  while (newKey !== null && valuesA.includes(newKey)) {
    showMessage('A列に同じ値が既に存在します。別の値を入力してください。');
    newKey = getInputValue('新しいKey(A列)の値を再度入力してください');
  }

  return newKey;
}

// カスタムメニュー「新規スペース登録」が選択された時に実行される関数
function addNewRecord() {
  var sheet = getSheet();
  sheet.activate();

  var newKey = getUniqueKey();
  if (newKey === null) return;

  var newRow = [newKey];

  var columnLabels = ['SpaceName', 'PJCD', 'PJCDName', 'Overview', 'ContractNumber'];
  for (var i = 0; i < columnLabels.length; i++) {
    var promptMessage = '新しい' + columnLabels[i] + '(' + String.fromCharCode(66 + i) + '列)の値を入力してください';
    var newValue = getInputValue(promptMessage);
    if (newValue === null) return;
    newRow.push(newValue);
  }

  sheet.appendRow(newRow);
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).activate();
  showMessage('新しいレコードを追加しました');
  adjustFontAndAlignment(); // フォントと配置の調整を行う
}

// セルの値を置換する
function replaceCellValue(row, column, value) {
  getSheet().getRange(row, column).setValue(value);
}

// カスタムメニュー「PJCD変更」が選択された時に実行される関数
function replaceValues() {
  var sheet = getSheet();
  sheet.activate();

  var searchValue = getInputValue('変更対象のKey(A列)を入力してください');
  if (searchValue === null) return;

  var replaceValue1 = getInputValue('置換するPJCD(C列)を入力してください');
  if (replaceValue1 === null) return;

  var replaceValue2 = getInputValue('置換するPJCDName(D列)を入力してください');
  if (replaceValue2 === null) return;

  var valuesA = sheet.getRange('A:A').getValues().flat();
  var rowIndex = valuesA.indexOf(searchValue);

  if (rowIndex !== -1) {
    replaceCellValue(rowIndex + 1, 3, replaceValue1);
    replaceCellValue(rowIndex + 1, 4, replaceValue2);
    sheet.getRange(rowIndex + 1, 1).activate();
    showMessage('置換が成功しました');
  } else {
    showMessage('入力したKeyは存在しませんでした');
  }
}

// カスタムメニュー「スペース削除」が選択された時に実行される関数
function deleteSpace() {
  var sheet = getSheet();
  sheet.activate();

  var searchValue = getInputValue('削除対象スペースのKey(A列)を入力してください');
  if (searchValue === null) return;

  var valuesA = sheet.getRange('A:A').getValues().flat();
  var rowIndex = valuesA.indexOf(searchValue);

  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 1);
    showMessage('該当のスペースを削除しました');
  } else {
    showMessage('一致するKeyは存在しませんでした');
  }
}
