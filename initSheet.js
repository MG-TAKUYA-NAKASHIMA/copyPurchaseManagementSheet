function initTriger() {
  popMessage();
}

function popMessage() {
  let result = Browser.msgBox("※実行前確認※", '当該スプレッドシートを初期化しますか?', Browser.Buttons.OK_CANCEL);
  if (result == 'ok') {
    initSheet();
  } else if (result == 'cancel') {
    return;
  }
}

function initSheet() {
  const targetSheets = callSheets();

  for (let i = 0; targetSheets.length > i; i++) {
    deleteData(targetSheets[i]);
  }
}

function callSheets() {
  let arr = [],
    noteKeysSheet = getNoteKeysSheet(),
    storeAccountsSheet = getStoreAccountsSheet(),
    byItemList = getByItemList(),
    purchaseManagementSheet = getPurchaseManagementSheet(),
    copiedPurchaseManagementSheet = getCopiedPurchaseManagementSheet(),
    archiveSheets = getArchiveSheet();

  arr.push([noteKeysSheet, 1]);
  arr.push([storeAccountsSheet, 1]);
  arr.push([byItemList, 2]);
  arr.push([purchaseManagementSheet, 3]);
  arr.push([copiedPurchaseManagementSheet, 4]);

  if (archiveSheets.length > 0) {
    for (let i = 0; archiveSheets.length > i; i++) {
      arr.push([archiveSheets[i], 4])
    }
  }
  return arr;
}

function getArchiveSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheets = ss.getSheets(),
    sheetName,
    match = /【/,
    archiveSheets = [];

  for (let i = 0; sheets.length > i; i++) {
    sheetName = sheets[i].getName();
    if (sheetName.search(match) >= 0) {
      archiveSheets.push(sheets[i]);
    }
  }
  return archiveSheets;
}

function deleteData(sheetObj) {
  switch (sheetObj[1]) {
    case 1:
      sheetObj[0].getDataRange().clearContent();
      break;
    case 2:
      sheetObj[0].getRange(2, 6, 1, 2).clearContent();
      sheetObj[0].getRange(4, 1, sheetObj[0].getLastRow(), sheetObj[0].getLastColumn()).clear();
      break;
    case 3:
      sheetObj[0].getRange(2, 11, sheetObj[0].getLastRow(), 13).clearContent();
      sheetObj[0].getRange(2, 11, sheetObj[0].getLastRow(), 13).clearNote();
      break;
    case 4:
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetObj[0]);
      break;
    default:
      break;
  }
}
