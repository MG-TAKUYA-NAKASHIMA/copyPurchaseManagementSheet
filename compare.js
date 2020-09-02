//前月までの台帳エリアを取得する
//現在の台帳エリアを取得する
//カスタムID、種別、科目名で金額が一致してないものを探す
//一致していなければエラーシートに吐く
function getLedgerArea() {
  const purchaseManagementSheet = getPurchaseManagementSheet(),
    copiedPurchaseManagementSheet = getCopiedPurchaseManagementSheet();
  let nowPurchaseData = purchaseManagementSheet.getDataRange().getValues(),
    prevPurchaseData = copiedPurchaseManagementSheet.getDataRange().getValues(),
    prevLastRow = 1,
    nowLastRow = 1,
    prevLedgerArr,
    nowLedgerArr;

  nowPurchaseData.forEach(value => {
    if (typeof value[1] === 'string') {
      if (value[1].indexOf('SM') >= 0) {
        nowLastRow++;
      }
    }
  });

  prevPurchaseData.forEach(value => {
    if (typeof value[1] === 'string') {
      if (value[1].indexOf('SM') >= 0) {
        prevLastRow++;
      }
    }
  });

  prevLedgerArr = prevPurchaseData.slice(0, prevLastRow);
  nowLedgerArr = nowPurchaseData.slice(0, nowLastRow);

  Logger.log(prevLedgerArr);
  Logger.log(nowLedgerArr);
  comparePrevLedgerAndNowLedger(prevLedgerArr, nowLedgerArr)
}




function comparePrevLedgerAndNowLedger(prevLedgerArr, nowLedgerArr) {

  prevLedgerArr.forEach((arr, i, self) => {
    nowLedgerArr.forEach((arr2, i2, self2) => {
      if (typeof self[i][5] == 'number' && self[i][5] === self2[i2][5] && self[i][7] === self2[i2][7] && self[i][8] === self2[i2][8]) {
        if (self[i][14] !== self2[i2][14]) {
          exportErrorLog(i2);
        }
      }
    })
  })
}

function exportErrorLog(i2) {
  const logSheet = getLogSheet();
  const purchaseManagementSheet = getPurchaseManagementSheet();
  let lastRow = logSheet.getLastRow(),
    tmp = [],
    d = new Date(),
    number = logSheet.getRange(lastRow, 1).getValue() + 1,
    spreadsheetId = '1XfB8xfJMUERp3WIskVrJBiReeAZ3mUuIztDKlbCPKyM',
    purchaseManagementSheetId = purchaseManagementSheet.getSheetId(),
    errorCell = purchaseManagementSheet.getRange(i2 + 1, 15).getA1Notation(),
    url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${purchaseManagementSheetId}&range=${errorCell}`,
    insertRows = [];

  tmp.push(number);
  tmp.push(d);
  tmp.push(errorCell);
  tmp.push(url);

  insertRows.push(tmp);
  tmp = [];

  logSheet.getRange(number + 2, 1, 1, 4).setValues(insertRows);
  logSheet.getRange(number + 2, 5).insertCheckboxes();
  purchaseManagementSheet.getRange(errorCell).setBackground('red');
  Logger.log(insertRows);
  logSheet.setTabColor('red');
}

function solveTriger() {
  const logSheet = getLogSheet();
  let valueOfLogSheet = logSheet.getDataRange().getValues();
  let result = false;
  result = valueOfLogSheet.some((arr, i, self) => self[i][4] == false)
  
  let question;
  if (result == true) {
    question = Browser.msgBox("修正は完了していますか？", Browser.Buttons.OK_CANCEL);
    if (question == 'calcel') {
      return;
    } else if (question == 'ok') {
      logSheet.setTabColor('white');
    }
  }else if(!result){
    logSheet.setTabColor('white');
  }
}