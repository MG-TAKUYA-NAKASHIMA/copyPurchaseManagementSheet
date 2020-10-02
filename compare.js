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

 


  comparePrevLedgerAndNowLedger(prevLedgerArr, nowLedgerArr)
}




function comparePrevLedgerAndNowLedger(prevLedgerArr, nowLedgerArr) {
  const configSheet = getConfigSheet(),
    creationMounth = configSheet.getRange('C3').getValue();
  let creationMouthPoint;

  //作成月が1月2月の時の調整
  if (creationMounth <= 1) {
    creationMouthPoint = creationMounth + 19;
  } else if (creationMounth > 2) {
    creationMouthPoint = creationMounth + 7;
  }

  //チェックする月が格納されている配列
  let judgeMouthPoints = new Array(creationMouthPoint - 10);
  let w = 0;
  let errorCels = [],
  tmp = [];

  //当期の前月分以前が何か月存在しているのかを取得する
  for (let c = creationMouthPoint; c > 10; c--) {
    judgeMouthPoints[w] = c - 1;
    w++;
  }


  // 前月分と当月分のkeyを比較
  // keyが一致した場合は
  // 	前月以前のデータと照合する
  // 		前月以前の金額が一致した場合
  // 		→何もしない
  // 		前月以前の金額とどれか一つでも差異が発生した場合
  // 		→エラーを吐く
  // keyが一致しない場合
  // 	当月分の以前のデータを参照する
  // 		前月以前に何も記入されていない場合
  // 		→何もしない
  // 		前月以前に何か記入されている場合
  // 		→エラーを吐く
   
  for (let i = 1; nowLedgerArr.length > i; i++) {
    nowCompareKey = nowLedgerArr[i][5] + nowLedgerArr[i][7] + nowLedgerArr[i][8];

    for (let j = 1; prevLedgerArr.length > j; j++) {
      prevCompareKey = prevLedgerArr[j][5] + prevLedgerArr[j][7] + prevLedgerArr[j][8];

      for (let t = 0; judgeMouthPoints.length > t; t++) {
        if (nowCompareKey == prevCompareKey) {
          Logger.log(nowLedgerArr[i][judgeMouthPoints[t]]);

          Logger.log(prevLedgerArr[j][judgeMouthPoints[t]]);

          Logger.log(typeof nowCompareKey)
          if (nowLedgerArr[i][judgeMouthPoints[t]] !== prevLedgerArr[j][judgeMouthPoints[t]]) {
            Logger.log('a当月')
            Logger.log(nowLedgerArr[i][judgeMouthPoints[t]])
            Logger.log('a前月')
            Logger.log(prevLedgerArr[j][judgeMouthPoints[t]]);
            tmp.push(i);
            tmp.push(judgeMouthPoints[t]);
            errorCels.push(tmp);
            tmp = [];
          }

        } else if (nowCompareKey !== prevCompareKey) {
          
          if (typeof nowLedgerArr[i][judgeMouthPoints[t]] == 'number' && nowLedgerArr[i][judgeMouthPoints[t]] !== prevLedgerArr[j][judgeMouthPoints[t]]) {
            Logger.log('b当月')
            Logger.log(nowLedgerArr[i][judgeMouthPoints[t]]);
            Logger.log('b前月')
            Logger.log(prevLedgerArr[j][judgeMouthPoints[t]]);

            tmp.push(i);
            tmp.push(judgeMouthPoints[t]);
            errorCels.push(tmp);
            tmp = [];
          }
        }
      }
    }
  }

  let uniqueErrorCels = errorCels.filter((e, index) => {
    return !errorCels.some((e2, index2) => {
      return index > index2 && e[0] == e2[0] && e[1] == e2[1];
    });
  });

  uniqueErrorCels.forEach(arr => {
    exportErrorLog(arr[0], arr[1]);
  })



}

function exportErrorLog(row, col) {
  const logSheet = getLogSheet();
  let purchaseManagementSheet = getPurchaseManagementSheet(),
    lastRow = logSheet.getLastRow(),
    tmp = [],
    d = new Date(),
    number = logSheet.getRange(lastRow, 1).getValue() + 1,
    spreadsheetId = '1XfB8xfJMUERp3WIskVrJBiReeAZ3mUuIztDKlbCPKyM',
    purchaseManagementSheetId = purchaseManagementSheet.getSheetId(),
    errorCell = purchaseManagementSheet.getRange(row + 1, col + 1).getA1Notation(),
    url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${purchaseManagementSheetId}&range=${errorCell}`,
    insertRows = [];

  tmp.push(number);
  tmp.push(d);
  tmp.push(errorCell);
  tmp.push(url);

  insertRows.push(tmp);


  logSheet.getRange(number + 2, 1, 1, 4).setValues(insertRows);
  logSheet.getRange(number + 2, 5).insertCheckboxes();
  purchaseManagementSheet.getRange(errorCell).setBackground('red');
  number, errorCell, url = '';
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
  } else if (!result) {
    logSheet.setTabColor('white');
  }
}