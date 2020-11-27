function prepareNextMonthTriger() {
	let creationMonth = getNextMonth();
	setNextMonth(creationMonth);
	renameIfNeeded(creationMonth);
	deletePrevMonth();
}

function getNextMonth() {
  const date = new Date(),//現在の時間情報を取得
		month = date.getMonth() + 1;//現在の月を取得
	let creationMonth;//のちで使う変数宣言
//1月の場合、補正し前月の月を取得
	if (month === 1) {
		creationMonth = 12;
	} else {
		creationMonth = month - 1;
	}
	return creationMonth;
}

function setNextMonth(creationMonth) {
	const configSheet = getConfigSheet();//「config」シートを特定
	configSheet.getRange('C3').setValue(creationMonth);//「creationMonth」を「config」シートのC3セルに出力
}

//仕入管理表の状態から次月作成のための処理をするかどうか判定する
//変数名とifのネストが汚いので改善が必要
function renameIfNeeded(creationMonth) {
	const purchaseManagementSheet = getPurchaseManagementSheet(),//「仕入管理表_{媒体名}」シートを特定
	copiedPurchaseManagementSheet = getCopiedPurchaseManagementSheet();//「仕入管理表_{媒体名} のコピー」シートを特定
	let creationMonthIncopiedPurchaseManagementSheet = retrieveCreationMonth(copiedPurchaseManagementSheet);//「仕入管理表_{媒体名}」シートの最新月を格納
        creationMonthInpurchaseManagementSheet = retrieveCreationMonth(purchaseManagementSheet);//「仕入管理表_{媒体名} のコピー」シートの最新月を格納
    
	if(creationMonth === 1) {
		if(creationMonthInpurchaseManagementSheet === 12 && creationMonthIncopiedPurchaseManagementSheet === 11) {
			renameSheet(creationMonth);
		}

	} else if(creationMonth === 2) {
		if(creationMonthInpurchaseManagementSheet === 1 && creationMonthIncopiedPurchaseManagementSheet === 12) {
			renameSheet(creationMonth);
		}

	} else if(creationMonth >= 3) {
		if(creationMonth - 1 === creationMonthInpurchaseManagementSheet && creationMonth - 2 === creationMonthIncopiedPurchaseManagementSheet) {
			renameSheet(creationMonth);
		}
	}
}

//仕入管理表と仕入管理表のコピーシートの作成月を取得する
function retrieveCreationMonth(sheet) {
	const valueOfSheet = sheet.getDataRange().getValues();//「sheet」シートデータを全件取得
	let lastRow = findLastRowInNextMonth(valueOfSheet),
	amountParts = sheet.getRange(2, 11, lastRow, 12).getValues(),//「sheet」の金額部分を全件取得
	creationMonth = 0,//のちで使う変数を宣言
	creationMonthPoint = 0;//のちで使う変数を宣言
 
	amountParts.forEach(arr => {
		arr.forEach((value, i2) => {
			if(typeof arr[i2] ===  'number' && creationMonthPoint - 1 < i2){
				creationMonthPoint = i2 + 1;
			}
		});
	});

	//「creationMonthPoint」が1月2月の位置の時の調整
	if (creationMonthPoint >= 11) {
		creationMonth = creationMonthPoint - 10;
	} else if (creationMonthPoint < 10) {
		creationMonth = creationMonthPoint + 2;
	}

	return creationMonth;
}

//仕入管理表と仕入管理表のコピーの最終行を取得する
function findLastRowInNextMonth(valueOfSheet) {
	const configSheet = getConfigSheet();//「config」シートを特定
	let lastRow = 0,//のちで使う変数を宣言
	stockingCode = configSheet.getRange('C4').getValue();//「config」シートのC4セルを取得
  
		valueOfSheet.forEach(value => {
			if (typeof value[1] === 'string') {
				if (value[1].indexOf(stockingCode) >= 0) {
					lastRow++;
				}
			}
		});
	return lastRow;
}

function renameSheet(creationMonth) {
	const copiedPurchaseManagementSheet = getCopiedPurchaseManagementSheet();//「仕入管理表_{媒体名} のコピー」シートを特定
	copiedPurchaseManagementSheet.setName(`【${creationMonth - 2}月分】仕入管理表{媒体名}`);//「仕入管理表_{媒体名} のコピー」シートをリネーム
}

function deletePrevMonth() {
	const noteKeysSheet = getNoteKeysSheet(),//「noteKeys」シートを特定
	storeAccountsSheet = getStoreAccountsSheet();//「storeAccount」シートを特定
	noteKeysSheet.getDataRange().clear();//「noteKeys」シートデータを削除
	storeAccountsSheet.getDataRange().clear();//「storeAccount」シートデータを削除
}