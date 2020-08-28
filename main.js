
function copyTriger() {
	getPrevPurchaseManagementSheet();
	let extractedData       = calculateCurrentPurchaseAmount();
	let currentPurchaseData = formateData(extractedData);
	exportCurrentPurchaseData(currentPurchaseData);
}

//過去の台帳の保持
function getPrevPurchaseManagementSheet() {
	const ss                      = SpreadsheetApp.openById('1XfB8xfJMUERp3WIskVrJBiReeAZ3mUuIztDKlbCPKyM');
	const purchaseManagementSheet = getPurchaseManagementSheet();
	purchaseManagementSheet.copyTo(ss);
}

//当月分の書き込み情報を計算
//カスタムIDと種別と勘定科目が一致したものを合計する
function calculateCurrentPurchaseAmount() {
	const byItemList      = getByItemList()
	let valueOfByItemList = byItemList.getDataRange().getValues(),
		deleteRows        = [],
		tmp               = 0;

	for (let c = 2; valueOfByItemList.length > c; c++) {
		//i=2の時に[i][6]が空だった場合のエラーを拾うスクリプトを実装する

		if (valueOfByItemList[c][6] === '') {
			valueOfByItemList[c][6] = valueOfByItemList[c - 1][6];
		}
	}

	let extractedData = extractData(valueOfByItemList);

	extractedData.forEach((value, i, self) => {
		self.forEach((value2, i2) => {
			if (i !== i2 && self[i][1] === value2[1] && self[i][2] === value2[2] && self[i][3] === value2[3] && self[i2][4] > 0) {
				self[i][4] += value2[4];
				deleteRows.push(i2);
			}
		})

		//重複のあるデータを削除している
		for (let j = 0; deleteRows.length > j; j++) {
			self.splice(deleteRows[j] - j, 1);
		}
		deleteRows = [];
	});

	extractedData.forEach((arr, i, self) => {
		tmp += self[i][4];
	})

	byItemList.getRange(1, 6).setValue(tmp);
	return extractedData;
}

function extractData(valueOfByItemList) {
	let tmp           = [],
		extractedData = [];

	valueOfByItemList.forEach((arr, i, self) => {

		//源泉分を弾く
		if (self[i][10] > 0) {
			tmp.push(self[i][12]);
			tmp.push(self[i][13]);
			tmp.push(self[i][42]);
			tmp.push(self[i][6]);
			tmp.push(self[i][10]);
			extractedData.push(tmp);
			tmp = [];
		}
	});

	return extractedData
}

//過去分の台帳と突合し、情報がある人間の金額を追加
//仕入れ管理表側の下準備
function formateData(extractedData) {
	const purchaseManagementSheet = getPurchaseManagementSheet(),
		configSheet               = getConfigSheet();
	let creationMounth            = configSheet.getRange('B7'),
		deleteRows                = [],
		valueOfPurchaseManagement = purchaseManagementSheet.getDataRange().getValues();

	valueOfPurchaseManagement.forEach((arr, i) => {
		extractedData.forEach((arr2, i2) => {
			if (valueOfPurchaseManagement[i][5] === extractedData[i2][1] && valueOfPurchaseManagement[i][7] === extractedData[i2][2] && valueOfPurchaseManagement[i][8] === extractedData[i2][3]) {
				valueOfPurchaseManagement[i][15] = extractedData[i2][4];
				deleteRows.push(i2);
			}
		})
	})

	deleteRows.sort((a, b) => {
		return (a < b ? -1 : 1);
	})

	//仕入れ管理表にすでに名前があるデータを削除
	//ここでextractedDataの変数名を変えた方がいい
	for (let i = 0; deleteRows.length > i; i++) {
		extractedData.splice(deleteRows[i] - i, 1)
	}

	valueOfPurchaseManagement = mergeValueOfPurchaseListToExtractedData(extractedData);
	return valueOfPurchaseManagement;
}


//新規追加者を仕入れ管理表配列に挿入する
function mergeValueOfPurchaseListToExtractedData(extractedData) {
	const purchaseManagementSheet = getPurchaseManagementSheet();
	let valueOfPurchaseManagement = purchaseManagementSheet.getDataRange().getValues(),
	lastRow                       = 1;

	//台帳最終行の判別にSMを使用しているが動的に取得できるように変更する必要がある。
	valueOfPurchaseManagement.forEach(value => {
		if (typeof value[1] === 'string') {
			if (value[1].indexOf('SM') >= 0) {
				lastRow++;
			}
		}
	})

	let insertData = shapeInsertData(extractedData, lastRow);

	//台帳配列に新規追加者を挿入する
	insertData.forEach(arr => {
		valueOfPurchaseManagement.splice(lastRow, 0, arr);
	})

	return valueOfPurchaseManagement;
}

//spliceで挿入するデータの整形
function shapeInsertData(extractedData, lastRow) {
	let tmp        = [],
		insertData = [];

	//A列~D列及びK列からX列は後ほど修正
	//媒体と作成月により動的に変更しなければいけない。
	extractedData.forEach((value, i, self) => {
		let supplierName = findSupplierName(value);

		tmp.push(' ');//A列
		tmp.push('SM');//B列
		tmp.push(' ');//C列
		tmp.push('32');//D列
		tmp.push('lifehacker');//E列
		tmp.push(self[i][1]);//F列カスタムID
		tmp.push(supplierName);//G列仕入先名
		tmp.push(self[i][2]);//H列種別
		tmp.push(self[i][3]);//I列勘定科目
		tmp.push(' ');//J列空白で出力
		tmp.push(' ');//K列
		tmp.push(' ');//L列
		tmp.push(' ');//M列
		tmp.push(' ');//N列
		tmp.push(' ');//O列
		tmp.push(self[i][4]);//P列金額
		tmp.push(' ');//Q列
		tmp.push(' ');//R列
		tmp.push(' ');//S列
		tmp.push(' ');//T列
		tmp.push(' ');//U列
		tmp.push(' ');//V列
		tmp.push(0)//W列
		tmp.push(false);//X列
		insertData.push(tmp);
		tmp = [];
		lastRow++;
	})
	return insertData;
}

//名前が空欄であるので仕入れ台帳と仕入先codeで突合し、名前を入力
function findSupplierName(value) {
	let ValueOfSupplierLedgerSheet = getSupplierLedgerSheet().getDataRange().getValues(),
		supplierName;

	ValueOfSupplierLedgerSheet.some((value2, i2, self2) => {
		if (value[1] === self2[i2][0]) {
			supplierName = self2[i2][1];
		}
	});

	return supplierName
}

//台帳に貼り付ける前に消去するスクリプトが必要
//特定箇所を関数に変更するための処理
//台帳に貼り付け
function exportCurrentPurchaseData(currentPurchaseData) {
	const purchaseManagementSheet = getPurchaseManagementSheet(),
	obj                           = processCurrentPurchaseData(currentPurchaseData);
	let ledgerArr = obj.ledgerArr,
	ensureArr     = obj.ensureArr;

	ledgerArr     = initLedgerArr(ledgerArr);
	ensureArr     = initEnsureArr(ensureArr, ledgerArr);

	purchaseManagementSheet.clear();
	purchaseManagementSheet.getRange(1, 1, ledgerArr.length, 24).setValues(ledgerArr);
	purchaseManagementSheet.getRange(ledgerArr.length + 1, 1, ensureArr.length, 24).setValues(ensureArr);

	insertCheckBox(purchaseManagementSheet, ledgerArr);
	writeBorder(purchaseManagementSheet, ledgerArr);

}

//確認用エリアと台帳エリアを分離し、加工用の関数に渡す
function processCurrentPurchaseData(currentPurchaseData) {
	let lastRow = 1,
		ledgerArr,
		ensureArr;

	//台帳最終行の判別にSMを使用しているが動的に取得できるように変更する必要がある。
	currentPurchaseData.forEach(value => {
		if (typeof value[1] === 'string') {
			if (value[1].indexOf('SM') >= 0) {
				lastRow++;
			}
		}
	});

	ledgerArr = currentPurchaseData.slice(0, lastRow + 1);
	ensureArr = currentPurchaseData.slice(lastRow + 1);

	const obj = {
		'ledgerArr': ledgerArr,
		'ensureArr': ensureArr
	}

	return obj;
}

//台帳に記入する前に仕入先codeの昇順にし、スプレッドシート関数を加える
function initLedgerArr(ledgerArr) {
	const purchaseManagementSheet = getPurchaseManagementSheet();
	let lastRowIndex              = ledgerArr.length - 1,
		startRangeByMouth,
		endRangeByMouth,
		startRangeByPerson,
		endRangeByPerson;

	ledgerArr = sortCustomId(ledgerArr);

	ledgerArr[lastRowIndex].forEach((value, i, self) => {
		if (typeof value == 'number') {
			startRangeByMouth = purchaseManagementSheet.getRange(2, i + 1).getA1Notation();
			endRangeByMouth = purchaseManagementSheet.getRange(ledgerArr.length - 1, i + 1).getA1Notation();
			self[i] = `=sum(${startRangeByMouth}:${endRangeByMouth})`;
		}
	});

	ledgerArr.forEach((value, i, self) => {
		if (typeof value[22] == 'number') {
			startRangeByPerson = purchaseManagementSheet.getRange(i + 1, 11).getA1Notation();
			endRangeByPerson   = purchaseManagementSheet.getRange(i + 1, 22).getA1Notation();
			self[i][22]        = `=sum(${startRangeByPerson}:${endRangeByPerson})`;
		}
	});

	Logger.log(ledgerArr);

	ledgerArr.forEach((arr, i, self) => {
		if (self[i][23] == true) {
			self[i][23] = false;
		}
	})

	return ledgerArr;
}

//先頭行と最終行を除いた台帳エリアデータに対し、仕入先codeの昇順ソートを行う
function sortCustomId(ledgerArr) {
	const header = ledgerArr[0];
	const footer = ledgerArr[ledgerArr.length - 1];
	ledgerArr.shift();
	ledgerArr.pop();

	ledgerArr.sort((a, b) => {
		return (a[5] < b[5] ? -1 : 1);
	})

	ledgerArr.forEach((arr, i, self) => {
		self[i][1] = `SM${i + 1}`;
	})

	ledgerArr.unshift(header);
	ledgerArr.push(footer);

	return ledgerArr;
}


//確認用エリアの初期化処理、スプレッドシート関数を追加する
function initEnsureArr(ensureArr, ledgerArr) {
	let totalFeeByItemList = getByItemList().getRange(1, 6).getValue(),
	lastRowInLedgerArr     = ledgerArr.length - 1;

	ensureArr.forEach((arr, i, self) => {
		if (typeof self[i][3] == 'string') {
			if (self[i][3].indexOf('8') === 0) {
				self[i][4] = `=sum(P2:P${lastRowInLedgerArr})`;
				self[i][5] = totalFeeByItemList;
			}
		}
	})

	return ensureArr;
}

function insertCheckBox(purchaseManagementSheet, ledgerArr) {
	let startCell = purchaseManagementSheet.getRange(2,24).getA1Notation(),
	endCell       = purchaseManagementSheet.getRange(ledgerArr.length - 1, 24).getA1Notation();

	purchaseManagementSheet.getRange(`${startCell}:${endCell}`).insertCheckboxes();
}

function writeBorder(purchaseManagementSheet, ledgerArr) {
	let lastRow = ledgerArr.length - 1;
	purchaseManagementSheet.getRange(2, 2, lastRow, 22).setBorder('DOTTED');
}










//貼り付け後、前月分のデータと相違がないように照合
//エラー発生時にはエラー吐き出す用のシートに情報を追加するのと同時にエラーシートのタブ色を赤色に変更する