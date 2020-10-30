
function copyTriger() {
	getPrevPurchaseManagementSheet();//過去の台帳の保持
	let extractedData = calculateCurrentPurchaseAmount(),//カスタムIDと種別と勘定科目が一致したものを合計する
	currentPurchaseData = formateData(extractedData);
	exportCurrentPurchaseData(currentPurchaseData);
}

//「削除」ボタンを押すと「請求書（明細別）_{媒体名}」シートの先頭2行を除き削除する
function deleteTriger() {
	const byItemList = getByItemList();//「請求書（明細別）_{媒体名}」シートを特定
	let lastRow = byItemList.getLastRow();//「請求書（明細別）_{媒体名}」のシートデータが存在している最終行を取得
	byItemList.getRange(4,1,lastRow,43).clear();//先頭2行を除いた「請求書（明細別）_{媒体名}」のシートデータを削除
	byItemList.getRange(2,6,1,2).clearContent();//「請求書（明細別）_{媒体名}」シートのF2:G2セルを削除
}

//過去の台帳の保持
function getPrevPurchaseManagementSheet() {
	const ss = SpreadsheetApp.getActiveSpreadsheet(),//アクティブなスプレッドシートを特定
	purchaseManagementSheet = getPurchaseManagementSheet();//「仕入管理表{媒体名}」シートを特定
	purchaseManagementSheet.copyTo(ss);//アクティブなスプレッドシートに「仕入管理表{媒体名}」シートをコピー
}

//「請求書(明細別)_{媒体名}」シートの当月分の書き込み情報を計算
//カスタムIDと種別と勘定科目が一致したものを合計する
function calculateCurrentPurchaseAmount() {
	const byItemList = getByItemList();//「請求書（明細別）_{媒体名}」シート特定
	let valueOfByItemList = byItemList.getDataRange().getValues(),//「請求書（明細別）_{媒体名}」シートデータを全件取得
		deleteRows = [],//削除する行数を保持しておく配列を宣言
		tmp = 0;//

	for (let c = 3; valueOfByItemList.length > c; c++) {//valueOfByItemListの数だけ下記を実行
		if (valueOfByItemList[c][4] === '-' && valueOfByItemList[c][6] === '') {//タスクIDが-かつ勘定科目が空白であった場合
			valueOfByItemList[c][6] = valueOfByItemList[c - 1][6];//1個上の勘定科目を入力
		} else if (typeof valueOfByItemList[c][4] == 'number' && valueOfByItemList[c][6] === '') {//タスクIDが数値型であり、勘定科目欄が空白である場合
			valueOfByItemList[c][6] = '勘定科目未入力';//勘定科目欄に勘定科目未入力と入力
		}

		if (valueOfByItemList[c][4] !== '-') {//タスクIDが-でなければ
			byItemList.getRange(c + 1, 7).setBackground('yellow');//勘定科目セルを黄色にする
		}
	}

	let extractedData = extractData(valueOfByItemList);//「valueOfByItemList」から仕入管理表作成に必要な情報だけを抽出する

	extractedData.forEach((value, i, self) => {
		self.forEach((value2, i2) => {
			if (i !== i2 && self[i][1] === value2[1] && self[i][2] === value2[2] && self[i][3] === value2[3]) {//カスタムID、種別、勘定科目が一致すれば
				self[i][4] += value2[4];//金額を足す
				deleteRows.push(i2);//金額を足した行は削除対象配列に記載する
			}
		})

		//重複のあるデータを削除している
		for (let j = 0; deleteRows.length > j; j++) {//deleteRowsの数だけ下記を実行
			self.splice(deleteRows[j] - j, 1);//削除する
		}
		deleteRows = [];//処理終了後にdeleteRowsをリセット
	});


	sumToalAmount();	//F2セルに合計金額を記載

	return extractedData;
}

//源泉徴収税も含んだ金額を出力している
function sumToalAmount() {
	const byItemList = getByItemList();//「請求書（明細別）_{媒体名}」シートを特定
	let valueOfByItemList = byItemList.getDataRange().getValues(),//「請求書（明細別）_{媒体名}」シートデータを全件取得
	tmp = 0;//金額保持用の変数宣言
	
	for(let i = 3; valueOfByItemList.length > i; i++){//valueOfByItemListの数だけ下記を実行
		tmp += valueOfByItemList[i][10];//tmpに金額を足す
	}
	byItemList.getRange('F2').setValue(tmp);//「請求書（明細別）_{媒体名}」のF2セルにtmpを出力
	byItemList.getRange('G2').setValue(`=sum(H2:K2)`);//「請求書（明細別）_{媒体名}」のG2セルにsum関数を出力
}

//「請求書(明細別)_{媒体名}」シートから必要情報のみを取得する
//extractedData = [請求元名, カスタムID,　種別, 勘定科目, 金額]
function extractData(valueOfByItemList) {
	let tmp = [],//2次元配列を作成するための一時的な変数を宣言
		extractedData = [];//[請求元名,カスタムID,種別,勘定科目,金額]を2次元配列として入れる変数を宣言
		valueOfByItemList.shift();//「請求書(明細別)_{媒体名}」の先頭2行は見出し行であるため削除
  	valueOfByItemList.shift();
  	valueOfByItemList.shift();
		valueOfByItemList.forEach((arr, i, self) => {

		//源泉分を弾く
		if (self[i][5] !== '源泉徴収税') {//品目が源泉徴収税でなければ下記を実行
      if(self[i][5] !== '源泉所得税（経費）'){//品目が源泉所得税(経費)でなければ下記を実行
			tmp.push(self[i][12]);//請求元名をtmpに挿入
			tmp.push(self[i][13]);//カスタムIDをtmpに挿入
			tmp.push(self[i][42]);//種別をtmpに挿入
			tmp.push(self[i][6]);//勘定科目をtmpに挿入
			tmp.push(Number(self[i][10]));//数値型にした金額をtmpに挿入
			extractedData.push(tmp);//tmpをextractedDataに挿入
			tmp = [];//tmpの中身を消去
      }
		}
	});
	return extractedData
}


//過去分の台帳と突合し、情報がある人間の金額を追加
function formateData(extractedData) {
	const purchaseManagementSheet = getPurchaseManagementSheet(),//「仕入管理表{媒体名}」シートを特定
		configSheet = getConfigSheet();//「config」シートを特定
	let creationMounth = configSheet.getRange('C3').getValue(),//「config」シートのC3セルの作成月を取得
		creationMouthPoint,//のちで使う変数を宣言
		deleteRows = [],//削除用の配列を宣言
		valueOfPurchaseManagement = purchaseManagementSheet.getDataRange().getValues();//「仕入管理表{媒体名}」のシートデータを全件取得

	//作成月が1月2月の時の調整
	if (creationMounth <= 1) {
		creationMouthPoint = creationMounth + 19;
	} else if (creationMounth > 2) {
		creationMouthPoint = creationMounth + 7;
	}

//先月以前に
	valueOfPurchaseManagement.forEach((arr, i) => {
		extractedData.forEach((arr2, i2) => {
			if (valueOfPurchaseManagement[i][5] === extractedData[i2][1] && valueOfPurchaseManagement[i][7] === extractedData[i2][2] && valueOfPurchaseManagement[i][8] === extractedData[i2][3]) {
				valueOfPurchaseManagement[i][creationMouthPoint] = extractedData[i2][4];
				deleteRows.push(i2);//extractedDataから削除する列を取得
			}
		})
	})

	//削除する列を昇順に整理
	deleteRows.sort((a, b) => {
		return (a < b ? -1 : 1);
	})

	//仕入れ管理表にすでに名前があるデータを削除

	for (let i = 0; deleteRows.length > i; i++) {
		extractedData.splice(deleteRows[i] - i, 1)
	}

	valueOfPurchaseManagement = mergeValueOfPurchaseListToExtractedData(extractedData, valueOfPurchaseManagement);
	return valueOfPurchaseManagement;
}


//新規追加者を仕入管理表配列に挿入する
function mergeValueOfPurchaseListToExtractedData(extractedData, valueOfPurchaseManagement) {
	const configSheet = getConfigSheet(),//「config」シートを特定
		stockingCode = configSheet.getRange('C4').getValue();//「config」シートのC4セルを取得
	let lastRow = 1;//後ほど使う変数を宣言
//既存の最終行を特定している
	valueOfPurchaseManagement.forEach(value => {
		if (typeof value[1] === 'string') {
			if (value[1].indexOf(stockingCode) >= 0) {
				lastRow++;
			}
		}
	})

	let insertData = shapeInsertData(extractedData);

	//新規追加者を挿入する
	insertData.forEach(arr => {
		valueOfPurchaseManagement.splice(lastRow, 0, arr);
	})

	return valueOfPurchaseManagement;
}

//spliceで挿入するデータの整形
function shapeInsertData(extractedData) {
	const configSheet = getConfigSheet(),//「config」シートを特定
		stockingCode = configSheet.getRange('C4').getValue(),//「config」シートのC4セルを取得
		mediaName = configSheet.getRange('C5').getValue(),//「config」シートのC5セルを取得
		mediaCode = configSheet.getRange('C6').getValue();//「config」シートのC6セルを取得
	let tmp = [],//のちで使う変数を宣言
		insertData = [],//のちで使う変数を宣言
		creationMounth = configSheet.getRange('C3').getValue(),//「config」シートのC3セルを取得
		creationMouthPoint;//のちで使う変数を宣言

			//作成月が1月2月の時の調整
	if (creationMounth <= 1) {
		creationMouthPoint = creationMounth + 19;
	} else if (creationMounth > 2) {
		creationMouthPoint = creationMounth + 7;
	}

	extractedData.forEach((value, i, self) => {
		let supplierName = findSupplierName(value);

		tmp.push(' ');//A列
		tmp.push(stockingCode);//B列
		tmp.push(' ');//C列
		tmp.push(mediaCode);//D列
		tmp.push(mediaName);//E列
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
		tmp.push(' ');//P列
		tmp.push(' ');//Q列
		tmp.push(' ');//R列
		tmp.push(' ');//S列
		tmp.push(' ');//T列
		tmp.push(' ');//U列
		tmp.push(' ');//V列
		tmp.push(0)//W列
		tmp.push(' ');//X列
		tmp[creationMouthPoint] = extractedData[i][4];
		insertData.push(tmp);
		tmp = [];

	});
	return insertData;
}

//名前が空欄であるので仕入先台帳と仕入先codeで突合し、名前を入力
function findSupplierName(value) {
	let ValueOfSupplierLedgerSheet = getSupplierLedgerSheet().getDataRange().getValues(),//仕入先台帳シートを特定
		supplierName;//のちに使う変数を宣言

	ValueOfSupplierLedgerSheet.some((value2, i2, self2) => {
		if (value[1] === self2[i2][0]) {
			supplierName = self2[i2][1];
		}
	});

	return supplierName
}

//特定箇所を関数に変更するための処理
//台帳に貼り付け

function exportCurrentPurchaseData(currentPurchaseData) {
	const purchaseManagementSheet = getPurchaseManagementSheet(),//「仕入管理表{媒体名}」シートを特定
		obj = processCurrentPurchaseData(currentPurchaseData);
	let ledgerArr = obj.ledgerArr,
		ensureArr = obj.ensureArr;

	ledgerArr = initLedgerArr(ledgerArr);
	ensureArr = initEnsureArr(ensureArr, ledgerArr);

	purchaseManagementSheet.clear();
	purchaseManagementSheet.getRange(1, 1, ledgerArr.length, 24).setValues(ledgerArr);
	// purchaseManagementSheet.getRange(ledgerArr.length + 1, 1, ensureArr.length, 24).setValues(ensureArr);

	writeBorder(purchaseManagementSheet, ledgerArr, ensureArr);

}

//確認用エリアと台帳エリアを分離し、加工用の関数に渡す
function processCurrentPurchaseData(currentPurchaseData) {
	const configSheet = getConfigSheet(),//「config」シートを特定
		stockingCode = configSheet.getRange('C4').getValue();//「config」シートのC4セルを取得
	let lastRow = 1,//金額部分の最終行を特定するために「lastRow」を宣言する
		ledgerArr,//のちに使う変数を宣言
		ensureArr;//のちに使う変数を宣言


		currentPurchaseData.forEach(value => {
		if (typeof value[1] === 'string') {
			if (value[1].indexOf(stockingCode) >= 0) {
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
	const purchaseManagementSheet = getPurchaseManagementSheet();//「仕入管理表{媒体名}」シートを特定
	let lastRowIndex = ledgerArr.length - 1,//スプレッドシート上で位置を指定したいので-1を行う
		startRangeByMouth,//あとで使う変数宣言
		endRangeByMouth,//あとで使う変数宣言
		startRangeByPerson,//あとで使う変数宣言
		endRangeByPerson;//あとで使う変数宣言

	ledgerArr = sortCustomId(ledgerArr);//カスタムIDを昇順にする

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
			endRangeByPerson = purchaseManagementSheet.getRange(i + 1, 22).getA1Notation();
			self[i][22] = `=sum(${startRangeByPerson}:${endRangeByPerson})`;
		}
	});


	return ledgerArr;
}

//先頭行と最終行を除いた台帳エリアデータに対し、仕入先codeの昇順ソートを行う
function sortCustomId(ledgerArr) {
	const header = ledgerArr[0],//見出し行をheaderに退避
		footer = ledgerArr[ledgerArr.length - 1],//最終行をfooterに退避
		configSheet = getConfigSheet(),//「config」シートを特定
		stockingCode = configSheet.getRange('C4').getValue();//「config」シートのC4セルを取得

	ledgerArr.shift();//先頭行を削除
	ledgerArr.pop();//最終行を削除
//カスタムIDの昇順にする
	ledgerArr.sort((a, b) => {
		return (a[5] < b[5] ? -1 : 1);
	})
//B列に存在している仕入codeを入力
	ledgerArr.forEach((arr, i, self) => {
		self[i][1] = `${stockingCode}${i + 1}`;
	})

	ledgerArr.unshift(header);
	ledgerArr.push(footer);

	return ledgerArr;
}


//確認用エリアの初期化処理、スプレッドシート関数を追加する
function initEnsureArr(ensureArr, ledgerArr) {
	let totalFeeByItemList = getByItemList().getRange(2, 6).getValue(),//「請求書（明細別）_{媒体名}」シートのF2セルを取得
		lastRowInLedgerArr = ledgerArr.length + 1,//スプレッドシート上での位置を指定したいので「ledgerArr」の最終行から-1をする
		configSheet = getConfigSheet(),//「config」シートを特定
	  creationMounth = configSheet.getRange('C3').getValue();//「config」シートのC3セルを取得

		// ensureArr.forEach((arr, i, self) => {
		// 	if (typeof self[i][2] == 'string') {
		// 		if (self[i][1].indexOf(creationMounth) === 0) {
		// 			self[i][6] = `=sum(C${lastRowInLedgerArr + i}:F${lastRowInLedgerArr + i})`
		// 			self[i][7] = totalFeeByItemList;
		// 		}
		// 	}
		// })

	return ensureArr;
}

//ボーダーの追加以外に数値の表示形式も設定しているので関数名をリネームしたいが、時間的な都合で行っていない。
function writeBorder(purchaseManagementSheet, ledgerArr, ensureArr) {
	let lastRowInLedgerArr = ledgerArr.length - 1,//スプレッドシート上での位置を指定したいので「ledgerArr」の最終行から-1をする
		lastRowInEnsureArr = ensureArr.length - 1,//現在使っていないが削除もしていない。次回更新時に削除予定
		supplierCodeFormats = [],//あとで使う変数宣言
		accountFormats = [],//あとで使う変数宣言
		totalFormats = [],//あとで使う変数宣言
		tmp = [];//あとで使う変数宣言

	ledgerArr.forEach((arr, i, self) => {
		if (self[i][8] === '勘定科目未入力') {
			purchaseManagementSheet.getRange(i + 1, 9).setBackground('#fa8072');
		}
	})


	//数値の表示形式を2次元配列で指定している
	ledgerArr.forEach(value => {
		supplierCodeFormats.push(['0']);

		for (let w = 0; 13 > w; w++) {
			tmp.push('#,##');
		}
		accountFormats.push(tmp);
		tmp = [];
	})

	for (let j = 0; 12 > j; j++) {
		for (let i = 0; 6 > i; i++) {
			tmp.push('#,##');
		}
		totalFormats.push(tmp);
		tmp = [];
	}

	//数値の表示形式を設定する
	purchaseManagementSheet.getRange(2, 6, lastRowInLedgerArr + 1, 1).setNumberFormats(supplierCodeFormats);
	purchaseManagementSheet.getRange(2, 11, lastRowInLedgerArr + 1, 13).setNumberFormats(accountFormats);
	purchaseManagementSheet.getRange(lastRowInLedgerArr + 4, 3, 12, 6).setNumberFormats(totalFormats);


	//ボーダーと着色を行っている
	purchaseManagementSheet.getRange(1, 2, 1, 22).setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
	purchaseManagementSheet.getRange(2, 2, lastRowInLedgerArr, 22).setBorder(null, null, null, null, true, true, "black", SpreadsheetApp.BorderStyle.DOTTED);
	purchaseManagementSheet.getRange(2, 2, lastRowInLedgerArr, 22).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
	purchaseManagementSheet.getRange(2, 2, lastRowInLedgerArr, 22).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
	//purchaseManagementSheet.getRange(lastRowInLedgerArr + 3, 2, 13, 9).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

	purchaseManagementSheet.getRange(1, 2, 1, 22).setBackground('#87ceeb');
	purchaseManagementSheet.getRange(lastRowInLedgerArr + 4, 11, 13,13).clear();


}
