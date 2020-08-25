
function copyTriger() {
	let valueOfPrevPurchaseManagement = getPrevPurchaseManagementSheet();
}

//過去の台帳の保持
function getPrevPurchaseManagementSheet() {
	const ss = SpreadsheetApp.openById('1XfB8xfJMUERp3WIskVrJBiReeAZ3mUuIztDKlbCPKyM');
	const purchaseManagementSheet = getPurchaseManagementSheet();
	purchaseManagementSheet.copyTo(ss);
	return purchaseManagementSheet.getDataRange().getValues();
}

//当月分の書き込み情報を計算
//カスタムIDと種別と勘定科目が一致したものを合計する
function calculateCurrentPurchaseAmount() {
	const valueOfByItemList = getByItemList().getDataRange().getValues();
	let deleteRows = [];

	for (let c = 2; valueOfByItemList.length > c; c++) {
		//i=2の時に[i][6]が空だった場合のエラーを拾うスクリプトを実装する

		if (valueOfByItemList[c][6] === '') {
			valueOfByItemList[c][6] = valueOfByItemList[c - 1][6];
		}
	}

	let extractedData = extractData(valueOfByItemList);
	
	extractedData.forEach((value, i, self) => {
		self.forEach((value2, i2) => {
			if (i !== i2 && self[i][1] === value2[1] && self[i][2] === value2[2] && self[i][3] === value2[3]) {
				self[i][4] += value2[4];
				deleteRows.push(i2);
			}
		})

		for (let j = 0; deleteRows.length > j; j++) {
			self.splice(deleteRows[j] - j, 1);
		}
		deleteRows = [];
	})
}

function extractData(valueOfByItemList) {
	let tmp = [],
		extractedData = [];

	valueOfByItemList.forEach((arr, i, self) => {
		tmp.push(self[i][12]);
		tmp.push(self[i][13]);
		tmp.push(self[i][42]);
		tmp.push(self[i][6]);
		tmp.push(self[i][10]);
		extractedData.push(tmp);
		tmp = [];
	});
	extractedData.shift();
	extractedData.shift();
	return extractedData
}

//過去分の台帳と突合し、情報がある人間の金額を追加


//台帳に追加する人間を算出


//仕入先台帳の昇順になるように挿入ポイントを算出


//名前が空欄であるので仕入れ台帳と仕入先codeで突合し、名前を入力


//台帳に貼り付け
