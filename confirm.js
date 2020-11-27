//「確認」ボタンを押すと勘定科目未入力欄の検知と明細別シートの税込合計金額を出力する
function confirmTriger() {
	const byItemList = getByItemList();//「請求書（明細別）_{媒体名}」シートを特定
	let valueOfByItemList = byItemList.getDataRange().getValues();//「請求書（明細別）_{媒体名}」シートデータを全件取得

	sumToalAmount();//F2セルに合計金額を記載

	for (let c = 3; valueOfByItemList.length > c; c++) {//valueOfByItemListの数だけ下記を実行
		if (valueOfByItemList[c][4] !== '-') {//タスクIDが-でなければ
			byItemList.getRange(c + 1, 7).setBackground('yellow');//勘定科目セルを黄色にする
		}
	}
}

//源泉徴収税も含んだ金額を出力している
function sumToalAmount() {
	const byItemList = getByItemList();//「請求書（明細別）_{媒体名}」シートを特定
	let valueOfByItemList = byItemList.getDataRange().getValues(),//「請求書（明細別）_{媒体名}」シートデータを全件取得
		tmp = 0;//金額保持用の変数宣言

	for (let i = 3; valueOfByItemList.length > i; i++) {//valueOfByItemListの数だけ下記を実行
		tmp += valueOfByItemList[i][10];//tmpに金額を足す
	}
	byItemList.getRange('F2').setValue(tmp);//「請求書（明細別）_{媒体名}」のF2セルにtmpを出力
	byItemList.getRange('G2').setValue(`=sum(H2:K2)`);//
}