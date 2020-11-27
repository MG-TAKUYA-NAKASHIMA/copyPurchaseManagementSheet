//「再転記準備」ボタンをクリックしたら発動する
function prepareReMakeTriger() {
	const ss = SpreadsheetApp.getActiveSpreadsheet(),//当該スプレッドシートを特定
		byItemList = getByItemList(),//「請求書(明細別)_{媒体名}」シートを特定
		purchaseManagementSheet = getPurchaseManagementSheet(),//「仕入管理表_{媒体名}」シートを特定
		copiedPurchaseManagementSheet = getCopiedPurchaseManagementSheet();//「仕入管理表_{媒体名} のコピー」シートを特定
	storeNotesIfNeeded();
	storeAccount();

	ss.deleteSheet(purchaseManagementSheet);//「仕入管理表_{媒体名}」シートを削除
	ss.setActiveSheet(copiedPurchaseManagementSheet);//「仕入管理表_{媒体名} のコピー」シートをアクティブにする
	ss.moveActiveSheet(7);//アクティブなシートを7番目に挿入
	copiedPurchaseManagementSheet.setName('仕入管理表{媒体名}');//「仕入管理表_{媒体名} のコピー」シートをリネーム
	ss.setActiveSheet(byItemList);//「請求書(明細別)_{媒体名}」シートをアクティブにする
}

//メモを別シートに退避する
function storeNotesIfNeeded() {
	let lastRow = findLastRows(),
		notes = searchNotes(lastRow),
		count = 0;

	notes.forEach(arr => {
		arr.forEach((value, i2) => {
			if (arr[i2] === '') {
				count++
			}
		});
	});

	if (notes.length * notes[0].length > count) {
		let notePositions = locateNotePostion(notes),	
		noteKeys = generateNotes(notePositions);
		const noteKeysSheet = getNoteKeysSheet();//「noteKeys」シートを特定
		noteKeysSheet.getDataRange().clear();//「noteKeys」シートデータを削除
		noteKeysSheet.getRange(1, 1, noteKeys.length, 3).setValues(noteKeys);//「noteKeys」に出力
	}
}

//仕入管理表の最終行を取得
function findLastRows() {
	const configSheet = getConfigSheet(),//「config」シートを特定
	purchaseManagementSheet = getPurchaseManagementSheet();//「仕入管理表_{媒体名}」シートを特定
	let lastRow = 0,//のちで使う変数を宣言
		stockingCode = configSheet.getRange('C4').getValue(),//「config」シートのC4セルを取得
		valueOfPurchaseManagementSheet = purchaseManagementSheet.getDataRange().getValues();//「仕入管理表_{媒体名}」シートデータを全件取得

	valueOfPurchaseManagementSheet.forEach(value => {
		if (typeof value[1] === 'string') {
			if (value[1].indexOf(stockingCode) >= 0) {
				lastRow++;
			}
		}
	});

	return lastRow;
}

//メモが存在するかしないかにかかわらず、検索し取得する
function searchNotes(lastRow) {
	const purchaseManagementSheet = getPurchaseManagementSheet();//「仕入管理表_{媒体名}」シートを特定
	let notes = purchaseManagementSheet.getRange(2, 11, lastRow, 12).getNotes();//金額部分のメモを全件取得
	return notes;
}

//取得した2次元配列のメモの位置を割り出し、2次元配列を生成する
function locateNotePostion(notes) {
	let tmp = [],//のちに使う変数を宣言
		notePositions = [];//のちに使う変数を宣言

	notes.forEach((arr, i) => {
		arr.forEach((value, i2) => {
			if (arr[i2] !== '') {
				tmp.push(i);
				tmp.push(i2);
				tmp.push(arr[i2]);
				notePositions.push(tmp);
				tmp = [];
			}
		});
	});

	return notePositions;
}

//keyを生成して連想配列を作成する
function generateNotes(notePositions) {
	const purchaseManagementSheet = getPurchaseManagementSheet();//「仕入管理表_{媒体名}」シートを特定
	let noteKeys = [],//のちに使う変数を宣言
		tmp = [],//のちに使う変数を宣言
		key,//のちに使う変数を宣言
		coustomId,//のちに使う変数を宣言
		type,//のちに使う変数を宣言
		account;//のちに使う変数を宣言

	notePositions.forEach(arr => {
		coustomId = purchaseManagementSheet.getRange(arr[0] + 2, 6).getValue();//カスタムIDを取得
		type = purchaseManagementSheet.getRange(arr[0] + 2, 8).getValue();//種別を取得
		account = purchaseManagementSheet.getRange(arr[0] + 2, 9).getValue();//勘定科目を取得
		key = coustomId + type + account;//上記3種を結合し、キーにする
		tmp.push(key);
		tmp.push(arr[1]);
		tmp.push(arr[2]);
		noteKeys.push(tmp);
		tmp = [];
	});

	return noteKeys;
}

//勘定科目を保存するための関数
function storeAccount() {
	let valueOfbyItemList = getInputData(),
	accounts = extractAccount(valueOfbyItemList);
	exportAccount(accounts);
}

//「請求書(明細別)_{媒体名}」シートデータを取得
function getInputData() {
	const byItemList = getByItemList();//「請求書(明細別)_{媒体名}」シートを特定
	let valueOfbyItemList = byItemList.getDataRange().getValues();//「請求書(明細別)_{媒体名}」シートデータを全件取得
	return valueOfbyItemList;
}

//タスクIDと勘定科目を抽出する
function extractAccount(valueOfbyItemList) {
	valueOfbyItemList.splice(0,3);//見出し3行を削除
	let accounts = [],//のちで使う変数を宣言
	tmp = [];//のちで使う変数を宣言

	valueOfbyItemList.forEach((arr, i) => {
		if(arr[4] !== '-' && arr[6] !== ''){
			tmp.push(arr[4]);
			tmp.push(arr[6]);
			accounts.push(tmp);
			tmp = [];
		}
	});
	return accounts;
}


//「storeAccount」に「accounts」を出力する
function exportAccount(accounts) {
	const storeAccountsSheet = getStoreAccountsSheet();//「storeAccount」シートを特定
	storeAccountsSheet.getDataRange().clear();//「storeAccount」シートのデータを削除
	storeAccountsSheet.getRange(1, 1, accounts.length, accounts[0].length).setValues(accounts);//「accounts」を出力
}

