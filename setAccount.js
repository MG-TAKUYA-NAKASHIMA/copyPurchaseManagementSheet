//「勘定科目補完」ボタンをクリックすると発動する
function accountTriger() {
  let valueOfStoreAccounts = getAccounts(),//「storeAccount」シートのシートデータを取得
  valueOfByItemList = setAccounts();//「請求書(明細別)_{媒体名}」シートを全件取得し、見出し行を削除
  mergeAccounts(valueOfStoreAccounts, valueOfByItemList);//タスクIDが一致し、勘定科目が空欄のところに勘定科目を補完する
}

//「storeAccount」シートに記載されたタスクIDと勘定科目のリストを取得する
function getAccounts() {
  const storeAccountsSheet = getStoreAccountsSheet();//「storeAccount」シートを特定
  let valueOfStoreAccounts = storeAccountsSheet.getDataRange().getValues();//「storeAccount」シートデータを全件取得
  return valueOfStoreAccounts;//シートデータを戻す
}

//「請求書(明細別)_{媒体名}」シートを全件取得し、見出し行を削除する
function setAccounts() {
  const byItemList = getByItemList();//「請求書(明細別)_{媒体名}」シートを特定
  let valueOfByItemList = byItemList.getDataRange().getValues();//「請求書(明細別)_{媒体名}」シートデータを全件取得
  valueOfByItemList.splice(0,3);//見出し3行を削除
  return valueOfByItemList;
}

//タスクIDが一致し、勘定科目が空欄のところに勘定科目を補完する
function mergeAccounts(valueOfStoreAccounts, valueOfByItemList) {
  const byItemList = getByItemList();//「請求書(明細別)_{媒体名}」シートを特定

  valueOfByItemList.forEach((arr, i) => {
    valueOfStoreAccounts.forEach(arr2 => {
      if(arr[4] !== '' && arr[6] === '') {
        if(arr[4] === arr2[0]) {
          byItemList.getRange(i + 4, 7).setValue(arr2[1]);
        }
      }
    });
  });
}