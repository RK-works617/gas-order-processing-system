function dailySalesAggregation(){
  let sheet3=SpreadsheetApp.openById("1liiKr_ivAYwlLtxrwH8S_OQL7UV5kUqL1G-XEtnzkRA").getSheetByName("売上管理表");
  // 元のスタンドアロンスクリプトから派生しているので、変数名がsheet3となっている。
  let last=sheet3.getLastRow();
  let daily=0;
  for (i=2; i<=last; i++){
    if(sheet3.getRange(i,1).getDisplayValue()==sheet3.getRange(2,4).getDisplayValue()){
      daily=daily+Number(sheet3.getRange(i,2).getValue());
    }
  }
  sheet3.getRange(2,5).setValue(daily);
  // console.log(i);
  // console.log(sheet3.getRange(i,1).getDisplayValue());
  // console.log(sheet3.getRange(2,4).getDisplayValue());
}
