function onOpen(){
  // スプレッドシートに集計実行のためのメニューを作成。
  SpreadsheetApp.getUi()
    .createMenu("売上管理")
    .addItem("日次集計","dailySalesAggregation")
    .addItem("月次集計","monthlySalesAggregation")
    .addToUi();
}

function dailySalesAggregation(){
  let sheet3=SpreadsheetApp.openById("**********").getSheetByName("売上管理表");
  // 元のスタンドアロンスクリプトから派生しているので、変数名がsheet3となっている。
  let last=sheet3.getLastRow();
  let daily=0;
  // 日付データを文字列型で取得し、売上集計日に入力した日付の売上額を集計。
  for (i=2; i<=last; i++){
    if(sheet3.getRange(i,1).getDisplayValue()==sheet3.getRange(2,4).getDisplayValue()){
      daily=daily+Number(sheet3.getRange(i,2).getValue());
    }
  }
  // 売上集計額表示欄に転記。
  sheet3.getRange(2,5).setValue(daily);
}

function monthlySalesAggregation(){
  let sheet3=SpreadsheetApp.openById("**********").getSheetByName("売上管理表");
  let last=sheet3.getLastRow();
  let monthly=0;
  // 表示形式を年月のみにして、文字列型として売上集計月に入力した月の売上額を集計。
  for (i=2; i<=last; i++){
    let date=sheet3.getRange(i,1).getValue();
    let yearmonth=Utilities.formatDate(date,Session.getScriptTimeZone(),"yyyy/MM");
    if(yearmonth==sheet3.getRange(2,6).getDisplayValue()){
      monthly=monthly+Number(sheet3.getRange(i,2).getValue());
    }
  }
  sheet3.getRange(2,7).setValue(monthly);
}
