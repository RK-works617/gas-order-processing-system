function orderProcessingSystem(e) {
  let responses=e.response.getItemResponses();
  let name=responses[0].getResponse();
  let address=responses[1].getResponse();
  let mailad=responses[2].getResponse();
  let mer=responses[3].getResponse();
  let amount=responses[4].getResponse();
  let sheet1=SpreadsheetApp.openById("*********").getSheetByName("注文管理表");
  let sheet2=SpreadsheetApp.openById("*********").getSheetByName("商品リスト");
  // console.log(name);
  let a=sheet1.getLastRow()+1;
  sheet1.getRange(a,1).setValue(name);
  sheet1.getRange(a,2).setValue(address);
  sheet1.getRange(a,3).setValue(mailad);
  sheet1.getRange(a,4).setValue(mer);
  sheet1.getRange(a,5).setValue(amount);
  let price="";
  let stock="";
  let b=sheet2.getLastRow();
  for (let i=2; i<=b; i++){
    if(mer==sheet2.getRange(i,1).getValue()){
      price=sheet2.getRange(i,2).getValue();
      stock=sheet2.getRange(i,3).getValue()-amount;
      sheet2.getRange(i,3).setValue(stock);
      if(stock<=10){
        GmailApp.sendEmail("**********@gmail.com","在庫が僅少です",mer+"の在庫が残り"+stock+"個です。");
      }
    }
  }
  let totalp=price*amount
  let body=name+"様"+"\n\n"+"ご注文を受け付けました。"+"\n\n"+"ご注文商品"+"\n"+mer+"\n\n"+"注文数量"+"\n"+amount+"\n\n"+"購入価格"+"\n"+totalp;
  GmailApp.sendEmail(mailad,"注文受付",body);
}

function createTrigger(){
  let form=FormApp.openById("**********");
  ScriptApp.newTrigger("orderProcessingSystem").forForm(form).onFormSubmit().create();
}
