function onOpen(){
  //メニューの作成
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("受信BOX整理").addItem("実行","deleteMail").addToUi();
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("起動時に戻す").addItem("実行","onOpen").addToUi();
  
  //A列（入力列）の拡張
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  spreadsheet.setColumnWidth(1, 200);
  
  //表の項目名
  spreadsheet.getRange("A1").setValue("対象キーワード").setFontWeight("bold").setBackground("lightcyan").setHorizontalAlignment("center");
  
}

//ユーザー定義関数
function myAlert(){
  var ui = SpreadsheetApp.getUi();
  var ui_response = ui.alert(
    "受信BOX整理が完了しました。",
    "OKを押して完了してください。",
    ui.ButtonSet.OK
  );
}


function deleteMail(){
  //整理対象メールの送信者をスプレッドシートから取得する
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var last_low = spreadsheet.getLastRow();
  
  //もしA列が空ならエラーをスローする
  if (spreadsheet.getRange("A2").getValue() === ""){
    var ui = SpreadsheetApp.getUi();
    var ui_response = ui.alert(
      "キーワードを入力してください。",
      "A列が空欄です。",
      ui.ButtonSet.OK
    );
    throw new Error("キーワードが入力されていません")
  }
  
  //整理対象キーワードを配列で取得する
  var from_lists = spreadsheet.getRange(2,1,last_low-1).getValues();
  
  var delete_threads = []; 
  
  //上記キーワードのクエリによる検索→削除を行う
  from_lists.forEach(function(value,index,array){
    value.forEach(function(v,i,a){
        var query = 'from:'+　"\""+v+"\"";
      var threads = GmailApp.search(query,0,10);
      threads.forEach(function(v,i,a){
      delete_threads.push(v.getFirstMessageSubject());
      v.moveToTrash();
      });
    });
  });
  
  //削除が完了したらアラート&確認メール送信
  myAlert();
  sendMail(delete_threads);
}

function sendMail(delete_subject){
  var subject = "Gmail_Manager3実行結果のお知らせ"
  
  var body = "";
  body += "Gmail_Manager3を実行し、以下のメールを削除しました。\n";
  body += "\n"
  body += "------------------------------\n"
  
  if (delete_subject[0] === undefined)　{
    body += "削除対象メールは見つかりませんでした。";
  }
  
  delete_subject.forEach(function(v,i,a){
    body += v+"\n";
  });
  
  GmailApp.sendEmail("sasakendodt3@gmail.com", subject, body);
}

