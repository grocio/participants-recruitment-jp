//名前空間の定義
var grocio ={};

grocio.experimenterName = "サンプル太郎"
grocio.experimenterMailAddress = "sample@sample.com"
grocio.experimenterPhone = "080XXXXAAAA"
grocio.experimentRoom = "ABC学部実験室XYZ"
grocio.experimentCalndar = "実験予約カレンダー"

function sendToCalendar() {
  try{
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    //新規仮予約された行番号を取得
    var num_row = sheet.getLastRow();

    //新規仮予約された行から名前を取得し、"仮予約：参加者名"の文字列を作る
    var thing = "仮予約:" + sheet.getRange(num_row, 2).getValue();

    //仮予約を記載するカレンダーを取得
    var cals = CalendarApp.getCalendarsByName(grocio.experimentCalndar);

    //仮予約の開始時間を取得
    var stime = new Date(sheet.getRange(num_row, 4).getValue());

    //仮予約の開始時間から終了時間を設定
    var etime = new Date(sheet.getRange(num_row, 4).getValue());
    etime.setMinutes(etime.getMinutes()+60);

    //仮予約情報をカレンダーに追加
    var r = cals[0].createEvent(thing, stime, etime); 

  }catch(exp){
    //実行に失敗した時に通知
    MailApp.sendEmail(grocio.experimenterMailAddress, exp.message, exp.message);
  }
}


function updateCalendar() {
  try{
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    //アクティブセル（値の変更があったセル）を取得
    var activeCell = sheet.getActiveCell();
    var activeCellRow = activeCell.getRow()
    
    if(activeCell.getColumn()==5 && activeCell.getValue()==1 && sheet.getRange(activeCellRow,6)!=1){
      //予約を記載するカレンダーを取得
      var cals = CalendarApp.getCalendarsByName(grocio.experimentCalndar);
        
      //予約の開始時間を取得
      var stime = new Date(sheet.getRange(activeCellRow, 4).getValue());

      //予約の開始時間から終了時間を設定
      var etime = new Date(sheet.getRange(activeCellRow, 4).getValue());
      etime.setMinutes(etime.getMinutes()+60);
        
      //変更した行から名前を取得し、"予約完了：参加者名"の文字列を作る
      var thing = "予約完了:" + sheet.getRange(activeCellRow, 2).getValue();
      
      //予約情報をカレンダーに追加
      var r = cals[0].createEvent(thing, stime, etime);
      
      //見やすい日付
      var hour = stime.getHours();
      var week = stime.getDay();
      var day = stime.getDate();
      var yobi= new Array("日","月","火","水","木","金","土");
      var hizuke = "3月"+day+"日 "+yobi[week]+"曜日"+hour+"時";
      
      //参加者の名前などを含む、メール本文の内容
      var ParticipantName = sheet.getRange(activeCellRow, 2).getValue();
      var text = ParticipantName + "様\n\nこの度は心理学実験への応募ありがとうございました。\n" +
          hizuke + "からの心理学実験の予約が完了しましたのでメールいたします。\n" +
          "場所は" + grocio.experimentRoom + "です。当日は直接お越しください。\n" +
          "ご不明な点などありましたら、" + grocio.experimenterMailAddress +"までご連絡ください。\n" +
          "当日もよろしくお願いいたします。\n\n実験責任者 " + grocio.experimenterName + "（当日は他の者が実験担当いたします）\n" +
          "当日の連絡は" + grocio.experimenterPhone + "までお願いいたします。";

      //参加者にメールを送る
      var ParticipantEmail = sheet.getRange(activeCellRow, 3).getValue();
      MailApp.sendEmail(ParticipantEmail, "実験予約完了いたしました", text);

      //実験者にも参加者に送ったものと同じ内容を念の為送っておく。いらない場合はコメントアウトしてください！  
      var explanatoryText = "以下の文を参加者にも送りました。\n"
      MailApp.sendEmail(grocio.experimenterMailAddress, "予約完了メール送信確認" + ParticipantName + hizuke, explanatoryText + text);

      sheet.getRange(activeCellRow, 6).setValue(1);
      }
  }catch(exp){
    //実行に失敗した時に通知
    MailApp.sendEmail(grocio.experimenterMailAddress, exp.message, exp.message);
  }
}
