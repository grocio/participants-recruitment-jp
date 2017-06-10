//仮予約があった際に、カレンダーに書き込む関数
function sendToCalendar(e){
  try{
    //エラーがあった際に報告するアドレスです。基本的にGMailのアドレスで良いと思います。
    var experimenterMailAddress = "exp.sample.taro@gmail.com";
    
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    //新規仮予約された行番号を取得
    var num_row = sheet.getLastRow();

    //新規仮予約された行から名前を取得し、"仮予約：参加者名"の文字列を作る
    var thing = "仮予約:" + sheet.getRange(num_row, 2).getValue();

    //仮予約を記載するカレンダーを取得
    var cal = CalendarApp.getCalendarById(experimenterMailAddress);

    //仮予約の開始時間を取得
    var stime = new Date(sheet.getRange(num_row, 5).getValue());

    //仮予約の開始時間から終了時間を設定
    var etime = new Date(sheet.getRange(num_row, 5).getValue());
    etime.setMinutes(etime.getMinutes()+60);

    //仮予約情報をカレンダーに追加
    cal.createEvent(thing, stime, etime); 

  }catch(exp){
    //実行に失敗した時に通知
    MailApp.sendEmail(experimenterMailAddress, exp.message, exp.message);
  }
}

//スプレッドシート上で予約を完了させ、メール送信及びカレンダーへの書き込みを行う関数
function updateCalendar(e){
  try{
    var experimenterName = "サンプル太郎";
    var experimenterMailAddress = "exp.sample.taro@gmail.com";
    var experimenterPhone = "080XXXXAAAA";
    var experimentRoom = "ABC学部実験室XYZ";
    
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    //アクティブセル（値の変更があったセル）を取得
    var activeCell = sheet.getActiveCell();
    var activeCellRow = activeCell.getRow();
    
    if(activeCell.getColumn()==6 && activeCell.getValue()==1 && sheet.getRange(activeCellRow,7)!=1){
      //予約を記載するカレンダーを取得
      var cal = CalendarApp.getCalendarById(experimenterMailAddress);
        
      //予約の開始時間を取得
      var stime = new Date(sheet.getRange(activeCellRow, 5).getValue());

      //予約の開始時間から終了時間を設定
      var etime = new Date(sheet.getRange(activeCellRow, 5).getValue());
      etime.setMinutes(etime.getMinutes()+60);
      
      //変更した行から名前を取得し、"予約完了：参加者名(ふりがな)"の文字列を作る
      var ParticipantName = sheet.getRange(activeCellRow, 2).getValue();
      var thing = "予約完了:" + ParticipantName +'('+sheet.getRange(activeCellRow, 3).getValue()+')';
      
      //予約情報をカレンダーに追加
      cal.createEvent(thing, stime, etime);
      
      //見やすい日付
      var month = stime.getMonth()+1;
      var hour = stime.getHours();
      var week = stime.getDay();
      var day = stime.getDate();
      var yobi= new Array("日","月","火","水","木","金","土");
      var hizuke = month+"月"+day+"日 "+yobi[week]+"曜日"+hour+"時";
      
      //参加者の名前などを含む、メール本文の内容（平日か土日かで文章を変える） 
      //メールの本文（土日に予約した場合）
      if(week == 0 || week == 6){
        var text = ParticipantName + "  様\n\nこの度は心理学実験への応募ありがとうございました。\n" +
            hizuke + "からの心理学実験の予約が完了しましたのでメールいたします。\n" +
            "場所は" + experimentRoom + "です。休日は教育学部棟玄関の鍵がかかっており、外から入ることができません。実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n" +
            "ご不明な点などありましたら、" + experimenterMailAddress +"までご連絡ください。\n" +
            "当日もよろしくお願いいたします。\n\n実験責任者 " + experimenterName +
            "（当日は他の者が実験担当する可能性があります）\n" +
            "当日の連絡は" + experimenterPhone + "までお願いいたします。";
        }
      
      //平日に予約した場合のメール本文
      else{
        var text = ParticipantName + "  様\n\nこの度は心理学実験への応募ありがとうございました。\n" +
            hizuke + "からの心理学実験の予約が完了しましたのでメールいたします。\n" +
            "場所は" + experimentRoom + "です。当日は直接お越しください。\n" +
            "ご不明な点などありましたら、" + experimenterMailAddress +"までご連絡ください。\n" +
            "当日もよろしくお願いいたします。\n\n実験責任者 " + experimenterName +
            "（当日は他の者が実験担当する可能性があります）\n" +
            "当日の連絡は" + experimenterPhone + "までお願いいたします。";
      }      
      //参加者にメールを送る
      var ParticipantEmail = sheet.getRange(activeCellRow, 4).getValue();
      MailApp.sendEmail(ParticipantEmail, "実験予約完了いたしました", text);

      //実験者にも参加者に送ったものと同じ内容を念の為送っておく。いらない場合はコメントアウトしてください！  
      var explanatoryText = "以下の文を参加者にも送りました。\n"
      MailApp.sendEmail(experimenterMailAddress, "予約完了メール送信確認" + ParticipantName + hizuke, explanatoryText + text);

      sheet.getRange(activeCellRow, 7).setValue(1);
      }
  }catch(exp){
    //実行に失敗した時に通知
    MailApp.sendEmail(experimenterMailAddress, exp.message, exp.message);
  }
}
