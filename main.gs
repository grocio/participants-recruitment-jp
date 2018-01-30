//それぞれの関数をトリガーを設定することを忘れないように！！
// sendToCalendar:スプレッドシートから -> フォーム送信時
// updateCalendar:スプレッドシートから -> 値の変更
// sendReminders:時間主導型 -> 日タイマー -> 午後7時〜8時
// 各関数の最初の変数の定義を確認・変更してください

//仮予約があった際に、カレンダーに書き込む関数
function sendToCalendar(e){
  try{
    // --- 各変数の定義セクション ---
    //ご自身の実験に合わせて各変数の値を変更してください。
    
    //エラーがあった際に報告するアドレスです。基本的にGMailのアドレスで良いと思います。
    var experimenterMailAddress = "exp.sample.taro@gmail.com";
    var experimenterName = '実験者太郎'; //実験者名
    var experimentLength = 60; //実験の長さ（単位は分）
    
    // 自分が実験を担当できる時間を設定する（この時間以外に予約されたらエラーメールを予約者に返す）
    var otime = 9; //実験開始可能時間
    var ctime = 18; //実験終了時間
    
    // --- 定義セクション終了 ---
    
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    //後の関数のための列を付け足す
    sheet.getRange(1, 10).setValue('予約ステータス(1=予約完了,222=既参加,333=実施期間外,444=定員超過)//,555=二重登録,999=予約キャンセル)');
    sheet.getRange(1, 11).setValue('連絡したか');
    sheet.getRange(1, 12).setValue('リマインド日時');
    sheet.getRange(1, 13).setValue('リマインダー');
    //新規仮予約された行番号を取得
    var num_row = sheet.getLastRow();

    //新規仮予約された行から名前を取得し、"仮予約：参加者名"の文字列を作る
    //ついでにアドレスも取得しておく
    var ParticipantName = sheet.getRange(num_row, 2).getValue();
    var ParticipantEmail = sheet.getRange(num_row, 8).getValue();
    var thing = "仮予約:" + ParticipantName

    //仮予約を記載するカレンダーを取得
    var cal = CalendarApp.getCalendarById(experimenterMailAddress);

    //仮予約の開始時間を取得
    var stime = new Date(sheet.getRange(num_row, 9).getValue());

    //仮予約の開始時間から終了時間を設定
    var etime = new Date(sheet.getRange(num_row, 9).getValue());
    etime.setMinutes(etime.getMinutes()+experimentLength);

    //メールに記載する、予約日時の変数（hidukeやappo）を作成する
    var month = stime.getMonth() + 1; //返ってくる値が0~11のため、+1をする必要がある。
    var hour = stime.getHours();
    var min = ('0' + stime.getMinutes()).slice(-2);
    var week = stime.getDay();
    var day = stime.getDate();
    var yobi = new Array("日", "月", "火", "水", "木", "金", "土");
    var ehour = etime.getHours(); //実験終了時間(〜時)
    var emin = ('0' + etime.getMinutes()).slice(-2); //実験終了時間(〜分)
    var appo = month + "月" + day + "日" + '(' + yobi[week] + ") " + hour + ":" + min + '〜' + ehour + ':' + emin;
    
    //取得するカレンダーの時間帯を決定する。
    var opening = new Date(stime); //予約時刻の取得
    opening.setHours(otime, 0, 0);
    var closing = new Date(stime);
    closing.setHours(ctime, 0, 0);
    //重複の確認
    var allEvents = cal.getEvents(stime, etime);

    // --- 実験実施可能時間外に応募してきた場合 ---
    if (stime < opening || closing < etime) {
        var text = ParticipantName + "  様\n\n心理学実験実施責任者の" + experimenterName +
          "です。\nこの度は心理学実験への応募ありがとうございました。\n" +
          '申し訳ありませんが、ご希望いただいた\n\n' + appo +
          '\n\nは実験実施可能時間（' + opening.getHours() + '時〜' + closing.getHours() + '時）外です。\n\n' +
          'お手数ですが、もう一度登録し直していただきますようお願いします。\n\n' + experimenterName;
        MailApp.sendEmail(ParticipantEmail, "実験実施可能時間外です", text, {bcc: experimenterMailAddress});
        sheet.getRange(num_row, 10).setValue('時間外');
        sheet.getRange(num_row, 11).setValue(1);
        sheet.getRange(num_row, 12).setValue('N/A');
        sheet.getRange(num_row, 13).setValue('N/A');
      
    // --- カレンダーに既に登録された予定や予約と重複する時間に応募してきた場合 ---
    } else if (allEvents.length > 0) {
        var text = ParticipantName + "  様\n\n心理学実験実施責任者の" + experimenterName +
          "です。\nこの度は心理学実験への応募ありがとうございました。\n" +
          '申し訳ありませんが、ご希望いただいた\n\n' + appo +
          '\n\nにはすでに予約（予定）が入っており（タッチの差で他の方が予約をされた可能性もあります）、実験を実施することができません。\n\n' +
          'お手数ですが、もう一度別の日時で登録し直していただきますようお願いします。\n\n' + experimenterName;
        MailApp.sendEmail(ParticipantEmail, "予約が重複しています", text, {bcc: experimenterMailAddress});
        sheet.getRange(num_row, 10).setValue('重複');
        sheet.getRange(num_row, 11).setValue(1);
        sheet.getRange(num_row, 12).setValue('N/A');
        sheet.getRange(num_row, 13).setValue('N/A');
    } else {
        var text = ParticipantName + "  様\n\n心理学実験実施責任者の" + experimenterName + "です。\n" +
            "この度は心理学実験への応募ありがとうございました。\n予約の確認メールを自動で送信しております。\n\n" +
            appo + '\n\nで' + ParticipantName + '様の予約を受け付けました（まだ確定はしていません。）\n\n' +
            '後日、予約完了のメールを送信いたします。\n\n' + 'もし日時の変更等がある場合は' + experimenterMailAddress +
            'までご連絡ください。\nそれでは失礼します。\n\n' + experimenterName;
        //予約確認メールを送信
        MailApp.sendEmail(ParticipantEmail, "予約の確認", text, {bcc: experimenterMailAddress});
        //仮予約情報をカレンダーに作成
        cal.createEvent(thing, stime, etime);
    }    
  }catch(exp){
    //実行に失敗した時に通知
    MailApp.sendEmail(experimenterMailAddress, exp.message, exp.message);
  }
}

//スプレッドシート上で予約を完了させ、メール送信及びカレンダーへの書き込みを行う関数
function updateCalendar(e){
  try{
    // --- 各変数の定義セクション ---
    
    var experimenterName = "実験者太郎";
    var experimenterMailAddress = "exp.sample.taro@gmail.com";
    var experimenterPhone = "080xxxxxxxx";
    var experimentRoom = "abc学部xyz実験室";
    var experimentLength = 60; //実験の長さ（単位は分）
    
    // --- 定義セクション終了 ---
    
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    //アクティブセル（値の変更があったセル）を取得
    var activeCell = sheet.getActiveCell();
    var activeCellRow = activeCell.getRow();
    
    //予約を記載するカレンダーを取得
    var cal = CalendarApp.getCalendarById(experimenterMailAddress);

    //予約の開始時間を取得
    var stime = new Date(sheet.getRange(activeCellRow, 9).getValue());

    //予約の開始時間から終了時間を設定
    var etime = new Date(sheet.getRange(activeCellRow, 9).getValue());
    etime.setMinutes(etime.getMinutes()+experimentLength);

    //変更した行から名前を取得し、"予約完了：参加者名(ふりがな)"の文字列を作る
    var ParticipantName = sheet.getRange(activeCellRow, 2).getValue();
    var thing = "予約完了:" + ParticipantName +'('+sheet.getRange(activeCellRow, 3).getValue()+')';
    //ついでに被験者のメールアドレスも取得
    var ParticipantEmail = sheet.getRange(activeCellRow, 8).getValue();

    //見やすい日付
    var month = stime.getMonth()+1;
    var hour = stime.getHours();
    var min = ('0' + stime.getMinutes()).slice(-2);
    var week = stime.getDay();
    var day = stime.getDate();
    var yobi= new Array("日","月","火","水","木","金","土");
    var hizuke = month + "月" + day + "日" + '(' + yobi[week] + ") " + hour + ":" + min;
    
    if(activeCell.getColumn()==10 && activeCell.getValue()==1 && sheet.getRange(activeCellRow,11)!=1){
        // 予約イベントを一旦削除
        var reserve = cal.getEvents(stime, etime);
        for (var i = 0; i < reserve.length; i++) {
            if (reserve[i].getTitle() == "仮予約:" + ParticipantName) {
                reserve[i].deleteEvent();
            }
        }
        //予約確定情報をカレンダーに追加
        cal.createEvent(thing, stime, etime);
        
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
      
      //平日に予約した場合のメール本文
      }else{
        var text = ParticipantName + "  様\n\nこの度は心理学実験への応募ありがとうございました。\n" +
            hizuke + "からの心理学実験の予約が完了しましたのでメールいたします。\n" +
            "場所は" + experimentRoom + "です。当日は直接お越しください。\n" +
            "ご不明な点などありましたら、" + experimenterMailAddress +"までご連絡ください。\n" +
            "当日もよろしくお願いいたします。\n\n実験責任者 " + experimenterName +
            "（当日は他の者が実験担当する可能性があります）\n" +
            "当日の連絡は" + experimenterPhone + "までお願いいたします。";
      }      
      //参加者にメールを送る(bccで実験者にも送信する)
      MailApp.sendEmail(ParticipantEmail, "実験予約完了いたしました", text, {bcc: experimenterMailAddress});
      //リマインダーのための設定をする
      var reminder = new Date(sheet.getRange(activeCellRow, 9).getValue());
      reminder.setDate(reminder.getDate() - 1); //reminderの時刻を予約時間の1日前に設定する。
      sheet.getRange(activeCellRow, 12).setValue(reminder);
      var time = new Date(); //現在時刻の取得
      time.setHours(19); //19時に設定
      //予約を完了させた日の19時にreminderの時刻が達していない場合、"送信準備"というコードを指定のセルに入力する
      if (reminder > time) {
          var code = "送信準備";
      } else {
          var code = "前日予約のため省略";
      }
      //sendMailsが参照するためのコードをセルに入力する
      sheet.getRange(activeCellRow, 13).setValue(code);
      sheet.getRange(activeCellRow, 11).setValue(1);
      }
  
    // --- 以前実験に参加したことがあり参加を断るの場合 ---
    else if (activeCell.getColumn() == 10 && activeCell.getValue() == 222 && sheet.getRange(activeCellRow, 11) != 1) {
        // 予約イベントを削除
        var reserve = cal.getEvents(stime, etime);
        for (var i = 0; i < reserve.length; i++) {
            if (reserve[i].getTitle() == "仮予約:" + ParticipantName) {
                reserve[i].deleteEvent();
            }
        }
        var text = ParticipantName + "  様\n\n心理学実験実施責任者の" + experimenterName + "です。\nこの度は心理学実験への応募ありがとうございました。\n" +
            "大変申し訳ありませんが、" + ParticipantName + "様は以前実施した同様の実験にご参加いただいており、" +
            "今回の実験にはご参加いただけません。ご了承ください。\n\n" +
            "ご不明な点などありましたら、" + experimenterMailAddress + "までご連絡ください。\n" +
            "今後ともよろしくお願いします。\n\n" + experimenterName;
        //参加者にメールを送る
        MailApp.sendEmail(ParticipantEmail, "以前に実験にご参加いただいたことがあります", text, {
            bcc: experimenterMailAddress
        });
        sheet.getRange(activeCellRow, 11).setValue(1);
        sheet.getRange(activeCellRow, 12).setValue('N/A');
        sheet.getRange(activeCellRow, 13).setValue('N/A');
    }

    // --- もし実験期間外に応募してきた場合 ---
    else if (activeCell.getColumn() == 10 && activeCell.getValue() == 333 && sheet.getRange(activeCellRow, 11) != 1) {
        // 予約イベントを削除
        var reserve = cal.getEvents(stime, etime);
        for (var i = 0; i < reserve.length; i++) {
            if (reserve[i].getTitle() == "仮予約:" + ParticipantName) {
                reserve[i].deleteEvent();
            }
        }
        var text = ParticipantName + "  様\n\n心理学実験実施責任者の" + experimenterName + "です。\nこの度は心理学実験への応募ありがとうございました。\n" +
            '申し訳ありませんが、ご希望の日時（' + month + '月' + day + '日' + '）は実験実施期間外です。\n\n' +
            'お手数ですが、もう一度、募集の掲示や応募サイトの文面を確認し、登録し直していただきますようお願いします。\n\n' +
            "ご不明な点などありましたら、" + experimenterMailAddress + "までご連絡ください。\n\n" + experimenterName;
        //参加者にメールを送る
        MailApp.sendEmail(ParticipantEmail, "実験実施期間外です。", text, {
            bcc: experimenterMailAddress
        });
        sheet.getRange(activeCellRow, 11).setValue(1);
        sheet.getRange(activeCellRow, 12).setValue('N/A');
        sheet.getRange(activeCellRow, 13).setValue('N/A');
    }

    // --- もし応募人数を超過した場合 ---
    else if (activeCell.getColumn() == 10 && activeCell.getValue() == 444 && sheet.getRange(activeCellRow, 11) != 1) {
        // 予約イベントを削除
        var reserve = cal.getEvents(stime, etime);
        for (var i = 0; i < reserve.length; i++) {
            if (reserve[i].getTitle() == "仮予約:" + ParticipantName) {
                reserve[i].deleteEvent();
            }
        }
        var text = ParticipantName + "  様\n\n心理学実験実施責任者の" + experimenterName + "です。\nこの度は心理学実験への応募ありがとうございました。\n" +
            "大変申し訳ありませんが、応募いただいた段階ですでに募集人数の定員に達していたため、" +
            "実験に参加していただくことができません。ご了承ください。\n\n" +
            "今後、次の実験を実施する際に再度応募していただけると幸いです。\n\n" +
            "ご不明な点などありましたら、" + experimenterMailAddress + "までご連絡ください。\n" +
            "今後ともよろしくお願いいたします。\n\n" + experimenterName;
        //参加者にメールを送る
        MailApp.sendEmail(ParticipantEmail, "定員に達してしまいました", text, {
            bcc: experimenterMailAddress
        });
        sheet.getRange(activeCellRow, 11).setValue(1);
        sheet.getRange(activeCellRow, 12).setValue('N/A');
        sheet.getRange(activeCellRow, 13).setValue('N/A');
    }
  }catch(exp){
    //実行に失敗した時に通知
    MailApp.sendEmail(experimenterMailAddress, exp.message, exp.message);
  }
}

//リマインダーを実行する関数
function sendReminders(e) {
  try {
      // --- 各変数の定義セクション ---

      //参照するスプレッドシートのURLを指定する
      var url = '予約フォームのスプレッドシートのURL';
      var experimenterName = "サンプル太郎";
      var experimenterMailAddress = "exp.sample.taro@gmail.com";
      var experimenterPhone = "080XXXXAAAA";
      var experimentRoom = "ABC学部実験室XYZ";

      // --- 定義セクション終了 ---

      var sheet = SpreadsheetApp.openByUrl(url); //openByIdを使うと時限式のトリガーが正常に働くらしい
      var data = sheet.getDataRange().getValues(); //たぶんシート全体のデータを取得する
      var time = new Date().getTime(); //現在時刻の取得
      // スプレッドシートを1列ずつ参照し、該当する被験者を探していく。
      for (var row = 1; row < data.length; row++) {
          //ステータスが送信準備になっていることを確認する
          if (data[row][12] == "送信準備") {
              var reminder = data[row][11];
              // もし現在時刻がリマインド日時を過ぎていたならメールを送信
              if ((reminder != "") && (reminder.getTime() <= time)) {
                  // メールの本文の内容を作成するための要素を定義
                  var ParticipantName = data[row][1] //被験者の名前
                  //見やすい日付
                  var stime = data[row][8]
                  var week = stime.getDay();
                  var hour = stime.getHours();
                  var min = ('0' + stime.getMinutes()).slice(-2);
                  //休日（前半,後半は平日）に予約した場合のメール本文
                  if (week == 0 || week == 6) {
                      var text = ParticipantName + "  様\n\n実験者の" + experimenterName + 'です。明日参加していただく実験についての確認のメールをお送りしています。\n\n' +
                          '明日  ' + hour + ":" + min + "  から実験に参加していただく予定となっております。\n" +
                          "場所は" + experimentRoom + "です。\n\nなお、明日は休日のため教育学部棟玄関の鍵がかかっており、外から入ることができません。実験者が実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n\n" +
                          '****************\n' +
                          "また、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。\n" +
                          '****************\n\n' +
                          "ご不明な点などありましたら、" + experimenterMailAddress + "までご連絡ください。\n" +
                          "それでは明日、よろしくお願いいたします。\n\n" + experimenterName;
                  } else {
                      var text = ParticipantName + "  様\n\n実験者の" + experimenterName + 'です。明日参加していただく実験についての確認のメールをお送りしています。\n\n' +
                          '明日  ' + hour + ":" + min + "  から実験に参加していただく予定となっております。\n" +
                          "場所は" + experimentRoom + "です。実験時間に実験室まで直接お越しください。\n\n" +
                          '****************\n' +
                          "なお、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。\n" +
                          '****************\n\n' +
                          "ご不明な点などありましたら、" + experimenterMailAddress + "までご連絡ください。\n" +
                          "それでは明日、よろしくお願いいたします。\n\n" + experimenterName;
                  }
                  //参加者にメールを送る
                  var ParticipantEmail = data[row][7]
                  MailApp.sendEmail(ParticipantEmail, "明日実施の心理学実験のリマインダー", text, {
                      bcc: experimenterMailAddress
                  });
                  sheet.getRange("M" + (row + 1)).setValue("送信済み");
              }
          }
      }
  } catch (exp) {
      //実行に失敗した時に通知
      MailApp.sendEmail(experimenterMailAddress, exp.message, exp.message);
  }
}
