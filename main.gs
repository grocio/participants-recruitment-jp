/*誰か綺麗に書きなおしてくれるとありがたい！
それぞれの関数をトリガーを設定することを忘れないように！！
sendToCalendar:スプレッドシートから -> フォーム送信時
updateCalendar:スプレッドシートから -> 値の変更
sendReminders:時間主導型 -> 日タイマー -> 午後7時〜8時
各関数の最初の変数の定義を確認・変更してください*/

// --- 各変数の定義セクション ---
//ご自身の実験に合わせて各変数の値を変更してください。
var experimenterName = '実験者太郎'; //実験者名
var experimenterMailAddress = "github.test.participant@gmail.com";
var experimenterPhone = '080-1234-5678';
var experimentRoom = "abc学部xyz実験室";
var experimentLength = 60; //実験の長さ（単位は分）
var url = 'https://docs.google.com/spreadsheets/d/180VT_tRqmYBWCvLPlJYk_gVfSWWT47FtCGsdqTbg6T0/edit#gid=1566578720';

// 自分が実験を担当できる1日の時間を設定する（この時間以外に予約されたらエラーメールを予約者に返す）
var openTime = 9; //実験開始可能時間
var closeTime = 18; //実験終了時間
//実験の開始日の前日の日と終了日を設定する
//6月20日スタートならopenDateは19日に設定しておく
//月に関しては、-1をする必要がある点に注意する（例. 7月ならopenMonth = 7 - 1）
var openMonth = 1 - 1;//1〜12月 -> 0 ~ 11に変更
var openDate = 29;//日
var closeMonth = 12 - 1;//1〜12月 -> 0 ~ 11に変更
var closeDate = 1;//日

// 既に実験に参加済みテキスト
var textAlreadyDone = 'PARTICIPANTNAME様\n'+
'\n'+
'心理学実験実施責任者の'+experimenterName+'です。\n'+
'この度は心理学実験への応募ありがとうございました。\n'+
'大変申し訳ありませんが、PARTICIPANTNAME様は以前実施した同様の実験にご参加いただいており、今回の実験にはご参加いただけません。ご了承ください。\n'+
'\n'+
'ご不明な点などありましたら、'+experimenterMailAddress+'までご連絡ください。\n'+
'今後ともよろしくお願いします。\n'+
'\n'+
experimenterName;
// 定員に達したテキスト
var textReachedCapacity = 'PARTICIPANTNAME様\n'+
'\n'+
'心理学実験実施責任者の'+experimenterName+'です。\n'+
'この度は心理学実験への応募ありがとうございました。\n'+
'大変申し訳ありませんが、応募いただいた段階ですでに募集人数の定員に達していたため、実験に参加していただくことができません。ご了承ください。\n'+
'\n'+
'今後、次の実験を実施する際に再度応募していただけると幸いです。\n'+
'\n'+
'ご不明な点などありましたら、'+experimenterMailAddress+'までご連絡ください。\n'+
'今後ともよろしくお願いいたします。\n'+
'\n'+
experimenterName;
// 時間外・期間外テキスト
var textOutoftime = 'PARTICIPANTNAME様\n'+
'\n'+
'心理学実験実施責任者の'+experimenterName+'です。\n'+
'この度は心理学実験への応募ありがとうございました。\n'+
'申し訳ありませんが、ご希望いただいた'+
'\n'+
'APPOINTMENT'+
'\n'+
'は実験実施可能時間（'+openTime+' 時〜'+closeTime+'時）外または、実施期間（'+openMonth+'月'+openDate+'日〜'+closeMonth+'月'+closeDate+'日）外です。\n'+
'お手数ですが、もう一度登録し直していただきますようお願いします。\n'+
'\n'+
experimenterName;
// 重複テキスト
var textOverlap = 'PARTICIPANTNAME様\n'+
'\n'+
'心理学実験実施責任者の'+experimenterName+'です。\n'+
'この度は心理学実験への応募ありがとうございました。\n'+
'申し訳ありませんが、ご希望いただいた'+
'\n'+
'APPOINTMENT'+
'\n'+
'にはすでに予約（予定）が入っており（タッチの差で他の方が予約をされた可能性もあります）、実験を実施することができません。\n'+
'お手数ですが、もう一度別の日時で登録し直していただきますようお願いします。\n'+
'\n'+
experimenterName;
// 仮予約完了テキスト
var textTentativeBooking = 'PARTICIPANTNAME様\n'+
'\n'+
'心理学実験実施責任者の'+experimenterName+'です。\n'+
'この度は心理学実験への応募ありがとうございました。\n'+
'予約の確認メールを自動で送信しております。\n'+
'\n'+
'APPOINTMENT'+
'\n'+
'で予約を受け付けました（まだ確定はしていません。)'+
'\n'+
'後日、予約完了のメールを送信いたします。\n'+
'もし日時の変更等がある場合は'+experimenterMailAddress+'までご連絡ください。\n'+
'どうぞよろしくお願いいたします。\n'+
'\n'+
experimenterName;
// 予約完了テキスト(休日)'+
var textHolidayBookingDone = 'PARTICIPANTNAME様\n'+
'\n'+
'この度は心理学実験への応募ありがとうございました。\n'+
'RESERVEAPPOからの心理学実験の予約が完了しましたのでメールいたします。\n'+
'場所は'+experimentRoom+'です。休日は教育学部棟玄関の鍵がかかっており、外から入ることができません。実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n'+
'ご不明な点などありましたら、'+experimenterMailAddress+'までご連絡ください。\n'+
'当日もよろしくお願いいたします。\n'+
'\n'+
'実験責任者'+experimenterName+'（当日は他の者が実験担当する可能性があります)'+
'当日の連絡は'+experimenterPhone+'までお願いいたします。';
// 予約完了テキスト(平日)'+
var textWeekdayBookingDone = 'PARTICIPANTNAME様\n'+
'\n'+
'この度は心理学実験への応募ありがとうございました。\n'+
'RESERVEAPPOからの心理学実験の予約が完了しましたのでメールいたします。\n'+
'場所は'+experimentRoom+'です。当日は直接お越しください。\n'+
'ご不明な点などありましたら、'+experimenterMailAddress+'までご連絡ください。\n'+
'当日もよろしくお願いいたします。\n'+
'\n'+
'実験責任者'+experimenterName+'（当日は他の者が実験担当する可能性があります)'+
'当日の連絡は'+experimenterPhone+'までお願いいたします。';
// リマインダー(休日)_
var textReminderHoliday ='PARTICIPANTNAME様\n'+
'\n'+
'実験者の'+experimenterName+'です。明日参加していただく実験についての確認のメールをお送りしています。\n'+
'\n'+
'明日 BOOKEDSCHEDULEから実験に参加していただく予定となっております。\n'+
'場所は'+experimentRoom+'です。\n'+
'\n'+
'なお、明日は休日のため教育学部棟玄関の鍵がかかっており、外から入ることができません。実験者が実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n'+
'\n'+
'また、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。\n'+
'ご不明な点などありましたら、'+experimenterMailAddress+'までご連絡ください。\n'+
'それでは明日、よろしくお願いいたします。\n'+
'\n'+
experimenterName;
// リマインダー(平日)'+
var textReminderWeekday ='PARTICIPANTNAME様\n'+
'\n'+
'実験者の'+experimenterName+'です。明日参加していただく実験についての確認のメールをお送りしています。\n'+
'\n'+
'明日 BOOKEDSCHEDULEから実験に参加していただく予定となっております。\n'+
'場所は'+experimentRoom+'です。実験時間に実験室まで直接お越しください。\n'+
'\n'+
'なお、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。\n'+
'ご不明な点などありましたら、'+experimenterMailAddress+'までご連絡ください。\n'+
'それでは明日、よろしくお願いいたします。\n'+
'\n'+
experimenterName;
// --- 定義セクション終了 ---
 
 
//見やすい形式にする関数
function beautifulDate(d,option){
  var month = d.getMonth() + 1; //返ってくる値が0~11のため、+1をする必要がある。
  var hour = d.getHours();
  var min = ('0' + d.getMinutes()).slice(-2);
  var date = d.getDate();
  var yobi = new Array("日", "月", "火", "水", "木", "金", "土")[d.getDay()];
  if (option === 'full') {
    return month + '月' + date +'日（' + yobi +'）' + hour + ':' + min;
  } else if (option === 'time') {
    return hour + ':' + min;
  } else if (option === 'month_n_date') {
    return month + '月' + date + '日';}
}
 
// sheetModifySendMail
function sendMailModifySheet(sheet, numRow, ParticipantEmail, mailText, mailTitle, operationDict){
    //MailApp.sendEmail(ParticipantEmail, mailTitle, mailText, {bcc: experimenterMailAddress});
  
    //Logger.log(typeof sheet);
    //Logger.log(sheet);
    if (typeof sheet.getSheets === 'function'){
    sheet = sheet.getSheets()[0];}
  
    if (typeof operationDict !== 'undefined'){
    for(var i = 0; i < Object.keys(operationDict).length; i++){
      var arrayKeys = Object.keys(operationDict);
      var key = arrayKeys[i];
      sheet.getRange(numRow, key).setValue(operationDict[key]);}
    }
}

var openExperimentDate = new Date(year=new Date().getFullYear(), month=openMonth, day=openDate, hour=openTime);
var closeExperimentDate = new Date(year=new Date().getFullYear(), month=closeMonth, day=closeDate, hour=closeTime);


//仮予約があった際に、カレンダーに書き込む関数
function sendToCalendar(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try{
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(1, 10).setValue('予約ステータス(111=予約完了,222=既参加,333=定員超過)'); //,444=二重登録,999=予約キャンセル)');
    sheet.getRange(1, 11).setValue('連絡したか');
    sheet.getRange(1, 12).setValue('リマインド日時');
    sheet.getRange(1, 13).setValue('リマインドしたか');
    //新規仮予約された行番号を取得
    var numRow = sheet.getLastRow();
    //新規仮予約された行から名前を取得し、"仮予約：参加者名"の文字列を作る
    //ついでにアドレスも取得しておく
    var ParticipantName = sheet.getRange(numRow, 2).getValue();
    var ParticipantEmail = sheet.getRange(numRow, 8).getValue();
    var thing = "仮予約:" + ParticipantName
    //仮予約を記載するカレンダーを取得
    var cal = CalendarApp.getCalendarById(experimenterMailAddress);
    //仮予約の開始時間を取得
    var stime = new Date(sheet.getRange(numRow, 9).getValue());
    //仮予約の開始時間から終了時間を設定
    var etime = new Date(sheet.getRange(numRow, 9).getValue());
    etime.setMinutes(etime.getMinutes()+experimentLength);
    //メールに記載する、予約日時の変数を作成する
    var appo = beautifulDate(stime, 'full') + '〜' + beautifulDate(etime, 'time');
    //取得するカレンダーの時間帯を決定する。
    var opening = new Date(stime); //予約時刻の取得
    opening.setHours(openTime, 0, 0);
    var closing = new Date(stime);
    closing.setHours(closeTime, 0, 0);
    //重複の確認
    var allEvents = cal.getEvents(stime, etime);

    //実験実施可能時間・期間外に応募してきた場合
    if (stime < opening || closing < etime || stime < openExperimentDate || closeExperimentDate < stime) {
      var mailText = textOutoftime.replace('PARTICIPANTNAME', ParticipantName)
                              .replace('APPOINTMENT', appo);
      sendMailModifySheet(sheet, numRow, ParticipantEmail, mailText, "実験実施可能時間外です", {10:'時期間外', 11:1, 12:'N/A', 13:'N/A'}); 
    //カレンダーに既に登録された予定や予約と重複する時間に応募してきた場合
    } else if (allEvents.length > 0) {
      var mailText = textOverlap.replace('PARTICIPANTNAME', ParticipantName)
                            .replace('APPOINTMENT', appo);
      sendMailModifySheet(sheet, numRow, ParticipantEmail, mailText, "予約が重複しています", {10:'重複', 11:1, 12:'N/A', 13:'N/A'});
      //上記以外では、仮予約の実行
    } else {
      var mailText = textTentativeBooking.replace('PARTICIPANTNAME', ParticipantName)
                                     .replace('APPOINTMENT', appo);
      sendMailModifySheet(sheet, numRow, ParticipantEmail, mailText, "予約の確認");
      cal.createEvent(thing, stime, etime); //仮予約情報をカレンダーに作成
    }    
  }catch(exp){
    //実行に失敗した時に通知
    //MailApp.sendEmail(experimenterMailAddress, exp.message, exp.message);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(1,15).setValue(exp.message);
  }
}

//スプレッドシート上で予約を完了させ、メール送信及びカレンダーへの書き込みを行う関数
function updateCalendar(e){
  try{
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
    //予約された日時（見やすい形式）
    var reservedAppo = beautifulDate(stime, 'full');

    //111:予約完了
    if(activeCell.getColumn()==10 && activeCell.getValue()== 111 && sheet.getRange(activeCellRow,11).getValue() !=1){
        // 予約イベントを一旦削除
        var reserve = cal.getEvents(stime, etime);
        for (var i = 0; i < reserve.length; i++) {
            if (reserve[i].getTitle() == "仮予約:" + ParticipantName) {
                reserve[i].deleteEvent();
            }
        }
        cal.createEvent(thing, stime, etime); //予約確定情報をカレンダーに追加
        
      //土日に予約した場合
      if(stime.getDay() == 0 || stime.getDay() == 6){
      var mailText = textHolidayBookingDone.replace('PARTICIPANTNAME', ParticipantName)
                                           .replace('RESERVEAPPO', reservedAppo);
      //平日に予約した場合
      }else{
      var mailText = textWeekdayBookingDone.replace('PARTICIPANTNAME', ParticipantName)
                                           .replace('RESERVEAPPO', reservedAppo);
      }
      //リマインダーのための設定をする
      var reminder = new Date(sheet.getRange(activeCellRow, 9).getValue());
      reminder.setDate(reminder.getDate() - 1); //reminderの時刻を予約時間の1日前に設定する。
      sheet.getRange(activeCellRow, 12).setValue(reminder);
      var time = new Date(); //現在時刻の取得
      time.setHours(19); //19時に設定
      //予約を完了させた日の19時にreminderの時刻が達していない場合、"送信準備"というコードを指定のセルに入力する
      if (reminder > time) {
          var reminderStatus = "送信準備";
      } else {
          var reminderStatus = "前日予約のため省略";
      }
      sendMailModifySheet(sheet, activeCellRow, ParticipantEmail, mailText, "実験予約完了いたしました", {11:1, 13:reminderStatus});
      }
  
    //222:以前実験に参加したことがあり参加を断るの場合
    else if (activeCell.getColumn() == 10 && activeCell.getValue() == 222 && sheet.getRange(activeCellRow, 11).getValue() != 1) {
        // 予約イベントを削除
        var reserve = cal.getEvents(stime, etime);
        for (var i = 0; i < reserve.length; i++) {
            if (reserve[i].getTitle() == "仮予約:" + ParticipantName) {
                reserve[i].deleteEvent();
            }
        }
      var mailText = textAlreadyDone.replace('PARTICIPANTNAME', ParticipantName);
      sendMailModifySheet(sheet, activeCellRow, ParticipantEmail, mailText, "以前に実験にご参加いただいたことがあります",
        {11:1, 12:'N/A', 13:'N/A'});
    }

    //333:応募人数を超過した場合
    else if (activeCell.getColumn() == 10 && activeCell.getValue() == 333 && sheet.getRange(activeCellRow, 11).getValue() != 1) {
        // 予約イベントを削除
        var reserve = cal.getEvents(stime, etime);
        for (var i = 0; i < reserve.length; i++) {
            if (reserve[i].getTitle() == "仮予約:" + ParticipantName) {
                reserve[i].deleteEvent();
            }
        }
      var mailText = textReachedCapacity.replace('PARTICIPANTNAME', ParticipantName);
        sendMailModifySheet(sheet, activeCellRow, ParticipantEmail, mailText, "定員に達してしまいました", {11:1, 12:'N/A',13:'N/A'});
    }
  }catch(exp){
    //実行に失敗した時に通知
    MailApp.sendEmail(experimenterMailAddress, exp.message, exp.message);
    Logger.log(exp.message);
  }
}

//リマインダーを実行する関数
function sendReminders(e) {
  try {
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
                  var ParticipantName = data[row][1]; //被験者の名前
                  var week = data[row][8].getDay();
                  //休日（前半,後半は平日）に予約した場合のメール本文
                  if (week == 0 || week == 6) {
                      var mailText = textReminderHoliday.replace('PARTICIPANTNAME', ParticipantName)
                                                        .replace('BOOKEDSCHEDULE', beautifulDate(stime,'time'));
                  } else {
                      var mailText = textReminderWeekday.replace('PARTICIPANTNAME', ParticipantName)
                                                        .replace('BOOKEDSCHEDULE', beautifulDate(stime,'time'));
                  }

                  //参加者にメールを送る
                  var ParticipantEmail = data[row][7]
                  sendMailModifySheet(sheet, (row+1), ParticipantEmail, mailText, "明日実施の心理学実験のリマインダー",
                    {13:'送信済み'});//row + 1 needed beacuse Spredsheet starts from 1 while list in js starts from 0.
              }
          }
      }
  } catch (exp) {
      //実行に失敗した時に通知
      MailApp.sendEmail(experimenterMailAddress, exp.message, exp.message);
      Logger.log(exp.message);
  }
}
