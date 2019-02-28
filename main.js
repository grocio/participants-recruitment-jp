/*誰か綺麗に書きなおしてくれるとありがたい！
それぞれの関数をトリガーを設定することを忘れないように！！
sendToCalendar:スプレッドシートから -> フォーム送信時
updateCalendar:スプレッドシートから -> 値の変更
sendReminders:時間主導型 -> 日タイマー -> 午後7時〜8時
各関数の最初の変数の定義を確認・変更してください*/

/*
1.使用するスプレッドシートのアドレスとフォームの項目数を記入してください。
2.上記メニューにある「関数を選択」から「init」を選択し，実行してください。
  メールの本文に使用される文章などを設定するためのシート（setting）が生成されます。
  デフォルトの文章やtriggerを変更することもできます。また，14行目以降に自作のtriggerとメール用の文章を足すことができます。
*/
var spreadSheetURL = 'https://docs.google.com/spreadsheets/d/1dAHFn6t9OWpToDKWCsqinUMC8qIsBYTsHckjd-xSTWE/edit#gid=1136118204';
var lastColumn = 10;// googleフォームの項目数

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

// スプレッドシートからメールの内容を取得する関数
function getMailContents(spreadsheet, trigger, weekend){
  var setting = spreadsheet.getSheetByName('setting');
  var contents = setting.getRange(7,1,setting.getLastRow()-6,4).getValues();
  Logger.log(contents);
  var contentsDict = {};
  for (var i = 0; i < contents.length; i++){
    if (contents[i][0] === trigger){
      contentsDict['subject'] = contents[i][1];
      if (weekend){ //もし週末指定があったら
        contentsDict['body'] = contents[i][3];
      }
      else{
        contentsDict['body'] = contents[i][2];
      }
    }
  }
  //Logger.log(contentsDict);
  return contentsDict;
}

// 本文中の変数を取得する関数
function getBodyVariables(body){
  var variables = body.match(/[A-Za-z]+/g);
  return variables;
}

function replaceVariables(body, variables, dict){
  for (var i = 0; i < variables.length; i++){
    var variable = variables[i];
    var body = body.replace(variable,dict[variable]);
   // Logger.log(body);
  }
  return body;
}

// 実験の情報を取得する関数
function getExpInfo(spreadsheet){
  var setting = spreadsheet.getSheetByName('setting');
  var expInfo = setting.getRange(1,1,5,2).getValues();
  var expInfoDict = {};
  // 取得した配列を連想配列に変換する
  for (var i = 0; i < expInfo.length; i++){
    expInfoDict[expInfo[i][0]] = expInfo[i][1];
  }
  return expInfoDict;
}

// sheetModifySendMail
function sendMailModifySheet(sheet, numRow, ParticipantEmail, experimenterMailAddress, mailText, mailTitle, operationDict){
  MailApp.sendEmail(ParticipantEmail, mailTitle, mailText, {bcc: experimenterMailAddress});

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
// 希望日時をdate型に変換する関数
function expDatetime(sheet, row){
  var start = new Date(sheet.getRange(row, 9).getValue());
  var end = new Date(sheet.getRange(row, 9).getValue());
  //Logger.log(start)
  var T = sheet.getRange(row, 10).getValue(); // この段階では文字列
  var T_li = T.replace(/\s+/g, '').split('~'); //空白を除去し，~で分けて要素２の配列に
  var sHM = T_li[0].split(':');
  var eHM = T_li[1].split(':');
  start.setHours(sHM[0],sHM[1]);
  end.setHours(eHM[0],eHM[1]);
  return [start, end];
}

// memberシートからbccアドレスを追加する関数
function getbccAddresses(firstaddress, row){
  var spreadsheet = SpreadsheetApp.openByUrl(spreadSheetURL);
  var activeSheet = spreadsheet.getActiveSheet();
  Logger.log(activeSheet);
  var members = spreadsheet.getSheetByName('member');
  var members_mtrx = members.getDataRange().getValues();
  Logger.log(members_mtrx);
  var lastCol = activeSheet.getLastColumn();
  var charge = activeSheet.getRange(row,lastCol).getValue();
  Logger.log(charge);
  var bccAddress = [firstaddress];
  for (var i = 0; i < members_mtrx.length; i++){
    if (members_mtrx[i][0] == charge){
      bccAddress.push(members_mtrx[i][2]);
    }
  }
  Logger.log(bccAddress);
  return bccAddress.join(',');
}

//仮予約があった際に、カレンダーに書き込む関数
function sendToCalendar(e){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //var url = 'https://docs.google.com/spreadsheets/d/1FlPOHi3pJM-Z66C9dpiugy_dEgdeWoE6xquj35t2FGA/edit#gid=51117778';
  //var spreadsheet = SpreadsheetApp.openByUrl(url); //openByIdを使うと時限式のトリガーが正常に働くらしい
  //var sheet = spreadsheet.getSheetByName('フォームの回答 1');
  try{
    //実験情報の取得
    var expInfo = getExpInfo(spreadsheet);

    //新規仮予約された行番号を取得
    var numRow = sheet.getLastRow();
    //新規仮予約された行から名前を取得し、"仮予約：参加者名"の文字列を作る
    //ついでにアドレスも取得しておく
    var ParticipantName = sheet.getRange(numRow, 2).getValue();
    var ParticipantEmail = sheet.getRange(numRow, 8).getValue();
    var thing = "仮予約:" + ParticipantName
    //仮予約を記載するカレンダーを取得
    var cal = CalendarApp.getCalendarById(expInfo['experimenterMailAddress']);
    var prefDT = expDatetime(sheet, numRow);
    //仮予約の開始時間を取得
    var stime = prefDT[0];
    //仮予約の開始時間から終了時間を設定
    var etime = prefDT[1];
    //メールに記載する、予約日時の変数を作成する
    var appo = beautifulDate(stime, 'full') + '〜' + beautifulDate(etime, 'time');
    expInfo['PARTICIPANTNAME'] = ParticipantName;
    expInfo['WHEN'] = appo;
    Logger.log(expInfo);
    //取得するカレンダーの時間帯を決定する。
    //var opening = new Date(stime); //予約時刻の取得
    //opening.setHours(openTime, 0, 0);
    //var closing = new Date(stime);
    //closing.setHours(closeTime, 0, 0);
    //重複の確認
    var allEvents = cal.getEvents(stime, etime);

    //実験実施可能時間・期間外に応募してきた場合
    //if (stime < opening || closing < etime || stime < openExperimentDate || closeExperimentDate < etime) {
    //  var mailText = textOutoftime.replace('PARTICIPANTNAME', ParticipantName)
    //  .replace('APPOINTMENT', appo);
    //  sendMailModifySheet(sheet, numRow, ParticipantEmail, mailText, "実験実施可能時間外です", {10:'時期間外', 11:1, 12:'N/A', 13:'N/A'});
      //カレンダーに既に登録された予定や予約と重複する時間に応募してきた場合
    //} else
    if (allEvents.length > 0) {
      var trigger = '重複';
      var operationDict = {11:trigger, 12:1, 13:'N/A', 14:'N/A'};
    } else {
      var trigger = '仮予約'
      var operationDict = {11:'', 12:'', 13:'', 14:''};
      //var mailText = textTentativeBooking.replace('PARTICIPANTNAME', ParticipantName)
      //.replace('APPOINTMENT', appo);
      //sendMailModifySheet(sheet, numRow, ParticipantEmail, mailText, "予約の確認");
      cal.createEvent(thing, stime, etime); //仮予約情報をカレンダーに作成
    }
    var contents = getMailContents(spreadsheet, trigger, false);
    Logger.log(contents);
    var variables = getBodyVariables(contents['body']);
    Logger.log(variables);
    var mailText = replaceVariables(contents['body'],variables,expInfo);
    Logger.log(mailText);
    //var mailText = textOverlap.replace('PARTICIPANTNAME', ParticipantName)
    //.replace('APPOINTMENT', appo);
    sendMailModifySheet(sheet, numRow, ParticipantEmail, expInfo['experimenterMailAddress'], mailText, contents['subject'], operationDict);
    //上記以外では、仮予約の実行

  }catch(exp){
    //実行に失敗した時に通知
    MailApp.sendEmail(expInfo['experimenterMailAddress'], exp.message, exp.message);
    Logger.log(exp.message);
  }
}

//スプレッドシート上で予約を完了させ、メール送信及びカレンダーへの書き込みを行う関数
function updateCalendar(e){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var expInfo = getExpInfo(spreadsheet);
  try{
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getSheetId() === spreadsheet.getSheets()[0].getSheetId()){//設定用のシートをいじっても何も起きないようにする
      //アクティブセル（値の変更があったセル）を取得
      var activeCell = sheet.getActiveCell();
      var activeCellRow = activeCell.getRow();
      //予約を記載するカレンダーを取得
      var cal = CalendarApp.getCalendarById(expInfo['experimenterMailAddress']);
      var prefDT = expDatetime(sheet, activeCellRow);
      //予約の開始時間を取得
      var stime = prefDT[0];
      //予約の開始時間から終了時間を設定
      var etime = prefDT[1];
      //変更した行から名前を取得し、"予約完了：参加者名(ふりがな)"の文字列を作る
      var ParticipantName = sheet.getRange(activeCellRow, 2).getValue();
      var thing = "予約完了:" + ParticipantName +'('+sheet.getRange(activeCellRow, 3).getValue()+')';
      //ついでに被験者のメールアドレスも取得
      var ParticipantEmail = sheet.getRange(activeCellRow, 8).getValue();
      //予約された日時（見やすい形式）
      var reservedAppo = beautifulDate(stime, 'full');
      expInfo['PARTICIPANTNAME'] = ParticipantName;
      expInfo['WHEN'] = reservedAppo;
      var activeColumn = activeCell.getColumn();
      var activeColname = sheet.getRange(1,activeColumn).getValue();
      var trigger = activeCell.getValue();
      var regex = /[0-9]+/;
      if (regex.test(trigger)){
        if (activeColname === '予約ステータス' && sheet.getRange(activeCellRow,activeColumn).getValue() !== 1){
          // まず予約イベントを削除する
          var reserve = cal.getEvents(stime, etime);
          for (var i = 0; i < reserve.length; i++) {
            if (reserve[i].getTitle() == "仮予約:" + ParticipantName) {
              reserve[i].deleteEvent();
            }
          }

          if (trigger === 111) {
            cal.createEvent(thing, stime, etime); //予約確定情報をカレンダーに追加
            //リマインダーのための設定をする
            var reminder = new Date(stime);
            reminder.setDate(stime.getDate() - 1); //reminderの時刻を予約時間の1日前に設定する。
            var time = new Date(); //現在時刻の取得
            time.setHours(19); //19時に設定
            //予約を完了させた日の19時にreminderの時刻が達していない場合、"送信準備"というコードを指定のセルに入力する
            if (reminder > time) {
              var reminderStatus = "送信準備";
            } else {
              var reminderStatus = "前日予約のため省略";
            }
            var insertValues = [1, reminder, reminderStatus];
            if (stime.getDay()==0 || stime.getDay()==6){
              var contents = getMailContents(spreadsheet, trigger, true);
            }
            else{
              var contents = getMailContents(spreadsheet, trigger, false);
            }
          }
          // triggerが111以外のとき
          else {
            var insertValues = [1,'N/A','N/A'];
            var contents = getMailContents(spreadsheet, trigger, false);
          }

          var variables = getBodyVariables(contents['body']);
          var mailText = replaceVariables(contents['body'],variables,expInfo);
          var operationDict = {};
          for (i = 0; i < insertValues.length; i++){
            operationDict[activeColumn+i+1] = insertValues[i];
          }
          // メールの送信
          sendMailModifySheet(sheet, activeCellRow, ParticipantEmail, getbccAddresses(expInfo['experimenterMailAddress'],activeCellRow), mailText, contents['subject'], operationDict);
        }
      }
    }
  }catch(exp){
    //実行に失敗した時に通知
    Logger.log(exp)
    //MailApp.sendEmail(expInfo['experimenterMailAddress'], exp.message, exp.message);
  }
}

//リマインダーを実行する関数
function sendReminders(e) {
  var spreadsheet = SpreadsheetApp.openByUrl(spreadSheetURL); //openByIdを使うと時限式のトリガーが正常に働くらしい
  var expInfo = getExpInfo(spreadsheet);
  try {
    var sheet = spreadsheet.getSheets()[0];
    var data = sheet.getDataRange().getValues(); //たぶんシート全体のデータを取得する
    var time = new Date().getTime(); //現在時刻の取得
    // スプレッドシートを1列ずつ参照し、該当する被験者を探していく。
    for (var row = 1; row < data.length; row++) {
      //ステータスが送信準備になっていることを確認する
      if (data[row][lastColumn+3] == "送信準備") {
        var reminder = data[row][lastColumn+2];
        // もし現在時刻がリマインド日時を過ぎていたならメールを送信
        if ((reminder != "") && (reminder.getTime() <= time)) {
          // メールの本文の内容を作成するための要素を定義
          var ParticipantName = data[row][1]; //被験者の名前

          var trigger = 'リマインダー';
          //予約の開始時間を取得
          var prefDT = expDatetime(sheet, row+1);
          Logger.log(prefDT);
          var stime = prefDT[0];
          var reservedAppo = Utilities.formatDate(stime, 'Asia/Tokyo', 'HH:mm');
          var week = stime.getDay();
          expInfo['PARTICIPANTNAME'] = ParticipantName;
          expInfo['WHEN'] = reservedAppo;
          //休日（前半,後半は平日）に予約した場合のメール本文
          if (stime.getDay()==0 || stime.getDay()==6){
            var contents = getMailContents(spreadsheet, trigger, true);
          }
          else{
            var contents = getMailContents(spreadsheet, trigger, false);
          }
          Logger.log(contents)
          var variables = getBodyVariables(contents['body']);
          var mailText = replaceVariables(contents['body'],variables,expInfo);
          var operationDict = {};
          var insertValues = ['送信済み'];
          for (i = 0; i < insertValues.length; i++){
            operationDict[lastColumn+4] = insertValues[i];
          }

          //参加者にメールを送る
          var ParticipantEmail = data[row][7];
          sendMailModifySheet(sheet, (row+1), ParticipantEmail, getbccAddresses(expInfo['experimenterMailAddress'], row+1),
                              mailText, contents['subject'], operationDict);//row + 1 needed beacuse Spredsheet starts from 1 while list in js starts from 0.
        }
      }
    }
  } catch (exp) {
    //実行に失敗した時に通知
    Logger.log(exp.message);
    //MailApp.sendEmail(expInfo['experimenterMailAddress'], exp.message, exp.message);
  }
}
// 設定用のシートおよびその見本を最初に作る関数
function init(){
  //var url = 'https://docs.google.com/spreadsheets/d/1FlPOHi3pJM-Z66C9dpiugy_dEgdeWoE6xquj35t2FGA/edit#gid=51117778';
  var spreadsheet = SpreadsheetApp.openByUrl(spreadSheetURL); //openByIdを使うと時限式のトリガーが正常に働くらしい
  var sheets = spreadsheet.getSheets();
  //有効なGooglesプレッドシートを開く
  sheets[0].getRange(1, lastColumn+1).setValue('予約ステータス'); //,444=二重登録,999=予約キャンセル)');
  sheets[0].getRange(1, lastColumn+2).setValue('連絡したか');
  sheets[0].getRange(1, lastColumn+3).setValue('リマインド日時');
  sheets[0].getRange(1, lastColumn+4).setValue('リマインドしたか');
  if (sheets.length < 2){

    spreadsheet.insertSheet('setting');

    var setting = spreadsheet.getSheetByName('setting');
    var expInfoDict = {'experimenterName':'実験太郎','experimenterMailAddress':'hogehoge@gmail.com','experimenterPhone':'xxx-xxx-xxx',
                        'experimentRoom':'実施場所','experimentLength':'実験の所要時間'};
    // 実験情報
    var rowN = 0;
    for (key in expInfoDict){
      setting.getRange(rowN+1,1).setValue(key);
      setting.getRange(rowN+1,2).setValue(expInfoDict[key]);
      rowN++;
    }
    Logger.log(Object.keys(expInfoDict).length);
    // メールの内容や起動のトリガーを指定する
    var colnames = ['trigger','subject','body[weekday]','body[weekend]'];
    var tempLastRow = setting.getLastRow();
    for (i = 0; i < colnames.length; i++){
      setting.getRange(tempLastRow+1, i+1).setValue(colnames[i]);
    }
    // successful 仮予約
    var textTentativeBooking = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                                '予約の確認メールを自動で送信しております。\n','WHEN','で予約を受け付けました（まだ確定はしていません。)',
                                '後日、予約完了のメールを送信いたします。','もし日時の変更等がある場合は experimenterMailAddress までご連絡ください。',
                                'どうぞよろしくお願いいたします。\n','experimenterName'];
    // --- Failed 仮予約シリーズ ---
    // 時間外・期間外
    var textOutoftime = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。',
                         'この度は心理学実験への応募ありがとうございました。','申し訳ありませんが、ご希望いただいた','WHEN',
                         'は実験実施可能時間（openTime時〜closeTime時）外または、実施期間（openMonth月openDate日〜closeMonth月closeDate日）外です。',
                         'お手数ですが、もう一度登録し直していただきますようお願いします。\n',
                         'experimenterName'];
    // 重複
    var textOverlap = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                       '申し訳ありませんが、ご希望いただいた','WHEN',
                       'にはすでに予約（予定）が入っており（タッチの差で他の方が予約をされた可能性もあります）、実験を実施することができません。',
                       'お手数ですが、もう一度別の日時で登録し直していただきますようお願いします。\n','experimenterName'];
    // --- Successful Booking ---
    // 予約完了テキスト(平日)
    var textWeekdayBookingDone = ['PARTICIPANTNAME 様\n','この度は心理学実験への応募ありがとうございました。',
                                  'WHENからの心理学実験の予約が完了しましたのでメールいたします。',
                                  '場所はexperimentRoomです。当日は直接お越しください。',
                                  'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
                                  '当日もよろしくお願いいたします。\n',
                                  '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)','当日の連絡はexperimenterPhoneまでお願いいたします。'];
    // 予約完了テキスト(休日)
    var textHolidayBookingDone = ['PARTICIPANTNAME 様\n','この度は心理学実験への応募ありがとうございました。',
                                  'WHENからの心理学実験の予約が完了しましたのでメールいたします。',
                                  '場所はexperimentRoomです。休日は教育学部棟玄関の鍵がかかっており、外から入ることができません。実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。',
                                  'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。','当日もよろしくお願いいたします。\n',
                                  '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)','当日の連絡はexperimenterPhoneまでお願いいたします。'];
    // --- Rejected Booking ---
    // 既参加
    var textAlreadyDone = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                           '大変申し訳ありませんが、以前実施した同様の実験にご参加いただいており、今回の実験にはご参加いただけません。ご了承ください。\n',
                           'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。','今後ともよろしくお願いします。\n','experimenterName'];
    // 定員オーバー
    var textReachedCapacity = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                               '大変申し訳ありませんが、応募いただいた段階ですでに募集人数の定員に達していたため、実験に参加していただくことができません。ご了承ください。\n',
                               '今後、次の実験を実施する際に再度応募していただけると幸いです。','ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
                               '今後ともよろしくお願いいたします。\n','experimenterName'];

    // --- Reminders ---
    // リマインダー(平日)
    var textReminderWeekday =['PARTICIPANTNAME 様\n','実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
                              '明日 WHENから実験に参加していただく予定となっております。','場所はexperimentRoomです。実験時間に実験室まで直接お越しください。\n',
                              'なお、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。','ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
                              'それでは明日、よろしくお願いいたします。\n','experimenterName'];
    // リマインダー(休日)
    var textReminderHoliday =['PARTICIPANTNAME 様\n','実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
                              '明日 WHENから実験に参加していただく予定となっております。','場所はexperimentRoomです。\n',
                              'なお、明日は休日のため教育学部棟玄関の鍵がかかっており、外から入ることができません。実験者が実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n',
                              'また、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。',
                              'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。','それでは明日、よろしくお願いいたします。\n','experimenterName'];

    var sample = {"trigger":['仮予約','時期間外','重複',111,222,333,'リマインダー'],
                  "subject":['予約の確認','実験実施可能時間外です','予約が重複しています',"実験予約が完了いたしました","以前に実験にご参加いただいたことがあります","定員に達してしまいました","明日実施の心理学実験のリマインダー"],
                  "body_wd":[textTentativeBooking,textOutoftime,textOverlap,textWeekdayBookingDone,textAlreadyDone,textReachedCapacity,textReminderWeekday],
                  "body_ho":[['not for use'],['not for use'],['not for use'],textHolidayBookingDone,['not for use'],['not for use'],textReminderHoliday]};
    var colN = 1;
    var tempLastRow = setting.getLastRow()
    for (key in sample) {
      Logger.log(key);
      for (var i = 0; i < sample[key].length; i++ ){
        if (key === 'body_wd' || key === 'body_ho'){
          var item = sample[key][i].join('\n');
        }
        else {
          var item = sample[key][i]
        }
        setting.getRange(tempLastRow+i+1, colN).setValue(item);
      }
      colN++
    }
  }
}
