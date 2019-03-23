var type = 1; // 1: 自由回答, 2: 選択式 どちらかの半角数字を入れてください。

function init() { // デフォルトの設定を作成する関数
  setting(type);
}

var ss = SpreadsheetApp.getActiveSpreadsheet(); // spreadsheet

try{
  if (ss.getSheets().length > 1) {
    var expInfo = getExpInfo();
    var answers = ss.getSheets()[0];
    var colParName = Number(expInfo['colParName']);
    var colParNameKana = Number(expInfo['colParNameKana']);
    var colCharge = answers.getLastColumn();
    var colReminded = colCharge - 1;
    var colRemindDate = colCharge - 2;
    var colMailed = colCharge - 3;
    var colStatus = colCharge - 4;
    if (type == 2) {
      colExpTime = Number(expInfo['colExpTime']);
    }
    var colExpDate = Number(expInfo['colExpDate']);
    var colAddress = Number(expInfo['colAddress']);
  }
} catch (err) {
  Logger.log(err.message);
}

function zenToHan(str) {
  if (typeof str == "string") {
    return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) { // 全角を半角に変換
      return String.fromCharCode(s.charCodeAt(0) - 65248); // 10進数の場合
    });
  } else {
    return str
  }
}

// 実験の設定を取得する関数
function getExpInfo() {
  var sheet = ss.getSheetByName('設定');
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var expInfo = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
  var expInfoDict = {};
  // 取得した配列を連想配列に変換する
  for (var i = 0; i < expInfo.length; i++) {
    expInfoDict[expInfo[i][0]] = zenToHan(expInfo[i][1]); // 念の為;
  }
  return expInfoDict;
}

// 希望日時を取得しdate型に変換する関数
function getExpDateTime(sheet, row) {
  if (type == 1) {
    var from = new Date(sheet.getRange(row, colExpDate).getValue());
    var to = new Date(from);
    var expLength = expInfo['experimentLength'];
    to.setMinutes(from.getMinutes() + expLength);
  } else {
    var from = new Date();
    var date = sheet.getRange(row, colExpDate).getValue();
    date = zenToHan(date);
    var dateInfo = date.match(/\d+/g);
    if (dateInfo.length == 3) { //年月日なら
      from.setFullYear(dateInfo[0], dateInfo[1], dateInfo[2]);
    } else if (dateInfo.length == 2) { //月日なら
      from.setMonth(dateInfo[0], dateInfo[1]);
    } else if (dateInfo.length == 1) { //日なら
      from.setDate(dateInfo[0]);
    }
    var to = new Date(from);
    var time = sheet.getRange(row, colExpTime).getValue(); // この段階では文字列
    time = zenToHan(time);
    var FromTo = datetime.match(/\d+/g); //空白を除去し，~で分けて要素２の配列に
    from.setHours(FromTo[0],FromTo[1]);
    if (FromTo.length == 4) {
      to.setHours(FromTo[2],FromTo[3]);
    } else if (FromTo.length == 2) {
      var expLength = expInfo['experimentLength'];
      to.setMinutes(from.getMinutes() + expLength);
    }
  }
  return {'from': from, 'to': to};
}

// メールの本文を作成する関数
function makeMailBody(body) {
  for (key in expInfo) {
    var regex = new RegExp(key,'g');
    var body = body.replace(regex, expInfo[key]);
  }
  return body;
}

// スプレッドシートからメールのテンプレートを取得する関数
function getTemplate(trigger, time) {
  var sheet = ss.getSheetByName('テンプレート');
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var contents = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var contentsDict = {};
  for (var i = 0; i < contents.length; i++){
    if (contents[i][0] === trigger){
      contentsDict['subject'] = contents[i][2];
      var changeByDay = contents[i][1];
      if (changeByDay === 1 && (time.getDay()==0 || time.getDay()==6)){ //もし週末なら
        contentsDict['body'] = contents[i][4];
      }
      else{
        contentsDict['body'] = contents[i][3];
      }
    }
  }
  return contentsDict;
}

function myFormatDate(datetime, pattern) {
  return Utilities.formatDate(datetime, 'JST', pattern);
}

// mailの内容を作成する関数
function sendEmail(name, address, from, to, trigger, row) {
  //メールに記載する、予約日時の変数を作成する
  var yobi = new Array("日", "月", "火", "水", "木", "金", "土")[from.getDay()];
  expInfo['participantName'] = name;
  expInfo['expDate'] = myFormatDate(from, 'MM/dd') + "（"+ yobi +"）";
  expInfo['fromWhen'] = myFormatDate(from, 'HH:mm');
  expInfo['toWhen'] = myFormatDate(to, 'HH:mm');
  expInfo['openDate'] = myFormatDate(expInfo['openDate'], 'yyyy/MM/dd');
  expInfo['closeDate'] = myFormatDate(expInfo['closeDate'], 'yyyy/MM/dd');
  var contents = getTemplate(trigger, from);
  var subject = makeMailBody(contents['subject']);
  var body = makeMailBody(contents['body']);
  var bccAddresses = getbccAddresses(row)
  MailApp.sendEmail(address, subject, body, {bcc: bccAddresses});
}

function modifySheet(sheet, numRow, columns, values) {
  for (var i = 0; i < columns.length; i++){
    sheet.getRange(numRow, columns[i]).setValue(values[i]);
  }
}

function getMemberInfo() {
  var sheet = ss.getSheetByName('メンバー');
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var memberInfo = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var memberInfoDict = {};
  // 取得した配列を連想配列に変換する
  for (var i = 0; i < memberInfo.length; i++) {
    key = zenToHan(memberInfo[i][0]);
    memberInfoDict[key] = zenToHan(memberInfo[i][2]); // 念の為;
  }
  return memberInfoDict;
}

// memberシートからbccアドレスを追加する関数
function getbccAddresses(row) {
  var activeSheet = ss.getActiveSheet();
  var memberInfo = getMemberInfo();
  var charges = activeSheet.getRange(row, colCharge).getValue();
  charges = zenToHan(charges);
  var bccAddress = [expInfo['experimenterMailAddress']];
  if (typeof charges == 'number') {// 一人だけが指定されている場合
    bccAddress.push(memberInfo[charges]);
  } else if (typeof charges == 'string') {// 複数人が指定されている場合
    if (charges.length > 0) {
      var chargeIDs = charges.match(/\d+/g);
      for (var i = 0; i < chargeIDs.length; i++) {
        var chargeID = chargeIDs[i];
        bccAddress.push(memberInfo[chargeID]);
      }
    }
  }
  return bccAddress.join(',');
}

//仮予約があった際に、カレンダーに書き込む関数
function sendToCalendar() {
  var sheet = ss.getActiveSheet();
  try{
    //実験情報の取得
    var numRow = sheet.getActiveRange().getRow() // 新規仮予約された行番号を取得。active行だけを取得するようにする。手動で参加者情報を追加しても対応できる。
    var participantName = sheet.getRange(numRow, colParName).getValue(); //新規仮予約された行から名前を取得し、"仮予約：参加者名"の文字列を作る
    var eventTitle = "仮予約:" + participantName;
    //重複の確認
    var cal = CalendarApp.getCalendarById(expInfo['experimenterMailAddress']); //仮予約を記載するカレンダーを取得
    var expDT = getExpDateTime(sheet, numRow);
    var from = expDT['from']; //仮予約の開始時間を取得
    var to = expDT['to'];//仮予約の開始時間から終了時間を設定
    var openTime = expInfo['openTime'];
    var closeTime = expInfo['closeTime'];
    var openDate = expInfo['openDate'];
    var closeDate = expInfo['closeDate'];

    var allEvents = cal.getEvents(from, to);
    if (allEvents.length > 0) {
      var trigger = '重複';
      var values = [trigger, 1, 'N/A', 'N/A'];
    } else if (from.getHours() < openTime || to.getHours() >= closeTime || from < openDate || from > closeDate) {
      var trigger = '時間外';
      var values = [trigger, 1, 'N/A', 'N/A'];
    } else {
      var trigger = '仮予約'
      var values = ['', '', '', ''];
      cal.createEvent(eventTitle, from, to); //仮予約情報をカレンダーに作成
    }
    var ParticipantEmail = sheet.getRange(numRow, colAddress).getValue();
    sendEmail(participantName, ParticipantEmail, from, to, trigger, numRow);
    modifySheet(sheet, numRow, [colStatus, colMailed, colRemindDate, colReminded],  values);
  } catch(err) {
    //実行に失敗した時に通知
    var fb = "[line " + err.lineNumber + "] " +err.message;
    Logger.log(fb);
    MailApp.sendEmail(expInfo['experimenterMailAddress'], "エラーが発生しました", fb);
  }
}

// スプレッドシート上で予約を完了させ、メール送信及びカレンダーへの書き込みを行う関数
function updateCalendar() {
  try {
    //有効なGooglesプレッドシートを開く
    var sheet = ss.getActiveSheet();
    if (sheet.getSheetId() === ss.getSheets()[0].getSheetId()){ //設定用のシートをいじっても何も起きないようにする
      //アクティブセル（値の変更があったセル）を取得
      var activeCell = sheet.getActiveCell();
      var activeRow = activeCell.getRow();

      var expDT = getExpDateTime(sheet, activeRow);
      var from = expDT['from']; //予約の開始時間を取得
      var to = expDT['to'];//予約の開始時間から終了時間を設定

      //予約された日時（見やすい形式）
      "要修正"
      var participantName = sheet.getRange(activeRow, colParName).getValue();
      var activeColumn = activeCell.getColumn();
      var activeColname = sheet.getRange(1,activeColumn).getValue();
      var trigger = activeCell.getValue();
      var regex = /[0-9]+/;
      if (regex.test(trigger)){
        var cal = CalendarApp.getCalendarById(expInfo['experimenterMailAddress']); //予約を記載するカレンダーを取得
        if (activeColname === '予約ステータス' && sheet.getRange(activeRow,activeColumn).getValue() !== 1){
          // まず予約イベントを削除する
          var reserve = cal.getEvents(from, to);
          for (var i = 0; i < reserve.length; i++) {
            if (reserve[i].getTitle() == "仮予約:" + participantName) {
              reserve[i].deleteEvent();
            }
          }
          if (trigger === 111) {
            // 変更した行から名前を取得し、"予約完了：参加者名(ふりがな)"の文字列を作る
            if (colParNameKana > 0) {
              var eventTitle = "予約完了:" + participantName +'('+sheet.getRange(activeRow, colParNameKana).getValue()+')';
            } else {
              var eventTitle = "予約完了:" + participantName;
            }
            cal.createEvent(eventTitle, from, to); //予約確定情報をカレンダーに追加
            // リマインダーのための設定をする
            var remindDate = new Date(from)
            remindDate.setDate(from.getDate() - 1); //remindDateの時刻を予約時間の1日前に設定する。
            var time = new Date(); //現在時刻の取得
            time.setHours(19); //19時に設定
            // 予約を完了させた日の19時にremindDateの時刻が達していない場合、"送信準備"というコードを指定のセルに入力する
            if (remindDate > time) {
              var reminderStatus = "送信準備";
            } else {
              var reminderStatus = "直前のため省略";
            }
            var values = [1, remindDate, reminderStatus];
          }
          // triggerが111以外のとき
          else {
            var values = [1,'N/A','N/A'];
          }
          // メールの送信
          var ParticipantEmail = sheet.getRange(activeRow, colAddress).getValue();
          sendEmail(participantName, ParticipantEmail, from, to, trigger, activeRow)
          modifySheet(sheet, activeRow, [colMailed, colRemindDate, colReminded], values)
        }
      }
    }
  } catch(err) {
    //実行に失敗した時に通知
    var fb = "[line " + err.lineNumber + "] " +err.message;
    if (expInfo['experimenterMailAddress'] == "hogehoge@gmail.com") {
      fb += "\\n実験者のアドレスがデフォルトのままです。"
    } else if (err.message == "予定の開始日時は終了日時より前にしてください。") {
      fb += "\\nバージョン情報が適切でないかもしれません。\\nコード1行目のtypeが適切な数字かどうか確認してください。"
    }
    Logger.log(fb);
    Browser.msgBox("エラーが発生しました", fb, Browser.Buttons.OK);
    MailApp.sendEmail(expInfo['experimenterMailAddress'], exp.message, exp.message);
  }
}

//リマインダーを実行する関数
function sendReminders() {
  try {
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues(); //シート全体のデータを取得。2次元の配列 [行 [列]]
    var time = new Date().getTime(); //現在時刻の取得
    // スプレッドシートを1列ずつ参照し、該当する被験者を探していく。
    for (var row = 1; row < data.length; row++) { // 0行目は列名
      //ステータスが送信準備になっていることを確認する
      if (data[row][colReminded - 1] == "送信準備") {
        var reminder = data[row][colRemindDate - 1];
        // もし現在時刻がリマインド日時を過ぎていたならメールを送信
        if ((reminder != "") && (reminder.getTime() <= time)) {
          // メールの本文の内容を作成するための要素を定義
          var participantName = data[row][colParName - 1]; //被験者の名前
          //参加者にメールを送る
          var ParticipantEmail = data[row][colAddress - 1];
          var trigger = 'リマインダー';
          var expDT = getExpDateTime(sheet, row + 1);
          var from = expDT['from'];
          var to = expDT['to'];
          sendEmail(participantName, ParticipantEmail, from, to, trigger, row + 1)
          modifySheet(sheet, row + 1, [colReminded], ['送信済み']);
        }
      }
    }
  } catch (err) {
    //実行に失敗した時に通知
    var fb = "[line " + err.lineNumber + "] " +err.message;
    Logger.log(fb);
    MailApp.sendEmail(expInfo['experimenterMailAddress'], "エラーが発生しました", fb);
  }
}

// 設定用のシートおよびその見本を最初に作る関数
function setting(){
  buttons = Browser.Buttons.OK_CANCEL;
  start = true;
  var sheets = ss.getSheets();
  if (sheets.length > 1) {
    msg = "一度設定を行ったことがあるようです（シートが2枚以上あります）。\\nもう一度初期化を行いますか？\\n"
    msg += "フォームの回答が一番初めのシートでないとこれまでの情報が失われる場合があります。"
    var choice = Browser.msgBox("設定の初期化を行います", msg, buttons);
    if (choice !== "ok") {
      start = false;
    }
  }
  if (type == 1) {
    msg = "自由回答形式の設定で初期化を行います";
  } else if (type == 2) {
    msg = "選択形式の設定で初期化を行います";
  } else {
    msg = "半角数字の1か2を入力して設定の形式を選択してください";
    buttons = Browser.Buttons.OK;
    start = false;
  }
  var choice = Browser.msgBox("設定の初期化", msg, buttons);
  if (choice !== "ok") {
    start = false;
  }
  if (start) {
    setTriggers();
    setDefault(type);
    msg = "初期設定が終了しました。\\n";
    msg += "「設定」シートの太枠に囲まれた項目を適切な情報に変更してください。";
    Browser.msgBox("設定の初期化", msg, Browser.Buttons.OK);
  } else {
    Browser.msgBox("設定の初期化", "初期化はキャンセルされました", Browser.Buttons.OK);
  }
}

function setTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('sendToCalendar').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('updateCalendar').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('sendReminders').timeBased().atHour(19).nearMinute(30).everyDays(1).create();
}

function setDefault(type){
  try {
    var sheets = ss.getSheets();
    var addNewCol = true;
    if (sheets.length > 2) {
      for (i = 1; i < sheets.length; i++) {
        ss.deleteSheet(sheets[i]);
      }
      addNewCol = false;
    }
    var sheet = sheets[0];
    ss.insertSheet('設定');
    ss.insertSheet('テンプレート');
    ss.insertSheet('メンバー');
    var colNames = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    var addColNames = [['予約ステータス', '連絡したか', 'リマインド日時', 'リマインドしたか', '担当']];
    if (addNewCol) {
      var newColNames = [colNames[0].concat(addColNames[0])];
      sheet.getRange(1, 1, 1, newColNames[0].length).setValues(newColNames);
    } else {
      sheet.getRange(1, colNames[0].length - addColNames[0].length + 1, 1, addColNames[0].length).setValues(addColNames);
    }
    lastCol = sheet.getLastColumn();
    // 設定シート
    var start = new Date();
    formattedStart = myFormatDate(start, 'yyyy/MM/dd');
    var end = new Date(start); end.setDate(start.getDate() + 13);
    formattedEnd = myFormatDate(end, 'yyyy/MM/dd');
    var config = ss.getSheetByName('設定');
    var note2 = '「フォームの回答」の列番号と一致しているか確認してください（A列が1）';
    var defaultExpInfo = [['設定項目','メール本文内でのキー','値','備考'],
                          ['実験責任者名','experimenterName','実験太郎', "実験責任者の名前を記入してください"],
                          ['実験責任者のGmailアドレス','experimenterMailAddress','hogehoge@gmail.com', "実験用のGmailアドレスを記入してください"],
                          ['実験責任者の電話番号','experimenterPhone','xxx-xxx-xxx', "電話番号を記入してください"],
                          ['実験の実施場所','experimentRoom','実施場所',"実験の実施場所を記入してください"],
                          ['実験の所要時間','experimentLength', 60, '実験の所要時間を記入してください。2列目は変更しないでください'],
                          ['実験開始可能時間','openTime', 9, '何時から実験できるかを記入してください（24時間表記）'],
                          ['実験最終時間','closeTime', 19,'何時まで実験可能かを記入してください（24時間表記）'],
                          ['実験開始日','openDate', formattedStart, '実験を開始する日付を記入してください（年/月/日で表記）'],
                          ['実験最終日','closeDate', formattedEnd, '実験の終了予定日を記入してください（年/月/日で表記）'],
                          ['参加者名の列番号','colParName', 2, note2],
                          ['ふりがなの列番号','colParNameKana', 3, note2 + 'もし利用しない場合は-1を入力してください。'],
                          ['参加者アドレスの列番号','colAddress',lastCol - 7, note2]];

    var verChoice = [['希望日の列番号','colExpDate', lastCol - 6, note2],
                     ['希望時間の列番号','colExpTime', lastCol - 5, note2]];

    var verAnswer = [['希望日時の列番号','colExpDate', lastCol - 6, note2]];

    if (type == 1) {
      defaultExpInfo = defaultExpInfo.concat(verAnswer)
    } else {
      defaultExpInfo = defaultExpInfo.concat(verChoice)
    }

    var configNRow = defaultExpInfo.length;
    var configNCol = defaultExpInfo[0].length;
    config.getRange(1, 1, configNRow, configNCol).setValues(defaultExpInfo);


    // メールのテンプレート用シート
    var template = ss.getSheetByName('テンプレート');

    // successful 仮予約
    var TentativeBooking = ['participantName 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                            '予約の確認メールを自動で送信しております。\n','expDatefromWhen〜toWhen','で予約を受け付けました（まだ確定はしていません。)',
                            '後日、予約完了のメールを送信いたします。','もし日時の変更等がある場合は experimenterMailAddress までご連絡ください。',
                            'どうぞよろしくお願いいたします。\n','experimenterName'];
    // --- Failed 仮予約シリーズ ---
    // 時間外・期間外
    var Outoftime = ['participantName 様\n','心理学実験実施責任者のexperimenterNameです。',
                     'この度は心理学実験への応募ありがとうございました。','申し訳ありませんが、ご希望いただいた','expDatefromWhen〜toWhen',
                     'は実験実施可能時間（openTime時〜closeTime時）外または、実施期間（openDate〜closeDate）外です。',
                     'お手数ですが、もう一度登録し直していただきますようお願いします。\n',
                     'experimenterName'];
    // 重複
    var Overlap = ['participantName 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                   '申し訳ありませんが、ご希望いただいた','expDatefromWhen〜toWhen',
                   'にはすでに予約（予定）が入っており（タッチの差で他の方が予約をされた可能性もあります）、実験を実施することができません。',
                   'お手数ですが、もう一度別の日時で登録し直していただきますようお願いします。\n','experimenterName'];
    // --- Successful Booking ---
    // 予約完了テキスト(平日)
    var WeekdayBookingDone = ['participantName 様\n','この度は心理学実験への応募ありがとうございました。',
                              'expDatefromWhen〜toWhenの心理学実験の予約が完了しましたのでメールいたします。',
                              '場所はexperimentRoomです。当日は直接お越しください。',
                              'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
                              '当日もよろしくお願いいたします。\n',
                              '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)','当日の連絡はexperimenterPhoneまでお願いいたします。'];
    // 予約完了テキスト(休日)
    var HolidayBookingDone = ['participantName 様\n','この度は心理学実験への応募ありがとうございました。',
                              'expDatefromWhen〜toWhenの心理学実験の予約が完了しましたのでメールいたします。',
                              '場所はexperimentRoomです。休日は教育学部棟玄関の鍵がかかっており、外から入ることができません。実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。',
                              'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。','当日もよろしくお願いいたします。\n',
                              '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)','当日の連絡はexperimenterPhoneまでお願いいたします。'];
    // --- Rejected Booking ---
    // 既参加
    var AlreadyDone = ['participantName 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                       '大変申し訳ありませんが、以前実施した同様の実験にご参加いただいており、今回の実験にはご参加いただけません。ご了承ください。\n',
                       'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。','今後ともよろしくお願いします。\n','experimenterName'];
    // 定員オーバー
    var ReachedCapacity = ['participantName 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                           '大変申し訳ありませんが、応募いただいた段階ですでに募集人数の定員に達していたため、実験に参加していただくことができません。ご了承ください。\n',
                           '今後、次の実験を実施する際に再度応募していただけると幸いです。','ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
                           '今後ともよろしくお願いいたします。\n','experimenterName'];

    // --- Reminders ---
    // リマインダー(平日)
    var ReminderWeekday =['participantName 様\n','実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
                          '明日 fromWhenから実験に参加していただく予定となっております。','場所はexperimentRoomです。実験時間に実験室まで直接お越しください。\n',
                          'なお、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。','ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
                          'それでは明日、よろしくお願いいたします。\n','experimenterName'];
    // リマインダー(休日)
    var ReminderHoliday =['participantName 様\n','実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
                          '明日 fromWhenから実験に参加していただく予定となっております。','場所はexperimentRoomです。\n',
                          'なお、明日は休日のため教育学部棟玄関の鍵がかかっており、外から入ることができません。実験者が実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n',
                          'また、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。',
                          'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。','それでは明日、よろしくお願いいたします。\n','experimenterName'];

    var notUsed = '利用する場合はここに本文を記載するとともに土日での変更の数字を1に変えてください。なお，改行は"alt + enter"です'

    var note = '適宜変更してください。参加者名は participantName ，実験実施時間は fromWhen および toWhen に代入されます。その他のキーは設定シートを参照してください。'

    // これでいけるかも
    var bodies = {
      "仮予約":TentativeBooking,
      '時間外':Outoftime,
      "重複":Overlap,
      "予約完了wd":WeekdayBookingDone,
      "予約完了we":HolidayBookingDone,
      222:AlreadyDone,
      333:ReachedCapacity,
      "リマインダーwd":ReminderWeekday,
      "リマインダーwe":ReminderHoliday
    };

    for (key in bodies) {
      bodies[key] = bodies[key].join('\n');
    }

    var defaultTemplate = [['トリガー', '土日での変更', '題名', '本文（平日）', '本文（土日）', '備考'],
                           ['仮予約', 0, '予約の確認', bodies['仮予約'], notUsed, note],
                           ['時間外', 0, '実験実施可能時間外です', bodies['時間外'], notUsed, note],
                           ['重複', 0, '予約が重複しています', bodies['重複'], notUsed, note],
                           [111, 1, '実験予約が完了いたしました', bodies['予約完了wd'], bodies['予約完了we'], note],
                           [222, 0, '以前に実験にご参加いただいたことがあります', bodies[222], notUsed, note],
                           [333, 0, '定員に達してしまいました', bodies[333], notUsed, note],
                           ['リマインダー', 1, '明日実施の心理学実験のリマインダー', bodies['リマインダーwd'], bodies['リマインダーwe'], note]];
    var tempNRow = defaultTemplate.length;
    var tempNCol = defaultTemplate[0].length;
    template.getRange(1, 1, tempNRow, tempNCol).setValues(defaultTemplate);

    var member = ss.getSheetByName('メンバー');
    var sh1Name = sheets[0].getName();
    var sh1LastCol = sheets[0].getLastColumn();
    var sh1LColNotation = sheets[0].getRange(1, sh1LastCol).getA1Notation().replace(/\d/,''); // 列のアルファベットを取得
    var formula = "=COUNTIF('" + sh1Name + "'!" + sh1LColNotation + ":" + sh1LColNotation + ", A2)"
    Logger.log([sh1Name, sh1LastCol, sh1LColNotation, formula]);
    var defaultMember = [['キー', '名前', 'アドレス', '担当回数', '備考'],
                         [1, 'りんご', 'apple@hogege.com', formula,'Gmailのアドレスでなくても大丈夫です。'],
                         [2, 'ごりら', 'gorilla@hogege.com','',''],
                         [3, 'らっぱ', 'horn@hogege.com','','']];
    var memNRow = defaultMember.length;
    var memNCol = defaultMember[0].length;
    member.getRange(1, 1, memNRow, memNCol).setValues(defaultMember);

    config.getRange(2, 3, 9, 1).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
    config.activate();
  } catch(err) {
    var fb = "[line " + err.lineNumber + "] " +err.message;
    Logger.log(fb);
    Browser.msgBox("エラーが発生しました", fb, Browser.Buttons.OK);
  }
}
