var type = 1; // 1: 自由回答, 2: 選択式 どちらかの半角数字を入れてください。

function init() { // デフォルトの設定を作成する関数
  setting();
}

var ss = SpreadsheetApp.getActiveSpreadsheet(); // spreadsheet
var sheets = ss.getSheets();

try{
  if (sheets.length > 1) {
    var expInfo = getExpInfo();
    var templates = getTemplate();
    var answers = sheets[0];
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
  console.log(err.message);
}

// https://qiita.com/Layzie/items/465e715dae14e2f601de より
function is(type, obj) {
  var clas = Object.prototype.toString.call(obj).slice(8, -1);
  return obj !== undefined && obj !== null && clas === type;
}

function zenToHan(str) {
  if (is('String', str)) {
    return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) { // 全角を半角に変換
      return String.fromCharCode(s.charCodeAt(0) - 65248); // 10進数の場合
    });
  } else {
    return str;
  }
}

// 実験の設定を取得する関数
function getExpInfo() {
  var sheet = ss.getSheetByName('設定');
  var expInfo = sheet.getDataRange().getValues();//tRange(2, 2, lastRow - 1, 2).getValues();
  var expInfoDict = {};
  // 取得した配列を連想配列に変換する
  for (var i = 1; i < expInfo.length; i++) {
    expInfoDict[expInfo[i][1]] = zenToHan(expInfo[i][2]); // 念の為;
  }
  return expInfoDict;
}

// 希望日時を取得しdate型に変換する関数
function getExpDateTime(array) {
  if (type == 1) {
    var from = new Date(array[colExpDate - 1]);
    var to = new Date(from);
    var expLength = expInfo['experimentLength'];
    to.setMinutes(from.getMinutes() + expLength);
  } else {
    // 日付の操作
    var from = new Date();
    var date = array[colExpDate - 1];
    date = zenToHan(date);
    var dateInfo = date.match(/\d+/g);
    if (dateInfo.length == 3) { //年月日なら
      from.setFullYear(dateInfo[0], dateInfo[1] - 1, dateInfo[2]);
    } else if (dateInfo.length == 2) { //月日なら
      from.setMonth(dateInfo[0] - 1, dateInfo[1]);
    } else if (dateInfo.length == 1) { //日なら
      from.setDate(dateInfo[0]);
    }
    from.setSeconds(0,0);
    var to = new Date(from);
    // 時間の操作
    var time = array[colExpTime - 1]; // この段階では文字列
    time = zenToHan(time);
    var FromTo = time.match(/\d+/g); //空白を除去し，~で分けて要素２の配列に
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
function getTemplate() {
  var sheet = ss.getSheetByName('テンプレート');
  var contents = sheet.getDataRange().getValues();
  var templateDict = {};
  for (var i = 0; i < contents.length; i++){
    var mailContents = {};
    mailContents['changeByDay'] = contents[i][1];
    mailContents['subject'] = contents[i][2];
    mailContents['bodywd'] = contents[i][3];
    mailContents['bodywe'] = contents[i][4];
    templateDict[contents[i][0]] = mailContents;
  }
  return templateDict;
}

function getMailContents(trigger, time) {
  var useTemplate = templates[trigger];
  var mailContents = {};
  mailContents['subject'] = useTemplate['subject'];
  if (useTemplate['changeByDay'] == 1 && (time.getDay()==0 || time.getDay()==6)){ //もし週末なら
    mailContents['body'] = useTemplate['bodywe'];
  }
  else{
    mailContents['body'] = useTemplate['bodywd'];
  }
  return mailContents;
}

function myFormatDate(datetime, pattern) {
  if (is('Date', datetime)) {
    return Utilities.formatDate(datetime, 'JST', pattern);
  }
  return datetime;
}

// mailの内容を作成する関数
function sendEmail(name, address, from, to, trigger, chargeID) {
  //メールに記載する、予約日時の変数を作成する
  var yobi = new Array("日", "月", "火", "水", "木", "金", "土")[from.getDay()];
  expInfo['participantName'] = name;
  expInfo['expDate'] = myFormatDate(from, 'MM/dd') + "（"+ yobi +"）";
  expInfo['fromWhen'] = myFormatDate(from, 'HH:mm');
  expInfo['toWhen'] = myFormatDate(to, 'HH:mm');
  expInfo['openDate'] = myFormatDate(expInfo['openDate'], 'yyyy/MM/dd');
  expInfo['closeDate'] = myFormatDate(expInfo['closeDate'], 'yyyy/MM/dd');
  var contents = getMailContents(trigger, from);
  var subject = makeMailBody(contents['subject']);
  var body = makeMailBody(contents['body']);
  var bccAddresses = getbccAddresses(chargeID);
  MailApp.sendEmail(address, subject, body, {bcc: bccAddresses});
}

function modifySheet(sheet, numRow, columns, values) {
  for (var i = 0; i < columns.length; i++){
    sheet.getRange(numRow, columns[i]).setValue(values[i]);
  }
}

function getMemberInfo() {
  var sheet = ss.getSheetByName('メンバー');
  var memberInfo = sheet.getDataRange().getValues();
  var memberInfoDict = {};
  // 取得した配列を連想配列に変換する
  for (var i = 0; i < memberInfo.length; i++) {
    key = zenToHan(memberInfo[i][0]);
    memberInfoDict[key] = zenToHan(memberInfo[i][2]); // 念の為;
  }
  return memberInfoDict;
}

// memberシートからbccアドレスを追加する関数
function getbccAddresses(charges) {
  // var activeSheet = ss.getActiveSheet();
  var memberInfo = getMemberInfo();
  // var charges = activeSheet.getRange(row, colCharge).getValue();
  charges = zenToHan(charges);
  var bccAddress = [expInfo['experimenterMailAddress']];
  if (is('Number', charges)) {// 一人だけが指定されている場合
    bccAddress.push(memberInfo[charges]);
  } else if (is('String', charges)) {// 複数人が指定されている場合
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

function detectDefault(){
  var defName = false; if (expInfo['experimenterName'] == '実験太郎') defName = true;
  var defPhone = false;if (expInfo['experimenterPhone'] == 'xxx-xxx-xxx') defPhone = true;
  var defPlace = false;if (expInfo['experimentRoom'] == '実施場所') defPlace = true;
  var title = "設定がデフォルトのままです"
  var fb = "以下の重要な設定がデフォルトのままだったので，参加希望者への予約確認メールの送信を中止しました。\n\n"
  if (defName || defPhone || defPlace) {
    if (defName) fb += "実験者名\n";
    if (defPhone) fb += "電話番号\n";
    if (defPlace) fb += "実施場所\n";
    fb += "\n変更後，再度参加者応募のテストをして，予約確認のメールが送信されるかどうか，およびその本文が適切かどうかを確認してください。"
    MailApp.sendEmail(expInfo['experimenterMailAddress'], title, fb);
    return true;
  }
  return false;
}

//仮予約があった際に、カレンダーに書き込む関数
function sendToCalendar(e) {
  var sheet = ss.getActiveSheet();
  try{
    //実験情報の取得
    var submittedInfo = e.values;
    var participantName = submittedInfo[colParName - 1];

    //重複の確認
    var expDT = getExpDateTime(submittedInfo);
    var from = expDT['from']; //仮予約の開始時間を取得
    var to = expDT['to'];//仮予約の開始時間から終了時間を設定
    var openTime = expInfo['openTime'];
    var closeTime = expInfo['closeTime'];
    var openDate = expInfo['openDate'];
    var closeDate = expInfo['closeDate'];

    var cal = CalendarApp.getCalendarById(expInfo['experimenterMailAddress']); //仮予約を記載するカレンダーを取得
    var allEvents = cal.getEvents(from, to);

    if (allEvents.length > 0) {
      var trigger = '重複';
      var values = [trigger, 1, 'N/A', 'N/A'];
    } else if (from.getHours() < openTime || to.getHours() > closeTime || from < openDate || from > closeDate) {
      var trigger = '時間外';
      var values = [trigger, 1, 'N/A', 'N/A'];
    } else {
      var trigger = '仮予約'
      var values = ['', '', '', ''];
      var eventTitle = "仮予約:" + participantName;
      cal.createEvent(eventTitle, from, to); //仮予約情報をカレンダーに作成
    }
    var ParticipantEmail = submittedInfo[colAddress - 1];
    if (!detectDefault()) {
      var numRow = e.range.getRow();
      sendEmail(participantName, ParticipantEmail, from, to, trigger, false);
      modifySheet(sheet, numRow, [colStatus, colMailed, colRemindDate, colReminded], values);
    }
    console.log('Success!');
  } catch(err) {
    //実行に失敗した時に通知
    var fb = "[line " + err.lineNumber + "] " +err.message;
    console.log(fb);
    MailApp.sendEmail(expInfo['experimenterMailAddress'], "エラーが発生しました", fb);
  }
}

function updateCalendar(prevTitle, newTitle, from, to, trigger) {
  var cal = CalendarApp.getCalendarById(expInfo['experimenterMailAddress']); //予約を記載するカレンダーを取得
  // まず予約イベントを削除する
  var reserve = cal.getEvents(from, to);
  for (var i = 0; i < reserve.length; i++) {
    if (reserve[i].getTitle() == prevTitle) {
      reserve[i].deleteEvent();
    }
  }
  if (trigger == expInfo['finalizeTrigger']) {
    cal.createEvent(newTitle, from, to); //予約確定情報をカレンダーに追加
  }
}

function setReminder(from, to, trigger) {
  if (trigger == expInfo['finalizeTrigger']) {
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
  } else { // triggerが指定のトリガー以外のとき
    var values = [1,'N/A','N/A'];
  }
  return values;
}

// スプレッドシート上で予約を完了させ、メール送信及びカレンダーへの書き込みを行う関数
function updateStatus(array) {
  try {
    //予約された日時（見やすい形式）
    var prepTriggers = Object.keys(templates);
    var trigger = String(array[colStatus - 1]);
    var same = function(value) {return value == trigger};
    if (prepTriggers.some(same)) {
      var participantName = array[colParName - 1];
      var expDT = getExpDateTime(array);
      var from = expDT['from']; //予約の開始時間を取得
      var to = expDT['to'];//予約の開始時間から終了時間を設定
      var prevTitle = "仮予約:" + participantName;
      if (colParNameKana > 0) {
        var newTitle = "予約完了:" + participantName +'('+array[colParNameKana - 1]+')';
      } else {
        var newTitle = "予約完了:" + participantName;
      }
      updateCalendar(prevTitle, newTitle, from, to, trigger);
      // メールの送信
      var ParticipantEmail = array[colAddress - 1];
      sendEmail(participantName, ParticipantEmail, from, to, trigger, array[colCharge - 1]);
      return setReminder(from, to, trigger);
    } else {
      return ['','',''];
    }
  } catch(err) {
    //実行に失敗した時に通知
    var fb = "[line " + err.lineNumber + "] " +err.message;
    console.log(fb);
    Browser.msgBox("エラーが発生しました", fb, Browser.Buttons.OK);
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
      var rowVals = data[row];
      if (rowVals[colReminded - 1] == "送信準備") {
        var reminder = rowVals[colRemindDate - 1];
        // もし現在時刻がリマインド日時を過ぎていたならメールを送信
        if ((reminder != "") && (reminder.getTime() <= time)) {
          // メールの本文の内容を作成するための要素を定義
          var participantName = rowVals[colParName - 1]; //被験者の名前
          //参加者にメールを送る
          var ParticipantEmail = rowVals[colAddress - 1];
          var trigger = 'リマインダー';
          var expDT = getExpDateTime(rowVals);//getExpDateTime(sheet, row + 1);
          var from = expDT['from'];
          var to = expDT['to'];
          sendEmail(participantName, ParticipantEmail, from, to, trigger, rowVals[colCharge - 1])
          modifySheet(sheet, row + 1, [colReminded], ['送信済み']);
          console.log('Success!');
        }
      }
    }
  } catch (err) {
    //実行に失敗した時に通知
    var fb = "[line " + err.lineNumber + "] " +err.message;
    console.log(fb);
    MailApp.sendEmail(expInfo['experimenterMailAddress'], "エラーが発生しました", fb);
  }
}

function updateTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    // sendRemindersのトリガーだけを削除する
    if (triggers[i].getEventType() == ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(triggers[i]);
      ScriptApp.newTrigger('sendReminders').timeBased().atHour(expInfo['remindHour']).nearMinute(30).everyDays(1).create();
    }
  }
}

function onFormSubmitted(e) {
  // 実際の回答に続けて値のない回答が送られることがあるので以下のif文で回避
  if (e.values[colAddress - 1].length > 0){
    sendToCalendar(e);
  } else {
    console.log(e.values);
  }
}

function onEdited(e) {
  var edRange = e.range;
  var edSheet = edRange.getSheet();
  var edSheetID = edSheet.getSheetId();
  if (edSheetID === ss.getSheets()[0].getSheetId()) {
    var edValues = edRange.getValues();
    var edFirstRowNum = edRange.getRow();
    var edColNum = edRange.getColumn();
    var lastCol = edSheet.getLastColumn();
    if (edColNum === colStatus) {
      for (var i = 0; i < edValues.length; i++){
        var edRowNum = edFirstRowNum + i;
        var edRowVals = edSheet.getRange(edRowNum, 1, 1, lastCol).getValues()[0];
        if (edRowVals[colMailed - 1] !== 1) {
          var values = updateStatus(edRowVals);
          modifySheet(edSheet, edRowNum, [colMailed, colRemindDate, colReminded], values);
          console.log('Success!');
        }
      }
    }
  } else if (sheetID == ss.getSheetByName('設定').getSheetId()) {
    if (expInfo['remindHour'] != 19) {
      updateTriggers();
    }
  }
}

function setTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('onFormSubmitted').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('onEdited').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('sendReminders').timeBased().atHour(19).nearMinute(30).everyDays(1).create();
}

// 設定用のシートおよびその見本を最初に作る関数
function setting(){
  buttons = Browser.Buttons.OK_CANCEL;
  start = true;
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
    setDefault();
    msg = "初期設定が終了しました。\\n";
    msg += "「設定」シートの太枠に囲まれた項目を適切な情報に変更してください。";
    Browser.msgBox("設定の初期化", msg, Browser.Buttons.OK);
  } else {
    Browser.msgBox("設定の初期化", "初期化はキャンセルされました", Browser.Buttons.OK);
  }
}

function setDefault() {
  try {
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
    var colNames = sheet.getDataRange().getValues();
    var addColNames = [['予約ステータス', '連絡したか', 'リマインド日時', 'リマインドしたか', '担当']];
    if (addNewCol) {
      var newColNames = [colNames[0].concat(addColNames[0])];
      sheet.getRange(1, 1, 1, newColNames[0].length).setValues(newColNames);
    } else {
      sheet.getRange(1, colNames[0].length - addColNames[0].length + 1, 1, addColNames[0].length).setValues(addColNames);
    }
    var lastCol = sheet.getLastColumn();
    // 設定シート
    var start = new Date();
    formattedStart = myFormatDate(start, 'yyyy/MM/dd');
    var end = new Date(start); end.setDate(start.getDate() + 13);
    formattedEnd = myFormatDate(end, 'yyyy/MM/dd');
    var config = ss.getSheetByName('設定');
    var note2 = '「フォームの回答」の列番号と一致しているか確認してください（A列が1）';
    var defaultExpInfo = [['設定項目','メール本文内でのキー','値','備考'],
                          ['実験責任者名','experimenterName','実験太郎', "実験責任者の名前を記入してください"],
                          ['実験責任者のGmailアドレス','experimenterMailAddress', Session.getActiveUser().getEmail(), "実験用のGmailアドレスを記入してください"],
                          ['実験責任者の電話番号','experimenterPhone','xxx-xxx-xxx', "電話番号を記入してください"],
                          ['実験の実施場所','experimentRoom','実施場所',"実験の実施場所を記入してください"],
                          ['実験の所要時間','experimentLength', 60, '実験の所要時間を記入してください。2列目は変更しないでください'],
                          ['実験開始可能時刻','openTime', 9, '何時から実験できるかを記入してください（24時間表記）'],
                          ['実験終了時刻','closeTime', 19,'何時まで実験可能かを記入してください（24時間表記）'],
                          ['実験開始日','openDate', formattedStart, '実験を開始する日付を記入してください（年/月/日で表記）'],
                          ['実験最終日','closeDate', formattedEnd, '実験の終了予定日を記入してください（年/月/日で表記）'],
                          ['リマインダー送信時刻','remindHour', 19, 'リマインダーを送信する時刻を記入してください（24時間表記）。なお指定した時刻から1時間以内に送信されます。'],
                          ['予約を完了させるトリガー','finalizeTrigger',111,'必要に応じて任意の数字・文字列に変更してください'],
                          ['参加者名の列番号','colParName', 2, note2],
                          ['ふりがなの列番号','colParNameKana', -1, note2 + 'もし利用しない場合は-1を入力してください。']];

    var verChoice = [['参加者アドレスの列番号','colAddress', lastCol - 7, note2],
                     ['希望日の列番号','colExpDate', lastCol - 6, note2],
                     ['希望時間の列番号','colExpTime', lastCol - 5, note2]];

    var verAnswer = [['参加者アドレスの列番号','colAddress', lastCol - 6, note2],
                     ['希望日時の列番号','colExpDate', lastCol - 5, note2]];

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
