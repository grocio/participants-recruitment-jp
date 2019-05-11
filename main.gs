var TYPE = 1; // 1: 自由回答, 2: 選択式 どちらかの半角数字を入れてください。

var SS = SpreadsheetApp.getActiveSpreadsheet(); // spreadsheet
var SHEETS = SS.getSheets();
// シート名とsettig時のkeyの対応させるための連想配列
// getSettingとonEditedで使用される
var NAME_TO_KEY = {"設定":"config", "テンプレート":"templates", "メンバー":"memberInfo"};
if (SHEETS.length > 1) {
  var SETTING = getSetting();
  var CONFIG = SETTING.config;
  var TEMPLATES = SETTING.templates;
  var MEMBER_INFO = SETTING.memberInfo;
}

function init(){
  setting();
}

// 型判定のための関数https://qiita.com/Layzie/items/465e715dae14e2f601de より
function is(type, obj) {
  const clas = Object.prototype.toString.call(obj).slice(8, -1);
  return obj !== undefined && obj !== null && clas === type;
}

// 全角を半角に変換する関数
function zenToHan(str) {
  if (is('String', str)) {
    return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) { // 全角を半角に変換
      return String.fromCharCode(s.charCodeAt(0) - 65248); // 10進数の場合
    });
  } else {
    return str;
  }
}

function fmtDate(datetime, pattern) {
  if (is('Date', datetime)) {
    return Utilities.formatDate(datetime, 'JST', pattern);
  }
  return datetime;
}

function getInfo(sheetName) {
  const sheetInfo = SS.getSheetByName(sheetName);
  const infoArray = sheetInfo.getDataRange().getValues();
  const infoObj = {};
  for (var i = 1; i < infoArray.length; i++) {
    // 参照するシートによって処理を変える
    if (sheetName == "設定") {
      var key = infoArray[i][1];
      var property = zenToHan(infoArray[i][2]); // 念の為
      if (key.indexOf("col") == 0) { // 列番号に関する設定は，Numberに変更しておく
        property = Number(property);
      }
    } else if (sheetName == "テンプレート") {
      var key = infoArray[i][0];
      var property = {};
      property.changeByDay = infoArray[i][1];
      property.title = infoArray[i][2];
      property.bodywd = infoArray[i][3];
      property.bodywe = infoArray[i][4];
    } else if (sheetName == "メンバー") {
      var key = zenToHan(infoArray[i][0]);
      var property = zenToHan(infoArray[i][2]);
    }
    infoObj[key] = property;
  }   
  return infoObj;
}

function getSetting(){
  const sheetCache = SS.getSheetByName("Cache");
  const cacheJson = sheetCache.getRange(1,1).getValue();
  if (cacheJson.length < 10){
    var settingObj = {};
    for (name in NAME_TO_KEY) {
      var key = NAME_TO_KEY[name];
      settingObj[key] = getInfo(name);
    }
  } else {
    var settingObj = JSON.parse(cacheJson);
    // parseしたままだと以下の2つがstringのままで機能しない
    settingObj.config.openDate = new Date(settingObj.config.openDate);
    if (settingObj.config.openDate < new Date()) {
      settingObj.config.openDate = new Date();
    }
    settingObj.config.closeDate = new Date(settingObj.config.closeDate);
    if (settingObj.config.closeDate < new Date()) {
      const title = "実験実施期間を修正してください";
      const text = "実験実施期間が過去になっています。早急に修正してください。";
      console.log(text);
      MailApp.sendEmail(settingObj.config.experimenterMailAddress, title, text);
    }
  }
  return settingObj;
}

// 実験期間に合わせてフォームの日にちの選択肢が変わるようにしたい
// function modifyFormItems() {
//   linkedFormURL = SS.getFormUrl();
//   linkedForm = FormApp.getByUrl(linkedFormURL);
// }

///////////////////////////////////////////////////////////////////////////////
// メインの関数群で利用されるミニ関数
///////////////////////////////////////////////////////////////////////////////

// 希望日時を取得しdate型に変換する関数
function getExpDateTime(array) {
  if (TYPE == 1) {
    var from = new Date(array[CONFIG.colExpDate - 1]);
    var to = new Date(from);
    var expLength = CONFIG.experimentLength;
    to.setMinutes(from.getMinutes() + expLength);
  } else {
    // 日付の操作
    var from = new Date();
    var date = array[CONFIG.colExpDate - 1];
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
    var time = array[CONFIG.colExpTime - 1]; // この段階では文字列
    time = zenToHan(time);
    var FromTo = time.match(/\d+/g); //空白を除去し，~で分けて要素２の配列に
    from.setHours(FromTo[0],FromTo[1]);
    if (FromTo.length == 4) {
      to.setHours(FromTo[2],FromTo[3]);
    } else if (FromTo.length == 2) {
      var expLength = CONFIG.experimentLength;
      to.setMinutes(from.getMinutes() + expLength);
    }
  }
  return {'from': from, 'to': to};
}

function getMailContents(trigger, time) {
  const template = TEMPLATES[trigger];
  var body = template.bodywd;
  if (template.changeByDay == 1 && (time.getDay()==0 || time.getDay()==6)) { //もし週末なら
    body = template.bodywe;
  }
  for (key in CONFIG) { // メールの本文の変数を置換する
    var regex = new RegExp(key,'g');
    body = body.replace(regex, CONFIG[key]);
  }
  return {title: template.title, body: body};
}

// memberシートからbccアドレスを追加する関数
function getBccAddresses(charges) {
  charges = zenToHan(charges);
  const bccArray = [CONFIG.experimenterMailAddress];
  if (is('Number', charges)) {// 一人だけが指定されている場合
    bccArray.push(MEMBER_INFO[charges]);
  } else if (is('String', charges)) {// 複数人が指定されている場合
    if (charges.length > 0) {
      const chargeIDs = charges.match(/\d+/g);
      for (var i = 0; i < chargeIDs.length; i++) {
        var chargeID = chargeIDs[i];
        bccArray.push(MEMBER_INFO[chargeID]);
      }
    }
  }
  return bccArray.join(',');
}

// mailの内容を作成する関数
function sendEmail(name, address, from, to, trigger, chargeID) {
  //メールに記載する、予約日時の変数を作成する
  const yobi = new Array("日", "月", "火", "水", "木", "金", "土")[from.getDay()];
  CONFIG.participantName = name;
  CONFIG.expDate = fmtDate(from, 'MM/dd') + "（"+ yobi +"）";
  CONFIG.fromWhen = fmtDate(from, 'HH:mm');
  CONFIG.toWhen = fmtDate(to, 'HH:mm');
  CONFIG.openDate = fmtDate(CONFIG.openDate, 'yyyy/MM/dd');
  CONFIG.closeDate = fmtDate(CONFIG.closeDate, 'yyyy/MM/dd');
  const mail = getMailContents(trigger, from);
  const bccAddresses = getBccAddresses(chargeID);
  MailApp.sendEmail(address, mail.title, mail.body, {bcc: bccAddresses});
}

function updateCalendar(oldEventName, newEventName, from, to, trigger) {
  const cal = CalendarApp.getCalendarById(CONFIG.experimenterMailAddress); //予約を記載するカレンダーを取得
  // まず予約イベントを削除する
  const reserve = cal.getEvents(from, to);
  for (var i = 0; i < reserve.length; i++) {
    if (reserve[i].getTitle() == oldEventName) {
      reserve[i].deleteEvent();
    }
  }
  if (trigger == CONFIG.finalizeTrigger) {
    cal.createEvent(newEventName, from, to); //予約確定情報をカレンダーに追加
  }
}

function setReminder(from, trigger) {
  if (trigger == CONFIG.finalizeTrigger) {
    // リマインダーのための設定をする
    const remindDate = new Date(from)
    remindDate.setDate(from.getDate() - 1); //remindDateの時刻を予約時間の1日前に設定する。
    const time = new Date(); //現在時刻の取得
    time.setHours(19); //19時に設定
    // 予約を完了させた日の19時にremindDateの時刻が達していない場合、"送信準備"というコードを指定のセルに入力する
    if (remindDate > time) {
      return [1, remindDate, "送信準備"];
    }
    return [1, remindDate, "直前のため省略"];
  }
  return [1,'N/A','N/A']; // triggerが指定のトリガー以外のとき
}

function isDefault(){
  const def = {
    Name: false,
    Phone:false,
    Place:false
  };
  if (CONFIG.experimenterName == '実験太郎') def.Name = true;
  if (CONFIG.experimenterPhone == 'xxx-xxx-xxx') def.Phone = true;
  if (CONFIG.experimentRoom == '実施場所') def.Place = true;

  const title = "設定がデフォルトのままです";
  var fb = "以下の重要な設定がデフォルトのままだったので，参加希望者への予約確認メールの送信を中止しました。\n\n";
  if (def.Name || def.Phone || def.Place) { // デフォルトのままなら
    if (def.Name)  fb += "実験者名\n";
    if (def.Phone) fb += "電話番号\n";
    if (def.Place) fb += "実施場所\n";
    fb += "\n変更後，再度参加者応募のテストをして，予約確認のメールが送信されるかどうか，およびその本文が適切かどうかを確認してください。";
    MailApp.sendEmail(CONFIG.experimenterMailAddress, title, fb);
    return true;
  }
  return false;
}

///////////////////////////////////////////////////////////////////////////////
// メインの関数群
///////////////////////////////////////////////////////////////////////////////

//仮予約があった際に、カレンダーに書き込む関数
function checkAppointment(e) {
  try{
    //実験情報の取得
    const answersArray = e.values;
    const participantName = answersArray[CONFIG.colParName - 1];

    //重複の確認
    const expDT = getExpDateTime(answersArray);
    // 設定がデフォルトかどうかを判定する
    if (!isDefault()) { // 設定がデフォルトから変更されていればメールを送る
      const cal = CalendarApp.getCalendarById(CONFIG.experimenterMailAddress); //仮予約を記載するカレンダーを取得
      const allEvents = cal.getEvents(expDT.from, expDT.to);
      var trigger = '仮予約';
      var values = ['', '', '', ''];
      if (allEvents.length > 0) {
        trigger = '重複';
        values = [trigger, 1, 'N/A', 'N/A'];
      } else if (expDT.from.getHours() < CONFIG.openTime || expDT.to.getHours() > CONFIG.closeTime || 
                 expDT.from < CONFIG.openDate || expDT.from > CONFIG.closeDate) {
        trigger = '時間外';
        values = [trigger, 1, 'N/A', 'N/A'];
      } else {
        const eventTitle = "仮予約:" + participantName;
        cal.createEvent(eventTitle, expDT.from, expDT.to); //仮予約情報をカレンダーに作成
      }
      const participantEmail = answersArray[CONFIG.colAddress - 1];
      const sheetAnswers = SHEETS[0];
      const numRow = e.range.getRow();
      const colNumArray = [CONFIG.colStatus, CONFIG.colMailed, CONFIG.colRemindDate, CONFIG.colReminded];
      sendEmail(participantName, participantEmail, expDT.from, expDT.to, trigger, false);
      // sheetの修正
      sheetAnswers.getRange(numRow, colNumArray[0], 1, colNumArray.length).setValues([values]);
    }
    console.log('Success!');
  } catch(err) {
    //実行に失敗した時に通知
    const fb = "[line " + err.lineNumber + "] " +err.message;
    console.log(fb);
    MailApp.sendEmail(CONFIG.experimenterMailAddress, "エラーが発生しました", fb);
  }
}

// スプレッドシート上で予約を完了させ、メール送信及びカレンダーへの書き込みを行う関数
function finalizeAppointment(array) {
  const prepTriggers = Object.keys(TEMPLATES);
  const trigger = String(array[CONFIG.colStatus - 1]);
  const identical = function(value) {return value == trigger};
  if (prepTriggers.some(identical)) {
    const participantName = array[CONFIG.colParName - 1];
    const expDT = getExpDateTime(array);  //予約された日時（見やすい形式）
    const oldEventName = "仮予約:" + participantName;
    var newEventName = "予約完了:" + participantName;
    if (CONFIG.colParNameKana > 0) {
      newEventName = newEventName + '('+array[CONFIG.colParNameKana - 1]+')';
    }
    updateCalendar(oldEventName, newEventName, expDT.from, expDT.to, trigger);
    // メールの送信
    const ParticipantEmail = array[CONFIG.colAddress - 1];
    sendEmail(participantName, ParticipantEmail, expDT.from, expDT.to, trigger, array[CONFIG.colCharge - 1]);
    return setReminder(expDT.from, trigger);
  } else {
    return ['','',''];
  }
}

//リマインダーを実行する関数
function sendReminders() {
  try {
    const sheetAnswers = SHEETS[0];
    const data = sheetAnswers.getDataRange().getValues(); //シート全体のデータを取得。2次元の配列 [行 [列]]
    const time = new Date().getTime(); //現在時刻の取得
    // スプレッドシートを1列ずつ参照し、該当する被験者を探していく。
    for (var row = 1; row < data.length; row++) { // 0行目は列名
      //ステータスが送信準備になっていることを確認する
      var rowVals = data[row];
      if (rowVals[CONFIG.colReminded - 1] == "送信準備") {
        var reminder = rowVals[CONFIG.colRemindDate - 1];
        // もし現在時刻がリマインド日時を過ぎていたならメールを送信
        if ((reminder != "") && (reminder.getTime() <= time)) {
          // メールの本文の内容を作成するための要素を定義
          var participantName = rowVals[CONFIG.colParName - 1]; //被験者の名前
          //参加者にメールを送る
          var participantEmail = rowVals[CONFIG.colAddress - 1];
          var expDT = getExpDateTime(rowVals);
          sendEmail(participantName, participantEmail, expDT.from, expDT.to, 'リマインダー', rowVals[CONFIG.colCharge - 1])
          sheetAnswers.getRange(row + 1, CONFIG.colReminded).setValue("送信済み"); // シートの修正
          console.log('Success!');
        }
      }
    }
  } catch (err) {
    //実行に失敗した時に通知
    const fb = "[line " + err.lineNumber + "] " +err.message;
    console.log(fb);
    MailApp.sendEmail(CONFIG.experimenterMailAddress, "エラーが発生しました", fb);
  }
}

///////////////////////////////////////////////////////////////////////////////
// トリガー用の関数
///////////////////////////////////////////////////////////////////////////////

function updateTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    // sendRemindersのトリガーだけを削除する
    if (triggers[i].getEventType() == ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(triggers[i]);
      ScriptApp.newTrigger('sendReminders').timeBased().atHour(CONFIG.remindHour).nearMinute(30).everyDays(1).create();
    }
  }
}

function onFormSubmitted(e) {
  // 実際の回答に続けて値のない回答が送られることがあるので以下のif文で回避
  if (e.values[CONFIG.colAddress - 1].length > 0){
    checkAppointment(e);
  } else {
    console.log(e.values);
  }
}

function onEdited(e) {
  try {
    const edRange = e.range;
    const edSheet = edRange.getSheet();
    const edSheetName = edSheet.getSheetName();
    if (edSheetName === SHEETS[0].getSheetName()) {
      const edColNum = edRange.getColumn();
      if (edColNum === CONFIG.colStatus) {
        const edValues = edRange.getValues();
        const edFirstRowNum = edRange.getRow();
        const answersArray = edSheet.getDataRange().getValues();
        for (var i = 0; i < edValues.length; i++) {
          var edRowNum = edFirstRowNum + i;
          var edRowVals = answersArray[edRowNum - 1];
          if (edRowVals[CONFIG.colMailed - 1] !== 1) {
            var values = finalizeAppointment(edRowVals);
            var colNumArray = [CONFIG.colMailed, CONFIG.colRemindDate, CONFIG.colReminded];
            edSheet.getRange(edRowNum, colNumArray[0], 1, colNumArray.length).setValues([values]);
            console.log('Success!');
          }
        }
      }
    } else {
      const newInfo = getInfo(edSheetName);
      // 変更を加えたシートだけcacheを変更する
      const cache = {};
      for (name in NAME_TO_KEY){
        key = NAME_TO_KEY[name];
        if (name == edSheetName) {
          cache[key] = newInfo;
        } else {
        cache[key] = SETTING[key];
        }
      }
      const sheetCache = SS.getSheetByName('Cache');
      sheetCache.getRange(1,1).setValue(JSON.stringify(cache));
      // Logger.log(edValues);
      if (edSheetName == '設定') {
        if (newInfo.remindHour != 19) {
          updateTriggers();
        }
      }
    }
  } catch (err) {
    //実行に失敗した時に通知
    const fb = "[line " + err.lineNumber + "] " +err.message;
    console.log(fb);
    Browser.msgBox("エラーが発生しました", fb, Browser.Buttons.OK);
  }
}

///////////////////////////////////////////////////////////////////////////////
// 初期設定に関わる関数
///////////////////////////////////////////////////////////////////////////////

function setTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('onFormSubmitted').forSpreadsheet(SS).onFormSubmit().create();
  ScriptApp.newTrigger('onEdited').forSpreadsheet(SS).onEdit().create();
  ScriptApp.newTrigger('sendReminders').timeBased().atHour(19).nearMinute(30).everyDays(1).create();
}

// 設定用のシートおよびその見本を最初に作る関数
function setting(){
  var buttons = Browser.Buttons.OK_CANCEL;
  var start = true;
  if (SHEETS.length > 1) {
    var msg = "一度設定を行ったことがあるようです（シートが2枚以上あります）。\\nもう一度初期化を行いますか？\\n"
    msg += "フォームの回答が一番初めのシートでないとこれまでの情報が失われる場合があります。"
    var choice = Browser.msgBox("設定の初期化を行います", msg, buttons);
    if (choice !== "ok") {
      start = false;
    }
  }
  if (TYPE == 1) {
    var msg = "自由回答形式の設定で初期化を行います";
  } else if (TYPE == 2) {
    var msg = "選択形式の設定で初期化を行います";
  } else {
    var msg = "半角数字の1か2を入力して設定の形式を選択してください";
    buttons = Browser.Buttons.OK;
    start = false;
  }
  choice = Browser.msgBox("設定の初期化", msg, buttons);
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
    if (SHEETS.length > 2) {
      for (i = 1; i < SHEETS.length; i++) {
        SS.deleteSheet(SHEETS[i]);
      }
      addNewCol = false;
    }
    const sheetAnswers = SHEETS[0];
    SS.insertSheet('設定');
    SS.insertSheet('テンプレート');
    SS.insertSheet('メンバー');
    SS.insertSheet('Cache');
    const colNames = sheetAnswers.getDataRange().getValues();
    const addColNames = [['予約ステータス', '連絡したか', 'リマインド日時', 'リマインドしたか', '担当']];
    if (addNewCol) {
      const newColNames = [colNames[0].concat(addColNames[0])];
      sheetAnswers.getRange(1, 1, 1, newColNames[0].length).setValues(newColNames);
    } else {
      sheetAnswers.getRange(1, colNames[0].length - addColNames[0].length + 1, 1, addColNames[0].length).setValues(addColNames);
    }
    const lastCol = sheetAnswers.getLastColumn();
    // 設定シート
    const start = new Date();
    formattedStart = fmtDate(start, 'yyyy/MM/dd');
    const end = new Date(start); end.setDate(start.getDate() + 13);
    formattedEnd = fmtDate(end, 'yyyy/MM/dd');
    const sheetConfig = SS.getSheetByName('設定');
    const note2 = '「フォームの回答」の列番号と一致しているか確認してください（A列が1）';
    var defaultConfig = [
      ['設定項目','メール本文内でのキー','値','備考'],
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
      ['ふりがなの列番号','colParNameKana', -1, note2 + 'もし利用しない場合は-1を入力してください。']
    ];

    const verChoice = [
      ['参加者アドレスの列番号','colAddress', lastCol - 7, note2],
      ['希望日の列番号','colExpDate', lastCol - 6, note2],
      ['希望時間の列番号','colExpTime', lastCol - 5, note2]
    ];

    const verAnswer = [
      ['参加者アドレスの列番号','colAddress', lastCol - 6, note2],
      ['希望日時の列番号','colExpDate', lastCol - 5, note2]
    ];

    const otherColConfig = [
      ['予約ステータスの列番号','colStatus', lastCol - 4, note2],
      ['「連絡したか」の列番号','colMailed', lastCol - 3, note2],
      ['リマインド日時の列番号','colRemindDate', lastCol - 2, note2],
      ['「リマインドしたか」の列番号','colReminded', lastCol - 1, note2],
      ['担当の列番号','colCharge', lastCol, note2]
    ];

    if (TYPE == 1) {
      defaultConfig = defaultConfig.concat(verAnswer).concat(otherColConfig);
    } else {
      defaultConfig = defaultConfig.concat(verChoice).concat(otherColConfig);
    }

    const configNRow = defaultConfig.length;
    const configNCol = defaultConfig[0].length;
    sheetConfig.getRange(1, 1, configNRow, configNCol).setValues(defaultConfig);


    // メールのテンプレート用シート
    const sheetTemplate = SS.getSheetByName('テンプレート');

    // successful 仮予約
    const tentativeBooking = [
      'participantName 様\n',
      '心理学実験実施責任者のexperimenterNameです。',
      'この度は心理学実験への応募ありがとうございました。',
      '予約の確認メールを自動で送信しております。\n',
      'expDate fromWhen〜toWhen',
      'で予約を受け付けました（まだ確定はしていません)。',
      '後日、予約完了のメールを送信いたします。',
      'もし日時の変更等がある場合は experimenterMailAddress までご連絡ください。',
      'どうぞよろしくお願いいたします。\n',
      'experimenterName'
    ];
    // --- Failed 仮予約シリーズ ---
    // 時間外・期間外
    const outOfTime = [
      'participantName 様\n',
      '心理学実験実施責任者のexperimenterNameです。',
      'この度は心理学実験への応募ありがとうございました。',
      '申し訳ありませんが、ご希望いただいた',
      'expDate fromWhen〜toWhen',
      'は実験実施可能時間（openTime時〜closeTime時）外または、実施期間（openDate〜closeDate）外です。',
      'お手数ですが、もう一度登録し直していただきますようお願いします。\n',
      'experimenterName'
    ];
    // 重複
    const overlap = [
      'participantName 様\n',
      '心理学実験実施責任者のexperimenterNameです。',
      'この度は心理学実験への応募ありがとうございました。',
      '申し訳ありませんが、ご希望いただいた',
      'expDate fromWhen〜toWhen',
      'にはすでに予約（予定）が入っており（タッチの差で他の方が予約をされた可能性もあります）、実験を実施することができません。',
      'お手数ですが、もう一度別の日時で登録し直していただきますようお願いします。\n',
      'experimenterName'
    ];
    // --- Successful Booking ---
    // 予約完了テキスト(平日)
    const weekdayBookingDone = [
      'participantName 様\n',
      'この度は心理学実験への応募ありがとうございました。',
      'expDate fromWhen〜toWhenの心理学実験の予約が完了しましたのでメールいたします。',
      '場所はexperimentRoomです。当日は直接お越しください。',
      'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
      '当日もよろしくお願いいたします。\n',
      '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)',
      '当日の連絡はexperimenterPhoneまでお願いいたします。'
    ];
    // 予約完了テキスト(休日)
    const holidayBookingDone = [
      'participantName 様\n',
      'この度は心理学実験への応募ありがとうございました。',
      'expDate fromWhen〜toWhenの心理学実験の予約が完了しましたのでメールいたします。',
      '場所はexperimentRoomです。休日は教育学部棟玄関の鍵がかかっており、外から入ることができません。実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。',
      'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
      '当日もよろしくお願いいたします。\n',
      '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)',
      '当日の連絡はexperimenterPhoneまでお願いいたします。'
    ];
    // --- Rejected Booking ---
    // 既参加
    const alreadyParticipated = [
      'participantName 様\n',
      '心理学実験実施責任者のexperimenterNameです。',
      'この度は心理学実験への応募ありがとうございました。',
      '大変申し訳ありませんが、以前実施した同様の実験にご参加いただいており、今回の実験にはご参加いただけません。ご了承ください。\n',
      'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
      '今後ともよろしくお願いします。\n',
      'experimenterName'
    ];
    // 定員オーバー
    const reachedCapacity = [
      'participantName 様\n',
      '心理学実験実施責任者のexperimenterNameです。',
      'この度は心理学実験への応募ありがとうございました。',
      '大変申し訳ありませんが、応募いただいた段階ですでに募集人数の定員に達していたため、実験に参加していただくことができません。ご了承ください。\n',
      '今後、次の実験を実施する際に再度応募していただけると幸いです。',
      'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
      '今後ともよろしくお願いいたします。\n',
      'experimenterName'
    ];

    // --- Reminders ---
    // リマインダー(平日)
    const reminderWeekday =[
      'participantName 様\n',
      '実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
      '明日 fromWhenから実験に参加していただく予定となっております。',
      '場所はexperimentRoomです。実験時間に実験室まで直接お越しください。\n',
      'なお、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。',
      'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
      'それでは明日、よろしくお願いいたします。\n',
      'experimenterName'
    ];
    // リマインダー(休日)
    const reminderHoliday =[
      'participantName 様\n',
      '実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
      '明日 fromWhenから実験に参加していただく予定となっております。',
      '場所はexperimentRoomです。\n',
      'なお、明日は休日のため教育学部棟玄関の鍵がかかっており、外から入ることができません。実験者が実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n',
      'また、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。',
      'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
      'それでは明日、よろしくお願いいたします。\n',
      'experimenterName'
    ];

    const notUsed = '利用する場合はここに本文を記載するとともに土日での変更の数字を1に変えてください。なお，改行は"alt + enter"です';

    const note = '適宜変更してください。参加者名は participantName ，実験実施時間は fromWhen および toWhen に代入されます。その他のキーは設定シートを参照してください。';

    // これでいけるかも
    const bodies = {
      "仮予約":tentativeBooking,
      '時間外':outOfTime,
      "重複":overlap,
      "予約完了wd":weekdayBookingDone,
      "予約完了we":holidayBookingDone,
      222:alreadyParticipated,
      333:reachedCapacity,
      "リマインダーwd":reminderWeekday,
      "リマインダーwe":reminderHoliday
    };

    for (key in bodies) {
      bodies[key] = bodies[key].join('\n');
    };

    const defaultTemplate = [
      ['トリガー', '土日での変更', '題名', '本文（平日）', '本文（土日）', '備考'],
      ['仮予約', 0, '予約の確認', bodies['仮予約'], notUsed, note],
      ['時間外', 0, '実験実施可能時間外です', bodies['時間外'], notUsed, note],
      ['重複', 0, '予約が重複しています', bodies['重複'], notUsed, note],
      [111, 1, '実験予約が完了いたしました', bodies['予約完了wd'], bodies['予約完了we'], note],
      [222, 0, '以前に実験にご参加いただいたことがあります', bodies[222], notUsed, note],
      [333, 0, '定員に達してしまいました', bodies[333], notUsed, note],
      ['リマインダー', 1, '明日実施の心理学実験のリマインダー', bodies['リマインダーwd'], bodies['リマインダーwe'], note]
    ];
    const tempNRow = defaultTemplate.length;
    const tempNCol = defaultTemplate[0].length;
    sheetTemplate.getRange(1, 1, tempNRow, tempNCol).setValues(defaultTemplate);

    const sheetMember = SS.getSheetByName('メンバー');
    const sh1Name = sheetAnswers.getName();
    const sh1LastCol = sheetAnswers.getLastColumn();
    const sh1LColNotation = sheetAnswers.getRange(1, sh1LastCol).getA1Notation().replace(/\d/,''); // 列のアルファベットを取得
    const formula = "=COUNTIF('" + sh1Name + "'!" + sh1LColNotation + ":" + sh1LColNotation + ", A2)"
    Logger.log([sh1Name, sh1LastCol, sh1LColNotation, formula]);
    const defaultMember = [
      ['キー', '名前', 'アドレス', '担当回数', '備考'],
      [1, 'りんご', 'apple@hogege.com', formula,'Gmailのアドレスでなくても大丈夫です。'],
      [2, 'ごりら', 'gorilla@hogege.com','',''],
      [3, 'らっぱ', 'horn@hogege.com','','']
    ];
    const memNRow = defaultMember.length;
    const memNCol = defaultMember[0].length;
    sheetMember.getRange(1, 1, memNRow, memNCol).setValues(defaultMember);

    sheetConfig.getRange(2, 3, 9, 1).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheetConfig.activate();
  } catch(err) {
    const fb = "[line " + err.lineNumber + "] " +err.message;
    Logger.log(fb);
    Browser.msgBox("エラーが発生しました", fb, Browser.Buttons.OK);
  }
}
