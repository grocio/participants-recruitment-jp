var ss = SpreadsheetApp.getActiveSpreadsheet(); // spreadsheet

try{
  if (ss.getSheets().length > 1) {
    var expInfo = getExpInfo();
    var answers = ss.getSheets()[0];
    var colParName = 2
    var colParNameKana = colParNameKana
    var colMember = answers.getLastColumn();
    var colReminded = colMember - 1;
    var colRemindDate = colMember - 2;
    var colMailed = colMember - 3;
    var colStatus = colMember - 4;
    var colExpTime = colMember - 5;
    var colExpDate = colMember - 6;
    var colAddress = colMember - 7;
    // ユーザーが変更していた場合はそれを反映するようにする
    if (Number(expInfo['colParName']) !== colParName) {colParName = Number(expInfo['colParName'])};
    if (Number(expInfo['colParNameKana']) !== colParName) {colParNameKana = Number(expInfo['colParNameKana'])};
    if (Number(expInfo['colAddress']) !== colAddress) {colAddress = Number(expInfo['colAddress'])};
    if (Number(expInfo['colExpDate']) !== colExpDate) {colExpDate = Number(expInfo['colExpDate'])};
    if (Number(expInfo['colExpTime']) !== colExpTime) {colExpTime = Number(expInfo['colExpTime'])};
  }
} catch (e) {
  Logger.log(e.message);
}

"メールに記載する時間のフォーマットなどをどうするか考える"

// 実験の情報を取得する関数
function getExpInfo() {
  var sheet = ss.getSheetByName('設定');
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  "要修正 設定シートとメールのテンプレシートを分ける"
  var expInfo = sheet.getRange(2, 2, lastRow, lastCol).getValues();
  var expInfoDict = {};
  // 取得した配列を連想配列に変換する
  for (var i = 0; i < expInfo.length; i++) {
    expInfoDict[expInfo[i][0]] = expInfo[i][1];
  }
  return expInfoDict;
}

// 希望日時を取得しdate型に変換する関数
function getExpDateTime(sheet, row) {
  var from = new Date(sheet.getRange(row, colExpTime).getValue());
  var to = new Date(sheet.getRange(row, colExpTime).getValue());
  var datetime = sheet.getRange(row, colExpDate).getValue(); // この段階では文字列
  var FromTo = datetime.replace(/\s+/g, '').split('~'); //空白を除去し，~で分けて要素２の配列に
  var sHM = FromTo[0].split(':'); // [hour, minute] の配列
  var eHM = FromTo[1].split(':');
  from.setHours(sHM[0],sHM[1]);
  to.setHours(eHM[0],eHM[1]);
  return {'from': from, 'to': to};
}

// メールの本文を作成する関数
function makeMailBody(body) {
  var variables = body.match(/[A-Z]+/g); // 本文中の変数を大文字だけにする
  for (var i = 0; i < variables.length; i++){
    var variable = variables[i];
    // ユーザーが施設名などで大文字を使用した場合にエラーを起こさないように修正する必要あり
    var body = body.replace(variable, info[variable]);
  }
  return body;
}

// スプレッドシートからメールのテンプレートを取得する関数
function getTemplate(trigger, time) {
  var sheet = ss.getSheetByName('テンプレート');
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  "要修正 設定シートとメールのテンプレシートを分ける"
  var contents = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var contentsDict = {};
  for (var i = 0; i < contents.length; i++){
    if (contents[i][0] === trigger){
      contentsDict['subject'] = contents[i][2];
      var changeByDay = contents[i][1];
      if (changeByDay === "あり" && (time.getDay()==0 || time.getDay()==6)){ //もし週末なら
        contentsDict['body'] = contents[i][4];
      }
      else{
        contentsDict['body'] = contents[i][3];
      }
    }
  }
  return contentsDict;
}

// mailの内容を作成する関数
function sendEmail(name, address, from, to, trigger) {
  //メールに記載する、予約日時の変数を作成する
  '要修正 beautifulDateをgasのbuit-in関数に置き換える'
  var reservedAppo = Utilities.formatDate(from, 'Asia/Tokyo', 'HH:mm');
  // たぶん上のやつ
  var appo = beautifulDate(from, 'full') + '〜' + beautifulDate(to, 'time');
  expInfo['PARTICIPANTNAME'] = name;
  expInfo['WHEN'] = appo;
  var contents = getTemplate(trigger, from);
  var subject = contents['subject'];
  var body = makeMailBody(contents['body']);
  var bccAddresses = getbccAddresses(expInfo['experimenterMailAddress'],activeCellRow)
  MailApp.sendEmail(address, subject, body, {bcc: bccAddresses});
}

function modifySheet(sheet, numRow, operationDict) {
  if (typeof sheet.getSheets === 'function') {
    sheet = sheet.getSheets()[0];
  }
  if (typeof operationDict !== 'undefined') {
    for(var i = 0; i < Object.keys(operationDict).length; i++){
      var arrayKeys = Object.keys(operationDict);
      var key = arrayKeys[i];
      sheet.getRange(numRow, key).setValue(operationDict[key]);}
  }
}

// memberシートからbccアドレスを追加する関数
function getbccAddresses(firstaddress, row) {
  var activeSheet = ss.getActiveSheet();
  // Logger.log(activeSheet);
  var members = ss.getSheetByName('member');
  var members_mtrx = members.getDataRange().getValues();
  // Logger.log(members_mtrx);
  var lastCol = activeSheet.getLastColumn();
  var charge = activeSheet.getRange(row,lastCol).getValue();
  // Logger.log(charge);
  var bccAddress = [firstaddress];
  for (var i = 0; i < members_mtrx.length; i++) {
    if (members_mtrx[i][0] == charge){
      bccAddress.push(members_mtrx[i][2]);
    }
  }
  // Logger.log(bccAddress);
  return bccAddress.join(',');
}

//仮予約があった際に、カレンダーに書き込む関数
function sendToCalendar(e) {
  var sheet = ss.getActiveSheet();
  try{
    //実験情報の取得
    var numRow = sheet.getActiveRange().getRow() // 新規仮予約された行番号を取得。active行だけを取得するようにする。手動で参加者情報を追加しても対応できる。

    var ParticipantName = sheet.getRange(numRow, colParName).getValue(); //新規仮予約された行から名前を取得し、"仮予約：参加者名"の文字列を作る
    var eventTitle = "仮予約:" + ParticipantName

    var cal = CalendarApp.getCalendarById(expInfo['experimenterMailAddress']); //仮予約を記載するカレンダーを取得
    var expDT = getExpDateTime(sheet, numRow);
    var from = expDT['from']; //仮予約の開始時間を取得
    var to = expDT['to'];//仮予約の開始時間から終了時間を設定

    //重複の確認
    var allEvents = cal.getEvents(from, to);
    if (allEvents.length > 0) {
      var trigger = '重複';
      var operationDict = {colStatus:trigger, colMailed:1, colRemindDate:'N/A', colReminded:'N/A'};
    } else {
      var trigger = '仮予約'
      var operationDict = {colStatus:'', colMailed:'', colRemindDate:'', colReminded:''};
      cal.createEvent(eventTitle, from, to); //仮予約情報をカレンダーに作成
    }

    var ParticipantEmail = sheet.getRange(numRow, colAddress).getValue();
    sendEmail(ParticipantName, ParticipantEmail, from, to, trigger)
    modifySheet(sheet, numRow, operationDict)
  } catch(exp) {
    //実行に失敗した時に通知
    MailApp.sendEmail(expInfo['experimenterMailAddress'], exp.message, exp.message);
    Logger.log(exp.message);
  }
}

// スプレッドシート上で予約を完了させ、メール送信及びカレンダーへの書き込みを行う関数
function updateCalendar(e) {
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
      var ParticipantName = sheet.getRange(activeRow, colParName).getValue();
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
            if (reserve[i].getTitle() == "仮予約:" + ParticipantName) {
              reserve[i].deleteEvent();
            }
          }
          if (trigger === 111) {
            // 変更した行から名前を取得し、"予約完了：参加者名(ふりがな)"の文字列を作る
            var eventTitle = "予約完了:" + ParticipantName +'('+sheet.getRange(activeRow, colParNameKana).getValue()+')';
            cal.createEvent(eventTitle, from, to); //予約確定情報をカレンダーに追加
            // リマインダーのための設定をする
            var remindDate = new Date(from).setDate(from.getDate() - 1);
            var time = new Date().setHours(19); //現在時刻の取得
            // var remindDate = new Date(from)
            // remindDate.setDate(from.getDate() - 1); //remindDateの時刻を予約時間の1日前に設定する。
            // var time = new Date(); //現在時刻の取得
            // time.setHours(19); //19時に設定
            // 予約を完了させた日の19時にremindDateの時刻が達していない場合、"送信準備"というコードを指定のセルに入力する
            if (remindDate > time) {
              var reminderStatus = "送信準備";
            } else {
              var reminderStatus = "翌日のため省略";
            }
            var insertValues = [1, remindDate, reminderStatus];
          }
          // triggerが111以外のとき
          else {
            var insertValues = [1,'N/A','N/A'];
          }
          var operationDict = {};
          for (i = 0; i < insertValues.length; i++){
            operationDict[activeColumn+i+1] = insertValues[i];
          }
          // メールの送信
          var ParticipantEmail = sheet.getRange(activeRow, colAddress).getValue();
          sendEmail(ParticipantName, ParticipantEmail, from, to, trigger)
          modifySheet(sheet, activeRow, operationDict)
        }
      }
    }
  } catch(exp) {
    //実行に失敗した時に通知
    Logger.log(exp)
    //MailApp.sendEmail(expInfo['experimenterMailAddress'], exp.message, exp.message);
  }
}

//リマインダーを実行する関数
function sendReminders(e) {
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
          var ParticipantName = data[row][colParName - 1]; //被験者の名前
          //参加者にメールを送る
          var ParticipantEmail = data[row][colAddress - 1];
          var trigger = 'リマインダー';
          var expDT = getExpDateTime(sheet, row + 1);
          var from = expDT['from'];
          "リマインダーでは to を使っていない"
          sendEmail(ParticipantName, ParticipantEmail, from, to, trigger)

          // シートの修正
          var reminded = sheet.getRange(row + 1, colReminded).setValue('送信済み');
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
  var sheets = ss.getSheets();
  var sheet = sheets[0];
  var colNames = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var addColNames = [['予約ステータス', '連絡したか', 'リマインド日時', 'リマインドしたか', '担当']];

  if (sheets.length < 2) {
    ss.insertSheet('設定');
    ss.insertSheet('テンプレート');
    ss.insertSheet('メンバー');
    var newColNames = [colNames[0].concat(addColNames[0])]
    sheet.getRange(1, 1, 1, newColNames[0].length).setValues(newColNames); //,444=二重登録,999=予約キャンセル)');
  } else {
    sheet.getRange(1, colNames[0].length - addColNames[0].length + 1, 1, addColNames[0].length).setValues(addColNames); //,444=二重登録,999=予約キャンセル)');
  }
  lastCol = sheet.getLastColumn();
  // 設定シート
  var config = ss.getSheetByName('設定');
  var note1 = '必要に応じて変更してください';
  var note2 = '「フォームの回答」に合わせて値を変更してください。この項目はメールでは使用されませんが，keyは変更しないでください。';
  var defaultExpInfo = [['設定項目','メール本文内でのkey','値','備考'],
                        ['実験責任者名','experimenterName','実験太郎', note1],
                        ['実験責任者のGmailアドレス','experimenterMailAddress','hogehoge@gmail.com',note1],
                        ['実験責任者の電話番号','experimenterPhone','xxx-xxx-xxx',note1],
                        ['実験の実施場所','experimentRoom','実施場所',note1],
                        ['実験の所要時間','experimentLength','実験の所要時間',note1],
                        ['参加者名の列番号','colParName', 2, note2],
                        ['ふりがなの列番号','colParNameKana', 3, note2],
                        ['参加者アドレスの列番号','colAddress',lastCol - 7, note2],
                        ['希望日の列番号','colExpDate', lastCol - 6, note2],
                        ['希望時間の列番号','colExpTime', lastCol - 5, note2]];

  var configNRow = defaultExpInfo.length;
  var configNCol = defaultExpInfo[0].length;
  config.getRange(1, 1, configNRow, configNCol).setValues(defaultExpInfo);


  // メールのテンプレート用シート
  var template = ss.getSheetByName('テンプレート');

  // successful 仮予約
  var TentativeBooking = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                          '予約の確認メールを自動で送信しております。\n','WHEN','で予約を受け付けました（まだ確定はしていません。)',
                          '後日、予約完了のメールを送信いたします。','もし日時の変更等がある場合は experimenterMailAddress までご連絡ください。',
                          'どうぞよろしくお願いいたします。\n','experimenterName'];
  // --- Failed 仮予約シリーズ ---
  // 時間外・期間外
  var Outoftime = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。',
                   'この度は心理学実験への応募ありがとうございました。','申し訳ありませんが、ご希望いただいた','WHEN',
                   'は実験実施可能時間（openTime時〜closeTime時）外または、実施期間（openMonth月openDate日〜closeMonth月closeDate日）外です。',
                   'お手数ですが、もう一度登録し直していただきますようお願いします。\n',
                   'experimenterName'];
  // 重複
  var Overlap = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                 '申し訳ありませんが、ご希望いただいた','WHEN',
                 'にはすでに予約（予定）が入っており（タッチの差で他の方が予約をされた可能性もあります）、実験を実施することができません。',
                 'お手数ですが、もう一度別の日時で登録し直していただきますようお願いします。\n','experimenterName'];
  // --- Successful Booking ---
  // 予約完了テキスト(平日)
  var WeekdayBookingDone = ['PARTICIPANTNAME 様\n','この度は心理学実験への応募ありがとうございました。',
                            'WHENからの心理学実験の予約が完了しましたのでメールいたします。',
                            '場所はexperimentRoomです。当日は直接お越しください。',
                            'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
                            '当日もよろしくお願いいたします。\n',
                            '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)','当日の連絡はexperimenterPhoneまでお願いいたします。'];
  // 予約完了テキスト(休日)
  var HolidayBookingDone = ['PARTICIPANTNAME 様\n','この度は心理学実験への応募ありがとうございました。',
                            'WHENからの心理学実験の予約が完了しましたのでメールいたします。',
                            '場所はexperimentRoomです。休日は教育学部棟玄関の鍵がかかっており、外から入ることができません。実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。',
                            'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。','当日もよろしくお願いいたします。\n',
                            '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)','当日の連絡はexperimenterPhoneまでお願いいたします。'];
  // --- Rejected Booking ---
  // 既参加
  var AlreadyDone = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                     '大変申し訳ありませんが、以前実施した同様の実験にご参加いただいており、今回の実験にはご参加いただけません。ご了承ください。\n',
                     'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。','今後ともよろしくお願いします。\n','experimenterName'];
  // 定員オーバー
  var ReachedCapacity = ['PARTICIPANTNAME 様\n','心理学実験実施責任者のexperimenterNameです。','この度は心理学実験への応募ありがとうございました。',
                         '大変申し訳ありませんが、応募いただいた段階ですでに募集人数の定員に達していたため、実験に参加していただくことができません。ご了承ください。\n',
                         '今後、次の実験を実施する際に再度応募していただけると幸いです。','ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
                         '今後ともよろしくお願いいたします。\n','experimenterName'];

  // --- Reminders ---
  // リマインダー(平日)
  var ReminderWeekday =['PARTICIPANTNAME 様\n','実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
                        '明日 WHENから実験に参加していただく予定となっております。','場所はexperimentRoomです。実験時間に実験室まで直接お越しください。\n',
                        'なお、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。','ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
                        'それでは明日、よろしくお願いいたします。\n','experimenterName'];
  // リマインダー(休日)
  var ReminderHoliday =['PARTICIPANTNAME 様\n','実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
                        '明日 WHENから実験に参加していただく予定となっております。','場所はexperimentRoomです。\n',
                        'なお、明日は休日のため教育学部棟玄関の鍵がかかっており、外から入ることができません。実験者が実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n',
                        'また、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。',
                        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。','それでは明日、よろしくお願いいたします。\n','experimenterName'];

  var notUsed = '利用する場合はここに本文を記載するとともに土日での変更の数字を1に変えてください。'

  var note = '必要に応じて変更してください'

  // これでいけるかも
  var bodies = {
    "仮予約":TentativeBooking,
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
  var defaultMember = [['キー', '名前', 'アドレス', '担当回数', '備考'],
                       [1, 'りんご', 'apple@hogege.com', '','Gmailのアドレスでなくても大丈夫です。'],
                       [2, 'ごりら', '実験実施可能時間外です','',''],
                       [3, 'らっぱ', '予約が重複しています','','']];
  var memNRow = defaultMember.length;
  var memNCol = defaultMember[0].length;
  member.getRange(1, 1, memNRow, memNCol).setValues(defaultMember);
}
