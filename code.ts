// https://github.com/Anti-Pattern-Inc/harbors_gas_api
function doPost(e: { parameter: { [x: string]: any; }; }): any {

  let eventName: string = "";
  //開始
  putlog("開始");

  try {
    let addData: any[] = [];
    const timeStamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    addData.push(timeStamp);
    const sheetName: string = e.parameter['sheetName'];
    let keys: string[] = [];
    let meilTemplateId: string = ''         // メールテンプレートID
    let slackMessage: string = '';         // Slack通知用メッセージ
    let sheetsRangeName: string = '';      // 添付ファイルのIDの設定場所

    const ret = checkContactType(e.parameter['contact_type']);
    const meet = ret.meet;
    const reserved = ret.reserved;
    const sendfile = ret.sendfile;
    const contactTypeMessage = ret.contactTypeMessage;
    e.parameter['contact_type'] = contactTypeMessage; //表示用内容書き換え
    e.parameter['tel'] = '\'' + e.parameter['tel'] ;  //前０が消えるので文字列に変換

    eventName = sheetName;

    switch (sheetName) {
      case 'コワーキング会員':
        keys = [
          "name",
          "mail",
          "tel",
          "contact_type",
          "preferred_visit_date",
          "preferred_visit_time",
          "frequency",
          "remarks"
        ];
        eventName = editEventName(reserved, meet, sendfile, "コワーキングスペース");
        meilTemplateId = PropertiesService.getScriptProperties().getProperty('RESERVE_CONFIRMATION_TEMPLATE');
        sheetsRangeName = 'CW_EMAIL_ATTACH_FILEID';
        break;
      case 'バーチャルオフィス会員':
        keys = [
          "name",
          "company_name",
          "mail",
          "tel",
          "contact_type",
          "preferred_visit_date",
          "preferred_visit_time",
          "remarks"
        ];
        eventName = editEventName(reserved, meet, sendfile, "バーチャルオフィス");
        meilTemplateId = PropertiesService.getScriptProperties().getProperty('RESERVE_CONFIRMATION_TEMPLATE');
        sheetsRangeName = 'VO_EMAIL_ATTACH_FILEID';
        break;
      case 'HarborSLP':
        keys = [
          "name",
          "mail",
          "tel",
          "frequency",
          "trigger",
          "remarks"
        ];
        break;
      case 'バーチャルLP':
        keys = [
          "name",
          "mail",
          "tel",
          "frequency",
          "trigger",
          "remarks"
        ];
        break;
      case 'extends':
        keys = [
          "name",
          "mail",
          "tel",
          "contact_type",
          "preferred_visit_date",
          "preferred_visit_time",
          "remarks"
        ];
        eventName = editEventName(reserved, meet, sendfile, "extends");
        meilTemplateId = PropertiesService.getScriptProperties().getProperty('RESERVE_CONFIRMATION_TEMPLATE_EXTENDS');
        sheetsRangeName = 'EX_EMAIL_ATTACH_FILEID';
        break;
      case 'testGas':
        keys = [
          "name",
          "company_name",
          "mail",
          "tel",
          "preferred_visit_date",
          "preferred_visit_time",
          "remarks"
        ];
        eventName = "testGas見学";
        break;
      default:
        throw new Error("イベント名不正[" + sheetName + "]");
    }

    //見学予約
    if (reserved == true){
      //オンライン説明会を予約したい
      //現地の見学説明会を予約したい
      const startDate = new Date(e.parameter['preferred_visit_date'] + " " + e.parameter['preferred_visit_time']); //予約開始日
      const endDate = new Date(e.parameter['preferred_visit_date'] + " " + e.parameter['preferred_visit_time']);
      endDate.setHours(endDate.getHours() + 1);//予約終了日（開始＋１時間）

      //カレンダー予約／Meet設定
      const ret = reservedEvent(eventName, e.parameter['name'], startDate, endDate, meet);
      if (ret.reserved == true){
        // 予約済み
        return result("reserved");
      }

      //予約成功のメール送信
      try{
        sendReserveMail(e.parameter['mail'],
                        e.parameter['name'],
                        eventName,
                        e.parameter['preferred_visit_date'],
                        e.parameter['preferred_visit_time'],
                        PropertiesService.getScriptProperties().getProperty('AP_CONTACT_EMAIL'),
                        ret.webUrl,
                        meilTemplateId
                        );
      }catch(error){
        putlog(error);
        throw new Error('メール送信エラー(' + error + ')');
      }

      //Slack通知用メッセージ作成
      slackMessage = '<!channel>「' + eventName + '」に申込みがありました。\n' +
          e.parameter['name']+ ' 様' +
          '\n予約日時：' + Utilities.formatDate(startDate,"Asia/Tokyo","yyyy/MM/dd HH:mm") +
          '\n問合せ内容:' + contactTypeMessage;

      // slack通知
      try{
        postMessageToContactChannel(slackMessage);
      }catch(error){
        putlog(error);
        throw new Error('slack送信エラー(' + error + ')');
      }
    }else if(sendfile == true){
      //資料を送付してほしい
      try{
        sendAttachEmail(e.parameter['mail'],
                    e.parameter['name'],
                    eventName,
                    PropertiesService.getScriptProperties().getProperty('AP_CONTACT_EMAIL'),
                    sheetsRangeName,
                    PropertiesService.getScriptProperties().getProperty('ATTACH_EMAIL_TEMPLATE')
                    );
      }catch(error){
        putlog(error);
        throw new Error('メール送信エラー(' + error + ')');
      }

      //Slack通知用メッセージ作成
      slackMessage = '<!channel>「' + eventName + '」に問合せがありました。\n' +
          e.parameter['name']+ ' 様\n' +
          '問合せ内容:' + contactTypeMessage;
      // slack通知
      try{
        postMessageToContactChannel(slackMessage);
      }catch(error){
        putlog(error);
        throw new Error('slack送信エラー(' + error + ')');
      }
    }

    for (let key of keys) {
      if (e.parameter[key]) {
        addData.push(e.parameter[key]);
        continue;
      }
      addData.push("");
    }

    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    // シートへの書き込み、getRange(開始行、開始列、行数、列数)
    sheet.appendRow(addData);

    return result("success");
  }
  catch (error) {
    putlog(error);
    postMessageToContactChannel('<!channel>「' + eventName + '」の申込でエラーが発生しました。\n```エラー内容:' + error + '\nパラメータ:' + JSON.stringify(e.parameter) + '```');
    return result("failed");
  }
}

/**
 * カレンダー予約を行う
 * @param eventName イベント名
 * @param name 予約者
 * @param startDate 予約開始日時
 * @param endDate 予約終了日時
 * @param meet Meet予約フラグ
 */
function reservedEvent(eventName: string, name: string , startDate: Date, endDate: Date, meet: boolean){
  let webUrl: string = ''

  //contact@harbors.sh（harborsお問い合わせスタッフ） のカレンダーID
  const CALENDAR_CONTACT_ID = PropertiesService.getScriptProperties().getProperty('CALENDAR_CONTACT_ID');
  try{
    // カレンダーIDでカレンダーを取得
    const calendarContact = CalendarApp.getCalendarById(CALENDAR_CONTACT_ID);
    if(calendarContact==null){
      putlog("カレンダーオブジェクト取得失敗");
      throw new Error("カレンダーオブジェクト取得失敗");
    }

    //予約情報
    putlog("Name:" + name +
      " StartDate:" + Utilities.formatDate(startDate,"Asia/Tokyo","yyyy/MM/dd HH:mm:ss") +
        " EndDate:" + Utilities.formatDate(endDate,"Asia/Tokyo","yyyy/MM/dd HH:mm:ss"));

    // 指定日時に予定が既にある場合は、予約済みステータスをセット
    if (existEventInCalendar(calendarContact, startDate, endDate) == true) {
      putlog("reserved");
      return {reserved: true, webUrl: ''}
    }

    //カレンダー登録
    let noticeName: string = eventName + "-" + name + "様";
    let event = createEventToCalendar(noticeName, startDate, endDate, meet);
    if (meet){
      //MeetのUrlを取得
      webUrl = event.conferenceData.entryPoints[0].uri;
    }

    return {reserved: false, webUrl: webUrl};
  }catch(error){
    throw new Error('カレンダー予約エラー(' + error + ')');
  }
}

/**
 * 申込種別判定
 * @param contactType フォームから送信される申込種別
 */
function checkContactType(contactType: string) {
  let reserved: boolean = false;
  let meet: boolean = false;
  let sendfile: boolean = false;
  let contactTypeMessage: string = "";

  switch(contactType){
    case '1':
      reserved = true;
      meet = true;
      contactTypeMessage = "オンライン説明会を予約したい";
      break;
    case '2':
      reserved = true;
      contactTypeMessage = "現地の見学説明会を予約したい";
      break;
    case '3':
      sendfile = true;
      contactTypeMessage = "資料を送付してほしい";
      break;
    case '': //未設定はOK
      break;
    default: //1-3,''未設定以外の場合NG
      throw new Error("問い合わせ種別不正[" + contactType + "]");
  }

  return { reserved, meet , sendfile, contactTypeMessage };
}

/**
 * イベント名を編集する
 * @param reserved 予約設定
 * @param meet meet設定
 * @param sendfile ファイル送信
 * @param eventName イベント名
 */
function editEventName(reserved: boolean, meet: boolean, sendfile: boolean, eventName: string): string{
  if (meet == true) return eventName + 'オンライン説明会';
  if (reserved == true) return eventName + '現地見学説明会';
  if (sendfile == true) return eventName;
  return eventName;
}
/*
クライアントへのレスポンス
*/
function result(msg: string): GoogleAppsScript.Content.TextOutput{
  const result = {
    message: msg
  };

  const out = ContentService.createTextOutput();
  //Mime TypeをJSONに設定
  out.setMimeType(ContentService.MimeType.JSON);
  //JSONテキストをセットする
  out.setContent(JSON.stringify(result));

  return out;
}

function putlog(msg: string): void{
  /* デバッグ用
  const timeStamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const stlog = SpreadsheetApp.getActive().getSheetByName("log");
  const addData = [];
  addData.push(timeStamp);
  addData.push(msg);

  stlog.appendRow(addData);
  */
  console.info(msg);
}

/**
 * カレンダーの指定日時にイベントがあるかチェック
 * @param  {GoogleAppsScript.Calendar.Calendar} calender カレンダー情報
 * @param  {Date}     startDate 開始日時
 * @param  {Date}     endDate   終了日時
 * @return {boolean} 指定日時にイベントが存在すればtrue、なければfalse
 */
function existEventInCalendar(
    calendar: GoogleAppsScript.Calendar.Calendar,
    startDate: Date,
    endDate: Date
  ): boolean {

    // 変数eventsは「CalendarEvent」を持つ配列
    const events = calendar.getEvents(startDate, endDate);

    console.log('イベント重複数 %d', events.length);

    // イベントがなければ、falseを返却
    if (events.length < 1) {
      return false;
    }
    // イベントが一つでもあれば、trueを返却
    return true;
}

/**
 * slackのチャンネルにメッセージを投稿する
 * @param  {string} message 投稿メッセージ
 * @return {void}
 */
function postMessageToContactChannel(message: string): void {
  // #contantへのwebhook URLを取得
  const webhookURL = PropertiesService.getScriptProperties().getProperty('WEBHOOK_URL');
  // 投稿に必要なデータを用意
  const jsonData =
  {
    "text": message  // 投稿メッセージ
  };
  // JSON文字列に変換
  const payload = JSON.stringify(jsonData);

  // 送信オプションを用意
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    contentType: "application/json",
    payload: payload
  }

  UrlFetchApp.fetch(webhookURL, options);
}

/**
 * 予約完了メールを送信する
 * @param {string} mailAddress 送信先アドレス
 * @param {string} contactName 予約者者名
 * @param {string} eventName   予約イベント名
 * @param {string} visitDate   予約日利用日
 * @param {string} visitTime   利用開始時間
 * @param {string} carbonCopyMail CCの送信先
 * @param {string} webUrl       Meetの接続先
 * @param {string} meilTemplateId テンプレートID
 * @return void
 */
function sendReserveMail(mailAddress: string,
                        contactName: string,
                        eventName: string,
                        visitDate: string,
                        visitTime: string,
                        carbonCopyMail: string,
                        webUrl: string,
                        meilTemplateId:string): void{

  // メールオプション
  const option = {from: 'contact@harbors.sh', name: 'HarborS運営スタッフ', cc:carbonCopyMail};
  // 件名
  const title = "[HarborS表参道]" + eventName + "申込のお知らせ";
  //　予約完了メールのテンプレートをドキュメントより取得
  const document = DocumentApp.openById(meilTemplateId);
  const bodyTemplate = document.getBody().getText();
  // 氏名をセット
  let body = bodyTemplate.replace("%contactName%", contactName);
  // イベントをセット
  body = body.replace("%eventName%", eventName);
  // 予約日をセット
  body = body.replace("%visitDate%", visitDate);
  // 予約時間をセット
  body = body.replace("%visitTime%", visitTime);
  // Web会議のURL
  if(webUrl==''){
    body = body.replace("%meetAddress%", '');
  }else{
    body = body.replace("%meetAddress%", '会議URL： ' + webUrl);
  }

  GmailApp.sendEmail(mailAddress, title, body, option);
}

/**
 * 資料送付メールを送信する
 * @param {string} mailAddress 送信先アドレス
 * @param {string} contactName 予約者者名
 * @param {string} eventName   予約イベント名
 * @param {string} carbonCopyMail CCの送信先
 * @param {string} sendFileId     送信ファイルの名前付き範囲の定義名
 * @param {string} meilTemplateId テンプレートID
 * @return void
 */
function sendAttachEmail(mailAddress: string,
                        contactName: string,
                        eventName: string,
                        carbonCopyMail: string,
                        sendFileId: string,
                        meilTemplateId: string): void{

  //添付ファイルを取得
  let arrReport:Array<GoogleAppsScript.Drive.File> = [];
  try{
    const  itemValues = SpreadsheetApp.getActive().getRangeByName(sendFileId).getValues();
    for (let item of itemValues){
      arrReport.push(DriveApp.getFileById(item.toString()));
    }
  }catch(error){
    throw new Error('添付ファイル取得エラー(' + error + ')');
  }

  // メールオプション
  const option = {from: 'contact@harbors.sh', name: 'HarborS運営スタッフ', cc:carbonCopyMail , attachments:arrReport};

  // 件名
  const title = "[HarborS表参道 資料請求]" +eventName + "に関する資料を送付致します";

  //　予約完了メールのテンプレートをドキュメントより取得
  const document = DocumentApp.openById(meilTemplateId);
  const bodyTemplate = document.getBody().getText();
  // 氏名をセット
  const body = bodyTemplate
                .replace(/%contactName%/g, contactName)
                .replace(/%eventName%/g, eventName);

  GmailApp.sendEmail(mailAddress, title, body, option);
}

/**
 *
 * @param noticeName
 * @param startDate
 * @param endDate
 * @param meet
 * @returns GoogleAppsScript.Calendar
 */
function createEventToCalendar(noticeName: string,
                              startDate: Date,
                              endDate: Date,
                              meet: boolean): GoogleAppsScript.Calendar.Schema.Event{
  // テレビ会議利用のため
  // Googel Calendar API にてカレンダー登録
  const CALENDAR_CONTACT_ID = PropertiesService.getScriptProperties().getProperty('CALENDAR_CONTACT_ID');

  // 作成するカレンダーのイベント定義
  const event: GoogleAppsScript.Calendar.Schema.Event  = {
    summary: noticeName,
    start: {
      dateTime: Utilities.formatDate(startDate,"Asia/Tokyo","yyyy-MM-dd'T'HH:mm:ssZ")
    },
    end: {
      dateTime: Utilities.formatDate(endDate,"Asia/Tokyo","yyyy-MM-dd'T'HH:mm:ssZ")
    },
    conferenceData: {
      createRequest: {
        conferenceSolutionKey: {
          type: "hangoutsMeet"
        },
        requestId: PropertiesService.getScriptProperties().getProperty('CALENDAR_REQUEST_ID')
      }
    }
  }

  if (meet == true){
    return Calendar.Events.insert(event, CALENDAR_CONTACT_ID, {conferenceDataVersion: 1})
  }else{
    return Calendar.Events.insert(event, CALENDAR_CONTACT_ID)
  }
}
