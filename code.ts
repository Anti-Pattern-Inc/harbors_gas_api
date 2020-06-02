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
    let meet: boolean = false;
    let reserved: boolean = false;
    let meilTemplateId: string = ''
    eventName = sheetName;
    
    switch (sheetName) {
      case 'HarborSコワーキング会員':
        keys = [
          "name",
          "mail",
          "tel",
          "preferred_visit_date",
          "preferred_visit_time",
          "frequency",
          "remarks"
        ];
        reserved = true;
        eventName = "コワーキング見学";
        meilTemplateId = PropertiesService.getScriptProperties().getProperty('RESERVE_CONFIRMATION_TEMPLATE');
        break;
      case 'バーチャルオフィス会員':
        keys = [
          "name",
          "company_name",
          "mail",
          "tel",
          "preferred_visit_date",
          "preferred_visit_time",
          "remarks"
        ];
        reserved = true;
        eventName = "バーチャルオフィス見学";
        meilTemplateId = PropertiesService.getScriptProperties().getProperty('RESERVE_CONFIRMATION_TEMPLATE');
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
          "preferred_visit_date",
          "preferred_visit_time",
          "remarks"
        ];
        reserved = true;
        meet = true;
        eventName = "extendsオンライン説明会";
        meilTemplateId = PropertiesService.getScriptProperties().getProperty('RESERVE_CONFIRMATION_TEMPLATE_EXTENDS');
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
        reserved = true;
        break;
      default:
        throw new Error("イベント名不正[" + sheetName + "]");
    }
    
    //見学予約
    if (reserved == true){
      putlog(eventName);

      //contact@harbors.sh（harborsお問い合わせスタッフ） のカレンダーID
      const CALENDAR_CONTACT_ID = PropertiesService.getScriptProperties().getProperty('CALENDAR_CONTACT_ID');
  
      // カレンダーIDでカレンダーを取得
      const calendarContact = CalendarApp.getCalendarById(CALENDAR_CONTACT_ID); 
      if(calendarContact==null){
        putlog("カレンダーオブジェクト取得失敗");
        throw new Error("カレンダーオブジェクト取得失敗");
      }
      
      const startDate = new Date(e.parameter['preferred_visit_date'] + " " + e.parameter['preferred_visit_time']); //予約開始日
      const endDate = new Date(e.parameter['preferred_visit_date'] + " " + e.parameter['preferred_visit_time']);
      endDate.setHours(endDate.getHours() + 1);//予約終了日（開始＋１時間）
      
      //予約情報
      putlog("Name:" + e.parameter['name'] +
        " StartDate:" + Utilities.formatDate(startDate,"Asia/Tokyo","yyyy/MM/dd HH:mm:ss") + 
          " EndDate:" + Utilities.formatDate(endDate,"Asia/Tokyo","yyyy/MM/dd HH:mm:ss"));
          
      // 指定日時に予定が既にある場合は、予約済みステータスをセット
      if (existEventInCalendar(calendarContact, startDate, endDate) == true) {
        putlog("reserved");
        return result("reserved");
      }
      
      //カレンダー登録
      let noticeName: string = `${eventName}-${e.parameter['name']}様`;
      let webUrl: string = "";
      let event = createEventToCalendar(noticeName, startDate, endDate, meet);
      if (meet){
        //MeetのUrlを取得
        webUrl = event.conferenceData.entryPoints[0].uri;
      }      
      putlog(eventName + " Id:" + event.id);
      try{        
        //予約成功のメール送信
        sendReserveMail(e.parameter['mail'], 
                        e.parameter['name'], 
                        eventName, 
                        e.parameter['preferred_visit_date'], 
                        e.parameter['preferred_visit_time'],
                        PropertiesService.getScriptProperties().getProperty('AP_CONTACT_EMAIL'),
                        webUrl,
                        meilTemplateId
                        );
      }catch(error){
        putlog(error);
        throw new Error('メール送信エラー(' + error + ')');
      }

      try{        
        // slack通知
        postMessageToContactChannel('<!channel>「' + eventName + '」に申し込みがありました。' + 
                                    '\n' + e.parameter['name']+ ' 様 予約日時：' + Utilities.formatDate(startDate,"Asia/Tokyo","yyyy/MM/dd HH:mm"));
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
      "text" : message  // 投稿メッセージ
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
function sendReserveMail(mailAddress :string, 
                         contactName :string, 
                         eventName :string, 
                         visitDate :string, 
                         visitTime :string, 
                         carbonCopyMail: string,
                         webUrl :string, 
                         meilTemplateId :string) :void{

  // メールオプション
  const option = {from: 'contact@harbors.sh', name: 'HarborS運営スタッフ', cc:carbonCopyMail};
  // 件名
  const title = eventName + "申込のお知らせ";
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
  body = body.replace("%meetAddress%", webUrl);

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