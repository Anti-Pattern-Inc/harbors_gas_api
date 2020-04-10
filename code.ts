// https://github.com/Anti-Pattern-Inc/harbors_gas_api
function doPost(e: { parameter: { [x: string]: any; }; }): any {
  
  //開始
  putlog("開始");
  
  //contact@harbors.sh（harborsお問い合わせスタッフ） のカレンダーID
  const CALENDAR_CONTACT_ID = PropertiesService.getScriptProperties().getProperty('CALENDAR_CONTACT_ID');
  
  try {
    let addData = [];
    const timeStamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    addData.push(timeStamp);
    const sheetName = e.parameter['sheetName'];
    let keys = [];
    let reserved = false;
    let eventName = "";
    
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
        eventName = "HarborSコワーキング見学予約";
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
        eventName = "バーチャルオフィス見学予約";
        reserved = true;
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
        eventName = "testGas見学予約";
        reserved = true;
        break;
      default:
        throw new Error("イベント名不正[" + sheetName + "]");
    }
    
    //見学予約
    if (reserved == true){
      putlog(eventName);
      // slack通知
      postMessageToContactChannel('<!channel>「' + eventName + '」に申し込みがありました。');
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
        " StartDate:" + Utilities.formatDate(startDate,"Asia/Tokyo","yyyy/MM/dd hh:mm:ss") + 
          " EndDate:" + Utilities.formatDate(endDate,"Asia/Tokyo","yyyy/MM/dd hh:mm:ss"));
      
      // 指定日時に予定が既にある場合は、予約済みステータスをセット
      if (existEventInCalendar(calendarContact, startDate, endDate) == true) {
        putlog("reserved");
        return result("reserved");
      }
      
      //カレンダー登録
      const event = calendarContact.createEvent(eventName, startDate, endDate);
      putlog(eventName + " Id:" + event.getId());
    }
    
    for (let _i = 0, keys_1 = keys; _i < keys_1.length; _i++) {
      const key = keys_1[_i];
      if (e.parameter[key]) {
        addData.push(e.parameter[key]);
        continue;
      }
      addData.push("");
    }
    
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    // シートへの書き込み、getRange(開始行、開始列、行数、列数)
    sheet.appendRow(addData);
    
    return result("success");
  }
  catch (error) {
    putlog(error);
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
  
  const timeStamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const stlog = SpreadsheetApp.getActive().getSheetByName("log");
  const addData = [];
  addData.push(timeStamp);
  addData.push(msg);
  
  stlog.appendRow(addData);
  
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
      "username" : '見学予約フォームbot',  // 通知時に表示されるユーザー名
      "icon_emoji": ':robot_face:',  // 通知時に表示されるアイコン
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