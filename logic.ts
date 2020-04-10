/*　
    ーーーーーーーーーーーーーーーーー
  　　　予約処理のロジックに関わる関数
    ーーーーーーーーーーーーーーーーー
*/　

/** 
 * カレンダーの指定日時にイベントがあるかチェック
 * @param  {GoogleAppsScript.Calendar.Calendar} calender カレンダー情報
 * @param  {Date}     startDate 開始日時
 * @param  {Date}     endDate   終了日時
 * @return {boolean} 指定日時にイベントが存在すればtrue、なければfalse
 */
export const existEventInCalendar = (
    calendar: GoogleAppsScript.Calendar.Calendar,
    startDate: Date,
    endDate: Date
 ): boolean => {
 
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
 
 