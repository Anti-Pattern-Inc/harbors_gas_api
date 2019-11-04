// https://github.com/Anti-Pattern-Inc/harbors_gas_api
function doPost(e) {
  try {
    let addData: any[] = [];
    const timeStamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');;
    addData.push(timeStamp);
    const sheetName: string[] = e.parameter['sheetName'];
    let keys: string[] = [];
    switch(sheetName) {
      case 'HarborSコワーキング会員':
        keys = [
          "name",
          "company_name",
          "mail",
          "tel",
          "preferred_visit_date",
          "frequency",
          "remarks"
        ];
        break;
      case 'バーチャルオフィス会員':
        keys = [
          "name",
          "company_name",
          "mail",
          "tel",
          "preferred_visit_date",
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
          "remarks"
        ];
        break;
      default:
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

    const result = {
      message: "success"
    };
    const out = ContentService.createTextOutput();
    //Mime TypeをJSONに設定
    out.setMimeType(ContentService.MimeType.JSON);
    //JSONテキストをセットする
    out.setContent(JSON.stringify(result));

    return out;
  } catch (error) {
    const result = {
      message: "failed"
    };
    const out = ContentService.createTextOutput();
    //Mime TypeをJSONに設定
    out.setMimeType(ContentService.MimeType.JSON);
    //JSONテキストをセットする
    out.setContent(JSON.stringify(result));
  }
}
