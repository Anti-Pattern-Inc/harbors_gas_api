function doPost(e) {
  try {
    const keys: string[] = [
      "name",
      "company_name",
      "mail",
      "tel",
      "preferred_visit_date",
      "remarks"
    ];
    let addData: any[] = [];
    for (let key of keys) {
      if (e.parameter[key]) {
        addData.push(e.parameter[key]);
        continue;
      }
      addData.push("");
    }
    const sheet = SpreadsheetApp.getActive().getSheetByName("testGas");
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
