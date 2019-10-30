function doPost(e) {
  //name,company_name,mail,tel,preferred_visit_date,remarks
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
    addData.push(e.parameters[key]);
  }
  const sheet = SpreadsheetApp.getActive().getSheetByName("testGas");

  // シートへの書き込み、getRange(開始行、開始列、行数、列数)
  sheet.appendRow(addData);
}
