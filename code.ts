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
  let addData: any[][] = [];
  for (let key of keys) {
    addData[0].push(e.parameters[key]);
  }
  const sheet = SpreadsheetApp.getActive().getSheetByName("testGas");
  const targetRow = sheet.getLastRow() + 1;

  // シートへの書き込み、getRange(開始行、開始列、行数、列数)
  sheet.getRange(targetRow, 1, 1, addData[0].length).setValues(addData);
}
