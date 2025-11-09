function fillPickups() {
 const spread = SpreadsheetApp.getActiveSpreadsheet();
 spread.setSpreadsheetTimeZone("Europe/Athens"); // explicitly set Athens timezone
 const emails = spread.getSheetByName("Sheet1");
 const pickups = spread.getSheetByName("list");


 const lastRowPickups = pickups.getLastRow();
 const dataA = pickups.getRange(1, /* num of column A */1, lastRowPickups).getValues(); // column A, ids
 const dataF = pickups.getRange(1, 6, lastRowPickups).getValues(); // column F, outbound/inbound
 const dataM = pickups.getRange(1, 13, lastRowPickups).getValues(); // column M, pickup times


 let row = 2;
 let logRow = 2;


 while (row <= lastRowPickups) {
   let id = dataA[row - 1][0];
   let transport = dataF[row - 1][0];
   if (!id || id == "" || transport != "outbound") {row++; continue;}
   const finder = emails.createTextFinder(String(id)).matchEntireCell(true);
   const match = finder.findNext();   // finds first match
   if (!match) {emails.getRange("T" + logRow).setValue("ID not found in emails sheet: " + id); logRow++; row++; continue;}
   emails.getRange("Q" + match.getRow()).setValue(dataM[row-1][0]).setNumberFormat("HH:mm"); 
   row++;
 }
}

