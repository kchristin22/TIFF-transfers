function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()       // name that shows in Extensions
    .addItem("Group passengers for drivers", "drivers") // menu item
    .addItem("Fill Pickup Times", "fillPickups")
    .addToUi();
}

function drivers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("list");
  SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetTimeZone("Europe/Athens");
  const lastRow = sheet.getLastRow();
  let row = 2;
  let logRow = 2;

  const dataB = sheet.getRange(2, /* num of column B */2, lastRow-1).getValues(); // column B, name
  const dataC = sheet.getRange(2, 3, lastRow-1).getValues(); // column C, name
  const dataE = sheet.getRange(2, 5, lastRow-1).getValues(); // column E, hotel
  const dataF = sheet.getRange(2, 6, lastRow-1).getValues(); // column F, outbound/inbound
  const dataG = sheet.getRange(2, 7, lastRow-1).getValues(); // column G, transport
  const dataJ = sheet.getRange(2, 10, lastRow-1).getValues(); // column J, Bus station
  const dataL = sheet.getRange(2, 12, lastRow-1).getValues(); // column L, arrival
  const dataM = sheet.getRange(2, 13, lastRow-1).getValues(); // column M, pickup
  const dataN = sheet.getRange(2, 14, lastRow-1).getValues(); // column N, passengers


  while (row <= lastRow) {
  const b = dataB[row - 2][0]; // array index starts from  0
  const c = dataC[row - 2][0];
  if (!b && !c) {
    row++; continue;
  }

  let text = "";
  let current = row;
  const outputCell = sheet.getRange("O" + row);
  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();


  const valF = dataF[current - 2][0]; // inbound/outbound
  const valG = dataG[current - 2][0]; // type of transport
  if (valF == 'inbound') {
    valL = dataL[current - 2][0]; // hour of landing
    valL = (typeof valL =='object') ? Utilities.formatDate(valL, "Europe/Athens", "HH:mm") :
      String(valL).trim();
    text = "ΑΦΙΞΗ " + valL;
    if (valG == "Bus") text += " KTEL " + dataJ[current - 2][0];
    const startIndex = text.indexOf(valL);
    let rich = SpreadsheetApp.newRichTextValue().setText(text);
    if (startIndex !== -1) {
      rich.setTextStyle(startIndex, startIndex + valL.length, boldStyle);
    }
    outputCell.setRichTextValue(rich.build());


    while (current <= lastRow && dataF[current - 2][0]  == "inbound") current++;
    row = current;
  } else if (valF == "outbound"){
    if (valG == "Plane") {
      text = "ΠΡΟΣ ΑΕΡΟΔΡΟΜΙΟ ";
    } else if (valG == "Train" || dataJ[current - 2][0] == "Monastiriou") {
      text = "ΠΡΟΣ ΟΣΕ "
    } else if (dataJ[current - 2][0] == "Makedonias") text = "ΠΡΟΣ ΚΤΕΛ ΜΑΚΕΔΟΝΙΑΣ ";
    else {
      sheet.getRange("Q" + logRow).setValue("Bus station not specified properly in line " + current + ". Please use 'Monastiriou' or 'Makedonias'.");
      logRow++;
    }

    let valEprev = "";
    let boldRanges = [];

    while (current <= lastRow && dataF[current - 2][0]  == "outbound") {
        const valB = dataB[current - 2][0]; // first name
        const valC = dataC[current - 2][0]; // last name or passengers
        const valN = dataN[current - 2][0]; // passengers

        if (!valB && !valC) break;

        const valE = dataE[current - 2][0]; // hotel
        if (!valE) {sheet.getRange("Q" + logRow).setValue("No hotel for departure for row " + current); logRow++; current++; continue;}
        if (!valEprev || valE != valEprev) {
          if (text != "ΠΡΟΣ ΑΕΡΟΔΡΟΜΙΟ " && text != "ΠΡΟΣ ΟΣΕ " && text != "ΠΡΟΣ ΚΤΕΛ ΜΑΚΕΔΟΝΙΑΣ ") text += " + ";
          valM = dataM[current - 2][0];
          if (!valM) {sheet.getRange("Q" + logRow).setValue("No pickup time for row " + current); logRow++;}
          valM = (typeof valM =='object') ? Utilities.formatDate(valM, "Europe/Athens", "HH:mm") :
            String(valM).trim();
          const valE = dataE[current - 2][0];
          text += valM + " " + valE + " - " + valB + " " + valC;
          valEprev = valE;

          const startIndex = text.indexOf(valM);
          const endIndex = text.indexOf(valE) + valE.length;
          boldRanges.push([startIndex, endIndex]);
        } else {
          text += ", " + valB + " " + valC;
        }

        if (valN && valN != "1 passenger") {
          text += " (" + valN + ")";
        }  
        current++;
      }
      let rich = SpreadsheetApp.newRichTextValue().setText(text.trim());
      for (const [s, e] of boldRanges) {
        rich = rich.setTextStyle(s, e, boldStyle);
      }

      outputCell.setRichTextValue(rich.build());
      row = current;
    } else {
      sheet.getRange("Q" + logRow).setValue("Specify outbound/inbound for row " + current); logRow++;
      row++;
    }
  }
}
