// Google tabulka → Rozšíření → Apps Script → vložit → Nasadit → Nová implementace
// Typ: Webová aplikace, Spustit jako: já, Přístup: kdokoli.
// Vzniklou URL (končí /exec) vložit do konstanty ENDPOINT v index.html.
// První řádek tabulky: Čas | Jméno | E-mail | Telefon | Vzkaz

function doPost(e) {
  var p = e.parameter;
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].appendRow([
    new Date(), p.jmeno, p.email, p.telefon, p.poznamka
  ]);
  return ContentService.createTextOutput("ok");
}
