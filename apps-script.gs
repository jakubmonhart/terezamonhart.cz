/**
 * Přihlášky z terezamonhart.cz do této tabulky.
 *
 * Nasazení: Rozšíření → Apps Script → vložit → Nasadit → Nová implementace
 * Typ: Webová aplikace, Spustit jako: já, Přístup: kdokoli.
 * Vzniklou URL (končí /exec) vložit do konstanty ENDPOINT v index.html.
 * První řádek tabulky: Čas | Jméno | E-mail | Telefon | Vzkaz | Potvrzení
 *
 * @OnlyCurrentDoc
 */

var KAPACITA = 0;   // největší počet přihlášených, 0 = bez omezení
var UPOZORNIT = ""; // kam chodí upozornění na novou přihlášku, prázdné = majitel tabulky

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var p = (e && e.parameter) || {};
    if (p.web) return odpoved({ ok: true }); // past na roboty, člověk pole nevidí

    var jmeno = cist(p.jmeno, 100);
    var email = cist(p.email, 150);
    var telefon = cist(p.telefon, 40);
    var vzkaz = cist(p.poznamka, 1000);
    if (!jmeno || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return odpoved({ ok: false, error: "udaje" });
    }

    var list = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (KAPACITA && list.getLastRow() - 1 >= KAPACITA) {
      return odpoved({ ok: false, error: "plno" });
    }

    list.appendRow([new Date(), jakoText(jmeno), jakoText(email), jakoText(telefon), jakoText(vzkaz), ""]);
    var radek = list.getLastRow();

    var potvrzeni = "odesláno";
    try {
      MailApp.sendEmail({
        to: email,
        subject: "Přihláška na workshop 15. 11. 2026",
        body:
          "Ahoj,\n\n" +
          "děkuju za přihlášku, počítám s tebou.\n\n" +
          "Kdy: neděle 15. 11. 2026\n" +
          "Kde: Centrum Jízdárna, Znojmo\n" +
          "Contemporary 14:30–16:00\n" +
          "POP-dance 16:15–17:45\n" +
          "Cena: 300 Kč\n\n" +
          "Kdyby se ti něco změnilo, odpověz prosím na tento e-mail.\n\n" +
          "Těším se,\nTereza"
      });
    } catch (err) {
      potvrzeni = "NEODESLÁNO";
    }
    list.getRange(radek, 6).setValue(potvrzeni);

    try {
      MailApp.sendEmail({
        to: UPOZORNIT || Session.getEffectiveUser().getEmail(),
        subject: "Nová přihláška: " + jmeno,
        body: jmeno + "\n" + email + "\n" + telefon + "\n\n" + vzkaz +
          "\n\nPotvrzení přihlášenému: " + potvrzeni
      });
    } catch (err) {}

    return odpoved({ ok: true });
  } catch (err) {
    return odpoved({ ok: false, error: "chyba" });
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

function cist(hodnota, max) {
  return String(hodnota || "").trim().slice(0, max);
}

// Tabulka by text začínající na = + - @ vyhodnotila jako vzorec.
function jakoText(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function odpoved(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
