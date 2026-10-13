/**
 * Přihlášky z terezamonhart.cz do této tabulky. Přihlášenému se nic neposílá,
 * odpovídá mu Tereza sama; skript jí jen pošle upozornění na novou přihlášku.
 *
 * Nasazení: Rozšíření → Apps Script → vložit → Nasadit → Nová implementace
 * Typ: Webová aplikace, Spustit jako: já, Přístup: kdokoli.
 * Vzniklou URL (končí /exec) vložit do konstanty ENDPOINT v index.html.
 * Záhlaví (Čas | Lekce | Jméno | E-mail | Telefon | Poznámka) si skript
 * do prázdné tabulky doplní sám při první přihlášce.
 *
 * @OnlyCurrentDoc
 */

var KAPACITA = 0;   // největší počet přihlášených, 0 = bez omezení
var UPOZORNIT = ""; // kam chodí každá nová přihláška e-mailem, prázdné = majitel tabulky

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var p = (e && e.parameter) || {};
    if (p.web) return odpoved({ ok: true }); // past na roboty, člověk pole nevidí

    var lekce = cist(p.lekce, 100);
    var jmeno = cist(p.jmeno, 100);
    var email = cist(p.email, 150);
    var telefon = cist(p.telefon, 40);
    var vzkaz = cist(p.poznamka, 1000);
    if (!lekce || !jmeno || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return odpoved({ ok: false, error: "udaje" });
    }

    var list = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (list.getLastRow() === 0) {
      list.appendRow(["Čas", "Lekce", "Jméno", "E-mail", "Telefon", "Poznámka"]);
    }
    if (KAPACITA && list.getLastRow() - 1 >= KAPACITA) {
      return odpoved({ ok: false, error: "plno" });
    }

    list.appendRow([new Date(), jakoText(lekce), jakoText(jmeno), jakoText(email), jakoText(telefon), jakoText(vzkaz)]);
    try {
      MailApp.sendEmail({
        to: UPOZORNIT || Session.getEffectiveUser().getEmail(),
        replyTo: email,
        subject: "Nová přihláška: " + jmeno,
        body: lekce + "\n" + jmeno + "\n" + email + "\n" + telefon + "\n\n" + vzkaz
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
