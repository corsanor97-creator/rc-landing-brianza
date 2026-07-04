/* ============================================================
   LEAD -> GOOGLE SHEET  (Google Apps Script)
   Riceve i dati del form della landing e scrive una riga nel
   foglio. Invia anche una email di notifica a ogni nuova lead.

   COME ATTIVARLO (una volta sola):
   1. Crea un nuovo Google Sheet (sheets.new).
   2. Menu: Estensioni -> Apps Script.
   3. Cancella tutto il codice di esempio e incolla QUESTO file.
   4. (Opzionale) Cambia NOTIFY_EMAIL con l'indirizzo dove vuoi
      ricevere gli avvisi. Lascia "" per non ricevere email.
   5. Salva (icona floppy).
   6. In alto: Distribuisci -> Nuova distribuzione.
   7. Tipo -> "App web".
        - Esegui come: Me
        - Chi ha accesso: Chiunque
   8. Distribuisci -> autorizza l'accesso (accetta i permessi).
   9. Copia l'"URL dell'app web" (finisce con /exec).
   10. Incolla quell'URL dentro index.html, nella riga:
         var SHEET_ENDPOINT = "...";
   11. Ricarica index.html + grazie.html su Vercel.

   NB: se in futuro modifichi questo script, vai su
   "Distribuisci -> Gestisci distribuzioni" e pubblica una
   nuova versione (l'URL resta lo stesso).
   ============================================================ */

var NOTIFY_EMAIL = "corsano.r97@gmail.com"; // <- avvisi via email (metti "" per disattivare)
var SHEET_NAME   = "Lead";                  // nome della scheda dove salvare

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // intestazioni alla prima esecuzione
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Data", "Nome e cognome", "Azienda", "P.IVA", "Email", "Telefono"]);
    }

    var p = (e && e.parameter) ? e.parameter : {};
    sheet.appendRow([
      new Date(),
      p.nome     || "",
      p.azienda  || "",
      p.piva     || "",
      p.email    || "",
      p.telefono || ""
    ]);

    if (NOTIFY_EMAIL) {
      var corpo =
        "Nuova richiesta dalla landing:\n\n" +
        "Nome e cognome: " + (p.nome     || "-") + "\n" +
        "Azienda: "        + (p.azienda  || "-") + "\n" +
        "P.IVA: "          + (p.piva     || "-") + "\n" +
        "Email: "          + (p.email    || "-") + "\n" +
        "Telefono: "       + (p.telefono || "-") + "\n";
      MailApp.sendEmail(NOTIFY_EMAIL, "Nuova lead dalla landing", corpo);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
