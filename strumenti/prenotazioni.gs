/**
 * Circolandia – gestione richieste di prenotazione
 * Script Google Apps Script da collegare a un Foglio Google.
 *
 * Per ogni richiesta inviata dal sito:
 *   1. aggiunge una riga al foglio "Prenotazioni"
 *   2. manda un'email di avviso allo staff
 *   3. manda al cliente un'email di riepilogo con il volantino PDF allegato
 *
 * Istruzioni di installazione: vedi il messaggio di Claude oppure il README del repository.
 */

// ====== DA COMPILARE ======
var EMAIL_STAFF = "prenotazioni@circolandia.it";   // inoltrato da OVH alla tua casella personale
var NOME_SALA = "Circolandia";
var URL_VOLANTINO = "https://circolandia.it/volantino-circolandia.pdf";
// ==========================

var COLONNE = ["Ricevuta il", "Codice", "Stato", "Occasione", "Data festa", "Fascia", "Ora extra",
               "Bambini", "Adulti", "Festeggiato", "Anni", "Genitore", "Telefono", "Email", "Note"];

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (!d.nome || !d.email || !d.data || !d.fascia) throw new Error("dati mancanti");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(d.email)) throw new Error("email non valida");

    registra(d);
    avvisaStaff(d);
    confermaCliente(d);
    return risposta({ ok: true });
  } catch (err) {
    return risposta({ ok: false, errore: String(err.message || err) });
  }
}

function doGet() {
  return risposta({ ok: true, messaggio: "Servizio prenotazioni " + NOME_SALA + " attivo" });
}

function risposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function pulisci(v) { return String(v == null ? "" : v).slice(0, 500); }

function registra(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var foglio = ss.getSheetByName("Prenotazioni") || ss.insertSheet("Prenotazioni");
  if (foglio.getLastRow() === 0) {
    foglio.appendRow(COLONNE);
    foglio.getRange(1, 1, 1, COLONNE.length).setFontWeight("bold");
    foglio.setFrozenRows(1);
  }
  // l'apostrofo iniziale evita che Fogli interpreti il testo come formula
  var testo = function (v) { return "'" + pulisci(v); };
  foglio.appendRow([new Date(), testo(d.codice), "Da confermare", testo(d.tipo), testo(d.data), testo(d.fascia),
    testo(d.ora_extra), testo(d.bambini), testo(d.adulti), testo(d.festeggiato), testo(d.anni),
    testo(d.nome), testo(d.telefono), testo(d.email), testo(d.note)]);
}

function riepilogo(d) {
  var righe = [
    ["Codice richiesta", d.codice], ["Occasione", d.tipo], ["Data", d.data], ["Fascia", d.fascia],
    ["Ora extra", d.ora_extra], ["Invitati", d.bambini + " bambini, " + d.adulti + " adulti"]
  ];
  if (d.festeggiato) righe.push(["Festeggiato", d.festeggiato + (d.anni ? " (" + d.anni + " anni)" : "")]);
  righe.push(["Genitore", d.nome], ["Telefono", d.telefono], ["Email", d.email]);
  if (d.note) righe.push(["Note", d.note]);
  return righe;
}

function esc(v) {
  return pulisci(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function tabellaHtml(righe) {
  return '<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:15px">' +
    righe.map(function (r) {
      return '<tr><td style="padding:6px 14px 6px 0;color:#6E5F6A">' + esc(r[0]) +
             '</td><td style="padding:6px 0;font-weight:bold;color:#2A1F2D">' + esc(r[1]) + '</td></tr>';
    }).join("") + "</table>";
}

function testoSemplice(righe) {
  return righe.map(function (r) { return r[0] + ": " + pulisci(r[1]); }).join("\n");
}

function avvisaStaff(d) {
  var dest = EMAIL_STAFF || Session.getEffectiveUser().getEmail();
  var righe = riepilogo(d);
  MailApp.sendEmail({
    to: dest,
    replyTo: d.email,
    subject: "Nuova richiesta " + pulisci(d.codice) + " – " + pulisci(d.data) + ", " + pulisci(d.fascia),
    body: "Nuova richiesta di prenotazione dal sito.\n\n" + testoSemplice(righe) +
          "\n\nRispondi a questa email per scrivere direttamente al cliente.",
    htmlBody: '<div style="font-family:Arial,sans-serif"><h2 style="color:#E2312B">Nuova richiesta di prenotazione</h2>' +
              tabellaHtml(righe) + '<p style="color:#6E5F6A">Rispondi a questa email per scrivere direttamente al cliente.</p></div>'
  });
}

function confermaCliente(d) {
  var righe = riepilogo(d);
  var opzioni = {
    to: d.email,
    name: NOME_SALA,
    subject: NOME_SALA + " – abbiamo ricevuto la tua richiesta " + pulisci(d.codice),
    body: "Ciao " + pulisci(d.nome) + ",\n\nabbiamo ricevuto la tua richiesta per la sala. " +
          "Controlliamo la disponibilità e ti scriviamo al più presto per confermare.\n" +
          "La prenotazione è valida solo dopo la nostra conferma.\n\n" + testoSemplice(righe) +
          "\n\nVolantino: " + URL_VOLANTINO + "\n\nA presto!\n" + NOME_SALA,
    htmlBody: '<div style="font-family:Arial,sans-serif;max-width:560px">' +
      '<div style="background:#E2312B;color:#fff;padding:22px 24px;border-radius:14px 14px 0 0">' +
      '<h1 style="margin:0;font-size:26px">' + esc(NOME_SALA) + '</h1><p style="margin:4px 0 0">La tua festa, il nostro tendone</p></div>' +
      '<div style="background:#FFFDF8;padding:22px 24px;border:1px solid #F0E3CC;border-top:0;border-radius:0 0 14px 14px">' +
      '<p>Ciao ' + esc(d.nome) + ',</p>' +
      '<p>abbiamo ricevuto la tua richiesta per la sala. Controlliamo la disponibilità e ti scriviamo al più presto per confermare.</p>' +
      '<p style="background:#FFF1D6;padding:10px 14px;border-radius:10px"><b>La prenotazione è valida solo dopo la nostra conferma.</b></p>' +
      tabellaHtml(righe) +
      '<p>In allegato trovi il volantino della sala.</p><p>A presto!<br>' + esc(NOME_SALA) + '</p></div></div>'
  };
  if (EMAIL_STAFF) opzioni.replyTo = EMAIL_STAFF;
  try {
    var pdf = UrlFetchApp.fetch(URL_VOLANTINO, { muteHttpExceptions: true });
    if (pdf.getResponseCode() === 200) {
      opzioni.attachments = [pdf.getBlob().setName("Volantino-" + NOME_SALA + ".pdf")];
    }
  } catch (err) { /* se il volantino non è raggiungibile l'email parte lo stesso, con il link */ }
  MailApp.sendEmail(opzioni);
}
