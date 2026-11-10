/**
 * Circolandia – gestione richieste di prenotazione
 * Script Google Apps Script da collegare a un Foglio Google.
 *
 * Per ogni richiesta inviata dal sito:
 *   1. aggiunge una riga al foglio "Prenotazioni"
 *   2. manda un'email di avviso allo staff
 *   3. manda al cliente un'email di riepilogo con il volantino PDF allegato
 *
 * Calendario: il sito blocca da solo le fasce
 *   - delle righe di "Prenotazioni" con Stato = "Confermata"
 *   - elencate nel foglio "Blocchi" (chiusure, ferie, feste prese per telefono)
 * Dopo aver incollato lo script, esegui una volta la funzione preparaFoglio.
 *
 * Istruzioni di installazione: vedi il messaggio di Claude oppure il README del repository.
 */

// ====== DA COMPILARE ======
var EMAIL_STAFF = "info@circolandia.it";     // dove arrivano gli avvisi (inoltrato alla tua Gmail)
var EMAIL_MITTENTE = "info@circolandia.it";  // da quale indirizzo parte il riepilogo al cliente:
                                             // deve essere configurato in Gmail come "Invia messaggio come"
var NOME_SALA = "Circolandia";
var URL_VOLANTINO = "https://circolandia.it/volantino-circolandia.pdf";
// ==========================

var COLONNE = ["Ricevuta il", "Codice", "Stato", "Occasione", "Data festa", "Fascia", "Ora extra",
               "Bambini", "Adulti", "Festeggiato", "Anni", "Genitore", "Telefono", "Email", "Note", "Giorno"];

var STATI = ["Da confermare", "Confermata", "Rifiutata", "Annullata"];
var FASCE = ["mattina", "pomeriggio", "sera"];
var FASCE_BLOCCHI = ["Tutto il giorno", "Mattina", "Pomeriggio", "Sera"];
// Un'ora extra confermata dopo il pomeriggio (fino alle 19:00) si sovrappone alla sera
var ORA_EXTRA_OCCUPA = { pomeriggio: "sera" };
var MESI = ["gennaio","febbraio","marzo","aprile","maggio","giugno","luglio","agosto","settembre","ottobre","novembre","dicembre"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse(e.postData.contents);
    if (!d.nome || !d.email || !d.data || !d.fascia) throw new Error("dati mancanti");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(d.email)) throw new Error("email non valida");

    lock.waitLock(15000);
    var occ = calcolaOccupate();
    if ((occ[d.data_iso] || []).indexOf(chiaveFascia(d.fascia)) !== -1) {
      return risposta({ ok: false, occupato: true, occupate: occ });
    }
    registra(d);
    lock.releaseLock();

    avvisaStaff(d);
    confermaCliente(d);
    return risposta({ ok: true });
  } catch (err) {
    return risposta({ ok: false, errore: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doGet() {
  return risposta({ ok: true, messaggio: "Servizio prenotazioni " + NOME_SALA + " attivo", occupate: occupateInCache() });
}

// Si attiva da solo a ogni modifica del Foglio: il sito vede subito le nuove conferme.
function onEdit() {
  CacheService.getScriptCache().remove("occupate");
}

function risposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function pulisci(v) { return String(v == null ? "" : v).slice(0, 500); }

// ---------- Foglio Prenotazioni ----------

function foglioPrenotazioni() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var foglio = ss.getSheetByName("Prenotazioni") || ss.insertSheet("Prenotazioni");
  if (foglio.getLastRow() === 0) {
    foglio.appendRow(COLONNE);
    foglio.getRange(1, 1, 1, COLONNE.length).setFontWeight("bold");
    foglio.setFrozenRows(1);
  } else {
    // aggiunge in fondo le colonne nuove (es. "Giorno") ai fogli creati con versioni precedenti
    var intest = foglio.getRange(1, 1, 1, foglio.getLastColumn()).getValues()[0];
    COLONNE.forEach(function (c) {
      if (intest.indexOf(c) === -1) {
        foglio.getRange(1, foglio.getLastColumn() + 1).setValue(c).setFontWeight("bold");
      }
    });
  }
  return foglio;
}

function indiceColonne(foglio) {
  var intest = foglio.getRange(1, 1, 1, foglio.getLastColumn()).getValues()[0];
  var idx = {};
  intest.forEach(function (c, i) { idx[String(c).trim()] = i; });
  return idx;
}

function registra(d) {
  var foglio = foglioPrenotazioni();
  var idx = indiceColonne(foglio);
  // l'apostrofo iniziale evita che Fogli interpreti il testo come formula
  var testo = function (v) { return "'" + pulisci(v); };
  var valori = {
    "Ricevuta il": new Date(), "Codice": testo(d.codice), "Stato": "Da confermare", "Occasione": testo(d.tipo),
    "Data festa": testo(d.data), "Fascia": testo(d.fascia), "Ora extra": testo(d.ora_extra),
    "Bambini": testo(d.bambini), "Adulti": testo(d.adulti), "Festeggiato": testo(d.festeggiato), "Anni": testo(d.anni),
    "Genitore": testo(d.nome), "Telefono": testo(d.telefono), "Email": testo(d.email), "Note": testo(d.note),
    "Giorno": /^\d{4}-\d{2}-\d{2}$/.test(d.data_iso || "") ? "'" + d.data_iso : ""
  };
  var riga = new Array(foglio.getLastColumn()).fill("");
  Object.keys(valori).forEach(function (k) { if (idx[k] !== undefined) riga[idx[k]] = valori[k]; });
  foglio.appendRow(riga);
  CacheService.getScriptCache().remove("occupate");
}

// ---------- Calendario: fasce occupate ----------

function chiaveFascia(testo) {
  var t = String(testo || "").toLowerCase();
  if (t.indexOf("tutto") !== -1) return "tutto";
  for (var i = 0; i < FASCE.length; i++) if (t.indexOf(FASCE[i]) !== -1) return FASCE[i];
  return "";
}

function isoDaValore(v, fuso) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, fuso, "yyyy-MM-dd");
  var t = String(v || "").replace(/^'/, "").trim().toLowerCase();
  var m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return t;
  m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);                       // 21/11/2026
  if (m) return m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  m = t.match(/(\d{1,2})\s+([a-z]+)\s+(\d{4})/);                         // sabato 21 novembre 2026
  if (m && MESI.indexOf(m[2]) !== -1) return m[3] + "-" + ("0" + (MESI.indexOf(m[2]) + 1)).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  return "";
}

function aggiungi(occ, giorno, fascia) {
  if (!giorno || !fascia) return;
  var lista = occ[giorno] || (occ[giorno] = []);
  (fascia === "tutto" ? FASCE : [fascia]).forEach(function (f) { if (lista.indexOf(f) === -1) lista.push(f); });
}

function calcolaOccupate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fuso = ss.getSpreadsheetTimeZone();
  var oggi = Utilities.formatDate(new Date(), fuso, "yyyy-MM-dd");
  var occ = {};

  var fp = ss.getSheetByName("Prenotazioni");
  if (fp && fp.getLastRow() > 1) {
    var idx = indiceColonne(fp);
    var righe = fp.getRange(2, 1, fp.getLastRow() - 1, fp.getLastColumn()).getValues();
    righe.forEach(function (r) {
      if (String(r[idx["Stato"]]).trim().toLowerCase() !== "confermata") return;
      var giorno = (idx["Giorno"] !== undefined && isoDaValore(r[idx["Giorno"]], fuso)) || isoDaValore(r[idx["Data festa"]], fuso);
      var fascia = chiaveFascia(r[idx["Fascia"]]);
      aggiungi(occ, giorno, fascia);
      if (/^s/i.test(String(r[idx["Ora extra"]]).replace(/^'/, "")) && ORA_EXTRA_OCCUPA[fascia]) aggiungi(occ, giorno, ORA_EXTRA_OCCUPA[fascia]);
    });
  }

  var fb = ss.getSheetByName("Blocchi");
  if (fb && fb.getLastRow() > 1) {
    fb.getRange(2, 1, fb.getLastRow() - 1, 2).getValues().forEach(function (r) {
      aggiungi(occ, isoDaValore(r[0], fuso), chiaveFascia(r[1] || "tutto") || "tutto");
    });
  }

  Object.keys(occ).forEach(function (g) { if (g < oggi) delete occ[g]; });
  return occ;
}

function occupateInCache() {
  var cache = CacheService.getScriptCache();
  var c = cache.get("occupate");
  if (c) return JSON.parse(c);
  var occ = calcolaOccupate();
  cache.put("occupate", JSON.stringify(occ), 300);
  return occ;
}

// ---------- Da eseguire una volta: menu a tendina e colori ----------

function preparaFoglio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fp = foglioPrenotazioni();
  var idx = indiceColonne(fp);
  var colStato = idx["Stato"] + 1;
  var area = fp.getRange(2, colStato, fp.getMaxRows() - 1, 1);
  area.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATI, true).setAllowInvalid(false).build());
  var colori = { "Da confermare": "#FFF1D6", "Confermata": "#D6F2E3", "Rifiutata": "#F1E6EA", "Annullata": "#F1E6EA" };
  var regole = fp.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (g) { return g.getColumn() === colStato; });
  });
  Object.keys(colori).forEach(function (st) {
    regole.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(st).setBackground(colori[st]).setRanges([area]).build());
  });
  fp.setConditionalFormatRules(regole);

  var fb = ss.getSheetByName("Blocchi") || ss.insertSheet("Blocchi");
  if (fb.getLastRow() === 0) {
    fb.appendRow(["Giorno", "Fascia", "Motivo"]);
    fb.getRange(1, 1, 1, 3).setFontWeight("bold");
    fb.setFrozenRows(1);
    fb.setColumnWidth(3, 280);
  }
  fb.getRange(2, 1, fb.getMaxRows() - 1, 1).setNumberFormat("dd/MM/yyyy")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  fb.getRange(2, 2, fb.getMaxRows() - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(FASCE_BLOCCHI, true).setAllowInvalid(false).build());

  CacheService.getScriptCache().remove("occupate");
  Logger.log("Foglio pronto. Fasce occupate ora: " + JSON.stringify(calcolaOccupate()));
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
  inviaDaAlias(opzioni);
}

// Invia dall'indirizzo @circolandia.it se è configurato in Gmail come "Invia messaggio come",
// altrimenti dall'indirizzo Gmail dell'account (con nome "Circolandia" e risposta a EMAIL_STAFF).
function inviaDaAlias(o) {
  var alias = [];
  try { alias = GmailApp.getAliases(); } catch (err) {}
  if (EMAIL_MITTENTE && alias.indexOf(EMAIL_MITTENTE) !== -1) {
    GmailApp.sendEmail(o.to, o.subject, o.body, {
      from: EMAIL_MITTENTE, name: o.name, replyTo: o.replyTo,
      htmlBody: o.htmlBody, attachments: o.attachments
    });
  } else {
    MailApp.sendEmail(o);
  }
}

// Da eseguire una volta dall'editor (menu Esegui) per controllare la configurazione.
function verificaConfigurazione() {
  var alias = GmailApp.getAliases();
  Logger.log("Alias disponibili in Gmail: " + (alias.length ? alias.join(", ") : "nessuno"));
  Logger.log(alias.indexOf(EMAIL_MITTENTE) !== -1
    ? "OK: i riepiloghi ai clienti partiranno da " + EMAIL_MITTENTE
    : "Attenzione: " + EMAIL_MITTENTE + " non è tra gli alias di Gmail, i riepiloghi partiranno dall'indirizzo Gmail");
}
