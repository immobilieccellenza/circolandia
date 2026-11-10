/**
 * Circolandia – gestione richieste di prenotazione
 * Script Google Apps Script collegato al Foglio Google delle prenotazioni.
 *
 * Per ogni richiesta inviata dal sito:
 *   1. aggiunge una riga al foglio "Prenotazioni"
 *   2. manda un'email di avviso allo staff
 *   3. manda al cliente un'email di riepilogo con il volantino PDF allegato
 *
 * Quando nel Foglio lo Stato diventa "Confermata":
 *   - controlla che la fascia non sia già occupata (altrimenti annulla la modifica e avvisa)
 *   - manda al cliente l'email di conferma (una sola volta)
 *   - il sito blocca quella fascia nel calendario
 *
 * Dopo ogni aggiornamento dello script:
 *   1. esegui una volta la funzione preparaFoglio (menu, controlli, istruzioni, automazioni)
 *   2. Esegui il deployment > Gestisci deployment > matita > Nuova versione
 */

// ====== DA COMPILARE ======
var EMAIL_STAFF = "info@circolandia.it";     // dove arrivano gli avvisi (inoltrato alla tua Gmail)
var EMAIL_MITTENTE = "info@circolandia.it";  // da quale indirizzo partono le email ai clienti:
                                             // deve essere configurato in Gmail come "Invia messaggio come"
var NOME_SALA = "Circolandia";
var URL_SITO = "https://circolandia.it";
var URL_VOLANTINO = "https://circolandia.it/volantino-circolandia.pdf";
var INVIA_EMAIL_CONFERMA = true;             // false = la conferma al cliente la scrivi tu
// ==========================

var COLONNE = ["Ricevuta il", "Codice", "Stato", "Occasione", "Data festa", "Fascia", "Ora extra",
               "Bambini", "Adulti", "Festeggiato", "Anni", "Genitore", "Telefono", "Email", "Note", "Giorno",
               "Prezzo (€)", "Messaggio al cliente", "Conferma inviata"];

var STATI = ["Da confermare", "Confermata", "Rifiutata", "Annullata"];
var OCCASIONI = ["Compleanno", "Festa di classe", "Altra occasione"];
var FASCE_TESTO = ["Mattina 9:30–12:30", "Pomeriggio 15:00–18:00", "Sera 18:30–21:30"];
var FASCE = ["mattina", "pomeriggio", "sera"];
var FASCE_BLOCCHI = ["Tutto il giorno", "Mattina", "Pomeriggio", "Sera"];
// Un'ora extra dopo il pomeriggio (fino alle 19:00) si sovrappone alla sera
var ORA_EXTRA_OCCUPA = { pomeriggio: "sera" };
var MESI = ["gennaio","febbraio","marzo","aprile","maggio","giugno","luglio","agosto","settembre","ottobre","novembre","dicembre"];
var GIORNI = ["domenica","lunedì","martedì","mercoledì","giovedì","venerdì","sabato"];

var NOTE_COLONNE = {
  "Ricevuta il": "Data e ora in cui è arrivata la richiesta dal sito. La compila lo script.",
  "Codice": "Codice della richiesta (CRL-…), lo stesso che il cliente ha nell'email di riepilogo.",
  "Stato": "Da confermare → Confermata (blocca la data sul sito e manda l'email di conferma al cliente).\nRifiutata o Annullata = la fascia resta libera.",
  "Occasione": "Compleanno, Festa di classe o Altra occasione.",
  "Data festa": "Data scritta per esteso, come l'ha scelta il cliente.",
  "Fascia": "Fascia oraria prenotata. Usa il menu a tendina.",
  "Ora extra": "Sì se il cliente vuole un'ora in più. Un'ora extra confermata dopo il pomeriggio blocca anche la sera.",
  "Bambini": "Numero di bambini invitati.",
  "Adulti": "Numero di adulti invitati.",
  "Anni": "Anni che compie il festeggiato.",
  "Email": "Email del cliente: qui arriva la conferma automatica.",
  "Giorno": "Data della festa in formato data. È quella che usa il calendario del sito: non cancellarla.",
  "Prezzo (€)": "Facoltativo. Se lo scrivi PRIMA di mettere Confermata, compare nell'email di conferma.",
  "Messaggio al cliente": "Facoltativo. Testo aggiunto all'email di conferma (es. indicazioni per l'arrivo, caparra).",
  "Conferma inviata": "Data e ora dell'email di conferma. La compila lo script. Per rimandare la conferma cancella questa cella e rimetti Confermata."
};

// =====================================================================
// Sito: ricezione richieste e calendario
// =====================================================================

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
    confermaRicezione(d);
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

function risposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Trigger semplice: a ogni modifica il sito rilegge subito le date occupate.
function onEdit() {
  svuotaCache();
}

function svuotaCache() {
  CacheService.getScriptCache().remove("occupate");
}

// =====================================================================
// Foglio Prenotazioni
// =====================================================================

function pulisci(v) { return String(v == null ? "" : v).slice(0, 500); }

function lettera(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function foglioPrenotazioni() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var foglio = ss.getSheetByName("Prenotazioni") || ss.insertSheet("Prenotazioni");
  if (foglio.getLastRow() === 0) {
    foglio.getRange(1, 1, 1, COLONNE.length).setValues([COLONNE]);
  } else {
    // aggiunge in fondo le colonne nuove ai fogli creati con versioni precedenti
    var intest = foglio.getRange(1, 1, 1, foglio.getLastColumn()).getValues()[0];
    COLONNE.forEach(function (c) {
      if (intest.indexOf(c) === -1) {
        if (foglio.getMaxColumns() <= foglio.getLastColumn()) foglio.insertColumnAfter(foglio.getLastColumn());
        foglio.getRange(1, foglio.getLastColumn() + 1).setValue(c);
        intest.push(c);
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

function formulaData(iso) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  return m ? "=DATE(" + (+m[1]) + "," + (+m[2]) + "," + (+m[3]) + ")" : "";
}

function numero(v) {
  var n = parseInt(v, 10);
  return isNaN(n) ? "" : n;
}

function registra(d) {
  var foglio = foglioPrenotazioni();
  var idx = indiceColonne(foglio);
  // l'apostrofo iniziale evita che Fogli interpreti il testo come formula
  var testo = function (v) { return "'" + pulisci(v); };
  var valori = {
    "Ricevuta il": new Date(), "Codice": testo(d.codice), "Stato": "Da confermare", "Occasione": pulisci(d.tipo),
    "Data festa": testo(d.data), "Fascia": pulisci(d.fascia), "Ora extra": d.ora_extra === "Sì" ? "Sì" : "No",
    "Bambini": numero(d.bambini), "Adulti": numero(d.adulti), "Festeggiato": testo(d.festeggiato), "Anni": numero(d.anni),
    "Genitore": testo(d.nome), "Telefono": testo(d.telefono), "Email": pulisci(d.email), "Note": testo(d.note),
    "Giorno": formulaData(d.data_iso)
  };
  // i valori che non sono formule non devono mai iniziare con "=": li forziamo a testo
  ["Occasione", "Fascia", "Email"].forEach(function (k) { if (/^[=+\-@]/.test(valori[k])) valori[k] = "'" + valori[k]; });
  var riga = new Array(foglio.getLastColumn()).fill("");
  Object.keys(valori).forEach(function (k) { if (idx[k] !== undefined) riga[idx[k]] = valori[k]; });
  foglio.appendRow(riga);
  svuotaCache();
}

// =====================================================================
// Calendario: fasce occupate
// =====================================================================

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
  m = t.match(/(\d{1,2})\s+([a-zì]+)\s+(\d{4})/);                       // sabato 21 novembre 2026
  if (m && MESI.indexOf(m[2]) !== -1) return m[3] + "-" + ("0" + (MESI.indexOf(m[2]) + 1)).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  return "";
}

function giornoRiga(r, idx, fuso) {
  return (idx["Giorno"] !== undefined && isoDaValore(r[idx["Giorno"]], fuso)) || isoDaValore(r[idx["Data festa"]], fuso);
}

// fasce occupate da una riga (la fascia più l'eventuale sovrapposizione dell'ora extra)
function fasceRiga(r, idx) {
  var f = chiaveFascia(r[idx["Fascia"]]);
  if (!f) return [];
  if (f === "tutto") return FASCE.slice();
  var out = [f];
  if (/^s/i.test(String(r[idx["Ora extra"]]).replace(/^'/, "")) && ORA_EXTRA_OCCUPA[f]) out.push(ORA_EXTRA_OCCUPA[f]);
  return out;
}

function aggiungi(occ, giorno, fasce) {
  if (!giorno) return;
  var lista = occ[giorno] || (occ[giorno] = []);
  fasce.forEach(function (f) { if (lista.indexOf(f) === -1) lista.push(f); });
}

// escludiRiga: numero di riga del foglio Prenotazioni da non contare (serve al controllo doppie conferme)
function calcolaOccupate(escludiRiga) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fuso = ss.getSpreadsheetTimeZone();
  var oggi = Utilities.formatDate(new Date(), fuso, "yyyy-MM-dd");
  var occ = {};

  var fp = ss.getSheetByName("Prenotazioni");
  if (fp && fp.getLastRow() > 1) {
    var idx = indiceColonne(fp);
    fp.getRange(2, 1, fp.getLastRow() - 1, fp.getLastColumn()).getValues().forEach(function (r, i) {
      if (i + 2 === escludiRiga) return;
      if (String(r[idx["Stato"]]).trim().toLowerCase() !== "confermata") return;
      aggiungi(occ, giornoRiga(r, idx, fuso), fasceRiga(r, idx));
    });
  }

  var fb = ss.getSheetByName("Blocchi");
  if (fb && fb.getLastRow() > 1) {
    fb.getRange(2, 1, fb.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (!r[0]) return;
      var f = chiaveFascia(r[1] || "tutto") || "tutto";
      aggiungi(occ, isoDaValore(r[0], fuso), f === "tutto" ? FASCE : [f]);
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

// =====================================================================
// Conferma: trigger installato da preparaFoglio
// =====================================================================

function alCambioStato(e) {
  if (!e || !e.range) return;
  var foglio = e.range.getSheet();
  if (foglio.getName() !== "Prenotazioni" && foglio.getName() !== "Blocchi") return;
  svuotaCache();
  if (foglio.getName() !== "Prenotazioni" || e.range.getRow() < 2) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  var idx = indiceColonne(foglio);
  if (e.range.getColumn() !== idx["Stato"] + 1) return;
  if (String(e.value || "") !== "Confermata") return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fuso = ss.getSpreadsheetTimeZone();
  var numRiga = e.range.getRow();
  var r = foglio.getRange(numRiga, 1, 1, foglio.getLastColumn()).getValues()[0];
  var giorno = giornoRiga(r, idx, fuso), fasce = fasceRiga(r, idx);

  if (!giorno || !fasce.length) {
    e.range.setValue(e.oldValue || "Da confermare");
    ss.toast("Per confermare servono Giorno e Fascia compilati in questa riga.", "Conferma non salvata", 10);
    return;
  }

  var occAltri = calcolaOccupate(numRiga);
  var presi = fasce.filter(function (f) { return (occAltri[giorno] || []).indexOf(f) !== -1; });
  if (presi.length) {
    e.range.setValue(e.oldValue || "Da confermare");
    svuotaCache();
    ss.toast("Il " + dataLeggibile(giorno) + " la fascia " + presi.join(" e ") +
      " è già confermata per un'altra festa o bloccata nel foglio Blocchi.", "Fascia già occupata", 15);
    return;
  }

  if (!INVIA_EMAIL_CONFERMA) { ss.toast("Prenotazione confermata. La data ora risulta occupata sul sito.", NOME_SALA, 6); return; }

  var colInviata = idx["Conferma inviata"];
  if (colInviata !== undefined && r[colInviata]) {
    ss.toast("Email di conferma già inviata il " + Utilities.formatDate(new Date(r[colInviata]), fuso, "dd/MM/yyyy HH:mm") +
      ". Per rimandarla cancella la cella Conferma inviata e rimetti Confermata.", NOME_SALA, 10);
    return;
  }
  var email = String(r[idx["Email"]] || "").replace(/^'/, "").trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
    ss.toast("Prenotazione confermata, ma manca un'email valida: avvisa il cliente a mano.", NOME_SALA, 10);
    return;
  }
  try {
    emailConferma(r, idx, giorno, email);
    if (colInviata !== undefined) foglio.getRange(numRiga, colInviata + 1).setValue(new Date());
    ss.toast("Email di conferma inviata a " + email + ".", NOME_SALA, 6);
  } catch (err) {
    ss.toast("Prenotazione confermata, ma l'email non è partita: " + err.message, "Attenzione", 15);
  }
}

// =====================================================================
// Email
// =====================================================================

function dataLeggibile(iso) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  if (!m) return iso;
  var d = new Date(+m[1], +m[2] - 1, +m[3]);
  return GIORNI[d.getDay()] + " " + (+m[3]) + " " + MESI[+m[2] - 1] + " " + m[1];
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
      return '<tr><td style="padding:6px 14px 6px 0;color:#6E5F6A;vertical-align:top">' + esc(r[0]) +
             '</td><td style="padding:6px 0;font-weight:bold;color:#2A1F2D">' + esc(r[1]).replace(/\n/g, "<br>") + '</td></tr>';
    }).join("") + "</table>";
}

function testoSemplice(righe) {
  return righe.map(function (r) { return r[0] + ": " + pulisci(r[1]); }).join("\n");
}

function cornice(titolo, corpoHtml) {
  return '<div style="font-family:Arial,sans-serif;max-width:560px">' +
    '<div style="background:#E2312B;color:#fff;padding:22px 24px;border-radius:14px 14px 0 0">' +
    '<h1 style="margin:0;font-size:26px">' + esc(NOME_SALA) + '</h1><p style="margin:4px 0 0">' + esc(titolo) + '</p></div>' +
    '<div style="background:#FFFDF8;padding:22px 24px;border:1px solid #F0E3CC;border-top:0;border-radius:0 0 14px 14px">' +
    corpoHtml + '<p style="color:#6E5F6A;font-size:13px;margin-top:22px">' + esc(URL_SITO.replace(/^https?:\/\//, "")) + '</p></div></div>';
}

function avvisaStaff(d) {
  var dest = EMAIL_STAFF || Session.getEffectiveUser().getEmail();
  var righe = riepilogo(d);
  MailApp.sendEmail({
    to: dest,
    replyTo: d.email,
    subject: "Nuova richiesta " + pulisci(d.codice) + " – " + pulisci(d.data) + ", " + pulisci(d.fascia),
    body: "Nuova richiesta di prenotazione dal sito.\n\n" + testoSemplice(righe) +
          "\n\nPer confermarla: nel Foglio metti Stato = Confermata.\nRispondi a questa email per scrivere direttamente al cliente.",
    htmlBody: '<div style="font-family:Arial,sans-serif"><h2 style="color:#E2312B">Nuova richiesta di prenotazione</h2>' +
              tabellaHtml(righe) + '<p style="color:#6E5F6A">Per confermarla: nel Foglio metti <b>Stato = Confermata</b>.<br>' +
              'Rispondi a questa email per scrivere direttamente al cliente.</p></div>'
  });
}

function confermaRicezione(d) {
  var righe = riepilogo(d);
  var opzioni = {
    to: d.email,
    name: NOME_SALA,
    subject: NOME_SALA + " – abbiamo ricevuto la tua richiesta " + pulisci(d.codice),
    body: "Ciao " + pulisci(d.nome) + ",\n\nabbiamo ricevuto la tua richiesta per la sala. " +
          "Controlliamo la disponibilità e ti scriviamo al più presto per confermare.\n" +
          "La prenotazione è valida solo dopo la nostra conferma.\n\n" + testoSemplice(righe) +
          "\n\nVolantino: " + URL_VOLANTINO + "\n\nA presto!\n" + NOME_SALA,
    htmlBody: cornice("La tua festa, il nostro tendone",
      '<p>Ciao ' + esc(d.nome) + ',</p>' +
      '<p>abbiamo ricevuto la tua richiesta per la sala. Controlliamo la disponibilità e ti scriviamo al più presto per confermare.</p>' +
      '<p style="background:#FFF1D6;padding:10px 14px;border-radius:10px"><b>La prenotazione è valida solo dopo la nostra conferma.</b></p>' +
      tabellaHtml(righe) +
      '<p>In allegato trovi il volantino della sala.</p><p>A presto!<br>' + esc(NOME_SALA) + '</p>')
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

function emailConferma(r, idx, giorno, email) {
  var v = function (k) { return idx[k] === undefined ? "" : String(r[idx[k]] == null ? "" : r[idx[k]]).replace(/^'/, "").trim(); };
  var nome = v("Genitore") || "";
  var fascia = v("Fascia") + (/^s/i.test(v("Ora extra")) ? " + 1 ora extra" : "");
  var righe = [];
  if (v("Codice")) righe.push(["Codice", v("Codice")]);
  righe.push(["Data", dataLeggibile(giorno)], ["Fascia", fascia]);
  if (v("Occasione")) righe.push(["Occasione", v("Occasione")]);
  if (v("Bambini") || v("Adulti")) righe.push(["Invitati", (v("Bambini") || "0") + " bambini, " + (v("Adulti") || "0") + " adulti"]);
  if (v("Festeggiato")) righe.push(["Festeggiato", v("Festeggiato") + (v("Anni") ? " (" + v("Anni") + " anni)" : "")]);
  var prezzo = idx["Prezzo (€)"] !== undefined ? r[idx["Prezzo (€)"]] : "";
  if (prezzo !== "" && prezzo !== null && !isNaN(Number(prezzo))) {
    righe.push(["Prezzo", Number(prezzo).toFixed(2).replace(".", ",").replace(",00", "") + " €"]);
  }
  var messaggio = v("Messaggio al cliente");

  inviaDaAlias({
    to: email,
    name: NOME_SALA,
    replyTo: EMAIL_STAFF || undefined,
    subject: NOME_SALA + " – la tua festa è confermata! " + dataLeggibile(giorno),
    body: "Ciao " + nome + ",\n\nla tua prenotazione è confermata: la sala è riservata per te.\n\n" +
          testoSemplice(righe) + (messaggio ? "\n\n" + messaggio : "") +
          "\n\nPer qualsiasi domanda rispondi a questa email.\n\nA presto sotto il tendone!\n" + NOME_SALA,
    htmlBody: cornice("Prenotazione confermata",
      '<p>Ciao ' + esc(nome) + ',</p>' +
      '<p style="background:#D6F2E3;padding:12px 14px;border-radius:10px;font-size:17px"><b>La tua festa è confermata!</b> La sala è riservata per te.</p>' +
      tabellaHtml(righe) +
      (messaggio ? '<p style="background:#FFF1D6;padding:10px 14px;border-radius:10px">' + esc(messaggio).replace(/\n/g, "<br>") + '</p>' : '') +
      '<p>Per qualsiasi domanda rispondi a questa email.</p><p>A presto sotto il tendone!<br>' + esc(NOME_SALA) + '</p>')
  });
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

// =====================================================================
// preparaFoglio: da eseguire una volta dopo ogni aggiornamento dello script
// =====================================================================

function preparaFoglio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fp = foglioPrenotazioni();
  migraDati(fp);
  formattaPrenotazioni(fp);
  preparaBlocchi(ss);
  preparaProssimeFeste(ss, fp);
  preparaIstruzioni(ss);
  installaAutomazioni(ss);
  svuotaCache();
  ss.setActiveSheet(ss.getSheetByName("Istruzioni"));
  ss.toast("Foglio pronto: istruzioni, controlli e conferma automatica attivi.", NOME_SALA, 8);
  Logger.log("Foglio pronto. Fasce occupate ora: " + JSON.stringify(calcolaOccupate()));
}

// converte i dati scritti dalle versioni precedenti (date e numeri salvati come testo)
function migraDati(fp) {
  if (fp.getLastRow() < 2) return;
  var idx = indiceColonne(fp), fuso = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var n = fp.getLastRow() - 1;
  var dati = fp.getRange(2, 1, n, fp.getLastColumn()).getValues();
  var colG = idx["Giorno"] + 1;
  var giorni = fp.getRange(2, colG, n, 1).getFormulas();
  dati.forEach(function (r, i) {
    var g = r[idx["Giorno"]];
    if (!(g instanceof Date) && !giorni[i][0]) {
      var iso = isoDaValore(g, fuso) || isoDaValore(r[idx["Data festa"]], fuso);
      if (iso) fp.getRange(i + 2, colG).setFormula(formulaData(iso));
    }
    ["Bambini", "Adulti", "Anni"].forEach(function (k) {
      var v = r[idx[k]];
      if (typeof v === "string" && /^'?\d+$/.test(v.trim())) fp.getRange(i + 2, idx[k] + 1).setValue(parseInt(v.replace("'", ""), 10));
    });
    ["Fascia", "Occasione", "Email"].forEach(function (k) {
      var v = r[idx[k]];
      if (typeof v === "string" && v.charAt(0) === "'") fp.getRange(i + 2, idx[k] + 1).setValue(v.slice(1));
    });
    var ox = String(r[idx["Ora extra"]] || "").replace(/^'/, "");
    if (ox && ox !== r[idx["Ora extra"]]) fp.getRange(i + 2, idx["Ora extra"] + 1).setValue(/^s/i.test(ox) ? "Sì" : "No");
  });
}

function colonna(fp, idx, nome) {
  var L = lettera(idx[nome] + 1);
  return fp.getRange(L + "2:" + L);
}

function formattaPrenotazioni(fp) {
  var idx = indiceColonne(fp), nCol = fp.getLastColumn();
  var dv = SpreadsheetApp.newDataValidation;

  // intestazione
  var testa = fp.getRange(1, 1, 1, nCol);
  testa.setBackground("#E2312B").setFontColor("#FFFFFF").setFontWeight("bold").setWrap(true).setVerticalAlignment("middle");
  fp.setRowHeight(1, 42);
  fp.setFrozenRows(1);
  fp.setFrozenColumns(3);
  Object.keys(NOTE_COLONNE).forEach(function (k) {
    if (idx[k] !== undefined) fp.getRange(1, idx[k] + 1).setNote(NOTE_COLONNE[k]);
  });

  // larghezze
  var larg = { "Ricevuta il": 130, "Codice": 90, "Stato": 120, "Occasione": 130, "Data festa": 190, "Fascia": 170, "Ora extra": 80,
    "Bambini": 75, "Adulti": 70, "Festeggiato": 120, "Anni": 55, "Genitore": 150, "Telefono": 120, "Email": 200, "Note": 220,
    "Giorno": 120, "Prezzo (€)": 90, "Messaggio al cliente": 240, "Conferma inviata": 140 };
  Object.keys(larg).forEach(function (k) { if (idx[k] !== undefined) fp.setColumnWidth(idx[k] + 1, larg[k]); });

  // formati
  colonna(fp, idx, "Ricevuta il").setNumberFormat("dd/MM/yyyy HH:mm");
  colonna(fp, idx, "Giorno").setNumberFormat("ddd dd/MM/yyyy");
  colonna(fp, idx, "Prezzo (€)").setNumberFormat('#,##0.00 "€"');
  colonna(fp, idx, "Conferma inviata").setNumberFormat("dd/MM/yyyy HH:mm");
  colonna(fp, idx, "Note").setWrap(true);
  colonna(fp, idx, "Messaggio al cliente").setWrap(true);

  // convalida dati
  var lista = function (valori, aiuto) { return dv().requireValueInList(valori, true).setAllowInvalid(false).setHelpText(aiuto).build(); };
  colonna(fp, idx, "Stato").setDataValidation(lista(STATI, "Scegli uno stato dal menu."));
  colonna(fp, idx, "Occasione").setDataValidation(lista(OCCASIONI, "Scegli l'occasione dal menu."));
  colonna(fp, idx, "Fascia").setDataValidation(lista(FASCE_TESTO, "Scegli la fascia oraria dal menu."));
  colonna(fp, idx, "Ora extra").setDataValidation(lista(["Sì", "No"], "Sì oppure No."));
  colonna(fp, idx, "Bambini").setDataValidation(dv().requireNumberBetween(0, 200).setAllowInvalid(false).setHelpText("Un numero da 0 a 200.").build());
  colonna(fp, idx, "Adulti").setDataValidation(dv().requireNumberBetween(0, 200).setAllowInvalid(false).setHelpText("Un numero da 0 a 200.").build());
  colonna(fp, idx, "Anni").setDataValidation(dv().requireNumberBetween(0, 99).setAllowInvalid(false).setHelpText("Un numero da 0 a 99, oppure lascia vuoto.").build());
  colonna(fp, idx, "Email").setDataValidation(dv().requireTextIsEmail().setAllowInvalid(false).setHelpText("Un indirizzo email valido: qui arriva la conferma.").build());
  colonna(fp, idx, "Giorno").setDataValidation(dv().requireDate().setAllowInvalid(false).setHelpText("Una data (es. 21/11/2026). Serve al calendario del sito.").build());
  colonna(fp, idx, "Prezzo (€)").setDataValidation(dv().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText("Solo il numero, es. 150. Va nell'email di conferma.").build());

  // colori: conflitti, feste passate, stato
  var L = function (k) { return "$" + lettera(idx[k] + 1); };
  var tutta = fp.getRange(2, 1, fp.getMaxRows() - 1, nCol);
  var areaStato = colonna(fp, idx, "Stato");
  var regole = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(' + L("Stato") + '2="Confermata",ISNUMBER(' + L("Giorno") + '2),COUNTIFS(' +
        L("Stato") + ':' + L("Stato") + ',"Confermata",' +
        L("Giorno") + ':' + L("Giorno") + ',' + L("Giorno") + '2,' +
        L("Fascia") + ':' + L("Fascia") + ',' + L("Fascia") + '2)>1)')
      .setBackground("#F8C9C5").setFontColor("#A81E1A").setRanges([tutta]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(' + L("Giorno") + '2),' + L("Giorno") + '2<TODAY())')
      .setFontColor("#A99BA6").setRanges([tutta]).build()
  ];
  var colori = { "Da confermare": "#FFF1D6", "Confermata": "#D6F2E3", "Rifiutata": "#EEE6EA", "Annullata": "#EEE6EA" };
  Object.keys(colori).forEach(function (st) {
    regole.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(st).setBackground(colori[st]).setRanges([areaStato]).build());
  });
  fp.setConditionalFormatRules(regole);

  // filtro sulle intestazioni
  if (!fp.getFilter()) fp.getRange(1, 1, fp.getMaxRows(), nCol).createFilter();

  // protezione "morbida" sulle colonne compilate dallo script: chiede conferma prima di modificarle
  fp.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (String(p.getDescription()).indexOf("Compilata dallo script") === 0) p.remove();
  });
  ["Ricevuta il", "Codice", "Conferma inviata"].forEach(function (k) {
    if (idx[k] !== undefined) colonna(fp, idx, k).protect().setDescription("Compilata dallo script: " + k).setWarningOnly(true);
  });
}

function preparaBlocchi(ss) {
  var fb = ss.getSheetByName("Blocchi") || ss.insertSheet("Blocchi");
  if (fb.getLastRow() === 0) fb.getRange(1, 1, 1, 3).setValues([["Giorno", "Fascia", "Motivo"]]);
  fb.getRange(1, 1, 1, 3).setBackground("#1F5FBF").setFontColor("#FFFFFF").setFontWeight("bold");
  fb.setFrozenRows(1);
  fb.setColumnWidth(1, 130); fb.setColumnWidth(2, 140); fb.setColumnWidth(3, 320);
  fb.getRange(1, 1).setNote("Giorno da bloccare sul sito (es. 25/12/2026).");
  fb.getRange(1, 2).setNote("Tutto il giorno, oppure una sola fascia.");
  fb.getRange(1, 3).setNote("Facoltativo: ferie, festivo, festa presa per telefono…");
  fb.getRange("A2:A").setNumberFormat("ddd dd/MM/yyyy")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText("Una data, es. 25/12/2026.").build());
  fb.getRange("B2:B")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(FASCE_BLOCCHI, true).setAllowInvalid(false).setHelpText("Scegli dal menu.").build());
  fb.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=AND(ISNUMBER($A2),$A2<TODAY())")
      .setFontColor("#A99BA6").setRanges([fb.getRange("A2:C")]).build()
  ]);
}

function preparaProssimeFeste(ss, fp) {
  var idx = indiceColonne(fp);
  var c = function (k) { return idx[k] === undefined ? null : "Col" + (idx[k] + 1); };
  var campi = ["Giorno", "Fascia", "Ora extra", "Occasione", "Festeggiato", "Anni", "Bambini", "Adulti", "Genitore", "Telefono", "Prezzo (€)", "Codice"]
    .map(c).filter(Boolean).join(", ");
  var ultima = lettera(fp.getLastColumn());
  var oggi = '"date \'"&YEAR(TODAY())&"-"&TEXT(MONTH(TODAY());"00")&"-"&TEXT(DAY(TODAY());"00")&"\'"';
  var formula = '=IFERROR(QUERY({Prenotazioni!A1:' + ultima + '};"select ' + campi + ' where ' + c("Stato") + ' = \'Confermata\' and ' +
    c("Giorno") + ' >= "&' + oggi + '&" order by ' + c("Giorno") + '";1);"Nessuna festa confermata in arrivo")';

  var sh = ss.getSheetByName("Prossime feste") || ss.insertSheet("Prossime feste");
  sh.clear();
  sh.getRange("A1").setValue("Prossime feste confermate").setFontSize(16).setFontWeight("bold").setFontColor("#14306B");
  sh.getRange("A2").setValue("Si aggiorna da solo dal foglio Prenotazioni. Non scrivere in questo foglio.").setFontColor("#6E5F6A");
  sh.getRange("A4").setFormula(formula.replace(/;/g, ","));
  sh.getRange("A5:A").setNumberFormat("ddd dd/MM/yyyy");
  sh.getRange(4, 1, 1, 12).setBackground("#22A06B").setFontColor("#FFFFFF").setFontWeight("bold");
  sh.setFrozenRows(4);
  [130, 170, 80, 130, 120, 55, 75, 70, 150, 120, 90, 90].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  sh.protect().setDescription("Foglio calcolato in automatico").setWarningOnly(true);
}

function preparaIstruzioni(ss) {
  var sh = ss.getSheetByName("Istruzioni") || ss.insertSheet("Istruzioni");
  sh.clear();
  var righe = [
    ["Circolandia – come usare questo Foglio", "titolo"],
    ["Ogni richiesta fatta dal sito circolandia.it arriva qui, nel foglio Prenotazioni, e ti arriva anche un'email di avviso.", ""],
    ["", ""],
    ["1. Confermare una festa", "sezione"],
    ["• Nel foglio Prenotazioni trova la riga (codice CRL-… o nome del genitore).", ""],
    ["• Facoltativo: scrivi il Prezzo (€) e un Messaggio al cliente (es. caparra, indicazioni). Vanno scritti PRIMA di confermare.", ""],
    ["• Nella colonna Stato scegli Confermata dal menu.", ""],
    ["• Il cliente riceve subito l'email di conferma da info@circolandia.it e la fascia si blocca sul calendario del sito.", ""],
    ["• Se quella fascia è già confermata per un'altra festa (o bloccata), lo Stato torna indietro e compare un avviso in basso a destra.", ""],
    ["", ""],
    ["2. Rifiutare o annullare", "sezione"],
    ["• Rifiutata: la data non è disponibile. Annullata: una festa confermata che salta. In entrambi i casi la fascia torna libera sul sito.", ""],
    ["• Per questi casi non parte nessuna email automatica: avvisa tu il cliente (rispondi all'email di avviso o scrivigli su WhatsApp).", ""],
    ["", ""],
    ["3. Feste prese per telefono, ferie e chiusure", "sezione"],
    ["• Foglio Blocchi: scrivi il Giorno e scegli la Fascia (o Tutto il giorno). Il sito non la propone più ai clienti.", ""],
    ["• In alternativa puoi aggiungere una riga a mano in Prenotazioni con Giorno, Fascia, Email e poi Stato = Confermata: in quel caso parte anche l'email.", ""],
    ["", ""],
    ["4. Prossime feste", "sezione"],
    ["• Il foglio Prossime feste mostra in ordine di data tutte le feste confermate da oggi in poi. Si aggiorna da solo.", ""],
    ["", ""],
    ["Colori e controlli", "sezione"],
    ["• Stato giallo = da confermare, verde = confermata, grigio = rifiutata o annullata.", ""],
    ["• Riga rossa = due feste confermate nello stesso giorno e nella stessa fascia: controlla subito.", ""],
    ["• Testo grigio = festa già passata.", ""],
    ["• Le colonne hanno menu a tendina e controlli: se un valore non è valido Fogli lo rifiuta. Passa il mouse sulle intestazioni per vedere a cosa serve ogni colonna.", ""],
    ["• Ricevuta il, Codice e Conferma inviata le compila lo script: Fogli ti chiede conferma prima di modificarle.", ""],
    ["• Per rimandare un'email di conferma: cancella la cella Conferma inviata, poi rimetti Stato = Confermata.", ""],
    ["", ""],
    ["Da non fare", "sezione"],
    ["• Non rinominare i fogli Prenotazioni e Blocchi e non cambiare i nomi delle colonne: il sito e lo script li cercano per nome.", ""],
    ["• Non cancellare la colonna Giorno: è quella che usa il calendario del sito.", ""],
    ["• Puoi invece ordinare e filtrare le righe con il filtro sulle intestazioni.", ""]
  ];
  sh.getRange(1, 1, righe.length, 1).setValues(righe.map(function (r) { return [r[0]]; }));
  sh.setColumnWidth(1, 900);
  sh.getRange(1, 1, righe.length, 1).setWrap(true).setVerticalAlignment("middle").setFontSize(11).setFontColor("#2A1F2D");
  righe.forEach(function (r, i) {
    var cella = sh.getRange(i + 1, 1);
    if (r[1] === "titolo") { cella.setFontSize(18).setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#E2312B"); sh.setRowHeight(i + 1, 44); }
    if (r[1] === "sezione") { cella.setFontSize(13).setFontWeight("bold").setFontColor("#14306B").setBackground("#FFF1D6"); sh.setRowHeight(i + 1, 30); }
  });
  sh.setHiddenGridlines(true);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
}

function installaAutomazioni(ss) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "alCambioStato") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("alCambioStato").forSpreadsheet(ss).onEdit().create();
}

// Da eseguire dall'editor (menu Esegui) per controllare la configurazione delle email.
function verificaConfigurazione() {
  var alias = GmailApp.getAliases();
  Logger.log("Alias disponibili in Gmail: " + (alias.length ? alias.join(", ") : "nessuno"));
  Logger.log(alias.indexOf(EMAIL_MITTENTE) !== -1
    ? "OK: le email ai clienti partiranno da " + EMAIL_MITTENTE
    : "Attenzione: " + EMAIL_MITTENTE + " non è tra gli alias di Gmail, le email partiranno dall'indirizzo Gmail");
  var trigger = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === "alCambioStato"; });
  Logger.log(trigger ? "OK: conferma automatica attiva" : "Conferma automatica non attiva: esegui preparaFoglio");
}
