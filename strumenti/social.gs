/**
 * Circolandia – pubblicazione automatica su Facebook e Instagram (API ufficiali Meta)
 * Da aggiungere come secondo file (social.gs) nello stesso progetto Apps Script delle prenotazioni.
 *
 * Come funziona:
 *   - nel foglio "Piano social" scrivi data e ora, link dell'immagine (JPEG) e testo del post
 *   - ogni 15 minuti lo script pubblica i post con Stato "Da pubblicare" la cui ora è arrivata
 *   - dopo la pubblicazione scrive lo Stato "Pubblicato" e i link ai post
 *
 * Prima configurazione (una volta sola):
 *   1. esegui preparaSocial            → crea il foglio "Piano social" e l'automazione ogni 15 minuti
 *   2. in Impostazioni progetto › Proprietà script aggiungi TOKEN_UTENTE = il token Meta a lunga durata
 *   3. esegui collegaMeta              → salva il token della Pagina (che non scade) e cancella TOKEN_UTENTE
 *   4. esegui provaConnessione         → controlla Pagina e Instagram senza pubblicare nulla
 */

// ====== CONFIGURAZIONE ======
var META_VERSIONE_API = "v23.0";
var META_PAGINA_ID = "1397469816786693";       // Pagina Facebook Circolandia
var META_IG_ID_PREDEFINITO = "17841462737313261"; // @circolandia.terni (verificato da collegaMeta)
var SOCIAL_FOGLIO = "Piano social";
// ============================

var SOCIAL_COLONNE = ["Data e ora", "Immagine (link)", "Testo", "Facebook", "Instagram", "Stato",
                      "Link Facebook", "Link Instagram", "Note"];
var SOCIAL_STATI = ["Da pubblicare", "Pubblicato", "Errore", "Sospeso"];

function metaUrl(percorso) {
  return "https://graph.facebook.com/" + META_VERSIONE_API + "/" + percorso;
}

function metaChiamata(metodo, percorso, parametri) {
  var opz = { method: metodo, muteHttpExceptions: true };
  var url = metaUrl(percorso);
  if (metodo === "get") {
    url += "?" + Object.keys(parametri).map(function (k) {
      return encodeURIComponent(k) + "=" + encodeURIComponent(parametri[k]);
    }).join("&");
  } else {
    opz.payload = parametri;
  }
  var r = UrlFetchApp.fetch(url, opz);
  var dati = {};
  try { dati = JSON.parse(r.getContentText()); } catch (e) { dati = { error: { message: r.getContentText().slice(0, 300) } }; }
  if (dati.error) throw new Error(dati.error.message || JSON.stringify(dati.error));
  return dati;
}

function proprieta() { return PropertiesService.getScriptProperties(); }

function tokenPagina() {
  var t = proprieta().getProperty("META_TOKEN_PAGINA");
  if (!t) throw new Error("Token Meta mancante: esegui collegaMeta (vedi istruzioni in cima a social.gs).");
  return t;
}

function igId() {
  return proprieta().getProperty("META_IG_ID") || META_IG_ID_PREDEFINITO;
}

// ---------------------------------------------------------------------------
// Configurazione
// ---------------------------------------------------------------------------

function preparaSocial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SOCIAL_FOGLIO) || ss.insertSheet(SOCIAL_FOGLIO);
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, SOCIAL_COLONNE.length).setValues([SOCIAL_COLONNE]);
  sh.getRange(1, 1, 1, SOCIAL_COLONNE.length).setBackground("#1F5FBF").setFontColor("#FFFFFF").setFontWeight("bold").setWrap(true);
  sh.setFrozenRows(1);
  [150, 300, 420, 85, 85, 120, 200, 200, 260].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  var note = {
    1: "Quando pubblicare (es. 05/12/2026 20:30). Lo script controlla ogni 15 minuti.",
    2: "Link diretto a un'immagine JPEG pubblica, es. https://circolandia.it/social/post-01-apertura.jpg. Formato consigliato 1080×1350.",
    3: "Testo del post, con emoji e hashtag.",
    4: "Sì = pubblica sulla Pagina Facebook.",
    5: "Sì = pubblica su Instagram.",
    6: "Da pubblicare → lo script lo pubblica all'ora indicata. Sospeso = non pubblicare per ora.",
    9: "Messaggi dello script (errori o conferme)."
  };
  Object.keys(note).forEach(function (c) { sh.getRange(1, +c).setNote(note[c]); });

  var dv = SpreadsheetApp.newDataValidation;
  sh.getRange("A2:A").setNumberFormat("ddd dd/MM/yyyy HH:mm")
    .setDataValidation(dv().requireDate().setAllowInvalid(false).setHelpText("Data e ora, es. 05/12/2026 20:30").build());
  sh.getRange("B2:B").setDataValidation(dv().requireTextIsUrl().setAllowInvalid(false).setHelpText("Link https all'immagine JPEG").build());
  sh.getRange("C2:C").setWrap(true).setVerticalAlignment("top");
  ["D2:D", "E2:E"].forEach(function (a) {
    sh.getRange(a).setDataValidation(dv().requireValueInList(["Sì", "No"], true).setAllowInvalid(false).build());
  });
  sh.getRange("F2:F").setDataValidation(dv().requireValueInList(SOCIAL_STATI, true).setAllowInvalid(false).build());
  sh.getRange("I2:I").setWrap(true);

  var colori = { "Da pubblicare": "#FFF1D6", "Pubblicato": "#D6F2E3", "Errore": "#F8C9C5", "Sospeso": "#EEE6EA" };
  sh.setConditionalFormatRules(Object.keys(colori).map(function (s) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colori[s]).setRanges([sh.getRange("F2:F")]).build();
  }));

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "pubblicaSocialProgrammati") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("pubblicaSocialProgrammati").timeBased().everyMinutes(15).create();
  ss.toast("Foglio Piano social pronto, controllo automatico ogni 15 minuti.", "Circolandia", 6);
}

// Trasforma il token utente (a lunga durata) nel token della Pagina, che non scade.
function collegaMeta() {
  var p = proprieta();
  var tu = (p.getProperty("TOKEN_UTENTE") || "").trim();
  if (!tu) throw new Error("Aggiungi prima TOKEN_UTENTE in Impostazioni progetto › Proprietà script.");
  var pagine = metaChiamata("get", "me/accounts", { access_token: tu, fields: "id,name,access_token,instagram_business_account{id,username}" }).data || [];
  var pag = pagine.filter(function (x) { return x.id === META_PAGINA_ID; })[0] || pagine.filter(function (x) { return /circolandia/i.test(x.name); })[0];
  if (!pag) throw new Error("La Pagina Circolandia non è tra quelle autorizzate dal token. Pagine trovate: " + pagine.map(function (x) { return x.name; }).join(", "));
  p.setProperty("META_TOKEN_PAGINA", pag.access_token);
  if (pag.instagram_business_account) p.setProperty("META_IG_ID", pag.instagram_business_account.id);
  p.deleteProperty("TOKEN_UTENTE");
  Logger.log("OK: collegata la Pagina " + pag.name + (pag.instagram_business_account ? " e Instagram @" + pag.instagram_business_account.username : " (Instagram non trovato: controlla il collegamento in Business Suite)"));
}

// Controlla i collegamenti senza pubblicare nulla.
function provaConnessione() {
  var t = tokenPagina();
  var pag = metaChiamata("get", META_PAGINA_ID, { access_token: t, fields: "name" });
  var ig = metaChiamata("get", igId(), { access_token: t, fields: "username" });
  var lim = {};
  try { lim = metaChiamata("get", igId() + "/content_publishing_limit", { access_token: t, fields: "quota_usage,config" }); } catch (e) {}
  Logger.log("OK: Pagina «" + pag.name + "», Instagram @" + ig.username +
    (lim.data && lim.data[0] ? " · post via API nelle ultime 24 ore: " + lim.data[0].quota_usage : ""));
}

// ---------------------------------------------------------------------------
// Pubblicazione
// ---------------------------------------------------------------------------

function pubblicaSocialProgrammati() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOCIAL_FOGLIO);
    if (!sh || sh.getLastRow() < 2) return;
    var righe = sh.getRange(2, 1, sh.getLastRow() - 1, SOCIAL_COLONNE.length).getValues();
    var ora = new Date();
    righe.forEach(function (r, i) {
      if (String(r[5]).trim() !== "Da pubblicare") return;
      if (!(r[0] instanceof Date) || r[0] > ora) return;
      pubblicaRiga(sh, i + 2, r);
    });
  } finally {
    lock.releaseLock();
  }
}

// Pubblica subito la riga selezionata nel foglio Piano social (per prove o urgenze).
function pubblicaRigaSelezionata() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sh.getName() !== SOCIAL_FOGLIO) throw new Error("Seleziona una riga del foglio " + SOCIAL_FOGLIO);
  var n = sh.getActiveRange().getRow();
  if (n < 2) throw new Error("Seleziona una riga con un post");
  pubblicaRiga(sh, n, sh.getRange(n, 1, 1, SOCIAL_COLONNE.length).getValues()[0]);
}

function pubblicaRiga(sh, n, r) {
  var img = String(r[1]).trim(), testo = String(r[2] || ""), suFb = /^s/i.test(r[3]), suIg = /^s/i.test(r[4]);
  var note = [], linkFb = r[6], linkIg = r[7], ok = true;
  if (!/^https:\/\//.test(img)) { sh.getRange(n, 6).setValue("Errore"); sh.getRange(n, 9).setValue("Manca il link https all'immagine"); return; }
  if (!suFb && !suIg) { sh.getRange(n, 6).setValue("Errore"); sh.getRange(n, 9).setValue("Scegli Sì per Facebook e/o Instagram"); return; }
  sh.getRange(n, 6).setValue("Pubblicato");   // evita doppie pubblicazioni se la funzione riparte
  SpreadsheetApp.flush();

  if (suFb && !linkFb) {
    try { linkFb = pubblicaFacebook(img, testo); } catch (e) { ok = false; note.push("Facebook: " + e.message); }
  }
  if (suIg && !linkIg) {
    try { linkIg = pubblicaInstagram(img, testo); } catch (e) { ok = false; note.push("Instagram: " + e.message); }
  }
  sh.getRange(n, 7, 1, 2).setValues([[linkFb || "", linkIg || ""]]);
  sh.getRange(n, 6).setValue(ok ? "Pubblicato" : "Errore");
  sh.getRange(n, 9).setValue(ok ? "Pubblicato il " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : note.join(" · "));
}

function pubblicaFacebook(img, testo) {
  var r = metaChiamata("post", META_PAGINA_ID + "/photos", { url: img, caption: testo, access_token: tokenPagina() });
  return "https://www.facebook.com/" + (r.post_id || r.id);
}

function pubblicaInstagram(img, testo) {
  var t = tokenPagina(), ig = igId();
  var cont = metaChiamata("post", ig + "/media", { image_url: img, caption: testo, access_token: t });
  for (var i = 0; i < 20; i++) {
    var st = metaChiamata("get", cont.id, { fields: "status_code", access_token: t }).status_code;
    if (st === "FINISHED") break;
    if (st === "ERROR" || st === "EXPIRED") throw new Error("Instagram non ha accettato l'immagine (" + st + "): usa un JPEG 1080×1350 o 1080×1080");
    Utilities.sleep(3000);
  }
  var pub = metaChiamata("post", ig + "/media_publish", { creation_id: cont.id, access_token: t });
  try { return metaChiamata("get", pub.id, { fields: "permalink", access_token: t }).permalink || pub.id; }
  catch (e) { return "https://www.instagram.com/circolandia.terni/"; }
}
