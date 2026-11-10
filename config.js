/* ============================================================
   CIRCOLANDIA – configurazione del modulo di prenotazione
   È l'unico file da modificare per aggiornare le date occupate.
   ============================================================ */
window.CIRCOLANDIA = {

  // Indirizzo dello script Google che riceve le richieste (finisce con /exec).
  // Finché è vuoto il modulo resta visibile ma l'invio è disattivato.
  FORM_ENDPOINT: "https://script.google.com/macros/s/AKfycbwXZa5_pOV9wLSUWrRNK0nV-4kitTafyCRrWQH5S6Y4YiMpXrv36tCyTuhfeJlwPIs/exec",

  // Fasce già prenotate. Formato: "AAAA-MM-GG": ["mattina", "pomeriggio", "sera"]
  // oppure "AAAA-MM-GG": "tutto" per bloccare l'intera giornata.
  DATE_OCCUPATE: {
    // "2026-11-21": ["pomeriggio"],
    // "2026-12-25": "tutto",
  },

  // Giorni minimi di anticipo per una richiesta (1 = da domani).
  ANTICIPO_MINIMO_GIORNI: 1,

  // Fino a quanti mesi avanti si può chiedere una data.
  MESI_PRENOTABILI: 12
};
