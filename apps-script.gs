// Quiz responses: paste this into Extensions > Apps Script in your Google Sheet.
// 1. Change RESULTS_KEY below to a private word only you know.
// 2. Deploy > New deployment > type "Web app", Execute as "Me", Who has access "Anyone".
// 3. Copy the Web app URL (ends in /exec) and send it to Claude.
const RESULTS_KEY = "change-me";

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(["Submitted", "Name", "Score", "Total", "Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8"]);
  return sh;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const clean = v => String(v == null ? "" : v).replace(/^[=+\-@]/, "'$&").slice(0, 300);
  const answers = (Array.isArray(d.answers) ? d.answers : []).slice(0, 8).map(clean);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { sheet_().appendRow([new Date(), clean(d.name).slice(0, 80), Number(d.score) || 0, Number(d.total) || 0].concat(answers)); }
  finally { lock.releaseLock(); }
  return ContentService.createTextOutput("ok");
}

function doGet(e) {
  if (!e.parameter.key || e.parameter.key !== RESULTS_KEY) {
    return ContentService.createTextOutput(JSON.stringify({error: "wrong key"})).setMimeType(ContentService.MimeType.JSON);
  }
  const rows = sheet_().getDataRange().getValues();
  const head = rows.shift();
  const out = rows.map(r => ({submitted: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), name: r[1], score: r[2], total: r[3], answers: r.slice(4).filter(x => x !== "")}));
  return ContentService.createTextOutput(JSON.stringify({rows: out})).setMimeType(ContentService.MimeType.JSON);
}
