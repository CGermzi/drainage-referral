/*
  REFERRAL PATHWAY FEEDBACK: receives feedback, new areas, notes to the developer and daily search counts
  from the web page, and writes them into a Google Sheet in your own Google account.

  ONE-TIME SETUP (about 5 minutes)
    1. Open the Google Sheet "Drainage app DATA" (its tabs are already set up).
    2. Extensions > Apps Script. Delete what is there, paste ALL of this file, click Save (the disk icon).
    3. Deploy > New deployment > gear icon > Web app.
         Execute as: Me      Who has access: Anyone
       Deploy > Authorize access > your Google account > Advanced > "Go to ... (unsafe)" > Allow.
       ("unsafe" only means Google hasn't reviewed a script you wrote yourself.)
    4. Copy the Web app URL (ends in /exec). Open it in your browser: it should say "Referral pathway feedback is working".
    5. Paste it into index.html at  const FEEDBACK_URL = "";  (or send it to Claude to do).
  The tabs (Daily searches, Feedback, New areas, Notes) appear by themselves with the first message.

  CHANGING THIS SCRIPT LATER: paste the new version, Save, then Deploy > Manage deployments > pencil >
  Version: New version > Deploy. That keeps the same URL ("New deployment" would make a new one).
*/
const TZ = "Africa/Johannesburg";
const TABS = {
  search: ["Daily searches", ["Date", "Searches", "Not found"]],
  feedback: ["Feedback", ["Time", "Verdict", "Area(s) shown", "Pathway shown", "Should be", "Comment", "From"]],
  area: ["New areas", ["Time", "Area", "Pathway", "Town / district", "How they know", "From", "Your decision"]],
  note: ["Notes", ["Time", "Note", "From", "Contact"]],
};

function tab_(type) {
  const [name, head] = TABS[type], ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight("bold");
  }
  return sh;
}

// text from a public page: kept short, and never run as a spreadsheet formula
function clean_(v) {
  const s = String(v == null ? "" : v).slice(0, 1000).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply_(ok) {
  return ContentService.createTextOutput(JSON.stringify({ ok: ok })).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput("Referral pathway feedback is working.");
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return reply_(false); }
  if (!d || !TABS[d.type]) return reply_(false);
  if (d.website) return reply_(true);   // a hidden field only robots fill in
  if (d.type === "area" && !(String(d.area || "").trim() && d.path)) return reply_(false);
  if (d.type === "note" && !String(d.note || "").trim()) return reply_(false);
  if (d.type === "feedback" && !/^(correct|wrong)$/.test(d.verdict)) return reply_(false);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = tab_(d.type), now = new Date(), c = clean_;
    if (d.type === "search") {
      // one row per day: the last row is today, or a new row is started
      const today = Utilities.formatDate(now, TZ, "yyyy-MM-dd"), last = sh.getLastRow();
      const v = last > 1 ? sh.getRange(last, 1).getValue() : "";
      const lastDay = v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd") : String(v).replace(/^'/, "");
      let row = last;
      if (lastDay !== today) { sh.appendRow(["'" + today, 0, 0]); row = last + 1; }
      const r = sh.getRange(row, 2, 1, 2), n = r.getValues()[0];
      r.setValues([[Number(n[0]) + 1, Number(n[1]) + (d.found ? 0 : 1)]]);
    } else if (d.type === "feedback") {
      sh.appendRow([now, d.verdict, c(d.areas), c(d.pathway), c(d.correct), c(d.comment), c(d.from)]);
    } else if (d.type === "area") {
      sh.appendRow([now, c(d.area), c(d.path), c(d.district), c(d.note), c(d.from), ""]);
    } else {
      sh.appendRow([now, c(d.note), c(d.from), c(d.contact)]);
    }
  } finally {
    lock.releaseLock();
  }
  return reply_(true);
}
