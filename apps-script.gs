/*
  CAPEREFER FEEDBACK: receives feedback, new areas, notes to the developer, daily search counts and anonymous
  daily visits (unique users),
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
  The tabs (Daily searches, Daily visits, Feedback, New areas, Notes) appear by themselves with the first message.

  CHANGING THIS SCRIPT LATER: paste the new version, Save, then Deploy > Manage deployments > pencil >
  Version: New version > Deploy. That keeps the same URL ("New deployment" would make a new one).
*/
const TZ = "Africa/Johannesburg";
const TABS = {
  search: ["Daily searches", ["Date", "Searches", "Not found"]],
  // each phone checks in at most once a day, saying which of these firsts it is (no ID): sums give unique users
  visit: ["Daily visits", ["Date", "Users today", "New users", "First this week", "First this month", "Installed app", "Browser", "iPhone", "Android", "Other"]],
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

// one row per day: the last row is today, or a new row of zeros is started; adds `add` to its counts
function addToday_(sh, now, add) {
  const today = Utilities.formatDate(now, TZ, "yyyy-MM-dd"), last = sh.getLastRow();
  const v = last > 1 ? sh.getRange(last, 1).getValue() : "";
  const lastDay = v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd") : String(v).replace(/^'/, "");
  let row = last;
  if (lastDay !== today) { sh.appendRow(["'" + today].concat(add.map(() => 0))); row = last + 1; }
  const r = sh.getRange(row, 2, 1, add.length), n = r.getValues()[0];
  r.setValues([n.map((x, i) => Number(x) + add[i])]);
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
      addToday_(sh, now, [1, d.found ? 0 : 1]);
    } else if (d.type === "visit") {
      if (!d.day) return reply_(false);
      const os = d.os === "iPhone" || d.os === "Android" ? d.os : "Other";
      addToday_(sh, now, [1, d.new ? 1 : 0, d.week ? 1 : 0, d.month ? 1 : 0, d.mode === "app" ? 1 : 0, d.mode === "app" ? 0 : 1,
        os === "iPhone" ? 1 : 0, os === "Android" ? 1 : 0, os === "Other" ? 1 : 0]);
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

/* DAILY REPORT EMAIL: the day's numbers to the developer every evening at about 8 pm (Cape Town).
   Set up once: in the Apps Script editor choose setupDailyReport in the function list, click Run, and allow it to
   send email and run on a schedule. It sends one report straight away to show it works. */
const REPORT_TO = "caspergermishuys2@gmail.com";

function dailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), now = new Date();
  const day = d => d instanceof Date ? Utilities.formatDate(d, TZ, "yyyy-MM-dd") : String(d).replace(/^'/, "");
  const ago = n => day(new Date(now.getTime() - n * 864e5));
  const today = day(now), weekAgo = ago(6), monday = ago(Number(Utilities.formatDate(now, TZ, "u")) - 1);
  const rows = name => { const sh = ss.getSheetByName(name);
    return sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : []; };
  const sum = (rs, col, keep) => rs.filter(keep || (() => true)).reduce((s, r) => s + (Number(r[col]) || 0), 0);
  const isToday = r => day(r[0]) === today;
  const S = rows("Daily searches"), V = rows("Daily visits"), v = V.find(isToday) || [];
  const fb = rows("Feedback").filter(isToday), areas = rows("New areas").filter(isToday), notes = rows("Notes").filter(isToday);
  const waiting = rows("New areas").filter(r => r[1] && !r[6]).length;
  const n = (r, i) => Number(r[i]) || 0, all = sum(S, 1), missed = sum(S, 2);
  const flagged = fb.filter(r => r[1] === "wrong");
  const ok = fb.length - flagged.length, list = (rs, f) => rs.map(f).join(", ");
  const lines = [
    `Today: ${sum(S, 1, isToday)} searches${sum(S, 2, isToday) ? ` (${sum(S, 2, isToday)} not found)` : ""}, ${n(v, 1)} users (${n(v, 2)} new)`,
    `Last 7 days: ${sum(S, 1, r => day(r[0]) >= weekAgo)} searches · this week ${sum(V, 3, r => day(r[0]) >= monday)} users`,
    `All time: ${all} searches (${all ? Math.round(100 * (all - missed) / all) : 0}% found), ${sum(V, 2)} users`,
    `Staff today: ${ok} correct, ${flagged.length} flagged, ${areas.length} areas, ${notes.length} notes`,
    ...flagged.map(r => `• Flagged ${r[2]}: shown ${r[3]}, should be ${r[4] || "not sure"}${r[5] ? " (" + r[5] + ")" : ""}`),
    ...(areas.length ? ["• Areas: " + list(areas, r => r[1] + (r[2] && r[2] !== "Not known" ? " → " + r[2] : ""))] : []),
    ...notes.map(r => `• Note: ${r[1]}${r[3] ? " (" + r[3] + ")" : ""}`),
    ...(waiting ? [`${waiting} new areas waiting for your decision`] : []),
    "",
    ss.getUrl(),
  ];
  MailApp.sendEmail(REPORT_TO, `CapeRefer ${Utilities.formatDate(now, TZ, "d MMM")}: ${sum(S, 1, isToday)} searches · ${n(v, 1)} users`, lines.join("\n"));
}

// run once from the editor: one report every evening between 8 and 9 pm, Cape Town time (and one now, to check)
function setupDailyReport() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "dailyReport").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("dailyReport").timeBased().everyDays(1).atHour(20).inTimezone(TZ).create();
  dailyReport();
}
