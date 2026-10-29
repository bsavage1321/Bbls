// Paste this into Extensions > Apps Script in the Google Sheet you want to log to.
// Set TOKEN to a long random string, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Paste the same TOKEN into the app.
const TOKEN = "CHANGE-ME-TO-A-LONG-RANDOM-STRING";
const HEAD = ["Key","Macro","Week","Dates","Day","Exercise","Set","Target weight","Target reps","Weight","Reps","Done","Updated","ts"];

function tab_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange("A:A").setNumberFormat("@");   // keep keys like 0-1-2-3 as text, not dates
    s.appendRow(head);
    s.setFrozenRows(1);
  }
  return s;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if ((e.parameter || {}).token !== TOKEN) return json_({ error: "auth" });
  const log = {}, v = tab_("Log", HEAD).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[0]) continue;
    log[String(r[0])] = { w: String(r[9]), r: String(r[10]), x: r[11] === true, t: Number(r[13]) || 0 };
  }
  const sv = tab_("Settings", ["Key","Value","ts"]).getDataRange().getValues();
  let settings = null;
  if (sv.length > 1 && sv[1][1]) { settings = JSON.parse(sv[1][1]); settings.t = Number(sv[1][2]) || 0; }
  return json_({ log: log, settings: settings });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.token !== TOKEN) return json_({ error: "auth" });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.rows && b.rows.length) {
      const sh = tab_("Log", HEAD), v = sh.getDataRange().getValues(), idx = {}, add = [];
      for (let i = 1; i < v.length; i++) idx[String(v[i][0])] = i;
      b.rows.forEach(function (r) {
        const vals = r.slice(0, 12).concat([new Date(r[12]), r[12]]);
        const i = idx[String(r[0])];
        if (i === undefined) add.push(vals);
        else if (r[12] > (Number(v[i][13]) || 0)) sh.getRange(i + 1, 1, 1, 14).setValues([vals]);
      });
      if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 14).setValues(add);
    }
    if (b.settings) {
      const sh = tab_("Settings", ["Key","Value","ts"]);
      const t = Number(sh.getRange("C2").getValue()) || 0;
      if (b.settings.t > t) sh.getRange("A2:C2").setValues([["settings", JSON.stringify(b.settings), b.settings.t]]);
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
