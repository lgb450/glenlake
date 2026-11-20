// U11 Glenlake Hawks 3 Blue - team leaderboard
// Paste this whole file into Extensions > Apps Script (replacing anything there), save, then Deploy > New deployment.

var SHEET_NAME = "Scores";
var MAX_SCORE = 1000;     // 12 plays, 40 to 100 points each with a streak, plus a 10 point perfect bonus. Change only if the Challenge changes.
var TOP_COUNT = 10;

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Jersey", "Best score", "Last updated"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 3).getValues().map(function (r, i) {
    return { row: i + 2, n: String(r[0]), score: Number(r[1]) || 0, t: r[2] ? new Date(r[2]).getTime() : 0 };
  });
}

function result_(list, n) {
  // highest score first; on a tie, whoever got there first ranks higher
  list.sort(function (a, b) { return b.score - a.score || a.t - b.t; });
  list.forEach(function (r) { r.rank = 1 + list.filter(function (z) { return z.score > r.score; }).length; });
  var me = null;
  list.forEach(function (r) { if (n && r.n === n) me = { rank: r.rank, best: r.score }; });
  return {
    ok: true,
    top: list.slice(0, TOP_COUNT).map(function (r) { return { rank: r.rank, n: r.n, score: r.score }; }),
    you: me,
    total: list.length
  };
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    var n = String(p.n || "").replace(/\D/g, "");
    if (n.length > 3) n = "";
    if (n) n = String(parseInt(n, 10));            // "07" and "7" are the same jersey
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sh = sheet_();
      if (p.action === "submit") {
        var s = parseInt(p.s, 10);
        if (!n || !isFinite(s) || s < 0 || s > MAX_SCORE) throw new Error("bad entry");
        var list = rows_(sh), found = null;
        list.forEach(function (r) { if (r.n === n) found = r; });
        if (!found) sh.appendRow([n, s, new Date()]);
        else if (s > found.score) sh.getRange(found.row, 2, 1, 2).setValues([[s, new Date()]]);
      }
      out = result_(rows_(sh), n);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
