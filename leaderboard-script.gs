// U11 Glenlake Hawks - team leaderboard (players and parents)
// Paste this whole file into Extensions > Apps Script (replacing anything there), save, then Deploy > New deployment.

var SHEET_NAME = "Scores";
var MAX_SCORE = 1000;     // 12 plays, 40 to 100 points each with a streak, plus a 10 point perfect bonus. Change only if the Challenge changes.
var TOP_COUNT = 10;

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Jersey", "Best score", "Last updated", "Team", "Role"]);
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, 4).getValue() === "") {
    sh.getRange(1, 4, 1, 2).setValues([["Team", "Role"]]);   // older sheet: add the new columns
  }
  return sh;
}

function rows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues().map(function (r, i) {
    return { row: i + 2, n: String(r[0]), score: Number(r[1]) || 0, t: r[2] ? new Date(r[2]).getTime() : 0,
             team: String(r[3] || ""), role: String(r[4]) === "parent" ? "parent" : "player" };
  });
}

function same_(r, n, team, role) { return r.n === n && r.team.toLowerCase() === team.toLowerCase() && r.role === role; }

function result_(all, me, view) {
  var list = view === "all" ? all.slice() : all.filter(function (r) { return r.role === view; });
  // highest score first; on a tie, whoever got there first ranks higher
  list.sort(function (a, b) { return b.score - a.score || a.t - b.t; });
  list.forEach(function (r) { r.rank = 1 + list.filter(function (z) { return z.score > r.score; }).length; });
  var you = null, fam = null;
  if (me && me.n) {
    list.forEach(function (r) { if (same_(r, me.n, me.team, me.role)) you = { rank: r.rank, best: r.score }; });
    fam = { player: null, parent: null };
    all.forEach(function (r) { if (r.n === me.n && r.team.toLowerCase() === me.team.toLowerCase()) fam[r.role] = r.score; });
  }
  return {
    ok: true,
    top: list.slice(0, TOP_COUNT).map(function (r) { return { rank: r.rank, n: r.n, t: r.team, r: r.role, score: r.score }; }),
    you: you,
    family: fam,
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
    var team = String(p.t || "").replace(/[^A-Za-z0-9 \-&.']/g, "").replace(/\s+/g, " ").trim().slice(0, 20);
    var role = p.r === "parent" ? "parent" : "player";
    var view = p.v === "player" || p.v === "parent" ? p.v : "all";
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sh = sheet_();
      if (p.action === "submit") {
        var s = parseInt(p.s, 10);
        if (!n || !isFinite(s) || s < 0 || s > MAX_SCORE) throw new Error("bad entry");
        var list = rows_(sh), found = null;
        list.forEach(function (r) { if (same_(r, n, team, role)) found = r; });
        if (!found) sh.appendRow([n, s, new Date(), team, role]);
        else if (s > found.score) sh.getRange(found.row, 2, 1, 2).setValues([[s, new Date()]]);
      }
      out = result_(rows_(sh), n ? { n: n, team: team, role: role } : null, view);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
