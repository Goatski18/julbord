// Julbordet 2026: sparar tingningar och OSA i det här kalkylarket.
// Klistra in i Tillägg > Apps Script och distribuera som webbapp (Alla har åtkomst).
var TABS = {
  claims: { name: "Tingningar", head: ["Rätt", "Sällskap", "Tid"] },
  extras: { name: "Egna rätter", head: ["Id", "Tur", "Rätt", "Sällskap", "Tid"] },
  osa:    { name: "OSA", head: ["Tid", "Sällskap", "Kommer", "Vuxna", "Barn", "Övrigt"] }
};

function tab(k) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(TABS[k].name);
  if (!sh) { sh = ss.insertSheet(TABS[k].name); sh.appendRow(TABS[k].head); sh.setFrozenRows(1); }
  return sh;
}
function rows(k) { return tab(k).getDataRange().getValues().slice(1); }
function clean(v) { var s = String(v == null ? "" : v).slice(0, 300); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function state() {
  return {
    claims: rows("claims").map(function (r) { return { item: String(r[0]), who: String(r[1]) }; }),
    extras: rows("extras").map(function (r) { return { id: String(r[0]), group: String(r[1]), name: String(r[2]), who: String(r[3]) }; })
  };
}
function removeWhere(k, col, value) {
  var sh = tab(k), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(v[i][col]) === value) sh.deleteRow(i + 1);
}

function doGet() { return out(state()); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var b = JSON.parse(e.postData.contents), now = new Date();
    if (b.action === "claim") {
      var taken = rows("claims").some(function (r) { return String(r[0]) === b.item; });
      if (taken) return out({ ok: false, error: "taken", claims: state().claims, extras: state().extras });
      tab("claims").appendRow([clean(b.item), clean(b.who), now]);
    } else if (b.action === "unclaim") {
      removeWhere("claims", 0, String(b.item));
    } else if (b.action === "extra") {
      tab("extras").appendRow([Utilities.getUuid(), clean(b.group), clean(b.name), clean(b.who), now]);
    } else if (b.action === "unextra") {
      removeWhere("extras", 0, String(b.id));
    } else if (b.action === "osa") {
      tab("osa").appendRow([now, clean(b.who), clean(b.coming), Number(b.adults) || 0, Number(b.kids) || 0, clean(b.notes)]);
      return out({ ok: true });
    }
    var s = state(); s.ok = true; return out(s);
  } finally {
    lock.releaseLock();
  }
}
