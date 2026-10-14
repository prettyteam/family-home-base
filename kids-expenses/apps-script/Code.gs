/**
 * Kids' Shared Expenses: the private data behind the site.
 * Writes to an "Expense Log" tab in the KIDS 2026 Sheet (the one already shared with John).
 * It only adds that tab. The older tabs are read, never changed.
 * Receipts are saved to a private "Kids Expense Receipts" folder in Alexandra's Drive and are only
 * shown through this script, so a receipt link can't be passed around.
 * Passcodes are kept in Script Properties, never in this file or in the shared Sheet:
 *   OWNER_PIN  Alexandra: add, edit, mark paid, delete
 *   VIEW_PIN   John: view items, totals and receipts
 */
const SHEET_ID = "12D_wZzJon9UAr5Mr3ciltY-Y945ISETXblD1aJUm8rA"; // "KIDS 2026"
const LOG_TAB = "Expense Log";
const OLD_TAB = "Owed Expenses (Split Amt)";
const OLD_BALANCE_CELL = "B3"; // "JOHN OWES ALEX:" on the older tab
const HEAD = ["ID", "Date", "Child", "Category", "What", "Full amount", "John's half", "Status", "Paid on", "Note", "Receipt", "Added"];
const STATUSES = ["Owed", "Paid", "Declined"];
const MAX_RECEIPT_BYTES = 15 * 1024 * 1024;

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function role_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("fails") || 0);
  if (fails >= 20) throw new Error("locked");
  const p = PropertiesService.getScriptProperties();
  const owner = p.getProperty("OWNER_PIN"), viewer = p.getProperty("VIEW_PIN");
  if (pin && owner && pin === owner) return "owner";
  if (pin && viewer && pin === viewer) return "viewer";
  cache.put("fails", String(fails + 1), 900);
  throw new Error("bad_pin");
}

function log_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(LOG_TAB);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange("F:G").setNumberFormat("$#,##0.00");
  }
  return sh;
}

function money_(n) { return Math.round(Number(n) * 100) / 100; }
function day_(d) { return d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d || ""); }

function items_() {
  const vals = log_().getDataRange().getValues(), list = [];
  for (let i = 1; i < vals.length; i++) {
    const v = vals[i];
    if (!v[0]) continue;
    list.push({
      id: String(v[0]), date: day_(v[1]), child: String(v[2]), category: String(v[3]), what: String(v[4]),
      full: money_(v[5]), half: money_(v[6]), status: String(v[7] || "Owed"), paidOn: day_(v[8]),
      note: String(v[9] || ""), hasReceipt: !!v[10]
    });
  }
  return list;
}

function findRow_(id) {
  const sh = log_(), ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  throw new Error("That item no longer exists.");
}

/** The balance on the older tab, shown as its own line so nothing entered there is lost. */
function oldBalance_() {
  if (PropertiesService.getScriptProperties().getProperty("INCLUDE_OLD_BALANCE") === "no") return null;
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(OLD_TAB);
    const n = Number(sh.getRange(OLD_BALANCE_CELL).getValue());
    return isFinite(n) && n ? money_(n) : null;
  } catch (e) { return null; }
}

/**
 * The items behind the older tab's "JOHN OWES ALEX" total: every cell its SUM formula adds up,
 * with the label to its left, the child whose block it sits in and the column's category
 * (MEDICAL, SCHOOL or ACTIVITIES). Read only. Amounts on that tab are already John's half.
 */
function oldItems_() {
  if (PropertiesService.getScriptProperties().getProperty("INCLUDE_OLD_BALANCE") === "no") return [];
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(OLD_TAB);
    const f = sh.getRange(OLD_BALANCE_CELL).getFormula();
    if (!f) return [];
    const vals = sh.getDataRange().getValues(), col = c => c.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    const cells = [], re = /\$?([A-Z]{1,2})\$?(\d+)(?::\$?([A-Z]{1,2})\$?(\d+))?/g;
    let m;
    while ((m = re.exec(f.replace(/^=\s*SUM/i, "")))) {
      const c1 = col(m[1]), r1 = Number(m[2]), c2 = m[3] ? col(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) cells.push([r, c]);
    }
    // Each child's block starts with his name in column A, right above the MEDICAL / SCHOOL / ACTIVITIES header row.
    const a = vals.map(r => String(r[0]).trim()), blocks = [];
    for (let i = 0; i < a.length - 1; i++) if (a[i] && a[i + 1].toUpperCase() === "MEDICAL") blocks.push({ row: i + 1, kid: a[i].charAt(0) + a[i].slice(1).toLowerCase() });
    const seen = {}, list = [];
    cells.forEach(([r, c]) => {
      if (seen[r + ":" + c] || c < 2) return;
      seen[r + ":" + c] = 1;
      const amount = money_((vals[r - 1] || [])[c - 1]);
      if (!(amount > 0)) return;
      const b = blocks.filter(x => x.row < r).pop() || { row: 0, kid: "" };
      const head = String(((vals[b.row] || [])[c - 2]) || "").trim();
      list.push({ kid: b.kid, category: head ? head.charAt(0) + head.slice(1).toLowerCase() : "Other",
        what: String((vals[r - 1] || [])[c - 2] || "").trim() || "Item", amount: amount });
    });
    return list;
  } catch (e) { console.warn("Couldn't read the older tab: " + e.message); return []; }
}

/**
 * The boys' upcoming appointments and activities for the next 6 weeks, from the calendars the Skylight shows:
 * Alexandra's own Google calendar (only events that name one of the boys, look medical or school related,
 * or that John is invited to) and the "Family Home Base" calendar (activities and games, never nanny shifts).
 * Work events (weddings, clients, video calls) are always left out. Kept for 10 minutes.
 */
function upcoming_() {
  const cache = CacheService.getScriptCache(), hit = cache.get("upcoming");
  if (hit) return JSON.parse(hit);
  const kids = kids_(), from = new Date(); from.setHours(0, 0, 0, 0);
  const to = new Date(from.getTime() + 42 * 864e5);
  const kidRe = kids.length ? new RegExp("\\b(" + kids.join("|") + ")\\b", "i") : null;
  const medical = /appoint|appt|\bdr\.?\b|doctor|dentist|ortho|therap|pediatric|chop|advocare|psych|clinic|physical|well visit|neck|chiro|urgent|x-?ray|eye exam|vaccin|amen/i;
  const school = /no school|school|early dismissal|single session|conference|back to school|field trip/i;
  const work = /zoom\.us|meet\.google\.com|teams\.microsoft|hair|makeup|\bmu\b|brid|wedding|trial|taffeta|salon|client|bootcamp|webinar|working$/i;
  const viewer = String(PropertiesService.getScriptProperties().getProperty("VIEWER_EMAIL") || "jenks020@gmail.com").toLowerCase();
  const list = [], seen = {};
  const add = (ev, kind) => {
    const title = ev.getTitle() || "", loc = (ev.getLocation() || "").split("\n")[0].slice(0, 120);
    if (work.test(title + " " + loc + " " + (ev.getDescription() || ""))) return;
    const allDay = ev.isAllDayEvent(), s = allDay ? ev.getAllDayStartDate() : ev.getStartTime(), e = allDay ? ev.getAllDayEndDate() : ev.getEndTime();
    const key = title.toLowerCase() + "|" + s.getTime();
    if (seen[key]) return;
    seen[key] = 1;
    const kidM = kidRe && title.match(kidRe);
    const type = medical.test(title) ? "Medical" : school.test(title + " " + (ev.getDescription() || "")) ? "School" : kind;
    list.push({ title: title, start: s.toISOString(), end: e.toISOString(), allDay: allDay, location: loc,
      kid: kidM ? kidM[1].charAt(0).toUpperCase() + kidM[1].slice(1).toLowerCase() : "", type: type });
  };
  try {
    CalendarApp.getDefaultCalendar().getEvents(from, to).forEach(ev => {
      const title = ev.getTitle() || "", desc = ev.getDescription() || "";
      const johnToo = ev.getGuestList().some(g => g.getEmail().toLowerCase() === viewer) || String(ev.getCreators()[0] || "").toLowerCase() === viewer;
      if ((kidRe && kidRe.test(title)) || medical.test(title) || school.test(title + " " + desc) || johnToo) add(ev, "Activity");
    });
  } catch (e) { console.warn("Couldn't read Alexandra's calendar: " + e.message); }
  try {
    CalendarApp.getCalendarsByName("Family Home Base").forEach(cal => cal.getEvents(from, to).forEach(ev => add(ev, "Activity")));
  } catch (e) { console.warn("Couldn't read the Family Home Base calendar: " + e.message); }
  list.sort((x, y) => x.start.localeCompare(y.start));
  const out = list.slice(0, 120);
  cache.put("upcoming", JSON.stringify(out), 600);
  return out;
}

function kids_() {
  const set = PropertiesService.getScriptProperties().getProperty("KIDS");
  return set ? set.split(",").map(s => s.trim()).filter(String) : [];
}

function clean_(d) {
  const full = money_(d.full);
  if (!(full > 0)) throw new Error("Enter the full amount you paid.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.date || ""))) throw new Error("Pick a date.");
  if (!String(d.what || "").trim()) throw new Error("Say what it was for.");
  const status = STATUSES.indexOf(d.status) >= 0 ? d.status : "Owed";
  return {
    date: d.date, child: String(d.child || "").slice(0, 40), category: String(d.category || "Other").slice(0, 40),
    what: String(d.what).trim().slice(0, 200), full: full, half: money_(full / 2), status: status,
    paidOn: status === "Paid" ? (d.paidOn || day_(new Date())) : "", note: String(d.note || "").slice(0, 300)
  };
}

function folder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty("RECEIPT_FOLDER");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder("Kids Expense Receipts");
  p.setProperty("RECEIPT_FOLDER", f.getId());
  return f;
}

function saveReceipt_(r, rec) {
  if (!r || !r.b64) return "";
  const mime = String(r.mime || "");
  if (!/^image\/|^application\/pdf$/.test(mime)) throw new Error("Receipts can be a photo or a PDF.");
  const bytes = Utilities.base64Decode(r.b64);
  if (bytes.length > MAX_RECEIPT_BYTES) throw new Error("That file is too big. Try a photo instead.");
  const ext = mime === "application/pdf" ? ".pdf" : ".jpg";
  const name = (rec.date + " " + rec.what).replace(/[\\/:*?"<>|]/g, "-").slice(0, 80) + ext;
  return folder_().createFile(Utilities.newBlob(bytes, mime, name)).getId();
}

function row_(id, rec, receipt, added) {
  return [id, rec.date, rec.child, rec.category, rec.what, rec.full, rec.half, rec.status, rec.paidOn, rec.note, receipt, added];
}

function doGet() {
  return out_({ ok: true, app: "Kids' Shared Expenses" });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const role = role_(String(body.pin || ""));
    const act = body.action;
    if (act === "list") return out_({ ok: true, role: role, items: items_(), oldBalance: oldBalance_(), oldItems: oldItems_(), kids: kids_(), upcoming: upcoming_() });
    if (act === "receipt") {
      const fileId = log_().getRange(findRow_(body.id), 11).getValue();
      if (!fileId) throw new Error("No receipt on this one yet.");
      const blob = DriveApp.getFileById(fileId).getBlob();
      return out_({ ok: true, mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) });
    }
    if (role !== "owner") throw new Error("Only Alexandra can change this.");
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = log_();
      if (act === "add") {
        const rec = clean_(body.data || {});
        const id = Utilities.getUuid().slice(0, 8);
        sh.appendRow(row_(id, rec, saveReceipt_(body.receipt, rec), new Date()));
        return out_({ ok: true, id: id });
      }
      const row = findRow_(body.id), old = sh.getRange(row, 1, 1, HEAD.length).getValues()[0];
      if (act === "update") {
        const rec = clean_(body.data || {});
        let receipt = old[10];
        if (body.receipt && body.receipt.b64) {
          if (receipt) { try { DriveApp.getFileById(receipt).setTrashed(true); } catch (x) {} }
          receipt = saveReceipt_(body.receipt, rec);
        }
        sh.getRange(row, 1, 1, HEAD.length).setValues([row_(old[0], rec, receipt, old[11])]);
        return out_({ ok: true, id: old[0] });
      }
      if (act === "delete") {
        if (old[10]) { try { DriveApp.getFileById(old[10]).setTrashed(true); } catch (x) {} }
        sh.deleteRow(row);
        return out_({ ok: true });
      }
      throw new Error("Unknown request.");
    } finally { lock.releaseLock(); }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

/**
 * One-time setup: run this once. It makes the two passcodes (if they don't exist yet), creates the
 * Expense Log tab and the receipts folder, and puts the passcodes in a separate, private Sheet
 * called "Kids Expenses Passcodes" in Alexandra's Drive. They are never written to KIDS 2026,
 * because John can see that Sheet.
 */
function setup() {
  const p = PropertiesService.getScriptProperties();
  const make = () => Utilities.getUuid().replace(/-/g, "").slice(0, 8);
  if (!p.getProperty("OWNER_PIN")) p.setProperty("OWNER_PIN", make());
  if (!p.getProperty("VIEW_PIN")) p.setProperty("VIEW_PIN", make());
  if (!p.getProperty("KIDS")) p.setProperty("KIDS", kidsFromOldTab_().join(", "));
  log_();
  folder_();
  let ss;
  const sid = p.getProperty("PASSCODE_SHEET");
  try { ss = sid && SpreadsheetApp.openById(sid); } catch (e) { ss = null; }
  if (!ss) { ss = SpreadsheetApp.create("Kids Expenses Passcodes"); p.setProperty("PASSCODE_SHEET", ss.getId()); }
  const lines = [
    ["Who", "Passcode"],
    ["Alexandra (add and change)", p.getProperty("OWNER_PIN")],
    ["John (view only)", p.getProperty("VIEW_PIN")],
    ["Note", "Keep this Sheet private. To change a passcode, ask Claude."]
  ];
  const sh = ss.getSheets()[0];
  sh.clear();
  sh.getRange(1, 1, lines.length, 2).setValues(lines);
  console.log("Passcodes are in: " + ss.getUrl());
}

/** The older tab has each child's name in column A, right above a "MEDICAL" header row. */
function kidsFromOldTab_() {
  try {
    const a = SpreadsheetApp.openById(SHEET_ID).getSheetByName(OLD_TAB).getRange("A1:A400").getValues().map(r => String(r[0]).trim());
    const names = [];
    for (let i = 0; i < a.length - 1; i++) {
      if (a[i] && a[i + 1].toUpperCase() === "MEDICAL") names.push(a[i].charAt(0) + a[i].slice(1).toLowerCase());
    }
    return names;
  } catch (e) { return []; }
}
