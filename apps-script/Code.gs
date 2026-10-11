/**
 * Family Home Base: the private data behind the site.
 * Lives inside a Google Sheet (Extensions > Apps Script). Each tab holds one
 * kind of record: column A is the id, column B the record, column C when it changed.
 * Passcodes are kept in Script Properties (Project Settings), never in this file.
 *   OWNER_PIN  full access (Alexandra)
 *   NANNY_PIN  nanny access: clock in/out, fix her times, check off to-dos and homework
 */
const SHEET_ID = "1TFjO180dVUddqTtr2cAeDgspTEejo0u8aYnfrUS_Kh0"; // "Family Home Base Data" in Alexandra's Drive
const COLLS = ["settings", "shifts", "assignments", "grades", "activities", "todos"];
const NANNY_WRITES = { shifts: ["set", "update"], assignments: ["set", "update"], todos: ["set", "update", "delete"] };

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function role_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("fails") || 0);
  if (fails >= 20) throw new Error("locked");
  const p = PropertiesService.getScriptProperties();
  const owner = p.getProperty("OWNER_PIN"), nanny = p.getProperty("NANNY_PIN");
  if (pin && owner && pin === owner) return "mom";
  if (pin && nanny && pin === nanny) return "nanny";
  cache.put("fails", String(fails + 1), 900);
  throw new Error("bad_pin");
}

function sheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(["id", "record", "updated"]); }
  return sh;
}

function rows_(name) {
  const sh = sheet_(name), vals = sh.getDataRange().getValues(), list = [];
  for (let i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    try { list.push(Object.assign(JSON.parse(vals[i][1]), { _id: String(vals[i][0]) })); } catch (e) {}
  }
  return list;
}

/** Checked-off to-dos are removed for good a week after they're checked (the app hides them after 2 days). */
function purgeOldTodos_() {
  const cutoff = Date.now() - 7 * 864e5;
  rows_("todos").filter(t => t.done && t.doneAt && new Date(t.doneAt).getTime() < cutoff).forEach(t => write_("delete", "todos", t._id));
}

function readAll_(role) {
  purgeOldTodos_();
  const data = {};
  COLLS.forEach(c => data[c] = rows_(c));
  if (role === "nanny") {
    data.settings.forEach(s => { if (s.showPayToNanny !== "yes") { delete s.rate; delete s.weeklyPay; } });
  }
  return data;
}

function write_(op, coll, id, data) {
  const sh = sheet_(coll), vals = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === String(id)) { row = i + 1; break; }
  if (["shifts", "activities", "settings"].indexOf(coll) >= 0) PropertiesService.getScriptProperties().setProperty("CAL_DIRTY", "1");
  if (op === "delete") { if (row > 0) sh.deleteRow(row); return id; }
  if (!id) id = Utilities.getUuid().slice(0, 12);
  let rec = Object.assign({}, data || {});
  delete rec._id;
  if (op === "update") {
    if (row < 0) throw new Error("That item no longer exists.");
    rec = Object.assign(JSON.parse(vals[row - 1][1] || "{}"), rec);
  }
  const line = [id, JSON.stringify(rec), new Date()];
  if (row > 0) sh.getRange(row, 1, 1, 3).setValues([line]); else sh.appendRow(line);
  return id;
}

function doGet() {
  return out_({ ok: true, app: "Family Home Base" });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const role = role_(String(body.pin || ""));
    if (body.action === "list") return out_({ ok: true, role: role, data: readAll_(role) });
    const op = body.action, coll = body.coll;
    if (["set", "update", "delete"].indexOf(op) < 0 || COLLS.indexOf(coll) < 0) throw new Error("Unknown request.");
    if (role === "nanny" && (NANNY_WRITES[coll] || []).indexOf(op) < 0) throw new Error("Only Alexandra can change that.");
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { return out_({ ok: true, id: write_(op, coll, body.id, body.data) }); }
    finally { lock.releaseLock(); }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

/**
 * One-time setup: run this once. It makes the two passcodes (if they don't exist yet),
 * writes them to a "Passcodes" tab in the Sheet for Alexandra, and turns on the Skylight calendar.
 */
function setup() {
  const p = PropertiesService.getScriptProperties();
  const make = () => Utilities.getUuid().replace(/-/g, "").slice(0, 8);
  if (!p.getProperty("OWNER_PIN")) p.setProperty("OWNER_PIN", make());
  if (!p.getProperty("NANNY_PIN")) p.setProperty("NANNY_PIN", make());
  const sh = sheet_("Passcodes");
  sh.clear();
  sh.getRange(1, 1, 4, 2).setValues([
    ["Who", "Passcode"],
    ["Alexandra (full access)", p.getProperty("OWNER_PIN")],
    ["Nanny", p.getProperty("NANNY_PIN")],
    ["Note", "Keep this tab private. To change a passcode, ask Claude."]
  ]);
  COLLS.forEach(sheet_);
  setupCalendarSync();
}

/* ---------- Skylight / Google Calendar ---------- */

/** Run once from the editor to turn on the calendar feed for the Skylight. */
function setupCalendarSync() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "syncIfChanged").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("syncIfChanged").timeBased().everyMinutes(15).create();
  syncCalendar();
}

function syncIfChanged() {
  const p = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  // Also catch edits made straight in the Sheet (not through the app).
  const digest = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    JSON.stringify(["settings", "shifts", "activities"].map(c => sheet_(c).getDataRange().getValues()))));
  const changed = p.getProperty("CAL_DIGEST") !== digest;
  if (changed || p.getProperty("CAL_DIRTY") === "1" || p.getProperty("CAL_LAST_DAY") !== today) {
    syncCalendar();
    p.setProperty("CAL_DIGEST", digest);
  }
}

/** Copies the weekly activities and the nanny's shifts for the next 4 weeks into the "Family Home Base" calendar. */
function syncCalendar() {
  const p = PropertiesService.getScriptProperties();
  p.setProperty("CAL_DIRTY", "0");
  p.setProperty("CAL_LAST_DAY", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
  const calId = p.getProperty("CALENDAR_ID");
  let cal = calId && CalendarApp.getCalendarById(calId);
  if (!cal) { cal = CalendarApp.createCalendar("Family Home Base"); p.setProperty("CALENDAR_ID", cal.getId()); }

  const start = new Date(); start.setHours(0, 0, 0, 0);
  const end = new Date(start); end.setDate(end.getDate() + 28);
  cal.getEvents(start, end).forEach(ev => { if (ev.getTag("fhb")) ev.deleteEvent(); });

  const settings = rows_("settings").filter(r => r._id === "main")[0] || {};
  const kids = settings.kids || [], nanny = settings.nannyName || "Nanny";
  const kidName = id => (kids.filter(k => k.id === id)[0] || {}).name || "Family";
  const at = (d, hm) => { const [h, m] = String(hm).split(":").map(Number); const x = new Date(d); x.setHours(h || 0, m || 0, 0, 0); return x; };

  rows_("activities").forEach(a => {
    if (!a.start) return;
    for (let d = new Date(start); d < end; d.setDate(d.getDate() + 1)) {
      if (d.getDay() !== Number(a.day)) continue;
      const s = at(d, a.start), e = a.end ? at(d, a.end) : new Date(s.getTime() + 3600000);
      const notes = [a.driver ? a.driver + " drives" : "", a.bring ? "Bring: " + a.bring : ""].filter(String).join("\n");
      cal.createEvent(kidName(a.kid) + ": " + a.name, s, e, { location: a.location || "", description: notes }).setTag("fhb", "1");
    }
  });

  if (p.getProperty("SKYLIGHT_SHOW_SHIFTS") === "no") return;
  rows_("shifts").forEach(sh => {
    if (!sh.date || !sh.start || !sh.end) return;
    const d = new Date(sh.date + "T00:00:00");
    if (d < start || d >= end) return;
    const s = at(d, sh.start), e = at(d, sh.end);
    if (e <= s) e.setDate(e.getDate() + 1);
    cal.createEvent(nanny + " working", s, e, { description: sh.note || "" }).setTag("fhb", "1");
  });
}

/** Optional: run once from the editor to load example entries for a preview. */
function loadExamples() {
  EXAMPLES.forEach(([coll, id, data]) => write_("set", coll, id, data));
}

/** Run once from the editor to remove every example entry. */
function clearExamples() {
  COLLS.forEach(c => rows_(c).filter(r => r.example).forEach(r => write_("delete", c, r._id)));
}

const EXAMPLES = [["settings","main",{"nannyName": "Maria","kids": [{"id": "k1","name": "Ava","grade": "6th","school": "Hammonton Middle School"},{"id": "k2","name": "Leo","grade": "3rd","school": "Warren E. Sooy Jr. Elementary"}],"payType": "hourly","rate": 20,"weeklyPay": 0,"overtime": "no","showPayToNanny": "yes","houseNotes": "Example notes:\nBedtime: Leo 8:00, Ava 9:00\nLeo is allergic to tree nuts\nScreen time after homework only, 1 hour on school nights\nPickup list: Grandma Jo, Aunt Kate","example": true}],["shifts","ex-s0",{"date": "2026-10-05","start": "07:00","end": "15:30","inAt": "06:58","outAt": "15:40","breakMin": 30,"note": "","example": true}],["shifts","ex-s1",{"date": "2026-10-06","start": "07:00","end": "15:30","inAt": "07:05","outAt": "15:30","breakMin": 30,"note": "","example": true}],["shifts","ex-s2",{"date": "2026-10-07","start": "07:00","end": "15:30","inAt": "07:00","outAt": "16:15","breakMin": 30,"note": "","example": true}],["shifts","ex-s3",{"date": "2026-10-08","start": "07:00","end": "15:30","inAt": "06:55","outAt": "15:35","breakMin": 30,"note": "","example": true}],["shifts","ex-s4",{"date": "2026-10-09","start": "07:00","end": "15:30","inAt": "07:00","outAt": "15:30","breakMin": 30,"note": "","example": true}],["shifts","ex-s5",{"date": "2026-10-10","start": "18:00","end": "23:00","inAt": "17:50","outAt": "23:20","breakMin": 0,"note": "Date night","example": true}],["shifts","ex-n0",{"date": "2026-10-12","start": "07:00","end": "15:30","inAt": "","outAt": "","breakMin": 30,"note": "","example": true}],["shifts","ex-n1",{"date": "2026-10-13","start": "07:00","end": "15:30","inAt": "","outAt": "","breakMin": 30,"note": "","example": true}],["shifts","ex-n2",{"date": "2026-10-14","start": "07:00","end": "15:30","inAt": "","outAt": "","breakMin": 30,"note": "Early dismissal","example": true}],["shifts","ex-n3",{"date": "2026-10-15","start": "07:00","end": "15:30","inAt": "","outAt": "","breakMin": 30,"note": "","example": true}],["shifts","ex-n4",{"date": "2026-10-16","start": "07:00","end": "15:30","inAt": "","outAt": "","breakMin": 30,"note": "","example": true}],["assignments","ex-a0",{"kid": "k1","cls": "Science","title": "Volcano model","due": "2026-10-13","source": "Google Classroom","notes": "Needs poster board","done": false,"example": true}],["assignments","ex-a1",{"kid": "k1","cls": "ELA","title": "Chapter 5 reading log","due": "2026-10-12","source": "Google Classroom","notes": "","done": false,"example": true}],["assignments","ex-a2",{"kid": "k1","cls": "Math","title": "Unit 2 test","due": "2026-10-15","source": "Genesis portal","notes": "Study guide in folder","done": false,"example": true}],["assignments","ex-a3",{"kid": "k2","cls": "Spelling","title": "Week 6 spelling test","due": "2026-10-16","source": "Genesis portal","notes": "","done": false,"example": true}],["assignments","ex-a4",{"kid": "k2","cls": "Reading","title": "Reading log signed","due": "2026-10-12","source": "Added by hand","notes": "Parent signature","done": false,"example": true}],["assignments","ex-a5",{"kid": "k1","cls": "Social Studies","title": "Map worksheet","due": "2026-10-09","source": "Google Classroom","notes": "","done": true,"example": true}],["grades","ex-g0",{"kid": "k1","cls": "Math","grade": "94","updated": "2026-10-09","example": true}],["grades","ex-g1",{"kid": "k1","cls": "ELA","grade": "88","updated": "2026-10-09","example": true}],["grades","ex-g2",{"kid": "k1","cls": "Science","grade": "76","updated": "2026-10-09","example": true}],["grades","ex-g3",{"kid": "k1","cls": "Social Studies","grade": "91","updated": "2026-10-09","example": true}],["grades","ex-g4",{"kid": "k2","cls": "Reading","grade": "B+","updated": "2026-10-09","example": true}],["grades","ex-g5",{"kid": "k2","cls": "Math","grade": "A","updated": "2026-10-09","example": true}],["activities","ex-c0",{"kid": "k1","name": "Soccer practice","day": 1,"start": "17:30","end": "19:00","location": "Hammonton Lake Park","driver": "Maria","bring": "Cleats, water","example": true}],["activities","ex-c1",{"kid": "k1","name": "Soccer practice","day": 3,"start": "17:30","end": "19:00","location": "Hammonton Lake Park","driver": "Alexandra","bring": "Cleats, water","example": true}],["activities","ex-c2",{"kid": "k2","name": "Swim lessons","day": 2,"start": "16:00","end": "16:45","location": "Example Swim Club","driver": "Maria","bring": "Goggles, towel","example": true}],["activities","ex-c3",{"kid": "k2","name": "Cub Scouts","day": 4,"start": "18:30","end": "19:30","location": "Example church hall","driver": "Alexandra","bring": "","example": true}],["activities","ex-c4",{"kid": "k1","name": "Soccer game","day": 6,"start": "10:00","end": "11:30","location": "Away, check team app","driver": "Carpool","bring": "","example": true}],["activities","ex-c5",{"kid": "all","name": "Family dinner","day": 0,"start": "17:00","end": "","location": "Grandma's","driver": "Alexandra","bring": "","example": true}],["todos","ex-t0",{"text": "Sign Ava's field trip form","who": "mom","kid": "k1","due": "2026-10-13","done": false,"example": true}],["todos","ex-t1",{"text": "Pick up poster board for volcano","who": "nanny","kid": "k1","due": "2026-10-12","done": false,"example": true}],["todos","ex-t2",{"text": "Wash soccer uniform","who": "nanny","kid": "k1","due": "","done": false,"example": true}],["todos","ex-t3",{"text": "Schedule Leo's dentist","who": "mom","kid": "k2","due": "","done": false,"example": true}]];
