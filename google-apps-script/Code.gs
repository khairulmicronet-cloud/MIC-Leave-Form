/**
 * Staff Leave Breakdown — Google Apps Script backend.
 *
 * Paste this whole file into Extensions > Apps Script inside your Google
 * Sheet (see SETUP-STAFF-LEAVE-BREAKDOWN.md for the full walkthrough),
 * then deploy it as a Web App. The deployed URL is what you paste into
 * config.js as LEAVE_LEDGER_API_URL.
 *
 * Sheets/tabs (created automatically on first run if missing):
 *
 *   StaffConfig: Staff | Year | AnnualEntitlement | AnnualCarryForward | SickEntitlement | SickCarryForward
 *   LeaveEntries: ID | Staff | StartDate | EndDate | Description | AnnualDays | SickDays | UnpaidDays | HospitalizeDays | CreatedAt
 *   StaffAuth: Staff | PinHash | Role | GoogleEmail | PinIsTemporary
 *
 * Access control:
 *   Every request must carry a signed token (obtained via the "login"
 *   action, staff name + PIN). A token proves who the caller is and
 *   whether they're "admin" or plain "staff":
 *     - admin: sees and edits every staff member's data (same as before
 *       this feature existed).
 *     - staff: sees ONLY their own entries/config, read-only. addEntry,
 *       upsertConfig (entitlement / carry-forward) and the spreadsheet
 *       import are all admin-only — only an admin tallies leave against
 *       entitlement, so there's no risk of a duplicate or stray entry
 *       from the staff member themselves. deleteEntry is admin-only for
 *       the same reason.
 *
 *   PIN model:
 *     - Every StaffAuth row has a PinIsTemporary flag. When TRUE, the
 *       front end forces that person to set their own permanent PIN
 *       (action "changePin") the moment they sign in — the PIN an admin
 *       handed them only works once.
 *     - A single MASTER PIN (its hash stored in Script Properties, see
 *       getMasterPinHash_()) lets an admin sign in AS any staff member by
 *       picking that person's name on the login screen and entering the
 *       master PIN instead of their personal PIN. That session is
 *       tagged viaMaster: true in its token. If the admin uses that
 *       session to change the PIN, the new PIN is left flagged temporary
 *       again (it's a reset on the person's behalf, not their own
 *       permanent choice) — whereas a person changing their own PIN
 *       (viaMaster: false) clears the temporary flag.
 *
 *   Run seedStaffAuth_() once (from the Apps Script editor's Run button,
 *   not via the web app) to create the StaffAuth sheet and generate a
 *   random PIN for every name in LEAVE_FORM_STAFF_NAMES, plus whichever
 *   admin names you list in ADMIN_STAFF_NAMES below. The generated PINs
 *   are written to the Execution log — copy them from there, they are
 *   NOT stored anywhere in plain text after that (only a hash is kept).
 *   New rows seed with PinIsTemporary = TRUE. Run setMasterPin_() once to
 *   generate the master PIN the same way (also logged once, never stored
 *   in plain text).
 */

const SHEET_STAFF_CONFIG = "StaffConfig";
const SHEET_ENTRIES = "LeaveEntries";
const SHEET_AUTH = "StaffAuth";

const STAFF_CONFIG_HEADERS = [
  "Staff", "Year", "AnnualEntitlement", "AnnualCarryForward", "SickEntitlement", "SickCarryForward"
];
const ENTRIES_HEADERS = [
  "ID", "Staff", "StartDate", "EndDate", "Description",
  "AnnualDays", "SickDays", "UnpaidDays", "HospitalizeDays", "CreatedAt"
];
const AUTH_HEADERS = ["Staff", "PinHash", "Role", "GoogleEmail", "PinIsTemporary"];

// Columns that hold a calendar date only (no time-of-day is meaningful).
// These are written as plain "yyyy-MM-dd" strings from the client, but
// Google Sheets' automatic type-detection can silently convert them into
// a real Date cell. When that happens, Sheets anchors the value at
// midnight UTC — NOT the spreadsheet's local timezone — so these must be
// read back using UTC too, or the calendar day shifts forward by one
// once formatted in a timezone ahead of UTC (e.g. "2026-02-06" becoming
// "2026-02-07"). Always format these as UTC to recover the exact date
// that was originally entered.
const DATE_ONLY_FIELDS = ["StartDate", "EndDate"];
const DATE_ONLY_TZ = "Etc/GMT";

// Columns that hold a genuine timestamp (when a row was created), where
// the local timezone is what a human reading it would expect.
const TIMESTAMP_FIELDS = ["CreatedAt"];

// Every staff first name that should get a login (kept in sync by hand
// with LEAVE_FORM_STAFF_NAMES in the web app's config.js — Apps Script
// projects and the GitHub-hosted web app are separate JS environments,
// so this list can't be shared automatically between them).
const ALL_STAFF_NAMES = [
  "Khairul", "Amal", "Nurlizam", "Lyana", "Veronica", "Hariz",
  "Aqilah", "Norain", "Maziyah", "Alisha", "Aslam", "Sheraden",
  "Ummi", "Izzaty", "Azimah", "Nurkhtamal", "Nurzahidah", "Nur Amelea",
  "Mustadim", "Sharon", "Crisanta", "Kalau"
];

// Names to seed as "admin" role in StaffAuth (case-insensitive match
// against the Staff column, must also appear in ALL_STAFF_NAMES above).
// Edit this list, then re-run seedStaffAuth_() if you need to add another
// admin later — it only fills in rows that don't already have a PinHash,
// so re-running is safe.
const ADMIN_STAFF_NAMES = ["Khairul", "Maziyah"];

// Secret used to sign/verify session tokens (HMAC-SHA256). This is
// generated once by seedStaffAuth_() and stored in Script Properties —
// never hard-code a secret in source. See getTokenSecret_().
function getTokenSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty("TOKEN_SECRET");
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty("TOKEN_SECRET", secret);
  }
  return secret;
}

// The master PIN's hash, or null if it hasn't been set up yet (see
// setMasterPin_() below). Stored in Script Properties, never in the
// spreadsheet — it isn't tied to any one staff name.
function getMasterPinHash_() {
  return PropertiesService.getScriptProperties().getProperty("MASTER_PIN_HASH") || null;
}

function doGet(e) {
  try {
    const auth = requireAuth_(e.parameter.token);
    const configSheet = getOrCreateSheet_(SHEET_STAFF_CONFIG, STAFF_CONFIG_HEADERS);
    const entriesSheet = getOrCreateSheet_(SHEET_ENTRIES, ENTRIES_HEADERS);
    let config = sheetToObjects_(configSheet);
    let entries = sheetToObjects_(entriesSheet);
    if (auth.role !== "admin") {
      config = config.filter(function (c) { return sameStaff_(c.Staff, auth.staff); });
      entries = entries.filter(function (en) { return sameStaff_(en.Staff, auth.staff); });
    }
    return jsonResponse_({ ok: true, role: auth.role, staff: auth.staff, config: config, entries: entries });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    // Actions that don't require an existing session.
    if (action === "login") {
      return jsonResponse_(login_(body.staff, body.pin));
    }

    const auth = requireAuth_(body.token);

    if (action === "addEntry") {
      requireAdmin_(auth);
      addEntry_(body.entry, auth);
    } else if (action === "upsertConfig") {
      requireAdmin_(auth);
      upsertConfig_(body.config);
    } else if (action === "deleteEntry") {
      requireAdmin_(auth);
      deleteEntry_(body.id, auth);
    } else if (action === "changePin") {
      changePin_(body.newPin, auth);
    } else {
      return jsonResponse_({ ok: false, error: "Unknown action: " + action });
    }

    return jsonResponse_({ ok: true });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

// ---------------------------------------------------------------------
// Auth
// ---------------------------------------------------------------------

function sameStaff_(a, b) {
  return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();
}

function sha256Hex_(str) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function base64url_(bytesOrStr) {
  const raw = typeof bytesOrStr === "string"
    ? Utilities.base64Encode(bytesOrStr, Utilities.Charset.UTF_8)
    : Utilities.base64Encode(bytesOrStr);
  return raw.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function base64urlDecodeToString_(s) {
  let b64 = s.replace(/-/g, "+").replace(/_/g, "/");
  while (b64.length % 4) b64 += "=";
  return Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString();
}

function signToken_(payloadObj) {
  const payloadJson = JSON.stringify(payloadObj);
  const payloadPart = base64url_(payloadJson);
  const sigBytes = Utilities.computeHmacSha256Signature(payloadPart, getTokenSecret_());
  const sigPart = base64url_(sigBytes);
  return payloadPart + "." + sigPart;
}

function verifyToken_(token) {
  if (!token || token.indexOf(".") === -1) return null;
  const parts = token.split(".");
  if (parts.length !== 2) return null;
  const payloadPart = parts[0];
  const sigPart = parts[1];
  const expectedSig = base64url_(Utilities.computeHmacSha256Signature(payloadPart, getTokenSecret_()));
  if (expectedSig !== sigPart) return null;
  let payload;
  try {
    payload = JSON.parse(base64urlDecodeToString_(payloadPart));
  } catch (e) {
    return null;
  }
  if (!payload || !payload.staff || !payload.role || !payload.iat) return null;
  // Tokens are valid for 30 days.
  const ageMs = Date.now() - Number(payload.iat);
  if (ageMs < 0 || ageMs > 30 * 24 * 60 * 60 * 1000) return null;
  return payload;
}

function requireAuth_(token) {
  const payload = verifyToken_(token);
  if (!payload) throw new Error("Not signed in. Please sign in again.");
  return payload;
}

function requireAdmin_(auth) {
  if (auth.role !== "admin") throw new Error("Admins only.");
}

function login_(staff, pin) {
  if (!staff || !pin) return { ok: false, error: "Staff and PIN are required." };
  const sheet = getOrCreateSheet_(SHEET_AUTH, AUTH_HEADERS);
  const values = sheet.getDataRange().getValues();
  const pinHash = sha256Hex_(String(pin).trim());
  const masterHash = getMasterPinHash_();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!sameStaff_(row[0], staff)) continue;

    const role = String(row[2] || "staff").trim().toLowerCase() === "admin" ? "admin" : "staff";

    // 1) That person's own PIN.
    if (String(row[1]) === pinHash) {
      const mustChangePin = String(row[4] || "").trim().toUpperCase() === "TRUE";
      const token = signToken_({ staff: row[0], role: role, iat: Date.now(), viaMaster: false });
      return { ok: true, token: token, role: role, staff: row[0], mustChangePin: mustChangePin, viaMaster: false };
    }

    // 2) The master PIN, signing in AS this staff member (admin override).
    // The operator here is the admin holding the master PIN, not the
    // impersonated staff member, so the session always gets admin
    // privileges (add/edit/delete for anyone) regardless of that staff
    // member's own configured role.
    if (masterHash && pinHash === masterHash) {
      const token = signToken_({ staff: row[0], role: "admin", iat: Date.now(), viaMaster: true });
      return { ok: true, token: token, role: "admin", staff: row[0], mustChangePin: false, viaMaster: true };
    }

    return { ok: false, error: "Incorrect PIN." };
  }
  return { ok: false, error: "Unknown staff member. Ask an admin to set up your account." };
}

// Lets a signed-in person set their own PIN (used both for the forced
// "your PIN is temporary" flow and as an ordinary "Change PIN" option).
// A change made on a normal (non-master) session clears the temporary
// flag, since the person just chose it themselves. A change made on a
// master-PIN session (an admin resetting it on someone's behalf) leaves
// it flagged temporary, so that person is prompted to set their own PIN
// the next time they sign in for real.
function changePin_(newPin, auth) {
  const trimmed = String(newPin || "").trim();
  if (!/^\d{6}$/.test(trimmed)) {
    throw new Error("PIN must be exactly 6 digits.");
  }
  const sheet = getOrCreateSheet_(SHEET_AUTH, AUTH_HEADERS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (sameStaff_(values[i][0], auth.staff)) {
      sheet.getRange(i + 1, 2).setValue(sha256Hex_(trimmed));
      sheet.getRange(i + 1, 5).setValue(auth.viaMaster ? "TRUE" : "FALSE");
      return;
    }
  }
  throw new Error("Staff account not found.");
}

// One-time setup: run this manually from the Apps Script editor (select
// seedStaffAuth_ in the function dropdown, click Run). It creates the
// StaffAuth sheet if missing, and for every name that doesn't already
// have a row, generates a random 6-digit PIN and writes {Staff, PinHash,
// Role, GoogleEmail, PinIsTemporary}. New rows always start with
// PinIsTemporary = TRUE, forcing a PIN change on first sign-in. The
// plain-text PINs are only ever visible in the Execution log output of
// this one run — copy them from there and share each person their own
// PIN privately. Re-running is safe: it skips names that already have a
// row (so it won't reset anyone's PIN).
function seedStaffAuth_() {
  const sheet = getOrCreateSheet_(SHEET_AUTH, AUTH_HEADERS);
  const values = sheet.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) existing[String(values[i][0]).trim().toLowerCase()] = true;
  }

  const allNames = ALL_STAFF_NAMES.slice();
  ADMIN_STAFF_NAMES.forEach(function (n) {
    if (allNames.indexOf(n) === -1) allNames.push(n);
  });

  const newRows = [];
  const report = [];
  allNames.forEach(function (name) {
    const key = String(name).trim().toLowerCase();
    if (existing[key]) return;
    const pin = String(Math.floor(100000 + Math.random() * 900000));
    const role = ADMIN_STAFF_NAMES.some(function (a) { return sameStaff_(a, name); }) ? "admin" : "staff";
    newRows.push([name, sha256Hex_(pin), role, "", "TRUE"]);
    report.push(name + " (" + role + "): PIN " + pin);
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, AUTH_HEADERS.length).setValues(newRows);
  }

  Logger.log("=== New StaffAuth PINs (copy these now, they will not be shown again) ===");
  if (report.length) {
    report.forEach(function (line) { Logger.log(line); });
  } else {
    Logger.log("No new staff to add — everyone already has a StaffAuth row.");
  }
  Logger.log("=== end ===");
}

// One-time setup: run this manually from the Apps Script editor to
// generate the master PIN (see getMasterPinHash_() above). Only ever
// visible in the Execution log of this one run — copy it from there and
// give it only to the admin(s) who should have master access. Safe to
// re-run: it does nothing if a master PIN already exists (delete the
// MASTER_PIN_HASH script property first if you deliberately want to
// generate a new one).
function setMasterPin_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("MASTER_PIN_HASH")) {
    Logger.log("Master PIN already set — delete the MASTER_PIN_HASH script property first if you want a new one.");
    return;
  }
  const pin = String(Math.floor(100000 + Math.random() * 900000));
  props.setProperty("MASTER_PIN_HASH", sha256Hex_(pin));
  Logger.log("=== New master PIN (copy this now, it will not be shown again) ===");
  Logger.log("Master PIN: " + pin);
  Logger.log("=== end ===");
}

// ---------------------------------------------------------------------
// Sheets helpers
// ---------------------------------------------------------------------

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Robust "is this cell actually a date" check. A plain `instanceof Date`
// can fail to catch values coming back from Sheets in some execution
// contexts. Checking the internal [[Class]] via Object.prototype.toString
// is the reliable way to detect a Date value regardless of which realm
// it was constructed in.
function isDateValue_(v) {
  return v && Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime());
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const localTz = Session.getScriptTimeZone();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const isBlank = row.every(function (c) { return c === "" || c === null; });
    if (isBlank) continue;
    const obj = {};
    headers.forEach(function (h, idx) {
      let v = row[idx];
      const isDateOnly = DATE_ONLY_FIELDS.indexOf(h) !== -1;
      const isTimestamp = TIMESTAMP_FIELDS.indexOf(h) !== -1;
      if (isDateValue_(v)) {
        if (isTimestamp) {
          v = Utilities.formatDate(v, localTz, "yyyy-MM-dd HH:mm:ss");
        } else if (isDateOnly) {
          v = Utilities.formatDate(v, DATE_ONLY_TZ, "yyyy-MM-dd");
        } else {
          v = Utilities.formatDate(v, localTz, "yyyy-MM-dd");
        }
      } else if (typeof v === "string" && isDateOnly) {
        // Already a plain string (e.g. "2026-02-06") — leave as-is, but
        // strip a stray time-of-day/ISO suffix if one ever sneaks in.
        const m = v.match(/^(\d{4}-\d{2}-\d{2})T/);
        if (m) v = m[1];
      }
      obj[h] = v;
    });
    rows.push(obj);
  }
  return rows;
}

function addEntry_(entry, auth) {
  const sheet = getOrCreateSheet_(SHEET_ENTRIES, ENTRIES_HEADERS);
  const id = Utilities.getUuid();
  const rowIndex = sheet.getLastRow() + 1;
  const createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  // addEntry is admin-only (enforced by requireAdmin_ in doPost before
  // this is called), so the staff name posted by the admin is trusted.
  const staffName = entry.staff || "";
  const rowValues = [
    id,
    staffName,
    entry.startDate || "",
    entry.endDate || entry.startDate || "",
    entry.description || "",
    Number(entry.annualDays) || 0,
    Number(entry.sickDays) || 0,
    Number(entry.unpaidDays) || 0,
    Number(entry.hospitalizeDays) || 0,
    createdAt
  ];
  // Force the StartDate/EndDate/CreatedAt columns to stay plain TEXT, so
  // Sheets' automatic date-detection never turns them into a Date cell —
  // that auto-conversion is what caused the day-shift bug in the first
  // place. Setting the number format to "@" before writing the values
  // keeps them as the exact strings we send.
  sheet.getRange(rowIndex, 3, 1, 2).setNumberFormat("@");
  sheet.getRange(rowIndex, 10, 1, 1).setNumberFormat("@");
  sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
}

function deleteEntry_(id, auth) {
  // deleteEntry is admin-only (enforced by requireAdmin_ in doPost before
  // this is called).
  const sheet = getOrCreateSheet_(SHEET_ENTRIES, ENTRIES_HEADERS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

function upsertConfig_(cfg) {
  const sheet = getOrCreateSheet_(SHEET_STAFF_CONFIG, STAFF_CONFIG_HEADERS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === cfg.staff && String(values[i][1]) === String(cfg.year)) {
      sheet.getRange(i + 1, 1, 1, STAFF_CONFIG_HEADERS.length).setValues([[
        cfg.staff,
        cfg.year,
        Number(cfg.annualEntitlement) || 0,
        Number(cfg.annualCarryForward) || 0,
        Number(cfg.sickEntitlement) || 0,
        Number(cfg.sickCarryForward) || 0
      ]]);
      return;
    }
  }
  sheet.appendRow([
    cfg.staff,
    cfg.year,
    Number(cfg.annualEntitlement) || 0,
    Number(cfg.annualCarryForward) || 0,
    Number(cfg.sickEntitlement) || 0,
    Number(cfg.sickCarryForward) || 0
  ]);
}

function jsonResponse_(obj) {
  // Plain-text output (not setHeader'd as JSON with CORS headers — Apps
  // Script Web Apps handle CORS for GET/POST automatically when deployed
  // with "Anyone" access). The client also POSTs with a text/plain
  // Content-Type to avoid a CORS preflight that Apps Script can't answer.
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// One-time migration: run this manually, once, after pasting this
// updated script over an existing deployment that already has staff in
// StaffAuth from before the PinIsTemporary column existed. It:
//   1) Adds the PinIsTemporary header/column if the sheet predates it.
//   2) Flags every EXISTING row as PinIsTemporary = TRUE (forcing
//      everyone to set their own permanent PIN next time they sign in),
//      except any name listed in KEEP_CURRENT_PIN below, whose PIN stays
//      exactly as-is and is NOT flagged temporary.
//   3) Adds StaffAuth rows for any name in ALL_STAFF_NAMES that doesn't
//      have one yet (new joiners), same as seedStaffAuth_().
//   4) Generates the master PIN if one doesn't already exist.
// Safe to re-run: it only touches rows that still say PinIsTemporary is
// blank/unset for step 2, and only adds rows that don't exist for step 3.
function migrateAuthForPinOverhaul_() {
  const KEEP_CURRENT_PIN = ["Khairul"]; // already set their own permanent PIN

  const sheet = getOrCreateSheet_(SHEET_AUTH, AUTH_HEADERS);
  const range = sheet.getDataRange();
  const values = range.getValues();

  // Make sure the header row has all 5 columns (older sheets only had 4).
  if (values[0].length < AUTH_HEADERS.length || values[0][4] !== AUTH_HEADERS[4]) {
    sheet.getRange(1, 1, 1, AUTH_HEADERS.length).setValues([AUTH_HEADERS]);
  }

  const report = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue;
    const already = row[4];
    if (already === "TRUE" || already === "FALSE") continue; // already migrated
    const keep = KEEP_CURRENT_PIN.some(function (n) { return sameStaff_(n, row[0]); });
    sheet.getRange(i + 1, 5).setValue(keep ? "FALSE" : "TRUE");
    report.push(row[0] + ": PinIsTemporary = " + (keep ? "FALSE (kept current PIN)" : "TRUE"));
  }

  // Add any brand-new names.
  const existing = {};
  const values2 = sheet.getDataRange().getValues();
  for (let i = 1; i < values2.length; i++) {
    if (values2[i][0]) existing[String(values2[i][0]).trim().toLowerCase()] = true;
  }
  const newRows = [];
  ALL_STAFF_NAMES.forEach(function (name) {
    const key = String(name).trim().toLowerCase();
    if (existing[key]) return;
    const pin = String(Math.floor(100000 + Math.random() * 900000));
    const role = ADMIN_STAFF_NAMES.some(function (a) { return sameStaff_(a, name); }) ? "admin" : "staff";
    newRows.push([name, sha256Hex_(pin), role, "", "TRUE"]);
    report.push(name + " (NEW, " + role + "): PIN " + pin);
  });
  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, AUTH_HEADERS.length).setValues(newRows);
  }

  Logger.log("=== Auth migration report (copy any PINs shown now — new joiners only) ===");
  report.forEach(function (line) { Logger.log(line); });
  Logger.log("=== end ===");

  setMasterPin_();
}
