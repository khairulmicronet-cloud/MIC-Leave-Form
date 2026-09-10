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
 *   StaffAuth: Staff | PinHash | Role | GoogleEmail
 *
 * Access control:
 *   Every request must carry a signed token (obtained via the "login"
 *   action, staff name + PIN). A token proves who the caller is and
 *   whether they're "admin" or plain "staff":
 *     - admin: sees and edits every staff member's data (same as before
 *       this feature existed).
 *     - staff: sees ONLY their own entries/config. addEntry is forced to
 *       their own name regardless of what's posted. deleteEntry is only
 *       allowed on their own entries. upsertConfig (entitlement /
 *       carry-forward) and the spreadsheet import are admin-only.
 *
 *   Run seedStaffAuth_() once (from the Apps Script editor's Run button,
 *   not via the web app) to create the StaffAuth sheet and generate a
 *   random PIN for every name in LEAVE_FORM_STAFF_NAMES, plus whichever
 *   admin names you list in ADMIN_STAFF_NAMES below. The generated PINs
 *   are written to the Execution log — copy them from there, they are
 *   NOT stored anywhere in plain text after that (only a hash is kept).
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
const AUTH_HEADERS = ["Staff", "PinHash", "Role", "GoogleEmail"];

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
  "Mustadim"
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
      addEntry_(body.entry, auth);
    } else if (action === "upsertConfig") {
      requireAdmin_(auth);
      upsertConfig_(body.config);
    } else if (action === "deleteEntry") {
      deleteEntry_(body.id, auth);
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
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (sameStaff_(row[0], staff)) {
      if (String(row[1]) === pinHash) {
        const role = String(row[2] || "staff").trim().toLowerCase() === "admin" ? "admin" : "staff";
        const token = signToken_({ staff: row[0], role: role, iat: Date.now() });
        return { ok: true, token: token, role: role, staff: row[0] };
      }
      return { ok: false, error: "Incorrect PIN." };
    }
  }
  return { ok: false, error: "Unknown staff member. Ask an admin to set up your account." };
}

// One-time setup: run this manually from the Apps Script editor (select
// seedStaffAuth_ in the function dropdown, click Run). It creates the
// StaffAuth sheet if missing, and for every name that doesn't already
// have a row, generates a random 6-digit PIN and writes {Staff, PinHash,
// Role, GoogleEmail}. The plain-text PINs are only ever visible in the
// Execution log output of this one run — copy them from there and share
// each person their own PIN privately. Re-running is safe: it skips
// names that already have a row (so it won't reset anyone's PIN).
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
    newRows.push([name, sha256Hex_(pin), role, ""]);
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
  // A non-admin can only ever add an entry for themselves, no matter what
  // staff name was posted.
  const staffName = auth.role === "admin" ? (entry.staff || "") : auth.staff;
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
  const sheet = getOrCreateSheet_(SHEET_ENTRIES, ENTRIES_HEADERS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      if (auth.role !== "admin" && !sameStaff_(values[i][1], auth.staff)) {
        throw new Error("You can only delete your own entries.");
      }
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
