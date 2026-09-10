/**
 * Staff Leave Breakdown — Google Apps Script backend.
 *
 * Paste this whole file into Extensions > Apps Script inside your Google
 * Sheet (see SETUP-STAFF-LEAVE-BREAKDOWN.md for the full walkthrough),
 * then deploy it as a Web App. The deployed URL is what you paste into
 * config.js as LEAVE_LEDGER_API_URL.
 *
 * Expects two sheets/tabs in the spreadsheet, created automatically on
 * first run if missing:
 *
 *   StaffConfig: Staff | Year | AnnualEntitlement | AnnualCarryForward | SickEntitlement | SickCarryForward
 *   LeaveEntries: ID | Staff | StartDate | EndDate | Description | AnnualDays | SickDays | UnpaidDays | HospitalizeDays | CreatedAt
 */

const SHEET_STAFF_CONFIG = "StaffConfig";
const SHEET_ENTRIES = "LeaveEntries";

const STAFF_CONFIG_HEADERS = [
  "Staff", "Year", "AnnualEntitlement", "AnnualCarryForward", "SickEntitlement", "SickCarryForward"
];
const ENTRIES_HEADERS = [
  "ID", "Staff", "StartDate", "EndDate", "Description",
  "AnnualDays", "SickDays", "UnpaidDays", "HospitalizeDays", "CreatedAt"
];

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

function doGet(e) {
  try {
    const configSheet = getOrCreateSheet_(SHEET_STAFF_CONFIG, STAFF_CONFIG_HEADERS);
    const entriesSheet = getOrCreateSheet_(SHEET_ENTRIES, ENTRIES_HEADERS);
    const config = sheetToObjects_(configSheet);
    const entries = sheetToObjects_(entriesSheet);
    return jsonResponse_({ ok: true, config: config, entries: entries });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === "addEntry") {
      addEntry_(body.entry);
    } else if (action === "upsertConfig") {
      upsertConfig_(body.config);
    } else if (action === "deleteEntry") {
      deleteEntry_(body.id);
    } else {
      return jsonResponse_({ ok: false, error: "Unknown action: " + action });
    }

    return jsonResponse_({ ok: true });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

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

function addEntry_(entry) {
  const sheet = getOrCreateSheet_(SHEET_ENTRIES, ENTRIES_HEADERS);
  const id = Utilities.getUuid();
  const rowIndex = sheet.getLastRow() + 1;
  const createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const rowValues = [
    id,
    entry.staff || "",
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

function deleteEntry_(id) {
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
