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

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const isBlank = row.every(function (c) { return c === "" || c === null; });
    if (isBlank) continue;
    const obj = {};
    headers.forEach(function (h, idx) {
      let v = row[idx];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
      obj[h] = v;
    });
    rows.push(obj);
  }
  return rows;
}

function addEntry_(entry) {
  const sheet = getOrCreateSheet_(SHEET_ENTRIES, ENTRIES_HEADERS);
  const id = Utilities.getUuid();
  sheet.appendRow([
    id,
    entry.staff || "",
    entry.startDate || "",
    entry.endDate || entry.startDate || "",
    entry.description || "",
    Number(entry.annualDays) || 0,
    Number(entry.sickDays) || 0,
    Number(entry.unpaidDays) || 0,
    Number(entry.hospitalizeDays) || 0,
    new Date()
  ]);
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
