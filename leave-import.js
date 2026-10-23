// ---------------------------------------------------------------------
// Staff Leave Breakdown — bulk import from a spreadsheet.
//
// Reads the same "one sheet per staff member" layout used by the
// Staff_Leave_Breakdown tracking workbook:
// Row 1: [Name, , "Annual Leave", "Sick Leave", "Unpaid Leave", "Hospitalize Leave"]
// Row 2: ["... brought forward from YYYY", , annualCarryForward, sickCarryForward]
// Row 3: ["... Entitlement for YYYY", , annualEntitlement, sickEntitlement]
// Then repeating month blocks: a month header row, "Annual leave balance
// as of ..." rows (skipped — these are running totals, not entries),
// and the actual leave-entry rows in between.
//
// Column B (the date/description cell) accepts any of these, so staff can
// keep whatever style their own tracking sheet already uses:
// - A real Excel date cell with no text at all (just the date value) —
// shown in the preview as "Leave (D Month YYYY)".
// - Free text with the date (or date range) embedded, in any of: "D Month
// YYYY", "Month D, YYYY", "DD-MM-YYYY", "DD/MM/YYYY", "YYYY-MM-DD", a
// range joined by "-", "to", "until" or "&" between two such dates, or
// "D-D Month YYYY" for a same-month range — with or without a leading
// label like "Leave (...)" or "MC (...)".
// - Anything else that still looks date-shaped (contains a 4-digit year)
// falls back to the browser's own date parser as a last resort.
// A row whose date genuinely can't be determined still appears in the
// preview with a blank, red-bordered date field for the admin to fill in
// by hand — it is never silently dropped.
//
// Everything happens client-side (via SheetJS) and the parsed entries are
// only written to the ledger once the admin reviews the preview and clicks
// "Import Checked Rows".
//
// This section is open to every signed-in staff member. An admin can
// import as any staff member and set entitlement/carry-forward from the
// spreadsheet (unchanged); a regular staff member can only import as
// themselves, can't touch entitlement/carry-forward, and their rows go
// to an admin for review (addEntry_ on the backend routes a non-admin's
// entry into PendingEntries instead of the live ledger) rather than
// landing on the ledger immediately. Either way the signed-in person's
// token — set on window.MIC_AUTH by leave-breakdown.js after sign-in —
// is attached to every request here.
// ---------------------------------------------------------------------

(function () {
const API_URL = typeof LEAVE_LEDGER_API_URL !== "undefined" ? LEAVE_LEDGER_API_URL : "";
const STAFF_NAMES = (typeof LEAVE_FORM_STAFF_NAMES !== "undefined" && Array.isArray(LEAVE_FORM_STAFF_NAMES))
? LEAVE_FORM_STAFF_NAMES : [];

const fileInput = document.getElementById("importFile");
const sheetRow = document.getElementById("importSheetRow");
const sheetSelect = document.getElementById("importSheetSelect");
const staffRow = document.getElementById("importStaffRow");
const staffSelect = document.getElementById("importStaffSelect");
const yearRow = document.getElementById("importYearRow");
const yearInput = document.getElementById("importYearInput");
const parseRow = document.getElementById("importParseRow");
const parseBtn = document.getElementById("importParseBtn");
const status = document.getElementById("importStatus");
const previewWrap = document.getElementById("importPreviewWrap");
const previewBody = document.getElementById("importPreviewBody");
const setConfigChk = document.getElementById("importSetConfig");
const setConfigRow = document.getElementById("importSetConfigRow");
const commitBtn = document.getElementById("importCommitBtn");
const commitStatus = document.getElementById("importCommitStatus");

if (!fileInput || typeof XLSX === "undefined") return;

let workbook = null;
let parsedConfig = null;
let previewRows = []; // { include, start, end, desc, annual, sick, unpaid, hosp }

const MONTHS = {
jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4,
may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8,
sep: 9, sept: 9, september: 9, oct: 10, october: 10, nov: 11, november: 11,
dec: 12, december: 12
};

function pad(n) { return String(n).padStart(2, "0"); }
function iso(y, m, d) { return y + "-" + pad(m) + "-" + pad(d); }
function monthNum(name) { return MONTHS[String(name).trim().toLowerCase()] || null; }

// Converts a real Excel date serial number (the 1899-12-30 epoch) to a JS
// Date. Guards against small numbers (day counts, leave totals, etc.) being
// mistaken for dates — real calendar dates in this tracker are all comfortably
// above 20000 (year ~1954).
function excelSerialToDate(serial) {
if (typeof serial !== "number" || !isFinite(serial) || serial < 20000) return null;
const ms = Math.round((serial - 25569) * 86400 * 1000);
const d = new Date(ms);
return isNaN(d.getTime()) ? null : d;
}

// Timezone-safe Date -> "YYYY-MM-DD". SheetJS builds date cells at UTC
// midnight, so local getters can shift the day by one depending on the
// browser's timezone — always read back with the UTC getters.
function dateToISO(d) {
return iso(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
}

const MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
"July", "August", "September", "October", "November", "December"];
function formatDisplayDate(d) {
return d.getUTCDate() + " " + MONTH_NAMES[d.getUTCMonth()] + " " + d.getUTCFullYear();
}

// Pulls a date (or date range) out of a free-text description. Tries the
// patterns actually seen in the tracking sheet, from most to least specific,
// then falls back to the browser's own date parser for anything else that
// still looks date-shaped. Returns { start: "", end: "" } only when nothing
// at all could be made of the text, so the preview row still shows up with
// a blank, editable date for manual fixing.
function parseDateFromDescription(desc, fallbackYear) {
const m = desc.match(/\(([^)]+)\)\s*$/);
const inner = m ? m[1].trim() : desc.trim();

let mm;

// "29-08-2026 until 31-08-2026" / "29-08-2026 to 31-08-2026" / "29-08-2026 - 31-08-2026" / "17-08-2026 & 18-08-2026"
mm = inner.match(/(\d{1,2})-(\d{1,2})-(\d{4})\s*(?:-|until|to|&)\s*(\d{1,2})-(\d{1,2})-(\d{4})/i);
if (mm) return { start: iso(mm[3], mm[2], mm[1]), end: iso(mm[6], mm[5], mm[4]) };

// "5 June 2026 - 6 June 2026" / "5 June 2026 to 6 June 2026" / "...until..." / "...&..."
mm = inner.match(/(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})\s*(?:-|to|until|&)\s*(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})/i);
if (mm) {
const mo1 = monthNum(mm[2]), mo2 = monthNum(mm[5]);
if (mo1 && mo2) return { start: iso(mm[3], mo1, mm[1]), end: iso(mm[6], mo2, mm[4]) };
}

// "12-14 March 2026" / "11-12 Sept 2026"
mm = inner.match(/(\d{1,2})\s*-\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
if (mm) {
const mo = monthNum(mm[3]);
if (mo) return { start: iso(mm[4], mo, mm[1]), end: iso(mm[4], mo, mm[2]) };
}

// "7th February 2026" / single ordinal date, day before month
mm = inner.match(/(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})/);
if (mm) {
const mo = monthNum(mm[2]);
if (mo) { const d = iso(mm[3], mo, mm[1]); return { start: d, end: d }; }
}

// "June 5, 2026 - June 6, 2026" / "Feb 7 2026 to Feb 9 2026" — month-first range
mm = inner.match(/([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\s*(?:-|to|until|&)\s*([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})/i);
if (mm) {
const mo1 = monthNum(mm[1]), mo2 = monthNum(mm[4]);
if (mo1 && mo2) return { start: iso(mm[3], mo1, mm[2]), end: iso(mm[6], mo2, mm[5]) };
}

// "February 7, 2026" / "Feb 7 2026" — month name first, single date
mm = inner.match(/([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})/);
if (mm) {
const mo = monthNum(mm[1]);
if (mo) { const d = iso(mm[3], mo, mm[2]); return { start: d, end: d }; }
}

// ISO "2026-06-05 to 2026-06-06" / "2026-06-05"
mm = inner.match(/(\d{4})-(\d{1,2})-(\d{1,2})\s*(?:-|to|until|&)\s*(\d{4})-(\d{1,2})-(\d{1,2})/i);
if (mm) return { start: iso(mm[1], mm[2], mm[3]), end: iso(mm[4], mm[5], mm[6]) };
mm = inner.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
if (mm) { const d = iso(mm[1], mm[2], mm[3]); return { start: d, end: d }; }

// "19-05-2026" single dd-mm-yyyy
mm = inner.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
if (mm) { const d = iso(mm[3], mm[2], mm[1]); return { start: d, end: d }; }

// "19/05/2026 - 21/05/2026" / "19/05/2026" slash dates (dd/mm/yyyy)
mm = inner.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s*(?:-|to|until|&)\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/i);
if (mm) return { start: iso(mm[3], mm[2], mm[1]), end: iso(mm[6], mm[5], mm[4]) };
mm = inner.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
if (mm) { const d = iso(mm[3], mm[2], mm[1]); return { start: d, end: d }; }

// Last resort: let the browser's own date parser take a shot — catches odd
// but valid styles not covered above (e.g. "5 Jun, 2026", "06/05/2026 (Fri)").
// Only trusted when the text actually contains a 4-digit year, so a stray
// number in the description is never mistaken for a date.
if (/\d{4}/.test(inner)) {
const native = new Date(inner);
if (!isNaN(native.getTime())) {
const d = dateToISO(native);
return { start: d, end: d };
}
}

return { start: "", end: "" };
}

function resetPreview() {
previewRows = [];
previewWrap.hidden = true;
previewBody.innerHTML = "";
commitStatus.textContent = "";
}

fileInput.addEventListener("change", async function () {
status.textContent = "";
resetPreview();
const file = fileInput.files[0];
if (!file) return;
try {
const buf = await file.arrayBuffer();
// cellDates: true makes SheetJS hand back real Excel date cells as JS Date
// objects instead of raw numeric serials, so a sheet that stores the leave
// date as an actual date cell (not typed into the description) parses
// correctly below instead of being silently skipped.
workbook = XLSX.read(buf, { type: "array", cellDates: true });

sheetSelect.innerHTML = "";
workbook.SheetNames.forEach(function (name) {
const opt = document.createElement("option");
opt.value = name;
opt.textContent = name;
sheetSelect.appendChild(opt);
});

const isAdmin = window.MIC_AUTH && window.MIC_AUTH.role === "admin";

staffSelect.innerHTML = "";
if (isAdmin) {
STAFF_NAMES.forEach(function (name) {
const opt = document.createElement("option");
opt.value = name;
opt.textContent = name;
staffSelect.appendChild(opt);
});
staffSelect.disabled = false;
} else {
const opt = document.createElement("option");
opt.value = window.MIC_AUTH.staff;
opt.textContent = window.MIC_AUTH.staff;
staffSelect.appendChild(opt);
staffSelect.value = window.MIC_AUTH.staff;
staffSelect.disabled = true;
}

// Entitlement / carry-forward stays admin-only, same as the manual "Add
// a Leave Entry" form and the Edit Entitlement section.
if (setConfigRow) setConfigRow.hidden = !isAdmin;
if (!isAdmin) setConfigChk.checked = false;

yearInput.value = new Date().getFullYear();

sheetRow.hidden = false;
staffRow.hidden = false;
yearRow.hidden = false;
parseRow.hidden = false;
} catch (err) {
status.textContent = "Could not read that file: " + err.message;
status.className = "status-msg error";
}
});

parseBtn.addEventListener("click", function () {
status.textContent = "";
resetPreview();
try {
const sheet = workbook.Sheets[sheetSelect.value];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
const year = yearInput.value || new Date().getFullYear();

parsedConfig = {
annualCarryForward: Number((rows[1] || [])[2]) || 0,
sickCarryForward: Number((rows[1] || [])[3]) || 0,
annualEntitlement: Number((rows[2] || [])[2]) || 0,
sickEntitlement: Number((rows[2] || [])[3]) || 0
};

for (let i = 3; i < rows.length; i++) {
const row = rows[i] || [];
const raw = row[1];
if (raw === null || raw === undefined || raw === "") continue;

let trimmed, dates;

if (raw instanceof Date) {
// A real Excel date cell with no description text alongside it — the
// case that used to be silently dropped. cellDates:true (above) gives
// us this as a proper Date.
const d = dateToISO(raw);
dates = { start: d, end: d };
trimmed = "Leave (" + formatDisplayDate(raw) + ")";
} else if (typeof raw === "number") {
// Some spreadsheets (or a raw:true re-read) can still hand back a bare
// Excel date serial instead of a Date object — recover it the same way.
const asDate = excelSerialToDate(raw);
if (!asDate) continue;
const d = dateToISO(asDate);
dates = { start: d, end: d };
trimmed = "Leave (" + formatDisplayDate(asDate) + ")";
} else if (typeof raw === "string") {
trimmed = raw.trim();
if (!trimmed) continue;
if (/^annual leave balance as of/i.test(trimmed)) continue;
// Month header rows are just a month name + 4-digit year, nothing else on the row.
const isMonthHeader = /^[A-Za-z]+\s+\d{4}$/.test(trimmed) &&
!row[2] && !row[3] && !row[4] && !row[5];
if (isMonthHeader) continue;
dates = parseDateFromDescription(trimmed, year);
} else {
continue;
}

const annual = Number(row[2]) || 0;
const sick = Number(row[3]) || 0;
const unpaid = Number(row[4]) || 0;
const hosp = Number(row[5]) || 0;
if (!annual && !sick && !unpaid && !hosp) continue;

previewRows.push({
include: true,
start: dates.start,
end: dates.end || dates.start,
desc: trimmed,
annual: annual, sick: sick, unpaid: unpaid, hosp: hosp
});
}

if (previewRows.length === 0) {
status.textContent = "No leave-entry rows found on that sheet. Check you picked the right tab.";
status.className = "status-msg error";
return;
}

renderPreview();
previewWrap.hidden = false;
status.textContent = previewRows.length + " row(s) found. Review dates before importing.";
status.className = "status-msg ok";
} catch (err) {
status.textContent = "Could not parse that sheet: " + err.message;
status.className = "status-msg error";
}
});

function renderPreview() {
previewBody.innerHTML = "";
previewRows.forEach(function (r, idx) {
const tr = document.createElement("tr");

const chkTd = document.createElement("td");
const chk = document.createElement("input");
chk.type = "checkbox";
chk.checked = r.include;
chk.addEventListener("change", function () { previewRows[idx].include = chk.checked; });
chkTd.appendChild(chk);

const startTd = document.createElement("td");
const startInput = document.createElement("input");
startInput.type = "date";
startInput.value = r.start;
if (!r.start) startInput.style.borderColor = "#c0392b";
startInput.addEventListener("input", function () {
previewRows[idx].start = startInput.value;
startInput.style.borderColor = startInput.value ? "" : "#c0392b";
});
startTd.appendChild(startInput);

const endTd = document.createElement("td");
const endInput = document.createElement("input");
endInput.type = "date";
endInput.value = r.end;
endInput.addEventListener("input", function () { previewRows[idx].end = endInput.value; });
endTd.appendChild(endInput);

const descTd = document.createElement("td");
descTd.textContent = r.desc;

const annualTd = document.createElement("td"); annualTd.textContent = r.annual || "";
const sickTd = document.createElement("td"); sickTd.textContent = r.sick || "";
const unpaidTd = document.createElement("td"); unpaidTd.textContent = r.unpaid || "";
const hospTd = document.createElement("td"); hospTd.textContent = r.hosp || "";

tr.appendChild(chkTd);
tr.appendChild(startTd);
tr.appendChild(endTd);
tr.appendChild(descTd);
tr.appendChild(annualTd);
tr.appendChild(sickTd);
tr.appendChild(unpaidTd);
tr.appendChild(hospTd);
previewBody.appendChild(tr);
});
}

async function postAction(payload) {
const token = window.MIC_AUTH && window.MIC_AUTH.token;
const withToken = Object.assign({ token: token }, payload);
const resp = await fetch(API_URL, {
method: "POST",
headers: { "Content-Type": "text/plain;charset=utf-8" },
body: JSON.stringify(withToken)
});
const data = await resp.json();
if (!data.ok) throw new Error(data.error || "Request failed");
return data;
}

commitBtn.addEventListener("click", async function () {
const staff = staffSelect.value;
const year = yearInput.value || new Date().getFullYear();
const toImport = previewRows.filter(function (r) { return r.include; });

if (!staff) { commitStatus.textContent = "Pick a staff member first."; commitStatus.className = "status-msg error"; return; }
const missingDate = toImport.find(function (r) { return !r.start; });
if (missingDate) {
commitStatus.textContent = "Every checked row needs a start date — fill in the row highlighted in red.";
commitStatus.className = "status-msg error";
return;
}
if (toImport.length === 0) {
commitStatus.textContent = "Nothing checked to import.";
commitStatus.className = "status-msg error";
return;
}

const isAdmin = window.MIC_AUTH && window.MIC_AUTH.role === "admin";

commitBtn.disabled = true;
commitStatus.className = "status-msg";
try {
if (isAdmin && setConfigChk.checked && parsedConfig) {
commitStatus.textContent = "Saving entitlement…";
await postAction({
action: "upsertConfig",
config: {
staff: staff,
year: year,
annualEntitlement: parsedConfig.annualEntitlement,
annualCarryForward: parsedConfig.annualCarryForward,
sickEntitlement: parsedConfig.sickEntitlement,
sickCarryForward: parsedConfig.sickCarryForward
}
});
}

for (let i = 0; i < toImport.length; i++) {
const r = toImport[i];
commitStatus.textContent = "Importing " + (i + 1) + " of " + toImport.length + "…";
await postAction({
action: "addEntry",
entry: {
staff: staff,
startDate: r.start,
endDate: r.end || r.start,
description: r.desc,
annualDays: r.annual,
sickDays: r.sick,
unpaidDays: r.unpaid,
hospitalizeDays: r.hosp
}
});
}

commitStatus.textContent = isAdmin
? "Imported " + toImport.length + " entry(ies). Reloading ledger…"
: "Submitted " + toImport.length + " entry(ies) for admin review. They'll appear on your ledger once approved. Reloading…";
commitStatus.className = "status-msg ok";
setTimeout(function () { window.location.reload(); }, 900);
} catch (err) {
commitStatus.textContent = "Import stopped: " + err.message + " (entries already imported before this stay saved — re-check the ledger before retrying).";
commitStatus.className = "status-msg error";
commitBtn.disabled = false;
}
});
})();
