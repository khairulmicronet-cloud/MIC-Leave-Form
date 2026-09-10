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
// Everything happens client-side (via SheetJS) and the parsed entries are
// only written to the ledger once the admin reviews the preview and clicks
// "Import Checked Rows".
//
// This whole section is admin-only (hidden in the UI for regular staff by
// leave-breakdown.js, and refused server-side for anyone else), so the
// signed-in admin's token — set on window.MIC_AUTH by leave-breakdown.js
// after sign-in — is attached to every request here too.
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

// Pulls a date (or date range) out of a free-text description. Tries the
// patterns actually seen in the tracking sheet, in order of specificity.
function parseDateFromDescription(desc, fallbackYear) {
const m = desc.match(/\(([^)]+)\)\s*$/);
const inner = m ? m[1].trim() : desc.trim();

let mm;

// "29-08-2026 until 31-08-2026"
mm = inner.match(/(\d{1,2})-(\d{1,2})-(\d{4})\s+until\s+(\d{1,2})-(\d{1,2})-(\d{4})/i);
if (mm) return { start: iso(mm[3], mm[2], mm[1]), end: iso(mm[6], mm[5], mm[4]) };

// "17-08-2026 & 18-08-2026"
mm = inner.match(/(\d{1,2})-(\d{1,2})-(\d{4})\s*&\s*(\d{1,2})-(\d{1,2})-(\d{4})/);
if (mm) return { start: iso(mm[3], mm[2], mm[1]), end: iso(mm[6], mm[5], mm[4]) };

// "12-14 March 2026" / "11-12 Sept 2026"
mm = inner.match(/(\d{1,2})\s*-\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
if (mm) {
const mo = monthNum(mm[3]);
if (mo) return { start: iso(mm[4], mo, mm[1]), end: iso(mm[4], mo, mm[2]) };
}

// "7th February 2026" / "16-09-2026" style single ordinal date
mm = inner.match(/(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})/);
if (mm) {
const mo = monthNum(mm[2]);
if (mo) { const d = iso(mm[3], mo, mm[1]); return { start: d, end: d }; }
}

// "19-05-2026" single dd-mm-yyyy
mm = inner.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
if (mm) { const d = iso(mm[3], mm[2], mm[1]); return { start: d, end: d }; }

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
workbook = XLSX.read(buf, { type: "array" });

sheetSelect.innerHTML = "";
workbook.SheetNames.forEach(function (name) {
const opt = document.createElement("option");
opt.value = name;
opt.textContent = name;
sheetSelect.appendChild(opt);
});

staffSelect.innerHTML = "";
STAFF_NAMES.forEach(function (name) {
const opt = document.createElement("option");
opt.value = name;
opt.textContent = name;
staffSelect.appendChild(opt);
});

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
const desc = row[1];
if (!desc || typeof desc !== "string") continue;
const trimmed = desc.trim();
if (!trimmed) continue;
if (/^annual leave balance as of/i.test(trimmed)) continue;
// Month header rows are just a month name + 4-digit year, nothing else on the row.
const isMonthHeader = /^[A-Za-z]+\s+\d{4}$/.test(trimmed) &&
!row[2] && !row[3] && !row[4] && !row[5];
if (isMonthHeader) continue;

const annual = Number(row[2]) || 0;
const sick = Number(row[3]) || 0;
const unpaid = Number(row[4]) || 0;
const hosp = Number(row[5]) || 0;
if (!annual && !sick && !unpaid && !hosp) continue;

const dates = parseDateFromDescription(trimmed, year);
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

commitBtn.disabled = true;
commitStatus.className = "status-msg";
try {
if (setConfigChk.checked && parsedConfig) {
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

commitStatus.textContent = "Imported " + toImport.length + " entry(ies). Reloading ledger…";
commitStatus.className = "status-msg ok";
setTimeout(function () { window.location.reload(); }, 900);
} catch (err) {
commitStatus.textContent = "Import stopped: " + err.message + " (entries already imported before this stay saved — re-check the ledger before retrying).";
commitStatus.className = "status-msg error";
commitBtn.disabled = false;
}
});
})();
