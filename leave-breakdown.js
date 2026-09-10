// ---------------------------------------------------------------------
// Staff Leave Breakdown — reads/writes leave records through the Google
// Apps Script Web App configured in config.js (LEAVE_LEDGER_API_URL).
//
// Data model:
// StaffConfig: one row per (staff, year) holding Annual/Sick
// entitlement + carry-forward for that year.
// LeaveEntries: one row per leave event (a date range, a description,
// and however many Annual/Sick/Unpaid/Hospitalize days it used).
//
// Running balances are computed here in the browser (Annual/Sick balance
// = carry-forward + entitlement − sum of days used so far this year);
// Unpaid/Hospitalize have no entitlement, they're just running totals.
//
// Access control:
// Every request to the backend must carry a signed token obtained by
// signing in with a staff name + PIN (see login_() in the Apps Script
// backend). Admins (currently Khairul and Maziyah) see and edit every
// staff member's data; everyone else sees and edits only their own
// entries, and never sees the "Import from spreadsheet" or "Edit
// entitlement / carry-forward" sections at all (those actions are also
// blocked server-side, so hiding them here is a UX convenience, not the
// actual security boundary).
//
// The signed-in session (token, staff, role) is kept in localStorage so
// people don't have to re-enter their PIN every visit, and is also
// exposed on window.MIC_AUTH so leave-import.js (admin-only) can attach
// it to its own requests.
// ---------------------------------------------------------------------

(function () {
const API_URL = typeof LEAVE_LEDGER_API_URL !== "undefined" ? LEAVE_LEDGER_API_URL : "";
const STAFF_NAMES = (typeof LEAVE_FORM_STAFF_NAMES !== "undefined" && Array.isArray(LEAVE_FORM_STAFF_NAMES))
? LEAVE_FORM_STAFF_NAMES : [];
const AUTH_STORAGE_KEY = "micLeaveAuth";

window.MIC_AUTH = { token: null, staff: null, role: null };

const setupNotice = document.getElementById("setupNotice");
const loginBox = document.getElementById("loginBox");
const loginStaffSelect = document.getElementById("loginStaffSelect");
const loginPin = document.getElementById("loginPin");
const loginBtn = document.getElementById("loginBtn");
const loginStatus = document.getElementById("loginStatus");

const appEl = document.getElementById("app");
const sessionLabel = document.getElementById("sessionLabel");
const signOutBtn = document.getElementById("signOutBtn");
const staffSelect = document.getElementById("staffSelect");
const yearSelect = document.getElementById("yearSelect");
const ledgerSummary = document.getElementById("ledgerSummary");
const ledgerBody = document.getElementById("ledgerBody");
const entryForm = document.getElementById("entryForm");
const entryStatus = document.getElementById("entryStatus");
const configStatus = document.getElementById("configStatus");
const saveConfigBtn = document.getElementById("saveConfigBtn");
const importDetails = document.getElementById("importDetails");
const configDetails = document.getElementById("configDetails");

let state = { config: [], entries: [] };

function init() {
if (!API_URL) {
setupNotice.hidden = false;
loginBox.hidden = true;
appEl.hidden = true;
return;
}
setupNotice.hidden = true;

loginBtn.addEventListener("click", onLogin);
loginPin.addEventListener("keydown", function (ev) {
if (ev.key === "Enter") onLogin();
});
signOutBtn.addEventListener("click", onSignOut);

const saved = readSavedAuth();
if (saved) {
window.MIC_AUTH = saved;
startApp();
} else {
showLogin();
}
}

function readSavedAuth() {
try {
const raw = localStorage.getItem(AUTH_STORAGE_KEY);
if (!raw) return null;
const parsed = JSON.parse(raw);
if (parsed && parsed.token && parsed.staff && parsed.role) return parsed;
return null;
} catch (e) {
return null;
}
}

function saveAuth(auth) {
try {
localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(auth));
} catch (e) {
// localStorage unavailable (private browsing, etc.) — session just
// won't survive a page reload, which is fine.
}
}

function clearSavedAuth() {
try { localStorage.removeItem(AUTH_STORAGE_KEY); } catch (e) { /* ignore */ }
}

function showLogin() {
appEl.hidden = true;
loginBox.hidden = false;

if (!loginStaffSelect.options.length) {
STAFF_NAMES.forEach(function (name) {
const opt = document.createElement("option");
opt.value = name;
opt.textContent = name;
loginStaffSelect.appendChild(opt);
});
}
loginPin.value = "";
loginStatus.textContent = "";
loginStatus.className = "status-msg";
}

async function onLogin() {
const staff = loginStaffSelect.value;
const pin = loginPin.value.trim();
if (!staff || !pin) {
loginStatus.textContent = "Pick your name and enter your PIN.";
loginStatus.className = "status-msg error";
return;
}
loginBtn.disabled = true;
loginStatus.textContent = "Signing in…";
loginStatus.className = "status-msg";
try {
const resp = await fetch(API_URL, {
method: "POST",
headers: { "Content-Type": "text/plain;charset=utf-8" },
body: JSON.stringify({ action: "login", staff: staff, pin: pin })
});
const data = await resp.json();
if (!data.ok) throw new Error(data.error || "Sign in failed.");

const auth = { token: data.token, staff: data.staff, role: data.role };
window.MIC_AUTH = auth;
saveAuth(auth);
startApp();
} catch (err) {
loginStatus.textContent = "Error: " + err.message;
loginStatus.className = "status-msg error";
} finally {
loginBtn.disabled = false;
}
}

function onSignOut() {
clearSavedAuth();
window.MIC_AUTH = { token: null, staff: null, role: null };
window.location.reload();
}

function startApp() {
loginBox.hidden = true;
appEl.hidden = false;

const auth = window.MIC_AUTH;
const isAdmin = auth.role === "admin";

sessionLabel.textContent = "Signed in as " + auth.staff + (isAdmin ? " (admin)" : "");

// Admin-only sections: hidden entirely for regular staff. The backend
// also refuses these actions for non-admins, so this is a convenience,
// not the actual boundary.
if (importDetails) importDetails.hidden = !isAdmin;
if (configDetails) configDetails.hidden = !isAdmin;

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
opt.value = auth.staff;
opt.textContent = auth.staff;
staffSelect.appendChild(opt);
staffSelect.value = auth.staff;
staffSelect.disabled = true;
}

if (!yearSelect.options.length) {
const currentYear = new Date().getFullYear();
for (let y = currentYear - 1; y <= currentYear + 1; y++) {
const opt = document.createElement("option");
opt.value = String(y);
opt.textContent = String(y);
if (y === currentYear) opt.selected = true;
yearSelect.appendChild(opt);
}
}

staffSelect.addEventListener("change", render);
yearSelect.addEventListener("change", render);
entryForm.addEventListener("submit", onAddEntry);
saveConfigBtn.addEventListener("click", onSaveConfig);

loadData();
}

async function loadData() {
ledgerSummary.textContent = "Loading…";
try {
const url = API_URL + "?token=" + encodeURIComponent(window.MIC_AUTH.token || "");
const resp = await fetch(url, { cache: "no-store" });
const data = await resp.json();
if (!data.ok) throw new Error(data.error || "Failed to load data");
state.config = data.config || [];
state.entries = (data.entries || []).map(function (e) {
return Object.assign({}, e, {
StartDate: normalizeDate(e.StartDate),
EndDate: normalizeDate(e.EndDate)
});
});
render();
} catch (err) {
if (isAuthError(err)) { handleAuthError(); return; }
ledgerSummary.textContent = "Could not load data: " + err.message;
}
}

function isAuthError(err) {
return /not signed in/i.test(err.message || "");
}

function handleAuthError() {
clearSavedAuth();
window.MIC_AUTH = { token: null, staff: null, role: null };
showLogin();
loginStatus.textContent = "Your session expired — please sign in again.";
loginStatus.className = "status-msg error";
}

// The Apps Script backend is supposed to hand back plain "yyyy-MM-dd"
// strings, but depending on which script version is actually deployed,
// Google Sheets' automatic date-detection can turn a stored date string
// into a real Date cell, which then comes back as a full ISO timestamp
// (e.g. "2026-02-06T16:00:00.000Z") — a calendar day off from what was
// entered once you're east of UTC. Normalize defensively here so the
// ledger always displays the correct calendar date regardless of what
// the backend currently returns.
function normalizeDate(v) {
if (!v) return v;
if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
const d = new Date(v);
if (isNaN(d.getTime())) return v;
try {
return new Intl.DateTimeFormat("en-CA", {
timeZone: "Asia/Brunei", year: "numeric", month: "2-digit", day: "2-digit"
}).format(d);
} catch (e) {
return v;
}
}

function currentStaff() { return staffSelect.value; }
function currentYear() { return yearSelect.value; }

function getConfigFor(staff, year) {
return state.config.find(function (c) {
return c.Staff === staff && String(c.Year) === String(year);
}) || { Staff: staff, Year: year, AnnualEntitlement: 0, AnnualCarryForward: 0, SickEntitlement: 0, SickCarryForward: 0 };
}

function entriesFor(staff, year) {
return state.entries
.filter(function (e) {
return e.Staff === staff && (e.StartDate || "").slice(0, 4) === String(year);
})
.sort(function (a, b) { return (a.StartDate || "").localeCompare(b.StartDate || ""); });
}

function render() {
const staff = currentStaff();
const year = currentYear();
if (!staff || !year) return;

const cfg = getConfigFor(staff, year);
const entries = entriesFor(staff, year);

document.getElementById("cfgAnnualEntitlement").value = cfg.AnnualEntitlement || 0;
document.getElementById("cfgAnnualCarryForward").value = cfg.AnnualCarryForward || 0;
document.getElementById("cfgSickEntitlement").value = cfg.SickEntitlement || 0;
document.getElementById("cfgSickCarryForward").value = cfg.SickCarryForward || 0;

let annualBal = Number(cfg.AnnualCarryForward || 0) + Number(cfg.AnnualEntitlement || 0);
let sickBal = Number(cfg.SickCarryForward || 0) + Number(cfg.SickEntitlement || 0);
let unpaidTotal = 0;
let hospitalizeTotal = 0;

ledgerBody.innerHTML = "";
entries.forEach(function (e) {
annualBal -= Number(e.AnnualDays || 0);
sickBal -= Number(e.SickDays || 0);
unpaidTotal += Number(e.UnpaidDays || 0);
hospitalizeTotal += Number(e.HospitalizeDays || 0);

const tr = document.createElement("tr");
const dateLabel = e.StartDate === e.EndDate || !e.EndDate ? e.StartDate : (e.StartDate + " to " + e.EndDate);
tr.innerHTML =
"<td>" + escapeHtml(dateLabel || "") + "</td>" +
"<td>" + escapeHtml(e.Description || "") + "</td>" +
"<td>" + fmtDays(e.AnnualDays) + "</td>" +
"<td>" + fmtDays(e.SickDays) + "</td>" +
"<td>" + fmtDays(e.UnpaidDays) + "</td>" +
"<td>" + fmtDays(e.HospitalizeDays) + "</td>" +
"<td>" + annualBal + "</td>" +
"<td>" + sickBal + "</td>" +
'<td><button type="button" class="btn-link delete-entry" data-id="' + escapeHtml(e.ID) + '">Delete</button></td>';
ledgerBody.appendChild(tr);
});

ledgerBody.querySelectorAll(".delete-entry").forEach(function (btn) {
btn.addEventListener("click", function () { onDeleteEntry(btn.getAttribute("data-id")); });
});

ledgerSummary.innerHTML =
'<div class="admin-preview">' +
"<div class=\"admin-row\"><strong>Annual Leave Balance</strong> &nbsp; " + annualBal + " day(s)</div>" +
"<div class=\"admin-row\"><strong>Sick Leave Balance</strong> &nbsp; " + sickBal + " day(s)</div>" +
"<div class=\"admin-row\"><strong>Unpaid Leave Taken (this year)</strong> &nbsp; " + unpaidTotal + " day(s)</div>" +
"<div class=\"admin-row\"><strong>Hospitalize Leave Taken (this year)</strong> &nbsp; " + hospitalizeTotal + " day(s)</div>" +
"</div>";
}

function fmtDays(v) {
const n = Number(v || 0);
return n === 0 ? "" : String(n);
}

function escapeHtml(str) {
return String(str)
.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
.replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

async function postAction(payload) {
const withToken = Object.assign({ token: window.MIC_AUTH.token }, payload);
const resp = await fetch(API_URL, {
method: "POST",
// text/plain avoids a CORS preflight that Apps Script Web Apps can't answer;
// the script itself still JSON.parses the body regardless of this header.
headers: { "Content-Type": "text/plain;charset=utf-8" },
body: JSON.stringify(withToken)
});
const data = await resp.json();
if (!data.ok) throw new Error(data.error || "Request failed");
return data;
}

async function onAddEntry(ev) {
ev.preventDefault();
entryStatus.textContent = "Saving…";
entryStatus.className = "status-msg";
try {
const start = document.getElementById("entryStart").value;
if (!start) throw new Error("Start date is required.");
const end = document.getElementById("entryEnd").value || start;
const desc = document.getElementById("entryDesc").value.trim();
if (!desc) throw new Error("Description is required.");

await postAction({
action: "addEntry",
entry: {
staff: currentStaff(),
startDate: start,
endDate: end,
description: desc,
annualDays: document.getElementById("entryAnnual").value || 0,
sickDays: document.getElementById("entrySick").value || 0,
unpaidDays: document.getElementById("entryUnpaid").value || 0,
hospitalizeDays: document.getElementById("entryHospitalize").value || 0
}
});

entryForm.reset();
entryStatus.textContent = "Entry added.";
entryStatus.className = "status-msg ok";
await loadData();
} catch (err) {
if (isAuthError(err)) { handleAuthError(); return; }
entryStatus.textContent = "Error: " + err.message;
entryStatus.className = "status-msg error";
}
}

async function onDeleteEntry(id) {
if (!confirm("Delete this leave entry?")) return;
try {
await postAction({ action: "deleteEntry", id: id });
await loadData();
} catch (err) {
if (isAuthError(err)) { handleAuthError(); return; }
alert("Could not delete: " + err.message);
}
}

async function onSaveConfig() {
configStatus.textContent = "Saving…";
configStatus.className = "status-msg";
try {
await postAction({
action: "upsertConfig",
config: {
staff: currentStaff(),
year: currentYear(),
annualEntitlement: document.getElementById("cfgAnnualEntitlement").value || 0,
annualCarryForward: document.getElementById("cfgAnnualCarryForward").value || 0,
sickEntitlement: document.getElementById("cfgSickEntitlement").value || 0,
sickCarryForward: document.getElementById("cfgSickCarryForward").value || 0
}
});
configStatus.textContent = "Saved.";
configStatus.className = "status-msg ok";
await loadData();
} catch (err) {
if (isAuthError(err)) { handleAuthError(); return; }
configStatus.textContent = "Error: " + err.message;
configStatus.className = "status-msg error";
}
}

init();
})();
