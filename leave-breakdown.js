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
// Running balances are computed here in the browser. Sick balance is a
// flat pool (carry-forward + entitlement − used). Annual balance accrues
// monthly instead of being available in full on 1 January: 1 day is
// released on the 1st of each month (Jan–Oct), 2 days on 1 Nov, 2 days
// on 1 Dec, plus 1 more on 31 Dec — 15 "shares" total for a staff member
// on the default 15-day entitlement. A different configured entitlement
// scales every release proportionally (see accruedAnnual() below).
// Unpaid/Hospitalize have no entitlement, they're just running totals.
//
// Access control:
// Every request to the backend must carry a signed token obtained by
// signing in with a staff name + PIN (see login_() in the Apps Script
// backend). Admins (currently Khairul and Maziyah) see and edit every
// staff member's data. Everyone else gets a read-only view of their own
// record: the "Add a Leave Entry" quick-add form and editing
// entitlement/carry-forward stay admin-only (so only admins tally leave
// directly), and those sections are hidden entirely for non-admins
// (also blocked server-side, so hiding them here is a UX convenience,
// not the actual boundary).
//
// The spreadsheet import is open to every signed-in staff member now: an
// admin's import lands straight in the ledger as before, but a regular
// staff member's import goes to a Review queue (PendingEntries on the
// backend) and only counts once an admin approves it — see
// leave-review.js and importDetails' role-based wiring below.
//
// PIN model: a PIN handed out by an admin is temporary — the first time
// it's used to sign in, this page forces a "set your own PIN" step
// before showing any data (see showChangePin()). Maziyah additionally
// has a master PIN: entering it on the login screen for ANY staff name
// signs her in as that person (their exact view/edit access), which is
// how she can check, edit or reset another staff member's data or PIN.
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
// Lets leave-review.js (approve/reject buttons) refresh the ledger +
// pending list after a decision, without duplicating the fetch logic.
window.MIC_ON_REVIEW_CHANGED = null;

const setupNotice = document.getElementById("setupNotice");
const loginBox = document.getElementById("loginBox");
const loginStaffSelect = document.getElementById("loginStaffSelect");
const loginPin = document.getElementById("loginPin");
const loginBtn = document.getElementById("loginBtn");
const loginStatus = document.getElementById("loginStatus");

const changePinBox = document.getElementById("changePinBox");
const changePinNote = document.getElementById("changePinNote");
const newPin1 = document.getElementById("newPin1");
const newPin2 = document.getElementById("newPin2");
const changePinBtn = document.getElementById("changePinBtn");
const changePinCancelBtn = document.getElementById("changePinCancelBtn");
const changePinStatus = document.getElementById("changePinStatus");
const changePinLinkBtn = document.getElementById("changePinLinkBtn");

const appEl = document.getElementById("app");
const sessionLabel = document.getElementById("sessionLabel");
const signOutBtn = document.getElementById("signOutBtn");
const staffSelect = document.getElementById("staffSelect");
const yearSelect = document.getElementById("yearSelect");
const ledgerSummary = document.getElementById("ledgerSummary");
const ledgerBody = document.getElementById("ledgerBody");
const entrySection = document.getElementById("entrySection");
const entryForm = document.getElementById("entryForm");
const entryStatus = document.getElementById("entryStatus");
const entrySubmitBtn = document.getElementById("entrySubmitBtn");
const entryCancelEditBtn = document.getElementById("entryCancelEditBtn");
const configStatus = document.getElementById("configStatus");
const saveConfigBtn = document.getElementById("saveConfigBtn");
const importDetails = document.getElementById("importDetails");
const importAdminNote = document.getElementById("importAdminNote");
const importStaffNote = document.getElementById("importStaffNote");
const configDetails = document.getElementById("configDetails");

let state = { config: [], entries: [], pending: [] };
let editingEntryId = null;

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
if (changePinBtn) changePinBtn.addEventListener("click", onChangePin);
if (changePinCancelBtn) changePinCancelBtn.addEventListener("click", function () {
changePinBox.hidden = true;
appEl.hidden = false;
});
if (changePinLinkBtn) changePinLinkBtn.addEventListener("click", function () { showChangePin(false); });
if (newPin1) newPin1.addEventListener("keydown", function (ev) { if (ev.key === "Enter") newPin2.focus(); });
if (newPin2) newPin2.addEventListener("keydown", function (ev) { if (ev.key === "Enter") onChangePin(); });

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
changePinBox.hidden = true;
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

// required === true: this is the forced first-sign-in flow (no Cancel,
// can't be dismissed). required === false: the person opened "Change
// PIN" from an already-signed-in session, and can cancel back to the app.
function showChangePin(required) {
loginBox.hidden = true;
appEl.hidden = true;
changePinBox.hidden = false;
newPin1.value = "";
newPin2.value = "";
changePinStatus.textContent = "";
changePinStatus.className = "status-msg";
changePinNote.textContent = required
? "Your PIN is temporary. Please set a new 6-digit PIN that only you know before continuing."
: "Choose a new 6-digit PIN.";
if (changePinCancelBtn) changePinCancelBtn.hidden = required;
newPin1.focus();
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

const auth = { token: data.token, staff: data.staff, role: data.role, viaMaster: !!data.viaMaster };
window.MIC_AUTH = auth;
if (data.mustChangePin) {
// Don't persist yet — the temporary PIN must not survive a reload
// without the person actually setting their own permanent one.
showChangePin(true);
} else {
saveAuth(auth);
startApp();
}
} catch (err) {
loginStatus.textContent = "Error: " + err.message;
loginStatus.className = "status-msg error";
} finally {
loginBtn.disabled = false;
}
}

async function onChangePin() {
const p1 = newPin1.value.trim();
const p2 = newPin2.value.trim();
if (!/^\d{6}$/.test(p1)) {
changePinStatus.textContent = "PIN must be exactly 6 digits.";
changePinStatus.className = "status-msg error";
return;
}
if (p1 !== p2) {
changePinStatus.textContent = "PINs don't match.";
changePinStatus.className = "status-msg error";
return;
}
changePinBtn.disabled = true;
changePinStatus.textContent = "Saving…";
changePinStatus.className = "status-msg";
try {
await postAction({ action: "changePin", newPin: p1 });
saveAuth(window.MIC_AUTH);
changePinBox.hidden = true;
startApp();
} catch (err) {
if (isAuthError(err)) { handleAuthError(); return; }
changePinStatus.textContent = "Error: " + err.message;
changePinStatus.className = "status-msg error";
} finally {
changePinBtn.disabled = false;
}
}

function onSignOut() {
clearSavedAuth();
window.MIC_AUTH = { token: null, staff: null, role: null };
window.location.reload();
}

function startApp() {
loginBox.hidden = true;
changePinBox.hidden = true;
appEl.hidden = false;

const auth = window.MIC_AUTH;
const isAdmin = auth.role === "admin";

// Khairul keeps full admin access (all staff visible, import/config
// shown) but the on-screen label reads as regular staff.
const HIDE_ADMIN_LABEL_FOR = ["Khairul"];
const showAdminLabel = isAdmin && HIDE_ADMIN_LABEL_FOR.indexOf(auth.staff) === -1;
sessionLabel.textContent = "Signed in as " + auth.staff +
(showAdminLabel ? " (admin)" : "") +
(auth.viaMaster ? " (via master PIN)" : "");

// Admin-only sections: hidden entirely for regular staff. The quick-add
// entry form and editing entitlement/carry-forward stay admin actions —
// only admins tally leave directly, so staff get a read-only ledger of
// their own record. The backend also refuses these actions for
// non-admins, so this is a convenience, not the actual boundary.
if (entrySection) entrySection.hidden = !isAdmin;
if (configDetails) configDetails.hidden = !isAdmin;

// The spreadsheet import is open to everyone; only the wording (and, in
// leave-import.js, who it can be imported "as") differs by role.
if (importDetails) importDetails.hidden = false;
if (importAdminNote) importAdminNote.hidden = !isAdmin;
if (importStaffNote) importStaffNote.hidden = isAdmin;

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
if (entryCancelEditBtn) entryCancelEditBtn.addEventListener("click", onCancelEdit);
saveConfigBtn.addEventListener("click", onSaveConfig);

window.MIC_ON_REVIEW_CHANGED = loadData;

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
state.pending = (data.pending || []).map(function (p) {
return Object.assign({}, p, {
StartDate: normalizeDate(p.StartDate),
EndDate: normalizeDate(p.EndDate)
});
});
render();
if (window.MIC_RENDER_REVIEW) window.MIC_RENDER_REVIEW(state.pending);
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

// How much of a staff member's Annual Leave entitlement has actually
// been released as of `asOf`, for the given `year`:
//   - a year already in the past is fully accrued (the whole entitlement)
//   - a year not yet started hasn't accrued anything
//   - the current year accrues progressively: 1 "share" released on the
//     1st of each month Jan–Oct, 2 shares on 1 Nov, 2 shares on 1 Dec,
//     and 1 more share on 31 Dec (15 shares total, matching a 15-day
//     entitlement). A different configured entitlement scales every
//     share proportionally, e.g. a 12-day entitlement releases 12/15 of
//     a share at each of the same release points.
function accruedAnnual(entitlement, year, asOf) {
const e = Number(entitlement) || 0;
const y = Number(year);
const currentYear = asOf.getFullYear();
if (y < currentYear) return e;
if (y > currentYear) return 0;

const sharesPerDay = e / 15;
const month = asOf.getMonth(); // 0 = Jan … 11 = Dec
const day = asOf.getDate();
let shares = 0;
for (let m = 0; m <= month; m++) {
if (m <= 9) { // Jan–Oct
shares += 1;
} else if (m === 10) { // Nov
shares += 2;
} else { // Dec
shares += 2;
if (day >= 31) shares += 1;
}
}
// Round to the nearest half-day for a clean display; exact for the
// default 15-day schedule, a close approximation for any other total.
return Math.round(shares * sharesPerDay * 2) / 2;
}

function render() {
const staff = currentStaff();
const year = currentYear();
if (!staff || !year) return;

const cfg = getConfigFor(staff, year);
const entries = entriesFor(staff, year);
const isAdmin = window.MIC_AUTH.role === "admin";

document.getElementById("cfgAnnualEntitlement").value = cfg.AnnualEntitlement || 0;
document.getElementById("cfgAnnualCarryForward").value = cfg.AnnualCarryForward || 0;
document.getElementById("cfgSickEntitlement").value = cfg.SickEntitlement || 0;
document.getElementById("cfgSickCarryForward").value = cfg.SickCarryForward || 0;

let annualBal = Number(cfg.AnnualCarryForward || 0) + accruedAnnual(cfg.AnnualEntitlement, year, new Date());
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
"<td>" + escapeHtml(fmtTimeRange(e.TimeFrom, e.TimeTo)) + "</td>" +
"<td>" + escapeHtml(e.Description || "") + "</td>" +
"<td>" + fmtDays(e.AnnualDays) + "</td>" +
"<td>" + fmtDays(e.SickDays) + "</td>" +
"<td>" + fmtDays(e.UnpaidDays) + "</td>" +
"<td>" + fmtDays(e.HospitalizeDays) + "</td>" +
"<td>" + annualBal + "</td>" +
"<td>" + sickBal + "</td>" +
"<td>" + (isAdmin ?
'<button type="button" class="btn-link edit-entry" data-id="' + escapeHtml(e.ID) + '">Edit</button> ' +
'<button type="button" class="btn-link delete-entry" data-id="' + escapeHtml(e.ID) + '">Delete</button>'
: "") + "</td>";
ledgerBody.appendChild(tr);
});

ledgerBody.querySelectorAll(".edit-entry").forEach(function (btn) {
btn.addEventListener("click", function () { onEditEntry(btn.getAttribute("data-id")); });
});
ledgerBody.querySelectorAll(".delete-entry").forEach(function (btn) {
btn.addEventListener("click", function () { onDeleteEntry(btn.getAttribute("data-id")); });
});

const annualNote = String(year) === String(new Date().getFullYear())
? "<p class=\"hint\">Annual leave accrues monthly: 1 day on the 1st of each month (Jan–Oct), 2 days on 1 Nov, 2 days on 1 Dec, plus 1 more on 31 Dec.</p>"
: "";

ledgerSummary.innerHTML =
'<div class="admin-preview">' +
"<div class=\"admin-row\"><strong>Annual Leave Balance</strong> &nbsp; " + annualBal + " day(s)</div>" +
"<div class=\"admin-row\"><strong>Sick Leave Balance</strong> &nbsp; " + sickBal + " day(s)</div>" +
"<div class=\"admin-row\"><strong>Unpaid Leave Taken (this year)</strong> &nbsp; " + unpaidTotal + " day(s)</div>" +
"<div class=\"admin-row\"><strong>Hospitalize Leave Taken (this year)</strong> &nbsp; " + hospitalizeTotal + " day(s)</div>" +
"</div>" + annualNote;
}

function fmtDays(v) {
const n = Number(v || 0);
return n === 0 ? "" : String(n);
}

function fmtTimeRange(from, to) {
from = (from || "").trim();
to = (to || "").trim();
if (!from && !to) return "";
if (from && to) return from + " - " + to;
return from || to;
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
const isEdit = !!editingEntryId;
entryStatus.textContent = isEdit ? "Saving changes…" : "Saving…";
entryStatus.className = "status-msg";
try {
const start = document.getElementById("entryStart").value;
if (!start) throw new Error("Start date is required.");
const end = document.getElementById("entryEnd").value || start;
const desc = document.getElementById("entryDesc").value.trim();
if (!desc) throw new Error("Description is required.");

const entryPayload = {
startDate: start,
endDate: end,
description: desc,
timeFrom: document.getElementById("entryTimeFrom").value || "",
timeTo: document.getElementById("entryTimeTo").value || "",
annualDays: document.getElementById("entryAnnual").value || 0,
sickDays: document.getElementById("entrySick").value || 0,
unpaidDays: document.getElementById("entryUnpaid").value || 0,
hospitalizeDays: document.getElementById("entryHospitalize").value || 0
};

if (isEdit) {
await postAction({ action: "editEntry", id: editingEntryId, entry: entryPayload });
} else {
entryPayload.staff = currentStaff();
await postAction({ action: "addEntry", entry: entryPayload });
}

entryForm.reset();
editingEntryId = null;
entrySubmitBtn.textContent = "+ Add Entry";
if (entryCancelEditBtn) entryCancelEditBtn.hidden = true;
entryStatus.textContent = isEdit ? "Entry updated." : "Entry added.";
entryStatus.className = "status-msg ok";
await loadData();
} catch (err) {
if (isAuthError(err)) { handleAuthError(); return; }
entryStatus.textContent = "Error: " + err.message;
entryStatus.className = "status-msg error";
}
}

function onEditEntry(id) {
const entry = state.entries.find(function (en) { return en.ID === id; });
if (!entry) return;
editingEntryId = id;
document.getElementById("entryStart").value = entry.StartDate || "";
document.getElementById("entryEnd").value = entry.EndDate || "";
document.getElementById("entryTimeFrom").value = entry.TimeFrom || "";
document.getElementById("entryTimeTo").value = entry.TimeTo || "";
document.getElementById("entryDesc").value = entry.Description || "";
document.getElementById("entryAnnual").value = entry.AnnualDays || 0;
document.getElementById("entrySick").value = entry.SickDays || 0;
document.getElementById("entryUnpaid").value = entry.UnpaidDays || 0;
document.getElementById("entryHospitalize").value = entry.HospitalizeDays || 0;
entrySubmitBtn.textContent = "Save Changes";
if (entryCancelEditBtn) entryCancelEditBtn.hidden = false;
entryStatus.textContent = "Editing this entry — update the fields above and Save Changes, or Cancel Edit.";
entryStatus.className = "status-msg";
entrySection.scrollIntoView({ behavior: "smooth", block: "start" });
}

function onCancelEdit() {
editingEntryId = null;
entryForm.reset();
entrySubmitBtn.textContent = "+ Add Entry";
if (entryCancelEditBtn) entryCancelEditBtn.hidden = true;
entryStatus.textContent = "";
entryStatus.className = "status-msg";
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
