// ---------------------------------------------------------------------
// Staff Leave Breakdown — reads/writes leave records through the Google
// Apps Script Web App configured in config.js (LEAVE_LEDGER_API_URL).
//
// Data model:
//   StaffConfig: one row per (staff, year) holding Annual/Sick
//     entitlement + carry-forward for that year.
//   LeaveEntries: one row per leave event (a date range, a description,
//     and however many Annual/Sick/Unpaid/Hospitalize days it used).
//
// Running balances are computed here in the browser (Annual/Sick balance
// = carry-forward + entitlement − sum of days used so far this year);
// Unpaid/Hospitalize have no entitlement, they're just running totals.
// ---------------------------------------------------------------------

(function () {
  const API_URL = typeof LEAVE_LEDGER_API_URL !== "undefined" ? LEAVE_LEDGER_API_URL : "";
  const STAFF_NAMES = (typeof LEAVE_FORM_STAFF_NAMES !== "undefined" && Array.isArray(LEAVE_FORM_STAFF_NAMES))
    ? LEAVE_FORM_STAFF_NAMES : [];

  const setupNotice = document.getElementById("setupNotice");
  const appEl = document.getElementById("app");
  const staffSelect = document.getElementById("staffSelect");
  const yearSelect = document.getElementById("yearSelect");
  const ledgerSummary = document.getElementById("ledgerSummary");
  const ledgerBody = document.getElementById("ledgerBody");
  const entryForm = document.getElementById("entryForm");
  const entryStatus = document.getElementById("entryStatus");
  const configStatus = document.getElementById("configStatus");
  const saveConfigBtn = document.getElementById("saveConfigBtn");

  let state = { config: [], entries: [] };

  function init() {
    if (!API_URL) {
      setupNotice.hidden = false;
      appEl.hidden = true;
      return;
    }
    setupNotice.hidden = true;
    appEl.hidden = false;

    STAFF_NAMES.forEach(function (name) {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      staffSelect.appendChild(opt);
    });

    const currentYear = new Date().getFullYear();
    for (let y = currentYear - 1; y <= currentYear + 1; y++) {
      const opt = document.createElement("option");
      opt.value = String(y);
      opt.textContent = String(y);
      if (y === currentYear) opt.selected = true;
      yearSelect.appendChild(opt);
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
      const resp = await fetch(API_URL, { cache: "no-store" });
      const data = await resp.json();
      if (!data.ok) throw new Error(data.error || "Failed to load data");
      state.config = data.config || [];
      state.entries = data.entries || [];
      render();
    } catch (err) {
      ledgerSummary.textContent = "Could not load data: " + err.message;
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
    const resp = await fetch(API_URL, {
      method: "POST",
      // text/plain avoids a CORS preflight that Apps Script Web Apps can't answer;
      // the script itself still JSON.parses the body regardless of this header.
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
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
      configStatus.textContent = "Error: " + err.message;
      configStatus.className = "status-msg error";
    }
  }

  init();
})();
