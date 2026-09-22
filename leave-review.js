// ---------------------------------------------------------------------
// Staff Leave Breakdown — Review area for entries submitted via
// spreadsheet import.
//
// The spreadsheet import (leave-import.js) is open to every signed-in
// staff member. An admin's import still lands straight in the ledger,
// but a regular staff member's import is written to PendingEntries on
// the backend and only counts against their balance once an admin
// approves it here. This file renders that queue:
//   - Admin: sees every staff member's pending entries, with Approve /
//     Reject buttons.
//   - Staff: sees only their own pending entries, read-only, so they can
//     tell what's waiting on review.
//
// leave-breakdown.js calls window.MIC_RENDER_REVIEW(pendingArray) every
// time it (re)loads data, and this file exposes window.MIC_ON_REVIEW_-
// CHANGED (also set by leave-breakdown.js, to its own loadData) so an
// Approve/Reject click can refresh the ledger + queue afterwards.
// ---------------------------------------------------------------------

(function () {
const API_URL = typeof LEAVE_LEDGER_API_URL !== "undefined" ? LEAVE_LEDGER_API_URL : "";

const reviewHeading = document.getElementById("reviewHeading");
const reviewNote = document.getElementById("reviewNote");
const reviewBody = document.getElementById("reviewBody");
const reviewEmptyMsg = document.getElementById("reviewEmptyMsg");

if (!reviewBody || !API_URL) return;

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
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

async function onDecide(id, decision, btnRow) {
  if (decision === "reject" && !confirm("Reject this entry? It will not be added to the ledger.")) return;
  btnRow.querySelectorAll("button").forEach(function (b) { b.disabled = true; });
  try {
    await postAction({ action: decision === "approve" ? "approveEntry" : "rejectEntry", id: id });
    if (typeof window.MIC_ON_REVIEW_CHANGED === "function") await window.MIC_ON_REVIEW_CHANGED();
  } catch (err) {
    alert("Could not " + decision + " this entry: " + err.message);
    btnRow.querySelectorAll("button").forEach(function (b) { b.disabled = false; });
  }
}

function render(pending) {
  const isAdmin = window.MIC_AUTH && window.MIC_AUTH.role === "admin";

  if (reviewHeading) {
    reviewHeading.textContent = isAdmin ? "Pending Import Entries — Review" : "Your Pending Import Submissions";
  }
  if (reviewNote) {
    reviewNote.textContent = isAdmin
      ? "Approve or reject entries staff submitted via spreadsheet import, before they count against their leave balance."
      : "These were submitted via spreadsheet import and will appear on your ledger once an admin approves them.";
  }

  reviewBody.innerHTML = "";
  const list = pending || [];
  if (!list.length) {
    if (reviewEmptyMsg) reviewEmptyMsg.hidden = false;
    return;
  }
  if (reviewEmptyMsg) reviewEmptyMsg.hidden = true;

  list.forEach(function (p) {
    const tr = document.createElement("tr");
    tr.innerHTML =
      "<td>" + escapeHtml(p.Staff || "") + "</td>" +
      "<td>" + escapeHtml(p.StartDate || "") + "</td>" +
      "<td>" + escapeHtml(p.EndDate || "") + "</td>" +
      "<td>" + escapeHtml(fmtTimeRange(p.TimeFrom, p.TimeTo)) + "</td>" +
      "<td>" + escapeHtml(p.Description || "") + "</td>" +
      "<td>" + fmtDays(p.AnnualDays) + "</td>" +
      "<td>" + fmtDays(p.SickDays) + "</td>" +
      "<td>" + fmtDays(p.UnpaidDays) + "</td>" +
      "<td>" + fmtDays(p.HospitalizeDays) + "</td>" +
      "<td>" + escapeHtml(p.SubmittedAt || "") + "</td>" +
      "<td></td>";

    const actionsTd = tr.lastElementChild;
    if (isAdmin) {
      const approveBtn = document.createElement("button");
      approveBtn.type = "button";
      approveBtn.className = "btn-link";
      approveBtn.textContent = "Approve";
      const rejectBtn = document.createElement("button");
      rejectBtn.type = "button";
      rejectBtn.className = "btn-link";
      rejectBtn.textContent = "Reject";
      approveBtn.addEventListener("click", function () { onDecide(p.ID, "approve", actionsTd); });
      rejectBtn.addEventListener("click", function () { onDecide(p.ID, "reject", actionsTd); });
      actionsTd.appendChild(approveBtn);
      actionsTd.appendChild(document.createTextNode(" "));
      actionsTd.appendChild(rejectBtn);
    } else {
      actionsTd.textContent = "Pending";
    }

    reviewBody.appendChild(tr);
  });
}

window.MIC_RENDER_REVIEW = render;
})();
