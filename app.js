/* Micronet Leave Application — form logic, docx export, email handoff */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const statusEl = $("statusMsg");

  const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  function setStatus(msg, kind) {
    statusEl.textContent = msg;
    statusEl.className = "status-msg" + (kind ? " " + kind : "");
  }

  // ---- Auto-fill "Day" from the date pickers (matches the form's Date/Day/Time layout) ----
  function wireDayAutofill(dateId, dayId) {
    const dateInput = $(dateId);
    const dayInput = $(dayId);
    dateInput.addEventListener("change", () => {
      if (!dateInput.value) { dayInput.value = ""; return; }
      // Parse as local date (avoid UTC off-by-one)
      const [y, m, d] = dateInput.value.split("-").map(Number);
      const dt = new Date(y, m - 1, d);
      dayInput.value = DAY_NAMES[dt.getDay()];
    });
  }
  wireDayAutofill("startDate", "startDay");
  wireDayAutofill("endDate", "endDay");
  wireDayAutofill("resumeDate", "resumeDay");

  // ---- Formatting helpers ----
  function formatDateLong(isoDate) {
    if (!isoDate) return "";
    const [y, m, d] = isoDate.split("-").map(Number);
    const dt = new Date(y, m - 1, d);
    const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
    return `${d} ${months[m - 1]} ${y}`;
  }

  function formatDateDDMMYYYY(isoDate) {
    if (!isoDate) return "";
    const [y, m, d] = isoDate.split("-");
    return `${d}-${m}-${y}`;
  }

  function formatTime12h(isoTime) {
    if (!isoTime) return "";
    let [h, min] = isoTime.split(":").map(Number);
    const suffix = h >= 12 ? "PM" : "AM";
    h = h % 12;
    if (h === 0) h = 12;
    return `${h}.${String(min).padStart(2, "0")} ${suffix}`;
  }

  // ---- Collect + validate form data ----
  function collectFormData() {
    const form = $("leaveForm");
    if (!form.checkValidity()) {
      form.reportValidity();
      return null;
    }
    return {
      name: $("name").value.trim(),
      startDate: $("startDate").value,
      startDay: $("startDay").value,
      startTime: $("startTime").value,
      endDate: $("endDate").value,
      endDay: $("endDay").value,
      endTime: $("endTime").value,
      resumeDate: $("resumeDate").value,
      resumeDay: $("resumeDay").value,
      resumeTime: $("resumeTime").value,
      daysApplied: $("daysApplied").value,
      reason: $("reason").value.trim(),
      address: $("address").value.trim(),
      telephone: $("telephone").value.trim(),
      signature: $("signature").value.trim(),
      sigDate: $("sigDate").value
    };
  }

  // ---- DOCX generation (mirrors the Micronet Leave Application Form layout) ----
  async function buildDocxBlob(data) {
    const {
      Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
      BorderStyle, WidthType, AlignmentType, ShadingType, HeadingLevel, VerticalAlign, TabStopType
    } = docx;

    const logoResp = await fetch("assets/micronet-logo.png");
    const logoBuf = await logoResp.arrayBuffer();

    const thinBottomBorder = {
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" }
    };
    const boxBorder = {
      top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "000000" }
    };

    function fieldRow(label, valueRuns) {
      return new Paragraph({
        border: thinBottomBorder,
        spacing: { after: 120 },
        tabStops: [{ type: TabStopType.LEFT, position: 2600 }],
        children: [
          new TextRun({ text: label, bold: true }),
          new TextRun({ text: "\t" }),
          ...valueRuns
        ]
      });
    }

    function dateDayTimeRow(label, dateVal, dayVal, timeVal) {
      return new Paragraph({
        border: thinBottomBorder,
        spacing: { after: 120 },
        tabStops: [{ type: TabStopType.LEFT, position: 2600 }],
        children: [
          new TextRun({ text: label, bold: true }),
          new TextRun({ text: "\t" }),
          new TextRun({ text: "Date: ", bold: true }),
          new TextRun({ text: formatDateLong(dateVal) + "   " }),
          new TextRun({ text: "Day: ", bold: true }),
          new TextRun({ text: dayVal + "   " }),
          new TextRun({ text: "Time: ", bold: true }),
          new TextRun({ text: formatTime12h(timeVal) })
        ]
      });
    }

    const children = [];

    // Header: logo + title
    children.push(new Paragraph({
      children: [new ImageRun({ data: logoBuf, type: "png", transformation: { width: 110, height: 98 } })]
    }));
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 300 },
      children: [new TextRun({ text: "LEAVE APPLICATION FORM", bold: true, size: 28 })]
    }));

    // Applicant section
    children.push(fieldRow("Name", [new TextRun({ text: data.name })]));
    children.push(dateDayTimeRow("Leave Start", data.startDate, data.startDay, data.startTime));
    children.push(dateDayTimeRow("Leave End", data.endDate, data.endDay, data.endTime));
    children.push(dateDayTimeRow("Resume Duty On", data.resumeDate, data.resumeDay, data.resumeTime));
    children.push(fieldRow("No. Of Days Applied For", [new TextRun({ text: String(data.daysApplied) })]));
    children.push(fieldRow("Reason For Applying Leave", [new TextRun({ text: data.reason })]));
    children.push(fieldRow("Address While On Leave", [new TextRun({ text: data.address })]));
    children.push(fieldRow("Telephone While On Leave", [new TextRun({ text: "+673 " + data.telephone })]));
    children.push(fieldRow("Signature", [new TextRun({ text: data.signature, italics: true })]));
    children.push(fieldRow("Date", [new TextRun({ text: formatDateLong(data.sigDate) })]));

    children.push(new Paragraph({ spacing: { before: 300, after: 200 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: "000000" } }, children: [] }));

    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: "To Be Completed by Administration Department", bold: true })]
    }));

    children.push(new Paragraph({
      spacing: { after: 200 },
      children: [
        new TextRun({ text: "Type of Leave", bold: true }),
        new TextRun({ text: "\t\tAnnual \u25CB    Sick \u25CB    Advance \u25CB    Unpaid \u25CB" })
      ]
    }));
    children.push(new Paragraph({ spacing: { after: 200 },
      children: [ new TextRun({ text: "Leave Available", bold: true }), new TextRun({ text: "\t\t________________________________________" }) ] }));
    children.push(new Paragraph({ spacing: { after: 200 },
      children: [ new TextRun({ text: "Leave Remain", bold: true }), new TextRun({ text: "\t\t________________________________________" }) ] }));

    function signOffBlock(role, showApproval) {
      children.push(new Paragraph({ spacing: { before: 200 },
        children: [ new TextRun({ text: role, bold: true }), new TextRun({ text: "\t\t\tDate", bold: true }) ] }));
      if (showApproval) {
        children.push(new Paragraph({ children: [ new TextRun({ text: "Approved / Not Approved", bold: true }) ] }));
      }
      children.push(new Paragraph({ spacing: { after: 200 },
        children: [ new TextRun({ text: "…………………………………….\t……………………………………………………………………………………." }) ] }));
    }
    signOffBlock("Administrator", false);
    signOffBlock("General Manager", true);
    signOffBlock("Executive Director", true);
    signOffBlock("Managing Director", true);

    // Page 2: To Be Filled by MIC Admin
    children.push(new Paragraph({ children: [], pageBreakBefore: true }));
    children.push(new Paragraph({
      children: [new ImageRun({ data: logoBuf, type: "png", transformation: { width: 90, height: 80 } })]
    }));

    const headerCell = new TableCell({
      columnSpan: 3,
      shading: { type: ShadingType.CLEAR, fill: "D9E2F3" },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "To Be Filled by MIC Admin", bold: true, size: 24 })] })]
    });

    function adminRow(leaveLabel, balanceText) {
      return new TableRow({
        children: [
          new TableCell({ width: { size: 1129, type: WidthType.DXA }, children: [new Paragraph({ children: [] })] }),
          new TableCell({ width: { size: 4111, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: leaveLabel })] })] }),
          new TableCell({ width: { size: 4105, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: balanceText ? [new TextRun({ text: balanceText })] : [] })] })
        ]
      });
    }

    const adminTable = new Table({
      width: { size: 9345, type: WidthType.DXA },
      columnWidths: [1129, 4111, 4105],
      rows: [
        new TableRow({ children: [headerCell] }),
        adminRow("Annual Leave", "Balance Annual Leave as of: ___________________"),
        adminRow("Sick Leave", ""),
        adminRow("Unpaid Leave", ""),
        adminRow("Hospitalized Leave", "")
      ]
    });

    const finalDoc = new Document({
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 } } }, // A4
        children: [...children, adminTable]
      }]
    });

    return await Packer.toBlob(finalDoc);
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  function buildFilename(data) {
    const startDD = data.startDate ? data.startDate.split("-")[2] : "";
    const endFormatted = formatDateDDMMYYYY(data.endDate);
    const shortName = (LEAVE_FORM_CONFIG.applicantDisplayName || "").split(" ").pop() || "Leave";
    return `Leave Form (${startDD} & ${endFormatted}) - ${shortName}.docx`;
  }

  async function handleDownload() {
    const data = collectFormData();
    if (!data) { setStatus("Please fill in all required fields.", "error"); return; }
    setStatus("Generating document…");
    try {
      const blob = await buildDocxBlob(data);
      downloadBlob(blob, buildFilename(data));
      setStatus("Downloaded. Check your Downloads folder.", "ok");
    } catch (err) {
      console.error(err);
      setStatus("Could not generate the document: " + err.message, "error");
    }
    return data;
  }

  async function handleEmail() {
    const data = collectFormData();
    if (!data) { setStatus("Please fill in all required fields.", "error"); return; }

    setStatus("Generating document…");
    let blob;
    try {
      blob = await buildDocxBlob(data);
    } catch (err) {
      console.error(err);
      setStatus("Could not generate the document: " + err.message, "error");
      return;
    }
    const filename = buildFilename(data);
    downloadBlob(blob, filename);

    const startDD = data.startDate.split("-")[2];
    const endFormatted = formatDateDDMMYYYY(data.endDate);
    const shortName = (LEAVE_FORM_CONFIG.applicantDisplayName || "").split(" ").pop() || "";
    const subject = `Leave Form (${startDD} & ${endFormatted}) - ${shortName}`;

    const bodyLines = [
      `Dear ${LEAVE_FORM_CONFIG.emailGreetingName},`,
      "",
      `This is the leave form attached for my leave on ${formatDateLong(data.startDate)}, ${data.startDay} until ${formatDateLong(data.endDate)}, ${data.endDay}. If you do have any query, please do ask me.`,
      "",
      "Thank you.",
      "",
      LEAVE_FORM_CONFIG.signatureBlock
    ];
    const body = bodyLines.join("\n");

    const mailto = `mailto:${encodeURIComponent(LEAVE_FORM_CONFIG.emailTo)}` +
      `?cc=${encodeURIComponent(LEAVE_FORM_CONFIG.emailCc)}` +
      `&subject=${encodeURIComponent(subject)}` +
      `&body=${encodeURIComponent(body)}`;

    setStatus(`Downloaded "${filename}" — your email app is opening. Attach that file before sending.`, "ok");
    window.location.href = mailto;
  }

  $("downloadBtn").addEventListener("click", handleDownload);
  $("emailBtn").addEventListener("click", handleEmail);

  // ---- PWA: register service worker ----
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("sw.js").catch((err) => console.warn("SW registration failed", err));
    });
  }
})();
