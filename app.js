/* Micronet Leave Application — form logic, docx export (from real template), and email handoff */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const statusEl = $("statusMsg");

  const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const TEMPLATE_URL = "assets/leave-form-template.docx";

  // Signature / MC attachment state, populated by the file inputs below.
  let signatureImage = null; // { buffer: ArrayBuffer, width, height, ext: "png"|"jpeg" }
  let mcImage = null;        // same shape, or null if not attached

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

  function escapeXml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  // ---- Image file handling (signature + optional MC attachment) ----
  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  function loadImageDimensions(objectUrl) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
      img.onerror = reject;
      img.src = objectUrl;
    });
  }

  function extFromMime(mime) {
    return mime && mime.indexOf("png") !== -1 ? "png" : "jpeg";
  }

  function wireImageUpload(inputId, previewId, onLoaded, onCleared) {
    const input = $(inputId);
    const preview = $(previewId);
    input.addEventListener("change", async (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) { onCleared(); preview.hidden = true; preview.src = ""; return; }
      try {
        const buffer = await readFileAsArrayBuffer(file);
        const objectUrl = URL.createObjectURL(file);
        const dims = await loadImageDimensions(objectUrl);
        preview.src = objectUrl;
        preview.hidden = false;
        onLoaded({ buffer, width: dims.width, height: dims.height, ext: extFromMime(file.type) });
      } catch (err) {
        console.error(err);
        setStatus("Could not read the selected image.", "error");
        onCleared();
        preview.hidden = true;
      }
    });
  }

  wireImageUpload("signature", "signaturePreview",
    (img) => { signatureImage = img; },
    () => { signatureImage = null; });

  wireImageUpload("mcAttachment", "mcPreview",
    (img) => { mcImage = img; },
    () => { mcImage = null; });

  // ---- Collect + validate form data ----
  function collectFormData() {
    const form = $("leaveForm");
    if (!form.checkValidity()) {
      form.reportValidity();
      return null;
    }
    if (!signatureImage) {
      setStatus("Please upload a signature image.", "error");
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
      sigDate: $("sigDate").value
    };
  }

  // ---- Build an <a:graphic> inline drawing run embedding a picture via relationship id ----
  function inlineImageRunXml(relId, docPrId, cx, cy) {
    return (
      '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>' +
      `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      `<wp:docPr id="${docPrId}" name="Image${docPrId}"/>` +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      `<pic:nvPicPr><pic:cNvPr id="${docPrId}" name="Image${docPrId}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${relId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    );
  }

  function scaledEmuSize(srcWidth, srcHeight, maxWidthEmu, maxHeightEmu) {
    const EMU_PER_PX = 9525; // 96 dpi
    const scale = Math.min(maxWidthEmu / (srcWidth * EMU_PER_PX), maxHeightEmu / (srcHeight * EMU_PER_PX));
    return {
      cx: Math.round(srcWidth * EMU_PER_PX * scale),
      cy: Math.round(srcHeight * EMU_PER_PX * scale)
    };
  }

  // ---- DOCX generation: fill the real Micronet template (XML string replacement), never rebuild from scratch ----
  async function buildDocxBlob(data) {
    const templateResp = await fetch(TEMPLATE_URL);
    if (!templateResp.ok) throw new Error("Could not load the leave form template.");
    const templateBuf = await templateResp.arrayBuffer();
    const zip = await JSZip.loadAsync(templateBuf);

    let xml = await zip.file("word/document.xml").async("string");

    const textTokens = {
      "{{NAME}}": escapeXml(data.name),
      "{{START_DATE}}": escapeXml(formatDateLong(data.startDate)),
      "{{START_DAY}}": escapeXml(data.startDay),
      "{{START_TIME}}": escapeXml(formatTime12h(data.startTime)),
      "{{END_DATE}}": escapeXml(formatDateLong(data.endDate)),
      "{{END_DAY}}": escapeXml(data.endDay),
      "{{END_TIME}}": escapeXml(formatTime12h(data.endTime)),
      "{{RESUME_DATE}}": escapeXml(formatDateLong(data.resumeDate)),
      "{{RESUME_DAY}}": escapeXml(data.resumeDay),
      "{{RESUME_TIME}}": escapeXml(formatTime12h(data.resumeTime)),
      "{{DAYS_APPLIED}}": escapeXml(data.daysApplied),
      "{{REASON}}": escapeXml(data.reason),
      "{{ADDRESS}}": escapeXml(data.address),
      "{{PHONE}}": escapeXml(data.telephone),
      "{{SIG_DATE}}": escapeXml(formatDateLong(data.sigDate))
    };
    for (const [token, value] of Object.entries(textTokens)) {
      if (xml.indexOf(token) === -1) throw new Error(`Template is missing expected placeholder ${token}`);
      xml = xml.replace(token, value);
    }

    // ---- Signature image: embed as a relationship + inline drawing in place of the token run ----
    const relsPath = "word/_rels/document.xml.rels";
    let relsXml = await zip.file(relsPath).async("string");
    const ctPath = "[Content_Types].xml";
    let ctXml = await zip.file(ctPath).async("string");

    function ensureContentTypeDefault(ext, mime) {
      const marker = `Extension="${ext}"`;
      if (ctXml.indexOf(marker) === -1) {
        ctXml = ctXml.replace("</Types>", `<Default Extension="${ext}" ContentType="${mime}"/></Types>`);
      }
    }

    const sigExt = signatureImage.ext === "png" ? "png" : "jpeg";
    ensureContentTypeDefault(sigExt, sigExt === "png" ? "image/png" : "image/jpeg");
    zip.file(`word/media/signature.${sigExt}`, signatureImage.buffer);
    relsXml = relsXml.replace(
      "</Relationships>",
      `<Relationship Id="rIdSignatureImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/signature.${sigExt}"/></Relationships>`
    );
    const sigSize = scaledEmuSize(signatureImage.width, signatureImage.height, 1500000, 230000);
    const sigToken = '<w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>{{SIGNATURE_IMAGE}}</w:t></w:r>';
    if (xml.indexOf(sigToken) === -1) throw new Error("Template is missing the signature placeholder run.");
    xml = xml.replace(sigToken, inlineImageRunXml("rIdSignatureImg", 900, sigSize.cx, sigSize.cy));

    // ---- Optional Medical Certificate attachment (page 2), only if the user attached one ----
    const mcToken = "<w:r><w:t>{{MC_ATTACHMENT}}</w:t></w:r>";
    if (xml.indexOf(mcToken) === -1) throw new Error("Template is missing the MC attachment placeholder.");
    if (mcImage) {
      const mcExt = mcImage.ext === "png" ? "png" : "jpeg";
      ensureContentTypeDefault(mcExt, mcExt === "png" ? "image/png" : "image/jpeg");
      zip.file(`word/media/mc-attachment.${mcExt}`, mcImage.buffer);
      relsXml = relsXml.replace(
        "</Relationships>",
        `<Relationship Id="rIdMcImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/mc-attachment.${mcExt}"/></Relationships>`
      );
      const mcSize = scaledEmuSize(mcImage.width, mcImage.height, 5940000, 7000000);
      const mcXml =
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Medical Certificate / Supporting Document Attached:</w:t></w:r>' +
        '<w:r><w:br/></w:r>' +
        inlineImageRunXml("rIdMcImg", 901, mcSize.cx, mcSize.cy);
      xml = xml.replace(mcToken, mcXml);
    } else {
      xml = xml.replace(mcToken, "");
    }

    zip.file("word/document.xml", xml);
    zip.file(relsPath, relsXml);
    zip.file(ctPath, ctXml);

    return await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      compression: "DEFLATE"
    });
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

  // Match the typed Name field against the known staff roster (config.js) so the
  // exported filename uses a recognisable short name (e.g. "Khairul", not the
  // last word of the full name, which is often a family/title component).
  function resolveShortName(fullName) {
    const typed = (fullName || "").trim();
    const lowerTyped = typed.toLowerCase();
    const roster = (typeof LEAVE_FORM_STAFF_NAMES !== "undefined" && Array.isArray(LEAVE_FORM_STAFF_NAMES))
      ? LEAVE_FORM_STAFF_NAMES : [];
    for (const candidate of roster) {
      if (candidate && lowerTyped.indexOf(candidate.toLowerCase()) !== -1) {
        return candidate;
      }
    }
    const parts = typed.split(/\s+/).filter(Boolean);
    return parts[parts.length - 1] || "Leave";
  }

  function buildFilename(data) {
    const startFormatted = formatDateDDMMYYYY(data.startDate);
    const endFormatted = formatDateDDMMYYYY(data.endDate);
    const shortName = resolveShortName(data.name);
    return `Leave Form (${startFormatted} to ${endFormatted}) - ${shortName}.docx`;
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

    const subject = filename.replace(/\.docx$/i, "");

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
