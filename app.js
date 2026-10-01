/* Micronet Leave Application — form logic, docx export (from real template), and email handoff */
(function () {
  "use strict";

  // If this page is restored from the browser's back-forward cache (e.g.
  // the user clicks "Back" after visiting Staff Leave Breakdown), the DOM
  // and scripts are NOT re-run — the browser just resurrects whatever was
  // on screen before navigating away, which can be a stale render (e.g. the
  // Name dropdown showing short names from an older app.js). Force a full
  // reload in that case so the page always reflects the current code.
  window.addEventListener("pageshow", function (event) {
    if (event.persisted) {
      location.reload();
    }
  });

  const $ = (id) => document.getElementById(id);
  const statusEl = $("statusMsg");

  const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const TEMPLATE_URL = "assets/leave-form-template.docx";

  // ---------------------------------------------------------------------
  // Draft autosave: on mobile, "Download Filled Form" leaves the page (the
  // browser opens its own file preview/share sheet to hand off the .docx),
  // and coming back via "Back" triggers the pageshow/persisted reload
  // above — which wipes every field the person just filled in, right
  // before they'd want to press "Prepare Submission Email" with the same
  // details. To fix that without giving up the reload (still needed to
  // avoid stale code), every field is saved to localStorage as it's typed
  // and silently restored the moment the page (re)loads, so Download →
  // Back → Prepare Submission Email keeps the same data. The draft is
  // cleared once an email is successfully prepared (that leave request is
  // done) and expires on its own after a day, so returning to this page
  // long after abandoning a form doesn't resurrect a stale, confusing fill.
  // ---------------------------------------------------------------------
  const DRAFT_KEY = "micLeaveFormDraft";
  const DRAFT_MAX_AGE_MS = 24 * 60 * 60 * 1000;
  const DRAFT_IMAGE_MAX_BYTES = 1.5 * 1024 * 1024; // skip persisting an image above this size
  const DRAFT_FIELD_IDS = [
    "name", "startDate", "startDay", "startTime",
    "endDate", "endDay", "endTime",
    "resumeDate", "resumeDay", "resumeTime",
    "daysApplied", "reason", "address", "telephone", "sigDate"
  ];

  function bufferToBase64(buffer) {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }

  function base64ToBuffer(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes.buffer;
  }

  function saveDraft() {
    try {
      const draft = { savedAt: Date.now(), fields: {} };
      DRAFT_FIELD_IDS.forEach((id) => {
        const el = $(id);
        if (el) draft.fields[id] = el.value;
      });
      if (signatureImage && signatureImage.buffer.byteLength <= DRAFT_IMAGE_MAX_BYTES) {
        draft.signature = {
          base64: bufferToBase64(signatureImage.buffer),
          width: signatureImage.width, height: signatureImage.height, ext: signatureImage.ext
        };
      }
      if (mcImage && mcImage.buffer.byteLength <= DRAFT_IMAGE_MAX_BYTES) {
        draft.mc = {
          base64: bufferToBase64(mcImage.buffer),
          width: mcImage.width, height: mcImage.height, ext: mcImage.ext
        };
      }
      localStorage.setItem(DRAFT_KEY, JSON.stringify(draft));
    } catch (e) {
      // Storage unavailable, over quota, or private browsing — the typed
      // fields just won't survive a reload; nothing else to do here.
    }
  }

  function clearDraft() {
    try { localStorage.removeItem(DRAFT_KEY); } catch (e) { /* ignore */ }
  }

  function restoreImagePreview(payload, previewId) {
    const mime = payload.ext === "png" ? "image/png" : "image/jpeg";
    const buffer = base64ToBuffer(payload.base64);
    const blob = new Blob([buffer], { type: mime });
    const url = URL.createObjectURL(blob);
    const preview = $(previewId);
    preview.src = url;
    preview.hidden = false;
    return { buffer, width: payload.width, height: payload.height, ext: payload.ext };
  }

  function restoreDraft() {
    let raw;
    try {
      raw = localStorage.getItem(DRAFT_KEY);
    } catch (e) {
      return;
    }
    if (!raw) return;
    let draft;
    try {
      draft = JSON.parse(raw);
    } catch (e) {
      clearDraft();
      return;
    }
    if (!draft || !draft.fields || typeof draft.savedAt !== "number" ||
      Date.now() - draft.savedAt > DRAFT_MAX_AGE_MS) {
      clearDraft();
      return;
    }

    DRAFT_FIELD_IDS.forEach((id) => {
      const el = $(id);
      if (el && draft.fields[id] !== undefined) el.value = draft.fields[id];
    });

    if (draft.signature) {
      try { signatureImage = restoreImagePreview(draft.signature, "signaturePreview"); }
      catch (e) { signatureImage = null; }
    }
    if (draft.mc) {
      try { mcImage = restoreImagePreview(draft.mc, "mcPreview"); }
      catch (e) { mcImage = null; }
    }
  }

  function wireDraftAutosave() {
    DRAFT_FIELD_IDS.forEach((id) => {
      const el = $(id);
      if (!el) return;
      const evt = (el.tagName === "SELECT" || el.type === "date" || el.type === "time") ? "change" : "input";
      el.addEventListener(evt, saveDraft);
    });
  }

  // ---- Populate the Name dropdown from the shared staff roster (config.js) ----
  (function populateNameDropdown() {
    const select = $("name");
    if (!select) return;
    const roster = (typeof LEAVE_FORM_STAFF_ROSTER !== "undefined" && Array.isArray(LEAVE_FORM_STAFF_ROSTER))
      ? LEAVE_FORM_STAFF_ROSTER : [];
    roster.forEach((entry) => {
      const opt = document.createElement("option");
      opt.value = entry.full;
      opt.textContent = entry.full;
      opt.dataset.short = entry.short;
      opt.dataset.branch = entry.branch || "";
      select.appendChild(opt);
    });
  })();

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
    (img) => { signatureImage = img; saveDraft(); },
    () => { signatureImage = null; saveDraft(); });

  wireImageUpload("mcAttachment", "mcPreview",
    (img) => { mcImage = img; saveDraft(); },
    () => { mcImage = null; saveDraft(); });

  wireDraftAutosave();
  restoreDraft();

  // ---- Collect + validate form data ----
  function collectFormData() {
    const form = $("leaveForm");
    if (!form.checkValidity()) {
      form.reportValidity();
      return null;
    }
    const nameSelect = $("name");
    const selectedOption = nameSelect.selectedOptions && nameSelect.selectedOptions[0];
    return {
      name: nameSelect.value.trim(),
      shortName: (selectedOption && selectedOption.dataset.short) || "",
      branch: (selectedOption && selectedOption.dataset.branch) || "",
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

    const sigToken = '<w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>{{SIGNATURE_IMAGE}}</w:t></w:r>';
    if (xml.indexOf(sigToken) === -1) throw new Error("Template is missing the signature placeholder run.");
    if (signatureImage) {
      const sigExt = signatureImage.ext === "png" ? "png" : "jpeg";
      ensureContentTypeDefault(sigExt, sigExt === "png" ? "image/png" : "image/jpeg");
      zip.file(`word/media/signature.${sigExt}`, signatureImage.buffer);
      relsXml = relsXml.replace(
        "</Relationships>",
        `<Relationship Id="rIdSignatureImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/signature.${sigExt}"/></Relationships>`
      );
      const sigSize = scaledEmuSize(signatureImage.width, signatureImage.height, 1700000, 500000);
      xml = xml.replace(sigToken, inlineImageRunXml("rIdSignatureImg", 900, sigSize.cx, sigSize.cy));
    } else {
      // No signature uploaded — leave the line blank so the printed copy can be signed by hand.
      xml = xml.replace(sigToken, "");
    }

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
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Supporting Document(s) Attached:</w:t></w:r>' +
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
    // The Name field is always picked from the roster dropdown now, so its
    // short name rides along as data.shortName — no need to fuzzy-match
    // the full name against the roster. resolveShortName() is kept only
    // as a fallback for the unlikely case that's ever missing.
    const shortName = data.shortName || resolveShortName(data.name);
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

  // Signature sign-off for the "Prepare Submission Email" body. Staff we
  // have full contact details for (currently just Khairul, in config.js
  // signatureOverrides) get their complete personal block; anyone else
  // gets their name plus the general College address block only — we
  // don't have per-staff extension/mobile/email on file, so those lines
  // are left out rather than guessed or borrowed from someone else.
  function buildSignatureBlock(name) {
    const overrides = (LEAVE_FORM_CONFIG && LEAVE_FORM_CONFIG.signatureOverrides) || {};
    if (overrides[name]) return overrides[name];
    const collegeBlock = (LEAVE_FORM_CONFIG && LEAVE_FORM_CONFIG.genericSignatureCollegeBlock) || "";
    return ["Best Regards,", "", name, "", collegeBlock].join("\n");
  }

    // gets their name plus the general College address block only — we
  // don't have per-staff extension/mobile/email on file, so those lines
  // are left out rather than guessed or borrowed from someone else.
  // shortName picks the override (signatureOverrides is keyed by short
  // name, e.g. "Khairul"); displayName is what's actually printed under
  // "Best Regards," for anyone without a personal override — their full
  // name, not just their first name.
  function buildSignatureBlock(shortName, displayName) {
    const overrides = (LEAVE_FORM_CONFIG && LEAVE_FORM_CONFIG.signatureOverrides) || {};
    if (overrides[shortName]) return overrides[shortName];
    const collegeBlock = (LEAVE_FORM_CONFIG && LEAVE_FORM_CONFIG.genericSignatureCollegeBlock) || "";
    return ["Best Regards,", "", displayName, "", collegeBlock].join("\n");
  }

  // CC depends on the applicant's branch (see LEAVE_FORM_STAFF_ROSTER in
  // config.js): Jerudong staff CC both Sharon and Aqilah, Gadong staff CC
  // Sharon only. Falls back to the Gadong list if branch is ever unknown.
  function resolveEmailCc(branch) {
    const byBranch = (LEAVE_FORM_CONFIG && LEAVE_FORM_CONFIG.emailCcByBranch) || {};
    return byBranch[branch] || byBranch.Gadong || "";
  }

  // ---- Build the email subject/body/mailto for a given generated document.
  // Shared by both hand-off paths below (native Share, and the old
  // download-then-mailto fallback) so the two stay in sync.
  function buildEmailContent(data, filename) {
    const subject = filename.replace(/\.docx$/i, "");
    const bodyLines = [
      `Dear ${LEAVE_FORM_CONFIG.emailGreetingName},`,
      "",
      `This is the leave form attached for my leave on ${formatDateLong(data.startDate)}, ${data.startDay} until ${formatDateLong(data.endDate)}, ${data.endDay}. If you do have any query, please do ask me.`,
      "",
      "Thank you.",
      "",
      buildSignatureBlock(data.shortName, data.name)
    ];
    const body = bodyLines.join("\n");
    const mailto = `mailto:${encodeURIComponent(LEAVE_FORM_CONFIG.emailTo)}` +
      `?cc=${encodeURIComponent(resolveEmailCc(data.branch))}` +
      `&subject=${encodeURIComponent(subject)}` +
      `&body=${encodeURIComponent(body)}`;
    return { subject, body, mailto };
  }

  // Web Share API (with a file) lets the OS share sheet hand the document
  // straight to Gmail/Outlook/WhatsApp etc. as a real attachment — no
  // Downloads-folder detour. Only some mobile browsers support sharing
  // files this way, so this is feature-detected per file, and the
  // download+mailto fallback below always stays available too.
  function canShareFile(file) {
    try {
      return !!(navigator.share && navigator.canShare && navigator.canShare({ files: [file] }));
    } catch (e) {
      return false;
    }
  }

  function hideDocPreview() {
    const overlay = $("previewOverlay");
    overlay.hidden = true;
    $("previewBody").innerHTML = "";
    const statusEl = $("previewStatusMsg");
    statusEl.textContent = "";
    statusEl.className = "status-msg";
  }

  function finalizeDownloadAndEmail(blob, filename, data) {
    downloadBlob(blob, filename);
    const { mailto } = buildEmailContent(data, filename);
    setStatus(`Downloaded "${filename}" — your email app is opening. Attach that file before sending.`, "ok");
    clearDraft();
    hideDocPreview();
    window.location.href = mailto;
  }

  async function finalizeShare(blob, filename, data) {
    const mime = blob.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const file = new File([blob], filename, { type: mime });
    const { subject, body } = buildEmailContent(data, filename);
    const previewStatus = $("previewStatusMsg");
    try {
      await navigator.share({ files: [file], title: subject, text: body });
      setStatus(`Shared "${filename}" — finish sending it from the app you picked.`, "ok");
      clearDraft();
      hideDocPreview();
    } catch (err) {
      if (err && err.name === "AbortError") {
        // User backed out of the share sheet — leave the preview open and
        // the draft intact so they can just try again.
        return;
      }
      console.error(err);
      previewStatus.textContent = "Could not share the document: " + err.message;
      previewStatus.className = "status-msg error";
    }
  }

  // Shows the generated .docx rendered on-page (via docx-preview) before the
  // person commits to downloading or sharing it, so they can check it looks
  // right first. Rendering is best-effort: if the preview library fails for
  // any reason, the actual document is still fine and both action buttons
  // still work — only the visual preview itself is skipped.
  function showDocPreview(blob, filename, data) {
    const overlay = $("previewOverlay");
    const loading = $("previewLoading");
    const body = $("previewBody");
    const shareBtn = $("previewShareBtn");
    const downloadBtn = $("previewDownloadBtn");
    const closeBtn = $("previewCloseBtn");
    const statusEl = $("previewStatusMsg");

    $("previewFilename").textContent = filename;
    statusEl.textContent = "";
    statusEl.className = "status-msg";
    body.innerHTML = "";
    loading.hidden = false;
    overlay.hidden = false;

    const mime = blob.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const file = new File([blob], filename, { type: mime });
    shareBtn.hidden = !canShareFile(file);

    shareBtn.onclick = () => finalizeShare(blob, filename, data);
    downloadBtn.onclick = () => finalizeDownloadAndEmail(blob, filename, data);
    closeBtn.onclick = () => hideDocPreview();

    if (typeof docx === "undefined" || !docx.renderAsync) {
      loading.hidden = true;
      body.innerHTML = '<p style="padding:16px;color:#666;">Preview isn\'t available right now, but your document was generated fine — use the buttons below.</p>';
      return;
    }

    blob.arrayBuffer()
      .then((buf) => docx.renderAsync(buf, body, body))
      .then(() => { loading.hidden = true; })
      .catch((err) => {
        console.error(err);
        loading.hidden = true;
        body.innerHTML = '<p style="padding:16px;color:#666;">Could not render a preview, but your document was generated fine — use the buttons below.</p>';
      });
  }

  // "Prepare Submission Email" now builds the document itself (no need to
  // press "Download Filled Form" first) and shows a preview before handing
  // it off, rather than silently downloading and jumping to the email app.
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
    setStatus("");
    showDocPreview(blob, filename, data);
    return data;
  }

  $("downloadBtn").addEventListener("click", handleDownload);
  $("emailBtn").addEventListener("click", handleEmail);

  $("previewOverlay").addEventListener("click", (e) => {
    if (e.target === $("previewOverlay")) hideDocPreview();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !$("previewOverlay").hidden) hideDocPreview();
  });

  // ---- PWA: service worker registration + auto-update now live in
  // sw-update.js (shared by every page) ----
})();
