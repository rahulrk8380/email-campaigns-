let contacts = [];
let currentCampaignId = null;
let selectedAttachments = [];
let selectedContactIds = new Set();
let composeSearchFilter = "";
let msStatus = { connected: false, email: null, name: null, clientIdConfigured: false };
let currentPdfFolderId = null;
let pdfFolderFileNames = [];
let excelHeaders = []; // All headers from last uploaded Excel
let _validationResult = null; // Cached validation result
let _pendingSendCallback = null; // Called after confirm modal approved
const $ = id => document.getElementById(id);

function toast(message, type = "info") {
  const el = $("toast");
  el.textContent = message;
  el.className = `toast show ${type}`;
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => el.classList.remove("show"), 3500);
}

async function api(url, options = {}) {
  const r = await fetch(url, options);
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(data.error || "Request failed");
  return data;
}

function esc(v) {
  return String(v ?? "").replace(/[&<>"']/g, m => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;"
  }[m]));
}

async function boot() {
  const params = new URLSearchParams(window.location.search);
  if (params.get("microsoft") === "connected") {
    toast("Microsoft Outlook account connected successfully!", "success");
    window.history.replaceState({}, document.title, window.location.pathname);
  } else if (params.get("error")) {
    const err = params.get("error");
    if (err === "microsoft_client_id_missing") {
      toast("Please enter your Microsoft Client ID first.", "error");
      setTimeout(() => openMsModal(), 400);
    } else {
      toast(`Microsoft: ${err}`, "error");
    }
    window.history.replaceState({}, document.title, window.location.pathname);
  }

  try {
    const me = await api("/api/me");
    if (me.loggedIn) showMain(me.email);
    else showLogin();
  } catch {
    showLogin();
  }
}

function showLogin() {
  $("loginView").classList.remove("hidden");
  $("mainView").classList.add("hidden");
}

function showMain(email) {
  $("loginView").classList.add("hidden");
  $("mainView").classList.remove("hidden");
  $("userPill").textContent = email;
  loadDashboard();
  loadContacts();
  loadCampaigns();
  loadMicrosoftStatus();
  loadSenders();
  updateSelectedCount();
  loadTemplates();
}

function showPage(page) {
  document.querySelectorAll(".page").forEach(x => x.classList.add("hidden"));
  $(`${page}Page`).classList.remove("hidden");
  document.querySelectorAll(".nav").forEach(x => x.classList.toggle("active", x.dataset.page === page));

  $("pageTitle").textContent =
    page === "dashboard" ? "Dashboard" :
    page === "contacts" ? "Contacts" :
    page === "compose" ? "Compose" : "Campaign History";

  if (page === "campaigns") loadCampaigns();
  if (page === "compose") updateSelectedCount();
}
window.showPage = showPage;

$("loginForm").addEventListener("submit", async e => {
  e.preventDefault();
  $("loginError").textContent = "";
  try {
    const me = await api("/api/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        email: $("loginEmail").value.trim(),
        password: $("loginPassword").value
      })
    });
    showMain(me.email);
  } catch (err) {
    $("loginError").textContent = err.message;
  }
});

$("logoutBtn").onclick = async () => {
  await api("/api/logout", { method: "POST" });
  showLogin();
};

document.querySelectorAll(".nav").forEach(b => {
  b.onclick = () => showPage(b.dataset.page);
});

async function loadDashboard() {
  try {
    const d = await api("/api/dashboard");
    $("statContacts").textContent = d.contacts;
    $("statCampaigns").textContent = d.campaigns;
    $("statSent").textContent = d.sent;
    $("statFailed").textContent = d.failed;
  } catch {}
}

function selectedIds() {
  return Array.from(selectedContactIds);
}

function updateSelectedCount() {
  const count = selectedContactIds.size;
  $("selectedCount").textContent = count;
  if ($("composeRecipientsSubtitle")) {
    $("composeRecipientsSubtitle").textContent = `${count} of ${contacts.length} rows selected`;
  }
  updateSendButtonLabel();
}

function updateSendButtonLabel() {
  const btn = $("sendCampaignBtn");
  if (!btn) return;
  const count = selectedContactIds.size;
  const select = $("senderSelect");
  let senderName = "Email";
  if (select && select.value === "microsoft") {
    senderName = "Microsoft Outlook";
  } else if (select && select.value === "smtp") {
    senderName = "Custom SMTP";
  }
  btn.innerHTML = `Send ${count} email${count === 1 ? "" : "s"} via ${senderName} <span>→</span>`;
}

// ── Contacts Table ─────────────────────────────────────────────────────────────

function getContactEmails(c) {
  let extras = {};
  try { extras = JSON.parse(c.extra_json || "{}"); } catch {}
  let emailArr = [];
  try { emailArr = JSON.parse(extras["__emails"] || "[]"); } catch {}
  if (!emailArr.length && c.email) emailArr = [c.email];
  return emailArr;
}

function getContactCC(c) {
  let extras = {};
  try { extras = JSON.parse(c.extra_json || "{}"); } catch {}
  return String(extras["__cc"] || "");
}

function getContactInvalid(c) {
  let extras = {};
  try { extras = JSON.parse(c.extra_json || "{}"); } catch {}
  return extras["__invalid"] || null;
}

function renderContacts() {
  $("contactCount").textContent = `${contacts.length} contacts (rows)`;
  $("contactsTable").innerHTML = contacts.map((c, i) => {
    const isChecked = selectedContactIds.has(c.id);
    const emails = getContactEmails(c);
    const cc = getContactCC(c);
    const invalid = getContactInvalid(c);

    const primaryEmail = emails[0] || "—";
    const extraCount = emails.length - 1;
    const extraBadge = extraCount > 0
      ? `<span class="email-more-badge">+${extraCount} more</span>` : "";

    const statusCls = invalid ? "invalid-row-status" : "ready";
    const statusLabel = invalid ? "⚠ No Email" : "Ready";

    const ccDisplay = cc
      ? cc.split(";").filter(Boolean).join(", ")
      : "—";

    return `
      <tr class="${invalid ? "row-invalid" : ""}">
        <td><input class="contact-check" type="checkbox" value="${c.id}" ${isChecked ? "checked" : ""}></td>
        <td>${i + 1}</td>
        <td><b>${esc(c.name)}</b></td>
        <td>${esc(primaryEmail)}${extraBadge}</td>
        <td>${extraCount > 0 ? emails.slice(1).map(esc).join("<br>") : "—"}</td>
        <td>${esc(ccDisplay)}</td>
        <td><span class="status ${statusCls}">${statusLabel}</span>${invalid ? `<div class="invalid-reason">${esc(invalid)}</div>` : ""}</td>
      </tr>
    `;
  }).join("") || `<tr><td colspan="7" class="empty">No contacts loaded. Upload an Excel file above.</td></tr>`;

  document.querySelectorAll(".contact-check").forEach(cb => {
    cb.addEventListener("change", e => {
      const id = Number(e.target.value);
      if (e.target.checked) selectedContactIds.add(id);
      else selectedContactIds.delete(id);
      $("selectAll").checked = contacts.length > 0 && selectedContactIds.size === contacts.length;
      updateSelectedCount();
      renderComposeRecipients();
    });
  });

  $("selectAll").checked = contacts.length > 0 && selectedContactIds.size === contacts.length;
  updateSelectedCount();
  renderComposeRecipients();
}

function renderComposeRecipients() {
  const tbody = $("composeRecipientsTbody");
  if (!tbody) return;

  const q = (composeSearchFilter || "").toLowerCase().trim();
  const list = contacts.filter(c => {
    if (!q) return true;
    return (c.name || "").toLowerCase().includes(q) || (c.email || "").toLowerCase().includes(q);
  });

  if (!list.length) {
    tbody.innerHTML = `<tr><td colspan="6" class="empty">${contacts.length === 0 ? "No contacts loaded yet. Upload in Contacts tab." : "No contacts match search filter."}</td></tr>`;
    return;
  }

  tbody.innerHTML = list.map((c, i) => {
    const isChecked = selectedContactIds.has(c.id);
    const emails = getContactEmails(c);
    const cc = getContactCC(c);
    const invalid = getContactInvalid(c);

    const primaryEmail = emails[0] || "—";
    const extraCount = emails.length - 1;
    const extraBadge = extraCount > 0 ? `<span class="email-more-badge">+${extraCount}</span>` : "";
    const ccDisplay = cc ? cc.split(";").filter(Boolean).slice(0, 1).join(", ") + (cc.split(";").filter(Boolean).length > 1 ? "..." : "") : "—";

    const statusCls = invalid ? "draft" : isChecked ? "ready" : "draft";
    const statusLabel = invalid ? "Invalid" : isChecked ? "Ready" : "Excluded";

    return `
      <tr class="${invalid ? "row-invalid" : ""}">
        <td><input class="compose-contact-check" type="checkbox" data-id="${c.id}" ${isChecked ? "checked" : ""}></td>
        <td>${i + 1}</td>
        <td><b>${esc(c.name)}</b></td>
        <td>${esc(primaryEmail)}${extraBadge}</td>
        <td>${esc(ccDisplay)}</td>
        <td><span class="status ${statusCls}">${statusLabel}</span></td>
      </tr>
    `;
  }).join("");

  tbody.querySelectorAll(".compose-contact-check").forEach(cb => {
    cb.addEventListener("change", e => {
      const id = Number(e.target.dataset.id);
      if (e.target.checked) selectedContactIds.add(id);
      else selectedContactIds.delete(id);
      renderContacts();
      updateSelectedCount();
    });
  });
}

function selectComposeRecipients(selectAll) {
  const q = (composeSearchFilter || "").toLowerCase().trim();
  const list = contacts.filter(c => {
    if (!q) return true;
    return (c.name || "").toLowerCase().includes(q) || (c.email || "").toLowerCase().includes(q);
  });

  list.forEach(c => {
    if (selectAll) selectedContactIds.add(c.id);
    else selectedContactIds.delete(c.id);
  });

  renderContacts();
  renderComposeRecipients();
  updateSelectedCount();
}
window.selectComposeRecipients = selectComposeRecipients;

async function loadContacts() {
  try {
    contacts = await api("/api/contacts");
    selectedContactIds = new Set(
      contacts.filter(c => !getContactInvalid(c)).map(c => c.id)
    );
    renderContacts();
    renderDynamicColumnChips();
  } catch {}
}

$("selectAll").addEventListener("change", e => {
  if (e.target.checked) {
    contacts.forEach(c => selectedContactIds.add(c.id));
  } else {
    selectedContactIds.clear();
  }
  renderContacts();
  updateSelectedCount();
});

const dz = $("dropzone");
const input = $("excelInput");

dz.onclick = () => input.click();

["dragenter", "dragover"].forEach(ev => dz.addEventListener(ev, e => {
  e.preventDefault();
  dz.classList.add("drag");
}));

["dragleave", "drop"].forEach(ev => dz.addEventListener(ev, e => {
  e.preventDefault();
  dz.classList.remove("drag");
}));

dz.addEventListener("drop", e => {
  if (e.dataTransfer.files[0]) uploadExcel(e.dataTransfer.files[0]);
});

input.addEventListener("change", () => {
  if (input.files[0]) uploadExcel(input.files[0]);
});

async function uploadExcel(file) {
  const fd = new FormData();
  fd.append("file", file);
  $("uploadResult").innerHTML = `<div class="loading">Uploading and validating <b>${esc(file.name)}</b>...</div>`;

  try {
    const d = await api("/api/upload", { method: "POST", body: fd });
    contacts = d.contacts;
    // Auto-select valid contacts, not invalid ones
    selectedContactIds = new Set(
      contacts.filter(c => !getContactInvalid(c)).map(c => c.id)
    );
    // Store Excel headers for dynamic variable display
    excelHeaders = d.headers || [];

    renderContacts();
    renderDynamicColumnChips();
    showAvailableVarsPanel();

    const invalidMsg = d.invalid > 0
      ? ` <span style="color:var(--red)">${d.invalid} row${d.invalid === 1 ? "" : "s"} have no valid email (shown in red — excluded from send).</span>`
      : "";

    $("uploadResult").innerHTML =
      `<div class="success-line">✓ ${d.added} rows with valid emails loaded. ${d.total} total rows processed.${invalidMsg}</div>`;
    loadDashboard();
    toast("Contacts imported successfully.", "success");
  } catch (e) {
    $("uploadResult").innerHTML = `<div class="error-line">${esc(e.message)}</div>`;
    toast(e.message, "error");
  }
}

async function clearContacts() {
  if (!confirm("Clear the current contact list?")) return;
  await api("/api/contacts", { method: "DELETE" });
  contacts = [];
  selectedContactIds.clear();
  excelHeaders = [];
  renderContacts();
  renderDynamicColumnChips();
  hideAvailableVarsPanel();
  loadDashboard();
  toast("Contact list cleared.", "success");
}
window.clearContacts = clearContacts;

// ── Available Variables Panel ─────────────────────────────────────────────────

function showAvailableVarsPanel() {
  const panel = $("availableVarsPanel");
  if (panel) panel.classList.remove("hidden");
  renderAvailableVarsChips();
}

function hideAvailableVarsPanel() {
  const panel = $("availableVarsPanel");
  if (panel) panel.classList.add("hidden");
}

function renderAvailableVarsChips() {
  const container = $("availableVarsChips");
  if (!container) return;

  // Use stored headers if available; fall back to deriving from contacts
  let headers = excelHeaders.length ? excelHeaders : [];
  if (!headers.length && contacts.length) {
    // Derive from extra_json keys (exclude internal __ keys)
    const seen = new Set();
    contacts.forEach(c => {
      try {
        const extras = JSON.parse(c.extra_json || "{}");
        Object.keys(extras).forEach(k => {
          if (!k.startsWith("__")) seen.add(k);
        });
      } catch {}
    });
    headers = [...seen];
  }

  const vars = headers.filter(h => !String(h).startsWith("__"));

  if (!vars.length) {
    container.innerHTML = `<span class="muted" style="font-size:11px">Upload an Excel file to see available variables.</span>`;
    return;
  }

  container.innerHTML = vars.map(h => `
    <button class="available-var-chip" onclick="insertVarFromPanel('{{${h}}}')" title="Insert {{${h}}} into selected field">
      {{${esc(h)}}}
    </button>
  `).join("");
}

function insertVarFromPanel(varStr) {
  // Determine target field from radio/select
  const target = getVarInsertTarget();
  const field = $(target);
  if (!field) return;
  field.focus();
  const a = field.selectionStart ?? field.value.length;
  const b = field.selectionEnd ?? a;
  field.value = field.value.slice(0, a) + varStr + field.value.slice(b);
  field.selectionStart = field.selectionEnd = a + varStr.length;
  field.dispatchEvent(new Event("input"));
  toast(`Inserted ${varStr} into ${target === "body" ? "Body" : "Subject"}`, "info");
}
window.insertVarFromPanel = insertVarFromPanel;

// ── Dynamic Column Chips (in compose variable-bar) ───────────────────────────

function renderDynamicColumnChips() {
  const bar = $("variableBar");
  if (!bar) return;

  // Remove any previously added dynamic chips
  bar.querySelectorAll(".dynamic-var-chip").forEach(el => el.remove());

  // Build variable list from Excel headers or extra_json
  let headers = excelHeaders.length ? excelHeaders : [];
  if (!headers.length && contacts.length) {
    const seen = new Set();
    contacts.forEach(c => {
      // Also include top-level fields
      ["name", "email", "number", "details"].forEach(k => seen.add(k));
      try {
        const extras = JSON.parse(c.extra_json || "{}");
        Object.keys(extras).forEach(k => {
          if (!k.startsWith("__")) seen.add(k);
        });
      } catch {}
    });
    headers = [...seen];
  }

  headers.filter(h => !String(h).startsWith("__")).forEach(key => {
    const varName = `{{${key}}}`;
    const btn = document.createElement("button");
    btn.textContent = varName;
    btn.className = "dynamic-var-chip";
    btn.title = `Insert ${varName} (from your Excel)`;
    btn.onclick = () => insertVar(varName);
    bar.appendChild(btn);
  });

  // Also update the available vars panel if visible
  renderAvailableVarsChips();
}

// ── Variable Insert Target (Subject or Body) ──────────────────────────────────

function getVarInsertTarget() {
  const sel = $("varTargetSelect");
  if (sel) return sel.value;
  // Fallback to radio buttons if present
  const radios = document.querySelectorAll("[name='varTarget']");
  for (const r of radios) {
    if (r.checked) return r.value;
  }
  return "body";
}

function insertVar(v) {
  const target = getVarInsertTarget();
  const t = $(target) || $("body"); // fallback to body
  if (!t) return;
  const a = t.selectionStart ?? t.value.length;
  const b = t.selectionEnd ?? a;
  t.value = t.value.slice(0, a) + v + t.value.slice(b);
  t.focus();
  t.selectionStart = t.selectionEnd = a + v.length;
  updatePreview();
}
window.insertVar = insertVar;

// ── Format helpers ─────────────────────────────────────────────────────────────

function formatBytes(bytes) {
  const n = Number(bytes || 0);
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

function fileIcon(type, name) {
  if (String(type).startsWith("image/")) return "🖼️";
  const ext = String(name).split(".").pop().toLowerCase();
  if (ext === "pdf") return "📕";
  if (["doc","docx"].includes(ext)) return "📘";
  if (["xls","xlsx","csv"].includes(ext)) return "📗";
  if (["ppt","pptx"].includes(ext)) return "📙";
  if (ext === "zip") return "🗜️";
  return "📄";
}

function renderAttachments() {
  const box = $("attachmentList");
  if (!selectedAttachments.length) {
    box.innerHTML = `<div class="attachment-empty">No files attached yet.</div>`;
    return;
  }
  box.innerHTML = selectedAttachments.map(a => `
    <div class="attachment-item">
      <div class="attachment-icon">${fileIcon(a.type, a.name)}</div>
      <div class="attachment-info"><b>${esc(a.name)}</b><span>${formatBytes(a.size)}</span></div>
      <button class="attachment-remove" title="Remove" onclick="removeAttachment(${a.id})">×</button>
    </div>
  `).join("");
}

async function uploadAttachments(files) {
  const incoming = [...files];
  if (!incoming.length) return;
  if (selectedAttachments.length + incoming.length > 8) return toast("You can attach up to 8 files.", "error");
  const total = selectedAttachments.reduce((s, a) => s + Number(a.size || 0), 0) + incoming.reduce((s, f) => s + f.size, 0);
  if (total > 24 * 1024 * 1024) return toast("Keep total attachments below 24 MB.", "error");

  const fd = new FormData();
  incoming.forEach(f => fd.append("files", f));
  setActionMessage("Uploading attachments...", "info");
  try {
    const d = await api("/api/attachments", { method: "POST", body: fd });
    selectedAttachments.push(...d.attachments);
    renderAttachments();
    setActionMessage(`${selectedAttachments.length} attachment${selectedAttachments.length === 1 ? "" : "s"} ready to send.`, "success");
    toast("Attachments added.", "success");
  } catch (e) {
    setActionMessage("✕ " + e.message, "error");
    toast(e.message, "error");
  }
}

// ── PDF Folder ────────────────────────────────────────────────────────────────

function renderPdfFolderStatus() {
  const statusEl = $("pdfFolderStatus");
  const dropzone = $("pdfFolderDropzone");
  if (!statusEl) return;

  if (!currentPdfFolderId || !pdfFolderFileNames.length) {
    statusEl.classList.add("hidden");
    statusEl.innerHTML = "";
    if (dropzone) dropzone.classList.remove("pdf-folder-loaded");
    return;
  }

  statusEl.classList.remove("hidden");
  const preview = pdfFolderFileNames.slice(0, 5).map(f => `<span class="pdf-name-chip">${esc(f)}</span>`).join("");
  const more = pdfFolderFileNames.length > 5 ? `<span class="pdf-name-chip muted">+${pdfFolderFileNames.length - 5} more</span>` : "";
  statusEl.innerHTML = `
    <div class="pdf-folder-loaded-row">
      <span>✅ <b>${pdfFolderFileNames.length} PDF${pdfFolderFileNames.length === 1 ? "" : "s"}</b> uploaded — will be matched by recipient name.</span>
      <button class="btn outline small-btn danger" onclick="clearPdfFolder()">✕ Clear</button>
    </div>
    <div class="pdf-name-chips">${preview}${more}</div>
  `;
  if (dropzone) dropzone.classList.add("pdf-folder-loaded");
}

async function uploadPdfFolder(files) {
  const pdfs = [...files].filter(f => /\.pdf$/i.test(f.name));
  if (!pdfs.length) return toast("No PDF files found. Please select PDF files only.", "error");

  setActionMessage(`Uploading ${pdfs.length} PDF${pdfs.length === 1 ? "" : "s"}...`, "info");

  if (currentPdfFolderId) {
    await api(`/api/pdf-folder/${currentPdfFolderId}`, { method: "DELETE" }).catch(() => {});
    currentPdfFolderId = null;
    pdfFolderFileNames = [];
  }

  const fd = new FormData();
  pdfs.forEach(f => fd.append("files", f));

  try {
    const d = await api("/api/pdf-folder", { method: "POST", body: fd });
    currentPdfFolderId = d.folderId;
    pdfFolderFileNames = d.filenames || [];
    renderPdfFolderStatus();
    setActionMessage(`✓ ${d.fileCount} personal PDFs ready. Contacts without a matching PDF will be skipped.`, "success");
    toast(`${d.fileCount} PDFs uploaded successfully.`, "success");
  } catch (e) {
    setActionMessage("✕ " + e.message, "error");
    toast(e.message, "error");
  }
}

async function clearPdfFolder() {
  if (currentPdfFolderId) {
    await api(`/api/pdf-folder/${currentPdfFolderId}`, { method: "DELETE" }).catch(() => {});
  }
  currentPdfFolderId = null;
  pdfFolderFileNames = [];
  renderPdfFolderStatus();
  setActionMessage("Personal PDFs cleared.", "info");
  toast("Personal PDF folder cleared.", "success");
}
window.clearPdfFolder = clearPdfFolder;

const pdfFolderInput = $("pdfFolderInput");
const pdfFolderDirInput = $("pdfFolderDirInput");
const pdfFolderZipInput = $("pdfFolderZipInput");
const pdfFolderDropzone = $("pdfFolderDropzone");

if (pdfFolderInput) pdfFolderInput.addEventListener("change", () => { if (pdfFolderInput.files.length) uploadPdfFolder(pdfFolderInput.files); pdfFolderInput.value = ""; });
if (pdfFolderDirInput) pdfFolderDirInput.addEventListener("change", () => { if (pdfFolderDirInput.files.length) uploadPdfFolder(pdfFolderDirInput.files); pdfFolderDirInput.value = ""; });
if (pdfFolderZipInput) pdfFolderZipInput.addEventListener("change", () => { if (pdfFolderZipInput.files.length) uploadPdfFolder(pdfFolderZipInput.files); pdfFolderZipInput.value = ""; });
if (pdfFolderDropzone) {
  ["dragenter", "dragover"].forEach(ev => pdfFolderDropzone.addEventListener(ev, e => { e.preventDefault(); pdfFolderDropzone.classList.add("drag"); }));
  ["dragleave", "drop"].forEach(ev => pdfFolderDropzone.addEventListener(ev, e => { e.preventDefault(); pdfFolderDropzone.classList.remove("drag"); }));
  pdfFolderDropzone.addEventListener("drop", e => { if (e.dataTransfer.files.length) uploadPdfFolder(e.dataTransfer.files); });
}

// ── Send Timer ────────────────────────────────────────────────────────────────
let delayUnit = "sec";

function setDelayUnit(unit) {
  delayUnit = unit;
  $("unitSec").classList.toggle("active", unit === "sec");
  $("unitMin").classList.toggle("active", unit === "min");
  updateDelayPreview();
}
window.setDelayUnit = setDelayUnit;

function getDelaySeconds() {
  const val = Number($("delayInput")?.value || 5);
  return delayUnit === "min" ? val * 60 : val;
}

function updateDelayPreview() {
  const val = Number($("delayInput")?.value || 5);
  const secs = delayUnit === "min" ? val * 60 : val;
  const label = delayUnit === "min"
    ? `${val} minute${val === 1 ? "" : "s"} (${secs} seconds)`
    : `${secs} second${secs === 1 ? "" : "s"}`;
  const el = $("delayPreview");
  if (el) el.textContent = `Each email will be sent ${label} apart`;
  const slider = $("delaySlider");
  if (slider) slider.value = Math.min(300, secs);
}

const delaySlider = $("delaySlider");
const delayInput = $("delayInput");
if (delaySlider && delayInput) {
  delaySlider.addEventListener("input", () => { delayInput.value = delaySlider.value; updateDelayPreview(); });
  delayInput.addEventListener("input", () => { delaySlider.value = Math.min(300, Number(delayInput.value) || 1); updateDelayPreview(); });
}

// ── Live Preview ──────────────────────────────────────────────────────────────

function updatePreview() {
  const subject = $("subject").value || "Your subject";
  const body = $("body").value || "Your personalized message will appear here.";

  // Use first selected contact's real data, or sample data
  const selectedList = contacts.filter(c => selectedContactIds.has(c.id));
  const sampleContact = selectedList[0] || null;

  let renderedSubject = subject;
  let renderedBody = body;
  let toLine = "";
  let ccLine = "";

  if (sampleContact) {
    // Client-side variable substitution using contact data
    renderedSubject = clientRenderTemplate(subject, sampleContact);
    renderedBody = clientRenderTemplate(body, sampleContact);
    const emails = getContactEmails(sampleContact);
    const cc = getContactCC(sampleContact);
    toLine = emails.join(", ") || sampleContact.email || "";
    ccLine = cc ? cc.split(";").filter(Boolean).join(", ") : "";
    if ($("previewSampleTag")) {
      $("previewSampleTag").textContent = `Preview: ${sampleContact.name}`;
    }
  } else {
    if ($("previewSampleTag")) $("previewSampleTag").textContent = "Sample: first row";
  }

  $("previewSubject").textContent = renderedSubject;
  $("previewBody").textContent = renderedBody;

  const toRowEl = $("previewToRow");
  if (toRowEl) {
    toRowEl.innerHTML = toLine
      ? `<span class="preview-to-label">To:</span> <span class="preview-to-val">${esc(toLine)}${ccLine ? ` &nbsp;|&nbsp; <span class="preview-cc-label">CC:</span> ${esc(ccLine)}` : ""}</span>`
      : "";
  }
}

// Client-side template rendering (mirrors server-side logic)
function clientRenderTemplate(template, contact) {
  let extras = {};
  try { extras = JSON.parse(contact.extra_json || "{}"); } catch {}
  const merged = { ...contact, ...extras };

  return String(template || "").replace(/\{\{([^}]+)\}\}/g, (match, varName) => {
    const key = varName.trim().toLowerCase().replace(/[^a-z0-9]/g, "");
    // Search merged keys case-insensitively
    for (const [k, v] of Object.entries(merged)) {
      if (k.startsWith("__")) continue;
      const normK = String(k).trim().toLowerCase().replace(/[^a-z0-9]/g, "");
      if (normK === key) return String(v ?? "");
    }
    return match; // keep original if not found
  });
}

$("subject").addEventListener("input", updatePreview);
$("body").addEventListener("input", updatePreview);

function setActionMessage(message, type = "info") {
  $("actionMessage").className = `action-message ${type}`;
  $("actionMessage").textContent = message;
}

// ── Microsoft Status & Configuration ──────────────────────────────────────────

async function loadMicrosoftStatus() {
  try {
    msStatus = await api("/api/microsoft/status");
    updateMsUI();
  } catch {}
}

function updateMsUI() {
  const pill = $("msStatusPill");
  const text = $("msStatusText");
  const dot = pill?.querySelector(".dot");
  const connectedBox = $("msConnectedBox");
  const emailSpan = $("msAccountEmail");
  const nameSpan = $("msAccountName");

  if (msStatus.connected) {
    pill?.classList.add("connected");
    dot?.classList.add("connected");
    if (text) text.textContent = `Outlook: ${msStatus.email}`;
    if (connectedBox) connectedBox.classList.remove("hidden");
    if (emailSpan) emailSpan.textContent = msStatus.email;
    if (nameSpan) nameSpan.textContent = msStatus.name || "Microsoft Account";
    if ($("msSubmitBtn")) $("msSubmitBtn").textContent = "Switch Microsoft Account ⊞";
  } else {
    pill?.classList.remove("connected");
    dot?.classList.remove("connected");
    if (text) text.textContent = msStatus.clientIdConfigured ? "Connect Outlook" : "Setup Outlook";
    if (connectedBox) connectedBox.classList.add("hidden");
    if ($("msSubmitBtn")) $("msSubmitBtn").textContent = "Sign in with Microsoft ⊞";
  }

  if (msStatus.config && $("msClientId") && !$("msClientId").value) {
    if (msStatus.config.clientId) $("msClientId").placeholder = `Configured (${msStatus.config.clientId})`;
    if (msStatus.config.tenantId) $("msTenantId").value = msStatus.config.tenantId;
  }
}

async function loadSenders() {
  try {
    const data = await api("/api/senders");
    const select = $("senderSelect");
    if (!select) return;

    if (!data.senders || !data.senders.length) {
      select.innerHTML = `<option value="microsoft">Microsoft Outlook (Not connected)</option>`;
      updateSendButtonLabel();
      return;
    }

    select.innerHTML = data.senders.map(s => `
      <option value="${s.type}" ${s.type === data.defaultSender ? "selected" : ""}>
        ${esc(s.label)} ${s.recommended ? "★ Recommended" : ""}
      </option>
    `).join("");

    updateSendButtonLabel();
  } catch {}
}

// ── Email Preview Modal ────────────────────────────────────────────────────────

function showEmailPreview() {
  const subject = $("subject").value.trim();
  const body = $("body").value.trim();

  if (!subject && !body) return toast("Write a subject and body first.", "error");

  const selectedList = contacts.filter(c => selectedContactIds.has(c.id));
  if (!selectedList.length) return toast("Select at least one recipient row.", "error");

  const c = selectedList[0];
  const emails = getContactEmails(c);
  const cc = getContactCC(c);

  const renderedSubject = clientRenderTemplate(subject || "(no subject)", c);
  const renderedBody = clientRenderTemplate(body || "(empty body)", c);

  // Check for unresolved variables
  const unresolvedSubj = findUnresolvedVars(renderedSubject);
  const unresolvedBody = findUnresolvedVars(renderedBody);
  const allUnresolved = [...new Set([...unresolvedSubj, ...unresolvedBody])];

  $("previewModalContact").textContent = `Preview for: ${c.name}`;
  $("previewModalTo").textContent = emails.join(", ") || c.email || "(no email)";
  $("previewModalSubject").textContent = renderedSubject;

  const ccRow = $("previewModalCcRow");
  const ccVal = cc ? cc.split(";").filter(Boolean).join(", ") : "";
  if (ccRow) ccRow.style.display = ccVal ? "flex" : "none";
  $("previewModalCc").textContent = ccVal;

  const bodyEl = $("previewModalBody");
  bodyEl.textContent = renderedBody;

  const warningsEl = $("previewModalWarnings");
  if (allUnresolved.length) {
    warningsEl.classList.remove("hidden");
    warningsEl.innerHTML = `⚠️ <b>Unknown variables detected:</b> ${allUnresolved.map(v => `<code>${esc(v)}</code>`).join(", ")} — these will be sent as-is.`;
  } else {
    warningsEl.classList.add("hidden");
  }

  $("previewModal").classList.remove("hidden");
}
window.showEmailPreview = showEmailPreview;

function closePreviewModal() {
  $("previewModal").classList.add("hidden");
}
window.closePreviewModal = closePreviewModal;

function findUnresolvedVars(text) {
  const matches = String(text || "").matchAll(/\{\{([^}]+)\}\}/g);
  return [...new Set([...matches].map(m => m[0]))];
}

// ── Validation Modal ──────────────────────────────────────────────────────────

async function runValidation(ids, subject, body) {
  const result = await api("/api/validate-recipients", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ selectedIds: ids, subject, body })
  });
  return result;
}

function showValidationModal(result, onProceed) {
  _validationResult = result;
  _pendingSendCallback = onProceed;

  const icon = $("validationModalIcon");
  const title = $("validationModalTitle");
  const subtitle = $("validationModalSubtitle");
  const banner = $("validationSummaryBanner");
  const issuesList = $("validationIssuesList");
  const issuesRows = $("validationIssuesRows");
  const proceedBtn = $("validationProceedBtn");

  const hasIssues = result.invalid > 0;

  icon.style.background = hasIssues
    ? "linear-gradient(135deg,#f59e0b,#d97706)"
    : "linear-gradient(135deg,#13a673,#0e7e5a)";
  icon.textContent = hasIssues ? "⚠" : "✓";

  title.textContent = hasIssues ? "Validation: Issues Found" : "Validation: All Clear";
  subtitle.textContent = hasIssues
    ? `${result.invalid} row(s) will be skipped`
    : "All selected rows are ready to send";

  banner.innerHTML = `
    <div class="validation-stat"><span>${result.total}</span><small>Total rows</small></div>
    <div class="validation-stat ok"><span>${result.valid}</span><small>✓ Will send</small></div>
    <div class="validation-stat ${result.invalid > 0 ? "warn" : ""}"><span>${result.invalid}</span><small>⚠ Will skip</small></div>
    <div class="validation-stat ${result.missingEmail > 0 ? "err" : ""}"><span>${result.missingEmail}</span><small>✕ No email</small></div>
  `;

  if (hasIssues) {
    issuesList.classList.remove("hidden");
    const badRows = result.rows.filter(r => !r.valid);
    issuesRows.innerHTML = badRows.slice(0, 20).map(r => `
      <div class="validation-issue-row">
        <span class="validation-issue-name">${esc(r.name)}</span>
        <span class="validation-issue-reason">${r.issues.map(esc).join(" · ")}</span>
      </div>
    `).join("") + (badRows.length > 20 ? `<div class="validation-issue-more">...and ${badRows.length - 20} more</div>` : "");
  } else {
    issuesList.classList.add("hidden");
  }

  if (result.valid === 0) {
    proceedBtn.disabled = true;
    proceedBtn.textContent = "No valid rows to send";
  } else {
    proceedBtn.disabled = false;
    proceedBtn.textContent = `Send ${result.valid} email${result.valid === 1 ? "" : "s"} →`;
  }

  $("validationModal").classList.remove("hidden");
}

function closeValidationModal() {
  $("validationModal").classList.add("hidden");
}
window.closeValidationModal = closeValidationModal;

function proceedAfterValidation() {
  closeValidationModal();
  if (_pendingSendCallback) _pendingSendCallback();
}
window.proceedAfterValidation = proceedAfterValidation;

// ── Confirm Modal ─────────────────────────────────────────────────────────────

function showConfirmModal(validResult, onConfirm) {
  _pendingSendCallback = onConfirm;

  $("confirmRows").textContent = validResult.total;
  $("confirmValid").textContent = validResult.valid;
  $("confirmInvalid").textContent = validResult.invalid;

  // Estimate total recipients
  const selectedList = contacts.filter(c => selectedContactIds.has(c.id) && !getContactInvalid(c));
  let totalRecipients = 0;
  selectedList.forEach(c => {
    totalRecipients += getContactEmails(c).length || 1;
  });
  $("confirmRecipients").textContent = totalRecipients;

  const warning = $("confirmWarning");
  if (validResult.invalid > 0) {
    warning.classList.remove("hidden");
    warning.textContent = `⚠ ${validResult.invalid} row(s) with no valid email will be skipped.`;
  } else {
    warning.classList.add("hidden");
  }

  $("confirmModal").classList.remove("hidden");
}

function closeConfirmModal() {
  $("confirmModal").classList.add("hidden");
}
window.closeConfirmModal = closeConfirmModal;

async function executeSend() {
  closeConfirmModal();
  if (_pendingSendCallback) {
    await _pendingSendCallback();
  }
}
window.executeSend = executeSend;

// ── Send Campaign Flow ────────────────────────────────────────────────────────

function openMsModal() {
  $("msModal").classList.remove("hidden");
  $("quickSmtpMsg") && ($("quickSmtpMsg").textContent = "");
  const base = window.location.origin;
  const redirectUri = base + "/auth/microsoft/callback";
  const el = $("redirectUriDisplay");
  if (el) el.textContent = redirectUri;
}

function closeMsModal() {
  $("msModal").classList.add("hidden");
  $("quickSmtpMsg") && ($("quickSmtpMsg").textContent = "");
}

async function sendCampaignDirect() {
  const ids = selectedIds();

  if (!ids.length) return toast("Select at least one recipient row.", "error");
  if (!$("subject").value.trim()) return toast("Enter a subject.", "error");
  if (!$("body").value.trim()) return toast("Enter a message body.", "error");

  const subject = $("subject").value;
  const body = $("body").value;
  const select = $("senderSelect");
  const senderType = select?.value || "microsoft";
  const senderLabel = senderType === "microsoft" ? "Microsoft Outlook" : "Custom SMTP";

  setActionMessage("Validating recipients...", "info");

  try {
    const validResult = await runValidation(ids, subject, body);

    if (validResult.valid === 0) {
      setActionMessage("✕ No valid rows to send. Check email addresses.", "error");
      return toast("No valid rows found. All selected rows have invalid/missing emails.", "error");
    }

    // Show validation modal — proceed → confirm modal → actual send
    showValidationModal(validResult, () => {
      showConfirmModal(validResult, () => doSendCampaign(ids, senderType, senderLabel, validResult));
    });

    setActionMessage("", "info");
  } catch (e) {
    setActionMessage("✕ " + e.message, "error");
    toast(e.message, "error");
  }
}
window.sendCampaignDirect = sendCampaignDirect;

async function doSendCampaign(ids, senderType, senderLabel, validResult) {
  const subject = $("subject").value;
  const body = $("body").value;

  // Warn if PDF folder active
  if (currentPdfFolderId) {
    if (!confirm(`Send campaign to ${validResult.valid} rows via ${senderLabel}?\n\n📂 Personal PDFs are active (${pdfFolderFileNames.length} PDFs). Rows without a matching PDF will be SKIPPED.\n\nProceed?`)) return;
  }

  try {
    const d = await api("/api/campaigns", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        name: $("campaignName").value.trim() || "Email Campaign",
        subject,
        body,
        selectedIds: ids,
        attachmentIds: selectedAttachments.map(a => a.id),
        senderType,
        pdfFolderId: currentPdfFolderId || null,
        delaySeconds: getDelaySeconds(),
        scheduledAt: getScheduledAt()
      })
    });

    if (d.usedPdfFolder) {
      if (d.skippedCount > 0) {
        toast(`📂 ${d.matchedCount} will be sent · ${d.skippedCount} skipped (no matching PDF)`, "info");
      } else {
        toast(`📂 All ${d.matchedCount} contacts matched!`, "success");
      }
    }

    currentCampaignId = d.campaignId;

    if (d.scheduled) {
      const when = new Date(d.scheduledAt).toLocaleString();
      toast(`📅 Campaign scheduled! Will send automatically at ${when}`, "success");
      showPage("campaigns");
      await loadCampaigns();
      return;
    }

    if (d.duplicates && d.duplicates.length > 0) {
      showDuplicateModal(d.duplicates, d.campaignId, senderLabel);
    } else {
      await proceedToSend(d.campaignId, senderLabel);
    }

  } catch (e) {
    toast(e.message, "error");
    setActionMessage("✕ " + e.message, "error");
    if (senderType === "microsoft" && e.message.includes("not connected")) {
      setTimeout(() => openMsModal(), 700);
    }
  }
}

// ── Duplicate Loan Number Modal ───────────────────────────────────────────────
let _pendingDuplicates = [];
let _pendingCampaignId = null;
let _pendingSenderLabel = "";

function showDuplicateModal(duplicates, campaignId, senderLabel) {
  _pendingDuplicates = duplicates;
  _pendingCampaignId = campaignId;
  _pendingSenderLabel = senderLabel;

  const list = $("duplicateList");
  list.innerHTML = duplicates.map(d => `
    <div class="duplicate-item">
      <span class="duplicate-loan">📄 Loan Number: <b>${esc(d.loanNumber)}</b></span>
      <span class="duplicate-badge">${d.count} contacts matched</span>
    </div>
  `).join("");

  $("duplicateModal").classList.remove("hidden");
}

async function handleDuplicates(action) {
  $("duplicateModal").classList.add("hidden");

  if (action === "deny") {
    const toSkip = [];
    for (const d of _pendingDuplicates) {
      toSkip.push(...d.recipientIds.slice(1));
    }
    if (toSkip.length) {
      try {
        await api(`/api/campaigns/${_pendingCampaignId}/deny-duplicates`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ recipientIds: toSkip })
        });
        toast(`${toSkip.length} duplicate contact(s) moved to Undeliverable.`, "info");
      } catch (e) {
        toast("Could not deny duplicates: " + e.message, "error");
      }
    }
  }

  await proceedToSend(_pendingCampaignId, _pendingSenderLabel);
}
window.handleDuplicates = handleDuplicates;

async function proceedToSend(campaignId, senderLabel) {
  try {
    await api(`/api/campaigns/${campaignId}/send`, { method: "POST" });
    toast(`Campaign started via ${senderLabel}. Tracking delivery...`, "success");
    showPage("campaigns");
    await loadCampaigns();
    openCampaign(campaignId);
  } catch (e) {
    toast(e.message, "error");
  }
}

// ── Schedule Helpers ──────────────────────────────────────────────────────────

function toggleSchedule() {
  const on = $("scheduleToggle") && $("scheduleToggle").checked;
  const box = $("scheduleBox");
  if (box) box.classList.toggle("hidden", !on);
  if (on && $("scheduledAt")) {
    const d = new Date(Date.now() + 60 * 60 * 1000);
    const pad = n => String(n).padStart(2, "0");
    $("scheduledAt").value = `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }
}
window.toggleSchedule = toggleSchedule;

function getScheduledAt() {
  if (!$("scheduleToggle") || !$("scheduleToggle").checked) return null;
  const val = $("scheduledAt") && $("scheduledAt").value;
  if (!val) return null;
  return new Date(val).toISOString();
}

async function removeAttachment(id) {
  try {
    await api(`/api/attachments/${id}`, { method: "DELETE" });
    selectedAttachments = selectedAttachments.filter(a => a.id !== id);
    renderAttachments();
    toast("Attachment removed.", "success");
  } catch (e) {
    toast(e.message, "error");
  }
}
window.removeAttachment = removeAttachment;

const attachmentInput = $("attachmentInput");
const attachmentDropzone = $("attachmentDropzone");
attachmentInput.addEventListener("change", () => { uploadAttachments(attachmentInput.files); attachmentInput.value = ""; });
["dragenter", "dragover"].forEach(ev => attachmentDropzone.addEventListener(ev, e => { e.preventDefault(); attachmentDropzone.classList.add("drag"); }));
["dragleave", "drop"].forEach(ev => attachmentDropzone.addEventListener(ev, e => { e.preventDefault(); attachmentDropzone.classList.remove("drag"); }));
attachmentDropzone.addEventListener("click", () => attachmentInput.click());
attachmentDropzone.addEventListener("drop", e => uploadAttachments(e.dataTransfer.files));
renderAttachments();

// ── Campaign History ──────────────────────────────────────────────────────────

function statusBadge(status) {
  const s = String(status || "").toLowerCase();
  const cls = s === "completed" ? "completed" : s === "sending" ? "sending" : s === "failed" ? "failed" : "draft";
  return `<span class="status ${cls}">${esc(s.toUpperCase())}</span>`;
}

async function loadCampaigns() {
  try {
    const rows = await api("/api/campaigns");

    $("campaignsTable").innerHTML = rows.map(c => `
      <tr>
        <td>
          <b>${esc(c.name)}</b>
          <small>${new Date(c.created_at + "Z").toLocaleString()}</small>
        </td>
        <td>${statusBadge(c.status)}</td>
        <td>${c.total}</td>
        <td class="green">${c.sent}</td>
        <td class="red">${c.failed}</td>
        <td>${Number(c.skipped || 0) > 0 ? `<span class="skipped-pill">${c.skipped}</span>` : "—"}</td>
        <td>${new Date(c.created_at + "Z").toLocaleDateString()}</td>
        <td style="display:flex;gap:5px;flex-wrap:wrap">
          <button class="btn small outline" onclick="openCampaign(${c.id})">View</button>
          <button class="btn small outline" onclick="downloadReport(${c.id})" title="Download full report">⬇ Report</button>
          ${c.failed > 0 ? `<button class="btn small outline failed-dl-btn" onclick="downloadFailed(${c.id})" title="Download failed records">⬇ Failed</button>` : ""}
          ${Number(c.skipped || 0) > 0 ? `<button class="btn small outline skipped-dl-btn" onclick="downloadSkipped(${c.id})" title="Download not-sent contacts">⬇ Not Sent</button>` : ""}
        </td>
      </tr>
    `).join("") || `<tr><td colspan="8" class="empty">No campaigns yet.</td></tr>`;
  } catch (e) {
    toast(e.message, "error");
  }
}
window.loadCampaigns = loadCampaigns;

async function refreshCampaigns() {
  await loadCampaigns();
  await loadDashboard();
  if (currentCampaignId) await openCampaign(currentCampaignId, false);
  toast("Campaign history refreshed.", "success");
}
window.refreshCampaigns = refreshCampaigns;

async function openCampaign(id, scroll = true) {
  try {
    currentCampaignId = id;
    const d = await api(`/api/campaigns/${id}`);

    $("progressPanel").classList.remove("hidden");
    $("progressTitle").textContent = d.campaign.name;
    $("progressMeta").textContent =
      `${d.campaign.sent} sent • ${d.campaign.failed} failed • ${d.campaign.total} total`;

    updateProgress(d.campaign);
    renderRecipients(d.recipients);
    renderCampaignAttachments(d.attachments || []);

    // Show failed download button if any failed
    const failedCount = Number(d.campaign.failed || 0);
    const failedBtn = $("downloadFailedBtn");
    if (failedBtn) failedBtn.classList.toggle("hidden", failedCount === 0);

    // Show skipped stat and download button
    const skippedCount = Number(d.campaign.skipped || 0);
    const skippedStat = $("skippedStat");
    const dlBtn = $("downloadSkippedBtn");
    if (skippedStat) {
      if (skippedCount > 0) {
        skippedStat.classList.remove("hidden");
        skippedStat.innerHTML = `⚠️ <b>${skippedCount} row${skippedCount === 1 ? "" : "s"} skipped</b> — no matching PDF or invalid email. Their emails were <b>not sent</b>.`;
        if (dlBtn) dlBtn.classList.remove("hidden");
      } else {
        skippedStat.classList.add("hidden");
        if (dlBtn) dlBtn.classList.add("hidden");
      }
    }

    if (scroll) $("progressPanel").scrollIntoView({ behavior: "smooth", block: "start" });

    if (d.campaign.status === "sending") pollProgress(id);
  } catch (e) {
    toast(e.message, "error");
  }
}
window.openCampaign = openCampaign;
window.viewCampaign = openCampaign;

function renderCampaignAttachments(items) {
  const box = $("campaignAttachments");
  if (!items.length) { box.innerHTML = ""; return; }
  box.innerHTML = `<div class="attachment-history"><b>📎 Global Attachments:</b> ${items.map(a => `<span>${esc(a.original_name)} <small>${formatBytes(a.size)}</small></span>`).join("")}</div>`;
}

function renderRecipients(recipients) {
  $("recipientsTable").innerHTML = recipients.map((r, i) => {
    const sent = r.status === "sent";
    const notSent = r.status === "not_sent";
    const statusCls = sent ? "completed" : r.status === "failed" ? "failed" : notSent ? "skipped-badge" : "sending";
    const statusLabel = notSent ? "NOT SENT" : r.status.toUpperCase();
    const pdfNote = r.pdf_match_status === "skipped" ? '<span class="pdf-skip-note" title="No matching PDF / invalid email">📄✕</span>' :
                    r.pdf_match_status === "matched" ? '<span class="pdf-skip-note" title="Personal PDF attached">📄✓</span>' : "";

    // Parse To emails for display
    const toStr = r.to_emails || r.email || "";
    const toDisplay = toStr.split(",").map(s => s.trim()).filter(Boolean).join(", ") || r.email || "—";
    const ccStr = r.cc ? r.cc.split(";").filter(Boolean).join(", ") : "—";

    return `
      <tr class="${notSent ? "row-not-sent" : ""}">
        <td>${i + 1}</td>
        <td><b>${esc(r.name)}</b>${pdfNote}</td>
        <td class="to-emails-cell">${esc(toDisplay)}</td>
        <td class="muted">${esc(ccStr)}</td>
        <td><span class="status ${statusCls}">${esc(statusLabel)}</span></td>
        <td class="${r.status === "failed" ? "error-cell" : "muted"}">${notSent ? "Not sent — no matching PDF or invalid email" : esc(r.error || "—")}</td>
        <td>${r.sent_at ? new Date(r.sent_at + "Z").toLocaleString() : "—"}</td>
      </tr>
    `;
  }).join("") || `<tr><td colspan="7" class="empty">No recipient details available.</td></tr>`;
}

function updateProgress(p) {
  const done = (p.sent || 0) + (p.failed || 0);
  const skipped = Number(p.skipped || 0);
  const pct = p.total ? Math.round(done / p.total * 100) : 0;

  $("progressTitle").textContent = p.name;
  const skippedNote = skipped > 0 ? ` · ${skipped} skipped` : "";
  $("progressText").textContent = `${done} / ${p.total} processed · ${p.sent} sent · ${p.failed} failed${skippedNote}`;
  $("progressBar").style.width = `${pct}%`;
}

function downloadFailed(campaignId) {
  if (!campaignId) return toast("No campaign selected.", "error");
  const a = document.createElement("a");
  a.href = `/api/campaigns/${campaignId}/failed/export`;
  a.target = "_blank";
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}
window.downloadFailed = downloadFailed;

function downloadSkipped(campaignId) {
  if (!campaignId) return toast("No campaign selected.", "error");
  const a = document.createElement("a");
  a.href = `/api/campaigns/${campaignId}/skipped/export`;
  a.target = "_blank";
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}
window.downloadSkipped = downloadSkipped;

function downloadReport(campaignId) {
  if (!campaignId) return toast("No campaign selected.", "error");
  const a = document.createElement("a");
  a.href = `/api/campaigns/${campaignId}/report`;
  a.target = "_blank";
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}
window.downloadReport = downloadReport;

async function pollProgress(id) {
  const tick = async () => {
    try {
      const p = await api(`/api/campaigns/${id}/progress`);
      updateProgress(p);

      const d = await api(`/api/campaigns/${id}`);
      renderRecipients(d.recipients);

      if (p.status === "sending") {
        setTimeout(tick, 900);
      } else {
        $("progressMeta").textContent = `${p.sent} sent • ${p.failed} failed • ${p.total} total`;
        await loadCampaigns();
        await loadDashboard();
      }
    } catch {}
  };
  tick();
}

function closeCampaignDetails() {
  $("progressPanel").classList.add("hidden");
}
window.closeCampaignDetails = closeCampaignDetails;

$("senderSelect")?.addEventListener("change", () => updateSendButtonLabel());
$("composeSearchInput")?.addEventListener("input", e => {
  composeSearchFilter = e.target.value;
  renderComposeRecipients();
});

// ── Templates ─────────────────────────────────────────────────────────────────

let savedTemplates = [];

async function loadTemplates() {
  try {
    savedTemplates = await api("/api/templates");
    renderTemplateSelect();
    renderTemplateList();
  } catch {}
}

function renderTemplateSelect() {
  const sel = $("templateSelect");
  if (!sel) return;
  sel.innerHTML = `<option value="">— Load a saved template —</option>` +
    savedTemplates.map(t => `<option value="${t.id}">${esc(t.name)}</option>`).join("");
}

function renderTemplateList() {
  const box = $("templateList");
  if (!box) return;
  if (!savedTemplates.length) {
    box.innerHTML = `<div class="template-empty">No saved templates yet. Write a message and click "Save Current".</div>`;
    return;
  }
  box.innerHTML = savedTemplates.map(t => `
    <div class="template-item">
      <span class="template-name">${esc(t.name)}</span>
      <div class="template-item-actions">
        <button class="btn outline small-btn" onclick="applyTemplate(${t.id})">Load</button>
        <button class="btn outline small-btn danger" onclick="deleteTemplate(${t.id})">✕</button>
      </div>
    </div>
  `).join("");
}

async function loadTemplate() {
  const sel = $("templateSelect");
  if (!sel || !sel.value) return toast("Select a template from the list.", "error");
  applyTemplate(Number(sel.value));
}
window.loadTemplate = loadTemplate;

function applyTemplate(id) {
  const t = savedTemplates.find(x => x.id === id);
  if (!t) return toast("Template not found.", "error");
  $("subject").value = t.subject;
  $("body").value = t.body;
  updatePreview();
  toast(`Template "${t.name}" loaded.`, "success");
}
window.applyTemplate = applyTemplate;

async function saveTemplate() {
  const subject = $("subject").value.trim();
  const body = $("body").value.trim();
  if (!subject || !body) return toast("Write a subject and message first.", "error");

  const name = prompt("Enter a name for this template:");
  if (!name || !name.trim()) return;

  try {
    await api("/api/templates", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: name.trim(), subject, body })
    });
    await loadTemplates();
    toast(`Template "${name.trim()}" saved!`, "success");
  } catch (e) {
    toast(e.message, "error");
  }
}
window.saveTemplate = saveTemplate;

async function deleteTemplate(id) {
  const t = savedTemplates.find(x => x.id === id);
  if (!t) return;
  if (!confirm(`Delete template "${t.name}"?`)) return;
  try {
    await api(`/api/templates/${id}`, { method: "DELETE" });
    await loadTemplates();
    toast("Template deleted.", "success");
  } catch (e) {
    toast(e.message, "error");
  }
}
window.deleteTemplate = deleteTemplate;

// ── Sender Check ──────────────────────────────────────────────────────────────

async function checkCurrentSender() {
  const select = $("senderSelect");
  const senderType = select?.value || "microsoft";
  setActionMessage("Checking sender connection...", "info");

  try {
    if (senderType === "microsoft") {
      const d = await api("/api/microsoft/check", { method: "POST" });
      setActionMessage(`✓ Microsoft Outlook connected as ${d.email} (${d.name || "Verified"})`, "success");
      toast(`Outlook verified: ${d.email}`, "success");
    } else {
      const d = await api("/api/test-smtp", { method: "POST" });
      setActionMessage("✓ " + d.message, "success");
      toast(d.message, "success");
    }
  } catch (e) {
    setActionMessage("✕ " + e.message, "error");
    toast(e.message, "error");
    if (senderType === "microsoft" && e.message.includes("not connected")) {
      setTimeout(() => openMsModal(), 700);
    }
  }
}
window.checkCurrentSender = checkCurrentSender;
window.checkSMTP = checkCurrentSender;

async function sendTest() {
  const to = $("testEmail").value.trim();
  if (!to) return toast("Enter a test email address first.", "error");

  const select = $("senderSelect");
  const senderType = select?.value || "microsoft";
  const senderName = senderType === "microsoft" ? "Microsoft Outlook" : "Custom SMTP";

  setActionMessage(`Sending test email via ${senderName}...`, "info");

  try {
    const d = await api("/api/test-email", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        to,
        subject: $("subject").value,
        body: $("body").value,
        sampleName: contacts[0]?.name || "Recipient",
        attachmentIds: selectedAttachments.map(a => a.id),
        senderType
      })
    });
    setActionMessage("✓ " + d.message, "success");
    toast(d.message, "success");
  } catch (e) {
    setActionMessage("✕ " + e.message, "error");
    toast(e.message, "error");
    if (senderType === "microsoft" && e.message.includes("not connected")) {
      setTimeout(() => openMsModal(), 700);
    }
  }
}
window.sendTest = sendTest;

// ── SMTP Provider / OTP setup ─────────────────────────────────────────────────

const SMTP_PROVIDERS = {
  outlook: { host: "smtp.office365.com", port: "587", secure: "false" },
  gmail:   { host: "smtp.gmail.com",      port: "587", secure: "false" },
  yahoo:   { host: "smtp.mail.yahoo.com", port: "587", secure: "false" },
  other:   { host: "",                    port: "587", secure: "false" }
};
let _activeProvider = "outlook";

function switchSetupTab(tab) {
  $("tabSimple").classList.toggle("active", tab === "simple");
  $("tabOauth").classList.toggle("active", tab === "oauth");
  $("setupSimple").classList.toggle("hidden", tab !== "simple");
  $("setupOauth").classList.toggle("hidden", tab !== "oauth");
}
window.switchSetupTab = switchSetupTab;

function pickProvider(p) {
  _activeProvider = p;
  ["outlook","gmail","yahoo","other"].forEach(id => {
    const b = $("prov" + id.charAt(0).toUpperCase() + id.slice(1));
    if (b) b.classList.toggle("active", id === p);
  });
  const otherHost = $("provOtherHost");
  if (otherHost) otherHost.classList.toggle("hidden", p !== "other");

  const placeholders = { outlook:"yourname@outlook.com", gmail:"yourname@gmail.com", yahoo:"yourname@yahoo.com", other:"yourname@yourdomain.com" };
  const inp = $("quickSmtpUser");
  if (inp) inp.placeholder = placeholders[p] || "your@email.com";

  const guides = {
    outlook: {
      title: "How to get an Outlook App Password:",
      steps: ["Go to <a href='https://account.microsoft.com/security' target='_blank'>account.microsoft.com/security</a>",
              "Turn on <b>Two-step verification</b> (required)",
              "Go to <b>Advanced security options</b> → <b>App passwords</b>",
              "Click <b>Create a new app password</b> → copy it",
              "Paste it in the Password field above"]
    },
    gmail: {
      title: "How to get a Gmail App Password:",
      steps: ["Go to <a href='https://myaccount.google.com/security' target='_blank'>myaccount.google.com/security</a>",
              "Turn on <b>2-Step Verification</b> (required)",
              "Go to <a href='https://myaccount.google.com/apppasswords' target='_blank'>App passwords</a>",
              "Select app: <b>Mail</b>, device: <b>Windows Computer</b>",
              "Copy the 16-character password and paste above"]
    },
    yahoo: {
      title: "How to get a Yahoo App Password:",
      steps: ["Go to <a href='https://login.yahoo.com/account/security' target='_blank'>Yahoo Account Security</a>",
              "Turn on <b>Two-step verification</b>",
              "Click <b>Generate app password</b>",
              "Select <b>Other app</b>, name it SmartMail",
              "Copy the password and paste above"]
    },
    other: { title: "Custom SMTP — enter host above and use your email password.", steps: [] }
  };
  const g = guides[p] || guides.outlook;
  const title = $("appPwdGuideTitle");
  const steps = $("appPwdGuideSteps");
  if (title) title.innerHTML = g.title;
  if (steps) steps.innerHTML = g.steps.map(s => `<li>${s}</li>`).join("");
}
window.pickProvider = pickProvider;

async function quickSmtpSendOtp() {
  const email    = $("quickSmtpUser").value.trim();
  const pass     = $("quickSmtpPass").value.trim();
  const msg      = $("quickSmtpMsg");
  const provider = SMTP_PROVIDERS[_activeProvider] || SMTP_PROVIDERS.outlook;
  const host     = _activeProvider === "other"
    ? ($("smtpCustomHost").value.trim() || "")
    : provider.host;

  if (!email || !email.includes("@")) { msg.textContent = "Please enter a valid email address."; msg.style.color = "#ef4444"; return; }
  if (!pass) { msg.textContent = "Please enter your password."; msg.style.color = "#ef4444"; return; }
  if (_activeProvider === "other" && !host) { msg.textContent = "Please enter the SMTP host."; msg.style.color = "#ef4444"; return; }

  msg.textContent = "Connecting and sending code…"; msg.style.color = "#6b7280";

  try {
    await api("/api/smtp-settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ host, port: provider.port, secure: provider.secure })
    });

    const r = await api("/api/smtp-verify/send-otp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, pass, fromName: "" })
    });

    closeMsModal();
    _smtpOtpEmail = email;
    $("smtpOtpTargetEmail").textContent = email;
    $("smtpOtpCode").value = "";
    $("smtpOtpMsg2").textContent = "";
    $("smtpOtpStep1").classList.add("hidden");
    $("smtpOtpStep2").classList.remove("hidden");
    $("smtpOtpModal").classList.remove("hidden");
    setTimeout(() => $("smtpOtpCode").focus(), 150);
    toast(r.message, "success");
  } catch (e) {
    msg.textContent = "✕ " + e.message;
    msg.style.color = "#ef4444";
  }
}
window.quickSmtpSendOtp = quickSmtpSendOtp;

async function saveMsConfig(e) {
  if (e) e.preventDefault();

  const clientId     = ($("msClientId")?.value     || "").trim();
  const clientSecret = ($("msClientSecret")?.value  || "").trim();
  const tenantId     = ($("msTenantId")?.value      || "common").trim();

  if (!clientId) {
    return toast("Please enter your Application (Client) ID.", "error");
  }

  // Warn if user pasted api:// prefix
  if (clientId.startsWith("api://") || clientId.startsWith("https://")) {
    return toast("Remove api:// from the Client ID — paste only the UUID part.", "error");
  }

  const btn = $("msSubmitBtn");
  if (btn) { btn.disabled = true; btn.textContent = "Saving & redirecting…"; }

  try {
    // Save config to DB
    await api("/api/microsoft/config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ clientId, clientSecret, tenantId })
    });

    toast("Config saved — redirecting to Microsoft login…", "success");

    // Redirect to OAuth flow (server will build the correct redirect URI)
    setTimeout(() => {
      window.location.href = "/auth/microsoft/login";
    }, 500);
  } catch (err) {
    toast("Error saving config: " + err.message, "error");
    if (btn) { btn.disabled = false; btn.textContent = "Sign in with Microsoft ⊞"; }
  }
}

function disconnectMicrosoft() {
  if (!confirm("Disconnect the Microsoft Outlook account?")) return;
  api("/api/microsoft/disconnect", { method: "POST" })
    .then(() => { toast("Microsoft account disconnected.", "success"); closeMsModal(); loadMicrosoftStatus(); })
    .catch(e => toast(e.message, "error"));
}

function copyRedirectUri() {
  const txt = $("redirectUriDisplay") ? $("redirectUriDisplay").textContent : "";
  if (!txt || txt === "loading...") return;
  navigator.clipboard.writeText(txt).then(() => toast("✓ Redirect URI copied!", "success")).catch(() => {});
}
function validateClientId(input) {
  const val = (input.value || "").trim();
  const warn = $("oauthWarning");
  if (!warn) return;
  if (val.startsWith("api://") || val.startsWith("https://")) {
    warn.classList.remove("hidden");
    const uuid = val.replace(/^api:\/\//, "").replace(/^https?:\/\/[^/]+\//, "");
    if (uuid !== val) input.value = uuid;
  } else {
    warn.classList.add("hidden");
  }
}
window.openMsModal = openMsModal;
window.closeMsModal = closeMsModal;
window.saveMsConfig = saveMsConfig;
window.disconnectMicrosoft = disconnectMicrosoft;
window.checkCurrentSender = checkCurrentSender;
window.copyRedirectUri = copyRedirectUri;
window.validateClientId = validateClientId;

// ── Switch Mail Modal ─────────────────────────────────────────────────────────
let _smtpOtpEmail = "";

function openSmtpOtpModal() {
  $("smtpOtpStep1").classList.remove("hidden");
  $("smtpOtpStep2").classList.add("hidden");
  $("smtpOtpMsg1").textContent = "";
  $("smtpOtpMsg2").textContent = "";
  $("smtpUser").value = "";
  $("smtpOtpCode") && ($("smtpOtpCode").value = "");
  $("smtpOtpModal").classList.remove("hidden");
  setTimeout(() => $("smtpUser").focus(), 120);
}
window.openSmtpOtpModal = openSmtpOtpModal;

function closeSmtpOtpModal() {
  $("smtpOtpModal").classList.add("hidden");
}
window.closeSmtpOtpModal = closeSmtpOtpModal;

function smtpOtpBack() {
  $("smtpOtpStep1").classList.remove("hidden");
  $("smtpOtpStep2").classList.add("hidden");
  $("smtpOtpMsg1").textContent = "";
}
window.smtpOtpBack = smtpOtpBack;

async function smtpSendOtp() {
  const email    = $("smtpUser").value.trim();
  const pass     = $("smtpPass").value.trim();
  const fromName = $("smtpFromName").value.trim();
  const msg1 = $("smtpOtpMsg1");

  if (!email || !email.includes("@")) { msg1.textContent = "Please enter a valid email address."; msg1.style.color = "#ef4444"; return; }
  if (!pass) { msg1.textContent = "Please enter your App Password."; msg1.style.color = "#ef4444"; return; }

  msg1.textContent = "Connecting to Outlook and sending code…"; msg1.style.color = "#6b7280";

  try {
    const r = await api("/api/smtp-verify/send-otp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, pass, fromName })
    });
    _smtpOtpEmail = email;
    $("smtpOtpTargetEmail").textContent = email;
    $("smtpOtpCode").value = "";
    $("smtpOtpMsg2").textContent = "";
    $("smtpOtpStep1").classList.add("hidden");
    $("smtpOtpStep2").classList.remove("hidden");
    setTimeout(() => $("smtpOtpCode").focus(), 120);
    toast(r.message, "success");
  } catch (e) {
    msg1.textContent = "✕ " + e.message;
    msg1.style.color = "#ef4444";
  }
}
window.smtpSendOtp = smtpSendOtp;

async function smtpConfirmOtp() {
  const otp  = $("smtpOtpCode").value.trim();
  const msg2 = $("smtpOtpMsg2");

  if (!otp || otp.length < 4) { msg2.textContent = "Please enter the 6-digit code from your inbox."; msg2.style.color = "#ef4444"; return; }

  msg2.textContent = "Verifying…"; msg2.style.color = "#6b7280";

  try {
    const r = await api("/api/smtp-verify/confirm-otp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email: _smtpOtpEmail, otp })
    });

    const tag = $("activeSenderTag");
    if (tag) {
      tag.textContent = _smtpOtpEmail;
      tag.classList.remove("hidden");
    }

    toast(r.message, "success");
    closeSmtpOtpModal();
  } catch (e) {
    msg2.textContent = "✕ " + e.message;
    msg2.style.color = "#ef4444";
  }
}
window.smtpConfirmOtp = smtpConfirmOtp;

boot();
