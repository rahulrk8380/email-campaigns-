let contacts = [];
let currentCampaignId = null;
let selectedAttachments = [];
let selectedContactIds = new Set();
let composeSearchFilter = "";
let msStatus = { connected: false, email: null, name: null, clientIdConfigured: false };
let currentPdfFolderId = null;
let pdfFolderFileNames = [];
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
  // Check for Microsoft OAuth redirect query params
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
    $("composeRecipientsSubtitle").textContent = `${count} of ${contacts.length} recipients selected`;
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
  btn.innerHTML = `Send to ${count} recipient${count === 1 ? "" : "s"} via ${senderName} <span>→</span>`;
}

function renderContacts() {
  $("contactCount").textContent = `${contacts.length} contacts`;
  $("contactsTable").innerHTML = contacts.map((c, i) => {
    const isChecked = selectedContactIds.has(c.id);
    return `
      <tr>
        <td><input class="contact-check" type="checkbox" value="${c.id}" ${isChecked ? "checked" : ""}></td>
        <td>${i + 1}</td>
        <td><b>${esc(c.name)}</b></td>
        <td>${esc(c.email)}</td>
        <td>${esc(c.number || "—")}</td>
        <td>${esc(c.branch || c.Branch || "—")}</td>
        <td>${esc(c.details || "—")}</td>
        <td><span class="status ready">Ready</span></td>
      </tr>
    `;
  }).join("") || `<tr><td colspan="8" class="empty">No contacts loaded. Upload an Excel file above.</td></tr>`;

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
    return `
      <tr>
        <td><input class="compose-contact-check" type="checkbox" data-id="${c.id}" ${isChecked ? "checked" : ""}></td>
        <td>${i + 1}</td>
        <td><b>${esc(c.name)}</b></td>
        <td>${esc(c.email)}</td>
        <td>${esc(c.branch || c.Branch || "—")}</td>
        <td><span class="status ${isChecked ? "ready" : "draft"}">${isChecked ? "Ready" : "Excluded"}</span></td>
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
    selectedContactIds = new Set(contacts.map(c => c.id));
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
    selectedContactIds = new Set(contacts.map(c => c.id));
    renderContacts();
    $("uploadResult").innerHTML =
      `<div class="success-line">✓ ${d.added} valid contacts loaded. ${d.invalid} rows skipped (missing/invalid/duplicate email).</div>`;
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
  renderContacts();
  loadDashboard();
  toast("Contact list cleared.", "success");
}
window.clearContacts = clearContacts;

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

// ── PDF Folder (Personal PDFs per recipient) ──────────────────────────────────

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

  // Clear any previous folder first
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

// Wire up all 3 PDF upload modes
const pdfFolderInput = $("pdfFolderInput");
const pdfFolderDirInput = $("pdfFolderDirInput");
const pdfFolderZipInput = $("pdfFolderZipInput");
const pdfFolderDropzone = $("pdfFolderDropzone");

if (pdfFolderInput) {
  pdfFolderInput.addEventListener("change", () => {
    if (pdfFolderInput.files.length) uploadPdfFolder(pdfFolderInput.files);
    pdfFolderInput.value = "";
  });
}
if (pdfFolderDirInput) {
  pdfFolderDirInput.addEventListener("change", () => {
    if (pdfFolderDirInput.files.length) uploadPdfFolder(pdfFolderDirInput.files);
    pdfFolderDirInput.value = "";
  });
}
if (pdfFolderZipInput) {
  pdfFolderZipInput.addEventListener("change", () => {
    if (pdfFolderZipInput.files.length) uploadPdfFolder(pdfFolderZipInput.files);
    pdfFolderZipInput.value = "";
  });
}
if (pdfFolderDropzone) {
  ["dragenter", "dragover"].forEach(ev => pdfFolderDropzone.addEventListener(ev, e => {
    e.preventDefault();
    pdfFolderDropzone.classList.add("drag");
  }));
  ["dragleave", "drop"].forEach(ev => pdfFolderDropzone.addEventListener(ev, e => {
    e.preventDefault();
    pdfFolderDropzone.classList.remove("drag");
  }));
  pdfFolderDropzone.addEventListener("drop", e => {
    if (e.dataTransfer.files.length) uploadPdfFolder(e.dataTransfer.files);
  });
}


// ── Send Timer ────────────────────────────────────────────────────────────────
let delayUnit = "sec"; // "sec" or "min"

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
  // Sync slider (clamp to slider max 300)
  const slider = $("delaySlider");
  if (slider) slider.value = Math.min(300, secs);
}

// Wire slider ↔ input
const delaySlider = $("delaySlider");
const delayInput = $("delayInput");
if (delaySlider && delayInput) {
  delaySlider.addEventListener("input", () => {
    delayInput.value = delaySlider.value;
    updateDelayPreview();
  });
  delayInput.addEventListener("input", () => {
    delaySlider.value = Math.min(300, Number(delayInput.value) || 1);
    updateDelayPreview();
  });
}

// ── Dynamic Extra Columns from contacts ──────────────────────────────────────
function renderDynamicColumnChips() {
  const extraKeys = new Set();
  contacts.forEach(c => {
    try {
      const extras = JSON.parse(c.extra_json || "{}");
      Object.keys(extras).forEach(k => extraKeys.add(k));
    } catch {}
  });

  const bar = document.querySelector(".variable-bar");
  if (!bar) return;

  // Remove any previously added dynamic chips
  bar.querySelectorAll(".dynamic-var-chip").forEach(el => el.remove());

  extraKeys.forEach(key => {
    const varName = `{{${key}}}`;
    const btn = document.createElement("button");
    btn.textContent = varName;
    btn.className = "dynamic-var-chip";
    btn.title = `Insert ${varName} (from your Excel)`;
    btn.onclick = () => insertVar(varName);
    bar.appendChild(btn);
  });
}

// ── Download Full Report ──────────────────────────────────────────────────────
function downloadReport(campaignId) {
  if (!campaignId) return toast("No campaign selected.", "error");
  // Open in new tab so session stays intact in the main tab
  const a = document.createElement("a");
  a.href = `/api/campaigns/${campaignId}/report`;
  a.target = "_blank";
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}
window.downloadReport = downloadReport;

// ── Schedule Send Helpers ─────────────────────────────────────────────────────
function toggleSchedule() {
  const on = $("scheduleToggle") && $("scheduleToggle").checked;
  const box = $("scheduleBox");
  if (box) box.classList.toggle("hidden", !on);
  if (on && $("scheduledAt")) {
    // Default to 1 hour from now
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
attachmentInput.addEventListener("change", () => {
  uploadAttachments(attachmentInput.files);
  attachmentInput.value = "";
});
["dragenter", "dragover"].forEach(ev => attachmentDropzone.addEventListener(ev, e => {
  e.preventDefault();
  attachmentDropzone.classList.add("drag");
}));
["dragleave", "drop"].forEach(ev => attachmentDropzone.addEventListener(ev, e => {
  e.preventDefault();
  attachmentDropzone.classList.remove("drag");
}));
attachmentDropzone.addEventListener("click", () => attachmentInput.click());
attachmentDropzone.addEventListener("drop", e => uploadAttachments(e.dataTransfer.files));
renderAttachments();

function updatePreview() {
  const subject = $("subject").value || "Your subject";
  const body = $("body").value || "Your personalized message will appear here.";

  $("previewSubject").textContent = subject
    .replaceAll("{{name}}", "Arun")
    .replaceAll("{{email}}", "arun@example.com");

  $("previewBody").textContent = body
    .replaceAll("{{name}}", "Arun")
    .replaceAll("{{email}}", "arun@example.com")
    .replaceAll("{{number}}", "9876543210")
    .replaceAll("{{details}}", "Sample details");
}

$("subject").addEventListener("input", updatePreview);
$("body").addEventListener("input", updatePreview);

function insertVar(v) {
  const t = $("body");
  const a = t.selectionStart;
  const b = t.selectionEnd;
  t.value = t.value.slice(0, a) + v + t.value.slice(b);
  t.focus();
  t.selectionStart = t.selectionEnd = a + v.length;
  updatePreview();
}
window.insertVar = insertVar;

function setActionMessage(message, type = "info") {
  $("actionMessage").className = `action-message ${type}`;
  $("actionMessage").textContent = message;
}

// Microsoft Status & Configuration
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

function openMsModal() {
  $("msModal").classList.remove("hidden");
  loadMicrosoftStatus();
}
window.openMsModal = openMsModal;

function closeMsModal() {
  $("msModal").classList.add("hidden");
}
window.closeMsModal = closeMsModal;

function copyRedirectUri() {
  const uri = $("redirectUriDisplay").textContent.trim();
  navigator.clipboard.writeText(uri);
  toast("Redirect URI copied to clipboard!", "success");
}
window.copyRedirectUri = copyRedirectUri;

async function saveMsConfig(e) {
  e.preventDefault();
  const clientId = $("msClientId").value.trim();
  const clientSecret = $("msClientSecret").value.trim();
  const tenantId = $("msTenantId").value.trim() || "common";

  if (!clientId && !msStatus.clientIdConfigured) {
    return toast("Application (Client) ID is required.", "error");
  }

  try {
    await api("/api/microsoft/config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ clientId, clientSecret, tenantId })
    });

    toast("Redirecting to Microsoft Sign-in...", "info");
    window.location.href = "/auth/microsoft/login";
  } catch (err) {
    toast(err.message, "error");
  }
}
window.saveMsConfig = saveMsConfig;

async function disconnectMicrosoft() {
  if (!confirm("Disconnect your Microsoft Outlook account?")) return;
  try {
    await api("/api/microsoft/disconnect", { method: "POST" });
    toast("Microsoft Outlook disconnected.", "success");
    await loadMicrosoftStatus();
    await loadSenders();
  } catch (err) {
    toast(err.message, "error");
  }
}
window.disconnectMicrosoft = disconnectMicrosoft;

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

async function sendCampaignDirect() {
  const ids = selectedIds();

  if (!ids.length) return toast("Select at least one recipient.", "error");
  if (!$("subject").value.trim()) return toast("Enter a subject.", "error");
  if (!$("body").value.trim()) return toast("Enter a message body.", "error");

  const select = $("senderSelect");
  const senderType = select?.value || "microsoft";
  const senderLabel = senderType === "microsoft" ? "Microsoft Outlook" : "Custom SMTP";

  // Warn if PDF folder is active — show how many will be skipped
  if (currentPdfFolderId) {
    const matchable = contacts.filter(c => ids.includes(c.id));
    if (!confirm(`Send this campaign to ${ids.length} selected recipient${ids.length === 1 ? "" : "s"} via ${senderLabel}?\n\n📂 Personal PDFs are active (${pdfFolderFileNames.length} PDFs uploaded).\nContacts without a matching PDF will be SKIPPED and not sent.\n\nProceed?`)) return;
  } else {
    if (!confirm(`Send this campaign to ${ids.length} selected recipient${ids.length === 1 ? "" : "s"} via ${senderLabel} now?`)) return;
  }

  try {
    const d = await api("/api/campaigns", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        name: $("campaignName").value.trim() || "Email Campaign",
        subject: $("subject").value,
        body: $("body").value,
        selectedIds: ids,
        attachmentIds: selectedAttachments.map(a => a.id),
        senderType,
        pdfFolderId: currentPdfFolderId || null,
        delaySeconds: getDelaySeconds(),
        scheduledAt: getScheduledAt()
      })
    });

    // Show matched/skipped breakdown if PDF folder was used
    if (d.usedPdfFolder) {
      if (d.skippedCount > 0) {
        toast(`📂 ${d.matchedCount} will be sent · ${d.skippedCount} skipped (no matching PDF found)`, "info");
      } else {
        toast(`📂 All ${d.matchedCount} contacts matched!`, "success");
      }
    }

    currentCampaignId = d.campaignId;

    // Scheduled — don't send now, show confirmation
    if (d.scheduled) {
      const when = new Date(d.scheduledAt).toLocaleString();
      toast(`📅 Campaign scheduled! Will send automatically at ${when}`, "success");
      showPage("campaigns");
      await loadCampaigns();
      return;
    }

    // Check for duplicate loan numbers — show warning modal before sending
    if (d.duplicates && d.duplicates.length > 0) {
      showDuplicateModal(d.duplicates, d.campaignId, senderLabel);
    } else {
      await proceedToSend(d.campaignId, senderLabel);
    }

  } catch (e) {
    toast(e.message, "error");
    if (senderType === "microsoft" && e.message.includes("not connected")) {
      setTimeout(() => openMsModal(), 700);
    }
  }
}
window.sendCampaignDirect = sendCampaignDirect;

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
    // Collect all recipientIds from the 2nd contact onwards (keep first, skip rest)
    const toSkip = [];
    for (const d of _pendingDuplicates) {
      // Skip all but the first recipient for each duplicate group
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
          <button class="btn small outline" onclick="downloadReport(${c.id})" title="Download full report (Sent + Undeliverable)">⬇ Report</button>
          ${Number(c.skipped || 0) > 0 ? `<button class="btn small outline skipped-dl-btn" onclick="downloadSkipped(${c.id})" title="Download not-sent contacts as Excel">⬇ Not Sent</button>` : ""}
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

    // Show skipped stat and download button if any were skipped
    const skippedCount = Number(d.campaign.skipped || 0);
    const skippedStat = $("skippedStat");
    const dlBtn = $("downloadSkippedBtn");
    if (skippedStat) {
      if (skippedCount > 0) {
        skippedStat.classList.remove("hidden");
        skippedStat.innerHTML = `⚠️ <b>${skippedCount} contact${skippedCount === 1 ? "" : "s"} skipped</b> — no matching PDF found. Their emails were <b>not sent</b>.`;
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
    const pdfNote = r.pdf_match_status === "skipped" ? '<span class="pdf-skip-note" title="No matching PDF found">📄✕</span>' :
                    r.pdf_match_status === "matched" ? '<span class="pdf-skip-note" title="Personal PDF attached">📄✓</span>' : "";
    return `
      <tr class="${notSent ? "row-not-sent" : ""}">
        <td>${i + 1}</td>
        <td><b>${esc(r.name)}</b>${pdfNote}</td>
        <td>${esc(r.email)}</td>
        <td><span class="status ${statusCls}">${esc(statusLabel)}</span></td>
        <td class="${r.status === "failed" ? "error-cell" : "muted"}">${notSent ? "No matching PDF in folder" : esc(r.error || "—")}</td>
        <td>${r.sent_at ? new Date(r.sent_at + "Z").toLocaleString() : "—"}</td>
      </tr>
    `;
  }).join("") || `<tr><td colspan="6" class="empty">No recipient details available.</td></tr>`;
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

// ── Simple Email Provider Setup (replaces Azure OAuth) ───────────────────────
const SMTP_PROVIDERS = {
  outlook: { host: "smtp.office365.com", port: "587", secure: "false" },
  gmail:   { host: "smtp.gmail.com",      port: "587", secure: "false" },
  yahoo:   { host: "smtp.mail.yahoo.com", port: "587", secure: "false" },
  other:   { host: "",                    port: "587", secure: "false" }
};
let _activeProvider = "outlook";

// Tab switcher for Simple / OAuth
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

  // Update email placeholder
  const placeholders = { outlook:"yourname@outlook.com", gmail:"yourname@gmail.com", yahoo:"yourname@yahoo.com", other:"yourname@yourdomain.com" };
  const inp = $("quickSmtpUser");
  if (inp) inp.placeholder = placeholders[p] || "your@email.com";

  // Update App Password guide
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
    other: {
      title: "Custom SMTP — enter host above and use your email password.",
      steps: []
    }
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

  if (!email || !email.includes("@")) {
    msg.textContent = "Please enter a valid email address.";
    msg.style.color = "#ef4444"; return;
  }
  if (!pass) {
    msg.textContent = "Please enter your password.";
    msg.style.color = "#ef4444"; return;
  }
  if (_activeProvider === "other" && !host) {
    msg.textContent = "Please enter the SMTP host.";
    msg.style.color = "#ef4444"; return;
  }

  msg.textContent = "Connecting and sending code…";
  msg.style.color = "#6b7280";

  try {
    // Save SMTP host/port/secure first
    await api("/api/smtp-settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ host, port: provider.port, secure: provider.secure })
    });

    // Then fire OTP using those credentials
    const r = await api("/api/smtp-verify/send-otp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, pass, fromName: "" })
    });

    // Close this modal, open OTP code entry step
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

// Stub old MS functions so no JS errors
function openMsModal() {
  $("msModal").classList.remove("hidden");
  $("quickSmtpMsg") && ($("quickSmtpMsg").textContent = "");
  // Auto-fill the correct redirect URI based on current URL
  const base = window.location.origin;
  const redirectUri = base + "/auth/microsoft/callback";
  const el = $("redirectUriDisplay");
  if (el) el.textContent = redirectUri;
}
function closeMsModal() {
  $("msModal").classList.add("hidden");
  $("quickSmtpMsg") && ($("quickSmtpMsg").textContent = "");
}
function saveMsConfig(e) { if(e) e.preventDefault(); }
function disconnectMicrosoft() {}
function checkCurrentSender() {}
function copyRedirectUri() {
  const txt = $("redirectUriDisplay") ? $("redirectUriDisplay").textContent : "";
  if (!txt || txt === "loading...") return;
  navigator.clipboard.writeText(txt).then(() => toast("✓ Redirect URI copied!", "success")).catch(() => {});
}
function validateClientId(input) {
  const val = (input.value || "").trim();
  const warn = $("oauthWarning");
  if (!warn) return;
  // Show warning if value starts with api:// or looks like an App ID URI
  if (val.startsWith("api://") || val.startsWith("https://")) {
    warn.classList.remove("hidden");
    // Auto-fix: strip the api:// prefix
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

// ── Switch Mail (topbar button) ───────────────────────────────────────────────
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

// Step 1 — send OTP using the entered credentials to test them
async function smtpSendOtp() {
  const email    = $("smtpUser").value.trim();
  const pass     = $("smtpPass").value.trim();
  const fromName = $("smtpFromName").value.trim();

  const msg1 = $("smtpOtpMsg1");

  if (!email || !email.includes("@")) {
    msg1.textContent = "Please enter a valid email address.";
    msg1.style.color = "#ef4444"; return;
  }
  if (!pass) {
    msg1.textContent = "Please enter your App Password.";
    msg1.style.color = "#ef4444"; return;
  }

  msg1.textContent = "Connecting to Outlook and sending code…";
  msg1.style.color = "#6b7280";

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

// Step 2 — verify code → save new SMTP → update topbar tag
async function smtpConfirmOtp() {
  const otp  = $("smtpOtpCode").value.trim();
  const msg2 = $("smtpOtpMsg2");

  if (!otp || otp.length < 4) {
    msg2.textContent = "Please enter the 6-digit code from your inbox.";
    msg2.style.color = "#ef4444"; return;
  }

  msg2.textContent = "Verifying…";
  msg2.style.color = "#6b7280";

  try {
    const r = await api("/api/smtp-verify/confirm-otp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email: _smtpOtpEmail, otp })
    });

    // Update "Switch Mail" topbar tag to show active sender
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
