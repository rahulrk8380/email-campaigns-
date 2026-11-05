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
// Body editor (native contenteditable modal approach)
let batches = []; // array of batch objects for custom batch scheduling
let scheduleMode = 'all'; // 'all' | 'batch'
let _campaignPollTimer = null;

// -- Native Body Editor (Modal + contenteditable) ---------------------------
function openBodyEditor() {
  renderBodyEditorVarChips();
  const content = document.getElementById('body').value || '';
  document.getElementById('bodyContentEditor').innerHTML = content;
  document.getElementById('bodyEditorModal').classList.remove('hidden');
  document.getElementById('bodyContentEditor').focus();
  syncBodyFromEditor();
}
window.openBodyEditor = openBodyEditor;

function closeBodyEditor() {
  syncBodyFromEditor();
  document.getElementById('bodyEditorModal').classList.add('hidden');
  updateBodyPreview();
  updatePreview();
}
window.closeBodyEditor = closeBodyEditor;

function clearBodyEditor() {
  document.getElementById('bodyContentEditor').innerHTML = '';
  syncBodyFromEditor();
  updateBodyPreview();
}
window.clearBodyEditor = clearBodyEditor;

function syncBodyFromEditor() {
  const el = document.getElementById('bodyContentEditor');
  const bodyEl = document.getElementById('body');
  if (el && bodyEl) {
    bodyEl.value = el.innerHTML;
    const count = (el.textContent || '').length;
    const countEl = document.getElementById('bodyEditorCharCount');
    const countEl2 = document.getElementById('bodyEditorCharCount2');
    if (countEl) countEl.textContent = count + ' characters';
    if (countEl2) countEl2.textContent = count + ' characters';
  }
}
window.syncBodyFromEditor = syncBodyFromEditor;

function fmtExec(cmd, val) {
  document.getElementById('bodyContentEditor').focus();
  document.execCommand(cmd, false, val || null);
  syncBodyFromEditor();
}
window.fmtExec = fmtExec;

function fmtInsertLink() {
  const url = prompt('Enter URL (e.g. https://example.com):');
  if (!url) return;
  document.getElementById('bodyContentEditor').focus();
  document.execCommand('createLink', false, url);
  syncBodyFromEditor();
}
window.fmtInsertLink = fmtInsertLink;

function fmtInsertImage() {
  const input = document.createElement('input');
  input.type = 'file'; input.accept = 'image/*'; input.click();
  input.onchange = () => {
    const file = input.files[0]; if (!file) return;
    if (file.size > 5 * 1024 * 1024) return toast('Image too large (max 5MB)', 'error');
    const reader = new FileReader();
    reader.onload = () => {
      document.getElementById('bodyContentEditor').focus();
      document.execCommand('insertImage', false, reader.result);
      syncBodyFromEditor();
    };
    reader.readAsDataURL(file);
  };
}
window.fmtInsertImage = fmtInsertImage;

function handleEditorPaste(e) {
  e.preventDefault();
  const html = (e.clipboardData || window.clipboardData).getData('text/html');
  const text = (e.clipboardData || window.clipboardData).getData('text/plain');
  document.execCommand('insertHTML', false, html || text.replace(/\r?\n/g, '<br>'));
  syncBodyFromEditor();
}
window.handleEditorPaste = handleEditorPaste;

function updateBodyPreview() {
  const body = document.getElementById('body').value || '';
  const previewEl = document.getElementById('bodyPreviewInner');
  if (!previewEl) return;
  if (!body || body.trim() === '' || body.trim() === '<br>') {
    previewEl.innerHTML = '<span class="body-click-hint">📝 Click here to write your email message...<br><small>Bold · Italic · Color · Lists · Variables · Images supported</small></span>';
  } else {
    const stripped = (body.replace(/<[^>]*>/g, '') || '').trim().substring(0, 150);
    previewEl.innerHTML = '<span class="body-preview-text">' + (stripped || '(empty)') + (body.length > 150 ? '…' : '') + '</span><span class="body-edit-hint">✎ Click to edit</span>';
  }
}
window.updateBodyPreview = updateBodyPreview;

function renderBodyEditorVarChips() {
  const container = document.getElementById('bodyEditorVarChips');
  if (!container) return;
  const headers = excelHeaders.length ? excelHeaders : [];
  if (!headers.length) {
    container.innerHTML = '<span style="font-size:11px;color:#64748b">Upload Excel to see variables</span>';
    return;
  }
  container.innerHTML = headers.filter(function(h){ return !String(h).startsWith('__'); }).map(function(h){
    var v = '{{' + h + '}}'; return '<button class="body-var-chip" onclick="insertVarAtCursor(\'' + v + '\')">"' + v + '</button>';
  }).join('');
}
window.renderBodyEditorVarChips = renderBodyEditorVarChips;

function insertVarAtCursor(varStr) {
  const editor = document.getElementById('bodyContentEditor');
  if (!editor) return;
  editor.focus();
  document.execCommand('insertText', false, varStr);
  syncBodyFromEditor();
  toast('Inserted ' + varStr, 'info');
}
window.insertVarAtCursor = insertVarAtCursor;

function getEmailBody() {
  const el = document.getElementById('body');
  return el ? el.value : '';
}
window.getEmailBody = getEmailBody;
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

// -- Global Dialog � replaces browser confirm() / alert() ---------------------
function showDialog({ icon = "??", title = "Confirm", message = "", okText = "OK", cancelText = null, danger = false } = {}) {
  return new Promise(resolve => {
    const dlg      = $("globalDialog");
    const titleEl  = $("globalDialogTitle");
    const msgEl    = $("globalDialogMsg");
    const iconEl   = $("globalDialogIcon");
    const okBtn    = $("globalDialogOk");
    const cancelBtn= $("globalDialogCancel");

    iconEl.textContent  = icon;
    titleEl.textContent = title;
    msgEl.textContent   = message;
    okBtn.textContent   = okText;
    okBtn.className     = `btn ${danger ? "danger" : "primary"}`;

    if (cancelText) {
      cancelBtn.textContent = cancelText;
      cancelBtn.style.display = "";
    } else {
      cancelBtn.style.display = "none";
    }

    dlg.classList.remove("hidden");

    const cleanup = (result) => {
      dlg.classList.add("hidden");
      okBtn.onclick     = null;
      cancelBtn.onclick = null;
      resolve(result);
    };

    okBtn.onclick     = () => cleanup(true);
    cancelBtn.onclick = () => cleanup(false);
  });
}

// Shorthand: confirmation dialog (returns true/false)
function showConfirm(message, { title = "Confirm", icon = "??", okText = "Yes", cancelText = "Cancel", danger = false } = {}) {
  return showDialog({ icon, title, message, okText, cancelText, danger });
}

// Shorthand: info/alert dialog (returns true when dismissed)
function showAlert(message, { title = "Info", icon = "??", okText = "OK" } = {}) {
  return showDialog({ icon, title, message, okText, cancelText: null });
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
    } else if (err === "admin_consent_required") {
      showAlert(
        "Your organization requires administrator approval before SmartMail Pro can access this mailbox.\n\nPlease contact your Microsoft 365 administrator and ask them to grant consent for SmartMail Pro.",
        { title: "?? Admin Approval Required", icon: "??" }
      );
    } else if (err.toLowerCase().includes("cancel")) {
      toast("Microsoft login was cancelled. Click Connect Outlook to try again.", "info");
    } else if (err.toLowerCase().includes("expired") || err.toLowerCase().includes("refresh")) {
      toast("Your Microsoft connection has expired. Please reconnect your account.", "error");
      setTimeout(() => openMsModal(), 700);
    } else if (err.toLowerCase().includes("configuration") || err.toLowerCase().includes("invalid_client")) {
      toast("Microsoft authentication configuration error. Please contact the administrator.", "error");
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
  const ok = await showConfirm("Are you sure you want to logout?", { title: "Logout", icon: "??", okText: "Logout", cancelText: "Stay" });
  if (!ok) return;
  try { await api("/api/logout", { method: "POST" }); } catch {}
  window.location.reload();
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
  btn.innerHTML = `Send ${count} email${count === 1 ? "" : "s"} via ${senderName} <span>?</span>`;
}

// -- Contacts Table -------------------------------------------------------------

function getContactEmails(c) {
  let extras = {};
  try { extras = JSON.parse(c.extra_json || "{}"); } catch {}
  let emailArr = [];
  try { emailArr = JSON.parse(extras["__emails"] || "[]"); } catch {}
  if (!emailArr.length && c.email) emailArr = [c.email];
  return emailArr;
}

// Returns array of CC email addresses (from CC1-CC10 or legacy CC column)
function getContactCC(c) {
  let extras = {};
  try { extras = JSON.parse(c.extra_json || "{}"); } catch {}
  const raw = String(extras["__cc"] || "");
  if (!raw) return [];
  return raw.split(";").map(s => s.trim()).filter(Boolean);
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
    const ccArr  = getContactCC(c);
    const invalid = getContactInvalid(c);

    const primaryEmail = emails[0] || "�";
    const extraCount   = emails.length - 1;
    const extraBadge   = extraCount > 0
      ? `<span class="email-more-badge">+${extraCount} more</span>` : "";

    const statusCls   = invalid ? "invalid-row-status" : "ready";
    const statusLabel = invalid ? "? No Email" : "Ready";

    // CC display: first CC + "+N more" badge
    let ccDisplay = "�";
    if (ccArr.length > 0) {
      const ccExtra = ccArr.length - 1;
      const ccBadge = ccExtra > 0 ? `<span class="email-more-badge">+${ccExtra} more</span>` : "";
      ccDisplay = esc(ccArr[0]) + ccBadge;
    }

    return `
      <tr class="${invalid ? "row-invalid" : ""}">
        <td><input class="contact-check" type="checkbox" value="${c.id}" ${isChecked ? "checked" : ""}></td>
        <td>${i + 1}</td>
        <td><b>${esc(c.name)}</b></td>
        <td>${esc(primaryEmail)}${extraBadge}</td>
        <td>${extraCount > 0 ? emails.slice(1).map(esc).join("<br>") : "�"}</td>
        <td>${ccDisplay}</td>
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
    const emails    = getContactEmails(c);
    const ccArr     = getContactCC(c);
    const invalid   = getContactInvalid(c);

    const primaryEmail = emails[0] || "�";
    const extraCount   = emails.length - 1;
    const extraBadge   = extraCount > 0 ? `<span class="email-more-badge">+${extraCount}</span>` : "";

    // CC: show first + "+N more"
    let ccDisplay = "�";
    if (ccArr.length > 0) {
      const ccExtra = ccArr.length - 1;
      const ccBadge = ccExtra > 0 ? `<span class="email-more-badge">+${ccExtra}</span>` : "";
      ccDisplay = esc(ccArr[0]) + ccBadge;
    }

    const statusCls   = invalid ? "draft" : isChecked ? "ready" : "draft";
    const statusLabel = invalid ? "Invalid" : isChecked ? "Ready" : "Excluded";

    return `
      <tr class="${invalid ? "row-invalid" : ""}">
        <td><input class="compose-contact-check" type="checkbox" data-id="${c.id}" ${isChecked ? "checked" : ""}></td>
        <td>${i + 1}</td>
        <td><b>${esc(c.name)}</b></td>
        <td>${esc(primaryEmail)}${extraBadge}</td>
        <td>${ccDisplay}</td>
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
      ? ` <span style="color:var(--red)">${d.invalid} row${d.invalid === 1 ? "" : "s"} have no valid email (shown in red � excluded from send).</span>`
      : "";

    $("uploadResult").innerHTML =
      `<div class="success-line">? ${d.added} rows with valid emails loaded. ${d.total} total rows processed.${invalidMsg}</div>`;
    loadDashboard();
    toast("Contacts imported successfully.", "success");
  } catch (e) {
    $("uploadResult").innerHTML = `<div class="error-line">${esc(e.message)}</div>`;
    toast(e.message, "error");
  }
}

async function clearContacts() {
  const ok = await showConfirm("Clear the current contact list? This cannot be undone.", { title: "Clear Contacts", icon: "???", okText: "Clear", danger: true });
  if (!ok) return;
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

// -- Available Variables Panel -------------------------------------------------

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
  const target = getVarInsertTarget();
  if (target === 'body') {
    insertVarAtCursor(varStr);
    toast(`Inserted ${varStr} into Body`, 'info');
    return;
  }
  const field = $(target);
  if (!field) return;
  field.focus();
  const a = field.selectionStart ?? field.value.length;
  const b = field.selectionEnd ?? a;
  field.value = field.value.slice(0, a) + varStr + field.value.slice(b);
  field.selectionStart = field.selectionEnd = a + varStr.length;
  field.dispatchEvent(new Event("input"));
  toast(`Inserted ${varStr} into Subject`, 'info');
}
window.insertVarFromPanel = insertVarFromPanel;

// -- Dynamic Column Chips (in compose variable-bar) ---------------------------

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

// -- Variable Insert Target (Subject or Body) ----------------------------------

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
  if (target === 'body') {
    insertVarAtCursor(v);
    updatePreview();
    return;
  }
  const t = $(target) || $('body');
  if (!t) return;
  const a = t.selectionStart ?? t.value.length;
  const b = t.selectionEnd ?? a;
  t.value = t.value.slice(0, a) + v + t.value.slice(b);
  t.focus();
  t.selectionStart = t.selectionEnd = a + v.length;
  updatePreview();
}
window.insertVar = insertVar;

// -- Format helpers -------------------------------------------------------------

function formatBytes(bytes) {
  const n = Number(bytes || 0);
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

function fileIcon(type, name) {
  if (String(type).startsWith("image/")) return "???";
  const ext = String(name).split(".").pop().toLowerCase();
  if (ext === "pdf") return "??";
  if (["doc","docx"].includes(ext)) return "??";
  if (["xls","xlsx","csv"].includes(ext)) return "??";
  if (["ppt","pptx"].includes(ext)) return "??";
  if (ext === "zip") return "???";
  return "??";
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
      <button class="attachment-remove" title="Remove" onclick="removeAttachment(${a.id})">�</button>
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
    setActionMessage("? " + e.message, "error");
    toast(e.message, "error");
  }
}

// -- PDF Folder ----------------------------------------------------------------

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
      <span>? <b>${pdfFolderFileNames.length} PDF${pdfFolderFileNames.length === 1 ? "" : "s"}</b> uploaded � will be matched by recipient name.</span>
      <button class="btn outline small-btn danger" onclick="clearPdfFolder()">? Clear</button>
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
    setActionMessage(`? ${d.fileCount} personal PDFs ready. Contacts without a matching PDF will be skipped.`, "success");
    toast(`${d.fileCount} PDFs uploaded successfully.`, "success");
  } catch (e) {
    setActionMessage("? " + e.message, "error");
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

// -- Row Range Preview ---------------------------------------------------------

function updateRowRangePreview() {
  const start = parseInt($('rowStart')?.value) || 1;
  const end = parseInt($('rowEnd')?.value);
  const total = contacts.length;
  const preview = $('rowRangePreview');
  if (!preview) return;
  if (!$('rowStart')?.value && !$('rowEnd')?.value) {
    preview.textContent = total > 0 ? `All ${total} rows` : '';
    return;
  }
  const effectiveEnd = end ? Math.min(end, total || end) : (total || '?');
  const count = typeof effectiveEnd === 'number' && total > 0
    ? Math.max(0, effectiveEnd - start + 1)
    : '?';
  preview.textContent = `Rows ${start}–${effectiveEnd} = ${count} email${count !== 1 ? 's' : ''}`;
}
window.updateRowRangePreview = updateRowRangePreview;

const rowStartEl = $('rowStart');
const rowEndEl = $('rowEnd');
if (rowStartEl) rowStartEl.addEventListener('input', updateRowRangePreview);
if (rowEndEl) rowEndEl.addEventListener('input', updateRowRangePreview);

// -- Send Timer ----------------------------------------------------------------
let sendMode = "delay"; // "immediate" | "delay"

function setSendMode(mode) {
  sendMode = mode;
  const modeImmediate = $("modeImmediate");
  const modeDelay     = $("modeDelay");
  const delayControls = $("delayControls");
  const delayPreview  = $("delayPreview");

  if (modeImmediate) modeImmediate.classList.toggle("active", mode === "immediate");
  if (modeDelay)     modeDelay.classList.toggle("active",     mode === "delay");
  if (delayControls) delayControls.style.display = mode === "delay" ? "" : "none";
  if (delayPreview)  delayPreview.style.display  = mode === "delay" ? "" : "none";
}
window.setSendMode = setSendMode;

function getDelaySeconds() {
  if (sendMode === "immediate") return 0;
  const mins = parseInt($('delayMinutes')?.value || 0) || 0;
  const secs = parseInt($('delaySeconds')?.value || 0) || 0;
  return Math.max(0, mins * 60 + secs);
}

function updateDelayPreview() {
  if (sendMode === 'immediate') {
    const el = $('delayPreview');
    if (el) el.textContent = 'Emails will be sent immediately (no delay)';
    return;
  }
  const mins = parseInt($('delayMinutes')?.value || 0) || 0;
  const secs = parseInt($('delaySeconds')?.value || 0) || 0;
  const el = $('delayPreview');
  if (!el) return;
  if (mins === 0 && secs === 0) {
    el.textContent = 'Emails will be sent immediately (no delay)';
  } else if (mins === 0) {
    el.textContent = `Each email will be sent ${secs} second${secs !== 1 ? 's' : ''} apart`;
  } else if (secs === 0) {
    el.textContent = `Each email will be sent ${mins} minute${mins !== 1 ? 's' : ''} apart`;
  } else {
    el.textContent = `Each email will be sent ${mins} min ${secs} sec apart`;
  }
}

function setDelayPreset(mins, secs) {
  const mEl = $('delayMinutes');
  const sEl = $('delaySeconds');
  if (mEl) mEl.value = mins;
  if (sEl) sEl.value = secs;
  document.querySelectorAll('.delay-preset-btn').forEach(b => b.classList.remove('active'));
  updateDelayPreview();
  if (mins === 0 && secs === 0) {
    setSendMode('immediate');
  } else {
    setSendMode('delay');
  }
}
window.setDelayPreset = setDelayPreset;

// New delay: minutes + seconds inputs
const delayMinsEl = $('delayMinutes');
const delaySecsEl = $('delaySeconds');
if (delayMinsEl) delayMinsEl.addEventListener('input', updateDelayPreview);
if (delaySecsEl) delaySecsEl.addEventListener('input', updateDelayPreview);
updateDelayPreview();

// -- Live Preview --------------------------------------------------------------

function updatePreview() {
  const subject = $("subject").value || "Your subject";
  const body = getEmailBody() || "Your personalized message will appear here.";

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
  $("previewBody").innerHTML = renderedBody;

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
// body is now Quill - Quill text-change event triggers updatePreview

function setActionMessage(message, type = "info") {
  $("actionMessage").className = `action-message ${type}`;
  $("actionMessage").textContent = message;
}

// -- Microsoft Status & Configuration ------------------------------------------

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
    if ($("msSubmitBtn")) $("msSubmitBtn").textContent = "Switch Microsoft Account ?";
  } else {
    pill?.classList.remove("connected");
    dot?.classList.remove("connected");
    if (text) text.textContent = msStatus.clientIdConfigured ? "Connect Outlook" : "Setup Outlook";
    if (connectedBox) connectedBox.classList.add("hidden");
    if ($("msSubmitBtn")) $("msSubmitBtn").textContent = "Sign in with Microsoft ?";
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
        ${esc(s.label)} ${s.recommended ? "? Recommended" : ""}
      </option>
    `).join("");

    updateSendButtonLabel();
  } catch {}
}

// -- Email Preview Modal --------------------------------------------------------

function showEmailPreview() {
  const subject = $("subject").value.trim();
  const body = getEmailBody().trim();

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
  bodyEl.innerHTML = renderedBody;

  const warningsEl = $("previewModalWarnings");
  if (allUnresolved.length) {
    warningsEl.classList.remove("hidden");
    warningsEl.innerHTML = `?? <b>Unknown variables detected:</b> ${allUnresolved.map(v => `<code>${esc(v)}</code>`).join(", ")} � these will be sent as-is.`;
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

// -- Validation Modal ----------------------------------------------------------

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
  icon.textContent = hasIssues ? "?" : "?";

  title.textContent = hasIssues ? "Validation: Issues Found" : "Validation: All Clear";
  subtitle.textContent = hasIssues
    ? `${result.invalid} row(s) will be skipped`
    : "All selected rows are ready to send";

  banner.innerHTML = `
    <div class="validation-stat"><span>${result.total}</span><small>Total rows</small></div>
    <div class="validation-stat ok"><span>${result.valid}</span><small>? Will send</small></div>
    <div class="validation-stat ${result.invalid > 0 ? "warn" : ""}"><span>${result.invalid}</span><small>? Will skip</small></div>
    <div class="validation-stat ${result.missingEmail > 0 ? "err" : ""}"><span>${result.missingEmail}</span><small>? No email</small></div>
  `;

  if (hasIssues) {
    issuesList.classList.remove("hidden");
    const badRows = result.rows.filter(r => !r.valid);
    issuesRows.innerHTML = badRows.slice(0, 20).map(r => `
      <div class="validation-issue-row">
        <span class="validation-issue-name">${esc(r.name)}</span>
        <span class="validation-issue-reason">${r.issues.map(esc).join(" � ")}</span>
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
    proceedBtn.textContent = `Send ${result.valid} email${result.valid === 1 ? "" : "s"} ?`;
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

// -- Confirm Modal -------------------------------------------------------------

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
    warning.textContent = `? ${validResult.invalid} row(s) with no valid email will be skipped.`;
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

// -- Send Campaign Flow --------------------------------------------------------

function openMsModal() {
  // If already configured via env vars and not yet connected, go straight to Microsoft login
  if (msStatus.clientIdConfigured && !msStatus.connected) {
    toast("Redirecting to Microsoft login�", "success");
    setTimeout(() => { window.location.href = "/auth/microsoft/login"; }, 400);
    return;
  }

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
  if (!getEmailBody().trim() || getEmailBody().trim() === '<p><br></p>') return toast("Enter a message body.", "error");

  const subject = $("subject").value;
  const body = getEmailBody();
  const select = $("senderSelect");
  const senderType = select?.value || "microsoft";
  const senderLabel = senderType === "microsoft" ? "Microsoft Outlook" : "Custom SMTP";

  setActionMessage("Validating recipients...", "info");

  try {
    const validResult = await runValidation(ids, subject, body);

    if (validResult.valid === 0) {
      setActionMessage("? No valid rows to send. Check email addresses.", "error");
      return toast("No valid rows found. All selected rows have invalid/missing emails.", "error");
    }

    // Show validation modal � proceed ? confirm modal ? actual send
    showValidationModal(validResult, () => {
      showConfirmModal(validResult, () => doSendCampaign(ids, senderType, senderLabel, validResult));
    });

    setActionMessage("", "info");
  } catch (e) {
    setActionMessage("? " + e.message, "error");
    toast(e.message, "error");
  }
}
window.sendCampaignDirect = sendCampaignDirect;

async function doSendCampaign(ids, senderType, senderLabel, validResult) {
  const subject = $("subject").value;
  const body = getEmailBody();

  // Warn if PDF folder active
  if (currentPdfFolderId) {
    const ok = await showConfirm(
      `?? Personal PDFs are active (${pdfFolderFileNames.length} PDFs).\n\nRows without a matching PDF will be SKIPPED.\n\nSend campaign to ${validResult.valid} rows via ${senderLabel}?`,
      { title: "Send with Personal PDFs", icon: "??", okText: "Send", cancelText: "Cancel" }
    );
    if (!ok) return;
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
        scheduledAt: getScheduledAt(),
        rowStart: parseInt($('rowStart')?.value) || null,
        rowEnd: parseInt($('rowEnd')?.value) || null,
        requestDeliveryReceipt: !!$('requestDeliveryReceipt')?.checked,
        requestReadReceipt: !!$('requestReadReceipt')?.checked
      })
    });

    if (d.usedPdfFolder) {
      if (d.skippedCount > 0) {
        toast(`?? ${d.matchedCount} will be sent � ${d.skippedCount} skipped (no matching PDF)`, "info");
      } else {
        toast(`?? All ${d.matchedCount} contacts matched!`, "success");
      }
    }

    currentCampaignId = d.campaignId;

    if (d.scheduled) {
      const when = new Date(d.scheduledAt).toLocaleString();
      toast(`?? Campaign scheduled! Will send automatically at ${when}`, "success");
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
    setActionMessage("? " + e.message, "error");
    if (senderType === "microsoft" && e.message.includes("not connected")) {
      setTimeout(() => openMsModal(), 700);
    }
  }
}

// -- Duplicate Loan Number Modal -----------------------------------------------
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
      <span class="duplicate-loan">?? Loan Number: <b>${esc(d.loanNumber)}</b></span>
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

// -- Schedule Helpers ----------------------------------------------------------

function toggleSchedule() {
  const on  = $("scheduleToggle") && $("scheduleToggle").checked;
  const box = $("scheduleBox");
  if (box) box.classList.toggle("hidden", !on);

  if (on) {
    // Pre-fill with 1 hour from now
    const d   = new Date(Date.now() + 60 * 60 * 1000);
    const pad = n => String(n).padStart(2, "0");
    if ($("scheduledDate")) $("scheduledDate").value = `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
    if ($("scheduledTime")) $("scheduledTime").value = `${pad(d.getHours())}:${pad(d.getMinutes())}`;
    updateSchedulePreview();

    // Live update preview when user changes date/time
    $("scheduledDate")?.addEventListener("change", updateSchedulePreview);
    $("scheduledTime")?.addEventListener("change", updateSchedulePreview);
  }
}
window.toggleSchedule = toggleSchedule;

function updateSchedulePreview() {
  const dateVal = $("scheduledDate")?.value;
  const timeVal = $("scheduledTime")?.value;
  const el      = $("schedulePreviewText");
  if (!el) return;
  if (dateVal && timeVal) {
    const dt = new Date(`${dateVal}T${timeVal}`);
    el.textContent = `?? Will send on ${dt.toLocaleDateString("en-IN", { weekday:"long", year:"numeric", month:"long", day:"numeric" })} at ${dt.toLocaleTimeString("en-IN", { hour:"2-digit", minute:"2-digit" })}`;
  } else {
    el.textContent = "Select a date and time above";
  }
}

function getScheduledAt() {
  if (!$("scheduleToggle") || !$("scheduleToggle").checked) return null;
  const dateVal = $("scheduledDate")?.value;
  const timeVal = $("scheduledTime")?.value;
  if (!dateVal || !timeVal) return null;
  return new Date(`${dateVal}T${timeVal}`).toISOString();
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

// -- Campaign History ----------------------------------------------------------

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
        <td>${Number(c.skipped || 0) > 0 ? `<span class="skipped-pill">${c.skipped}</span>` : "�"}</td>
        <td>${new Date(c.created_at + "Z").toLocaleDateString()}</td>
        <td style="display:flex;gap:5px;flex-wrap:wrap">
          <button class="btn small outline" onclick="openCampaign(${c.id})">View</button>
          <button class="btn small outline" onclick="downloadReport(${c.id})" title="Download full report">? Report</button>
          ${c.failed > 0 ? `<button class="btn small outline failed-dl-btn" onclick="downloadFailed(${c.id})" title="Download failed records">? Failed</button>` : ""}
          ${Number(c.skipped || 0) > 0 ? `<button class="btn small outline skipped-dl-btn" onclick="downloadSkipped(${c.id})" title="Download not-sent contacts">? Not Sent</button>` : ""}
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
      `${d.campaign.sent} sent � ${d.campaign.failed} failed � ${d.campaign.total} total`;

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
        skippedStat.innerHTML = `?? <b>${skippedCount} row${skippedCount === 1 ? "" : "s"} skipped</b> � no matching PDF or invalid email. Their emails were <b>not sent</b>.`;
        if (dlBtn) dlBtn.classList.remove("hidden");
      } else {
        skippedStat.classList.add("hidden");
        if (dlBtn) dlBtn.classList.add("hidden");
      }
    }

    if (scroll) $("progressPanel").scrollIntoView({ behavior: "smooth", block: "start" });

    // Show/hide campaign control buttons based on status
    const status = d.campaign.status;
    const pauseBtn = $('pauseCampaignBtn');
    const resumeBtn = $('resumeCampaignBtn');
    const stopBtn = $('stopCampaignBtn');
    const cancelBtn = $('cancelCampaignBtn');
    if (pauseBtn) pauseBtn.classList.toggle('hidden', !['sending','resumed'].includes(status));
    if (resumeBtn) resumeBtn.classList.toggle('hidden', status !== 'paused');
    if (stopBtn) stopBtn.classList.toggle('hidden', !['sending','resumed','paused'].includes(status));
    if (cancelBtn) cancelBtn.classList.toggle('hidden', !['paused','scheduled'].includes(status));

    // Show/hide campaign control buttons based on status
    const _cStatus = d.campaign.status;
    const _pauseBtn = $('pauseCampaignBtn');
    const _resumeBtn = $('resumeCampaignBtn');
    const _stopBtn = $('stopCampaignBtn');
    const _cancelBtn = $('cancelCampaignBtn');
    if (_pauseBtn) _pauseBtn.classList.toggle('hidden', !['sending','resumed'].includes(_cStatus));
    if (_resumeBtn) _resumeBtn.classList.toggle('hidden', _cStatus !== 'paused');
    if (_stopBtn) _stopBtn.classList.toggle('hidden', !['sending','resumed','paused'].includes(_cStatus));
    if (_cancelBtn) _cancelBtn.classList.toggle('hidden', !['paused','scheduled'].includes(_cStatus));

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
  box.innerHTML = `<div class="attachment-history"><b>?? Global Attachments:</b> ${items.map(a => `<span>${esc(a.original_name)} <small>${formatBytes(a.size)}</small></span>`).join("")}</div>`;
}

function renderRecipients(recipients) {
  $("recipientsTable").innerHTML = recipients.map((r, i) => {
    const sent = r.status === "sent";
    const notSent = r.status === "not_sent";
    const statusCls = sent ? "completed" : r.status === "failed" ? "failed" : notSent ? "skipped-badge" : "sending";
    const statusLabel = notSent ? "NOT SENT" : r.status.toUpperCase();
    const pdfNote = r.pdf_match_status === "skipped" ? '<span class="pdf-skip-note" title="No matching PDF / invalid email">???</span>' :
                    r.pdf_match_status === "matched" ? '<span class="pdf-skip-note" title="Personal PDF attached">???</span>' : "";

    // Parse To emails for display
    const toStr = r.to_emails || r.email || "";
    const toDisplay = toStr.split(",").map(s => s.trim()).filter(Boolean).join(", ") || r.email || "�";
    const ccStr = r.cc ? r.cc.split(";").filter(Boolean).join(", ") : "�";

    return `
      <tr class="${notSent ? "row-not-sent" : ""}">
        <td>${i + 1}</td>
        <td><b>${esc(r.name)}</b>${pdfNote}</td>
        <td class="to-emails-cell">${esc(toDisplay)}</td>
        <td class="muted">${esc(ccStr)}</td>
        <td><span class="status ${statusCls}">${esc(statusLabel)}</span></td>
        <td class="${r.status === "failed" ? "error-cell" : "muted"}">${notSent ? "Not sent � no matching PDF or invalid email" : esc(r.error || "�")}</td>
        <td>${r.sent_at ? new Date(r.sent_at + "Z").toLocaleString() : "�"}</td>
      </tr>
    `;
  }).join("") || `<tr><td colspan="7" class="empty">No recipient details available.</td></tr>`;
}

function updateProgress(p) {
  const done = (p.sent || 0) + (p.failed || 0);
  const skipped = Number(p.skipped || 0);
  const pct = p.total ? Math.round(done / p.total * 100) : 0;

  $("progressTitle").textContent = p.name;
  const skippedNote = skipped > 0 ? ` � ${skipped} skipped` : "";
  $("progressText").textContent = `${done} / ${p.total} processed � ${p.sent} sent � ${p.failed} failed${skippedNote}`;
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

// ============================================================
// Failure-Pause Modal — auto-pause on recipient failure
// ============================================================
let _fpCampaignId = null;
let _fpBatchMode = false;
let _fpBatches = [];
let _fpCurrentTab = 'continue';

function showFailurePauseModal(campaignId, info) {
  _fpCampaignId = campaignId;
  _fpBatchMode = Boolean(info.batch_mode);
  _fpBatches = info.batches || [];
  _fpCurrentTab = 'continue';

  const rec = info.fail_pause_recipient || 'Unknown';
  const err = info.fail_pause_error || 'Unknown error';
  const recEl = document.getElementById('fpErrorRecipient');
  const errEl = document.getElementById('fpErrorMsg');
  if (recEl) recEl.textContent = '\u{1F464} Recipient: ' + rec;
  if (errEl) errEl.textContent = '\u274C Error: ' + err;

  const totalSecs = Number(info.delay_seconds || 5);
  const dMin = document.getElementById('fpDelayMins');
  const dSec = document.getElementById('fpDelaySecs');
  if (dMin) dMin.value = Math.floor(totalSecs / 60);
  if (dSec) dSec.value = totalSecs % 60 || 5;

  fpSetResumeNow();
  fpSelectTab('continue');

  const batchSection = document.getElementById('fpBatchSection');
  if (batchSection) {
    if (_fpBatchMode && _fpBatches.length > 1) {
      const bc = document.getElementById('fpBatchCount');
      if (bc) bc.textContent = _fpBatches.length;
      renderFpBatchTimings();
      batchSection.classList.remove('hidden');
    } else {
      batchSection.classList.add('hidden');
    }
  }

  const modal = document.getElementById('failurePauseModal');
  if (modal) modal.classList.remove('hidden');
}
window.showFailurePauseModal = showFailurePauseModal;

function closeFpModal() {
  const modal = document.getElementById('failurePauseModal');
  if (modal) modal.classList.add('hidden');
  if (_fpCampaignId) openCampaign(_fpCampaignId, false);
}
window.closeFpModal = closeFpModal;

function fpSelectTab(tab) {
  _fpCurrentTab = tab;
  ['continue','delay','stop'].forEach(function(t) {
    const id0 = t.charAt(0).toUpperCase() + t.slice(1);
    const btn = document.getElementById('fpTab' + id0);
    const panel = document.getElementById('fpPanel' + id0);
    if (btn) btn.classList.toggle('active', t === tab);
    if (panel) panel.classList.toggle('hidden', t !== tab);
  });
  const confirmBtn = document.getElementById('fpConfirmBtn');
  if (!confirmBtn) return;
  if (tab === 'stop') {
    confirmBtn.textContent = '\u25A0 Stop Campaign';
    confirmBtn.style.background = '#ef4444';
  } else if (tab === 'delay') {
    confirmBtn.textContent = '\u23F0 Schedule Resume';
    confirmBtn.style.background = '';
  } else {
    confirmBtn.textContent = '\u25B6 Continue Campaign';
    confirmBtn.style.background = '';
  }
}
window.fpSelectTab = fpSelectTab;

function fpSetResumeNow() {
  const now = new Date();
  now.setSeconds(0, 0);
  const local = new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
  const el = document.getElementById('fpResumeAt');
  if (el) { el.value = local; fpUpdateBatchPreview(); }
}
window.fpSetResumeNow = fpSetResumeNow;

function fpSetResumeIn(minutes) {
  const t = new Date(Date.now() + minutes * 60000);
  t.setSeconds(0, 0);
  const local = new Date(t.getTime() - t.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
  const el = document.getElementById('fpResumeAt');
  if (el) { el.value = local; fpUpdateBatchPreview(); }
}
window.fpSetResumeIn = fpSetResumeIn;

function fpSetResumeTomorrow() {
  const t = new Date();
  t.setDate(t.getDate() + 1);
  t.setHours(9, 0, 0, 0);
  const local = new Date(t.getTime() - t.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
  const el = document.getElementById('fpResumeAt');
  if (el) { el.value = local; fpUpdateBatchPreview(); }
}
window.fpSetResumeTomorrow = fpSetResumeTomorrow;

function fpBatchAdjustChanged(val) { fpUpdateBatchPreview(); }
window.fpBatchAdjustChanged = fpBatchAdjustChanged;

function renderFpBatchTimings() {
  const container = document.getElementById('fpBatchTimings');
  if (!container) return;
  container.innerHTML = _fpBatches.map(function(b, i) {
    return '<div class="fp-batch-row">' +
      '<span class="fp-batch-num">Batch ' + (i+1) + '</span>' +
      '<span class="fp-batch-time">' + new Date(b.scheduled_at).toLocaleString() + '</span>' +
      '<span class="fp-batch-count">' + (b.row_count || b.recipient_count || '?') + ' emails</span>' +
    '</div>';
  }).join('');
}
window.renderFpBatchTimings = renderFpBatchTimings;

function fpUpdateBatchPreview() {
  const previewEl = document.getElementById('fpBatchAutoPreview');
  if (!previewEl || !_fpBatches.length) return;
  const adjustInput = document.querySelector('input[name="fpBatchAdjust"]:checked');
  if (!adjustInput || adjustInput.value !== 'auto') { previewEl.innerHTML = ''; return; }
  const resumeVal = document.getElementById('fpResumeAt') ? document.getElementById('fpResumeAt').value : '';
  if (!resumeVal || !_fpBatches[0]) { previewEl.innerHTML = ''; return; }

  const newResume = new Date(resumeVal).getTime();
  const origFirst = new Date(_fpBatches[0].scheduled_at).getTime();
  const offsetMs = newResume - origFirst;

  if (Math.abs(offsetMs) < 60000) {
    previewEl.innerHTML = '<span style="color:#22c55e;font-size:12px">\u2713 No change needed</span>';
    return;
  }

  const sign = offsetMs > 0 ? '+' : '-';
  const absMins = Math.round(Math.abs(offsetMs) / 60000);
  const absHrs = Math.floor(absMins / 60);
  const remMins = absMins % 60;
  const delta = absHrs > 0 ? absHrs + 'h ' + remMins + 'm' : absMins + 'm';

  previewEl.innerHTML = '<div class="fp-preview-header">\uD83D\uDCC5 New batch schedule (all shifted ' + sign + delta + '):</div>' +
    _fpBatches.map(function(b, i) {
      const newT = new Date(new Date(b.scheduled_at).getTime() + offsetMs);
      return '<div class="fp-preview-row"><b>Batch ' + (i+1) + ':</b> ' + newT.toLocaleString() + '</div>';
    }).join('');
}
window.fpUpdateBatchPreview = fpUpdateBatchPreview;

(function() {
  const el = document.getElementById('fpResumeAt');
  if (el) el.addEventListener('change', fpUpdateBatchPreview);
})();

async function fpConfirm() {
  if (!_fpCampaignId) return;
  const tab = _fpCurrentTab;
  const btn = document.getElementById('fpConfirmBtn');
  if (btn) { btn.disabled = true; btn.textContent = '\u23F3 Processing...'; }

  try {
    var body = { action: tab };

    if (tab === 'stop') {
      body.action = 'stop';
    } else if (tab === 'delay') {
      const resumeEl = document.getElementById('fpResumeAt');
      const resumeVal = resumeEl ? resumeEl.value : '';
      if (!resumeVal) { toast('Please select a resume date/time', 'error'); if (btn) btn.disabled = false; fpSelectTab(tab); return; }
      const mins = parseInt(document.getElementById('fpDelayMins') ? document.getElementById('fpDelayMins').value : '0') || 0;
      const secs = parseInt(document.getElementById('fpDelaySecs') ? document.getElementById('fpDelaySecs').value : '5') || 5;
      const totalSecs = mins * 60 + secs;
      const adjustInput = document.querySelector('input[name="fpBatchAdjust"]:checked');
      body = {
        action: 'delay',
        resume_at: new Date(resumeVal).toISOString(),
        delay_seconds: totalSecs,
        postpone_batches: adjustInput && adjustInput.value === 'auto'
      };
    } else {
      const mins = parseInt(document.getElementById('fpDelayMins') ? document.getElementById('fpDelayMins').value : '0') || 0;
      const secs = parseInt(document.getElementById('fpDelaySecs') ? document.getElementById('fpDelaySecs').value : '5') || 5;
      body.delay_seconds = mins * 60 + secs;
      body.action = 'continue';
    }

    await api('/api/campaigns/' + _fpCampaignId + '/resume-after-failure', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });

    const modal = document.getElementById('failurePauseModal');
    if (modal) modal.classList.add('hidden');

    if (tab === 'stop') {
      toast('Campaign stopped.', 'info');
      await loadCampaigns();
    } else if (tab === 'delay') {
      const resumeEl = document.getElementById('fpResumeAt');
      const resumeVal = resumeEl ? resumeEl.value : '';
      toast('Campaign scheduled to resume at ' + (resumeVal ? new Date(resumeVal).toLocaleString() : 'scheduled time'), 'success');
      await openCampaign(_fpCampaignId, false);
    } else {
      toast('Campaign continuing...', 'success');
      await openCampaign(_fpCampaignId, false);
      pollProgress(_fpCampaignId);
    }
  } catch(e) {
    toast('Error: ' + (e.message || e), 'error');
  } finally {
    if (btn) { btn.disabled = false; fpSelectTab(_fpCurrentTab); }
  }
}
window.fpConfirm = fpConfirm;



async function pollProgress(id) {
  const tick = async () => {
    try {
      const p = await api(`/api/campaigns/${id}/progress`);
      updateProgress(p);

      const d = await api(`/api/campaigns/${id}`);
      renderRecipients(d.recipients);

      if (p.status === "sending") {
        setTimeout(tick, 900);
      } else if (p.status === "paused" || p.status === "scheduled") {
        try {
          const fpInfo = await api(`/api/campaigns/${id}/fail-pause-info`);
          if (fpInfo.fail_pause) { showFailurePauseModal(id, fpInfo); }
        } catch {}
        await loadCampaigns();
      } else {
        $("progressMeta").textContent = `${p.sent} sent � ${p.failed} failed � ${p.total} total`;
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

// -- Templates -----------------------------------------------------------------

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
  sel.innerHTML = `<option value="">� Load a saved template �</option>` +
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
        <button class="btn outline small-btn danger" onclick="deleteTemplate(${t.id})">?</button>
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
  const bodyEditor = document.getElementById('bodyContentEditor');
  if (bodyEditor) bodyEditor.innerHTML = t.body || '';
  document.getElementById('body').value = t.body || '';
  updateBodyPreview();
  updatePreview();
  toast(`Template "${t.name}" loaded.`, "success");
}
window.applyTemplate = applyTemplate;

async function saveTemplate() {
  const subject = $("subject").value.trim();
  const body = getEmailBody().trim();
  if (!subject || !body) return toast("Write a subject and message first.", "error");

  const name = prompt("Enter a name for this template:");
  if (!name || !name.trim()) return;

  try {
    await api("/api/templates", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: name.trim(), subject, body: getEmailBody() })
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
  const ok = await showConfirm(`Delete template "${t.name}"? This cannot be undone.`, { title: "Delete Template", icon: "???", okText: "Delete", danger: true });
  if (!ok) return;
  try {
    await api(`/api/templates/${id}`, { method: "DELETE" });
    await loadTemplates();
    toast("Template deleted.", "success");
  } catch (e) {
    toast(e.message, "error");
  }
}
window.deleteTemplate = deleteTemplate;

// -- Sender Check --------------------------------------------------------------

async function checkCurrentSender() {
  const select = $("senderSelect");
  const senderType = select?.value || "microsoft";
  setActionMessage("Checking sender connection...", "info");

  try {
    if (senderType === "microsoft") {
      const d = await api("/api/microsoft/check", { method: "POST" });
      setActionMessage(`? Microsoft Outlook connected as ${d.email} (${d.name || "Verified"})`, "success");
      toast(`Outlook verified: ${d.email}`, "success");
    } else {
      const d = await api("/api/test-smtp", { method: "POST" });
      setActionMessage("? " + d.message, "success");
      toast(d.message, "success");
    }
  } catch (e) {
    setActionMessage("? " + e.message, "error");
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
        body: getEmailBody(),
        sampleName: contacts[0]?.name || "Recipient",
        attachmentIds: selectedAttachments.map(a => a.id),
        senderType
      })
    });
    setActionMessage("? " + d.message, "success");
    toast(d.message, "success");
  } catch (e) {
    setActionMessage("? " + e.message, "error");
    toast(e.message, "error");
    if (senderType === "microsoft" && e.message.includes("not connected")) {
      setTimeout(() => openMsModal(), 700);
    }
  }
}
window.sendTest = sendTest;

// -- Campaign Pause / Resume / Stop / Cancel -----------------------------------

async function pauseCurrentCampaign() {
  if (!currentCampaignId) return;
  try {
    await api(`/api/campaigns/${currentCampaignId}/pause`, { method: 'POST' });
    toast('Campaign paused', 'success');
    setTimeout(() => openCampaign(currentCampaignId, false), 1000);
  } catch(e) { toast(e.message, 'error'); }
}
async function resumeCurrentCampaign() {
  if (!currentCampaignId) return;
  try {
    await api(`/api/campaigns/${currentCampaignId}/resume`, { method: 'POST' });
    toast('Campaign resumed', 'success');
    setTimeout(() => openCampaign(currentCampaignId, false), 1000);
  } catch(e) { toast(e.message, 'error'); }
}
async function stopCurrentCampaign() {
  if (!currentCampaignId) return;
  const ok = await showConfirm('Stop this campaign? It cannot be restarted.', { title: 'Stop Campaign', icon: '⏹', okText: 'Stop', cancelText: 'Cancel', danger: true });
  if (!ok) return;
  try {
    await api(`/api/campaigns/${currentCampaignId}/stop`, { method: 'POST' });
    toast('Campaign stopped', 'success');
    setTimeout(() => openCampaign(currentCampaignId, false), 1000);
  } catch(e) { toast(e.message, 'error'); }
}
async function cancelCurrentCampaign() {
  if (!currentCampaignId) return;
  const ok = await showConfirm('Cancel this campaign permanently?', { title: 'Cancel Campaign', icon: '✕', okText: 'Cancel Campaign', cancelText: 'Keep', danger: true });
  if (!ok) return;
  try {
    await api(`/api/campaigns/${currentCampaignId}/cancel`, { method: 'POST' });
    toast('Campaign cancelled', 'success');
    setTimeout(() => { openCampaign(currentCampaignId, false); loadCampaigns(); }, 1000);
  } catch(e) { toast(e.message, 'error'); }
}
window.pauseCurrentCampaign = pauseCurrentCampaign;
window.resumeCurrentCampaign = resumeCurrentCampaign;
window.stopCurrentCampaign = stopCurrentCampaign;
window.cancelCurrentCampaign = cancelCurrentCampaign;

// -- SMTP Provider / OTP setup -------------------------------------------------

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
              "Go to <b>Advanced security options</b> ? <b>App passwords</b>",
              "Click <b>Create a new app password</b> ? copy it",
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
    other: { title: "Custom SMTP � enter host above and use your email password.", steps: [] }
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

  msg.textContent = "Connecting and sending code�"; msg.style.color = "#6b7280";

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
    msg.textContent = "? " + e.message;
    msg.style.color = "#ef4444";
  }
}
window.quickSmtpSendOtp = quickSmtpSendOtp;

async function saveMsConfig(e) {
  if (e) e.preventDefault();

  let clientId     = ($("msClientId")?.value     || "").trim();
  const clientSecret = ($("msClientSecret")?.value  || "").trim();
  const tenantId     = ($("msTenantId")?.value      || "common").trim();

  // If input is empty but server already has client ID configured (via env vars), go straight to login
  if (!clientId && msStatus.clientIdConfigured) {
    toast("Redirecting to Microsoft login�", "success");
    setTimeout(() => { window.location.href = "/auth/microsoft/login"; }, 400);
    return;
  }

  if (!clientId) {
    return toast("Please enter your Application (Client) ID.", "error");
  }

  // Warn if user pasted api:// prefix
  if (clientId.startsWith("api://") || clientId.startsWith("https://")) {
    return toast("Remove api:// from the Client ID � paste only the UUID part.", "error");
  }

  const btn = $("msSubmitBtn");
  if (btn) { btn.disabled = true; btn.textContent = "Saving & redirecting�"; }

  try {
    // Save config to DB (only if user actually typed something)
    await api("/api/microsoft/config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ clientId, clientSecret, tenantId })
    });

    toast("Config saved � redirecting to Microsoft login�", "success");
    setTimeout(() => { window.location.href = "/auth/microsoft/login"; }, 500);
  } catch (err) {
    toast("Error saving config: " + err.message, "error");
    if (btn) { btn.disabled = false; btn.textContent = "Sign in with Microsoft ?"; }
  }
}

async function disconnectMicrosoft() {
  const ok = await showConfirm("Disconnect the Microsoft Outlook account?", { title: "Disconnect Outlook", icon: "??", okText: "Disconnect", danger: true });
  if (!ok) return;
  api("/api/microsoft/disconnect", { method: "POST" })
    .then(() => { toast("Microsoft account disconnected.", "success"); closeMsModal(); loadMicrosoftStatus(); })
    .catch(e => toast(e.message, "error"));
}

function copyRedirectUri() {
  const txt = $("redirectUriDisplay") ? $("redirectUriDisplay").textContent : "";
  if (!txt || txt === "loading...") return;
  navigator.clipboard.writeText(txt).then(() => toast("? Redirect URI copied!", "success")).catch(() => {});
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

// -- Switch Mail Modal ---------------------------------------------------------
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

  msg1.textContent = "Connecting to Outlook and sending code�"; msg1.style.color = "#6b7280";

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
    msg1.textContent = "? " + e.message;
    msg1.style.color = "#ef4444";
  }
}
window.smtpSendOtp = smtpSendOtp;

async function smtpConfirmOtp() {
  const otp  = $("smtpOtpCode").value.trim();
  const msg2 = $("smtpOtpMsg2");

  if (!otp || otp.length < 4) { msg2.textContent = "Please enter the 6-digit code from your inbox."; msg2.style.color = "#ef4444"; return; }

  msg2.textContent = "Verifying�"; msg2.style.color = "#6b7280";

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
    msg2.textContent = "? " + e.message;
    msg2.style.color = "#ef4444";
  }
}
window.smtpConfirmOtp = smtpConfirmOtp;

// ── Batch Scheduler ───────────────────────────────────────────────────────────

function setScheduleMode(mode) {
  scheduleMode = mode;
  const modeAllBtn = $('modeAllRows');
  const modeBatchBtn = $('modeBatch');
  const rowRangePanel = $('rowRangePanel');
  const batchPanel = $('batchSchedulerPanel');
  if (modeAllBtn) modeAllBtn.classList.toggle('active', mode === 'all');
  if (modeBatchBtn) modeBatchBtn.classList.toggle('active', mode === 'batch');
  if (rowRangePanel) rowRangePanel.classList.toggle('hidden', mode === 'batch');
  if (batchPanel) batchPanel.classList.toggle('hidden', mode === 'all');
  if (mode === 'batch') renderBatchList();
}
window.setScheduleMode = setScheduleMode;

function addBatch() {
  batches.push({ rowStart: '', rowEnd: '', scheduledDate: '', scheduledTime: '', delayMinutes: 1, delaySeconds: 0 });
  renderBatchList();
}
window.addBatch = addBatch;

function removeBatch(i) {
  batches.splice(i, 1);
  renderBatchList();
}
window.removeBatch = removeBatch;

function updateBatch(i, key, val) {
  batches[i][key] = val;
  updateBatchSummary();
}
window.updateBatch = updateBatch;

function getBatchRowPreview(b) {
  if (!b.rowStart && !b.rowEnd) return '';
  const s = parseInt(b.rowStart) || 1;
  const e = parseInt(b.rowEnd);
  const count = e ? Math.max(0, e - s + 1) : '?';
  const dt = b.scheduledDate && b.scheduledTime ? ` — ${b.scheduledDate} ${b.scheduledTime}` : '';
  return `Rows ${s}${e ? '–' + e : '+'} (${count} emails)${dt}`;
}

function detectBatchOverlaps() {
  const overlaps = [];
  for (let i = 0; i < batches.length; i++) {
    for (let j = i + 1; j < batches.length; j++) {
      const a = batches[i], b = batches[j];
      const aS = parseInt(a.rowStart) || 1, aE = parseInt(a.rowEnd) || 999999;
      const bS = parseInt(b.rowStart) || 1, bE = parseInt(b.rowEnd) || 999999;
      if (aS <= bE && bS <= aE) overlaps.push([i, j]);
    }
  }
  return overlaps;
}

function renderBatchList() {
  const list = $('batchList');
  if (!list) return;
  list.innerHTML = batches.map((b, i) => `
    <div class="batch-item">
      <div class="batch-item-head">
        <span class="batch-item-label">Batch ${i + 1}</span>
        <button class="batch-remove-btn" onclick="removeBatch(${i})">✕</button>
      </div>
      <div class="batch-fields">
        <div class="batch-field"><label>Start Row</label><input type="number" min="1" value="${b.rowStart || ''}" placeholder="1" onchange="updateBatch(${i},'rowStart',this.value)"></div>
        <div class="batch-field"><label>End Row</label><input type="number" min="1" value="${b.rowEnd || ''}" placeholder="last" onchange="updateBatch(${i},'rowEnd',this.value)"></div>
        <div class="batch-field"><label>Date</label><input type="date" value="${b.scheduledDate || ''}" onchange="updateBatch(${i},'scheduledDate',this.value)"></div>
        <div class="batch-field"><label>Time</label><input type="time" value="${b.scheduledTime || ''}" onchange="updateBatch(${i},'scheduledTime',this.value)"></div>
        <div class="batch-field"><label>Delay Min</label><input type="number" min="0" max="99" value="${b.delayMinutes || 0}" onchange="updateBatch(${i},'delayMinutes',+this.value)"></div>
        <div class="batch-field"><label>Delay Sec</label><input type="number" min="0" max="59" value="${b.delaySeconds || 0}" onchange="updateBatch(${i},'delaySeconds',+this.value)"></div>
      </div>
      <div class="batch-row-preview">${getBatchRowPreview(b)}</div>
    </div>`).join('');
  updateBatchSummary();
}

function updateBatchSummary() {
  const summary = $('batchSummary');
  if (!summary) return;
  const total = contacts.length;
  const assigned = batches.reduce((acc, b) => {
    const s = parseInt(b.rowStart) || 1;
    const e = parseInt(b.rowEnd) || total;
    return acc + Math.max(0, e - s + 1);
  }, 0);
  const overlapping = detectBatchOverlaps();
  summary.innerHTML = `${batches.length} batch${batches.length !== 1 ? 'es' : ''} · ${assigned} assigned rows · ${Math.max(0, total - assigned)} unassigned${overlapping.length ? ' · <span style="color:#ef4444">⚠ Overlapping ranges detected</span>' : ''}`;
}

function validateBatches() {
  const overlaps = detectBatchOverlaps();
  if (overlaps.length) return 'Row ranges in batches overlap. Please fix before sending.';
  for (let i = 0; i < batches.length; i++) {
    const b = batches[i];
    if (!b.rowStart || !b.rowEnd) return `Batch ${i + 1}: Start Row and End Row are required.`;
    if (parseInt(b.rowStart) > parseInt(b.rowEnd)) return `Batch ${i + 1}: Start Row cannot be greater than End Row.`;
    if (!b.scheduledDate || !b.scheduledTime) return `Batch ${i + 1}: Schedule date and time are required.`;
  }
  return null;
}

boot();
