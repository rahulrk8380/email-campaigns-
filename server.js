require("dotenv").config();

const express = require("express");
const session = require("express-session");
const Database = require("better-sqlite3");
const multer = require("multer");
const XLSX = require("xlsx");
const nodemailer = require("nodemailer");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const AdmZip = require("adm-zip");

const app = express();
const PORT = Number(process.env.PORT || 3000);
const MAX = Number(process.env.MAX_RECIPIENTS_PER_CAMPAIGN || 2000);

fs.mkdirSync(path.join(__dirname, "data"), { recursive: true });
fs.mkdirSync(path.join(__dirname, "uploads"), { recursive: true });
fs.mkdirSync(path.join(__dirname, "uploads", "attachments"), { recursive: true });
fs.mkdirSync(path.join(__dirname, "uploads", "pdf-folders"), { recursive: true });

const db = new Database(path.join(__dirname, "data", "smartmail.db"));
db.pragma("journal_mode = WAL");

db.exec(`
CREATE TABLE IF NOT EXISTS contacts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  email TEXT NOT NULL,
  number TEXT DEFAULT '',
  details TEXT DEFAULT '',
  extra_json TEXT DEFAULT '{}',
  created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS campaigns (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  subject TEXT NOT NULL,
  body TEXT NOT NULL,
  status TEXT DEFAULT 'draft',
  total INTEGER DEFAULT 0,
  sent INTEGER DEFAULT 0,
  failed INTEGER DEFAULT 0,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  started_at TEXT,
  completed_at TEXT
);

CREATE TABLE IF NOT EXISTS attachments (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  original_name TEXT NOT NULL,
  stored_name TEXT NOT NULL,
  file_path TEXT NOT NULL,
  mime_type TEXT DEFAULT 'application/octet-stream',
  size INTEGER DEFAULT 0,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS campaign_attachments (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  campaign_id INTEGER NOT NULL,
  attachment_id INTEGER NOT NULL,
  FOREIGN KEY(campaign_id) REFERENCES campaigns(id),
  FOREIGN KEY(attachment_id) REFERENCES attachments(id)
);

CREATE TABLE IF NOT EXISTS campaign_recipients (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  campaign_id INTEGER NOT NULL,
  contact_id INTEGER,
  name TEXT NOT NULL,
  email TEXT NOT NULL,
  rendered_subject TEXT NOT NULL,
  rendered_body TEXT NOT NULL,
  status TEXT DEFAULT 'pending',
  error TEXT DEFAULT '',
  sent_at TEXT,
  FOREIGN KEY(campaign_id) REFERENCES campaigns(id)
);

CREATE TABLE IF NOT EXISTS settings (
  key TEXT PRIMARY KEY,
  value TEXT
);
`);

try { db.exec("ALTER TABLE campaigns ADD COLUMN sender_type TEXT DEFAULT 'microsoft'"); } catch {}
try { db.exec("ALTER TABLE campaigns ADD COLUMN sender_email TEXT DEFAULT ''"); } catch {}
try { db.exec("ALTER TABLE campaigns ADD COLUMN pdf_folder_id TEXT DEFAULT NULL"); } catch {}
try { db.exec("ALTER TABLE campaigns ADD COLUMN skipped INTEGER DEFAULT 0"); } catch {}
try { db.exec("ALTER TABLE campaigns ADD COLUMN delay_seconds INTEGER DEFAULT 5"); } catch {}
try { db.exec("ALTER TABLE campaign_recipients ADD COLUMN pdf_match_status TEXT DEFAULT 'na'"); } catch {}
try { db.prepare("ALTER TABLE campaigns ADD COLUMN scheduled_at TEXT").run(); } catch {}
try { db.exec("ALTER TABLE campaign_recipients ADD COLUMN to_emails TEXT DEFAULT ''"); } catch {}
try { db.exec("ALTER TABLE campaign_recipients ADD COLUMN cc TEXT DEFAULT ''"); } catch {}


db.exec(`
  CREATE TABLE IF NOT EXISTS templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    subject TEXT NOT NULL,
    body TEXT NOT NULL,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  );
`);

const spreadsheetUpload = multer({
  dest: path.join(__dirname, "uploads"),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = /\.(xlsx|xls|csv)$/i.test(file.originalname);
    cb(ok ? null : new Error("Only .xlsx, .xls or .csv files are allowed."), ok);
  }
});

const attachmentStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, path.join(__dirname, "uploads", "attachments")),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const safeBase = path.basename(file.originalname, ext).replace(/[^a-z0-9_-]/gi, "_").slice(0, 60) || "attachment";
    cb(null, `${Date.now()}-${Math.random().toString(36).slice(2, 10)}-${safeBase}${ext}`);
  }
});

const allowedAttachmentExt = /\.(pdf|doc|docx|xls|xlsx|ppt|pptx|txt|csv|jpg|jpeg|png|gif|webp|zip)$/i;
const attachmentUpload = multer({
  storage: attachmentStorage,
  limits: { fileSize: 15 * 1024 * 1024, files: 8 },
  fileFilter: (req, file, cb) => {
    const ok = allowedAttachmentExt.test(file.originalname);
    cb(ok ? null : new Error("Unsupported attachment type. Use PDF, Word, Excel, PowerPoint, TXT, CSV, JPG, PNG, GIF, WEBP or ZIP."), ok);
  }
});

// PDF folder upload — stores PDFs under uploads/pdf-folders/<folderId>/
const pdfFolderStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    const folderId = req.pdfFolderId || (req.pdfFolderId = `pf-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`);
    const dir = path.join(__dirname, "uploads", "pdf-folders", folderId);
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    // Keep original filename so we can match by name later
    cb(null, file.originalname);
  }
});

const pdfFolderUpload = multer({
  storage: pdfFolderStorage,
  limits: { fileSize: 50 * 1024 * 1024, files: 500 },
  fileFilter: (req, file, cb) => {
    const ok = /\.(pdf|zip)$/i.test(file.originalname);
    cb(ok ? null : new Error("Only PDF files or a ZIP of PDFs are allowed."), ok);
  }
});

// Find a PDF in folderPath matching recipientName
function findPdfForRecipient(folderPath, recipientName) {
  if (!folderPath || !fs.existsSync(folderPath)) return null;
  let files;
  try { files = fs.readdirSync(folderPath); } catch { return null; }

  const norm = s => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const target = norm(recipientName);

  // 1. Full name match: "Rahul Sharma.pdf"
  let match = files.find(f =>
    /\.pdf$/i.test(f) && norm(path.basename(f, path.extname(f))) === target
  );
  if (match) return path.join(folderPath, match);

  // 2. First word / first name match: "Rahul.pdf" matches "Rahul Sharma"
  const firstName = norm((recipientName || "").split(/\s+/)[0]);
  if (firstName && firstName !== target) {
    match = files.find(f =>
      /\.pdf$/i.test(f) && norm(path.basename(f, path.extname(f))) === firstName
    );
    if (match) return path.join(folderPath, match);
  }

  return null; // No match — this contact will be skipped
}


app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));
app.set("trust proxy", 1); // Required for Render/Heroku — tells Express the X-Forwarded-Proto header is trustworthy
app.use(session({
  secret: process.env.SESSION_SECRET || "dev-only-change-me-not-for-production",
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    sameSite: "lax",
    maxAge: 8 * 60 * 60 * 1000,                          // 8 hours
    secure: process.env.NODE_ENV === "production"          // HTTPS-only on Render
  }
}));
app.use(express.static(path.join(__dirname, "public")));

function auth(req, res, next) {
  if (!req.session.user) return res.status(401).json({ error: "Please log in." });
  next();
}


function normalizeKey(k) {
  return String(k || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

// Find a value in a row object by trying multiple normalized key aliases
function pick(row, names) {
  const keys = Object.keys(row);
  for (const wanted of names) {
    const hit = keys.find(k => normalizeKey(k) === normalizeKey(wanted));
    if (hit !== undefined) return row[hit];
  }
  return "";
}

function cleanEmail(v) {
  return String(v || "").trim().toLowerCase();
}

function validEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function getContacts() {
  return db.prepare("SELECT * FROM contacts ORDER BY id ASC").all().map(c => {
    let extras = {};
    try { extras = JSON.parse(c.extra_json || "{}"); } catch {}
    return { ...c, ...extras };
  });
}

// Fully dynamic, case-insensitive template rendering.
// Builds a lookup table: normalizedKey → value, then replaces every
// {{variable}} in the template regardless of case or spacing.
function renderTemplate(template, contact) {
  let out = String(template || "");

  // Build a flat map of ALL fields: normalizedKey → value
  const valueMap = new Map();
  const addToMap = (obj) => {
    for (const [key, value] of Object.entries(obj || {})) {
      if (key === "extra_json") continue;
      valueMap.set(normalizeKey(key), String(value ?? ""));
    }
  };

  addToMap(contact);
  // Also expand extra_json fields
  try { addToMap(JSON.parse(contact.extra_json || "{}")); } catch {}

  // Replace every {{...}} with the matched value (case-insensitive)
  out = out.replace(/\{\{([^}]+)\}\}/g, (match, varName) => {
    const key = normalizeKey(varName);
    return valueMap.has(key) ? valueMap.get(key) : match; // keep original if not found
  });

  return out;
}

// Returns a list of unresolved variables in a template given a contact
function findUnknownVariables(template, contact) {
  const valueMap = new Set();
  const addToMap = (obj) => {
    for (const key of Object.keys(obj || {})) {
      if (key === "extra_json") continue;
      valueMap.add(normalizeKey(key));
    }
  };
  addToMap(contact);
  try { addToMap(JSON.parse(contact.extra_json || "{}")); } catch {}

  const unknowns = [];
  const matches = String(template || "").matchAll(/\{\{([^}]+)\}\}/g);
  for (const m of matches) {
    if (!valueMap.has(normalizeKey(m[1]))) {
      unknowns.push(m[0]);
    }
  }
  return [...new Set(unknowns)];
}



function getAttachmentRecords(ids) {
  if (!Array.isArray(ids) || !ids.length) return [];
  const cleanIds = [...new Set(ids.map(Number).filter(Number.isInteger))];
  if (!cleanIds.length) return [];
  const placeholders = cleanIds.map(() => "?").join(",");
  return db.prepare(`SELECT id, original_name, file_path, mime_type FROM attachments WHERE id IN (${placeholders})`).all(...cleanIds)
    .filter(a => fs.existsSync(a.file_path))
    .map(a => ({ filename: a.original_name, path: a.file_path, contentType: a.mime_type }));
}

// Settings helpers
function getSetting(key, fallback = null) {
  try {
    const row = db.prepare("SELECT value FROM settings WHERE key = ?").get(key);
    return row ? row.value : fallback;
  } catch {
    return fallback;
  }
}

function setSetting(key, value) {
  db.prepare("INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(key, String(value));
}

function deleteSetting(key) {
  db.prepare("DELETE FROM settings WHERE key = ?").run(key);
}

// Microsoft OAuth & Graph API helpers
function getMicrosoftConfig(req) {
  const clientId     = process.env.MICROSOFT_CLIENT_ID     || getSetting("ms_client_id")     || "";
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET || getSetting("ms_client_secret") || "";

  // "consumers" = personal @outlook.com/@hotmail.com accounts (Azure app set to Personal accounts only)
  // "common"    = both personal AND work accounts (Azure app set to All accounts)
  // "organizations" = work/school accounts only
  // Default to "consumers" because most SmartMail users use personal Outlook accounts.
  // If MICROSOFT_TENANT_ID is set to "common" in env but app is Personal-only, auto-correct to "consumers".
  let tenantId = process.env.MICROSOFT_TENANT_ID || getSetting("ms_tenant_id") || "consumers";
  if (!tenantId || tenantId.trim() === "" || tenantId.trim() === "common") {
    // "common" fails when Azure app is set to "Personal accounts only" — use "consumers" instead
    tenantId = "consumers";
  }

  // Auto-build redirect URI from the ACTUAL request host (works on both localhost and Render)
  let redirectUri;
  if (req) {
    const proto = req.protocol || (req.secure ? "https" : "http");
    const host  = req.get("host") || "localhost:3000";
    redirectUri = `${proto}://${host}/auth/microsoft/callback`;
  } else {
    redirectUri = process.env.MICROSOFT_REDIRECT_URI || getSetting("ms_redirect_uri") || `http://localhost:${PORT}/auth/microsoft/callback`;
  }

  return {
    clientId:     String(clientId     || "").trim(),
    clientSecret: String(clientSecret || "").trim(),
    tenantId:     String(tenantId     || "consumers").trim(),
    redirectUri:  String(redirectUri  || "").trim()
  };
}

async function getValidMicrosoftAccessToken() {
  const raw = getSetting("ms_tokens");
  if (!raw) {
    throw new Error("Microsoft Outlook account is not connected. Please connect your account first.");
  }
  let tokens;
  try {
    tokens = JSON.parse(raw);
  } catch {
    throw new Error("Corrupted Microsoft token storage. Please reconnect your Outlook account.");
  }

  // If token is valid for at least 3 more minutes, return it
  if (tokens.access_token && tokens.expires_at && tokens.expires_at > Date.now() + 3 * 60 * 1000) {
    return tokens.access_token;
  }

  if (!tokens.refresh_token) {
    throw new Error("No refresh token available. Please reconnect your Microsoft Outlook account.");
  }

  const config = getMicrosoftConfig();
  if (!config.clientId) {
    throw new Error("Microsoft Client ID is missing. Configure it in .env or Settings.");
  }

  const refreshParams = new URLSearchParams({
    client_id: config.clientId,
    grant_type: "refresh_token",
    refresh_token: tokens.refresh_token,
    scope: "offline_access User.Read Mail.Send"
  });
  if (config.clientSecret) {
    refreshParams.append("client_secret", config.clientSecret);
  }

  const refreshRes = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: refreshParams.toString()
  });

  const data = await refreshRes.json();
  if (!refreshRes.ok || data.error) {
    throw new Error(data.error_description || data.error || "Failed to refresh Microsoft token. Please reconnect Outlook.");
  }

  tokens.access_token = data.access_token;
  if (data.refresh_token) tokens.refresh_token = data.refresh_token;
  tokens.expires_at = Date.now() + (Number(data.expires_in) || 3600) * 1000;

  setSetting("ms_tokens", JSON.stringify(tokens));
  return tokens.access_token;
}

async function sendMailViaMicrosoftGraph({ to, toEmails, cc, subject, body, attachments = [] }) {
  const accessToken = await getValidMicrosoftAccessToken();

  const graphAttachments = [];
  for (const att of attachments) {
    const filePath = att.file_path || att.path;
    if (filePath && fs.existsSync(filePath)) {
      const content = fs.readFileSync(filePath);
      graphAttachments.push({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: att.original_name || att.filename || path.basename(filePath),
        contentType: att.mime_type || att.contentType || "application/octet-stream",
        contentBytes: content.toString("base64")
      });
    }
  }

  const isHtml = /<[a-z][\s\S]*>/i.test(body);

  // Build To recipients — support array (new multi-email) or legacy single string
  const toList = Array.isArray(toEmails) && toEmails.length
    ? toEmails
    : (to ? String(to).split(",").map(s => s.trim()).filter(Boolean) : []);

  const toRecipients = toList.filter(validEmail).map(addr => ({
    emailAddress: { address: addr }
  }));

  if (!toRecipients.length) {
    throw new Error("No valid To recipients for this email.");
  }

  // Build CC recipients
  const ccList = String(cc || "").split(/[,;]/).map(s => s.trim()).filter(s => s && validEmail(s));
  const ccRecipients = ccList.map(addr => ({ emailAddress: { address: addr } }));

  const payload = {
    message: {
      subject: String(subject || ""),
      body: {
        contentType: isHtml ? "HTML" : "Text",
        content: String(body || "")
      },
      toRecipients,
      ccRecipients,
      attachments: graphAttachments
    },
    saveToSentItems: true
  };

  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (res.status === 202) {
    return { ok: true, messageId: `Graph-${Date.now()}` };
  }

  const errData = await res.json().catch(() => ({}));
  const msg = errData?.error?.message || `HTTP ${res.status}: ${res.statusText}`;
  throw new Error(`Microsoft Graph error: ${msg}`);
}

function getTransporter() {
  if (!process.env.SMTP_HOST || !process.env.SMTP_USER || !process.env.SMTP_PASS) {
    throw new Error("SMTP is not configured. Open .env and add SMTP_HOST, SMTP_USER and SMTP_PASS.");
  }

  return nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 465),
    secure: String(process.env.SMTP_SECURE || "true").toLowerCase() === "true",
    family: 4,
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    }
  });
}

// Convert plain text body to clean HTML — improves inbox deliverability
function textToHtml(text) {
  const escaped = String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  const paragraphs = escaped.split(/\n{2,}/).map(p =>
    `<p style="margin:0 0 14px 0;line-height:1.6">${p.replace(/\n/g, "<br>")}</p>`
  ).join("");
  return `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;background:#fff;padding:24px;max-width:600px;margin:auto">
${paragraphs}
</body></html>`;
}

// Build anti-spam compliant mail options for SMTP
function buildSmtpMailOptions({ from, to, toEmails, cc, subject, body, attachments = [] }) {
  const fromName = getSetting("smtp_from_name") || process.env.MAIL_FROM_NAME || process.env.SMTP_USER || "SmartMail";
  const fromAddr = getSetting("smtp_from") || process.env.MAIL_FROM || process.env.SMTP_USER;
  const isHtml = /^<!DOCTYPE|^<html/i.test(String(body || "").trim());
  const htmlBody = isHtml ? body : textToHtml(body);
  const textBody = isHtml
    ? String(body).replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim()
    : String(body || "");

  // Build To — support array (multi-email) or single string
  const toList = Array.isArray(toEmails) && toEmails.length
    ? toEmails.filter(validEmail)
    : (to ? String(to).split(",").map(s => s.trim()).filter(s => s && validEmail(s)) : []);

  if (!toList.length) throw new Error("No valid To recipients for this email.");

  const toField = toList.join(", ");

  // Build CC
  const ccList = String(cc || "").split(/[,;]/).map(s => s.trim()).filter(s => s && validEmail(s));
  const ccField = ccList.length ? ccList.join(", ") : undefined;

  const domain = (fromAddr || "").split("@")[1] || "localhost";
  const messageId = `<${Date.now()}.${Math.random().toString(36).slice(2)}@${domain}>`;

  const opts = {
    from: `"${fromName}" <${fromAddr}>`,
    replyTo: fromAddr,
    to: toField,
    subject,
    text: textBody,
    html: htmlBody,
    attachments,
    messageId,
    headers: {
      "X-Mailer": "SmartMail-Pro",
      "X-Priority": "3",
      "Mime-Version": "1.0",
      "List-Unsubscribe": `<mailto:${fromAddr}?subject=Unsubscribe>`
    }
  };
  if (ccField) opts.cc = ccField;
  return opts;
}



let sendLock = false;

// Microsoft OAuth Endpoints
app.get("/auth/microsoft/login", (req, res) => {
  const config = getMicrosoftConfig(req);
  if (!config.clientId) {
    return res.redirect("/?error=microsoft_client_id_missing");
  }

  const state = crypto.randomBytes(16).toString("hex");
  req.session.ms_oauth_state = state;

  const scopes = "offline_access User.Read Mail.Send";
  const authUrl = `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/authorize?` +
    new URLSearchParams({
      client_id: config.clientId,
      response_type: "code",
      redirect_uri: config.redirectUri,
      response_mode: "query",
      scope: scopes,
      state: state,
      prompt: "select_account"
    }).toString();

  res.redirect(authUrl);
});

app.get("/auth/microsoft/callback", async (req, res) => {
  const { code, state, error, error_description } = req.query;

  if (error) {
    console.error("Microsoft OAuth Callback Error:", error, error_description);
    return res.redirect(`/?error=${encodeURIComponent(error_description || error)}`);
  }

  if (!code) {
    return res.redirect("/?error=no_authorization_code_received");
  }

  const config = getMicrosoftConfig(req);
  if (!config.clientId) {
    return res.redirect("/?error=microsoft_client_id_missing");
  }

  try {
    const tokenParams = new URLSearchParams({
      client_id: config.clientId,
      grant_type: "authorization_code",
      code: String(code),
      redirect_uri: config.redirectUri,
      scope: "offline_access User.Read Mail.Send"
    });
    if (config.clientSecret) {
      tokenParams.append("client_secret", config.clientSecret);
    }

    const tokenRes = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: tokenParams.toString()
    });

    const tokenData = await tokenRes.json();
    if (!tokenRes.ok || tokenData.error) {
      throw new Error(tokenData.error_description || tokenData.error || "Failed to exchange token with Microsoft");
    }

    // Retrieve user profile to determine email
    const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${tokenData.access_token}` }
    });
    const profile = await profileRes.json();
    const email = profile.mail || profile.userPrincipalName || "";
    const name = profile.displayName || email;

    const tokenRecord = {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expires_at: Date.now() + (Number(tokenData.expires_in) || 3600) * 1000,
      email,
      name,
      scope: tokenData.scope
    };

    setSetting("ms_tokens", JSON.stringify(tokenRecord));
    setSetting("ms_account_email", email);
    setSetting("ms_account_name", name);

    res.redirect("/?microsoft=connected");
  } catch (err) {
    console.error("Microsoft token exchange failed:", err);
    res.redirect(`/?error=${encodeURIComponent(err.message)}`);
  }
});

app.get("/api/microsoft/status", (req, res) => {
  const config = getMicrosoftConfig(req);
  const raw = getSetting("ms_tokens");
  let connected = false;
  let email = getSetting("ms_account_email") || null;
  let name = getSetting("ms_account_name") || null;

  if (raw) {
    try {
      const tokens = JSON.parse(raw);
      if (tokens.access_token || tokens.refresh_token) {
        connected = true;
        if (!email) email = tokens.email || null;
        if (!name) name = tokens.name || null;
      }
    } catch {}
  }

  res.json({
    connected,
    email,
    name,
    clientIdConfigured: Boolean(config.clientId),
    config: {
      clientId: config.clientId ? `${config.clientId.slice(0, 8)}...` : "",
      tenantId: config.tenantId,
      redirectUri: config.redirectUri
    }
  });
});

app.post("/api/microsoft/disconnect", auth, (req, res) => {
  deleteSetting("ms_tokens");
  deleteSetting("ms_account_email");
  deleteSetting("ms_account_name");
  res.json({ ok: true, message: "Microsoft account disconnected." });
});

app.post("/api/microsoft/config", auth, (req, res) => {
  const { clientId, clientSecret, tenantId } = req.body || {};
  if (clientId !== undefined) setSetting("ms_client_id", String(clientId).trim());
  if (clientSecret !== undefined) setSetting("ms_client_secret", String(clientSecret).trim());
  if (tenantId !== undefined) setSetting("ms_tenant_id", String(tenantId).trim() || "common");
  res.json({ ok: true, message: "Microsoft OAuth configuration updated." });
});

app.post("/api/microsoft/check", auth, async (req, res) => {
  try {
    const accessToken = await getValidMicrosoftAccessToken();
    const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    const profile = await profileRes.json();
    if (!profileRes.ok) throw new Error(profile?.error?.message || "Failed to fetch Microsoft profile");
    res.json({ ok: true, email: profile.mail || profile.userPrincipalName, name: profile.displayName });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.get("/api/senders", auth, (req, res) => {
  const senders = [];
  const msEmail = getSetting("ms_account_email");
  const msTokens = getSetting("ms_tokens");
  const isMsConnected = Boolean(msTokens && msEmail);

  if (isMsConnected) {
    senders.push({
      id: "microsoft",
      type: "microsoft",
      label: `Microsoft Outlook (${msEmail})`,
      email: msEmail,
      recommended: true
    });
  }

  if (process.env.SMTP_HOST && process.env.SMTP_USER) {
    const smtpEmail = process.env.MAIL_FROM || process.env.SMTP_USER;
    senders.push({
      id: "smtp",
      type: "smtp",
      label: `Custom SMTP (${smtpEmail})`,
      email: smtpEmail,
      recommended: !isMsConnected
    });
  }

  res.json({
    senders,
    defaultSender: isMsConnected ? "microsoft" : (senders[0]?.id || "microsoft"),
    msConnected: isMsConnected
  });
});

app.post("/api/login", (req, res) => {
  const { email, password } = req.body || {};
  const adminEmail = process.env.ADMIN_EMAIL || "admin@example.com";
  const adminPassword = process.env.ADMIN_PASSWORD || "change-me";

  if (email !== adminEmail || password !== adminPassword) {
    return res.status(401).json({ error: "Invalid admin credentials." });
  }

  req.session.user = { email };
  res.json({ ok: true, email });
});

app.post("/api/logout", (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

app.get("/api/me", (req, res) => {
  res.json({ loggedIn: !!req.session.user, email: req.session.user?.email || null });
});

app.get("/api/dashboard", auth, (req, res) => {
  const contacts = db.prepare("SELECT COUNT(*) c FROM contacts").get().c;
  const campaigns = db.prepare("SELECT COUNT(*) c FROM campaigns").get().c;
  const sent = db.prepare("SELECT COALESCE(SUM(sent),0) c FROM campaigns").get().c;
  const failed = db.prepare("SELECT COALESCE(SUM(failed),0) c FROM campaigns").get().c;
  const pending = db.prepare("SELECT COUNT(*) c FROM campaign_recipients WHERE status='pending'").get().c;
  res.json({ contacts, campaigns, sent, failed, pending });
});

app.get("/api/contacts", auth, (req, res) => res.json(getContacts()));

// ── SMTP Settings (read + update from UI) ─────────────────────────────────────
app.get("/api/smtp-settings", auth, (req, res) => {
  res.json({
    host: process.env.SMTP_HOST || "",
    port: process.env.SMTP_PORT || "587",
    secure: process.env.SMTP_SECURE || "false",
    user: process.env.SMTP_USER || "",
    from: process.env.MAIL_FROM || process.env.SMTP_USER || "",
    fromName: process.env.MAIL_FROM_NAME || ""
  });
});

// Save SMTP settings into DB so they survive without editing .env
app.post("/api/smtp-settings", auth, (req, res) => {
  const { host, port, secure, user, pass, from, fromName } = req.body || {};
  if (host !== undefined) setSetting("smtp_host", host);
  if (port !== undefined) setSetting("smtp_port", String(port));
  if (secure !== undefined) setSetting("smtp_secure", String(secure));
  if (user !== undefined) setSetting("smtp_user", user);
  if (pass && pass.trim()) setSetting("smtp_pass", pass.trim());
  if (from !== undefined) setSetting("smtp_from", from);
  if (fromName !== undefined) setSetting("smtp_from_name", fromName);
  res.json({ ok: true, message: "SMTP settings saved." });
});

// Send a verification OTP — tests the new credentials directly
let smtpOtpStore = {}; // { email: { otp, expires, pass, fromName } }

app.post("/api/smtp-verify/send-otp", auth, async (req, res) => {
  const { email, pass, fromName } = req.body || {};
  if (!email) return res.status(400).json({ error: "Email is required." });
  if (!pass)  return res.status(400).json({ error: "Password is required." });

  const otp = String(Math.floor(100000 + Math.random() * 900000));

  // Use the saved host (set just before by /api/smtp-settings), fall back to Outlook
  const smtpHost   = getSetting("smtp_host")   || process.env.SMTP_HOST   || "smtp.office365.com";
  const smtpPort   = Number(getSetting("smtp_port")   || process.env.SMTP_PORT   || 587);
  const smtpSecure = String(getSetting("smtp_secure") || process.env.SMTP_SECURE || "false").toLowerCase() === "true";

  try {
    const transporter = nodemailer.createTransport({
      host: smtpHost, port: smtpPort, secure: smtpSecure, family: 4,
      auth: { user: email, pass }
    });

    await transporter.sendMail({
      from: fromName ? `"${fromName}" <${email}>` : email,
      to: email,
      subject: "SmartMail Pro — Email Verification Code",
      text: `Your SmartMail Pro verification code is:\n\n${otp}\n\nThis code expires in 5 minutes.\nIf you did not request this, ignore this email.`,
      html: `<div style="font-family:sans-serif;max-width:400px;margin:auto;padding:32px;border-radius:12px;border:1px solid #e5e7eb">
        <h2 style="margin:0 0 16px;color:#111">SmartMail Pro</h2>
        <p style="color:#6b7280;margin:0 0 24px">Your email verification code:</p>
        <div style="font-size:36px;font-weight:900;letter-spacing:12px;font-family:monospace;color:#635bff;background:#f5f3ff;padding:16px;border-radius:10px;text-align:center">${otp}</div>
        <p style="color:#9ca3af;font-size:12px;margin-top:20px">Expires in 5 minutes. Ignore if you didn't request this.</p>
      </div>`
    });

    // Store OTP + credentials (saved permanently only after correct code is entered)
    smtpOtpStore[email] = { otp, expires: Date.now() + 5 * 60 * 1000, pass, fromName: fromName || "" };
    res.json({ ok: true, message: `Code sent to ${email}. Check your inbox.` });
  } catch (e) {
    res.status(400).json({ error: `Connection failed: ${e.message}` });
  }
});

app.post("/api/smtp-verify/confirm-otp", auth, (req, res) => {
  const { email, otp } = req.body || {};
  const record = smtpOtpStore[email];
  if (!record) return res.status(400).json({ error: "No OTP found. Please request a new code." });
  if (Date.now() > record.expires) {
    delete smtpOtpStore[email];
    return res.status(400).json({ error: "OTP expired. Please request a new code." });
  }
  if (String(otp).trim() !== record.otp) {
    return res.status(400).json({ error: "Incorrect code. Try again." });
  }

  // ✅ Verified — save as the active SMTP sender
  setSetting("smtp_host",      "smtp.office365.com");
  setSetting("smtp_port",      "587");
  setSetting("smtp_secure",    "false");
  setSetting("smtp_user",      email);
  setSetting("smtp_pass",      record.pass);
  setSetting("smtp_from",      email);
  setSetting("smtp_from_name", record.fromName);
  setSetting("smtp_verified",  "true");

  delete smtpOtpStore[email];
  res.json({ ok: true, email, message: `✓ Switched! Campaigns will now send from ${email}.` });

});

app.delete("/api/contacts", auth, (req, res) => {
  db.prepare("DELETE FROM contacts").run();
  res.json({ ok: true });
});

app.post("/api/upload", auth, spreadsheetUpload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No Excel file uploaded." });

  try {
    const workbook = XLSX.readFile(req.file.path, { cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) throw new Error("The first sheet is empty.");

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    if (!rows.length) throw new Error("The first sheet is empty.");

    // Collect all unique column headers from the Excel file (for dynamic variables)
    const allHeaders = [];
    const seenHeaders = new Set();
    for (const row of rows) {
      for (const key of Object.keys(row)) {
        if (!seenHeaders.has(key)) {
          seenHeaders.add(key);
          allHeaders.push(key);
        }
      }
    }

    const insert = db.prepare(`
      INSERT INTO contacts (name,email,number,details,extra_json)
      VALUES (?,?,?,?,?)
    `);

    const clear = db.prepare("DELETE FROM contacts");

    const result = db.transaction(items => {
      clear.run();
      let added = 0;
      let invalid = 0;
      let invalidNoEmail = 0;

      for (let rowIdx = 0; rowIdx < items.length; rowIdx++) {
        const row = items[rowIdx];

        // ── Determine row identifier (name / loan number) ──
        const name = String(
          pick(row, ["loan number","loannumber","loan_number","loanno","loan no","name","full name","fullname","customer name","customername"]) || ""
        ).trim() || `Row ${rowIdx + 1}`;

        // ── Collect Email1..Email9 ──
        const emailCols = [];
        for (let i = 1; i <= 9; i++) {
          const val = cleanEmail(pick(row, [`email${i}`, `email ${i}`, `e-mail${i}`, `e-mail ${i}`]));
          if (val) emailCols.push(val);
        }
        // Fallback: if no EmailN columns, try plain "email"/"mail" column
        if (!emailCols.length) {
          const fallback = cleanEmail(pick(row, ["email","mail","email address","emailaddress"]));
          if (fallback) emailCols.push(fallback);
        }

        // Filter to valid emails
        const validEmails = emailCols.filter(validEmail);

        // ── CC columns: CC1..CC10 (same pattern as Email1..Email9) ──
        const ccCols = [];
        for (let i = 1; i <= 10; i++) {
          const val = cleanEmail(pick(row, [`cc${i}`, `cc ${i}`, `c-c${i}`]));
          if (val) ccCols.push(val);
        }
        // Fallback: plain "CC" column — supports semicolon/comma-separated multiple emails
        if (!ccCols.length) {
          const ccRaw = String(pick(row, ["cc", "carbon copy", "carboncopy"]) || "").trim();
          if (ccRaw) {
            ccRaw.split(/[,;]/).map(s => s.trim()).filter(s => s && validEmail(s))
                 .forEach(e => ccCols.push(e));
          }
        }
        const ccEmails = ccCols.filter(e => validEmail(e));

        // ── Other standard fields ──
        const number = String(pick(row, ["number","phone","mobile","phone number","mobile number"]) || "").trim();
        const details = String(pick(row, ["details","detail","description"]) || "").trim();

        // ── Store ALL columns in extra_json (for dynamic variables) ──
        // Every column header becomes a variable, including Email1, Email2, CC, etc.
        const extras = {};
        for (const [k, v] of Object.entries(row)) {
          extras[k] = v == null ? "" : String(v);
        }
        // Also store structured email list and cc for campaign processing
        extras["__emails"] = JSON.stringify(validEmails);
        extras["__cc"] = ccEmails.join(";");
        extras["__allEmails"] = validEmails.join(",");
        extras["__hasValidEmail"] = String(validEmails.length > 0);

        // Primary email = Email1 (first valid email)
        const primaryEmail = validEmails[0] || "";

        // ── Validation flag ──
        if (!validEmails.length) {
          // Row has no valid email — keep it visible but mark as invalid
          extras["__invalid"] = "No valid email recipient";
          invalidNoEmail++;
        }

        // Note: we insert ALL rows (including those with no valid email)
        // The __invalid flag controls whether they get sent or not
        insert.run(name, primaryEmail, number, details, JSON.stringify(extras));
        if (validEmails.length > 0) added++;
        else invalid++;
      }

      return { added, invalid: invalidNoEmail, total: items.length, headers: allHeaders };
    })(rows);

    fs.unlink(req.file.path, () => {});
    res.json({ ok: true, ...result, contacts: getContacts() });
  } catch (e) {
    fs.unlink(req.file.path, () => {});
    res.status(400).json({ error: e.message });
  }
});



app.post("/api/attachments", auth, (req, res) => {
  attachmentUpload.array("files", 8)(req, res, err => {
    if (err) return res.status(400).json({ error: err.message || "Attachment upload failed." });
    const files = req.files || [];
    if (!files.length) return res.status(400).json({ error: "No attachment files selected." });

    const totalSize = files.reduce((sum, f) => sum + Number(f.size || 0), 0);
    if (totalSize > 24 * 1024 * 1024) {
      files.forEach(f => fs.unlink(f.path, () => {}));
      return res.status(400).json({ error: "Attachments are too large. Keep the total below 24 MB for reliable email sending." });
    }

    const insert = db.prepare(`
      INSERT INTO attachments (original_name, stored_name, file_path, mime_type, size)
      VALUES (?,?,?,?,?)
    `);
    const attachments = files.map(file => {
      const id = insert.run(
        file.originalname,
        file.filename,
        file.path,
        file.mimetype || "application/octet-stream",
        file.size
      ).lastInsertRowid;
      return { id, name: file.originalname, size: file.size, type: file.mimetype || "application/octet-stream" };
    });

    res.json({ ok: true, attachments });
  });
});

app.get("/api/attachments/:id", auth, (req, res) => {
  const a = db.prepare("SELECT * FROM attachments WHERE id=?").get(req.params.id);
  if (!a || !fs.existsSync(a.file_path)) return res.status(404).json({ error: "Attachment not found." });
  res.download(a.file_path, a.original_name);
});

app.delete("/api/attachments/:id", auth, (req, res) => {
  const a = db.prepare("SELECT * FROM attachments WHERE id=?").get(req.params.id);
  if (!a) return res.status(404).json({ error: "Attachment not found." });
  const used = db.prepare("SELECT COUNT(*) c FROM campaign_attachments WHERE attachment_id=?").get(a.id).c;
  if (used) return res.status(400).json({ error: "This attachment is already linked to a campaign and cannot be removed." });
  db.prepare("DELETE FROM attachments WHERE id=?").run(a.id);
  fs.unlink(a.file_path, () => {});
  res.json({ ok: true });
});

app.post("/api/pdf-folder", auth, (req, res) => {
  pdfFolderUpload.array("files", 500)(req, res, err => {
    if (err) return res.status(400).json({ error: err.message || "PDF folder upload failed." });
    const files = req.files || [];
    if (!files.length) return res.status(400).json({ error: "No files uploaded." });
    const folderId = req.pdfFolderId;
    if (!folderId) return res.status(500).json({ error: "Folder ID could not be determined." });

    const folderPath = path.join(__dirname, "uploads", "pdf-folders", folderId);
    let extractedPdfNames = [];

    // Extract any ZIP files found, pull out PDFs
    for (const file of files) {
      if (/\.zip$/i.test(file.originalname)) {
        try {
          const zip = new AdmZip(file.path);
          const entries = zip.getEntries();
          for (const entry of entries) {
            if (!entry.isDirectory && /\.pdf$/i.test(entry.entryName)) {
              const pdfName = path.basename(entry.entryName);
              const destPath = path.join(folderPath, pdfName);
              zip.extractEntryTo(entry, folderPath, false, true);
              extractedPdfNames.push(pdfName);
            }
          }
        } catch (e) {
          console.error("ZIP extraction error:", e);
        }
        // Remove the ZIP file after extraction
        fs.unlink(file.path, () => {});
      }
    }

    // Collect all PDF filenames now in the folder (directly uploaded + extracted from ZIP)
    let allPdfNames = [];
    try {
      allPdfNames = fs.readdirSync(folderPath).filter(f => /\.pdf$/i.test(f));
    } catch {}

    if (!allPdfNames.length) {
      return res.status(400).json({ error: "No PDF files found. Upload PDF files or a ZIP containing PDFs." });
    }

    res.json({
      ok: true,
      folderId,
      fileCount: allPdfNames.length,
      filenames: allPdfNames
    });
  });
});

app.delete("/api/pdf-folder/:folderId", auth, (req, res) => {
  const folderId = req.params.folderId;
  if (!folderId || folderId.includes("..") || !folderId.startsWith("pf-")) {
    return res.status(400).json({ error: "Invalid folder ID." });
  }
  const folderPath = path.join(__dirname, "uploads", "pdf-folders", folderId);
  fs.rm(folderPath, { recursive: true, force: true }, () => {});
  res.json({ ok: true });
});



app.post("/api/test-smtp", auth, async (req, res) => {
  try {
    const transporter = getTransporter();
    await transporter.verify();
    res.json({ ok: true, message: "SMTP connection and authentication are working." });
  } catch (e) {
    res.status(400).json({ error: `SMTP check failed: ${e.message}` });
  }
});

app.post("/api/test-email", auth, async (req, res) => {
  const { to, subject, body, sampleName = "there", attachmentIds = [], senderType } = req.body || {};

  if (!to || !validEmail(to)) {
    return res.status(400).json({ error: "Enter a valid test email address." });
  }
  if (!subject || !body) {
    return res.status(400).json({ error: "Subject and message body are required." });
  }

  const renderedSubject = String(subject).replaceAll("{{name}}", sampleName);
  const renderedBody = String(body)
    .replaceAll("{{name}}", sampleName)
    .replaceAll("{{email}}", to);

  const testAttachments = getAttachmentRecords(attachmentIds);
  const msTokens = getSetting("ms_tokens");
  const isMsConnected = Boolean(msTokens);
  const useMicrosoft = senderType === "microsoft" || (!senderType && isMsConnected);

  try {
    if (useMicrosoft) {
      const result = await sendMailViaMicrosoftGraph({
        to,
        subject: renderedSubject,
        body: renderedBody,
        attachments: testAttachments
      });
      return res.json({
        ok: true,
        message: `Test email sent successfully via Microsoft Outlook to ${to}.`
      });
    } else {
      const transporter = getTransporter();
      const info = await transporter.sendMail(buildSmtpMailOptions({
        to,
        subject: renderedSubject,
        body: renderedBody,
        attachments: testAttachments
      }));

      return res.json({
        ok: true,
        message: `Test email accepted by SMTP. Message ID: ${info.messageId || "created"}`
      });
    }
  } catch (e) {
    res.status(400).json({ error: `Test email failed: ${e.message}` });
  }
});

app.post("/api/campaigns", auth, (req, res) => {
  const { name, subject, body, selectedIds, attachmentIds = [], senderType, pdfFolderId, delaySeconds } = req.body || {};
  const scheduledAt = req.body.scheduledAt || null; // ISO datetime string

  if (!subject || !body) {
    return res.status(400).json({ error: "Subject and message body are required." });
  }

  // Validate delay: between 1s and 600s
  const cleanDelay = Math.max(1, Math.min(600, Number.isFinite(Number(delaySeconds)) ? Number(delaySeconds) : 5));


  let contacts = getContacts();

  if (Array.isArray(selectedIds) && selectedIds.length) {
    const ids = new Set(selectedIds.map(Number));
    contacts = contacts.filter(c => ids.has(c.id));
  }

  if (!contacts.length) return res.status(400).json({ error: "No recipients selected." });
  if (contacts.length > MAX) {
    return res.status(400).json({ error: `Campaign exceeds the configured limit of ${MAX} recipients.` });
  }

  // Validate pdfFolderId if provided
  const cleanPdfFolderId = (pdfFolderId && String(pdfFolderId).startsWith("pf-") && !String(pdfFolderId).includes(".."))
    ? String(pdfFolderId) : null;

  const pdfFolderPath = cleanPdfFolderId
    ? path.join(__dirname, "uploads", "pdf-folders", cleanPdfFolderId)
    : null;

  const usePdfFolder = Boolean(pdfFolderPath && fs.existsSync(pdfFolderPath));

  const msEmail = getSetting("ms_account_email") || "";
  const isMsConnected = Boolean(getSetting("ms_tokens"));
  const finalSenderType = senderType || (isMsConnected ? "microsoft" : "smtp");
  const finalSenderEmail = finalSenderType === "microsoft" ? msEmail : (process.env.MAIL_FROM || process.env.SMTP_USER || "");

  const create = db.prepare(`
    INSERT INTO campaigns (name,subject,body,total,sender_type,sender_email,pdf_folder_id,skipped,delay_seconds,scheduled_at)
    VALUES (?,?,?,?,?,?,?,?,?,?)
  `);

  const insert = db.prepare(`
    INSERT INTO campaign_recipients
    (campaign_id,contact_id,name,email,to_emails,cc,rendered_subject,rendered_body,status,pdf_match_status)
    VALUES (?,?,?,?,?,?,?,?,?,?)
  `);

  let matchedCount = 0;
  let skippedCount = 0;
  let invalidCount = 0;

  const campaignId = db.transaction(() => {
    const id = create.run(
      String(name || "Untitled Campaign"),
      String(subject),
      String(body),
      contacts.length,
      finalSenderType,
      finalSenderEmail,
      cleanPdfFolderId || null,
      0,
      cleanDelay,
      scheduledAt
    ).lastInsertRowid;

    for (const c of contacts) {
      let status = "pending";
      let pdfMatchStatus = "na";

      // Extract emails from extra_json
      let extras = {};
      try { extras = JSON.parse(c.extra_json || "{}"); } catch {}

      // Check if row was flagged invalid (no valid email)
      if (extras["__invalid"]) {
        status = "not_sent";
        pdfMatchStatus = "skipped";
        invalidCount++;
      } else if (usePdfFolder) {
        const pdfPath = findPdfForRecipient(pdfFolderPath, c.name);
        if (pdfPath) {
          pdfMatchStatus = "matched";
          matchedCount++;
        } else {
          status = "not_sent";
          pdfMatchStatus = "skipped";
          skippedCount++;
        }
      }

      // Build to_emails: parse from stored __emails or fallback to primary email
      let toEmails = [];
      try { toEmails = JSON.parse(extras["__emails"] || "[]"); } catch {}
      if (!toEmails.length && c.email && validEmail(c.email)) toEmails = [c.email];
      const toEmailsStr = toEmails.join(",");

      // CC from stored __cc
      const cc = String(extras["__cc"] || "");

      insert.run(
        id,
        c.id,
        c.name,
        c.email,     // primary email for display
        toEmailsStr, // all To emails comma-joined
        cc,          // CC emails semicolon-joined
        renderTemplate(subject, c),
        renderTemplate(body, c),
        status,
        pdfMatchStatus
      );
    }

    // Update skipped count on the campaign (includes invalid rows)
    const totalSkipped = skippedCount + invalidCount;
    if (totalSkipped > 0) {
      db.prepare("UPDATE campaigns SET skipped=? WHERE id=?").run(totalSkipped, id);
    }

    return id;
  })();


  const cleanAttachmentIds = [...new Set((Array.isArray(attachmentIds) ? attachmentIds : []).map(Number).filter(Number.isInteger))];
  const linkAttachment = db.prepare("INSERT INTO campaign_attachments (campaign_id, attachment_id) VALUES (?,?)");
  for (const attachmentId of cleanAttachmentIds) {
    const exists = db.prepare("SELECT id FROM attachments WHERE id=?").get(attachmentId);
    if (exists) linkAttachment.run(campaignId, attachmentId);
  }

  // Detect duplicate loan numbers among matched recipients (same loan number → same PDF sent twice)
  const allMatched = db.prepare(`
    SELECT name, GROUP_CONCAT(id) as ids, COUNT(*) as cnt
    FROM campaign_recipients
    WHERE campaign_id=? AND pdf_match_status='matched'
    GROUP BY name
    HAVING COUNT(*) > 1
  `).all(campaignId);

  const duplicates = allMatched.map(row => ({
    loanNumber: row.name,
    count: row.cnt,
    recipientIds: row.ids.split(",").map(Number)
  }));

  // If scheduledAt is set, do NOT start send loop — just return scheduled response
  if (scheduledAt) {
    return res.json({ campaignId, scheduled: true, scheduledAt });
  }

  res.json({
    ok: true,
    campaignId,
    attachments: cleanAttachmentIds,
    senderType: finalSenderType,
    matchedCount: usePdfFolder ? matchedCount : null,
    skippedCount: usePdfFolder ? skippedCount : null,
    usedPdfFolder: usePdfFolder,
    duplicates  // array of {loanNumber, count, recipientIds} — empty if no duplicates
  });
});

// Mark specific recipients as not_sent (denied duplicates) before sending
app.post("/api/campaigns/:id/deny-duplicates", auth, (req, res) => {
  const { recipientIds } = req.body || {};
  if (!Array.isArray(recipientIds) || !recipientIds.length) {
    return res.status(400).json({ error: "recipientIds array required." });
  }

  const campaign = db.prepare("SELECT * FROM campaigns WHERE id=?").get(req.params.id);
  if (!campaign) return res.status(404).json({ error: "Campaign not found." });

  const markNotSent = db.prepare(`
    UPDATE campaign_recipients
    SET status='not_sent', pdf_match_status='skipped', error='Denied — duplicate loan number'
    WHERE id=? AND campaign_id=?
  `);

  let denied = 0;
  for (const rid of recipientIds.map(Number)) {
    const r = markNotSent.run(rid, campaign.id);
    if (r.changes) denied++;
  }

  // Update skipped count
  db.prepare("UPDATE campaigns SET skipped=skipped+? WHERE id=?").run(denied, campaign.id);

  res.json({ ok: true, denied });
});


// Reusable send function — used by the /send route and the scheduler
async function sendCampaign(campaignId) {
  const campaign = db.prepare("SELECT * FROM campaigns WHERE id=?").get(campaignId);
  if (!campaign) throw new Error(`Campaign ${campaignId} not found`);

  if (campaign.status === "sending" || campaign.status === "completed") return;

  const useMicrosoft = campaign.sender_type === "microsoft" || (!campaign.sender_type && Boolean(getSetting("ms_tokens")));

  sendLock = true;

  db.prepare(`
    UPDATE campaigns
    SET status='sending', started_at=CURRENT_TIMESTAMP, sent=0, failed=0
    WHERE id=?
  `).run(campaign.id);

  // Only reset recipients that are actually pending — leave 'not_sent' (skipped) ones untouched
  db.prepare(`
    UPDATE campaign_recipients
    SET status='pending', error='', sent_at=NULL
    WHERE campaign_id=? AND pdf_match_status != 'skipped'
  `).run(campaign.id);

  try {
    let transporter = null;
    if (!useMicrosoft) {
      transporter = getTransporter();
    }

    const pending = db.prepare(`
      SELECT *
      FROM campaign_recipients
      WHERE campaign_id=? AND status='pending'
      ORDER BY id
    `).all(campaign.id);

    for (const recipient of pending) {
      try {
        // Global attachments — same for every recipient
        const globalAttachments = db.prepare(`
          SELECT a.original_name, a.file_path, a.mime_type
          FROM campaign_attachments ca
          JOIN attachments a ON a.id=ca.attachment_id
          WHERE ca.campaign_id=?
        `).all(campaign.id).filter(a => fs.existsSync(a.file_path)).map(a => ({
          filename: a.original_name,
          original_name: a.original_name,
          path: a.file_path,
          file_path: a.file_path,
          contentType: a.mime_type,
          mime_type: a.mime_type
        }));

        // Personal PDF — matched by recipient name from the pdf folder
        const personalAttachments = [];
        if (campaign.pdf_folder_id) {
          const pdfFolderPath = path.join(__dirname, "uploads", "pdf-folders", campaign.pdf_folder_id);
          const personalPdfPath = findPdfForRecipient(pdfFolderPath, recipient.name);
          if (personalPdfPath && fs.existsSync(personalPdfPath)) {
            personalAttachments.push({
              filename: path.basename(personalPdfPath),
              original_name: path.basename(personalPdfPath),
              path: personalPdfPath,
              file_path: personalPdfPath,
              contentType: "application/pdf",
              mime_type: "application/pdf"
            });
          }
        }

        const allAttachments = [...personalAttachments, ...globalAttachments];


        let successNote = "";

        // Build the recipients list from to_emails (comma-joined) or fallback to email
        const toEmailsStr = recipient.to_emails || recipient.email || "";
        const toEmailsArr = toEmailsStr.split(",").map(s => s.trim()).filter(s => s && validEmail(s));
        const ccStr = recipient.cc || "";

        if (!toEmailsArr.length) {
          throw new Error("No valid email address for this recipient.");
        }

        const toDisplay = toEmailsArr.join(", ");

        if (useMicrosoft) {
          const graphResult = await sendMailViaMicrosoftGraph({
            toEmails: toEmailsArr,
            cc: ccStr,
            subject: recipient.rendered_subject,
            body: recipient.rendered_body,
            attachments: allAttachments
          });
          successNote = `Sent to ${toDisplay}${ccStr ? ` (CC: ${ccStr})` : ""} via Microsoft Outlook (${graphResult.messageId || "Accepted"})`;
        } else {
          const info = await transporter.sendMail(buildSmtpMailOptions({
            toEmails: toEmailsArr,
            cc: ccStr,
            subject: recipient.rendered_subject,
            body: recipient.rendered_body,
            attachments: allAttachments
          }));
          successNote = `Sent to ${toDisplay}${ccStr ? ` (CC: ${ccStr})` : ""} via SMTP. Message ID: ${info.messageId || "created"}`;
        }



        db.prepare(`
          UPDATE campaign_recipients
          SET status='sent', sent_at=CURRENT_TIMESTAMP, error=?
          WHERE id=?
        `).run(successNote, recipient.id);

        db.prepare("UPDATE campaigns SET sent=sent+1 WHERE id=?").run(campaign.id);
      } catch (e) {
        db.prepare(`
          UPDATE campaign_recipients
          SET status='failed', error=?
          WHERE id=?
        `).run(String(e.message || e).slice(0, 1000), recipient.id);

        db.prepare("UPDATE campaigns SET failed=failed+1 WHERE id=?").run(campaign.id);
      }

      // Use campaign-specific delay (set by user in compose page)
      const delayMs = Math.max(1000, Number(campaign.delay_seconds || 5) * 1000);
      await new Promise(resolve => setTimeout(resolve, delayMs));

    }

    db.prepare(`
      UPDATE campaigns
      SET status='completed', completed_at=CURRENT_TIMESTAMP
      WHERE id=?
    `).run(campaign.id);
  } catch (e) {
    db.prepare(`
      UPDATE campaigns
      SET status='failed', completed_at=CURRENT_TIMESTAMP
      WHERE id=?
    `).run(campaign.id);
  } finally {
    sendLock = false;
  }
}

app.post("/api/campaigns/:id/send", auth, async (req, res) => {
  if (sendLock) return res.status(409).json({ error: "Another campaign is currently sending. Please wait." });

  const campaign = db.prepare("SELECT * FROM campaigns WHERE id=?").get(req.params.id);
  if (!campaign) return res.status(404).json({ error: "Campaign not found." });

  if (campaign.status === "sending") return res.status(409).json({ error: "Campaign is already sending." });
  if (campaign.status === "completed") return res.status(400).json({ error: "Campaign already completed." });

  const useMicrosoft = campaign.sender_type === "microsoft" || (!campaign.sender_type && Boolean(getSetting("ms_tokens")));

  if (useMicrosoft) {
    try {
      await getValidMicrosoftAccessToken();
    } catch (e) {
      return res.status(400).json({ error: `Microsoft Outlook error: ${e.message}` });
    }
  } else {
    try {
      getTransporter();
    } catch (e) {
      return res.status(400).json({ error: e.message });
    }
  }

  res.json({ ok: true, message: "Campaign sending started.", campaignId: campaign.id });

  setImmediate(() => sendCampaign(campaign.id).catch(console.error));
});

app.get("/api/campaigns", auth, (req, res) => {
  res.json(db.prepare(`
    SELECT *
    FROM campaigns
    ORDER BY id DESC
  `).all());
});

app.get("/api/campaigns/:id", auth, (req, res) => {
  const campaign = db.prepare("SELECT * FROM campaigns WHERE id=?").get(req.params.id);
  if (!campaign) return res.status(404).json({ error: "Campaign not found." });

  const recipients = db.prepare(`
    SELECT id,name,email,to_emails,cc,status,error,sent_at,pdf_match_status
    FROM campaign_recipients
    WHERE campaign_id=?
    ORDER BY id
  `).all(req.params.id);

  const attachments = db.prepare(`
    SELECT a.id, a.original_name, a.mime_type, a.size
    FROM campaign_attachments ca
    JOIN attachments a ON a.id=ca.attachment_id
    WHERE ca.campaign_id=?
    ORDER BY ca.id
  `).all(req.params.id);

  res.json({ campaign, recipients, attachments });
});

// ── Pre-send validation endpoint ──────────────────────────────────────────────
app.post("/api/validate-recipients", auth, (req, res) => {
  const { selectedIds, subject, body } = req.body || {};

  let contacts = getContacts();
  if (Array.isArray(selectedIds) && selectedIds.length) {
    const ids = new Set(selectedIds.map(Number));
    contacts = contacts.filter(c => ids.has(c.id));
  }

  const rows = contacts.map(c => {
    let extras = {};
    try { extras = JSON.parse(c.extra_json || "{}"); } catch {}

    let toEmails = [];
    try { toEmails = JSON.parse(extras["__emails"] || "[]"); } catch {}
    if (!toEmails.length && c.email && validEmail(c.email)) toEmails = [c.email];

    const cc = String(extras["__cc"] || "");
    const invalid = extras["__invalid"] || null;

    const issues = [];
    if (invalid) issues.push(invalid);

    // Check for unknown variables in subject and body
    const unknownSubj = subject ? findUnknownVariables(subject, c) : [];
    const unknownBody = body ? findUnknownVariables(body, c) : [];
    const allUnknown = [...new Set([...unknownSubj, ...unknownBody])];
    if (allUnknown.length) {
      issues.push(`Unknown variable(s): ${allUnknown.join(", ")}`);
    }

    return {
      id: c.id,
      name: c.name,
      toEmails,
      cc,
      issues,
      valid: issues.length === 0 && toEmails.length > 0
    };
  });

  const valid = rows.filter(r => r.valid).length;
  const invalid = rows.filter(r => !r.valid).length;
  const missingEmail = rows.filter(r => r.toEmails.length === 0).length;

  res.json({ total: contacts.length, valid, invalid, missingEmail, rows });
});

// ── Download failed recipients as Excel ───────────────────────────────────────
app.get("/api/campaigns/:id/failed/export", auth, (req, res) => {
  const campaign = db.prepare("SELECT * FROM campaigns WHERE id=?").get(req.params.id);
  if (!campaign) return res.status(404).json({ error: "Campaign not found." });

  const failed = db.prepare(`
    SELECT cr.name, cr.email, cr.to_emails, cr.cc, cr.status, cr.error
    FROM campaign_recipients cr
    WHERE cr.campaign_id=? AND cr.status='failed'
    ORDER BY cr.id
  `).all(req.params.id);

  if (!failed.length) {
    return res.status(404).json({ error: "No failed recipients found for this campaign." });
  }

  const wsData = [
    ["Loan Number / Name", "Primary Email", "All To Emails", "CC", "Status", "Failure Reason"],
    ...failed.map(r => [
      r.name || "",
      r.email || "",
      r.to_emails || "",
      r.cc || "",
      r.status || "failed",
      r.error || "Unknown error"
    ])
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws["!cols"] = [20, 30, 40, 30, 12, 50].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws, "Failed");

  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  const safeName = (campaign.name || "campaign").replace(/[^a-z0-9]/gi, "_").slice(0, 40);
  const filename = `failed-${safeName}-${campaign.id}.xlsx`;

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(buffer);
});



// Download skipped (not sent) contacts as Excel
app.get("/api/campaigns/:id/skipped/export", auth, (req, res) => {
  const campaign = db.prepare("SELECT * FROM campaigns WHERE id=?").get(req.params.id);
  if (!campaign) return res.status(404).json({ error: "Campaign not found." });

  const skipped = db.prepare(`
    SELECT cr.name, cr.email,
           c.number, c.details
    FROM campaign_recipients cr
    LEFT JOIN contacts c ON c.id = cr.contact_id
    WHERE cr.campaign_id=? AND cr.pdf_match_status='skipped'
    ORDER BY cr.id
  `).all(req.params.id);

  if (!skipped.length) {
    return res.status(404).json({ error: "No skipped contacts found for this campaign." });
  }

  const wsData = [
    ["Name", "Email", "Phone", "Details"],
    ...skipped.map(r => [r.name || "", r.email || "", r.number || "", r.details || ""])
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "Not Sent");

  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  const filename = `not-sent-${campaign.name.replace(/[^a-z0-9]/gi, "_").slice(0, 40)}-${campaign.id}.xlsx`;

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(buffer);
});

app.get("/api/campaigns/:id/progress", auth, (req, res) => {
  const c = db.prepare(`
    SELECT id,name,status,total,sent,failed,skipped,started_at,completed_at
    FROM campaigns
    WHERE id=?
  `).get(req.params.id);

  if (!c) return res.status(404).json({ error: "Campaign not found." });
  res.json(c);
});

// Full campaign report: Sent + Undeliverable as two Excel sheets
app.get("/api/campaigns/:id/report", auth, (req, res) => {
  const campaign = db.prepare("SELECT * FROM campaigns WHERE id=?").get(req.params.id);
  if (!campaign) return res.status(404).json({ error: "Campaign not found." });

  const allRecipients = db.prepare(`
    SELECT cr.name AS loanNumber, cr.email, cr.status, cr.error, cr.sent_at,
           c.number, c.details, c.extra_json
    FROM campaign_recipients cr
    LEFT JOIN contacts c ON c.id = cr.contact_id
    WHERE cr.campaign_id=?
    ORDER BY cr.id
  `).all(req.params.id);

  const sentRows     = allRecipients.filter(r => r.status === "sent");
  const undelivRows  = allRecipients.filter(r => r.status === "failed" || r.status === "not_sent");

  // ── Sent Items sheet ─────────────────────────────────────────────────────────
  const sentHdr = ["Loan Number", "Email", "Phone", "Status", "Sent At"];
  const sentData = sentRows.map(r => [
    r.loanNumber || "—",
    r.email      || "—",
    r.number     || "—",
    "✓ Sent",
    r.sent_at ? new Date(r.sent_at).toLocaleString() : "—"
  ]);

  // ── Undeliverable sheet ───────────────────────────────────────────────────────
  const undelHdr = ["Loan Number", "Email", "Phone", "Status", "Failure Reason"];
  const undelData = undelivRows.map(r => [
    r.loanNumber || "—",
    r.email      || "—",
    r.number     || "—",
    r.status === "not_sent" ? "Not Sent" : "Failed",
    r.error      || "No matching PDF / manually denied"
  ]);

  const wb = XLSX.utils.book_new();

  const wsSent = XLSX.utils.aoa_to_sheet([sentHdr, ...sentData]);
  // Style the header row bold-ish by setting column widths
  wsSent["!cols"] = [18, 30, 14, 10, 22].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsSent, "Sent Items");

  const wsUndel = XLSX.utils.aoa_to_sheet([undelHdr, ...undelData]);
  wsUndel["!cols"] = [18, 30, 14, 14, 40].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsUndel, "Undeliverable");

  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  const safeName = (campaign.name || "campaign").replace(/[^a-z0-9]/gi, "_").slice(0, 40);
  const filename = `report-${safeName}-${campaign.id}.xlsx`;

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(buffer);
});


// Templates CRUD
app.get("/api/templates", auth, (req, res) => {
  res.json(db.prepare("SELECT id, name, subject, body, created_at FROM templates ORDER BY id DESC").all());
});

app.post("/api/templates", auth, (req, res) => {
  const { name, subject, body } = req.body || {};
  if (!name || !subject || !body) return res.status(400).json({ error: "Name, subject and body are required." });
  const id = db.prepare("INSERT INTO templates (name, subject, body) VALUES (?, ?, ?)").run(
    String(name).trim().slice(0, 100),
    String(subject).trim(),
    String(body).trim()
  ).lastInsertRowid;
  const t = db.prepare("SELECT * FROM templates WHERE id=?").get(id);
  res.json({ ok: true, template: t });
});

app.delete("/api/templates/:id", auth, (req, res) => {
  const t = db.prepare("SELECT id FROM templates WHERE id=?").get(req.params.id);
  if (!t) return res.status(404).json({ error: "Template not found." });
  db.prepare("DELETE FROM templates WHERE id=?").run(t.id);
  res.json({ ok: true });
});

// Express 5 compatible SPA fallback: do not use app.get("*").
app.use((req, res, next) => {
  if (req.method === "GET" && !req.path.startsWith("/api/")) {
    return res.sendFile(path.join(__dirname, "public", "index.html"));
  }
  next();
});

app.use((err, req, res, next) => {
  console.error(err);
  res.status(400).json({ error: err.message || "Request failed." });
});

app.listen(PORT, () => {
  console.log(`SmartMail Pro running at http://localhost:${PORT}`);
  console.log(`SMTP configured: ${Boolean(process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS)}`);

  // Auto-send scheduled campaigns
  setInterval(() => {
    try {
      const now = new Date().toISOString();
      const due = db.prepare(`
        SELECT id FROM campaigns 
        WHERE status = 'pending' 
          AND scheduled_at IS NOT NULL 
          AND scheduled_at <= ?
      `).all(now);
      for (const c of due) {
        // Clear scheduled_at so it doesn't re-trigger
        db.prepare("UPDATE campaigns SET scheduled_at = NULL WHERE id = ?").run(c.id);
        setImmediate(() => sendCampaign(c.id).catch(console.error));
        console.log(`[Scheduler] Auto-sending campaign ${c.id}`);
      }
    } catch(e) { console.error("Scheduler error:", e); }
  }, 60 * 1000);
});

