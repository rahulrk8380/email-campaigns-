# SmartMail Pro 3.1 — Professional Email Automation

A full-stack Node.js + Express application for turning Excel contact lists into personalized one-to-one email campaigns.

## Features
- Professional responsive dashboard UI
- Excel / XLS / CSV upload
- Email validation and duplicate removal
- Select individual recipients or all contacts
- Personalized variables: `{{name}}`, `{{email}}`, `{{number}}`, `{{details}}`, and extra Excel columns
- Live email preview
- Test email before campaign send
- Multiple file attachments
- Photos: JPG, JPEG, PNG, GIF, WEBP
- Documents: PDF, DOC, DOCX, TXT
- Office files: XLS, XLSX, PPT, PPTX, CSV
- ZIP attachments
- Up to 8 attachments per upload, 15 MB per file, 24 MB total recommended
- Recipient-level Sent / Failed results
- Campaign history and retry support
- Gmail SMTP with IPv4 preference to avoid common IPv6 timeout issues
- SQLite persistence
- Express 5 compatible SPA fallback

## Run on Windows PowerShell

1. Extract the ZIP and open the folder in VS Code.
2. Open Terminal in that folder.
3. Install dependencies:

```powershell
npm install
```

4. Create your `.env` file:

```powershell
Copy-Item .env.example .env
```

5. Edit `.env` and enter your SMTP settings. For Gmail with an App Password:

```text
PORT=3000
SESSION_SECRET=replace-with-a-long-random-secret
ADMIN_EMAIL=admin@example.com
ADMIN_PASSWORD=change-me
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=yourgmail@gmail.com
SMTP_PASS=your-16-character-google-app-password
MAIL_FROM="SmartMail <yourgmail@gmail.com>"
MAX_RECIPIENTS_PER_CAMPAIGN=2000
```

6. Start:

```powershell
npm start
```

7. Open `http://localhost:3000`.

## Important

A successful SMTP send means the receiving mail server accepted the message; it does not guarantee Inbox placement. Gmail and other providers decide whether a message goes to Inbox, Spam, Promotions, etc. Domain authentication (SPF, DKIM, DMARC), sender reputation, consent, content quality and sending volume all affect deliverability.

Never paste your real Gmail App Password into chat or commit it to GitHub.
