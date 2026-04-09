# 📧 Email to PDF — Outlook Add-in

Convert any open email to a professionally formatted PDF — **100% private, processed entirely on your device.**

---

## ✨ Features

| Feature | Detail |
|---|---|
| **PDF Filename** | Auto-set from email subject (editable before saving) |
| **Metadata Header** | Subject · From · To · CC · Date in a styled table |
| **Full Email Body** | HTML body with inline styles preserved |
| **Multi-page PDF** | Automatically paginated for long emails |
| **Page Numbers** | "Page X of Y" footer on every page |
| **Page Sizes** | A4 · Letter · Legal |
| **Toggle Options** | Per-save control over what's included |
| **Privacy** | Zero network calls — all processing is local |
| **Permissions** | `ReadItem` only (minimum required by Office.js) |
| **Platforms** | Outlook Desktop (Win/Mac) + Outlook on the Web |

---

## 📁 File Structure

```
EMail to PDF/
├── manifest.xml                ← Sideload this into Outlook
├── taskpane/
│   ├── taskpane.html           ← Task pane UI
│   ├── taskpane.css            ← Premium dark theme
│   ├── taskpane.js             ← Core logic
│   └── commands.html           ← Required placeholder
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
└── libs/
    ├── jspdf.umd.min.js        ← Bundled locally (no CDN)
    └── html2canvas.min.js      ← Bundled locally (no CDN)
```

---

## 🚀 Installation

### Step 1 — Host the Files

Outlook add-ins require all files to be served over **HTTPS**.

**Option A — GitHub Pages (Free, Recommended):**
1. Push this folder to a GitHub repository
2. Enable GitHub Pages (Settings → Pages → Deploy from branch)
3. Your base URL will be: `https://github.com/NaqeebCh/email_to_pdf_outlook_addin.git`

**Option B — Any HTTPS Web Server:**
- IIS, Apache, Nginx, or any static host (Netlify, Vercel, etc.)

**Option C — Local Dev Testing (localhost):**
- Use [office-addin-dev-certs](https://www.npmjs.com/package/office-addin-dev-certs) to generate a trusted self-signed cert
- Run `npx office-addin-dev-certs install` then serve on `https://localhost:3000`

---

### Step 2 — Update manifest.xml

Once you have your hosting URL, open `manifest.xml` and replace **all** occurrences of:
```
https://localhost:3000
```
with your actual HTTPS URL, e.g.:
```
https://your-username.github.io/email-to-pdf
```

There are **6 URL references** in manifest.xml to update (search & replace all).

---

### Step 3 — Sideload into Outlook

#### Outlook Desktop (Windows):
1. Open Outlook → **File → Manage Add-ins** (opens browser)
2. Click **"Add a custom add-in"** → **"Add from file..."**
3. Browse to `manifest.xml` → Click **Open**
4. Confirm the security warning → Add-in installs
5. Open any email → Look for **"Save as PDF"** button in the ribbon

#### Outlook on the Web:
1. Go to [outlook.office.com](https://outlook.office.com) → Gear ⚙ → **View all Outlook settings**
2. **Mail → Customize actions → Manage add-ins**
3. Click **"+"** → **"Add from file..."** → Upload `manifest.xml`
4. Open any email → Click the **"Save as PDF"** button in the toolbar

---

## 🔒 Privacy Architecture

- **No server required** — PDF generation runs entirely in your browser
- **No CDN calls at runtime** — jsPDF and html2canvas are bundled locally
- **External images stripped** — tracking pixels and remote images are replaced with `[image]` placeholders
- **Script tags removed** — all `<script>` content is sanitized from the email body before rendering
- **Event handlers stripped** — `onclick`, `onload`, etc. are removed from HTML
- **`ReadItem` only** — the add-in cannot send, modify, or delete emails

---

## 🎛️ PDF Options

When you click **Save as PDF**, the task pane gives you full control:

| Option | Default | Effect |
|---|---|---|
| Include metadata header | ✅ On | Adds a styled table with Subject/From/To/Date |
| Include sender & recipients | ✅ On | Shows From/To/CC in the header |
| Include date & time | ✅ On | Shows the sent timestamp |
| Preserve HTML styling | ✅ On | Keeps email fonts, colors, layout |
| Add page numbers | ✅ On | Footer: "Page X of Y" |
| Page size | A4 | Also: Letter, Legal |
| Filename | (email subject) | Editable — illegal chars auto-removed |

---

## 🛠️ Updating the Libraries

The bundled libraries are:
- **jsPDF 2.5.1** → https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js
- **html2canvas 1.4.1** → https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js

To update, download newer versions and replace the files in `libs/`.

---

## 📋 Requirements

- Microsoft Outlook (Desktop 2016+ or Microsoft 365, or Outlook Web)
- A modern browser (Chromium-based) inside Outlook
- Mailbox API requirement set: **1.3** or higher

---

*Naqeeb Ch — Built for privacy, speed, and simplicity.*
