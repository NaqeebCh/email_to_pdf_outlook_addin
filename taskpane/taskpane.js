/**
 * Email to PDF — Core Logic
 * Saim Studios
 *
 * Architecture:
 *  1. Office.onReady()  → initialise add-in
 *  2. loadEmailData()   → read all email metadata + body via Office.js
 *  3. generatePDF()     → render HTML to canvas → compose A4 PDF → download
 *
 * Privacy: 100% client-side. No data leaves the browser.
 * Permissions required: ReadItem only.
 */

'use strict';

/* ═══════════════════════════════════════════════════════════════════
   STATE
   ═══════════════════════════════════════════════════════════════════ */
const state = {
  subject:     '',
  from:        '',
  to:          '',
  cc:          '',
  date:        '',
  bodyHtml:    '',
  bodyText:    '',
  loaded:      false,
};

/* ═══════════════════════════════════════════════════════════════════
   OFFICE INIT
   ═══════════════════════════════════════════════════════════════════ */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    loadEmailData();
  } else {
    showError('This add-in only works inside Outlook.');
  }
});

/* ═══════════════════════════════════════════════════════════════════
   LOAD EMAIL DATA
   ═══════════════════════════════════════════════════════════════════ */
function loadEmailData() {
  const item = Office.context.mailbox.item;

  if (!item) {
    showError('Could not access the email item. Please open an email and try again.');
    return;
  }

  // ── Subject ──────────────────────────────────────────────────────
  state.subject = item.subject || '(No Subject)';

  // ── From ─────────────────────────────────────────────────────────
  if (item.from) {
    const f = item.from;
    state.from = f.displayName
      ? `${f.displayName} <${f.emailAddress}>`
      : (f.emailAddress || '—');
  }

  // ── To ───────────────────────────────────────────────────────────
  if (item.to && Array.isArray(item.to)) {
    state.to = item.to
      .map(r => r.displayName ? `${r.displayName} <${r.emailAddress}>` : r.emailAddress)
      .join('; ') || '—';
  }

  // ── CC ───────────────────────────────────────────────────────────
  if (item.cc && Array.isArray(item.cc)) {
    state.cc = item.cc
      .map(r => r.displayName ? `${r.displayName} <${r.emailAddress}>` : r.emailAddress)
      .join('; ') || '';
  }

  // ── Date ─────────────────────────────────────────────────────────
  const d = item.dateTimeCreated || item.dateTimeModified;
  if (d) {
    state.date = new Date(d).toLocaleString('en-US', {
      weekday: 'long', year: 'numeric', month: 'long',
      day: 'numeric', hour: '2-digit', minute: '2-digit', timeZoneName: 'short'
    });
  }

  // ── Body (HTML preferred, text fallback) ─────────────────────────
  item.body.getAsync(Office.CoercionType.Html, { asyncContext: 'html' }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      state.bodyHtml = result.value || '';
      finaliseLoad();
    } else {
      // Fallback to plain text
      item.body.getAsync(Office.CoercionType.Text, {}, function (r2) {
        if (r2.status === Office.AsyncResultStatus.Succeeded) {
          state.bodyText = r2.value || '';
          state.bodyHtml = `<pre style="white-space:pre-wrap;font-family:monospace;">${escapeHtml(state.bodyText)}</pre>`;
        }
        finaliseLoad();
      });
    }
  });
}

/* ── Finalise after async body load ─────────────────────────────── */
function finaliseLoad() {
  state.loaded = true;
  populateUI();
  show('main-content');
  hide('loading-state');
}

/* ═══════════════════════════════════════════════════════════════════
   POPULATE UI
   ═══════════════════════════════════════════════════════════════════ */
function populateUI() {
  setText('meta-subject', state.subject);
  setText('meta-from',    state.from    || '—');
  setText('meta-to',      state.to      || '—');
  setText('meta-cc',      state.cc      || '(none)');
  setText('meta-date',    state.date    || '—');

  // Hide CC row if empty
  if (!state.cc) {
    const row = document.getElementById('cc-row');
    if (row) row.style.display = 'none';
  }

  // Set sanitised filename
  const input = document.getElementById('pdf-filename');
  if (input) {
    input.value = sanitiseFilename(state.subject);
  }
}

/* ═══════════════════════════════════════════════════════════════════
   GENERATE PDF
   (called by the button onclick in HTML)
   ═══════════════════════════════════════════════════════════════════ */
async function generatePDF() {  // eslint-disable-line no-unused-vars
  if (!state.loaded) {
    showError('Email data is still loading. Please wait a moment.');
    return;
  }

  // ── Read options ─────────────────────────────────────────────────
  const opts = {
    metadata:    document.getElementById('opt-metadata').checked,
    recipients:  document.getElementById('opt-recipients').checked,
    date:        document.getElementById('opt-date').checked,
    html:        document.getElementById('opt-html').checked,
    pageNumbers: document.getElementById('opt-pagenumbers').checked,
    pageSize:    document.getElementById('opt-pagesize').value,
  };

  const filenameInput = document.getElementById('pdf-filename');
  const filename = sanitiseFilename(filenameInput.value.trim() || state.subject) + '.pdf';

  // ── UI: start progress ───────────────────────────────────────────
  setBtnDisabled(true);
  hide('toast-success');
  show('progress-wrap');
  setProgress(5, 'Preparing email content…');

  try {
    // ── 1. Build render HTML ──────────────────────────────────────
    setProgress(15, 'Sanitising email body…');
    const sanitised = sanitiseBody(state.bodyHtml, opts.html);
    const renderHtml = buildRenderHtml(sanitised, opts);

    // ── 2. Inject into off-screen render target ───────────────────
    setProgress(25, 'Rendering layout…');
    const renderTarget = document.getElementById('render-target');
    renderTarget.innerHTML = renderHtml;

    // Give the browser one frame to paint the injected DOM
    await sleep(80);
    setProgress(40, 'Capturing page…');

    // ── 3. Render to canvas via html2canvas ───────────────────────
    const canvas = await html2canvas(renderTarget, {
      scale: 2,               // 2x for high-DPI / print quality
      useCORS: false,         // no external images — privacy
      allowTaint: false,
      backgroundColor: '#ffffff',
      logging: false,
      imageTimeout: 0,
      removeContainer: false,
    });

    setProgress(70, 'Composing PDF…');

    // ── 4. Compose PDF with jsPDF ─────────────────────────────────
    // Choose format
    const formatMap = { a4: 'a4', letter: 'letter', legal: 'legal' };
    const format = formatMap[opts.pageSize] || 'a4';

    // A4: 210×297mm  |  Letter: 216×279mm  |  Legal: 216×356mm
    const pageDimensions = {
      a4:     { w: 210, h: 297 },
      letter: { w: 216, h: 279 },
      legal:  { w: 216, h: 356 },
    };
    const pageDim = pageDimensions[format];

    const pdf = new jspdf.jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: format,
      compress: true,
    });

    // Margins in mm
    const marginL = 15;
    const marginR = 15;
    const marginT = 15;
    const marginB = 15;

    const canvasWidth    = canvas.width;
    const canvasHeight   = canvas.height;
    const pdfContentW    = pageDim.w - marginL - marginR;  // usable width in mm

    // Scale factor: canvas px → mm
    const scale          = pdfContentW / (canvasWidth / 2); // /2 because we used scale:2
    const totalContentH  = (canvasHeight / 2) * scale;       // total height in mm

    // Usable height per page (leave footer space if page numbers on)
    const footerH        = opts.pageNumbers ? 8 : 0;
    const pageContentH   = pageDim.h - marginT - marginB - footerH;

    let sourceY          = 0;  // in CANVAS pixels (full scale)
    let pageNumber       = 1;
    const totalPages     = Math.ceil(totalContentH / pageContentH);

    while (sourceY < canvasHeight) {
      if (pageNumber > 1) pdf.addPage(format, 'portrait');

      // How many mm of content this page can hold → how many px that is
      const pxPerPage = Math.round(pageContentH / scale * 2); // back to full-scale px

      // Clamp to remaining canvas
      const pxThisPage = Math.min(pxPerPage, canvasHeight - sourceY);

      // Create a slice canvas
      const sliceCanvas = document.createElement('canvas');
      sliceCanvas.width  = canvasWidth;
      sliceCanvas.height = pxThisPage;
      const ctx = sliceCanvas.getContext('2d');
      ctx.drawImage(canvas, 0, sourceY, canvasWidth, pxThisPage, 0, 0, canvasWidth, pxThisPage);

      const imgData = sliceCanvas.toDataURL('image/png');
      const sliceHmm = (pxThisPage / 2) * scale;

      pdf.addImage(imgData, 'PNG', marginL, marginT, pdfContentW, sliceHmm, '', 'FAST');

      // ── Page number footer ──────────────────────────────────────
      if (opts.pageNumbers) {
        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(8);
        pdf.setTextColor(150, 150, 150);
        const footerY = pageDim.h - marginB - 2;
        pdf.text(`Page ${pageNumber} of ${totalPages}`, pageDim.w / 2, footerY, { align: 'center' });
        // Thin separator line
        pdf.setDrawColor(220, 220, 220);
        pdf.setLineWidth(0.2);
        pdf.line(marginL, footerY - 3, pageDim.w - marginR, footerY - 3);
      }

      sourceY    += pxThisPage;
      pageNumber += 1;
    }

    setProgress(90, 'Saving…');
    await sleep(100);

    // ── 5. Clean up render target ─────────────────────────────────
    renderTarget.innerHTML = '';

    // ── 6. Download ───────────────────────────────────────────────
    pdf.save(filename);

    setProgress(100, 'Done!');
    await sleep(400);

    // ── 7. Success feedback ───────────────────────────────────────
    hide('progress-wrap');
    show('toast-success');
    setBtnDisabled(false);

    // Auto-hide toast after 5s
    setTimeout(() => hide('toast-success'), 5000);

  } catch (err) {
    hide('progress-wrap');
    setBtnDisabled(false);
    console.error('[Email to PDF] Error:', err);
    showError('Failed to generate PDF: ' + (err.message || String(err)));
  }
}

/* ═══════════════════════════════════════════════════════════════════
   BUILD RENDER HTML
   Constructs the white A4-width document that will be screenshot-d
   ═══════════════════════════════════════════════════════════════════ */
function buildRenderHtml(body, opts) {
  const parts = [];

  parts.push(`
    <div style="
      font-family: Arial, Helvetica, sans-serif;
      font-size: 13px;
      color: #111111;
      background: #ffffff;
      padding: 32px 36px 40px;
      -webkit-print-color-adjust: exact;
    ">
  `);

  // ── Metadata header table ───────────────────────────────────────
  if (opts.metadata) {
    parts.push(`
      <table style="
        width: 100%; border-collapse: collapse;
        margin-bottom: 24px;
        border: 1px solid #dee2e6;
        border-radius: 6px;
        overflow: hidden;
        font-size: 12px;
      ">
        <thead>
          <tr style="background: #1e3a5f;">
            <td colspan="2" style="padding: 10px 14px; color: #ffffff; font-weight: 700; font-size: 13px; letter-spacing: 0.02em;">
              📧 Email Details
            </td>
          </tr>
        </thead>
        <tbody>
    `);

    const rows = [];

    if (state.subject) {
      rows.push(['Subject', escapeHtml(state.subject)]);
    }
    if (opts.recipients) {
      if (state.from)      rows.push(['From',  escapeHtml(state.from)]);
      if (state.to)        rows.push(['To',    escapeHtml(state.to)]);
      if (state.cc)        rows.push(['CC',    escapeHtml(state.cc)]);
    }
    if (opts.date && state.date) {
      rows.push(['Date', escapeHtml(state.date)]);
    }

    rows.forEach(([label, value], i) => {
      const bg = i % 2 === 0 ? '#f8fafc' : '#ffffff';
      parts.push(`
        <tr style="background: ${bg};">
          <td style="
            padding: 7px 14px;
            font-weight: 600;
            color: #374151;
            width: 72px;
            border-bottom: 1px solid #e9ecef;
            vertical-align: top;
            white-space: nowrap;
          ">${label}</td>
          <td style="
            padding: 7px 14px;
            color: #4b5563;
            border-bottom: 1px solid #e9ecef;
            word-break: break-word;
          ">${value}</td>
        </tr>
      `);
    });

    parts.push('</tbody></table>');

    // Divider
    parts.push(`<hr style="border: none; border-top: 2px solid #e2e8f0; margin: 0 0 24px;">`);
  }

  // ── Email body ─────────────────────────────────────────────────
  parts.push(`<div class="email-body" style="
    line-height: 1.6;
    word-break: break-word;
    overflow-wrap: break-word;
  ">`);
  parts.push(body);
  parts.push('</div>');

  parts.push('</div>');

  return parts.join('\n');
}

/* ═══════════════════════════════════════════════════════════════════
   SANITISE HTML BODY
   Strips dangerous tags & external resources for security + privacy
   ═══════════════════════════════════════════════════════════════════ */
function sanitiseBody(html, preserveStyling) {
  if (!html || html.trim() === '') {
    return '<p style="color:#888;">(This email has no body content.)</p>';
  }

  const parser  = new DOMParser();
  const doc     = parser.parseFromString(html, 'text/html');

  // ── Remove dangerous elements ───────────────────────────────────
  const REMOVE_TAGS = ['script', 'style[data-remove]', 'iframe', 'frame',
                       'object', 'embed', 'applet', 'form', 'input',
                       'button', 'link[rel="stylesheet"]'];
  REMOVE_TAGS.forEach(selector => {
    doc.querySelectorAll(selector).forEach(el => el.remove());
  });

  // Remove ALL <script> tags absolutely
  doc.querySelectorAll('script').forEach(el => el.remove());

  // ── Remove external image sources (privacy) ─────────────────────
  doc.querySelectorAll('img').forEach(img => {
    const src = img.getAttribute('src') || '';
    // Allow data URIs (inline images) only
    if (!src.startsWith('data:')) {
      // Replace with a small placeholder text
      const span = doc.createElement('span');
      span.style.cssText = 'display:inline-block; padding:2px 6px; background:#f1f5f9; border:1px solid #e2e8f0; border-radius:3px; font-size:11px; color:#94a3b8;';
      span.textContent = '[image]';
      img.parentNode && img.parentNode.replaceChild(span, img);
    }
  });

  // ── Strip event handlers ────────────────────────────────────────
  const EVENT_ATTRS = ['onclick','onload','onerror','onmouseover','onmouseout',
                       'onfocus','onblur','onchange','onsubmit','onkeydown',
                       'onkeyup','onkeypress'];
  doc.querySelectorAll('*').forEach(el => {
    EVENT_ATTRS.forEach(attr => el.removeAttribute(attr));
    // Strip javascript: hrefs
    const href = el.getAttribute('href') || '';
    if (/^javascript:/i.test(href)) el.removeAttribute('href');
  });

  // ── Remove external link targets (open in new tabs safely) ─────
  doc.querySelectorAll('a[href]').forEach(a => {
    a.setAttribute('target', '_blank');
    a.setAttribute('rel', 'noopener noreferrer');
  });

  if (!preserveStyling) {
    // Strip inline styles if user opted out
    doc.querySelectorAll('[style]').forEach(el => el.removeAttribute('style'));
  }

  // Extract just the body content
  const body = doc.body;
  return body ? body.innerHTML : html;
}

/* ═══════════════════════════════════════════════════════════════════
   HELPERS
   ═══════════════════════════════════════════════════════════════════ */

/**
 * Sanitise a string for use as a filename.
 * Removes illegal characters, collapses whitespace, trims to 200 chars.
 */
function sanitiseFilename(name) {
  if (!name || name.trim() === '') {
    return 'email_' + Date.now();
  }
  return name
    .trim()
    .replace(/[/\\:*?"<>|]/g, '_')  // illegal filename chars
    .replace(/\s+/g, ' ')           // collapse whitespace
    .replace(/\.+$/, '')            // no trailing dots
    .slice(0, 200)                  // max 200 chars
    || 'email_' + Date.now();
}

/** Escape HTML special chars */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Sleep for ms milliseconds */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/* ── DOM Helpers ──────────────────────────────────────────────────── */
function show(id) {
  const el = document.getElementById(id);
  if (el) el.classList.remove('hidden');
}

function hide(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add('hidden');
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function setProgress(pct, text) {
  const fill = document.getElementById('progress-fill');
  const label = document.getElementById('progress-text');
  if (fill) fill.style.width = pct + '%';
  if (label) label.textContent = text;
}

function setBtnDisabled(disabled) {
  const btn = document.getElementById('btn-generate');
  const label = document.getElementById('btn-label');
  if (btn) btn.disabled = disabled;
  if (label) label.textContent = disabled ? 'Generating PDF…' : 'Generate & Download PDF';
}

/* ── Error display ────────────────────────────────────────────────── */
function showError(msg) {
  hide('loading-state');
  hide('main-content');
  const errMsg = document.getElementById('error-message');
  if (errMsg) errMsg.textContent = msg;
  show('error-state');
}
