/**
 * Email to PDF — Core Logic (ES5 Compatible for maximum Outlook support)
 * Saim Studios
 *
 * This version uses standard 'var' and 'function' to ensure it runs even on
 * older Outlook installations using IE11 rendering.
 */

'use strict';

/* ── State Object ────────────────────────────────────────────────── */
var state = {
  subject:     '',
  from:        '',
  to:          '',
  cc:          '',
  date:        '',
  bodyHtml:    '',
  bodyText:    '',
  loaded:      false
};

/* ── Office Init ─────────────────────────────────────────────────── */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    // Small delay to ensure mailbox items are fully populated
    setTimeout(loadEmailData, 500);
  } else {
    showError('This add-in only works inside Outlook.');
  }
});

/* ── Load Email Data ─────────────────────────────────────────────── */
function loadEmailData() {
  try {
    var item = Office.context.mailbox.item;

    if (!item) {
      showError('Could not access the email item. Please select an email.');
      return;
    }

    // Subject
    state.subject = item.subject || '(No Subject)';

    // From
    if (item.from) {
      state.from = item.from.displayName
        ? item.from.displayName + ' <' + item.from.emailAddress + '>'
        : (item.from.emailAddress || '—');
    }

    // To
    if (item.to && item.to.length) {
      state.to = item.to.map(function(r) {
        return r.displayName ? r.displayName + ' <' + r.emailAddress + '>' : r.emailAddress;
      }).join('; ');
    }

    // CC
    if (item.cc && item.cc.length) {
      state.cc = item.cc.map(function(r) {
        return r.displayName ? r.displayName + ' <' + r.emailAddress + '>' : r.emailAddress;
      }).join('; ');
    }

    // Date
    var d = item.dateTimeCreated || item.dateTimeModified;
    if (d) {
      state.date = new Date(d).toLocaleString();
    }

    // Body (with 10-second timeout safety)
    var timeout = setTimeout(function() {
      if (!state.loaded) {
        showError('Timeout: Failed to retrieve email content from Outlook.');
      }
    }, 10000);

    item.body.getAsync(Office.CoercionType.Html, function (result) {
      clearTimeout(timeout);
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        state.bodyHtml = result.value || '';
        finaliseLoad();
      } else {
        // Fallback to text
        item.body.getAsync(Office.CoercionType.Text, function (r2) {
          if (r2.status === Office.AsyncResultStatus.Succeeded) {
            state.bodyText = r2.value || '';
            state.bodyHtml = '<pre style="white-space:pre-wrap;font-family:monospace;">' + escapeHtml(state.bodyText) + '</pre>';
          }
          finaliseLoad();
        });
      }
    });

  } catch (err) {
    showError('Error loading email data: ' + err.message);
  }
}

function finaliseLoad() {
  state.loaded = true;
  populateUI();
  show('main-content');
  hide('loading-state');
}

/* ── UI Helpers ──────────────────────────────────────────────────── */
function populateUI() {
  setText('meta-subject', state.subject);
  setText('meta-from',    state.from || '—');
  setText('meta-to',      state.to || '—');
  setText('meta-cc',      state.cc || '(none)');
  setText('meta-date',    state.date || '—');

  if (!state.cc) {
    var row = document.getElementById('cc-row');
    if (row) row.style.display = 'none';
  }

  var input = document.getElementById('pdf-filename');
  if (input) {
    input.value = sanitiseFilename(state.subject);
  }
}

/* ── PDF Generation ──────────────────────────────────────────────── */
function generatePDF() {
  if (!state.loaded) return;

  var opts = {
    metadata:    document.getElementById('opt-metadata').checked,
    recipients:  document.getElementById('opt-recipients').checked,
    date:        document.getElementById('opt-date').checked,
    html:        document.getElementById('opt-html').checked,
    pageNumbers: document.getElementById('opt-pagenumbers').checked,
    pageSize:    document.getElementById('opt-pagesize').value
  };

  var filenameInput = document.getElementById('pdf-filename');
  var filename = sanitiseFilename(filenameInput.value.trim() || state.subject) + '.pdf';

  setBtnDisabled(true);
  hide('toast-success');
  show('progress-wrap');
  setProgress(10, 'Preparing PDF engine...');

  // Using a slight delay to allow UI updates
  setTimeout(function() {
    executePDFCreation(opts, filename);
  }, 100);
}

function executePDFCreation(opts, filename) {
  try {
    setProgress(20, 'Building document...');
    var sanitised = sanitiseBody(state.bodyHtml, opts.html);
    var renderHtml = buildRenderHtml(sanitised, opts);
    var renderTarget = document.getElementById('render-target');
    renderTarget.innerHTML = renderHtml;

    setProgress(40, 'Rendering frames (this may take a moment)...');

    // Use html2canvas
    html2canvas(renderTarget, {
      scale: 2,
      useCORS: false,
      backgroundColor: '#ffffff'
    }).then(function(canvas) {
      setProgress(70, 'Composing PDF pages...');
      
      var format = opts.pageSize || 'a4';
      var pdf = new jspdf.jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: format,
        compress: true
      });

      var pageW = pdf.internal.pageSize.getWidth();
      var pageH = pdf.internal.pageSize.getHeight();
      
      var margin = 15;
      var contentW = pageW - (margin * 2);
      var scale = contentW / (canvas.width / 2);
      var contentH_mm = (canvas.height / 2) * scale;
      var pageContentH = pageH - (margin * 2) - (opts.pageNumbers ? 10 : 0);

      var sourceY_px = 0;
      var pageNum = 1;
      var totalPages = Math.ceil(contentH_mm / pageContentH);

      while (sourceY_px < canvas.height) {
        if (pageNum > 1) pdf.addPage(format, 'portrait');

        var pxThisPage = Math.min(Math.round(pageContentH / scale * 2), canvas.height - sourceY_px);
        var slice = document.createElement('canvas');
        slice.width = canvas.width;
        slice.height = pxThisPage;
        var ctx = slice.getContext('2d');
        ctx.drawImage(canvas, 0, sourceY_px, canvas.width, pxThisPage, 0, 0, canvas.width, pxThisPage);

        var imgData = slice.toDataURL('image/png');
        pdf.addImage(imgData, 'PNG', margin, margin, contentW, (pxThisPage / 2) * scale, '', 'FAST');

        if (opts.pageNumbers) {
          pdf.setFontSize(8);
          pdf.setTextColor(150, 150, 150);
          pdf.text('Page ' + pageNum + ' of ' + totalPages, pageW / 2, pageH - 10, { align: 'center' });
        }

        sourceY_px += pxThisPage;
        pageNum++;
      }

      pdf.save(filename);
      
      setProgress(100, 'Success!');
      setTimeout(function() {
        hide('progress-wrap');
        show('toast-success');
        setBtnDisabled(false);
      }, 500);
      renderTarget.innerHTML = '';

    }).catch(function(err) {
      showError('PDF Rendering failed: ' + err.message);
      setBtnDisabled(false);
    });

  } catch (err) {
    showError('Error creating PDF: ' + err.message);
    setBtnDisabled(false);
  }
}

/* ── Rendering Support ───────────────────────────────────────────── */
function buildRenderHtml(body, opts) {
  var html = '<div style="font-family:Arial,sans-serif;font-size:13px;color:#111;background:#fff;padding:30px;">';
  
  if (opts.metadata) {
    html += '<div style="margin-bottom:20px;border:1px solid #ddd;padding:15px;background:#f9f9f9;border-radius:4px;">';
    html += '<h3 style="margin:0 0 10px;font-size:14px;color:#333;">📧 Email Information</h3>';
    if (state.subject) html += '<div><b>Subject:</b> ' + escapeHtml(state.subject) + '</div>';
    if (opts.recipients) {
      if (state.from) html += '<div><b>From:</b> ' + escapeHtml(state.from) + '</div>';
      if (state.to)   html += '<div><b>To:</b> ' + escapeHtml(state.to) + '</div>';
    }
    if (opts.date && state.date) html += '<div><b>Date:</b> ' + escapeHtml(state.date) + '</div>';
    html += '</div><hr style="border:none;border-top:1px solid #eee;margin-bottom:20px;">';
  }

  html += '<div class="body-content">' + body + '</div>';
  html += '</div>';
  return html;
}

function sanitiseBody(html, preserveStyling) {
  if (!html) return '(No content)';
  var div = document.createElement('div');
  div.innerHTML = html;

  var scripts = div.getElementsByTagName('script');
  for (var i = scripts.length - 1; i >= 0; i--) {
    scripts[i].parentNode.removeChild(scripts[i]);
  }

  var imgs = div.getElementsByTagName('img');
  for (var j = 0; j < imgs.length; j++) {
    var src = imgs[j].getAttribute('src') || '';
    if (src.indexOf('data:') !== 0) {
      imgs[j].style.display = 'none'; // Privacy: hide remote images
    }
  }

  if (!preserveStyling) {
    var all = div.getElementsByTagName('*');
    for (var k = 0; k < all.length; k++) {
      all[k].removeAttribute('style');
    }
  }

  return div.innerHTML;
}

/* ── General Helpers ─────────────────────────────────────────────── */
function sanitiseFilename(name) {
  return (name || 'email').replace(/[/\\:*?"<>|]/g, '_').slice(0, 50);
}

function escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function show(id) {
  var el = document.getElementById(id);
  if (el) el.classList.remove('hidden');
}

function hide(id) {
  var el = document.getElementById(id);
  if (el) el.classList.add('hidden');
}

function setText(id, val) {
  var el = document.getElementById(id);
  if (el) el.textContent = val;
}

function setProgress(pct, text) {
  var fill = document.getElementById('progress-fill');
  var lbl = document.getElementById('progress-text');
  if (fill) fill.style.width = pct + '%';
  if (lbl) lbl.textContent = text;
}

function setBtnDisabled(d) {
  var btn = document.getElementById('btn-generate');
  if (btn) btn.disabled = d;
}

function showError(msg) {
  hide('loading-state');
  show('error-state');
  setText('error-message', msg);
}
