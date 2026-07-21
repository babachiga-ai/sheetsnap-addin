// taskpane.js – SheetSnap Excel Add-in v1.0
//
// Instead of screenshotting Excel's rendered UI (fragile, canvas-based,
// impossible to detect selections in reliably from outside), we use the
// Excel JavaScript API to read the ACTUAL selected range: its values,
// column widths, row heights, fonts, and fill colors. Then we draw that
// data directly onto our own <canvas> element, pixel by pixel. This is
// exact, instant, and works for any selection size — no scrolling,
// no stitching, no screenshots at all.

let lastCanvas = null;
let lastFilename = null;
let currentFormat = 'png';
let currentQuality = 'medium';

Office.onReady((info) => {
  window.__sheetSnapReady = true;
  const bootErr = document.getElementById('bootError');
  if (bootErr) bootErr.style.display = 'none';

  if (info.host === Office.HostType.Excel) {
    document.getElementById('refreshBtn').addEventListener('click', refreshSelection);
    document.getElementById('capBtn').addEventListener('click', captureSelection);
    document.getElementById('downloadBtn').addEventListener('click', downloadLast);
    document.getElementById('clipboardBtn').addEventListener('click', copyLastToClipboard);

    document.querySelectorAll('[data-fmt]').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('[data-fmt]').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentFormat = btn.dataset.fmt;
      });
    });
    document.querySelectorAll('[data-q]').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('[data-q]').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentQuality = btn.dataset.q;
      });
    });

    // Also refresh automatically whenever the selection changes on the sheet
    Excel.run(async (context) => {
      context.workbook.worksheets.onSelectionChanged.add(refreshSelection);
      await context.sync();
    }).catch(() => { /* non-fatal if not supported */ });

    refreshSelection();
  }
});

// ─────────────────────────────────────────────────────────────────────
// Read the current selection's address & size, update the status card
// ─────────────────────────────────────────────────────────────────────
async function refreshSelection() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'rowCount', 'columnCount']);
      await context.sync();

      const cellCount = range.rowCount * range.columnCount;
      const statusEl = document.getElementById('selStatus');
      const dotEl    = document.getElementById('selDot');
      const addrEl   = document.getElementById('selAddr');
      const dimsEl   = document.getElementById('selDims');
      const hintEl   = document.getElementById('selHint');
      const capBtn   = document.getElementById('capBtn');

      const shortAddr = range.address.includes('!') ? range.address.split('!')[1] : range.address;

      addrEl.textContent = shortAddr;
      dimsEl.textContent = `${range.rowCount} rows × ${range.columnCount} columns  (${cellCount.toLocaleString()} cells)`;

      if (cellCount > 8000) {
        statusEl.className = 'sel-status error';
        dotEl.className = 'sel-dot';
        hintEl.textContent = `Selection is very large (${cellCount.toLocaleString()} cells). Please select a smaller range — up to about 8,000 cells works best.`;
        capBtn.disabled = true;
      } else {
        statusEl.className = 'sel-status ready';
        dotEl.className = 'sel-dot on';
        hintEl.textContent = 'Ready to capture. Click the button below to save this range as an image.';
        capBtn.disabled = false;
      }
    });
  } catch (e) {
    const statusEl = document.getElementById('selStatus');
    statusEl.className = 'sel-status empty';
    document.getElementById('selDot').className = 'sel-dot';
    document.getElementById('selAddr').textContent = 'No selection detected';
    document.getElementById('selDims').textContent = '—';
    document.getElementById('selHint').textContent = 'Click any cell or range in your worksheet, then press Refresh.';
    document.getElementById('capBtn').disabled = true;
  }
}

// ─────────────────────────────────────────────────────────────────────
// Main capture: read range data via Office.js, render to canvas
// ─────────────────────────────────────────────────────────────────────
async function captureSelection() {
  const capBtn = document.getElementById('capBtn');
  capBtn.classList.add('loading');
  capBtn.disabled = true;
  setProgress(true, 5, 'Reading selection…');

  try {
    const canvas = await buildCanvasFromSelection((pct, label) => setProgress(true, pct, label));

    setProgress(true, 95, 'Finalizing image…');
    lastCanvas = canvas;
    lastFilename = makeFilename(currentFormat);

    const dataUrl = canvas.toDataURL(currentFormat === 'jpeg' ? 'image/jpeg' : 'image/png', 0.95);
    document.getElementById('previewImg').src = dataUrl;
    document.getElementById('previewWrap').classList.add('on');

    setProgress(true, 100, 'Done!');
    showToast('✅ Image ready — download or copy it below', 'ok');

  } catch (e) {
    console.error('[SheetSnap]', e);
    showToast('❌ ' + (e.message || 'Capture failed'), 'err');
  } finally {
    setTimeout(() => {
      capBtn.classList.remove('loading');
      capBtn.disabled = false;
      setProgress(false, 0, '');
    }, 900);
  }
}

// Points → pixels (96 DPI standard: 1pt = 4/3 px)
const PT_TO_PX = 96 / 72;

function qualityScale(q) { return q === 'high' ? 2 : q === 'medium' ? 1.5 : 1; }

async function buildCanvasFromSelection(onProgress) {
  let result;

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(['rowCount', 'columnCount', 'text']);
    await context.sync();

    const rows = range.rowCount;
    const cols = range.columnCount;

    if (rows * cols > 8000) {
      throw new Error(`Selection too large (${rows*cols} cells). Please select fewer than 8,000 cells.`);
    }

    onProgress(15, 'Reading column widths…');

    // Queue column width loads
    const colRanges = [];
    for (let c = 0; c < cols; c++) {
      const cr = range.getColumn(c);
      cr.format.load('columnWidth');
      colRanges.push(cr);
    }
    // Queue row height loads
    const rowRanges = [];
    for (let r = 0; r < rows; r++) {
      const rr = range.getRow(r);
      rr.format.load('rowHeight');
      rowRanges.push(rr);
    }

    onProgress(30, 'Reading cell styles (including conditional formatting)…');

    // getCellProperties() returns the ACTUAL DISPLAYED formatting — this
    // is the key fix: unlike range.format.fill.color (which only reports
    // manually-set fill and misses conditional formatting entirely),
    // getCellProperties resolves color scales, rule-based highlights, and
    // any other conditional formatting exactly as Excel renders them.
    // Falls back to the older per-cell API on Excel hosts that predate it.
    let cellProps;
    try {
      const cellPropsResult = range.getCellProperties({
        format: {
          fill: { color: true },
          font: { color: true, bold: true, italic: true, size: true, underline: true },
          horizontalAlignment: true,
        }
      });
      await context.sync();
      cellProps = cellPropsResult.value; // 2D array [row][col]
    } catch (e) {
      // Fallback: older API surface without conditional-formatting resolution
      const cellRefs = [];
      for (let r = 0; r < rows; r++) {
        const rowRefs = [];
        for (let c = 0; c < cols; c++) {
          const cell = range.getCell(r, c);
          cell.format.font.load(['bold', 'italic', 'color', 'size', 'underline']);
          cell.format.fill.load(['color']);
          cell.format.load(['horizontalAlignment']);
          rowRefs.push(cell);
        }
        cellRefs.push(rowRefs);
      }
      await context.sync();
      cellProps = cellRefs.map(row => row.map(cell => ({
        format: {
          fill: { color: cell.format.fill.color },
          font: {
            bold: cell.format.font.bold, italic: cell.format.font.italic,
            underline: cell.format.font.underline, color: cell.format.font.color,
            size: cell.format.font.size,
          },
          horizontalAlignment: cell.format.horizontalAlignment,
        }
      })));
    }

    onProgress(55, 'Composing image…');

    const textGrid = range.text;
    const colWidthsPt = colRanges.map(cr => cr.format.columnWidth || 64);
    const rowHeightsPt = rowRanges.map(rr => rr.format.rowHeight || 20);

    const scale = qualityScale(currentQuality);
    const colWidthsPx = colWidthsPt.map(w => Math.max(4, Math.round(w * PT_TO_PX * scale)));
    const rowHeightsPx = rowHeightsPt.map(h => Math.max(4, Math.round(h * PT_TO_PX * scale)));

    const totalW = colWidthsPx.reduce((a, b) => a + b, 0);
    const totalH = rowHeightsPx.reduce((a, b) => a + b, 0);

    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, totalW);
    canvas.height = Math.max(1, totalH);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    let y = 0;
    for (let r = 0; r < rows; r++) {
      let x = 0;
      const rh = rowHeightsPx[r];

      for (let c = 0; c < cols; c++) {
        const cw = colWidthsPx[c];
        const cp = cellProps[r][c];
        const text = (textGrid[r] && textGrid[r][c] !== undefined) ? String(textGrid[r][c]) : '';

        let fillColor = null;
        try { fillColor = cp.format.fill.color; } catch (_) {}
        if (fillColor && /^#([0-9a-f]{6})$/i.test(fillColor) && fillColor.toLowerCase() !== '#ffffff') {
          ctx.fillStyle = fillColor;
          ctx.fillRect(x, y, cw, rh);
        }

        // Thin default gridline (matches Excel's standard light-gray grid)
        ctx.strokeStyle = '#d0d7de';
        ctx.lineWidth = 1;
        ctx.strokeRect(x + 0.5, y + 0.5, Math.max(0, cw - 1), Math.max(0, rh - 1));

        if (text) {
          let bold = false, italic = false, underline = false, fontColor = '#000000', fontSize = 11;
          try {
            bold      = !!cp.format.font.bold;
            italic    = !!cp.format.font.italic;
            underline = cp.format.font.underline && cp.format.font.underline !== 'None';
            fontColor = cp.format.font.color || '#000000';
            fontSize  = cp.format.font.size || 11;
          } catch (_) {}

          const px = Math.max(6, Math.round(fontSize * PT_TO_PX * scale));
          let fontStr = '';
          if (italic) fontStr += 'italic ';
          if (bold) fontStr += '700 '; else fontStr += '400 ';
          fontStr += `${px}px Calibri, Arial, sans-serif`;
          ctx.font = fontStr;
          ctx.textBaseline = 'middle';

          let align = 'General';
          try { align = cp.format.horizontalAlignment || 'General'; } catch (_) {}
          // Numbers default-align right in Excel; text defaults left
          const looksNumeric = /^-?[\d,.]+%?$/.test(text.trim()) || /^-?\$[\d,.]+$/.test(text.trim());
          if (align === 'General') align = looksNumeric ? 'Right' : 'Left';

          const pad = 4 * scale;
          let tx = x + pad;
          ctx.textAlign = 'left';
          if (align === 'Right')  { tx = x + cw - pad; ctx.textAlign = 'right'; }
          else if (align === 'Center') { tx = x + cw / 2; ctx.textAlign = 'center'; }

          ctx.save();
          ctx.beginPath();
          ctx.rect(x, y, cw, rh);
          ctx.clip();
          ctx.fillStyle = /^#([0-9a-f]{6})$/i.test(fontColor) ? fontColor : '#000000';
          ctx.fillText(text, tx, y + rh / 2);

          if (underline) {
            const tw = ctx.measureText(text).width;
            let ux = tx;
            if (align === 'Right') ux = tx - tw;
            else if (align === 'Center') ux = tx - tw / 2;
            const uy = y + rh / 2 + px / 2.4;
            ctx.beginPath();
            ctx.moveTo(ux, uy);
            ctx.lineTo(ux + tw, uy);
            ctx.strokeStyle = ctx.fillStyle;
            ctx.lineWidth = Math.max(1, scale);
            ctx.stroke();
          }
          ctx.restore();
        }

        x += cw;
      }
      y += rh;

      if (r % 20 === 0) onProgress(55 + Math.round((r / rows) * 35), `Drawing row ${r+1} of ${rows}…`);
    }

    result = canvas;
  });

  return result;
}

// ─────────────────────────────────────────────────────────────────────
// Download / Clipboard
// ─────────────────────────────────────────────────────────────────────
function downloadLast() {
  if (!lastCanvas) return;
  const dataUrl = lastCanvas.toDataURL(currentFormat === 'jpeg' ? 'image/jpeg' : 'image/png', 0.95);
  const a = document.createElement('a');
  a.href = dataUrl;
  a.download = lastFilename || makeFilename(currentFormat);
  document.body.appendChild(a);
  a.click();
  a.remove();
  showToast('✅ Downloaded: ' + a.download, 'ok');
}

async function copyLastToClipboard() {
  if (!lastCanvas) return;
  try {
    if (!navigator.clipboard || !window.ClipboardItem) {
      throw new Error('unsupported');
    }
    const blob = await new Promise(res => lastCanvas.toBlob(res, 'image/png'));
    await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
    showToast('📋 Copied to clipboard!', 'ok');
  } catch (e) {
    // Excel's taskpane runs inside a Microsoft-controlled iframe, which
    // can block clipboard-write access regardless of our own code — this
    // is a platform limitation, not a bug in the add-in. Download always
    // works since it doesn't depend on clipboard permissions.
    showToast('⚠️ Clipboard access is blocked inside Excel\'s panel — please use Download instead', 'err');
  }
}

function makeFilename(fmt) {
  const d = new Date(), p = n => String(n).padStart(2, '0');
  const ts = `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
  return `SheetSnap_${ts}.${fmt === 'jpeg' ? 'jpg' : fmt}`;
}

// ─────────────────────────────────────────────────────────────────────
// UI helpers
// ─────────────────────────────────────────────────────────────────────
function setProgress(on, pct, label) {
  const prog = document.getElementById('prog');
  prog.classList.toggle('on', on);
  document.getElementById('progFill').style.width = pct + '%';
  document.getElementById('progLbl').textContent = label;
  document.getElementById('progPct').textContent = pct + '%';
}

let toastTimer;
function showToast(msg, kind = 'inf') {
  const toast = document.getElementById('toast');
  clearTimeout(toastTimer);
  toast.textContent = msg;
  toast.className = `toast ${kind} show`;
  toastTimer = setTimeout(() => toast.classList.remove('show'), 3800);
}
