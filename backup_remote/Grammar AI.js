/***** 🧠 AI GRAMMAR ASSIST — Google Sheets (LanguageTool, no API key) *****
 * Menu: 🧠 AI Assist
 *   • Fix Grammar — Selected (Quick)
 *   • Fix Grammar — Column (by Header)
 *   • Undo Last Fix (Selected)
 *
 * Behavior:
 *   • Uses LanguageTool (https://api.languagetool.org) to fix grammar/spelling.
 *   • Skips header row, formulas, and cells with rich hyperlinks to avoid breaking links.
 *   • Stores original text in a cell Note so you can Undo Last Fix (Selected).
 *   • Default language: en-US (change AIGA_LT_LANG if needed).
 ***********************************************************************/

// === CONFIG ==========================================================
const AIGA_LT_ENDPOINT = 'https://api.languagetool.org/v2/check';
const AIGA_LT_LANG = 'en-US';            // Change to 'en-GB', etc., if you prefer
const AIGA_NOTE_PREFIX = '[AI Grammar] Original:'; // Note prefix for undo

// === ONE-TIME INSTALL ===============================================
function AIGA_install() {
  // Create a dedicated onOpen trigger for this menu (won’t touch your other triggers)
  const ssId = SpreadsheetApp.getActive().getId();
  // Remove older duplicates of our handler
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'AIGA_onOpen_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('AIGA_onOpen_').forSpreadsheet(ssId).onOpen().create();
  SpreadsheetApp.getUi().alert('🧠 AI Assist installed. Reopen the spreadsheet to see the menu.');
}

// === MENU ============================================================
function AIGA_onOpen_() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 AI Assist')
    .addItem('Fix Grammar — Selected (Quick)', 'AIGA_fixSelectedQuick')
    .addItem('Fix Grammar — Column (by Header)', 'AIGA_fixColumnByHeader')
    .addSeparator()
    .addItem('Undo Last Fix (Selected)', 'AIGA_restoreFromNotesSelected')
    .addToUi();
}

// === ACTIONS =========================================================
function AIGA_fixSelectedQuick() {
  const sh = SpreadsheetApp.getActiveSheet();
  const rg = sh.getActiveRange();
  if (!rg) return SpreadsheetApp.getUi().alert('Select one or more cells to fix.');
  const rows = rg.getNumRows(), cols = rg.getNumColumns();
  const startR = rg.getRow(), startC = rg.getColumn();

  let processed = 0, changed = 0, skipped = 0, errors = 0;

  for (let r = 0; r < rows; r++) {
    const rowIdx = startR + r;
    if (rowIdx === 1) { skipped++; continue; } // skip header row
    for (let c = 0; c < cols; c++) {
      const colIdx = startC + c;
      const cell = sh.getRange(rowIdx, colIdx);

      // Skip formulas
      if (String(cell.getFormula() || '').trim()) { skipped++; continue; }

      // Skip rich-text hyperlink cells to avoid breaking links
      if (AIGA_cellHasLinks_(cell)) { skipped++; continue; }

      const txt = String(cell.getDisplayValue() || '').trim();
      if (!txt) { skipped++; continue; }

      processed++;
      try {
        const fixed = AIGA_grammarFixText_(txt);
        if (fixed && fixed !== txt) {
          // Preserve original in Note (append if note already exists)
          const existingNote = cell.getNote() || '';
          const stamp = new Date().toISOString();
          const origBlock = `${AIGA_NOTE_PREFIX} ${txt}\n(ts: ${stamp})`;
          const newNote = existingNote
            ? (existingNote + '\n\n' + origBlock)
            : origBlock;
          try { cell.setNote(newNote); } catch (_) {}
          cell.setValue(fixed);
          changed++;
        } else {
          skipped++;
        }
      } catch (e) {
        errors++;
      }
    }
  }

  SpreadsheetApp.getActive().toast(
    `AI Assist — Selected:\nChanged: ${changed}\nProcessed: ${processed}\nSkipped: ${skipped}${errors ? `\nErrors: ${errors}` : ''}`,
    '🧠 AI Assist', 7
  );
}

function AIGA_fixColumnByHeader() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
  if (lastRow < 2) return ui.alert('No data to process.');

  const headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v => String(v||'').trim());
  const prompt = ui.prompt(
    'Fix Grammar — Column by Header',
    `Enter the column header exactly as it appears (e.g., "Task Details" or "NOTE").\nAvailable headers:\n${headers.join(' | ')}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  const target = String(prompt.getResponseText() || '').trim().toLowerCase();
  if (!target) return ui.alert('No header entered.');

  const col = AIGA_findHeaderIndex_(headers, target);
  if (!col) return ui.alert(`Could not find header: ${target}`);

  let processed = 0, changed = 0, skipped = 0, errors = 0;

  for (let r = 2; r <= lastRow; r++) {
    const cell = sh.getRange(r, col);

    // Skip formulas
    if (String(cell.getFormula() || '').trim()) { skipped++; continue; }

    // Skip rich links
    if (AIGA_cellHasLinks_(cell)) { skipped++; continue; }

    const txt = String(cell.getDisplayValue() || '').trim();
    if (!txt) { skipped++; continue; }

    processed++;
    try {
      const fixed = AIGA_grammarFixText_(txt);
      if (fixed && fixed !== txt) {
        const existingNote = cell.getNote() || '';
        const stamp = new Date().toISOString();
        const origBlock = `${AIGA_NOTE_PREFIX} ${txt}\n(ts: ${stamp})`;
        const newNote = existingNote
          ? (existingNote + '\n\n' + origBlock)
          : origBlock;
        try { cell.setNote(newNote); } catch (_) {}
        cell.setValue(fixed);
        changed++;
      } else {
        skipped++;
      }
    } catch (e) {
      errors++;
    }
  }

  ui.alert(`AI Assist — Column "${headers[col-1]}":\nChanged: ${changed}\nProcessed: ${processed}\nSkipped: ${skipped}${errors ? `\nErrors: ${errors}` : ''}`);
}

function AIGA_restoreFromNotesSelected() {
  const sh = SpreadsheetApp.getActiveSheet();
  const rg = sh.getActiveRange();
  if (!rg) return SpreadsheetApp.getUi().alert('Select the cell(s) to restore.');
  const rows = rg.getNumRows(), cols = rg.getNumColumns();
  const startR = rg.getRow(), startC = rg.getColumn();

  let restored = 0, skipped = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = sh.getRange(startR + r, startC + c);
      const note = cell.getNote() || '';
      if (!note || note.indexOf(AIGA_NOTE_PREFIX) === -1) { skipped++; continue; }

      // Restore the most recent original block (the last occurrence)
      const blocks = note.split('\n\n').filter(x => x.indexOf(AIGA_NOTE_PREFIX) === 0 || x.indexOf('\n' + AIGA_NOTE_PREFIX) >= 0);
      const lastBlock = blocks[blocks.length - 1];
      const m = lastBlock.match(new RegExp('^' + AIGA_escapeRegExp_(AIGA_NOTE_PREFIX) + '\\s([\\s\\S]+?)(?:\\n\\(ts:.*)?$'));
      if (m && m[1]) {
        const original = m[1];
        cell.setValue(original);
        // Remove only the last block from the note
        const idx = note.lastIndexOf(lastBlock);
        const newNote = (idx >= 0)
          ? (note.slice(0, idx).trim() + (idx > 0 ? '' : ''))
          : '';
        cell.setNote(newNote);
        restored++;
      } else {
        skipped++;
      }
    }
  }

  SpreadsheetApp.getActive().toast(`Undo complete. Restored: ${restored}${skipped ? ` | Skipped: ${skipped}` : ''}`, '🧠 AI Assist', 5);
}

// === GRAMMAR ENGINE (LanguageTool) ==================================
function AIGA_grammarFixText_(text) {
  // Call public LanguageTool endpoint (rate-limited). For heavier use, host your own LT server.
  const payload = {
    text: text,
    language: AIGA_LT_LANG
  };
  const resp = UrlFetchApp.fetch(AIGA_LT_ENDPOINT, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('LanguageTool error: ' + resp.getContentText());

  const data = JSON.parse(resp.getContentText());
  const matches = (data && data.matches) ? data.matches : [];
  if (!matches.length) return text;

  // Apply replacements left→right adjusting offsets
  // Use first suggested replacement for each match
  let out = text;
  let delta = 0;
  matches
    .filter(m => m && m.replacements && m.replacements.length && Number.isInteger(m.offset) && Number.isInteger(m.length))
    .sort((a,b) => a.offset - b.offset)
    .forEach(m => {
      const repl = m.replacements[0].value || '';
      const start = m.offset + delta;
      const end = start + m.length;
      // basic bounds check
      if (start < 0 || end > out.length) return;
      out = out.substring(0, start) + repl + out.substring(end);
      delta += (repl.length - m.length);
    });

  return out;
}

// === HELPERS =========================================================
function AIGA_cellHasLinks_(cell) {
  try {
    if (/^=HYPERLINK\(/i.test(String(cell.getFormula() || ''))) return true;
    const rtv = cell.getRichTextValue && cell.getRichTextValue();
    if (rtv && rtv.getRuns) {
      const runs = rtv.getRuns();
      for (let i = 0; i < runs.length; i++) {
        if (runs[i].getLinkUrl && runs[i].getLinkUrl()) return true;
      }
    }
  } catch (_) {}
  return false;
}

function AIGA_findHeaderIndex_(headers, targetLower) {
  const normed = headers.map(h => String(h||'').toLowerCase().trim());
  let idx = normed.indexOf(targetLower);
  if (idx >= 0) return idx + 1;
  // mild fuzzy contains
  for (let i = 0; i < normed.length; i++) {
    if (normed[i].includes(targetLower)) return i + 1;
  }
  return 0;
}

function AIGA_escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
