/* =========================
   MOVER (gated + strict + force)
========================= */
function KNB_moverOnEdit_(e){
  try {
    if (!e || !e.range) return;

    const sh  = e.range.getSheet();
    const gid = sh.getSheetId();
    if (!KNB_allGids_().includes(gid)) return;

    // FAST header map
    const map = KNB_headerIndex_CACHED_(sh);
    const cStatus = map[KNB_CFG.COL.STATUS];
    if (!cStatus || e.range.getColumn() !== cStatus) return;
    if (KNB_isSuppressed_()) return;

    const newStatus = String((e.value ?? e.range.getValue()) || '').trim();
    const oldStatus = (typeof e.oldValue !== 'undefined') ? String(e.oldValue || '') : null;
    const destGid   = KNB_CFG.ROUTE[newStatus];

    const willMove = !!destGid && destGid !== gid;
    const row = e.range.getRow();
    if (row === 1) return;

    // Rate limit & busy guard (before doing anything heavy)
    if (willMove && (!KNB_isRateOk_() || KNB_isBusy_())){
      KNB_suppressEdits_(750);
      e.range.setValue(oldStatus || KNB_CFG.REVERT_FALLBACK);
      SpreadsheetApp.getActive().toast('Please wait…', 'Mover', 3);
      return;
    }

    // If it’s not a routed status (or stays on same board), exit quickly
    if (!willMove){
      KNB_noteRate_();
      SpreadsheetApp.getActive().toast(`Status set to ${newStatus}.`, 'Mover', 2);
      return;
    }

    // === Show spinner ASAP (then do the work behind it) ===
    KNB_setBusy_(3000);
    KNB_suppressEdits_(1500);
    KNB_UI_showBusy_('Moving task…', 3500);
    SpreadsheetApp.flush();

    // Set Start Date cheaply if entering In Progress
    const cStart = map[KNB_CFG.COL.START] || 0;
    if (newStatus === 'In Progress' && cStart){
      const cell = sh.getRange(row, cStart);
      if (!cell.getValue()){
        cell.setValue(new Date());
        try { cell.setNumberFormat('yyyy-mm-dd'); } catch(_){}
      }
    }

    // Gate checks while spinner is visible
    if (!KNB_gateAllows_(sh, row, map, newStatus)){
      KNB_suppressEdits_(750);
      e.range.setValue(oldStatus || KNB_CFG.REVERT_FALLBACK);
      return;
    }

    // Move the row (preserves formats/notes/rich text + hides storage cols)
    KNB_moveRow_(sh, row, destGid);
    KNB_noteRate_();
    SpreadsheetApp.getActive().toast('Task moved.', 'Mover', 3);

  } catch (err){
    Logger.log(err && err.stack ? err.stack : err);
    SpreadsheetApp.getActive().toast(`Mover error: ${String(err && err.message || err)}`, 'Mover', 8);
  } finally {
    // Optional: the busy modal auto-closes; keep this only if you want a hard close.
    // try { KNB_UI_closeBusy_(); } catch(_){}
  }
}

function KNB_resetSheetBodyBackgroundHere(){
  const sh = SpreadsheetApp.getActiveSheet();
  const lastRow = Math.max(2, sh.getLastRow());
  const lastCol = Math.max(1, sh.getLastColumn());
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, lastCol).setBackground(null); // or '#ffffff'
  }
}

// GATES
function KNB_gateAllows_(sh, row, idx, newStatus){
  if (newStatus === 'In Progress'){
    const need = [KNB_CFG.COL.DEPARTMENT, KNB_CFG.COL.ASSIGNEE, KNB_CFG.COL.CLIENT,
                  KNB_CFG.COL.TASK, KNB_CFG.COL.PRIORITY, KNB_CFG.COL.CREATED, KNB_CFG.COL.DETAILS];
    const missing = need.filter(h=>{
      const c = idx[h]; if(!c) return true;
      const v = String(sh.getRange(row, c).getDisplayValue()||'').trim();
      if (h === KNB_CFG.COL.DETAILS){
        // accept if NOTE has HTML or the cell shows 📝
        const cell = sh.getRange(row, c);
        const hasNote = String(cell.getNote()||'').trim().length>0;
        return !(hasNote || v === '📝');
      }
      return !v;
    });
    if (missing.length){
      SpreadsheetApp.getActive().toast('Fill required: '+missing.join(', '), 'Gate', 6);
      return false;
    }
  }

  if (newStatus === 'For Approval'){
    // Priority mid/high, Deliverable non-empty, Screenshot is valid URL / HYPERLINK
    const prC = idx[KNB_CFG.COL.PRIORITY], dlC = idx[KNB_CFG.COL.DELIVERABLE], scC = idx[KNB_CFG.COL.SCREENSHOT];
    const issues = [];

    const pr = prC ? String(sh.getRange(row, prC).getDisplayValue()||'').trim() : '';
    if (!['Mid Prio','High Prio'].includes(pr)) issues.push('Task Priority must be Mid/High');

    const dl = dlC ? String(sh.getRange(row, dlC).getDisplayValue()||'').trim() : '';
    if (!dl) issues.push('Deliverable required');

    const okUrl = scC ? KNB_cellHasValidUrl_(sh.getRange(row, scC)) : false;
    if (!okUrl) issues.push('Screenshot must be a valid http(s) URL / HYPERLINK');

    if (issues.length){
      SpreadsheetApp.getActive().toast('Fix: '+issues.join(', '), 'Gate', 8);
      return false;
    }
  }
  return true;
}

// Reconcile (Gated)
function KNB_reconcileGated(){
  const ss = SpreadsheetApp.getActive();
  // Requested -> In Progress if passes gate
  KNB_eachRow(ss, KNB_CFG.GID.REQUESTED, (sh, r, idx)=>{
    const st = String(sh.getRange(r, idx[KNB_CFG.COL.STATUS]).getDisplayValue()||'').trim();
    if (st === 'In Progress' && KNB_gateAllows_(sh, r, idx, 'In Progress')){
      KNB_moveRow_(sh, r, KNB_CFG.GID.INPROGRESS);
    }
  });
  // In Progress -> For Approval if passes gate
  KNB_eachRow(ss, KNB_CFG.GID.INPROGRESS, (sh, r, idx)=>{
    const st = String(sh.getRange(r, idx[KNB_CFG.COL.STATUS]).getDisplayValue()||'').trim();
    if (st === 'For Approval' && KNB_gateAllows_(sh, r, idx, 'For Approval')){
      KNB_moveRow_(sh, r, KNB_CFG.GID.FORAPPROVAL);
    }
  });
  SpreadsheetApp.getUi().alert('Reconcile (Gated) complete ✅');
}

// Reconcile (Strict)
function KNB_reconcileStrict(){
  const ss = SpreadsheetApp.getActive();
  const map = {
    'Requested':    KNB_CFG.GID.REQUESTED,
    'In Progress':  KNB_CFG.GID.INPROGRESS,
    'For Approval': KNB_CFG.GID.FORAPPROVAL,
    'Done':         KNB_CFG.GID.DONE
  };
  [KNB_CFG.GID.REQUESTED, KNB_CFG.GID.INPROGRESS, KNB_CFG.GID.FORAPPROVAL].forEach(g=>{
    KNB_eachRow(ss, g, (sh, r, idx)=>{
      const st = String(sh.getRange(r, idx[KNB_CFG.COL.STATUS]).getDisplayValue()||'').trim();
      const want = map[st];
      if (want && want !== g) KNB_moveRow_(sh, r, want);
    });
  });
  SpreadsheetApp.getUi().alert('Reconcile (Strict) complete ✅');
}

// Force move (Selected)
function KNB_forceMoveSelected(){
  const sh = SpreadsheetApp.getActiveSheet();
  const rg = sh.getActiveRange();
  if (!rg) return SpreadsheetApp.getUi().alert('Select rows to move.');
  const idx = KNB_headerIndex_(sh);
  const map = {
    'Requested': KNB_CFG.GID.REQUESTED,
    'In Progress': KNB_CFG.GID.INPROGRESS,
    'For Approval': KNB_CFG.GID.FORAPPROVAL,
    'For Revision': KNB_CFG.GID.REQUESTED,
    'Done': KNB_CFG.GID.DONE
  };
  for (let i=0;i<rg.getNumRows();i++){
    const row = rg.getRow()+i; if (row<2) continue;
    const st = String(sh.getRange(row, idx[KNB_CFG.COL.STATUS]).getDisplayValue()||'').trim();
    const dest = map[st];
    if (dest && dest !== sh.getSheetId()) KNB_moveRow_(sh, row, dest);
  }
  SpreadsheetApp.getActive().toast('Force Move complete.', 'Mover', 4);
}

// Core move (preserve values, formats, validation, NOTES, rich text)
// Core move (preserve values, formats, validation, NOTES, rich text)
// + ensure & hide Task Details storage columns on both ends
function KNB_moveRow_(fromSheet, row, destGid){
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(200)) lock.waitLock(5000);
  try{
    const toSheet = KNB_sheetById_(destGid);
    if (!toSheet) throw new Error('Destination sheet not found.');

    // Ensure storage columns exist (and get hidden) on both ends
    try { KNB_TASK_ensureDetailsColumns_Here(fromSheet); } catch(_){}
    try { KNB_TASK_ensureDetailsColumns_Here(toSheet);   } catch(_){}

    // Make sure destination has enough columns
    const srcCols = fromSheet.getLastColumn();
    const dstMax  = toSheet.getMaxColumns();
    if (dstMax < srcCols) toSheet.insertColumnsAfter(dstMax, srcCols - dstMax);

    // Copy the entire row with formats/notes/rich text
    const destRow = Math.max(2, toSheet.getLastRow() + 1);
    fromSheet.getRange(row, 1, 1, srcCols)
      .copyTo(toSheet.getRange(destRow, 1, 1, srcCols), { contentsOnly:false });

    // Preserve row height, then remove source
    try { toSheet.setRowHeight(destRow, fromSheet.getRowHeight(row)); } catch(_){}
    fromSheet.deleteRow(row);

    // Put user on the moved row
    try {
      SpreadsheetApp.getActive()
        .setActiveSheet(toSheet)
        .setActiveSelection(toSheet.getRange(destRow, 1));
    } catch(_){}

    // Re-hide storage columns on destination in case copy made them visible
    try { KNB_TASK_ensureDetailsColumns_Here(toSheet); } catch(_){}
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

function KNB_jumpToLastRowHere(){
  const sh = SpreadsheetApp.getActiveSheet();
  const r = Math.max(2, sh.getLastRow());
  SpreadsheetApp.getActive().setActiveSelection(sh.getRange(r, 1));
  SpreadsheetApp.getActive().toast('Jumped to row '+r, 'Nav', 3);
}