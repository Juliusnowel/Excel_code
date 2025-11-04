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

    if (!willMove){
      // If the status is a lifecycle event that needs stamping, do it even if row stays on same sheet
      // try {
      //   const idx = map;
      //   // ensure late column exists
      //   KNB_ensureLateColumnHere(sh);
      //   // handle For Revision immediately; For Approval/Done stamping also acceptable here
      //   if (['For Revision','For Approval','Done'].includes(newStatus)) {
      //     KNB_handleLateForStatusChange_(sh, row, idx, newStatus);
      //   }
      // } catch(_){}
      KNB_manageForApprovalFreeze_(sh, row, map, oldStatus, newStatus); 
      KNB_noteRate_();
      SpreadsheetApp.getActive().toast(`Status set to ${newStatus}.`, 'Mover', 2);
      return;
    }


    // === Signal ASAP via toast; keep it lightweight ===
    // show toast FIRST to avoid waiting on any other calls
    SpreadsheetApp.getActive().toast('Please wait…', 'Mover', 3);

    // then set lightweight flags
    KNB_setBusy_(3000);
    KNB_suppressEdits_(1500);

    // DO NOT flush here; it delays toast rendering
    // SpreadsheetApp.flush();

    // Set Start Date cheaply if entering In Progress
    const cStart = map[KNB_CFG.COL.START] || 0;
    if (newStatus === 'In Progress' && cStart){
      const cell = sh.getRange(row, cStart);
      if (!cell.getValue()){
        cell.setValue(new Date());
        try { cell.setNumberFormat('yyyy-mm-dd'); } catch(_){}
      }
    }

    const cEnd = map[KNB_CFG.COL.END] || 0;
    if (newStatus === 'Done' && cEnd){
      const cell = sh.getRange(row, cEnd);
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

    // Freeze/Unfreeze driver for "For Approval"
    KNB_manageForApprovalFreeze_(sh, row, map, oldStatus, newStatus);

    // Stamp Day Count / Late / Revision resets on source row BEFORE moving,
    // so the moved copy contains the frozen values.
    // try {
    //   KNB_ensureLateColumnHere(sh);
    //   if (['For Revision','For Approval','Done'].includes(newStatus)){
    //     KNB_handleLateForStatusChange_(sh, row, map, newStatus);
    //   }
    // } catch(_){}

    // Move the row, then toast
    KNB_moveRow_(sh, row, destGid);
    KNB_noteRate_();
    SpreadsheetApp.getActive().toast('Task moved.', 'Mover', 3);
  } catch (err){
    Logger.log(err && err.stack ? err.stack : err);
    // SpreadsheetApp.getActive().toast(`Mover error: ${String(err && err.message || err)}`, 'Mover', 8);
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
    if (!['Adhoc Task','Low Prio','Mid Prio','High Prio', 'Urgent'].includes(pr)) issues.push('Invalid Task Priority Value');

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
// function KNB_reconcileGated(){
//   const ss = SpreadsheetApp.getActive();
//   // Requested -> In Progress if passes gate
//   KNB_eachRow(ss, KNB_CFG.GID.REQUESTED, (sh, r, idx)=>{
//     const st = String(sh.getRange(r, idx[KNB_CFG.COL.STATUS]).getDisplayValue()||'').trim();
//     if (st === 'In Progress' && KNB_gateAllows_(sh, r, idx, 'In Progress')){
//       KNB_moveRow_(sh, r, KNB_CFG.GID.INPROGRESS);
//     }
//   });
//   // In Progress -> For Approval if passes gate
//   KNB_eachRow(ss, KNB_CFG.GID.INPROGRESS, (sh, r, idx)=>{
//     const st = String(sh.getRange(r, idx[KNB_CFG.COL.STATUS]).getDisplayValue()||'').trim();
//     if (st === 'For Approval' && KNB_gateAllows_(sh, r, idx, 'For Approval')){
//       KNB_manageForApprovalFreeze_(sh, r, idx, /*old*/null, /*new*/'For Approval'); // <-- add
//       KNB_moveRow_(sh, r, KNB_CFG.GID.FORAPPROVAL);
//     }
//   });
//   SpreadsheetApp.getUi().alert('Reconcile (Gated) complete ✅');
// }

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
      if (want && want !== g) {
        // keep freeze column consistent even when status was pasted/edited in bulk
        KNB_manageForApprovalFreeze_(sh, r, idx, /*old*/null, /*new*/st, {forceReconcile:true});
        KNB_moveRow_(sh, r, want);
      }
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
    const st  = String(sh.getRange(row, idx[KNB_CFG.COL.STATUS]).getDisplayValue()||'').trim();
    const dest= map[st];
    if (dest && dest !== sh.getSheetId()) {
      KNB_manageForApprovalFreeze_(sh, row, idx, /*old*/null, /*new*/st, {forceReconcile:true});
      KNB_moveRow_(sh, row, dest);
    }
  }
  SpreadsheetApp.getActive().toast('Force Move complete.', 'Mover', 4);
}

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

    // Ensure the "Late or not?" column exists (and hidden) on both sheets
    // try { KNB_ensureLateColumnHere(fromSheet); } catch(_){}
    // try { KNB_ensureLateColumnHere(toSheet); } catch(_){}

    // Make sure destination has enough columns (to avoid range errors)
    const srcCols = fromSheet.getLastColumn();
    const dstMax  = toSheet.getMaxColumns();
    if (dstMax < srcCols) toSheet.insertColumnsAfter(dstMax, srcCols - dstMax);

    // === smarter copy: map by header name when possible, fallback to same column index ===
    const destRow = Math.max(2, toSheet.getLastRow() + 1);

    // Build header maps (name -> col) and reverse source lookup (col -> name)
    const srcHdr = KNB_headerIndex_CACHED_(fromSheet) || {};
    const dstHdr = KNB_headerIndex_CACHED_(toSheet) || {};
    const reverseSrc = [];
    Object.keys(srcHdr).forEach(h => { reverseSrc[srcHdr[h]] = h; });

    // Copy each source column cell to matching destination column:
    //  - if header exists in source and destination, copy to destination's column for that header
    //  - otherwise, copy to same numeric column index on destination (fallback)
    for (let c = 1; c <= srcCols; c++){
      try {
        const headerName = reverseSrc[c] || '';
        const destCol = (headerName && dstHdr[headerName]) ? dstHdr[headerName] : c;
        // ensure destination has that column (we made sure dstMax >= srcCols above)
        const srcRange = fromSheet.getRange(row, c, 1, 1);
        const dstRange = toSheet.getRange(destRow, destCol, 1, 1);
        // copy everything for that cell (formats, rich text, notes)
        srcRange.copyTo(dstRange, { contentsOnly: false });
      } catch (errCell) {
        // continue copying other cells even if one cell copy fails
        try { Logger.log('Copy cell failed col=' + c + ' -> ' + (reverseSrc[c] || c) + ' : ' + (errCell && errCell.message || errCell)); } catch(_){}
      }
    }

    // Preserve row height, then remove source
    try { KNB_withRetry_(() => toSheet.setRowHeight(destRow, fromSheet.getRowHeight(row)), 3, 'setRowHeight'); } catch(_){}
    // Remove source row
    KNB_withRetry_(() => fromSheet.deleteRow(row), 4, 'deleteRow');

    // Put user on the moved row
    try {
      KNB_withRetry_(() => {
        SpreadsheetApp.getActive().setActiveSheet(toSheet);
        SpreadsheetApp.getActive().setActiveSelection(toSheet.getRange(destRow, 1));
      }, 3, 'activate moved row');
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

// function KNB_backfillEndDateForDoneHere(){
//   const sh  = SpreadsheetApp.getActiveSheet();
//   const idx = KNB_headerIndex_(sh);
//   const cSts = idx[KNB_CFG.COL.STATUS], cEnd = idx[KNB_CFG.COL.END];
//   if (!cSts || !cEnd) return;

//   const last = Math.max(2, sh.getLastRow());
//   const sts  = sh.getRange(2, cSts, last-1, 1).getDisplayValues();
//   const ends = sh.getRange(2, cEnd, last-1, 1).getValues();

//   const toWrite = [];
//   for (let i=0;i<sts.length;i++){
//     if (String(sts[i][0]||'').trim()==='Done' && !ends[i][0]){
//       toWrite.push([new Date()]);
//     } else {
//       toWrite.push([ends[i][0]]);
//     }
//   }
//   if (toWrite.length){
//     sh.getRange(2, cEnd, toWrite.length, 1).setValues(toWrite);
//     try { sh.getRange(2, cEnd, toWrite.length, 1).setNumberFormat('yyyy-mm-dd'); } catch(_){}
//   }
//   SpreadsheetApp.getActive().toast('Backfilled End Date for Done rows.', 'Done', 3);
// }

function KNB_manageForApprovalFreeze_(sh, row, map, oldStatus, newStatus, opts){
  const cFrz = KNB_ensureFreezeColumn_(sh, map);
  const force = opts && opts.forceReconcile === true;

  if (newStatus === 'For Approval'){
    const cell = sh.getRange(row, cFrz);
    if (!cell.getValue()){
      cell.setValue(new Date());
      try { cell.setNumberFormat('yyyy-mm-dd'); } catch(_){}
    }
    return;
  }

  if (force || oldStatus === 'For Approval' || oldStatus == null){
    sh.getRange(row, cFrz).clearContent();
  }
}



