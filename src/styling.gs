/* =========================
   OPTIONAL styling / DV helpers (never changes headers text)
========================= */
// function KNB_applyStatusDropdownsAll(){
//   KNB_allGids_().forEach(gid=>{
//     const sh = KNB_sheetById_(gid); if (!sh) return;
//     const idx = KNB_headerIndex_(sh);
//     const col = idx[KNB_CFG.COL.STATUS]; if (!col) return;

//     const lastRow = sh.getLastRow();
//     if (lastRow < 2) return;

//     const rule = SpreadsheetApp.newDataValidation()
//       .requireValueInList(KNB_CFG.STATUSES, true)
//       .setAllowInvalid(false)
//       .setHelpText('Choose a Status')
//       .build();

//     sh.getRange(2, col, lastRow - 1, 1).setDataValidation(rule);
//   });
// }

// function KNB_applyDepartmentDropdownHere(){
//   const sh = SpreadsheetApp.getActiveSheet();
//   const idx = KNB_headerIndex_(sh);
//   const col = idx[KNB_CFG.COL.DEPARTMENT]; if (!col) return SpreadsheetApp.getUi().alert('Column "Department" not found.');
//   const lastRow = Math.max(2, sh.getLastRow());
//   const rule = SpreadsheetApp.newDataValidation()
//     .requireValueInList(KNB_CFG.DEPARTMENTS, true)
//     .setAllowInvalid(false)
//     .setHelpText('Choose a Department')
//     .build();
//   sh.getRange(2, col, lastRow-1, 1).setDataValidation(rule);
//   SpreadsheetApp.getActive().toast('Department dropdown applied.', 'Tasks', 3);
// }

// --- CF helper: tolerant, case-insensitive, strips NBSP & weird whitespace
function KNB_cfMatch_(a1, literal) {
  // Normalize the cell: TRIM, remove NBSP, collapse whitespace, LOWER
  // Then compare to a LOWER-cased literal
  const safe = String(literal || '').toLowerCase().replace(/"/g,'""');
  return `=LOWER(REGEXREPLACE(SUBSTITUTE(TRIM($${a1}2),CHAR(160)," "), "\\s+", " "))="${safe}"`;
}

function KNB_applyOwnerDropdownHere(){
  const sh = SpreadsheetApp.getActiveSheet();
  const idx = KNB_headerIndex_(sh);
  const col = idx[KNB_CFG.COL.OWNER];
  if (!col) return SpreadsheetApp.getUi().alert('Column "Owner" not found.');
  const lastRow = Math.max(2, sh.getLastRow());
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(KNB_CFG.ASSIGNEES, true)
    .setAllowInvalid(false)
    .setHelpText('Choose an Owner')
    .build();
  sh.getRange(2, col, lastRow-1, 1).setDataValidation(rule);
  SpreadsheetApp.getActive().toast('Owner dropdown applied.', 'Tasks', 3);
}

function KNB_ensureStyleHere(){
  const sh = SpreadsheetApp.getActiveSheet();

  // Header styling
  const lastCol = Math.max(1, sh.getLastColumn());
  sh.setFrozenRows(1);
  try { sh.setHiddenGridlines(true); } catch(_){}
  sh.getRange(1,1,1,lastCol)
    .setFontWeight('bold')
    .setBackground(KNB_CFG.HEADER_BG)
    .setFontColor(KNB_CFG.HEADER_FG)
    .setFontSize(14)
    .setVerticalAlignment('middle');

  const lastRow = Math.max(2, sh.getLastRow());
  const rows    = Math.max(1, lastRow - 1);
  const idx     = KNB_headerIndex_(sh);

  // Start clean: wipe ALL existing conditional rules (including any UI-made zebra/banding)
  try { sh.setConditionalFormatRules([]); } catch(_) {}
  const rules = [];

  // --- CHIP RULES ---
  // Status
  const cStatus = idx[KNB_CFG.COL.STATUS];
  if (cStatus){
    const a1 = KNB_colToA1_(cStatus);
    const r  = sh.getRange(2, cStatus, rows, 1);
    Object.entries(KNB_CFG.STATUS_COLORS || {}).forEach(([val,bg])=>{
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(KNB_cfMatch_(a1, val))
          .setBackground(bg)
          .setFontColor(KNB_pickTextColor_(bg))
          .setRanges([r])
          .build()
      );
    });
  }

  // Assignee
  const cAss = idx[KNB_CFG.COL.ASSIGNEE];
  if (cAss){
    const a1 = KNB_colToA1_(cAss);
    const r  = sh.getRange(2, cAss, rows, 1);
    Object.entries(KNB_CFG.ASSIGNEE_COLORS || {}).forEach(([val,bg])=>{
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(KNB_cfMatch_(a1, val))
          .setBackground(bg)
          .setRanges([r])
          .build()
      );
    });
  }

  // NEW: Owner (reuse assignee chips)
  const cOwner = idx[KNB_CFG.COL.OWNER];
  if (cOwner){
    const a1 = KNB_colToA1_(cOwner);
    const r  = sh.getRange(2, cOwner, rows, 1);
    Object.entries(KNB_CFG.ASSIGNEE_COLORS || {}).forEach(([val,bg])=>{
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(KNB_cfMatch_(a1, val))
          .setBackground(bg)
          .setRanges([r])
          .build()
      );
    });
  }

  // Priority
  const cPrio = idx[KNB_CFG.COL.PRIORITY];
  if (cPrio){
    const a1 = KNB_colToA1_(cPrio);
    const r  = sh.getRange(2, cPrio, rows, 1);
    Object.entries(KNB_CFG.PRIORITY_COLORS || {}).forEach(([val,bg])=>{
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(KNB_cfMatch_(a1, val))
          .setBackground(bg)
          .setFontColor(KNB_pickTextColor_(bg))
          .setRanges([r])
          .build()
      );
    });
  }

  // Client
  const cClient = idx[KNB_CFG.COL.CLIENT];
  if (cClient){
    const a1 = KNB_colToA1_(cClient);
    const r  = sh.getRange(2, cClient, rows, 1);
    Object.entries(KNB_CFG.CLIENT_COLORS || {}).forEach(([val,bg])=>{
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(KNB_cfMatch_(a1, val))
          .setBackground(bg)
          .setFontColor(KNB_pickTextColor_(bg))
          .setRanges([r])
          .build()
      );
    });
  }

  // Apply only the chip rules (no zebra!)
  sh.setConditionalFormatRules(rules);

  // Static column fills (these don’t block CF)
  Object.entries(KNB_CFG.COLUMN_COLORS || {}).forEach(([headerText,bg])=>{
    const c = idx[headerText]; if (!c) return;
    sh.getRange(2, c, rows, 1).setBackground(bg);
  });
}

function KNB_normalizeTextColumnsHere(){
  const sh  = SpreadsheetApp.getActiveSheet();
  const idx = KNB_headerIndex_(sh);
  const cols = [idx[KNB_CFG.COL.OWNER], idx[KNB_CFG.COL.ASSIGNEE], idx[KNB_CFG.COL.CLIENT], idx[KNB_CFG.COL.PRIORITY]].filter(Boolean);
  const last = sh.getLastRow();
  if (!cols.length || last < 2) return;

  cols.forEach(c => {
    const rng = sh.getRange(2, c, last-1, 1);
    const vals = rng.getValues().map(r => {
      const t = String(r[0]||'')
        .replace(/\u00A0/g,' ')     
        .replace(/\s+/g,' ')        
        .trim();
      return [t];
    });
    rng.setValues(vals);
  });
  SpreadsheetApp.getActive().toast('Normalized Owner/Assignee/Client/Priority text.', 'Styling', 3);
}

// Optional: set tab colors for all boards
function KNB_applyTabColorsAll(){
  const map = {
    [KNB_CFG.GID.REQUESTED]:   KNB_CFG.TAB_COLORS && KNB_CFG.TAB_COLORS.REQUESTED,
    [KNB_CFG.GID.INPROGRESS]:  KNB_CFG.TAB_COLORS && KNB_CFG.TAB_COLORS.INPROGRESS,
    [KNB_CFG.GID.FORAPPROVAL]: KNB_CFG.TAB_COLORS && KNB_CFG.TAB_COLORS.FORAPPROVAL,
    [KNB_CFG.GID.DONE]:        KNB_CFG.TAB_COLORS && KNB_CFG.TAB_COLORS.DONE
  };
  Object.entries(map).forEach(([gid,color])=>{
    const sh = KNB_sheetById_(Number(gid));
    if (sh && color) { try { sh.setTabColor(color); } catch(_){ } }
  });
}

function KNB_purgePinkHere(){
  const sh = SpreadsheetApp.getActiveSheet();
  sh.setConditionalFormatRules([]);
  try { (sh.getBandings() || []).forEach(b => b.remove()); } catch(_) {}
  const lastRow = Math.max(2, sh.getMaxRows());
  const lastCol = Math.max(1, sh.getMaxColumns());
  if (lastRow > 1) sh.getRange(2,1,lastRow-1,lastCol).setBackground(null);
  try { KNB_ensureStyleHere(); } catch(_) {}
  SpreadsheetApp.getActive().toast('Cleared CF/banding; reset body.', 'Styling', 4);
}

function KNB_makeBodyPlainHere(){
  const sh = SpreadsheetApp.getActiveSheet();

  // 0) Unmerge any body cells that could swallow pasted values
  const lastRow = Math.max(2, sh.getMaxRows());
  const lastCol = Math.max(1, sh.getMaxColumns());
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, lastCol).breakApart();

  // 1) Remove CF + banding
  sh.setConditionalFormatRules([]);
  try { (sh.getBandings() || []).forEach(b => b.remove()); } catch(_){}

  // 2) Reset body formatting portfolio-wide
  if (lastRow > 1) {
    const body = sh.getRange(2, 1, lastRow - 1, lastCol);
    body
      .setBackground('#ffffff')
      .setFontColor('#000000')
      .setFontWeight('normal')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('bottom')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    try { sh.autoResizeRows(2, lastRow - 1); } catch(_) {}
  }

  // 3) Reapply header + status chips
  // try { KNB_ensureStyleHere(); } catch(_){}
  // SpreadsheetApp.getActive().toast('Body normalized (unmerged + reset).', 'Styling', 4);
}

function KNB_clearFiltersHere(){
  const sh = SpreadsheetApp.getActiveSheet();
  // Remove standard filter
  try { if (sh.getFilter()) sh.getFilter().remove(); } catch(_) {}
  // Remove filter views scoped to this sheet
  try {
    SpreadsheetApp.getActive().getFilterViews()
      .filter(v => v.getRange().getSheet() === sh)
      .forEach(v => v.remove());
  } catch(_) {}
  // Reapply baseline styling
  try { KNB_ensureStyleHere(); } catch(_) {}
  SpreadsheetApp.getActive().toast('Filters cleared for this sheet.', 'Styling', 4);
}

// Reapply status DV + baseline style for the current sheet (cheap + local)
function KNB_touchStyleHere_() {
  const sh = SpreadsheetApp.getActiveSheet();
  const idx = KNB_headerIndex_(sh);
  const lastRow = Math.max(2, sh.getLastRow());

  // Status dropdown on existing rows
  const cStatus = idx[KNB_CFG.COL.STATUS];
  if (cStatus && lastRow > 1) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(KNB_CFG.STATUSES, true)
      .setAllowInvalid(false)
      .setHelpText('Choose a Status')
      .build();
    sh.getRange(2, cStatus, lastRow - 1, 1).setDataValidation(rule);
  }

  // Baseline header + chips + column fills
  try { KNB_ensureStyleHere(); } catch(_) {}
}
