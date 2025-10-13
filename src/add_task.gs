/* =========================
   ADD NEW TASK — FORM & WRITER
   - Status is left BLANK on create
   - Task Details saved to hidden "Task Details (HTML)" (no tooltip)
========================= */
function KNB_TASK_openNewTaskForm(){
  const tpl = HtmlService.createTemplateFromFile('views/NewTask');
  tpl.deptOptions = KNB_CFG.DEPARTMENTS || [];
  tpl.owners     = KNB_CFG.ASSIGNEES   || [];
  tpl.assignees  = KNB_CFG.ASSIGNEES   || [];   
  tpl.client  = KNB_CFG.CLIENTS   || [];   
  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(620).setHeight(720),
    'Add New Task'
  );
}

// Public wrapper so google.script.run can always find it
function KNB_TASK_add(p) {
  return KNB_TASK_add_(p);
}

function KNB_TASK_add_(p){
  // --- Destination sheet selection (robust) ---
  function step(label, fn){
    try { return fn(); }
    catch(e){ throw new Error('STEP['+label+'] → ' + (e && e.message || e)); }
  }

  const ss = SpreadsheetApp.getActive();
  const byGid   = ss.getSheets().find(x => x.getSheetId() === Number(KNB_CFG.GID.REQUESTED));
  const byName  = ss.getSheets().find(x => String(x.getName()).trim().toLowerCase() === 'requested');
  const needCols = [
    KNB_CFG.COL.DEPARTMENT, KNB_CFG.COL.ASSIGNEE, KNB_CFG.COL.CLIENT,
    KNB_CFG.COL.TASK, KNB_CFG.COL.PRIORITY, KNB_CFG.COL.CREATED,
    KNB_CFG.COL.STATUS, KNB_CFG.COL.DETAILS
  ];
  const hasAllHeaders = (sh)=>{
    const idx = KNB_headerIndex_(sh);
    return needCols.every(h => !!idx[h]);
  };
  const byHeaders = ss.getSheets().find(hasAllHeaders);
  const sh = byGid || byName;
  if (!sh) throw new Error('Requested board not found. Configure KNB_CFG.GID.REQUESTED or name a tab "Requested".');


  // Validate headers on chosen sheet
  const idx = KNB_headerIndex_(sh);

  // assertWritable_(sh);

  // function assertWritable_(sheet){
  //   // Block if whole sheet is protected
  //   var prot = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  //   if (prot && prot.isProtected()) {
  //     throw new Error('Sheet "'+sheet.getName()+'" is protected. Add yourself as an editor or remove the sheet protection.');
  //   }
  //   // Block if header row is protected (we otherwise insert/hide near headers)
  //   var header = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn()));
  //   var headerProtected = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
  //     .some(function(p){ return p.isProtected() && header.intersects(p.getRange()); });
  //   if (headerProtected) {
  //     throw new Error('Row 1 is protected on "'+sheet.getName()+'". Unprotect headers or add yourself as an editor to the protection.');
  //   }
  // }

  const missingCols = needCols.filter(h=>!idx[h]);
  if (missingCols.length) {
    const hdrs = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getDisplayValues()[0];
    throw new Error(
      'Destination sheet "'+sh.getName()+'" is missing required headers: ' + missingCols.join(', ') +
      '\n\nTip: Row 1 must contain these exact texts (any order):\n' + needCols.join(' | ') +
      '\n\nDetected headers:\n' + hdrs.join(' | ')
    );
  }

  // --- Make sure Priority DV includes the latest list (e.g., "Urgent") ---
  // (function ensurePriorityDV_(){
  //   const cPrio = idx[KNB_CFG.COL.PRIORITY] || 0;
  //   if (!cPrio) return;

  //   // Guard: detect protection on the Priority column body
  //   const body = sh.getRange(2, cPrio, Math.max(1, sh.getMaxRows() - 1), 1);
  //   const protectedBody = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
  //     .some(p => p.isProtected() && body.intersects(p.getRange()));

  //   if (protectedBody) {
  //     // Don’t attempt DV writes if protected; surface actionable guidance
  //     throw new Error('Priority column is protected. Remove range protection or add editors to that protection to allow dropdown refresh.');
  //   }

  //   const dv = SpreadsheetApp.newDataValidation()
  //     .requireValueInList(KNB_CFG.PRIORITIES || [], true)
  //     .setAllowInvalid(false)
  //     .setHelpText('Choose a Priority')
  //     .build();

  //   body.setDataValidation(dv);
  // })();


  // Robust date parsing
  function parseDate_(s){
    if(!s) return null;
    const t = String(s).trim();
    let m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);                 // yyyy-mm-dd
    if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);               // mm/dd/yyyy
    if (m) return new Date(Number(m[3]), Number(m[1])-1, Number(m[2]));
    m = t.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);                 // dd-mm-yyyy
    if (m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
    const d = new Date(t); return isNaN(d.getTime()) ? null : d;  // fallback
  }

  // Server-side safety defaults
  const created = parseDate_(p.created) || new Date();
  const start   = parseDate_(p.start);
  const due     = parseDate_(p.due);

  // Required payload checks (server-side)
  const missingReq = [];
  if (!String(p.department||'').trim()) missingReq.push('Department');
  if (!String(p.assignee||'').trim())   missingReq.push('Assignee');
  if (!String(p.client||'').trim())     missingReq.push('Client Name');
  if (!String(p.task||'').trim())       missingReq.push('Task Name');
  if (!String(p.prio||'').trim())       missingReq.push('Task Priority');
  if (!created)                         missingReq.push('Creation Date');
  if (missingReq.length) {
    throw new Error('Please complete required fields: ' + missingReq.join(', '));
  }

  // Append row
  const row = Math.max(sh.getLastRow()+1, 2);

  (function probeWrite_(){
    const targets = [
      KNB_CFG.COL.TASK,          
      KNB_CFG.COL.ASSIGNEE,
      KNB_CFG.COL.CLIENT,
      KNB_CFG.COL.DEPARTMENT
    ].filter(Boolean);

    const first = targets.find(h => idx[h]);
    if (!first) return; 

    const c = idx[first];
    const rg = sh.getRange(row, c);
    const val = rg.getValue();
    try {
      rg.setValue(val); 
    } catch (e) {
      throw new Error('You do not have edit rights on this sheet or the required columns. Underlying: ' + (e.message || e));
    }
  })();


  // assertTargetsWritable_(sh, idx, row);

  // function assertTargetsWritable_(sheet, idx, row){
  //   const targets = [
  //     KNB_CFG.COL.DEPARTMENT,
  //     KNB_CFG.COL.OWNER,
  //     KNB_CFG.COL.ASSIGNEE,
  //     KNB_CFG.COL.CLIENT,
  //     KNB_CFG.COL.TASK,
  //     KNB_CFG.COL.PRIORITY,
  //     KNB_CFG.COL.CREATED,
  //     KNB_CFG.COL.START,
  //     KNB_CFG.COL.DUE,
  //     KNB_CFG.COL.STATUS,
  //     'Task Details',            // visible icon cell
  //     'Task Details (HTML)'      // hidden storage
  //   ];

  //   const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
  //     .filter(p => p.isProtected());

  //   const blocked = [];
  //   for (const h of targets){
  //     const c = idx[h]; if (!c) continue;
  //     const rg = sheet.getRange(row, c, 1, 1);
  //     if (protections.some(p => rg.intersects(p.getRange()))) blocked.push(h);
  //   }

  //   if (blocked.length){
  //     throw new Error(
  //       'No edit rights on: ' + blocked.join(', ') +
  //       '. Update “Protected sheets and ranges” on "' + sheet.getName() + '" to include your account as an editor for those columns.'
  //     );
  //   }
  // }

  const set = (headerName, value) => {
    const c = idx[headerName];
    if (!c) return;
    try { sh.getRange(row, c).setValue(value); }
    catch (e) {
      const msg = String(e && e.message || e);
      if (/permission|forbidden|denied/i.test(msg)) {
        throw new Error('Write blocked at column "' + headerName + '". ' + msg);
      }
      throw e;
    }
  };



  set(KNB_CFG.COL.DEPARTMENT, p.department);
  set(KNB_CFG.COL.OWNER,      p.owner);
  set(KNB_CFG.COL.ASSIGNEE,   p.assignee);
  set(KNB_CFG.COL.CLIENT,     p.client);
  set(KNB_CFG.COL.TASK,       p.task);

  // Priority: write with a small retry if DV complained
  (function safeSetPriority_(){
    const cPrio = idx[KNB_CFG.COL.PRIORITY] || 0;
    if (!cPrio) return;
    const value = p.prio || 'Mid Prio';
    try {
      sh.getRange(row, cPrio).setValue(value);
    } catch (e){
      if (String(e).toLowerCase().includes('data validation')) {
        // Re-ensure DV and retry once
        const rows = Math.max(1, sh.getMaxRows() - 1);
        const dv = SpreadsheetApp.newDataValidation()
          .requireValueInList(KNB_CFG.PRIORITIES || [], true)
          .setAllowInvalid(false)
          .setHelpText('Choose a Priority')
          .build();
        sh.getRange(2, cPrio, rows, 1).setDataValidation(dv);
        sh.getRange(row, cPrio).setValue(value);
      } else {
        throw e;
      }
    }
  })();

  set(KNB_CFG.COL.CREATED,    created);
  if (idx[KNB_CFG.COL.START] && start) set(KNB_CFG.COL.START, start);  // optional
  if (idx[KNB_CFG.COL.DUE]   && due)   set(KNB_CFG.COL.DUE,   due);    // optional

  // Leave Status BLANK on create
  if (idx[KNB_CFG.COL.STATUS]) set(KNB_CFG.COL.STATUS, '');

  // --- Task Details: ensure hidden storage columns exist ---
  KNB_TASK_ensureDetailsColumns_(sh);

  // Re-read columns safely (case-insensitive) and assert existence
  const dCol = KNB_findHeaderColumn_(sh, 'Task Details');
  const hCol = KNB_findHeaderColumn_(sh, 'Task Details (HTML)');
  if (!dCol || !hCol) {
    throw new Error('Could not find "Task Details" / "Task Details (HTML)" after ensuring columns. Check row 1 headers.');
  }

  // Write HTML into hidden storage, set 📝 in visible cell, and ensure no tooltip
  const NOTE_LIMIT = 48000;
  let html = String(p.html||'').trim();
  if (html.length > NOTE_LIMIT) {
    throw new Error(`Task Details too large (${html.length} chars). Keep it under ${NOTE_LIMIT.toLocaleString()} characters.`);
  }

  try {
    sh.getRange(row, hCol).setValue(html);
  } catch(e){
    const m = String(e && e.message || e);
    if (/permission|forbidden|denied/i.test(m))
      throw new Error('Write blocked at "Task Details (HTML)". Update the protection to include you. ' + m);
    throw e;
  }

  try {
    const cell = sh.getRange(row, dCol);
    cell.setNote('');
    cell.setValue(html ? '📝' : '');
  } catch(e){
    const m = String(e && e.message || e);
    if (/permission|forbidden|denied/i.test(m))
      throw new Error('Write blocked at "Task Details". Update the protection to include you. ' + m);
    throw e;
  }


  try { KNB_touchStyleHere_(); } catch(_) {}

  try {
    KNB_applyOwnerDropdownHere();
    KNB_applyAssigneeDropdownHere();
    KNB_applyPriorityDropdownHere();
    KNB_applyClientDropdownHere();
  } catch(_) {}

  SpreadsheetApp.getActive().toast('Task added (Status blank).', 'Tasks', 4);
  return 'Task added to "'+sh.getName()+'" (row '+row+') — Status left blank.';
}


/* =========================
   Hidden storage columns helper
   - Ensures "Task Details (HTML)" and "Task Details (Draft)" exist (case-insensitive locate)
   - Hides them
========================= */
function KNB_TASK_ensureDetailsColumns_(sh){
  // Validate presence only; no structural mutations at runtime
  const dCol = KNB_findHeaderColumn_(sh, 'Task Details');
  const hCol = KNB_findHeaderColumn_(sh, 'Task Details (HTML)');
  const drCol = KNB_findHeaderColumn_(sh, 'Task Details (Draft)');
  if (!dCol || !hCol || !drCol) {
    throw new Error('Missing storage columns. An editor must add "Task Details (HTML)" and "Task Details (Draft)" immediately after "Task Details".');
  }
  // Optional: re-hide if someone exposed them
  try { sh.hideColumn(sh.getRange(1, hCol)); } catch(_){}
  try { sh.hideColumn(sh.getRange(1, drCol)); } catch(_){}
}


/* =========================
   Header finder (case-insensitive, trims)
========================= */
function KNB_findHeaderColumn_(sh, headerName){
  const n = Math.max(1, sh.getLastColumn());
  const row1 = sh.getRange(1,1,1,n).getDisplayValues()[0];
  const want = String(headerName||'').trim().toLowerCase();
  for (let i=0;i<row1.length;i++){
    const h = String(row1[i]||'').trim().toLowerCase();
    if (h === want) return i+1; // 1-based col
  }
  return 0; // not found
}
