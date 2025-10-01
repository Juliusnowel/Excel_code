/* =========================
   ADD NEW TASK — FORM & WRITER
   - Status is left BLANK on create
   - Task Details saved to hidden "Task Details (HTML)" (no tooltip)
========================= */
function KNB_TASK_openNewTaskForm(){
  const tpl = HtmlService.createTemplateFromFile('views/NewTask');
  tpl.deptOptions = KNB_CFG.DEPARTMENTS || [];
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
  const sh = byGid || byName || byHeaders || ss.getActiveSheet();

  // Validate headers on chosen sheet
  const idx = KNB_headerIndex_(sh);
  const missingCols = needCols.filter(h=>!idx[h]);
  if (missingCols.length) {
    const hdrs = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getDisplayValues()[0];
    throw new Error(
      'Destination sheet "'+sh.getName()+'" is missing required headers: ' + missingCols.join(', ') +
      '\n\nTip: Row 1 must contain these exact texts (any order):\n' + needCols.join(' | ') +
      '\n\nDetected headers:\n' + hdrs.join(' | ')
    );
  }

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
  const set = (name, val)=>{ const c=idx[name]; if(c) sh.getRange(row,c).setValue(val); };

  set(KNB_CFG.COL.DEPARTMENT, p.department);
  set(KNB_CFG.COL.ASSIGNEE,   p.assignee);
  set(KNB_CFG.COL.CLIENT,     p.client);
  set(KNB_CFG.COL.TASK,       p.task);
  set(KNB_CFG.COL.PRIORITY,   p.prio || 'Mid Prio');
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

  sh.getRange(row, hCol).setValue(html);      
  const cell = sh.getRange(row, dCol);
  cell.setNote('');                          
  cell.setValue(html ? '📝' : '');

  try { KNB_touchStyleHere_(); } catch(_) {}

  try {
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
  // Find "Task Details" (must exist)
  const dCol = KNB_findHeaderColumn_(sh, 'Task Details');
  if (!dCol) throw new Error('Missing header "Task Details" on this sheet.');

  // Ensure "Task Details (HTML)"
  let hCol = KNB_findHeaderColumn_(sh, 'Task Details (HTML)');
  if (!hCol) {
    sh.insertColumnAfter(dCol);
    sh.getRange(1, dCol + 1).setValue('Task Details (HTML)');
    hCol = dCol + 1;
  }

  // Ensure "Task Details (Draft)" (after HTML)
  let drCol = KNB_findHeaderColumn_(sh, 'Task Details (Draft)');
  if (!drCol) {
    // hCol might have moved; re-find it safely
    hCol = KNB_findHeaderColumn_(sh, 'Task Details (HTML)');
    sh.insertColumnAfter(hCol);
    sh.getRange(1, hCol + 1).setValue('Task Details (Draft)');
    drCol = hCol + 1;
  }

  // Hide storage columns
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
