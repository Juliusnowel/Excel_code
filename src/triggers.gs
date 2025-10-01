/* =========================
   TRIGGERS (installable onEdit)
========================= */
function KNB_installOnEditTrigger(){
  const ss = SpreadsheetApp.getActive();
  // Remove existing
  ScriptApp.getProjectTriggers().forEach(t=>{
    const fn = t.getHandlerFunction && t.getHandlerFunction();
    if (fn === 'KNB_onEdit_') ScriptApp.deleteTrigger(t);
  });
  // IMPORTANT: pass spreadsheet ID (string), not the object
  ScriptApp.newTrigger('KNB_onEdit_').forSpreadsheet(ss.getId()).onEdit().create();
  ss.toast('onEdit trigger installed. You may need to authorize once.', 'Kanban Suite', 4);
}
function KNB_removeOnEditTrigger(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    const fn = t.getHandlerFunction && t.getHandlerFunction();
    if (fn === 'KNB_onEdit_') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getActive().toast('onEdit trigger removed.', 'Kanban Suite', 4);
}

// Single installable onEdit handler: mover + task editor auto-open if Task Details touched
function KNB_onEdit_(e){
  try{
    if(!e || !e.range) return;
    // HTML editor auto-open when touching Task Details cell
    try { KNB_MTE_onEditOpen_(e); } catch(_){}
    // Mover reacts when Status is changed
    KNB_moverOnEdit_(e);
  } catch(err){
    console.error(err);
  }
}

// Optional: open editor on click of Task Details (no trigger install needed)
function onSelectionChange(e){ try{ KNB_MTE_onSelectionChange_(e); }catch(_){} }
