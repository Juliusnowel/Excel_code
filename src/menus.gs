/* =========================
   MENUS
========================= */
function KNB_RVN_openForActiveRow(){ KNB_RTE_openEditorForActiveRow('revision'); }
function KNB_RVN_backfillIcons(){ KNB_RTE_backfillIcons('revision'); }
function KNB_installRTETriggers(){
  // Clean existing installables
  ScriptApp.getProjectTriggers().forEach(t=>{
    const h = t.getHandlerFunction && t.getHandlerFunction();
    if (h === 'KNB_MTE_onEditOpen_' || h === 'KNB_RTE_onEditOpen_' || h === 'KNB_RTE_onSelectionChange_'){
      ScriptApp.deleteTrigger(t);
    }
  });
  // Install new
  ScriptApp.newTrigger('KNB_RTE_onEditOpen_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  // Optional but nice UX: click-to-open on select
  ScriptApp.newTrigger('KNB_RTE_onSelectionChange_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onSelectionChange()
    .create();
}

function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Tools') 
    .addSubMenu(
      ui.createMenu('Mover')
        .addItem('Install onEdit trigger', 'KNB_installOnEditTrigger')
        .addItem('Remove onEdit trigger',  'KNB_removeOnEditTrigger')
        .addSeparator()
        // .addItem('Reconcile (Gated)', 'KNB_reconcileGated')
        .addItem('Reconcile (Strict)', 'KNB_reconcileStrict')
        .addItem('Force Move (Selected Rows)', 'KNB_forceMoveSelected')
        .addSeparator()
        // .addItem('Apply Status Dropdowns (All Boards)', 'KNB_applyStatusDropdownsAll')
        // .addItem('Ensure Baseline Style (This Sheet)', 'KNB_ensureStyleHere')
        .addItem('Apply Tab Colors (All Boards)', 'KNB_applyTabColorsAll')
    )
    .addSubMenu(
      ui.createMenu('Tasks')
        .addItem('Add New Task…', 'KNB_TASK_openNewTaskForm')
        .addItem('Open HTML Editor (Active Row)…', 'KNB_MTE_openEditorForActiveRow')
        // .addItem('Backfill 📝 Icons', 'KNB_MTE_backfillIcons')
        .addItem('Open Revision Notes Editor (Active Row)…', 'KNB_RVN_openForActiveRow')
        .addItem('Backfill 🧾 Icons (Revision Notes)', 'KNB_RVN_backfillIcons')
        .addSeparator()
        // .addItem('Apply Department Dropdown (This Sheet)', 'KNB_applyDepartmentDropdownHere')
    )
    .addSubMenu(
      ui.createMenu('Notifier')
        .addItem('Open notifier (selected row)…', 'KNB_NTF_openNotifier')
        .addItem('Test webhook', 'KNB_NTF_testWebhook')
    )
    // .addSubMenu(
    //   ui.createMenu('Private Views')
    //     .addItem('Open My Private', 'KNB_PVX_openMyPrivate')
    //     .addItem('Refresh My Private', 'KNB_PVX_refreshMyPrivate')
    //     .addSeparator()
    //     .addItem('Publish All (Now)', 'KNB_PVX_publishAll')
    //     .addSeparator()
    //     .addItem('Install Hourly Refresh', 'KNB_PVX_installHourlyRefresh')
    //     .addItem('Remove Hourly Refresh', 'KNB_PVX_removeHourlyRefresh')
    // )
    .addSubMenu(
      ui.createMenu('Debug')
        .addItem('Diagnose Add Task Environment', 'KNB_TASK_diagnoseEnvironment_')
        .addItem('List Triggers to Logs', 'KNB_debugTriggers')
        .addItem('Reset Body Background (This Sheet)', 'KNB_resetSheetBodyBackgroundHere')
        // .addItem('Purge Pink (This Sheet)', 'KNB_purgePinkHere')
        // .addItem('Clear Filters (This Sheet)', 'KNB_clearFiltersHere') 
        // .addItem('Plain Body (This Sheet)', 'KNB_makeBodyPlainHere') 
        // .addItem('Plain Body (All Boards)', 'KNB_makeBodyPlain_AllBoards')
        // .addItem('Jump to Bottom (This Sheet)', 'KNB_jumpToLastRowHere')
        .addItem('Move Active Row → Done (DEBUG)', 'KNB_DEBUG_moveActiveRowToDone')
        // .addItem('Hide Task Details storage (All Boards)', 'KNB_hideTaskDetailsStorage_AllBoards')
        // .addItem('Ensure + Hide Task Details storage (All Boards)', 'KNB_TASK_ensureDetailsColumns_AllBoards')
        // .addItem('Apply Day Count (This Sheet)', 'KNB_applyDayCountHere')
        // .addItem('Apply Day Count (All Boards)', 'KNB_applyDayCount_AllBoards')
        // .addItem('Reset Day Count (This Sheet)', 'KNB_resetDayCountHere')
        .addItem('Reset Day Count (All Boards)', 'KNB_resetDayCount_AllBoards')
        // .addItem('Back Fill End Date', 'KNB_backfillEndDateForDoneHere')
        // .addItem('Install Owner Column (All Boards)', 'KNB_installOwnerColumnAll')
        // .addItem('Apply Owner Dropdown (All Boards)', 'KNB_applyOwnerDropdown_AllBoards')
        // .addItem('Refresh Priorities (All Boards)', 'KNB_refreshPriority_AllBoards')
        // .addItem('For Approval Date Column','KNB_installForApprovalDateColumn_AllBoards')
      )
    .addToUi();
}