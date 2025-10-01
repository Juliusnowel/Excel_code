/* =========================
   MENUS
========================= */
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Tools') // renamed from "Kanban Suite"
    .addSubMenu(
      ui.createMenu('Mover')
        .addItem('Install onEdit trigger', 'KNB_installOnEditTrigger')
        .addItem('Remove onEdit trigger',  'KNB_removeOnEditTrigger')
        .addSeparator()
        .addItem('Reconcile (Gated)', 'KNB_reconcileGated')
        .addItem('Reconcile (Strict)', 'KNB_reconcileStrict')
        .addItem('Force Move (Selected Rows)', 'KNB_forceMoveSelected')
        .addSeparator()
        .addItem('Apply Status Dropdowns (All Boards)', 'KNB_applyStatusDropdownsAll')
        .addItem('Ensure Baseline Style (This Sheet)', 'KNB_ensureStyleHere')
        .addItem('Apply Tab Colors (All Boards)', 'KNB_applyTabColorsAll')
    )
    .addSubMenu(
      ui.createMenu('Tasks')
        .addItem('Add New Task…', 'KNB_TASK_openNewTaskForm')
        .addItem('Open HTML Editor (Active Row)…', 'KNB_MTE_openEditorForActiveRow')
        .addItem('Backfill 📝 Icons', 'KNB_MTE_backfillIcons')
        .addSeparator()
        .addItem('Apply Department Dropdown (This Sheet)', 'KNB_applyDepartmentDropdownHere')
    )
    .addSubMenu(
      ui.createMenu('Notifier')
        .addItem('Open notifier (selected row)…', 'KNB_NTF_openNotifier')
        .addItem('Test webhook', 'KNB_NTF_testWebhook')
    )
    .addSubMenu(
      ui.createMenu('Private Views')
        .addItem('Open My Private', 'KNB_PVX_openMyPrivate')
        .addItem('Refresh My Private', 'KNB_PVX_refreshMyPrivate')
        .addSeparator()
        .addItem('Publish All (Now)', 'KNB_PVX_publishAll')
        .addSeparator()
        .addItem('Install Hourly Refresh', 'KNB_PVX_installHourlyRefresh')
        .addItem('Remove Hourly Refresh', 'KNB_PVX_removeHourlyRefresh')
    )
    .addSubMenu(
      ui.createMenu('Debug')
        .addItem('Diagnose Add Task Environment', 'KNB_TASK_diagnoseEnvironment_')
        .addItem('List Triggers to Logs', 'KNB_debugTriggers')
        .addItem('Reset Body Background (This Sheet)', 'KNB_resetSheetBodyBackgroundHere')
        .addItem('Purge Pink (This Sheet)', 'KNB_purgePinkHere')
        .addItem('Clear Filters (This Sheet)', 'KNB_clearFiltersHere') 
        .addItem('Plain Body (This Sheet)', 'KNB_makeBodyPlainHere') 
        .addItem('Plain Body (All Boards)', 'KNB_makeBodyPlain_AllBoards')
        .addItem('Jump to Bottom (This Sheet)', 'KNB_jumpToLastRowHere')
        .addItem('Move Active Row → Done (DEBUG)', 'KNB_DEBUG_moveActiveRowToDone')
        .addItem('Hide Task Details storage (All Boards)', 'KNB_hideTaskDetailsStorage_AllBoards')
        .addItem('Ensure + Hide Task Details storage (All Boards)', 'KNB_TASK_ensureDetailsColumns_AllBoards')
      )
    .addToUi();
}