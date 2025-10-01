/* =========================
   NOTIFIER (click-only, de-dupe 90s)
========================= */
function KNB_NTF_openNotifier(){
  const ui = HtmlService.createHtmlOutputFromFile('views/Notifier')
    .setWidth(480).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Send Discord Notification');
}

function KNB_NTF_getSelectedRowInfo() {
  const ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet();
  const sel = sh.getActiveRange(); if (!sel) return null;
  const row = sel.getRow(); if (row <= 1) return null;

  const idx = KNB_headerIndex_(sh);
  const get = (h)=> (idx[h] ? sh.getRange(row, idx[h]).getDisplayValue() : '');
  const taskName = KNB_sanitize_(String(get(KNB_CFG.COL.TASK) || 'Open row'));
  const assignee = String(get(KNB_CFG.COL.ASSIGNEE) || 'Unassigned').trim();
  const status   = String(get(KNB_CFG.COL.STATUS) || '').trim();
  return { row, taskName, assignee, status, defaultNote: KNB_NTF_CFG.DEFAULT_NOTE };
}

function KNB_NTF_manualSendSelected(statusVal, noteText) {
  const ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet();
  const sel = sh.getActiveRange(); if (!sel) return;
  const row = sel.getRow(); if (row <= 1) return;

  const idx = KNB_headerIndex_(sh);
  const get = (h)=> (idx[h] ? sh.getRange(row, idx[h]).getDisplayValue() : '');
  const assignee = String(get(KNB_CFG.COL.ASSIGNEE) || 'Unassigned').trim();
  const assigneeId = KNB_NTF_CFG.IDS[assignee] || null;

  const taskName = KNB_sanitize_(String(get(KNB_CFG.COL.TASK) || 'Open row'));
  const url = ss.getUrl().replace(/#.*$/,'') + '#gid=' + sh.getSheetId() + '&range=A' + row;

  const newStatus = String(statusVal || 'Requested').trim();
  const note = String(noteText || KNB_NTF_CFG.DEFAULT_NOTE);

  const key = KNB_NTF_CFG.P_SENT_PREFIX + sh.getSheetId() + '_' + row + '_' + newStatus;
  const props = PropertiesService.getDocumentProperties();
  const now = Date.now(), last = Number(props.getProperty(key) || 0);
  if (now - last < KNB_NTF_CFG.DEDUPE_MS) {
    SpreadsheetApp.getActive().toast('Sent recently. Try again later.', 'Notifier', 4);
    return;
  }
  props.setProperty(key, String(now));

  const msg = [
    assigneeId ? `<@${assigneeId}>` : '',
    'Task Update',
    `Assignee: ${assignee} | Row: ${row}`,
    `Task: [${taskName}](${url})`,
    `Status: ${newStatus}`,
    `Note: ${note}`
  ].filter(Boolean).join('\n');

  const payload = {
    content: msg,
    allowed_mentions: assigneeId ? { parse: [], users: [String(assigneeId)] } : { parse: [] }
  };

  // Prefer Script Property KNB_DISCORD_WEBHOOK; fallback to code constant
  const webhook = PropertiesService.getScriptProperties().getProperty('KNB_DISCORD_WEBHOOK') || KNB_NTF_CFG.WEBHOOK;

  const res = UrlFetchApp.fetch(webhook, {
    method:'post', contentType:'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions:true
  });
  const code = res.getResponseCode();
  SpreadsheetApp.getActive().toast(code<300 ? 'Notification sent.' : ('Discord error: ' + code), 'Notifier', 4);
}

function KNB_NTF_testWebhook() {
  const webhook = PropertiesService.getScriptProperties().getProperty('KNB_DISCORD_WEBHOOK') || KNB_NTF_CFG.WEBHOOK;
  const res = UrlFetchApp.fetch(webhook, {
    method:'post', contentType:'application/json',
    payload: JSON.stringify({ content:'🔔 Test ping from Google Sheets' }),
    muteHttpExceptions:true
  });
  SpreadsheetApp.getActive().toast('Webhook status: ' + res.getResponseCode(), 'Notifier', 4);
}


