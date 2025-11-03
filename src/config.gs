/* =========================
   CONFIG
========================= */
const KNB_CFG = {
  // Boards (update GIDs as needed)
  GID: {
    REQUESTED:   551438246,
    INPROGRESS:  1673786856,
    FORAPPROVAL: 1097345134,
    DONE:        232929486
  },
  // Canonical header names (MUST match your headers row text)
  COL: {
    ROWID:       'Row ID',
    DEPARTMENT:  'Department',
    OWNER:       'Owner',      
    ASSIGNEE:    'Assignee',
    CLIENT:      'Client Name',
    TASK:        'Task Name',
    DETAILS:     'Task Details',          // visible cell (📝 when saved)
    PRIORITY:    'Task Priority',
    CREATED:     'Creation Date',
    START:       'Start Date',
    DAYCOUNT:    'Day Count',
    END:         'End Date',
    DUE:         'Due Date',              // optional; used by Add New Task / Private views
    DELIVERABLE: 'Deliverable',
    SCREENSHOT:  'Screenshot',
    STATUS:      'Status',
    FREEZE:      'For Approval Date'
    // LATE:        'Late or not?'
  },
  // Status list and routing map
  STATUSES: ['Requested','In Progress','For Approval','For Revision','Blocked','On Hold','Done'],
  ROUTE: {
    'Requested':    551438246,
    'In Progress':  1673786856,
    'For Approval': 1097345134,
    'For Revision': 551438246,   // bounce back to Requested
    'Done':         232929486
    // 'Blocked' / 'On Hold' stay in place
  },
  // Mover UX
  RATE_LIMIT_MS: 3000,
  REVERT_FALLBACK: 'Requested',
  // Column shading (by header text)
  COLUMN_COLORS: {
    'Department':    '#d0eafc',
    'Owner':      '#d0eafc',
    'Task Name':     '#dfd3fa',
    'Creation Date': '#fcebd6',
    'Deliverable':   '#f5c5d7',   
    'Screenshot':    '#F8F8AD'
  },

  // Add this to KNB_CFG (optional seed)
  ASSIGNEES: ['Jane','Isay','Jake','Julius','Prince','Ellis','Ivan','Ferdinand','Andrew','Ray'],

  // Assignee chips (exact cell text)
  ASSIGNEE_COLORS: {
    'Julius':    '#00D9FF',
    'Jane':      '#FF0080',
    'Ivan':      '#FFC800',
    'Isay':      '#FF0000',
    'Jake':      '#88FF00',
    'Prince':    '#B700FF',
    'Ellis':     '#4C00FF',
    'Ferdinand': '#00FF88',
    'Andrew':    '#FF9900',
    'Ray':       '#00796B'
  },
  // Status color chips (used only by the optional styling helper)
  STATUS_COLORS: {
    'Requested':    '#3FA6FF',
    'In Progress':  '#1C7CFF',
    'For Approval': '#FFB300',
    'For Revision': '#8E24AA',
    'Blocked':      '#E53935',
    'On Hold':      '#757575',
    'Done':         '#2ECC71'
  },
  // Status color chips (used only by the optional styling helper)
  PRIORITY_COLORS: {
    'Adhoc Task': '#7EF0FF',
    'Low Prio':   '#7EFDA4',
    'Mid Prio':   '#FFB300',
    'High Prio':  '#FF3535',
    'Urgent':     '#FF3535'
  },
  // Status color chips (used only by the optional styling helper)
  CLIENT_COLORS: {
    'Allinclusive': '#3FA6FF',
    'Creceri':      '#1C7CFF',
    'Wix Media':    '#FFB300',
    'Ilegiants':    '#E53935',
    'Yzenshun':     '#757575',
    'Windows Live': '#2ECC71',
    'Vite SEO':    '#8E24AA'
  },
  HEADER_BG: '#0B1221',
  HEADER_FG: '#FFFFFF',
  // Dropdown choices (used by dialog + optional sheet validation helper)
  DEPARTMENTS: ['SEO Department','QA Department','Dev Department','Design Department'],
  PRIORITIES: ['Adhoc Task','Low Prio','Mid Prio','High Prio', 'Urgent'],
  CLIENTS: [
    'Allinclusive','Creceri','Wix Media','Ilegiants','Yzenshun','Windows Live','Vite SEO'
  ]
};

// Notifier config (click-only)
const KNB_NTF_CFG = {
  // Fallback webhook. Prefer setting Script Property 'KNB_DISCORD_WEBHOOK'
  WEBHOOK: 'https://discord.com/api/webhooks/1389807409274425466/BnwDqhrWOJ-GmBDL08exqk1IMwVL6fvfspIrzNBeG04Ss7bCet20GWls2iRVc80kZOFV',
  DEFAULT_NOTE: 'Please check the task details and let me know if you have any questions.',
  IDS: {
    'Jane':'1206775761772486677','Isay':'1387359431913902092','Jake':'921388311988273223',
    'Julius':'1420979154441998456','Prince':'1212022069215371264','Ellis':'1409729146355187763',
    'Ivan':'945666323646656512','Ferdinand':'353470938769260564','Andrew':'1389065637820760196',
    'Ray':'1396751676546879628'
  },
  P_SENT_PREFIX: 'KNB_NTF_SENT_',
  DEDUPE_MS: 90 * 1000
};

// Private views config
// const KNB_PVX_DEST_IDS = {
//   'Jane':      '1Te0Ul4MOglyi44j4eJwsUcjEJLz9krDXKBC0gHWdOeo',
//   'Isay':      '1EwMLS30HCscGLUhC-7nV2-8j235-V2O8ULdi2BAtHfE',
//   'Jake':      '1mtPP1CDgxXbFh7ruYr5iGZcBH82F2GXrTl0VSzlw47E',
//   'Julius':    '1pj7Oz1rCp1KEDpcppHd4vaqRbyJ_QWn1vViJs1M_xm8',
//   'Prince':    '1kRAx8YvgDvXrJCODxgOFdv4ciB6GGItDwE7lQjmKZ-E',
//   'Ellis':     '1p3CzP1Gzi57Q-OcKtJOgs989DTQhKQGoQzy_erXnhsg',
//   'Ivan':      '1lAUQZBnTpyKLI2-N3n4xPpdpZoWUcyldXiisEo0nRkI',
//   'Ferdinand': '1G8SsAQnIZfYO2yAau90I6KihNwfXLO_-kjDr7ibIhiI',
//   'Ray':       ''
// };

// optional roll-ups
const KNB_PVX_GROUPS = {
  // 'Andrew (Mgr)': ['Jane','Isay','Jake','Julius','Prince','Ellis','Ivan','Ferdinand', 'Ray']
};

// columns to export to private Overview
const KNB_PVX_COLS = [
  'Status','Task Priority','Client Name','Task Name','Task Details','Owner',
  'Assignee','Start Date','Due Date','End Date','Day Count'
];