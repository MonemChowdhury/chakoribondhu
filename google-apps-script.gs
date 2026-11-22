// ChakoriBondhu survey → Google Sheet
// 1. Put your Sheet ID below (the long part of the Sheet URL between /d/ and /edit).
// 2. New sheet only: run setup() once (pick "setup" in the function menu → Run).
//    WARNING: setup() deletes every response. On a sheet that already has responses,
//    run upgrade() instead: it adds the new columns and rebuilds only the Summary tab.
// 3. Deploy → Manage deployments → pencil icon → Version: New version → Deploy (URL stays the same).

const SHEET_ID = 'PASTE_SHEET_ID_HERE';
const TZ = 'Asia/Dhaka';

const COLUMNS = [
  ['Submitted (BD time)', 150],
  ['Contact',             140],
  ['Education',           150],
  ['Looking for',         200],
  ['Biggest problem',     200],
  ['Wants',               260],
  ['Plan / would pay',    150],
  ['District',            110],
  ['Birth year',           80],
  ['Source',              180],
  ['Paid before for',     200],   // added in v3 — new columns go at the end so old rows stay aligned
  ['Would pay for',       240],
];

const LABELS = {
  jobtype: { govt: 'Govt / BCS', bank: 'Bank', private: 'Private', ngo: 'NGO', abroad: 'Abroad' },
  pain: {
    'missing-circulars': 'Miss circulars / deadlines',
    'cv': 'CV & cover letter',
    'exam-prep': 'Exam prep',
    'interview': 'Interview nerves',
    'experience': 'No experience',
  },
  want: {
    'alerts': 'Circular alerts',
    'reminders': 'Deadline reminders',
    'english-cv': 'English CV',
    'cover-letter': 'Cover letters',
    'interview-practice': 'Interview practice',
    'after-apply': 'After-apply reminders',
  },
  paidbefore: {
    'coaching': 'Coaching',
    'prep-app': 'Prep app / course',
    'job-site': 'Paid job-site plan',
    'cv-service': 'CV service',
    'never': 'Never paid',
  },
  paywant: {
    'eligibility': 'Eligibility check',
    'reminders': 'Deadline/fee/exam reminders',
    'cv': 'English CV (once)',
    'cover-letter': 'Cover letter per job',
    'none': 'Nothing, free only',
  },
  price: {
    'free-only': 'Free only',
    '99-1m': 'Tk 99 / month',
    '249-3m': 'Tk 249 / 3 months',
    '449-6m': 'Tk 449 / 6 months',
    // older answers, before the plan-based question
    '30': 'Tk 30', '50': 'Tk 50', '100': 'Tk 100', '200+': 'Tk 200+',
  },
};

function label_(field, raw) {
  if (!raw) return '';
  const map = LABELS[field];
  return String(raw).split(',').map(s => s.trim()).filter(Boolean)
    .map(v => (map && map[v]) || v).join(', ');
}

function book_() { return SpreadsheetApp.openById(SHEET_ID); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = book_();
    const sh = ss.getSheetByName('Responses') || ss.getSheets()[0];
    sh.appendRow([
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'),
      "'" + String(d.contact || '').trim(),          // keep the leading 0
      d.education || '',
      label_('jobtype', d.jobtype),
      label_('pain', d.pain),
      label_('want', d.want),
      label_('price', d.price),
      d.district || '',
      d.birthyear || '',
      d.source || '',
      label_('paidbefore', d.paidbefore),
      label_('paywant', d.paywant),
    ]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// Run once. Rebuilds the Responses tab (clears old rows) and creates a Summary tab.
function setup() {
  const ss = book_();
  let sh = ss.getSheetByName('Responses');
  if (!sh) { sh = ss.getSheets()[0]; sh.setName('Responses'); }
  sh.clear();
  sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => c[0])])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#0f6b4f')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);
  sh.setFrozenRows(1);
  COLUMNS.forEach((c, i) => sh.setColumnWidth(i + 1, c[1]));
  sh.getRange('B:B').setNumberFormat('@');
  sh.getRange(2, 1, sh.getMaxRows() - 1, COLUMNS.length).setWrap(true).setVerticalAlignment('top');
  if (sh.getMaxColumns() > COLUMNS.length) sh.deleteColumns(COLUMNS.length + 1, sh.getMaxColumns() - COLUMNS.length);
  sh.getBandings().forEach(b => b.remove());
  sh.getRange(1, 1, sh.getMaxRows(), COLUMNS.length).applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, true, false);

  buildSummary_(ss);
  ss.setActiveSheet(sh);
}

// Rebuilds only the Summary tab. Never touches responses.
function buildSummary_(ss) {
  let sum = ss.getSheetByName('Summary');
  if (!sum) sum = ss.insertSheet('Summary');
  sum.clear();
  const rows = [['ChakoriBondhu survey summary', '']];
  rows.push(['Total responses', '=COUNTA(Responses!B2:B)']);
  rows.push(['', '']);
  const block = (title, col, labels, exact) => {
    rows.push([title, 'Count']);
    Object.values(labels).forEach(l => rows.push([l,
      exact ? `=COUNTIF(Responses!${col}2:${col},"${l}")` : `=COUNTIF(Responses!${col}2:${col},"*${l}*")`]));
    rows.push(['', '']);
  };
  block('Plan chosen', 'G', ['Free only', 'Tk 99 / month', 'Tk 249 / 3 months', 'Tk 449 / 6 months'], true);
  block('Older price answers', 'G', ['Tk 30', 'Tk 50', 'Tk 100', 'Tk 200+'], true);
  block('Biggest problem', 'E', LABELS.pain, true);
  block('Wants', 'F', LABELS.want, false);
  block('Looking for', 'D', LABELS.jobtype, false);
  block('Paid before for', 'K', LABELS.paidbefore, false);
  block('Would pay for', 'L', LABELS.paywant, false);
  sum.getRange(1, 1, rows.length, 2).setValues(rows);
  sum.getRange('A1').setFontSize(14).setFontWeight('bold');
  rows.forEach((r, i) => { if (r[1] === 'Count') sum.getRange(i + 1, 1, 1, 2).setFontWeight('bold').setBackground('#e6f2ec'); });
  sum.getRange('A2:B2').setFontWeight('bold');
  sum.setColumnWidth(1, 240); sum.setColumnWidth(2, 80);
}

// Safe for a sheet that already has responses: adds any missing column headers
// at the end of the Responses tab and rebuilds the Summary tab. Deletes nothing.
function upgrade() {
  const ss = book_();
  const sh = ss.getSheetByName('Responses');
  if (!sh) throw new Error('No Responses tab: run setup() on a new sheet instead.');
  const need = COLUMNS.length;
  if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  COLUMNS.forEach((c, i) => {
    const cell = sh.getRange(1, i + 1);
    if (!cell.getValue()) {
      cell.setValue(c[0]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#0f6b4f');
      sh.setColumnWidth(i + 1, c[1]);
    }
  });
  buildSummary_(ss);
}
