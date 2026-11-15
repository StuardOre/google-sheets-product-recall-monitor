// Offline test of Code.gs with real CPSC API answers and a small fake of the Apps Script services
// (SpreadsheetApp, Session, LockService, MailApp, Utilities). No Google account or network needed.
//
// 1. Download the API answers once (any browser or curl):
//    recalls_2025.json  https://www.saferproducts.gov/RestWebServices/Recall?format=json&RecallDateStart=2025-01-01&RecallDateEnd=2025-12-31
//    recalls_2026.json  https://www.saferproducts.gov/RestWebServices/Recall?format=json&RecallDateStart=2026-01-01&RecallDateEnd=2026-11-15
//    lastpub.json       https://www.saferproducts.gov/RestWebServices/Recall?format=json&LastPublishDateStart=2026-11-01&LastPublishDateEnd=2026-11-15
// 2. Run: node test/test_code.js Code.gs recalls_2025.json recalls_2026.json lastpub.json
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const [codePath, f25, f26, fLast] = process.argv.slice(2);

// ---------- fake services
function a1ToRC(a1) {
  const m = a1.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d*))?$/);
  const col = (s) => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const r1 = Number(m[2]), c1 = col(m[1]);
  const c2 = m[3] ? col(m[3]) : c1;
  const r2 = m[3] ? (m[4] ? Number(m[4]) : 1000) : r1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}
class FakeSheet {
  constructor(name) { this.name = name; this.cells = {}; this.maxRows = 1000; this.filter = null; }
  key(r, c) { return r + ',' + c; }
  getName() { return this.name; }
  getLastRow() { let m = 0; for (const k in this.cells) { const [r] = k.split(',').map(Number); if (this.cells[k].v !== '' && r > m) m = r; } return m; }
  getLastColumn() { let m = 0; for (const k in this.cells) { const c = Number(k.split(',')[1]); if (this.cells[k].v !== '' && c > m) m = c; } return m; }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(after, n) { this.maxRows += n; }
  insertRowAfter(after) {
    const moved = {};
    for (const k in this.cells) { const [r, c] = k.split(',').map(Number); moved[(r > after ? r + 1 : r) + ',' + c] = this.cells[k]; }
    this.cells = moved; this.maxRows += 1;
  }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const [r, cc, nr, nc] = a1ToRC(a); return new FakeRange(this, r, cc, nr, nc); }
    return new FakeRange(this, a, b, c || 1, d || 1);
  }
  getFilter() { return this.filter; }
}
class FakeRange {
  constructor(sheet, r, c, nr, nc) {
    if (r + nr - 1 > sheet.maxRows) throw new Error('Range outside sheet: ' + sheet.name + ' rows ' + r + '+' + nr);
    Object.assign(this, { sheet, r, c, nr, nc });
  }
  cell(i, j) { return this.sheet.cells[this.sheet.key(this.r + i, this.c + j)] || { v: '', f: '' }; }
  getValues() { return [...Array(this.nr)].map((_, i) => [...Array(this.nc)].map((_, j) => { const x = this.cell(i, j); return x.f ? 'CPSC page' : x.v; })); }
  getFormulas() { return [...Array(this.nr)].map((_, i) => [...Array(this.nc)].map((_, j) => this.cell(i, j).f || '')); }
  setValues(values) {
    if (values.length !== this.nr || values[0].length !== this.nc) throw new Error('setValues size mismatch ' + values.length + 'x' + values[0].length + ' vs ' + this.nr + 'x' + this.nc);
    values.forEach((row, i) => row.forEach((v, j) => {
      const isF = typeof v === 'string' && v.startsWith('=');
      this.sheet.cells[this.sheet.key(this.r + i, this.c + j)] = { v: isF ? '' : v, f: isF ? v : '' };
    }));
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) delete this.sheet.cells[this.sheet.key(this.r + i, this.c + j)]; return this; }
  setNumberFormat() { return this; }
  setVerticalAlignment() { return this; }
  setWrap() { return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
  setFontWeight() { return this; }
  setHorizontalAlignment() { return this; }
  createFilter() { this.sheet.filter = { remove: () => { this.sheet.filter = null; } }; return this.sheet.filter; }
}
const sheets = {};
const ss = {
  getSheetByName: (n) => sheets[n] || null,
  toast: () => {},
  getUrl: () => 'https://docs.google.com/spreadsheets/d/TEST',
};
const sent = [];
const ctx = {
  console,
  SpreadsheetApp: { getActive: () => ss },
  Session: { getScriptTimeZone: () => 'Australia/Melbourne', getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  MailApp: { sendEmail: (o) => sent.push(o) },
  Utilities: {
    sleep: () => {},
    formatDate: (d, tz, fmt) => {
      const p = (n) => String(n).padStart(2, '0');
      const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      if (fmt === 'yyyy-MM-dd') return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
      if (fmt === 'd MMM yyyy') return d.getDate() + ' ' + mon[d.getMonth()] + ' ' + d.getFullYear();
      return d.toISOString();
    },
  },
  UrlFetchApp: { fetch: () => { throw new Error('no network in test'); } },
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(codePath, 'utf8'), ctx);

// ---------- seed the tabs the way setup() does (only the parts the sync reads)
function addSheet(name) { sheets[name] = new FakeSheet(name); return sheets[name]; }
const S = vm.runInContext('SETTING', ctx);
const settings = addSheet('Settings');
settings.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Notes']]);
settings.getRange(2, 1, 7, 2).setValues([
  [S.KEEP_FROM, new Date(2025, 0, 1)], [S.EMAIL, ''], [S.ALERTS, 'Yes'], [S.HOUR, 7], [S.LOOK_BACK, 3],
  [S.LAST_RUN, ''], [S.LAST_RESULT, ''],
]);
const wl = addSheet('Watchlist');
const DW = vm.runInContext('DEFAULT_WATCHLIST', ctx);
wl.getRange(1, 1, 1, 3).setValues([['Keyword', 'Area', 'Active']]);
wl.getRange(2, 1, DW.length, 3).setValues(DW);
const rules = addSheet('Rules');
const DR = vm.runInContext('DEFAULT_HAZARD_RULES', ctx), DM = vm.runInContext('DEFAULT_MARKETPLACES', ctx);
rules.getRange(1, 1, 1, 2).setValues([['Hazard category', 'Pattern']]);
rules.getRange(2, 1, DR.length, 2).setValues(DR);
rules.getRange(2, 4, DM.length, 2).setValues(DM);
const rec = addSheet('Recalls');
const H = vm.runInContext('HEADERS', ctx);
rec.getRange(1, 1, 1, H.length).setValues([H]);
const log = addSheet('Log');
log.getRange(1, 1, 1, 11).setValues([vm.runInContext('LOG_HEADERS', ctx)]);

// ---------- unit checks
const parseUnits = vm.runInContext('parseUnits_', ctx);
const cases = { 'About 265,000': 265000, '83': 83, '5.9 million': 5900000,
  'About 29,000 (In addition, about 4,500 were sold in Canada)': 29000, 'About 1.2 million': 1200000, '': null, 'None': null };
for (const [k, v] of Object.entries(cases)) {
  const got = parseUnits(k);
  console.log((got === v ? 'ok  ' : 'FAIL') + ' parseUnits(' + JSON.stringify(k) + ') = ' + got);
}
const kre = vm.runInContext('keywordRegex_', ctx);
const kwCases = [['toy', 'Plush Toys Recalled', true], ['toy', 'Toyota Floor Mats', false], ["children's", 'Children’s Costumes', true],
  ['baby carrier', 'Baby Carriers Recalled', true], ['crib', 'Cribbage Boards', false], ['kids', "Kids' Bikes", true]];
const matchWL = vm.runInContext('matchWatchlist_', ctx);
kwCases.forEach(([k, t, want]) => {
  const got = matchWL(t, [{ keyword: k, area: 'x', re: kre(k) }]).keywords.length > 0;
  console.log((got === want ? 'ok  ' : 'FAIL') + ' keyword ' + k + ' in "' + t + '" -> ' + got);
});

// ---------- run 1: backfill with 2025 + 2026
const sync = vm.runInContext('syncRecalls_', ctx);
const d25 = JSON.parse(fs.readFileSync(f25)), d26 = JSON.parse(fs.readFileSync(f26)), dl = JSON.parse(fs.readFileSync(fLast));
console.log('\nrun 1 (backfill):', sync(d25.concat(d26), 'Backfill'), '| emails sent:', sent.length, '| recipients:', sent.map((s) => s.to).join(','));
const rows = rec.getRange(2, 1, rec.getLastRow() - 1, H.length).getValues();
const forms = rec.getRange(2, 18, rec.getLastRow() - 1, 1).getFormulas();
console.log('rows:', rows.length, '| first:', rows[0][2], rows[0][3].slice(0, 60), '| last:', rows[rows.length - 1][2]);
console.log('links kept:', forms.filter((f) => f[0].startsWith('=HYPERLINK')).length);
const count = (idx, filter) => rows.filter(filter || (() => true)).reduce((m, r) => { (m[r[idx]] = (m[r[idx]] || 0) + 1); return m; }, {});
const isDate = (x) => Object.prototype.toString.call(x) === '[object Date]';
const y26 = (r) => isDate(r[2]) && r[2].getFullYear() === 2026;
console.log('hazard categories 2026:', count(8, y26));
console.log('watchlist matches 2026:', rows.filter((r) => y26(r) && r[14]).length, 'of', rows.filter(y26).length);
console.log('areas 2026:', count(15, (r) => y26(r) && r[15]));
console.log('amazon 2026:', rows.filter((r) => y26(r) && /Amazon/.test(r[12])).length);
console.log('units 2026:', rows.filter(y26).reduce((s, r) => s + (Number(r[6]) || 0), 0));
console.log('made in 2026 top:', Object.entries(count(11, y26)).sort((a, b) => b[1] - a[1]).slice(0, 6));
console.log('remedy 2026:', count(10, y26));
console.log('sample row:', JSON.stringify(rows.find((r) => r[14]).slice(0, 17)).slice(0, 900));
console.log('email subject:', sent[0] && sent[0].subject, '| html chars:', sent[0] && sent[0].htmlBody.length);

// ---------- run 2: same data again -> everything unchanged
console.log('\nrun 2 (same data):', sync(d25.concat(d26), 'Daily'), '| emails sent:', sent.length);
// ---------- run 3: the "last published" answer (includes 1998 updates) -> skipped/unchanged
console.log('run 3 (last published Nov 1-15):', sync(dl, 'Daily'));
// ---------- run 3b: a new matching record in a daily run -> 1 new, alert email
const fresh = JSON.parse(JSON.stringify(d26.find((r) => /Toy/.test(r.Title))));
fresh.RecallID = 99999; fresh.RecallNumber = 'TEST1'; fresh.RecallDate = '2026-11-15T00:00:00'; fresh.LastPublishDate = '2026-11-15T00:00:00';
const before3b = sent.length;
console.log('run 3b (new match):', sync([fresh], 'Daily'), '| new emails:', sent.length - before3b, '|', sent.length ? sent[sent.length - 1].subject : '');
// ---------- run 4: one record with a newer LastPublishDate -> 1 updated, first seen kept
const bumped = JSON.parse(JSON.stringify(d26[5]));
bumped.LastPublishDate = '2026-11-14T00:00:00'; bumped.Title = bumped.Title + ' (updated)';
const before = rec.getRange(2, 1, rec.getLastRow() - 1, H.length).getValues().find((r) => r[0] === bumped.RecallID);
console.log('run 4 (one updated):', sync([bumped], 'Daily'));
const after = rec.getRange(2, 1, rec.getLastRow() - 1, H.length).getValues().find((r) => r[0] === bumped.RecallID);
console.log('updated title ok:', after[3].endsWith('(updated)'), '| first seen kept:', String(after[19]) === String(before[19]), '| rows:', rec.getLastRow() - 1);
console.log('links kept after rewrite:', rec.getRange(2, 18, rec.getLastRow() - 1, 1).getFormulas().filter((f) => f[0].startsWith('=HYPERLINK')).length);
console.log('\nlog:');
log.getRange(2, 1, 4, 11).getValues().forEach((r) => console.log('  ', r.slice(1).join(' | ')));
console.log('settings last result:', settings.getRange(8, 2).getValues()[0][0]);
// The alert email of run 3b, to open in a browser.
const lastMail = sent[sent.length - 1];
fs.writeFileSync(path.join(__dirname, 'email_preview.html'), lastMail ? lastMail.htmlBody : '');
