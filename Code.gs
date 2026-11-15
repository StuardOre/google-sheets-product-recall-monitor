/**
 * Product Recall Monitor (Google Sheets + Apps Script)
 *
 * Pulls U.S. Consumer Product Safety Commission (CPSC) recalls from the public
 * Recalls API, keeps one row per recall, flags the recalls that match your
 * watchlist and can email you the new matches. Runs every day on its own.
 *
 * Data source: CPSC Recalls API, https://www.saferproducts.gov/RestWebServices/Recall
 * (U.S. government data, public domain).
 *
 * Start: open a blank Google Sheet, paste this file in Extensions > Apps Script,
 * save, reload the sheet and use the menu Recall Monitor > Set up the sheet.
 *
 * @OnlyCurrentDoc
 */

const API_URL = 'https://www.saferproducts.gov/RestWebServices/Recall';

const SHEET = {
  DASHBOARD: 'Dashboard',
  RECALLS: 'Recalls',
  WATCHLIST: 'Watchlist',
  RULES: 'Rules',
  SETTINGS: 'Settings',
  LOG: 'Log',
  GUIDE: 'How to use',
};

const HEADERS = [
  'Recall ID', 'Recall number', 'Recall date', 'Title', 'Product', 'Product type',
  'Units', 'Units (as published)', 'Hazard category', 'Hazard', 'Remedy', 'Made in',
  'Sold online at', 'Sold at (as published)', 'Watchlist match', 'Watchlist area',
  'Injuries and incidents', 'CPSC page', 'Last published', 'First seen', 'Last updated',
];

// 1-based column numbers in the Recalls tab.
const COL = {
  ID: 1, NUMBER: 2, DATE: 3, TITLE: 4, PRODUCT: 5, TYPE: 6, UNITS: 7, UNITS_TEXT: 8,
  HAZARD_CATEGORY: 9, HAZARD: 10, REMEDY: 11, MADE_IN: 12, ONLINE: 13, SOLD_AT: 14,
  MATCH: 15, AREA: 16, INJURIES: 17, URL: 18, LAST_PUBLISHED: 19, FIRST_SEEN: 20,
  LAST_UPDATED: 21,
};

const SETTING = {
  KEEP_FROM: 'Keep recalls dated from',
  EMAIL: 'Alert email',
  ALERTS: 'Send email alerts',
  HOUR: 'Daily run hour (0-23)',
  LOOK_BACK: 'Days to look back for updates',
  LAST_RUN: 'Last run',
  LAST_RESULT: 'Last run result',
};

const LOG_HEADERS = [
  'Run time', 'Mode', 'Recalls received', 'New', 'Updated', 'Unchanged',
  'Skipped (older than start date)', 'New watchlist matches', 'Email sent', 'Seconds', 'Result',
];

const DEFAULT_WATCHLIST = [
  ['crib', 'Nursery and sleep', 'Yes'],
  ['bassinet', 'Nursery and sleep', 'Yes'],
  ['play yard', 'Nursery and sleep', 'Yes'],
  ['nursery', 'Nursery and sleep', 'Yes'],
  ['high chair', 'Feeding and bath', 'Yes'],
  ['bath seat', 'Feeding and bath', 'Yes'],
  ['pacifier', 'Feeding and bath', 'Yes'],
  ['teether', 'Feeding and bath', 'Yes'],
  ['stroller', 'On the go', 'Yes'],
  ['car seat', 'On the go', 'Yes'],
  ['baby carrier', 'On the go', 'Yes'],
  ['toy', 'Toys', 'Yes'],
  ['baby', 'Baby and kids (general)', 'Yes'],
  ['infant', 'Baby and kids (general)', 'Yes'],
  ['toddler', 'Baby and kids (general)', 'Yes'],
  ["children's", 'Baby and kids (general)', 'Yes'],
  ['kids', 'Baby and kids (general)', 'Yes'],
];

// First rule whose pattern matches the hazard text wins. Patterns are regular expressions.
const DEFAULT_HAZARD_RULES = [
  ['Poisoning / chemical', 'levels of lead|contains? lead|lead (content|paint)|lead poisoning|phthalate|cadmium|carbon monoxide|methylene chloride|benzene|child-resistant|poison prevention packaging'],
  ['Choking / ingestion', 'chok|ingest|swallow|aspirat|magnet|button cell|coin cell|button batter|coin batter|water bead'],
  ['Drowning', 'drown'],
  ['Strangulation / suffocation', 'strangul|suffocat|entangle|asphyx'],
  ['Tip-over / entrapment', 'tip-over|tip over|tipover|entrapment'],
  ['Fire / burn', 'fire|burn|overheat|ignit|flame|thermal|melt'],
  ['Electric shock', 'shock|electrocut'],
  ['Fall', '\\bfall|collapse'],
  ['Injury / crash', 'crash|loss of control|lacerat|\\bcut|puncture|impact|struck|crush|pinch|amputat|injur'],
  ['Health / bacteria', 'bacteri|mold|microb|infection|fung|irritat'],
];

const DEFAULT_MARKETPLACES = [
  ['Amazon', 'amazon'],
  ['Walmart.com', 'walmart\\.com'],
  ['Temu', 'temu'],
  ['Shein', 'shein'],
  ['TikTok Shop', 'tiktok'],
  ['eBay', 'ebay'],
  ['Etsy', 'etsy'],
  ['AliExpress', 'aliexpress'],
  ['Wayfair', 'wayfair'],
  ['Target.com', 'target\\.com'],
];

const NAVY = '#1F3864';
const LIGHT = '#F2F4F7';
const MATCH_FILL = '#FFF2CC';

// ---------------------------------------------------------------- menu

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Recall Monitor')
    .addItem('Check for new recalls now', 'checkNow')
    .addItem('Re-check watchlist and rules', 'recheckAll')
    .addItem('Send a test alert', 'sendTestAlert')
    .addSeparator()
    .addItem('Set up the sheet', 'setup')
    .addItem('Rebuild the dashboard', 'rebuildDashboard')
    .addItem('Load all recalls since the start date', 'backfill')
    .addItem('Turn on the daily check', 'installDailyTrigger')
    .addItem('Turn off the daily check', 'removeDailyTrigger')
    .addToUi();
}

/** Builds any missing tab with its defaults. Safe to run again: it never deletes data. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());
  buildSettings_(ss);
  buildWatchlist_(ss);
  buildRules_(ss);
  buildRecalls_(ss);
  buildLog_(ss);
  buildGuide_(ss);
  buildDashboard_(ss);
  const order = [SHEET.DASHBOARD, SHEET.RECALLS, SHEET.WATCHLIST, SHEET.RULES, SHEET.SETTINGS, SHEET.LOG, SHEET.GUIDE];
  order.forEach(function (name, i) {
    ss.setActiveSheet(ss.getSheetByName(name));
    ss.moveActiveSheet(i + 1);
  });
  ss.getSheets().forEach(function (sh) {
    const ours = order.indexOf(sh.getName()) >= 0;
    if (!ours && sh.getLastRow() === 0 && sh.getLastColumn() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  ss.setActiveSheet(ss.getSheetByName(SHEET.DASHBOARD));
}

/** Deletes and rebuilds the Dashboard tab (its formulas and charts). Data tabs are not touched. */
function rebuildDashboard() {
  const ss = SpreadsheetApp.getActive();
  const old = ss.getSheetByName(SHEET.DASHBOARD);
  if (old) ss.deleteSheet(old);
  buildDashboard_(ss);
  ss.setActiveSheet(ss.getSheetByName(SHEET.DASHBOARD));
  ss.moveActiveSheet(1);
}

// ---------------------------------------------------------------- runs

/** Loads every recall dated from the start date in Settings until today. */
function backfill() {
  const settings = getSettings_();
  const from = asDate_(settings[SETTING.KEEP_FROM]) || new Date(2025, 0, 1);
  const records = fetchRecalls_({ RecallDateStart: ymd_(from), RecallDateEnd: ymd_(new Date()) });
  const result = syncRecalls_(records, 'Backfill');
  SpreadsheetApp.getActive().toast(result, 'Recall Monitor', 8);
  return result;
}

/** Daily run: asks the API for recalls published or updated since the last run. */
function dailySync() {
  const settings = getSettings_();
  const lastRun = asDate_(settings[SETTING.LAST_RUN]);
  const keepFrom = asDate_(settings[SETTING.KEEP_FROM]) || new Date(2025, 0, 1);
  const lookBack = Number(settings[SETTING.LOOK_BACK]) || 3;
  const since = lastRun ? new Date(lastRun.getTime() - lookBack * 86400000) : keepFrom;
  const records = fetchRecalls_({ LastPublishDateStart: ymd_(since), LastPublishDateEnd: ymd_(new Date()) });
  return syncRecalls_(records, 'Daily');
}

/** Menu: same as the daily run, started by hand. */
function checkNow() {
  const result = dailySync();
  SpreadsheetApp.getActive().toast(result, 'Recall Monitor', 8);
}

/** Applies the current watchlist and rules to every row, without calling the API. */
function recheckAll() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.RECALLS);
  const last = sheet.getLastRow();
  if (last < 2) return;
  const watchlist = getWatchlist_();
  const rules = getHazardRules_();
  const markets = getMarketplaces_();
  const range = sheet.getRange(2, 1, last - 1, HEADERS.length);
  const values = range.getValues();
  values.forEach(function (row) {
    const text = [row[COL.TITLE - 1], row[COL.PRODUCT - 1], row[COL.TYPE - 1]].join(' | ');
    const match = matchWatchlist_(text, watchlist);
    row[COL.MATCH - 1] = match.keywords.join(', ');
    row[COL.AREA - 1] = match.areas.join(', ');
    row[COL.HAZARD_CATEGORY - 1] = classifyHazard_(row[COL.HAZARD - 1], rules);
    row[COL.ONLINE - 1] = detectMarketplaces_(row[COL.SOLD_AT - 1], markets);
  });
  const n = values.length;
  sheet.getRange(2, COL.HAZARD_CATEGORY, n, 1).setValues(values.map(function (r) { return [r[COL.HAZARD_CATEGORY - 1]]; }));
  sheet.getRange(2, COL.ONLINE, n, 1).setValues(values.map(function (r) { return [r[COL.ONLINE - 1]]; }));
  sheet.getRange(2, COL.MATCH, n, 2).setValues(values.map(function (r) { return [r[COL.MATCH - 1], r[COL.AREA - 1]]; }));
  SpreadsheetApp.getActive().toast('Watchlist and rules applied to ' + values.length + ' recalls.', 'Recall Monitor', 6);
}

/**
 * Upserts the API records into the Recalls tab (one row per Recall ID), logs the run
 * and emails the new watchlist matches when alerts are on.
 */
function syncRecalls_(records, mode) {
  const started = new Date();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another run is in progress. Try again in a minute.');
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET.RECALLS);
    const settings = getSettings_();
    const keepFrom = asDate_(settings[SETTING.KEEP_FROM]) || new Date(2025, 0, 1);
    const watchlist = getWatchlist_();
    const rules = getHazardRules_();
    const markets = getMarketplaces_();

    const last = sheet.getLastRow();
    const existing = last > 1 ? sheet.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
    if (last > 1) {
      // getValues() returns the text of a formula cell, so put the CPSC links back.
      const links = sheet.getRange(2, COL.URL, last - 1, 1).getFormulas();
      existing.forEach(function (row, i) { if (links[i][0]) row[COL.URL - 1] = links[i][0]; });
    }
    const byId = {};
    existing.forEach(function (row, i) { byId[String(row[COL.ID - 1])] = i; });

    const counts = { received: records.length, added: 0, updated: 0, unchanged: 0, skipped: 0 };
    const newMatches = [];
    const seen = {};
    records.forEach(function (rec) {
      const id = String(rec.RecallID);
      if (seen[id]) return;  // the API can repeat a record inside one answer
      seen[id] = true;
      const recallDate = parseApiDate_(rec.RecallDate);
      if (!recallDate || recallDate < keepFrom) { counts.skipped++; return; }
      const row = toRow_(rec, watchlist, rules, markets, started);
      if (byId.hasOwnProperty(id)) {
        const old = existing[byId[id]];
        const oldPublished = asDate_(old[COL.LAST_PUBLISHED - 1]);
        const newPublished = row[COL.LAST_PUBLISHED - 1];
        if (oldPublished && newPublished && newPublished.getTime() <= oldPublished.getTime()) {
          counts.unchanged++;
          return;
        }
        row[COL.FIRST_SEEN - 1] = old[COL.FIRST_SEEN - 1] || started;
        existing[byId[id]] = row;
        counts.updated++;
      } else {
        byId[id] = existing.length;
        existing.push(row);
        counts.added++;
        if (row[COL.MATCH - 1]) newMatches.push(row);
      }
    });

    existing.sort(function (a, b) {
      const da = asDate_(a[COL.DATE - 1]), db = asDate_(b[COL.DATE - 1]);
      const diff = (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
      return diff !== 0 ? diff : Number(b[COL.ID - 1]) - Number(a[COL.ID - 1]);
    });
    if (existing.length) {
      if (last > 1) sheet.getRange(2, 1, last - 1, HEADERS.length).clearContent();
      const missing = existing.length + 1 - sheet.getMaxRows();
      if (missing > 0) sheet.insertRowsAfter(sheet.getMaxRows(), missing + 100);
      sheet.getRange(2, 1, existing.length, HEADERS.length).setValues(existing);
      formatRecallRows_(sheet, existing.length);
    }

    let emailed = 'No';
    // A backfill loads history, so it never emails; only the daily or manual checks do.
    if (mode !== 'Backfill' && newMatches.length && isYes_(settings[SETTING.ALERTS])) {
      sendAlert_(newMatches, false);
      emailed = 'Yes (' + newMatches.length + ')';
    }

    const seconds = Math.round((new Date() - started) / 1000);
    const result = counts.added + ' new, ' + counts.updated + ' updated, ' + newMatches.length + ' new watchlist matches';
    log_([started, mode, counts.received, counts.added, counts.updated, counts.unchanged,
          counts.skipped, newMatches.length, emailed, seconds, 'OK: ' + result]);
    setSetting_(SETTING.LAST_RUN, started);
    setSetting_(SETTING.LAST_RESULT, 'OK: ' + result);
    return result;
  } catch (err) {
    log_([started, mode, records.length, '', '', '', '', '', 'No',
          Math.round((new Date() - started) / 1000), 'ERROR: ' + err.message]);
    setSetting_(SETTING.LAST_RESULT, 'ERROR: ' + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- API

/** Calls the CPSC Recalls API with retries and returns an array of recall records. */
function fetchRecalls_(params) {
  const query = Object.keys(params).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
  const url = API_URL + '?format=json&' + query;
  let lastError = '';
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { Accept: 'application/json' } });
      const code = response.getResponseCode();
      if (code === 200) {
        const data = JSON.parse(response.getContentText());
        if (!Array.isArray(data)) throw new Error('Unexpected answer from the CPSC API.');
        return data;
      }
      lastError = 'HTTP ' + code;
    } catch (err) {
      lastError = err.message;
    }
    Utilities.sleep(2000 * attempt);
  }
  throw new Error('CPSC API did not answer after 3 tries (' + lastError + ').');
}

// ---------------------------------------------------------------- record to row

/** Turns one API record into a row of the Recalls tab. */
function toRow_(rec, watchlist, rules, markets, now) {
  const products = rec.Products || [];
  const productNames = unique_(products.map(function (p) { return clean_(p.Name); }));
  const productTypes = unique_(products.map(function (p) { return clean_(p.Type); }));
  const unitsText = unique_(products.map(function (p) { return clean_(p.NumberOfUnits); }));
  let units = 0, anyUnits = false;
  products.forEach(function (p) {
    const n = parseUnits_(p.NumberOfUnits);
    if (n !== null) { units += n; anyUnits = true; }
  });
  const hazard = names_(rec.Hazards);
  const soldAt = cleanSoldAt_(names_(rec.Retailers));
  const matchText = [rec.Title, productNames.join(' | '), productTypes.join(' | ')].join(' | ');
  const match = matchWatchlist_(matchText, watchlist);
  const remedy = unique_((rec.RemedyOptions || []).map(function (o) { return clean_(o.Option); }));
  const madeIn = unique_((rec.ManufacturerCountries || []).map(function (c) { return clean_(c.Country); }));
  return [
    Number(rec.RecallID),
    String(rec.RecallNumber || ''),
    parseApiDate_(rec.RecallDate),
    clean_(rec.Title),
    productNames.join('; '),
    productTypes.join('; '),
    anyUnits ? units : '',
    unitsText.join('; '),
    classifyHazard_(hazard, rules),
    truncate_(hazard, 500),
    remedy.join(', '),
    madeIn.join(', '),
    detectMarketplaces_(soldAt, markets),
    truncate_(soldAt, 500),
    match.keywords.join(', '),
    match.areas.join(', '),
    truncate_(names_(rec.Injuries), 500),
    rec.URL ? '=HYPERLINK("' + String(rec.URL).replace(/"/g, '%22') + '","CPSC page")' : '',
    parseApiDate_(rec.LastPublishDate),
    now,
    now,
  ];
}

/** "About 265,000" -> 265000; "5.9 million" -> 5900000; the Canadian extra in brackets is ignored. */
function parseUnits_(text) {
  if (!text) return null;
  const head = String(text).split('(')[0];
  const m = head.match(/(\d[\d,]*(?:\.\d+)?)\s*(million|thousand)?/i);
  if (!m) return null;
  let value = parseFloat(m[1].replace(/,/g, ''));
  if (m[2]) value *= m[2].toLowerCase() === 'million' ? 1e6 : 1e3;
  return Math.round(value);
}

function classifyHazard_(text, rules) {
  const t = String(text || '').toLowerCase().replace(/\s+/g, ' ');
  if (!t) return 'Other';
  for (let i = 0; i < rules.length; i++) {
    if (rules[i].re.test(t)) return rules[i].category;
  }
  return 'Other';
}

function detectMarketplaces_(soldAt, markets) {
  const t = String(soldAt || '');
  return markets.filter(function (m) { return m.re.test(t); })
    .map(function (m) { return m.name; }).join(', ');
}

/** Whole-word match (plurals included) of each active keyword in the title and product fields. */
function matchWatchlist_(text, watchlist) {
  text = String(text || '').replace(/[\u2018\u2019]/g, "'");
  const keywords = [], areas = [];
  watchlist.forEach(function (w) {
    if (w.re.test(text)) {
      if (keywords.indexOf(w.keyword) < 0) keywords.push(w.keyword);
      if (w.area && areas.indexOf(w.area) < 0) areas.push(w.area);
    }
  });
  return { keywords: keywords, areas: areas };
}

function keywordRegex_(keyword) {
  const escaped = String(keyword).trim().replace(/[\u2018\u2019]/g, "'").replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
  return new RegExp('(^|[^A-Za-z])' + escaped + '(s|es)?(?![A-Za-z])', 'i');
}

// ---------------------------------------------------------------- email

function sendAlert_(rows, isTest) {
  const to = String(getSettings_()[SETTING.EMAIL] || '').trim() || Session.getEffectiveUser().getEmail();
  const tz = Session.getScriptTimeZone();
  const subject = (isTest ? '[TEST] ' : '') + 'Recall Monitor: ' + rows.length + ' new recall' +
    (rows.length === 1 ? ' matches' : 's match') + ' your watchlist';
  const shown = rows.slice(0, 25);
  const items = shown.map(function (r) {
    const url = (String(r[COL.URL - 1]).match(/HYPERLINK\("([^"]+)"/) || [])[1] || '';
    const date = r[COL.DATE - 1] instanceof Date ? Utilities.formatDate(r[COL.DATE - 1], tz, 'd MMM yyyy') : '';
    return '<tr>' +
      td_('<a href="' + url + '">' + esc_(r[COL.TITLE - 1]) + '</a>') +
      td_(esc_(date)) + td_(esc_(r[COL.HAZARD_CATEGORY - 1])) + td_(esc_(r[COL.REMEDY - 1])) +
      td_(esc_(r[COL.ONLINE - 1] || '-')) + td_(esc_(r[COL.MATCH - 1])) + '</tr>';
  }).join('');
  const html =
    '<p>' + (isTest ? '<b>This is a test.</b> ' : '') + 'These CPSC recalls match your watchlist:</p>' +
    '<table style="border-collapse:collapse;font-family:Arial;font-size:13px">' +
    '<tr style="background:' + NAVY + ';color:#fff">' +
    ['Recall', 'Date', 'Hazard', 'Remedy', 'Sold online at', 'Matched'].map(function (h) { return td_(h); }).join('') +
    '</tr>' + items + '</table>' +
    (rows.length > shown.length ? '<p style="font-family:Arial;font-size:13px">And ' + (rows.length - shown.length) +
      ' more in the sheet.</p>' : '') +
    '<p style="font-family:Arial;font-size:12px;color:#595959">Full details: ' +
    '<a href="' + SpreadsheetApp.getActive().getUrl() + '">Recall Monitor sheet</a>. Data: U.S. CPSC Recalls API.</p>';
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html });
}

/** Menu: emails the 5 most recent watchlist matches, marked as a test. */
function sendTestAlert() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.RECALLS);
  const last = sheet.getLastRow();
  if (last < 2) throw new Error('No recalls yet. Load them first.');
  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(function (r) { return r[COL.MATCH - 1]; }).slice(0, 5);
  const formulas = sheet.getRange(2, COL.URL, last - 1, 1).getFormulas();
  const urlById = {};
  sheet.getRange(2, COL.ID, last - 1, 1).getValues().forEach(function (r, i) { urlById[r[0]] = formulas[i][0]; });
  rows.forEach(function (r) { r[COL.URL - 1] = urlById[r[COL.ID - 1]] || ''; });
  if (!rows.length) throw new Error('No watchlist matches yet.');
  sendAlert_(rows, true);
  SpreadsheetApp.getActive().toast('Test alert sent with ' + rows.length + ' recalls.', 'Recall Monitor', 6);
}

// ---------------------------------------------------------------- triggers

function installDailyTrigger() {
  removeDailyTrigger();
  const hour = Math.min(23, Math.max(0, Number(getSettings_()[SETTING.HOUR]) || 7));
  ScriptApp.newTrigger('dailySync').timeBased().everyDays(1).atHour(hour).create();
  SpreadsheetApp.getActive().toast('Daily check on, around ' + hour + ':00 every day.', 'Recall Monitor', 6);
}

function removeDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailySync') ScriptApp.deleteTrigger(t);
  });
}

// ---------------------------------------------------------------- settings and lists

function getSettings_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
  const values = sheet.getRange(2, 1, 12, 2).getValues();
  const out = {};
  values.forEach(function (r) { if (r[0]) out[String(r[0]).trim()] = r[1]; });
  return out;
}

function setSetting_(label, value) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
  const labels = sheet.getRange(2, 1, 12, 1).getValues();
  for (let i = 0; i < labels.length; i++) {
    if (String(labels[i][0]).trim() === label) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
}

function getWatchlist_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.WATCHLIST);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 3).getValues()
    .filter(function (r) { return String(r[0]).trim() && !/^no$/i.test(String(r[2]).trim()); })
    .map(function (r) {
      return { keyword: String(r[0]).trim(), area: String(r[1]).trim(), re: keywordRegex_(r[0]) };
    });
}

function getHazardRules_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.RULES);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 2).getValues()
    .filter(function (r) { return String(r[0]).trim() && String(r[1]).trim(); })
    .map(function (r) { return { category: String(r[0]).trim(), re: new RegExp(String(r[1]).trim(), 'i') }; });
}

function getMarketplaces_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.RULES);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 4, last - 1, 2).getValues()
    .filter(function (r) { return String(r[0]).trim() && String(r[1]).trim(); })
    .map(function (r) { return { name: String(r[0]).trim(), re: new RegExp(String(r[1]).trim(), 'i') }; });
}

function log_(row) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.LOG);
  sheet.insertRowAfter(1);  // newest run on top
  sheet.getRange(2, 1, 1, row.length).setValues([row]);
  // An inserted row copies the header's style, so reset the style of the log rows.
  sheet.getRange(2, 1, sheet.getLastRow() - 1, LOG_HEADERS.length)
    .setBackground(null).setFontColor('#000000').setFontWeight('normal');
  sheet.getRange(2, 1).setNumberFormat('yyyy-mm-dd hh:mm');
}

// ---------------------------------------------------------------- building the tabs

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader_(range) {
  range.setFontWeight('bold').setFontColor('#FFFFFF').setBackground(NAVY)
    .setVerticalAlignment('middle').setWrap(true);
}

function buildSettings_(ss) {
  const sheet = getOrCreate_(ss, SHEET.SETTINGS);
  if (sheet.getLastRow() > 0) return;
  sheet.getRange('A1:C1').setValues([['Setting', 'Value', 'Notes']]);
  styleHeader_(sheet.getRange('A1:C1'));
  sheet.getRange('A2:C8').setValues([
    [SETTING.KEEP_FROM, new Date(2025, 0, 1), 'Recalls dated before this are not kept.'],
    [SETTING.EMAIL, '', 'Blank = the email of the sheet owner.'],
    [SETTING.ALERTS, 'No', 'Yes = email the new watchlist matches after each run.'],
    [SETTING.HOUR, 7, 'Hour of the daily check, in the time zone of the sheet.'],
    [SETTING.LOOK_BACK, 3, 'The daily check also picks up recalls the CPSC updated in these days.'],
    [SETTING.LAST_RUN, '', 'Written by the script.'],
    [SETTING.LAST_RESULT, '', 'Written by the script.'],
  ]);
  sheet.getRange('B2').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('B7').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange('B2:B6').setBackground('#FFFFCC').setFontColor('#0000FF');
  sheet.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No']).build());
  sheet.getRange('A10').setValue('Data source: U.S. Consumer Product Safety Commission (CPSC) Recalls API, ' +
    API_URL + ' (U.S. government data, public domain).').setFontStyle('italic').setFontColor('#595959');
  sheet.setColumnWidth(1, 240); sheet.setColumnWidth(2, 170); sheet.setColumnWidth(3, 470);
  sheet.setFrozenRows(1);
}

function buildWatchlist_(ss) {
  const sheet = getOrCreate_(ss, SHEET.WATCHLIST);
  if (sheet.getLastRow() > 0) return;
  sheet.getRange('A1:C1').setValues([['Keyword', 'Area', 'Active']]);
  styleHeader_(sheet.getRange('A1:C1'));
  sheet.getRange(2, 1, DEFAULT_WATCHLIST.length, 3).setValues(DEFAULT_WATCHLIST);
  sheet.getRange(2, 3, 200, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No']).setAllowInvalid(false).build());
  sheet.getRange('E1').setValue('How matching works').setFontWeight('bold');
  sheet.getRange('E2:E5').setValues([
    ['Each active keyword is searched as a whole word, plurals included, in the recall title, product name and product type.'],
    ['Add your own categories, brands or product words in new rows. Area groups them in the alert.'],
    ['Set Active to No to pause a keyword without deleting it.'],
    ['After editing, use Recall Monitor > Re-check watchlist and rules.'],
  ]);
  sheet.setColumnWidth(1, 160); sheet.setColumnWidth(2, 200); sheet.setColumnWidth(3, 70); sheet.setColumnWidth(5, 640);
  sheet.setFrozenRows(1);
}

function buildRules_(ss) {
  const sheet = getOrCreate_(ss, SHEET.RULES);
  if (sheet.getLastRow() > 0) return;
  sheet.getRange('A1:B1').setValues([['Hazard category', 'Words in the hazard text (regular expression; first match wins)']]);
  sheet.getRange('D1:E1').setValues([['Online store', 'Words in "Sold at" (regular expression)']]);
  styleHeader_(sheet.getRange('A1:B1'));
  styleHeader_(sheet.getRange('D1:E1'));
  sheet.getRange(2, 1, DEFAULT_HAZARD_RULES.length, 2).setValues(DEFAULT_HAZARD_RULES);
  sheet.getRange(2, 4, DEFAULT_MARKETPLACES.length, 2).setValues(DEFAULT_MARKETPLACES);
  sheet.getRange(DEFAULT_HAZARD_RULES.length + 3, 1)
    .setValue('A hazard text that matches no rule is classed as "Other". Rules are checked from top to bottom.')
    .setFontStyle('italic').setFontColor('#595959');
  sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 520); sheet.setColumnWidth(3, 30);
  sheet.setColumnWidth(4, 130); sheet.setColumnWidth(5, 260);
  sheet.setFrozenRows(1);
}

function buildRecalls_(ss) {
  const sheet = getOrCreate_(ss, SHEET.RECALLS);
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  styleHeader_(sheet.getRange(1, 1, 1, HEADERS.length));
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);
  const widths = [70, 80, 90, 360, 220, 160, 80, 140, 150, 300, 110, 100, 130, 300, 140, 170, 300, 90, 95, 120, 120];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$O2<>""').setBackground(MATCH_FILL)
    .setRanges([sheet.getRange('A2:U')]).build();
  sheet.setConditionalFormatRules([rule]);
}

function formatRecallRows_(sheet, n) {
  if (n < 1) return;
  sheet.getRange(2, COL.DATE, n, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, COL.UNITS, n, 1).setNumberFormat('#,##0');
  sheet.getRange(2, COL.LAST_PUBLISHED, n, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, COL.FIRST_SEEN, n, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(2, 1, n, HEADERS.length).setVerticalAlignment('top').setWrap(false);
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.getRange(1, 1, n + 1, HEADERS.length).createFilter();
}

function buildLog_(ss) {
  const sheet = getOrCreate_(ss, SHEET.LOG);
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  styleHeader_(sheet.getRange(1, 1, 1, LOG_HEADERS.length));
  sheet.setFrozenRows(1);
  [130, 80, 110, 60, 70, 80, 150, 140, 90, 70, 420].forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
}

function buildGuide_(ss) {
  const sheet = getOrCreate_(ss, SHEET.GUIDE);
  if (sheet.getLastRow() > 0) return;
  const lines = [
    ['What this sheet does'],
    ['Every day it asks the U.S. CPSC Recalls API for new and updated product recalls, keeps one row per recall in Recalls, and flags the recalls that match your Watchlist.'],
    ['It classes each hazard (fire, choking, tip-over...) with the rules in Rules, spots online stores in "Sold at" (Amazon, Walmart.com, Temu...) and can email you the new matches.'],
    [''],
    ['First time'],
    ['1. Recall Monitor > Set up the sheet (Google asks you to allow the script once).'],
    ['2. Recall Monitor > Load all recalls since the start date (Settings).'],
    ['3. Recall Monitor > Turn on the daily check.'],
    ['4. Optional: Settings > Send email alerts = Yes, and Recall Monitor > Send a test alert.'],
    [''],
    ['Every day (automatic)'],
    ['The daily check picks up recalls published or updated since the last run. Each run is written in Log.'],
    ['Edit Watchlist or Rules at any time, then use Recall Monitor > Re-check watchlist and rules.'],
    [''],
    ['Columns in Recalls'],
    ['Units: units recalled in the U.S., read from the published text ("About 265,000" -> 265,000; units sold in Canada are left out).'],
    ['Hazard category: first rule in Rules that matches the hazard text. Sold online at: online stores named in "Sold at".'],
    ['Watchlist match / area: the keywords found in the title, product name or product type.'],
    [''],
    ['Data source'],
    ['U.S. Consumer Product Safety Commission (CPSC) Recalls API: ' + API_URL + ' (U.S. government data, public domain).'],
  ];
  sheet.getRange(1, 1, lines.length, 1).setValues(lines);
  [1, 5, 11, 15, 20].forEach(function (r) { sheet.getRange(r, 1).setFontWeight('bold').setFontColor(NAVY).setFontSize(12); });
  sheet.setColumnWidth(1, 980);
  sheet.setHiddenGridlines(true);
}

function buildDashboard_(ss) {
  const sheet = getOrCreate_(ss, SHEET.DASHBOARD);
  if (sheet.getLastRow() > 0) return;
  sheet.setHiddenGridlines(true);
  const R = "'" + SHEET.RECALLS + "'!";
  const yearStart = 'DATE(YEAR(TODAY()),1,1)';

  sheet.getRange('A1').setValue('Product Recall Monitor').setFontSize(20).setFontWeight('bold').setFontColor(NAVY);
  sheet.getRange('A2').setValue('U.S. CPSC recalls from the public API, checked every day by Apps Script · watchlist: baby and kids products')
    .setFontStyle('italic').setFontColor('#595959');
  sheet.getRange('A3').setFormula('="Last run: "&IF(Settings!B7="","not yet",TEXT(Settings!B7,"yyyy-mm-dd hh:mm"))&"   ·   "&Settings!B8')
    .setFontColor('#595959');

  const cards = [
    ['RECALLS THIS YEAR', '=COUNTIFS(' + R + 'C:C,">="&' + yearStart + ')', '#,##0'],
    ['SAME PERIOD LAST YEAR', '=COUNTIFS(' + R + 'C:C,">="&DATE(YEAR(TODAY())-1,1,1),' + R + 'C:C,"<="&EDATE(TODAY(),-12))', '#,##0'],
    ['CHANGE', '=IF(B6=0,"",A6/B6-1)', '+0%;-0%;0%'],
    ['LAST 30 DAYS', '=COUNTIFS(' + R + 'C:C,">="&TODAY()-30)', '#,##0'],
    ['SOLD ON AMAZON', '=IF(A6=0,"",COUNTIFS(' + R + 'C:C,">="&' + yearStart + ',' + R + 'M:M,"*Amazon*")/A6)', '0%'],
    ['WATCHLIST MATCHES', '=COUNTIFS(' + R + 'C:C,">="&' + yearStart + ',' + R + 'O:O,"<>")', '#,##0'],
    ['UNITS RECALLED', '=SUMIFS(' + R + 'G:G,' + R + 'C:C,">="&' + yearStart + ')', '#,##0,,"M"'],
    ['NEW IN LAST RUN', '=IFERROR(INDEX(Log!D:D,2),0)', '#,##0'],  // whole column: new log rows are inserted at row 2
  ];
  cards.forEach(function (c, i) {
    sheet.getRange(5, i + 1).setValue(c[0]);
    sheet.getRange(6, i + 1).setFormula(c[1]).setNumberFormat(c[2]);
  });
  sheet.getRange('A5:H5').setFontSize(8).setFontWeight('bold').setFontColor('#595959').setBackground(LIGHT)
    .setHorizontalAlignment('center').setWrap(true);
  sheet.getRange('A6:H6').setFontSize(18).setFontWeight('bold').setFontColor(NAVY).setBackground(LIGHT)
    .setHorizontalAlignment('center');
  sheet.getRange('A5:H6').setBorder(true, true, true, true, true, false, '#D9DEE7', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A7').setValue('"This year" = 1 Jan to today. Same period last year = 1 Jan to the same day last year.')
    .setFontSize(8).setFontStyle('italic').setFontColor('#595959');

  // Recalls by month, this year vs last year
  sheet.getRange('A9').setValue('Recalls by month').setFontWeight('bold').setFontColor(NAVY);
  sheet.getRange('A10:C10').setValues([['Month', 'Last year', 'This year']]);
  styleHeader_(sheet.getRange('A10:C10'));
  for (let m = 0; m < 12; m++) {
    const r = 11 + m;
    sheet.getRange(r, 1).setFormula(m === 0 ? '=DATE(YEAR(TODAY()),1,1)' : '=EDATE(A' + (r - 1) + ',1)');
    sheet.getRange(r, 2).setFormula('=COUNTIFS(' + R + 'C:C,">="&EDATE(A' + r + ',-12),' + R + 'C:C,"<"&EDATE(A' + r + ',-11))');
    sheet.getRange(r, 3).setFormula('=IF(A' + r + '>TODAY(),"",COUNTIFS(' + R + 'C:C,">="&A' + r + ',' + R + 'C:C,"<"&EDATE(A' + r + ',1)))');
  }
  sheet.getRange('A11:A22').setNumberFormat('mmm').setHorizontalAlignment('left');
  sheet.getRange('A23').setValue('Total');
  sheet.getRange('B23').setFormula('=SUM(B11:B22)');
  sheet.getRange('C23').setFormula('=SUM(C11:C22)');
  sheet.getRange('A23:C23').setFontWeight('bold').setBackground(LIGHT);

  // Hazard categories this year
  const q = function (select, extra) {
    return '=IFERROR(QUERY(' + R + 'A2:U,"select ' + select + ' where C >= date \'"&TEXT(' + yearStart +
      ',"yyyy-mm-dd")&"\' ' + (extra || '') + '",0),"")';
  };
  sheet.getRange('A26').setValue('Hazards this year').setFontWeight('bold').setFontColor(NAVY);
  sheet.getRange('A27:B27').setValues([['Hazard category', 'Recalls']]);
  styleHeader_(sheet.getRange('A27:B27'));
  sheet.getRange('A28').setFormula(q("I, count(A)", "group by I order by count(A) desc label I '', count(A) ''"));

  // Online stores this year
  sheet.getRange('A43').setValue('Sold online at (this year)').setFontWeight('bold').setFontColor(NAVY);
  sheet.getRange('A44:C44').setValues([['Online store', 'Recalls', 'Share']]);
  styleHeader_(sheet.getRange('A44:C44'));
  for (let i = 0; i < DEFAULT_MARKETPLACES.length; i++) {
    const r = 45 + i;
    sheet.getRange(r, 1).setFormula('=IF(Rules!D' + (i + 2) + '="","",Rules!D' + (i + 2) + ')');
    sheet.getRange(r, 2).setFormula('=IF(A' + r + '="","",COUNTIFS(' + R + 'C:C,">="&' + yearStart + ',' + R + 'M:M,"*"&A' + r + '&"*"))');
    sheet.getRange(r, 3).setFormula('=IF(OR(A' + r + '="",$A$6=0),"",B' + r + '/$A$6)').setNumberFormat('0%');
  }

  // Made in and remedies this year
  sheet.getRange('E43').setValue('Made in (this year, top 6)').setFontWeight('bold').setFontColor(NAVY);
  sheet.getRange('E44:F44').setValues([['Country', 'Recalls']]);
  styleHeader_(sheet.getRange('E44:F44'));
  sheet.getRange('E45').setFormula(q("L, count(A)", "and L <> \'\' group by L order by count(A) desc limit 6 label L '', count(A) ''"));
  sheet.getRange('H43').setValue('Remedy (this year)').setFontWeight('bold').setFontColor(NAVY);
  sheet.getRange('H44:I44').setValues([['Remedy', 'Recalls']]);
  styleHeader_(sheet.getRange('H44:I44'));
  ['Refund', 'Repair', 'Replace'].forEach(function (rem, i) {
    sheet.getRange(45 + i, 8).setValue(rem);
    sheet.getRange(45 + i, 9).setFormula('=COUNTIFS(' + R + 'C:C,">="&' + yearStart + ',' + R + 'K:K,"*' + rem + '*")');
  });

  // Latest watchlist matches
  sheet.getRange('A57').setValue('Latest watchlist matches').setFontWeight('bold').setFontColor(NAVY);
  sheet.getRange('A58:I58').setValues([['Recall date', 'Title', '', '', '', 'Matched', 'Hazard', 'Remedy', 'Sold online at']]);
  styleHeader_(sheet.getRange('A58:I58'));
  // Two queries with the same order: the title spills over the empty columns C:E.
  const latest = function (select) {
    return '=IFERROR(QUERY(' + R + 'A2:U,"select ' + select + " where O <> '' order by C desc, A desc limit 12\",0),\"\")";
  };
  sheet.getRange('A59').setFormula(latest('C, D'));
  sheet.getRange('F59').setFormula(latest('O, I, K, M'));
  sheet.getRange('A59:A70').setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('left');
  sheet.getRange('B59:I70').setWrap(false);

  [190, 120, 110, 110, 150, 120, 120, 120, 90].forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setRowHeight(5, 34);
  sheet.setRowHeight(6, 40);
  sheet.setFrozenRows(3);

  // Charts
  const monthChart = sheet.newChart().asColumnChart()
    .addRange(sheet.getRange('A10:C22'))
    .setNumHeaders(1)
    .setOption('title', 'Recalls by month: last year vs this year')
    .setOption('titleTextStyle', { color: NAVY, fontSize: 14, bold: true })
    .setOption('colors', ['#A5A5A5', '#2E75B6'])
    .setOption('legend', { position: 'bottom' })
    .setOption('hAxis', { format: 'MMM' })
    .setPosition(9, 5, 0, 0)
    .setOption('width', 620).setOption('height', 300)
    .build();
  sheet.insertChart(monthChart);

  const hazardChart = sheet.newChart().asBarChart()
    .addRange(sheet.getRange('A27:B38'))
    .setNumHeaders(1)
    .setOption('title', 'Recalls this year by hazard')
    .setOption('titleTextStyle', { color: NAVY, fontSize: 14, bold: true })
    .setOption('colors', ['#2E75B6'])
    .setOption('legend', { position: 'none' })
    .setPosition(26, 5, 0, 0)
    .setOption('width', 620).setOption('height', 300)
    .build();
  sheet.insertChart(hazardChart);
}

// ---------------------------------------------------------------- small helpers

function parseApiDate_(value) {
  if (!value) return '';
  const m = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : '';
}

function asDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)) return parseApiDate_(value);
  return null;
}

function ymd_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function names_(list) {
  return unique_((list || []).map(function (x) { return clean_(x && x.Name); })).join(' | ');
}

function clean_(value) {
  return String(value === null || value === undefined ? '' : value).replace(/\s+/g, ' ').trim();
}

function cleanSoldAt_(text) {
  return clean_(text).replace(/^sold (at|exclusively at|online at|exclusively online at)\s*:?\s*/i, '');
}

function unique_(list) {
  const out = [];
  list.forEach(function (x) { if (x && out.indexOf(x) < 0) out.push(x); });
  return out;
}

function truncate_(text, max) {
  const t = String(text || '');
  return t.length > max ? t.slice(0, max - 1) + '…' : t;
}

function isYes_(value) {
  return /^(yes|y|true|si|sí)$/i.test(String(value).trim());
}

function esc_(text) {
  return String(text || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function td_(html) {
  return '<td style="border:1px solid #D9DEE7;padding:6px;vertical-align:top">' + html + '</td>';
}
