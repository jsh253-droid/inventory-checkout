/**
 * Inventory check-out backend and admin portal.
 * Paste into a Google Apps Script project (with Admin.html) and deploy twice:
 *   public API:   Execute as Me, access Anyone
 *   admin portal: Execute as User accessing, access Anyone with Google account
 * See README.md.
 */

const SPREADSHEET_ID = '1EvEk_jlDcMM_yMkKPpHd5VjLJ0KWqfQkATLVHVpzJVQ';

// Check-outs and check-ins are written as rows in the Google Form's response
// sheet, in the form's own format, so each inventory tab's Total Out /
// Current Qty formulas count them alongside real form submissions.
const FORM_SHEET = 'Form Responses 1';
const FORM_COLS = 66; // A..BN, through "Notes?"
const ITEMS_PER_BLOCK = 10;
const TAKE_OUT = 'Take out';
const CHECK_IN = 'Check in';

// Inventory tab columns (1-based). Total Out and Current Qty are formulas.
const COL = { id: 1, name: 2, type: 3, category: 4, location: 5, unit: 6, start: 7, out: 8, qty: 9, notes: 10 };
const INV_COLS = 10;
const EDITABLE = [1, 2, 3, 4, 5, 6, 7, 10]; // everything but the formulas

const CACHE_KEY = 'data';
const CACHE_SECONDS = 20;

// Used until an admin saves settings in the portal.
const DEFAULT_CONFIG = {
  title: 'Inventory Check Out',
  requireEvent: false,
  offices: [
    { name: 'Warehouse', sheet: 'Warehouse', firstRow: 2 },
    { name: 'OFSA Office', sheet: 'OFSA', firstRow: 2 },
    { name: 'DM Office', sheet: 'DM Office', firstRow: 2 },
  ],
};

function doGet(e) {
  if (e.parameter.page === 'admin') {
    if (!runsAsViewer_()) {
      return HtmlService.createHtmlOutput('The admin portal is on its own deployment. Use the Admin link on the check-out site.');
    }
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('Inventory Admin')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return handle_(() => getData_(false));
}

// Writes return the fresh data, so the page needs only one round trip.
function doPost(e) {
  return handle_(() => {
    const body = JSON.parse(e.postData.contents);
    withLock_(() => {
      if (body.action === 'checkout') return checkout_(body);
      if (body.action === 'checkin') return checkin_(body);
      throw new Error('Unknown action: ' + body.action);
    });
    return getData_(true);
  });
}

function handle_(fn) {
  let out;
  try {
    out = { ok: true, data: fn() };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const result = fn();
    invalidate_();
    return result;
  } finally {
    lock.releaseLock();
  }
}

// After a write: recalculate formulas and drop the cached page data.
function invalidate_() {
  SpreadsheetApp.flush();
  CacheService.getScriptCache().remove(CACHE_KEY);
}

// True on the admin deployment (Execute as user accessing) and in the editor.
// False on the public API, which runs as the owner: there, google.script.run
// calls would run with the owner's access instead of the admin's.
function runsAsViewer_() {
  const active = Session.getActiveUser().getEmail().toLowerCase();
  return !!active && active === Session.getEffectiveUser().getEmail().toLowerCase();
}

let SS;
function ss_() {
  return SS || (SS = SpreadsheetApp.openById(SPREADSHEET_ID));
}

// ---------- Settings ----------

const CONFIG_KEY = 'config:' + SPREADSHEET_ID;

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG_KEY);
  if (raw) return JSON.parse(raw);
  // First run on this spreadsheet: keep the admin list from the previous one.
  const legacy = JSON.parse(props.getProperty('config') || '{}');
  return Object.assign({}, DEFAULT_CONFIG, { admins: legacy.admins || [] });
}

function writeConfig_(cfg) {
  PropertiesService.getScriptProperties().setProperty(CONFIG_KEY, JSON.stringify(cfg));
}

function findOffice_(name) {
  const office = getConfig_().offices.find(o => o.name === name);
  if (!office) throw new Error('Unknown office: ' + name);
  return office;
}

function saveConfig_(c) {
  const sheetNames = ss_().getSheets().map(s => s.getName());
  const offices = (c.offices || []).map(o => ({
    name: String(o.name || '').trim(),
    sheet: String(o.sheet || ''),
    firstRow: Math.max(1, Math.floor(Number(o.firstRow)) || 2),
  }));
  if (!offices.length) throw new Error('Add at least one office.');
  offices.forEach(o => {
    if (!o.name) throw new Error('Every office needs a name.');
    if (!sheetNames.includes(o.sheet)) throw new Error('Sheet tab not found: ' + o.sheet);
    formBlock_(ss_().getSheetByName(o.sheet), o.firstRow);
  });
  if (new Set(offices.map(o => o.name)).size !== offices.length) throw new Error('Office names must be different.');
  const cfg = {
    title: String(c.title || '').trim() || DEFAULT_CONFIG.title,
    requireEvent: !!c.requireEvent,
    offices,
    admins: getConfig_().admins || [],
  };
  writeConfig_(cfg);
  return cfg;
}

// ---------- Inventory tabs ----------

function sheetFor_(office) {
  const sh = ss_().getSheetByName(office.sheet);
  if (!sh) throw new Error('Sheet tab not found: ' + office.sheet);
  return sh;
}

// Rows from firstRow down with anything in the editable columns.
function readRows_(office) {
  const sh = sheetFor_(office);
  const first = Number(office.firstRow) || 2;
  const last = sh.getLastRow();
  if (last < first) return [];
  return sh.getRange(first, 1, last - first + 1, INV_COLS).getDisplayValues()
    .map((values, i) => ({ row: first + i, values }))
    .filter(r => EDITABLE.some(c => String(r.values[c - 1]).trim()));
}

// Items on an office tab, plus the form-sheet column holding its "Item 1".
function readOffice_(office) {
  const sh = sheetFor_(office);
  const items = readRows_(office)
    .filter(r => r.values[COL.name - 1].trim())
    .map(r => {
      const v = r.values;
      return {
        name: v[COL.name - 1].trim(), type: v[COL.type - 1], category: v[COL.category - 1],
        location: v[COL.location - 1], unit: v[COL.unit - 1], qty: v[COL.qty - 1],
      };
    });
  return { name: office.name, items, block: formBlock_(sh, Number(office.firstRow) || 2) };
}

// Finds the office's item columns from its Total Out formula, e.g.
// ('Form Responses 1'!$Z$2:$AR=B2) -> column Z.
function formBlock_(sh, firstRow) {
  const formula = sh.getRange(firstRow, COL.out).getFormula();
  const m = formula.match(/'Form Responses 1'!\$([A-Z]+)\$\d+:\$[A-Z]+\s*=\s*B\d/);
  if (!m) throw new Error(`Can't find the form columns in ${sh.getName()}!H${firstRow}'s Total Out formula.`);
  return m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

function stockNumber_(qty) {
  const n = Number(String(qty).replace(/,/g, ''));
  return isFinite(n) && String(qty).trim() !== '' ? n : null;
}

// Guards admin edits against rows that moved since the portal loaded them.
function checkRow_(sh, office, row, original) {
  if (row < (Number(office.firstRow) || 2)) throw new Error('Row is above the inventory.');
  const current = sh.getRange(row, 1, 1, INV_COLS).getDisplayValues()[0];
  if (JSON.stringify(EDITABLE.map(c => current[c - 1])) !== JSON.stringify(original)) {
    throw new Error('That row changed in the sheet since you loaded it. Reload and try again.');
  }
}

// values: the EDITABLE columns, in order (A–G, then J).
function writeRow_(sh, row, values) {
  sh.getRange(row, 1, 1, 7).setValues([values.slice(0, 7)]);
  sh.getRange(row, COL.notes).setValue(values[7]);
}

function saveRow_(p) {
  const office = findOffice_(p.office);
  const sh = sheetFor_(office);
  checkRow_(sh, office, p.row, p.original);
  writeRow_(sh, p.row, p.values);
}

function addRow_(p) {
  if (!String(p.values[1] || '').trim()) throw new Error('Enter an item name.');
  const office = findOffice_(p.office);
  const sh = sheetFor_(office);
  const rows = readRows_(office);
  const row = rows.length ? rows[rows.length - 1].row + 1 : Number(office.firstRow) || 2;
  writeRow_(sh, row, p.values);
  // New rows at the bottom may not have the Total Out / Current Qty formulas yet.
  if (!sh.getRange(row, COL.out).getFormula() && row > (Number(office.firstRow) || 2)) {
    sh.getRange(row - 1, COL.out, 1, 2).copyTo(sh.getRange(row, COL.out, 1, 2));
  }
}

function deleteRow_(p) {
  const office = findOffice_(p.office);
  const sh = sheetFor_(office);
  checkRow_(sh, office, p.row, p.original);
  sh.deleteRow(p.row);
}

// ---------- Form log ----------

function formSheet_() {
  const sh = ss_().getSheetByName(FORM_SHEET);
  if (!sh) throw new Error('Sheet tab not found: ' + FORM_SHEET);
  return sh;
}

// Every form row as {row, ts, who, team, action, event, notes, items: [{office, item, qty}]}.
function readForm_(offices) {
  const sh = formSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, FORM_COLS).getValues().map((r, i) => {
    const items = [];
    offices.forEach(o => {
      for (let k = 0; k < ITEMS_PER_BLOCK; k++) {
        const item = String(r[o.block - 1 + 2 * k] || '').trim();
        if (item) items.push({ office: o.name, item, qty: Number(r[o.block + 2 * k]) || 0 });
      }
    });
    return {
      row: i + 2,
      ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      who: String(r[1] || '').trim(), team: String(r[2] || '').trim(),
      action: String(r[3] || '').trim(), event: String(r[4] || '').trim(),
      notes: String(r[FORM_COLS - 1] || '').trim(), items,
    };
  }).filter(e => e.who || e.items.length);
}

// "Emma F, egf48" -> {name: 'Emma F', netid: 'egf48'}; "egf48" -> {name: 'egf48', netid: 'egf48'}.
function person_(who) {
  const m = who.match(/^(.*?),\s*([^,\s]+)\s*$/);
  if (m) return { name: m[1].trim() || m[2], netid: m[2].toLowerCase() };
  return { name: who, netid: /^[a-z]+\d+$/i.test(who) ? who.toLowerCase() : '' };
}

const personKey_ = who => person_(who).netid || who.toLowerCase();

// ---------- Names ----------

// Saved display names per NetID: {netid: {name, source: 'admin'|'directory', checked}}.
const NAMES_KEY = 'names:' + SPREADSHEET_ID;
const LOOKUPS_PER_LOAD = 5;
const RETRY_MS = 7 * 24 * 3600 * 1000;

function names_() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty(NAMES_KEY) || '{}');
}

// Re-reads the saved names under the script lock, so a directory lookup and an
// admin's edit can't overwrite each other.
function updateNames_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const names = names_();
    fn(names);
    PropertiesService.getScriptProperties().setProperty(NAMES_KEY, JSON.stringify(names));
  } finally {
    lock.releaseLock();
  }
}

// Name from the Google Workspace directory (People API advanced service):
// '' if not listed, null if the lookup itself failed (e.g. permission missing).
function directoryName_(netid) {
  if (typeof People === 'undefined') return null;
  try {
    const res = People.People.searchDirectoryPeople({
      query: netid,
      readMask: 'names,emailAddresses',
      sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE', 'DIRECTORY_SOURCE_TYPE_DOMAIN_CONTACT'],
    });
    const match = (res.people || []).find(p => (p.emailAddresses || []).some(e => e.value.toLowerCase().startsWith(netid + '@')));
    return (match && match.names && match.names[0] && match.names[0].displayName) || '';
  } catch (err) {
    console.warn('Directory lookup failed: ' + err);
    return null;
  }
}

// Run from the editor to (re)approve the directory permission and see a sample.
function testDirectoryLookup() {
  if (!runsAsViewer_()) throw new Error('Run this from the Apps Script editor.');
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, ['https://www.googleapis.com/auth/directory.readonly']);
  ['egf48', 'jsh253', 'vjp91'].forEach(id => {
    const name = directoryName_(id);
    Logger.log(id + ' -> ' + (name === null ? '(lookup failed, see warning above)' : name || '(not in directory)'));
  });
}

// NetID -> display name. An admin's entry wins, then the Google directory,
// then the name typed on the check-out site. Looks up a few new NetIDs per load.
// Directory names are left out of public data: anyone can check out under any
// NetID, which would otherwise turn the site into a directory search.
function resolveNames_(log, isPublic) {
  let saved = names_();
  const typed = {};
  const netids = new Set(Object.keys(saved));
  log.forEach(e => {
    const p = person_(e.who);
    if (!p.netid) return;
    netids.add(p.netid);
    if (p.name !== p.netid) typed[p.netid] = p.name;
  });
  const found = {};
  netids.forEach(id => {
    const s = saved[id];
    if ((s && (s.name || Date.now() - s.checked < RETRY_MS)) || Object.keys(found).length >= LOOKUPS_PER_LOAD) return;
    const name = directoryName_(id);
    // A failed lookup is retried in an hour; "not listed" in a week.
    const checked = name === null ? Date.now() - RETRY_MS + 3600 * 1000 : Date.now();
    found[id] = { name: name || '', source: 'directory', checked };
  });
  if (Object.keys(found).length) {
    // An admin may have set one of these names while the lookups ran; theirs wins.
    updateNames_(names => Object.keys(found).forEach(id => { if (!names[id] || names[id].source !== 'admin') names[id] = found[id]; }));
    saved = names_();
  }
  const out = {};
  netids.forEach(id => {
    const s = saved[id];
    const usable = s && s.name && !(isPublic && s.source === 'directory');
    const name = (usable && s.name) || typed[id];
    if (name) out[id] = name;
  });
  return out;
}

// Adds `name` (resolved) and `netid` to log entries and checked-out lines.
function withNames_(log, out, isPublic) {
  const names = resolveNames_(log, isPublic);
  const label = who => { const p = person_(who); return { name: names[p.netid] || p.name, netid: p.netid }; };
  log.forEach(e => Object.assign(e, label(e.who)));
  out.forEach(o => Object.assign(o, label(o.who)));
  return names;
}

function setName_(p) {
  const netid = String(p.netid || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]+$/.test(netid)) throw new Error('Enter a NetID like egf48.');
  const name = String(p.name || '').trim();
  // Clearing a name goes back to the directory / typed name.
  updateNames_(saved => { if (name) saved[netid] = { name, source: 'admin', checked: Date.now() }; else delete saved[netid]; });
}

function listNames_() {
  const { log } = snapshot_();
  const names = resolveNames_(log);
  const saved = names_();
  return Object.keys(names).sort().map(id => ({
    netid: id, name: names[id], source: saved[id] && saved[id].name ? saved[id].source : 'typed on site',
  }));
}

// Reusable items each person still has: take-outs minus check-ins.
function outstanding_(log, typeOf) {
  const map = new Map();
  log.forEach(e => {
    const sign = e.action === TAKE_OUT ? 1 : e.action === CHECK_IN ? -1 : 0;
    if (!sign) return;
    const key = personKey_(e.who);
    e.items.forEach(it => {
      const k = [key, it.office, it.item].join('|');
      if (!map.has(k)) map.set(k, { key, who: e.who, office: it.office, item: it.item, qty: 0, since: e.ts, event: '', team: '' });
      const o = map.get(k);
      if (sign > 0) {
        if (o.qty <= 0) o.since = e.ts;
        Object.assign(o, { who: e.who, event: e.event || o.event, team: e.team || o.team });
      }
      o.qty += sign * it.qty;
    });
  });
  return [...map.values()]
    .filter(o => o.qty > 0 && !/consum/i.test(typeOf(o.office, o.item)))
    .map(o => Object.assign(o, person_(o.who)));
}

function snapshot_() {
  const cfg = getConfig_();
  const offices = cfg.offices.map(readOffice_);
  const types = {};
  offices.forEach(o => o.items.forEach(it => { types[o.name + '|' + it.name] = it.type; }));
  const log = readForm_(offices);
  return { cfg, offices, log, out: outstanding_(log, (office, item) => types[office + '|' + item] || '') };
}

function getData_(fresh) {
  const cache = CacheService.getScriptCache();
  if (!fresh) {
    const hit = cache.get(CACHE_KEY);
    if (hit) return JSON.parse(hit);
  }
  const { cfg, offices, log, out } = snapshot_();
  const people = withNames_(log, out, true);
  const events = new Set();
  const teams = new Set();
  log.forEach(e => {
    if (e.event) events.add(e.event);
    if (e.team) teams.add(e.team);
  });
  const data = {
    config: { title: cfg.title, requireEvent: !!cfg.requireEvent, offices: offices.map(o => o.name) },
    items: Object.fromEntries(offices.map(o => [o.name, o.items])),
    out: out.sort((a, b) => (a.since < b.since ? 1 : -1)),
    recent: log.slice(-30).reverse().map(e => ({ ts: e.ts, who: e.who, name: e.name, netid: e.netid, action: e.action, event: e.event, items: e.items })),
    people,
    events: [...events],
    teams: [...teams],
  };
  try {
    cache.put(CACHE_KEY, JSON.stringify(data), CACHE_SECONDS);
  } catch (err) {
    // Over the 100 KB cache limit: serve uncached.
  }
  return data;
}

// Appends form rows: up to 10 items per office per row, like the form does.
function writeForm_(action, who, team, event, items, offices) {
  const blocks = Object.fromEntries(offices.map(o => [o.name, o.block]));
  const byOffice = {};
  items.forEach(it => (byOffice[it.office] = byOffice[it.office] || []).push(it));
  const rows = [];
  const now = new Date();
  for (let r = 0; ; r++) {
    const row = new Array(FORM_COLS).fill('');
    let any = false;
    Object.keys(byOffice).forEach(office => {
      byOffice[office].slice(r * ITEMS_PER_BLOCK, (r + 1) * ITEMS_PER_BLOCK).forEach((it, k) => {
        row[blocks[office] - 1 + 2 * k] = it.item;
        row[blocks[office] + 2 * k] = it.qty;
        any = true;
      });
    });
    if (!any) break;
    Object.assign(row, { 0: now, 1: who, 2: team || '', 3: action, 4: event || '', [FORM_COLS - 1]: 'Via check-out site' });
    rows.push(row);
  }
  const sh = formSheet_();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, FORM_COLS).setValues(rows);
}

// Same office + item twice in one request -> one line with the total.
function mergeItems_(items) {
  const merged = new Map();
  (items || []).forEach(it => {
    const line = { office: String(it.office || ''), item: String(it.item || '').trim(), qty: Math.floor(Number(it.qty)) || 0 };
    if (!line.item) throw new Error('Missing item');
    if (line.qty < 1) throw new Error('Quantity must be at least 1: ' + line.item);
    const k = line.office + '|' + line.item;
    if (merged.has(k)) merged.get(k).qty += line.qty; else merged.set(k, line);
  });
  if (!merged.size) throw new Error('No items selected');
  return [...merged.values()];
}

function checkout_(b) {
  const cfg = getConfig_();
  const name = String(b.name || '').trim();
  const netid = String(b.netid || '').trim().toLowerCase();
  if (!name) throw new Error('Missing name');
  if (!netid) throw new Error('Missing NetID');
  if (cfg.requireEvent && !String(b.event || '').trim()) throw new Error('Missing event');
  const items = mergeItems_(b.items);
  const offices = cfg.offices.map(readOffice_);
  items.forEach(it => {
    const office = offices.find(o => o.name === it.office);
    const found = office && office.items.find(x => x.name === it.item);
    if (!found) throw new Error(`${it.item} isn't listed in ${it.office} anymore. Reload and try again.`);
    const left = stockNumber_(found.qty);
    if (left !== null && it.qty > left) throw new Error(`Only ${left} ${it.item} left in ${it.office}.`);
  });
  writeForm_(TAKE_OUT, `${name}, ${netid}`, String(b.team || '').trim(), String(b.event || '').trim(), items, offices);
}

// b: {who, items: [{office, item, qty}]}; who is the "Name, Net ID" the items were taken out under.
function checkin_(b) {
  const who = String(b.who || '').trim();
  if (!who) throw new Error('Missing person');
  const items = mergeItems_(b.items);
  const { offices, out } = snapshot_();
  const key = personKey_(who);
  items.forEach(it => {
    const o = out.find(x => x.key === key && x.office === it.office && x.item === it.item);
    if (!o) throw new Error(`${who} doesn't have ${it.item} checked out.`);
    if (it.qty > o.qty) throw new Error(`${who} only has ${o.qty} ${it.item} checked out.`);
  });
  writeForm_(CHECK_IN, who, '', '', items, offices);
}

// ---------- Admin portal ----------

function sheetEditors_() {
  const ss = ss_();
  return ss.getEditors().concat(ss.getOwner() || []).map(u => u.getEmail().toLowerCase());
}

// Admins are the emails in the admin list (Settings). Until that list is
// set, every editor of the spreadsheet is an admin.
function requireAdmin_() {
  if (!runsAsViewer_()) throw new Error('Open the admin portal from the Admin link on the check-out site.');
  const email = Session.getActiveUser().getEmail().toLowerCase();
  if (!email) throw new Error('Sign in to Google to use the admin portal.');
  const admins = getConfig_().admins || [];
  if (admins.length) {
    if (!admins.includes(email)) throw new Error(email + ' is not an admin. Ask an admin to add you in Settings.');
    return email;
  }
  let editors;
  try {
    editors = sheetEditors_();
  } catch (err) {
    throw new Error(email + ' does not have edit access to the inventory sheet.');
  }
  if (!editors.includes(email)) throw new Error(email + ' does not have edit access to the inventory sheet.');
  return email;
}

// Adds an admin and, if needed, shares the spreadsheet with them as an
// editor (the portal runs as the signed-in user, so admins need edit access).
function addAdmin_(p, me) {
  const email = String(p.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a full email address.');
  const cfg = getConfig_();
  // The first admin added also locks in whoever is adding, so they aren't shut out.
  const admins = new Set((cfg.admins || []).length ? cfg.admins : [me]);
  admins.add(email);
  let shared = false;
  if (!sheetEditors_().includes(email)) {
    ss_().addEditor(email);
    shared = true;
  }
  cfg.admins = [...admins];
  writeConfig_(cfg);
  return { admins: cfg.admins, shared };
}

function removeAdmin_(p, me) {
  const email = String(p.email || '').trim().toLowerCase();
  if (email === me) throw new Error("You can't remove yourself. Ask another admin.");
  const cfg = getConfig_();
  cfg.admins = (cfg.admins || []).filter(a => a !== email);
  writeConfig_(cfg);
  return { admins: cfg.admins };
}

// Deletes one form row, after checking it is still the row the portal showed.
function deleteLog_(p) {
  const sh = formSheet_();
  const [ts, who] = sh.getRange(p.row, 1, 1, 2).getValues()[0];
  const iso = ts instanceof Date ? ts.toISOString() : String(ts);
  if (p.row < 2 || iso !== p.ts || String(who).trim() !== p.who) {
    throw new Error('That row changed in the sheet since you loaded it. Reload and try again.');
  }
  sh.deleteRow(p.row);
}

// Single entry point for Admin.html (google.script.run).
function adminApi(action, p) {
  const email = requireAdmin_();
  if (action === 'init') {
    const ss = ss_();
    return {
      email,
      config: getConfig_(),
      sheets: ss.getSheets().map(s => s.getName()).filter(n => n !== FORM_SHEET),
      sheetUrl: ss.getUrl(),
    };
  }
  if (action === 'log') {
    const { log, out } = snapshot_();
    withNames_(log, out);
    return { log: log.reverse(), out };
  }
  if (action === 'names') return listNames_();
  if (action === 'setName') {
    setName_(p);
    CacheService.getScriptCache().remove(CACHE_KEY);
    return listNames_();
  }
  if (action === 'rows') return readRows_(findOffice_(p.office));
  return withLock_(() => {
    switch (action) {
      case 'checkin': return checkin_(p);
      case 'deleteLog': return deleteLog_(p);
      case 'saveRow': return saveRow_(p);
      case 'addRow': return addRow_(p);
      case 'deleteRow': return deleteRow_(p);
      case 'saveConfig': return saveConfig_(p.config);
      case 'addAdmin': return addAdmin_(p, email);
      case 'removeAdmin': return removeAdmin_(p, email);
      default: throw new Error('Unknown action: ' + action);
    }
  });
}
