import { readSpreadsheetFile } from './xlsx-reader.js';

const $ = (id) => document.getElementById(id);
const el = {
  drop: $('drop'), fileInput: $('fileInput'), status: $('status'), clearBtn: $('clearBtn'),
  mapCard: $('mapCard'), colFirst: $('colFirst'), colLast: $('colLast'), colCong: $('colCong'), dateInput: $('dateInput'),
  listCard: $('listCard'), search: $('search'), congFilter: $('congFilter'), selVisible: $('selVisible'), selNone: $('selNone'),
  count: $('count'), list: $('list'), printCard: $('printCard'), printBtn: $('printBtn'),
  guides: $('guides'), empty: $('empty'), sheets: $('sheets'), installBtn: $('installBtn'),
  tplRadios: document.querySelectorAll('input[name="tpl"]'),
};

const CACHE_KEY = 'giving-envelopes.table.v1';
const TEMPLATE_KEY = 'giving-envelopes.template.v1';
const FUNDS = ['Freewill Giving', 'Special Offering', "Pastor’s Fund", 'Farm Fund', 'Building Fund', 'Kai Fund'];

const state = {
  table: [],          // all rows of the sheet (arrays of strings)
  headerRow: 0,
  fileName: '',
  givers: [],         // {id, cong, first, last}
  selected: new Set(),
  template: 'freewill', // 'freewill' | 'special'
};

/* ---------- helpers ---------- */
function setStatus(msg, kind = '') {
  el.status.textContent = msg;
  el.status.className = 'status ' + kind;
}

function nextSundayISO(from = new Date()) {
  const d = new Date(from.getFullYear(), from.getMonth(), from.getDate());
  d.setDate(d.getDate() + ((7 - d.getDay()) % 7)); // today if Sunday
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function formatDate(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  return m ? `${m[3]}/${m[2]}/${m[1]}` : '';
}

function findHeaderRow(table) {
  for (let i = 0; i < Math.min(table.length, 15); i++) {
    const cells = table[i].map((c) => c.toLowerCase());
    const hasFirst = cells.some((c) => /christian|first|given/.test(c));
    const hasLast = cells.some((c) => /surname|last|family/.test(c));
    if (hasFirst && hasLast) return i;
  }
  return 0;
}

function guessColumn(headers, regex, fallback) {
  const i = headers.findIndex((h) => regex.test(h.toLowerCase()));
  return i >= 0 ? i : fallback;
}

/* ---------- loading ---------- */
function fillSelect(select, headers, selected, allowNone) {
  select.innerHTML = '';
  if (allowNone) select.add(new Option('(none)', '-1'));
  headers.forEach((h, i) => select.add(new Option(h || `Column ${i + 1}`, String(i))));
  select.value = String(selected);
}

function applyTable(table, fileName, { fromCache = false } = {}) {
  state.table = table;
  state.fileName = fileName;
  state.headerRow = findHeaderRow(table);
  const headers = table[state.headerRow] || [];
  if (!headers.length || table.length <= state.headerRow + 1) {
    setStatus('That spreadsheet looks empty.', 'error');
    return;
  }
  fillSelect(el.colFirst, headers, guessColumn(headers, /christian|first|given/, 0), false);
  fillSelect(el.colLast, headers, guessColumn(headers, /surname|last|family/, Math.min(1, headers.length - 1)), false);
  fillSelect(el.colCong, headers, guessColumn(headers, /cong|group|number|no\.?$/, -1), true);
  el.mapCard.hidden = el.listCard.hidden = el.printCard.hidden = false;
  el.clearBtn.hidden = false;
  state.selected.clear();
  rebuildGivers();
  setStatus(
    (fromCache ? 'Using saved list from ' : 'Loaded ') + `“${fileName}” — ${state.givers.length} people.` +
      (fromCache ? ' Choose the file again to refresh it.' : ''),
    'ok'
  );
}

function rebuildGivers() {
  const f = +el.colFirst.value, l = +el.colLast.value, c = +el.colCong.value;
  const rows = state.table.slice(state.headerRow + 1);
  state.givers = rows
    .map((r, i) => ({ id: i, first: (r[f] || '').trim(), last: (r[l] || '').trim(), cong: c >= 0 ? (r[c] || '').trim() : '' }))
    .filter((g) => g.first || g.last);
  const keep = new Set([...state.selected].filter((id) => state.givers.some((g) => g.id === id)));
  state.selected = keep;
  // cong filter options
  const congs = [...new Set(state.givers.map((g) => g.cong).filter(Boolean))].sort((a, b) => (+a || 0) - (+b || 0) || a.localeCompare(b));
  el.congFilter.innerHTML = '';
  el.congFilter.add(new Option('All cong.', ''));
  congs.forEach((v) => el.congFilter.add(new Option('Cong. ' + v, v)));
  el.congFilter.hidden = congs.length === 0;
  renderList();
  renderSheets();
}

async function loadFile(file) {
  try {
    setStatus('Reading ' + file.name + '…');
    const table = await readSpreadsheetFile(file);
    applyTable(table, file.name);
    try { localStorage.setItem(CACHE_KEY, JSON.stringify({ table, fileName: file.name, savedAt: Date.now() })); } catch { /* storage unavailable */ }
  } catch (err) {
    console.error(err);
    setStatus(err.message || 'Could not read that file.', 'error');
  }
}

/* ---------- list ---------- */
function visibleGivers() {
  const q = el.search.value.trim().toLowerCase();
  const cong = el.congFilter.value;
  return state.givers.filter((g) => {
    if (cong && g.cong !== cong) return false;
    return !q || `${g.first} ${g.last} ${g.last} ${g.first}`.toLowerCase().includes(q);
  });
}

function renderList() {
  const shown = visibleGivers();
  el.list.innerHTML = '';
  const frag = document.createDocumentFragment();
  for (const g of shown) {
    const li = document.createElement('li');
    const label = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = state.selected.has(g.id);
    cb.addEventListener('change', () => {
      cb.checked ? state.selected.add(g.id) : state.selected.delete(g.id);
      updateCount();
      renderSheets();
    });
    const name = document.createElement('span');
    name.textContent = `${g.first} ${g.last}`.trim();
    label.append(cb, name);
    if (g.cong) {
      const n = document.createElement('span');
      n.className = 'num';
      n.textContent = g.cong;
      label.append(n);
    }
    li.append(label);
    frag.append(li);
  }
  el.list.append(frag);
  updateCount();
}

function updateCount() {
  const n = state.selected.size;
  const sheets = Math.ceil(n / 4);
  el.count.textContent = n ? `${n} selected · ${sheets} sheet${sheets === 1 ? '' : 's'}` : 'None selected';
  el.printBtn.disabled = n === 0;
}

/* ---------- envelopes ---------- */
function div(cls, text) {
  const d = document.createElement('div');
  if (cls) d.className = cls;
  if (text != null) d.textContent = text;
  return d;
}

function optsEl(spacer) {
  const o = div('opts' + (spacer ? ' spacer' : ''));
  for (const t of ['Cash', 'D/C']) {
    const s = document.createElement('span');
    const b = document.createElement('i');
    b.className = 'box';
    s.append(t + ' ', b);
    o.append(s);
  }
  return o;
}

const TITLE = 'MT ROSKILL EVANGELICAL FELLOWSHIP';

// "Label: ______ suffix" row used by the Special Offering envelope.
function soLine(label, lineClass, suffix) {
  const r = div('so-row');
  r.append(div('', label), div('so-line ' + lineClass));
  if (suffix) r.append(div('', suffix));
  return r;
}

// Name line: name centered, cong. number pinned to the far right.
function nameEl(cls, g) {
  const n = div(cls, `${g.first} ${g.last}`.trim());
  if (g.cong) n.append(div('cong-num', g.cong));
  return n;
}

function specialEnvelopeEl(g, dateText) {
  const env = div('envelope');
  const box = div('so-box');
  box.append(div('so-title', TITLE));
  box.append(nameEl('so-name', g));
  box.append(div('so-date', 'Date:  ' + dateText));
  box.append(soLine('Special Offering Total:', 'w1', '(Given Weekly)'));
  box.append(soLine('Special Offering Total:', 'w1', '(Given Today)'));
  const pay = div('so-row');
  pay.append(optsEl(false));
  box.append(pay);
  box.append(soLine('Overall Total:', 'w2', ''));
  env.append(box);
  return env;
}

function envelopeEl(g, dateText) {
  if (state.template === 'special') return specialEnvelopeEl(g, dateText);
  const env = div('envelope');
  env.append(div('env-title', TITLE));
  env.append(nameEl('env-name', g));
  env.append(div('env-date', 'Date:  ' + dateText));
  const rows = div('env-rows');
  for (const f of FUNDS) {
    const r = div('row');
    r.append(div('label', f + ':'), div('line'), optsEl(false));
    rows.append(r);
  }
  const total = div('row');
  total.append(div('label', 'Total:'), div('line'), optsEl(true));
  rows.append(total);
  env.append(rows);
  return env;
}

function renderSheets() {
  const picked = state.givers.filter((g) => state.selected.has(g.id));
  el.sheets.innerHTML = '';
  el.empty.hidden = picked.length > 0;
  const dateText = formatDate(el.dateInput.value);
  for (let i = 0; i < picked.length; i += 4) {
    const frame = div('sheet-frame');
    const sheet = div('sheet' + (el.guides.checked ? ' guides' : ''));
    picked.slice(i, i + 4).forEach((g) => sheet.append(envelopeEl(g, dateText)));
    frame.append(sheet);
    el.sheets.append(frame);
  }
}

/* ---------- events ---------- */
el.fileInput.addEventListener('change', () => { if (el.fileInput.files[0]) loadFile(el.fileInput.files[0]); el.fileInput.value = ''; });
['dragenter', 'dragover'].forEach((t) => el.drop.addEventListener(t, (e) => { e.preventDefault(); el.drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((t) => el.drop.addEventListener(t, (e) => { e.preventDefault(); el.drop.classList.remove('over'); }));
el.drop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) loadFile(f); });
el.drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); el.fileInput.click(); } });

el.tplRadios.forEach((r) => r.addEventListener('change', () => {
  if (!r.checked) return;
  state.template = r.value;
  try { localStorage.setItem(TEMPLATE_KEY, r.value); } catch { /* ignore */ }
  renderSheets();
}));

[el.colFirst, el.colLast, el.colCong].forEach((s) => s.addEventListener('change', rebuildGivers));
el.dateInput.addEventListener('change', renderSheets);
el.guides.addEventListener('change', renderSheets);
el.search.addEventListener('input', renderList);
el.congFilter.addEventListener('change', renderList);
el.selVisible.addEventListener('click', () => { visibleGivers().forEach((g) => state.selected.add(g.id)); renderList(); renderSheets(); });
el.selNone.addEventListener('click', () => { state.selected.clear(); renderList(); renderSheets(); });
el.printBtn.addEventListener('click', () => window.print());
el.clearBtn.addEventListener('click', () => {
  try { localStorage.removeItem(CACHE_KEY); } catch { /* ignore */ }
  state.table = []; state.givers = []; state.selected.clear();
  el.mapCard.hidden = el.listCard.hidden = el.printCard.hidden = el.clearBtn.hidden = true;
  renderSheets();
  setStatus('Saved list removed. Choose a spreadsheet to begin.');
});

/* ---------- PWA: install, file handler, service worker ---------- */
let deferredPrompt = null;
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); deferredPrompt = e; el.installBtn.hidden = false; });
el.installBtn.addEventListener('click', async () => {
  if (!deferredPrompt) return;
  deferredPrompt.prompt();
  await deferredPrompt.userChoice;
  deferredPrompt = null;
  el.installBtn.hidden = true;
});
window.addEventListener('appinstalled', () => { el.installBtn.hidden = true; });

if ('launchQueue' in window) {
  window.launchQueue.setConsumer(async (params) => {
    if (!params.files || !params.files.length) return;
    loadFile(await params.files[0].getFile());
  });
}
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
}

/* ---------- init ---------- */
el.dateInput.value = nextSundayISO();
try {
  const t = localStorage.getItem(TEMPLATE_KEY);
  if (t === 'freewill' || t === 'special') {
    state.template = t;
    el.tplRadios.forEach((r) => { r.checked = r.value === t; });
  }
} catch { /* ignore */ }
try {
  const cached = JSON.parse(localStorage.getItem(CACHE_KEY) || 'null');
  if (cached && Array.isArray(cached.table)) applyTable(cached.table, cached.fileName, { fromCache: true });
} catch { /* ignore */ }
renderSheets();
