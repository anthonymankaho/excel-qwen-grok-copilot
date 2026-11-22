const API = '';
const CONTEXT_CAP = 45_000; // Relay rejects messages over 50,000 chars; leave room for the prompt itself.
const MAX_CONTEXT_ROWS = 150;
const MAX_CONTEXT_COLS = 30;
const READ_MAX_CELLS = 2000; // One read-back batch. Full column width is kept; rows are trimmed to fit.
const PROFILE_MAX_ROWS = 2000; // Profiling reads whole ranges, so it is bounded separately from READ_MAX_CELLS.
const PROFILE_MAX_COLS = 60;
const PROFILE_DISTINCT = 12; // Sample values shown per column; enough to identify a grouping field.
const CELL_CHAR_LIMIT = 32767;
const KEY_STORE = 'eqgc.keys.v1';
const RELAY_STORE = 'eqgc.relay.v1';
const RISKY_FORMULA = /^=.*\b(?:WEBSERVICE|FILTERXML|RTD|REGISTER\.ID|CALL|DDE|INFO|HYPERLINK)\s*\(/i;
const state = { messages: [], lastAnswer: '', busy: false, officeReady: false, forceSheet: false, turn: 0, host: '', audit: [] };
const $ = id => document.getElementById(id);
let abortChat = null;

function storedKeys() {
  try { return JSON.parse(localStorage.getItem(KEY_STORE) || '{}'); } catch { return {}; }
}
function rememberKey(provider, key) {
  const all = storedKeys();
  if (key) all[provider] = key; else delete all[provider];
  try { localStorage.setItem(KEY_STORE, JSON.stringify(all)); } catch { /* storage unavailable: session-only */ }
}
let keysReady = null;
let sessionRelayToken = '';
function hostedRelay() {
  const host = globalThis.location?.hostname || '';
  return host !== '' && host !== '127.0.0.1' && host !== 'localhost';
}
function relayToken() {
  if (sessionRelayToken) return sessionRelayToken;
  try { return localStorage.getItem(RELAY_STORE) || ''; } catch { return ''; }
}
function rememberRelay(token) {
  sessionRelayToken = token || '';
  try {
    if (token) localStorage.setItem(RELAY_STORE, token);
    else localStorage.removeItem(RELAY_STORE);
  } catch { /* storage unavailable: the token stays for this pane session */ }
}
function ensureKeys() {
  if (hostedRelay()) return Promise.resolve();
  if (!keysReady) keysReady = (async () => {
    const stored = storedKeys();
    let allOk = true;
    for (const provider of ['qwen', 'grok']) {
      if (stored[provider]) { try { await api('configure', { provider, apiKey: stored[provider] }); } catch { allOk = false; } }
    }
    if (!allOk) keysReady = null; // relay was offline; retry on next question
  })();
  return keysReady;
}

async function api(path, body) {
  const headers = {};
  if (body) headers['Content-Type'] = 'application/json';
  const token = relayToken();
  if (token) headers.Authorization = `Bearer ${token}`;
  const response = await fetch(`${API}/api/${path}`, { method: body ? 'POST' : 'GET', headers, body: body ? JSON.stringify(body) : undefined });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `Request failed (${response.status})`);
  return data;
}

function toast(message) {
  const el = $('toast'); el.textContent = message; el.classList.remove('hidden');
  clearTimeout(toast.timer); toast.timer = setTimeout(() => el.classList.add('hidden'), 3200);
}

function setRelayDot(mode) {
  const dot = $('statusDot');
  if (!dot) return;
  dot.classList.toggle('warn', mode === 'warn');
  dot.classList.toggle('offline', mode === 'offline');
}

async function loadLimited(context, range, nullable) {
  if (nullable) {
    range.load('isNullObject');
    await context.sync();
    if (range.isNullObject) return null;
  }
  range.load(['rowCount', 'columnCount', 'address']);
  await context.sync();
  if (range.rowCount === 0 || range.columnCount === 0) return null;
  let view = range;
  let gridTruncated = false;
  const fullRows = range.rowCount;
  const fullCols = range.columnCount;
  if (Number.isFinite(fullRows) && Number.isFinite(fullCols) && (fullRows > MAX_CONTEXT_ROWS || fullCols > MAX_CONTEXT_COLS)) {
    const rows = Math.min(fullRows, MAX_CONTEXT_ROWS);
    const cols = Math.min(fullCols, MAX_CONTEXT_COLS);
    view = range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    gridTruncated = true;
  }
  view.load(['address', 'text', 'formulas', 'rowCount', 'columnCount']);
  await context.sync();
  return { address: view.address, text: view.text, formulas: view.formulas, rowCount: view.rowCount ?? fullRows, columnCount: view.columnCount ?? fullCols, gridTruncated, fullRows, fullCols };
}

async function readSelection() {
  return Excel.run(async context => loadLimited(context, context.workbook.getSelectedRange(), false));
}

async function readUsedRange() {
  return Excel.run(async context => loadLimited(context, context.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject(), true));
}

function flatCell(value) {
  return String(value ?? '').replace(/[\t\r\n]+/g, ' ');
}

function annotateRows(rows, formulas) {
  return (rows || []).map((row, i) => (row || []).map((value, j) => {
    const formula = formulas?.[i]?.[j];
    const shown = flatCell(value);
    return formula && String(formula).startsWith('=') ? `${shown} [${flatCell(formula)}]` : shown;
  }).join('\t')).join('\n');
}

async function buildWordContext() {
  const source = await Word.run(async context => {
    const selection = context.document.getSelection();
    const body = context.document.body;
    selection.load('text');
    body.load('text');
    await context.sync();
    if (String(selection.text || '').trim()) return { kind: 'selection', text: selection.text };
    if (String(body.text || '').trim()) return { kind: 'document', text: body.text };
    return null;
  });
  if (!source) return null;
  let text = source.text;
  let truncated = false;
  if (text.length > CONTEXT_CAP) {
    const marker = `\n[...truncated: ${text.length - CONTEXT_CAP} chars omitted — ask about a smaller range for full detail]`;
    text = text.slice(0, Math.max(0, CONTEXT_CAP - marker.length)) + marker;
    truncated = true;
  }
  return { header: `[Word ${source.kind}]`, text, address: source.kind === 'selection' ? 'Selection' : 'Document', kind: source.kind, truncated, chars: text.length };
}

async function buildContext() {
  if (!state.officeReady || !$('includeContext').checked) return null;
  if (state.host === 'Word') return buildWordContext();
  let source = null, kind = 'selection';
  if (!state.forceSheet) {
    const selection = await readSelection();
    if (selection && selection.rowCount * selection.columnCount > 1) source = selection;
  }
  if (!source) { source = await readUsedRange(); kind = 'used range'; }
  if (!source) return null;
  let text = annotateRows(source.text, source.formulas);
  if (!text) return null;
  let truncated = false;
  if (source.gridTruncated) {
    const rows = Math.min(source.fullRows || MAX_CONTEXT_ROWS, MAX_CONTEXT_ROWS);
    const cols = Math.min(source.fullCols || MAX_CONTEXT_COLS, MAX_CONTEXT_COLS);
    text += `\n[...truncated: read ${rows} of ${source.fullRows} rows and ${cols} of ${source.fullCols} columns — select a smaller range for the rest]`;
    truncated = true;
  }
  if (text.length > CONTEXT_CAP) {
    const marker = `\n[...truncated: ${text.length - CONTEXT_CAP} chars omitted — ask about a smaller range for full detail]`;
    text = text.slice(0, Math.max(0, CONTEXT_CAP - marker.length)) + marker;
    truncated = true;
  }
  return { header: `[Excel ${kind} ${source.address}]`, text, address: source.address, kind, truncated, chars: text.length };
}

async function refreshContextCard() {
  if (!state.officeReady) { $('selectionName').textContent = 'Open Excel or Word to read context'; return; }
  try {
    const ctx = await buildContext();
    if (!$('includeContext').checked) {
      $('selectionName').textContent = 'Context off';
      $('selectionPreview').textContent = 'Questions are sent without workbook data.';
      $('contextHint').textContent = 'Nothing will be sent';
      return;
    }
    if (!ctx) {
      $('selectionName').textContent = 'No workbook data';
      $('selectionPreview').textContent = '(The active sheet and selection are empty)';
      $('contextHint').textContent = 'Nothing will be sent';
      return;
    }
    $('selectionName').textContent = ctx.address;
    $('selectionPreview').textContent = ctx.text;
    $('contextHint').textContent = `Next question sends ${ctx.kind} · ${ctx.chars.toLocaleString()} chars${ctx.truncated ? ' · truncated' : ''} · formulas included`;
  } catch (error) { toast(`Could not read workbook context: ${error.message}`); }
}

function watchSelection() {
  const doc = Office.context?.document;
  if (!doc || typeof doc.addHandlerAsync !== 'function' || !Office.EventType?.DocumentSelectionChanged) return;
  doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    clearTimeout(watchSelection.timer);
    watchSelection.timer = setTimeout(refreshContextCard, 400);
  });
}

function renderMessage(role, content) {
  const conversation = $('conversation');
  conversation.querySelector('.welcome')?.remove();
  const bubble = document.createElement('div'); bubble.className = `message ${role}`;
  if (role === 'assistant') {
    const label = document.createElement('div'); label.className = 'message-label'; label.textContent = $('activeProvider').value.toUpperCase(); bubble.append(label);
  }
  const body = document.createElement('div'); body.textContent = content; bubble.append(body); conversation.append(bubble); conversation.scrollTop = conversation.scrollHeight;
  return bubble;
}

function columnIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function normalizeCell(cell) {
  if (typeof cell !== 'string') return '';
  const match = cell.trim().match(/^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/);
  if (!match) return '';
  const col = match[1].toUpperCase();
  const row = Number(match[2]);
  if (columnIndex(col) > 16384 || row > 1048576) return '';
  return col + String(row);
}

function normalizeSheet(name) {
  if (typeof name !== 'string') return '';
  let sheet = name.trim();
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  if (!sheet || sheet.length > 31 || /[:\\/?*[\]]/.test(sheet) || sheet.startsWith("'") || sheet.endsWith("'")) return '';
  return sheet;
}

function parseEdits(text) {
  const blocks = [...text.matchAll(/```(?:json)?\s*([\s\S]*?)```/gi)];
  for (let i = blocks.length - 1; i >= 0; i--) {
    try {
      const parsed = JSON.parse(blocks[i][1]);
      const edits = parsed?.edits;
      if (!Array.isArray(edits) || edits.length < 1 || edits.length > 200) continue;
      const clean = [];
      let skipped = 0;
      for (const edit of edits) {
        const sheet = normalizeSheet(edit?.sheet);
        const cell = normalizeCell(edit?.cell);
        const value = edit?.value;
        const valueOk = (typeof value === 'string' || (typeof value === 'number' && Number.isFinite(value))) && String(value).length <= 5000;
        if (!sheet || !cell || !valueOk) { skipped += 1; continue; }
        clean.push({ sheet, cell, value });
      }
      if (!clean.length) continue;
      const prose = (text.slice(0, blocks[i].index) + text.slice(blocks[i].index + blocks[i][0].length)).trim();
      return { prose, edits: clean, skipped };
    } catch { /* not an edits block; keep looking */ }
  }
  return { prose: text, edits: null, skipped: 0 };
}

function riskyEdit(value) {
  return typeof value === 'string' && RISKY_FORMULA.test(value.trim());
}

const CHART_TYPES = {
  column: 'ColumnClustered', bar: 'BarClustered', line: 'Line', pie: 'Pie',
  columnstacked: 'ColumnStacked', columnstacked100: 'ColumnStacked100',
  barstacked: 'BarStacked', barstacked100: 'BarStacked100',
  linestacked: 'LineStacked', linemarkers: 'LineMarkers', linemarkersstacked: 'LineMarkersStacked',
  area: 'Area', areastacked: 'AreaStacked', areastacked100: 'AreaStacked100',
  doughnut: 'Doughnut', doughnutexploded: 'DoughnutExploded', pieexploded: 'PieExploded',
  pieofpie: 'PieOfPie', barofpie: 'BarOfPie',
  scatter: 'XYScatter', xyscatter: 'XYScatter', xyscattersmooth: 'XYScatterSmooth',
  xyscatterlines: 'XYScatterLines', xyscattersmoothnomarkers: 'XYScatterSmoothNoMarkers',
  radar: 'Radar', radarmarkers: 'RadarMarkers', radarfilled: 'RadarFilled',
  bubble: 'Bubble', surface: 'Surface', stockhlc: 'StockHLC', stockohlc: 'StockOHLC'
};
const PIVOT_AGGREGATIONS = {
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min',
  countnumbers: 'CountNumbers', product: 'Product', stddev: 'StandardDeviation', variance: 'Variance'
};
const PIVOT_LAYOUTS = { tabular: 'Tabular', compact: 'Compact', outline: 'Outline' };
const BORDER_EDGES = {
  top: 'EdgeTop', bottom: 'EdgeBottom', left: 'EdgeLeft', right: 'EdgeRight',
  insidevertical: 'InsideVertical', insidehorizontal: 'InsideHorizontal',
  diagonaldown: 'DiagonalDown', diagonalup: 'DiagonalUp',
  edgetop: 'EdgeTop', edgebottom: 'EdgeBottom', edgeleft: 'EdgeLeft', edgeright: 'EdgeRight',
  all: 'all', outline: 'outline'
};
const BORDER_STYLES = {
  none: 'None', continuous: 'Continuous', solid: 'Continuous', dash: 'Dash', dashed: 'Dash',
  dashdot: 'DashDot', dashdotdot: 'DashDotDot', dot: 'Dot', dotted: 'Dot', double: 'Double', slantdashdot: 'SlantDashDot'
};
const BORDER_WEIGHTS = { hairline: 'Hairline', thin: 'Thin', medium: 'Medium', thick: 'Thick' };
const CF_OPERATORS = {
  between: 'Between', notbetween: 'NotBetween', equalto: 'EqualTo', equal: 'EqualTo', '==': 'EqualTo',
  notequalto: 'NotEqualTo', notequal: 'NotEqualTo', '!=': 'NotEqualTo',
  greaterthan: 'GreaterThan', '>': 'GreaterThan', lessthan: 'LessThan', '<': 'LessThan',
  greaterthanorequal: 'GreaterThanOrEqual', '>=': 'GreaterThanOrEqual',
  lessthanorequal: 'LessThanOrEqual', '<=': 'LessThanOrEqual'
};
const FILTER_ON = {
  values: 'Values', custom: 'Custom', topitems: 'TopItems', bottomitems: 'BottomItems',
  toppercent: 'TopPercent', bottompercent: 'BottomPercent', cellcolor: 'CellColor', fontcolor: 'FontColor'
};
const WORD_STYLES = { 'heading 1': 'Heading 1', 'heading 2': 'Heading 2', 'heading 3': 'Heading 3', title: 'Title', normal: 'Normal' };

function indexToColumn(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function rangeSize(addr) {
  const parts = String(addr).split(':');
  const start = parts[0].match(/^([A-Z]+)(\d+)$/);
  const end = (parts[1] || parts[0]).match(/^([A-Z]+)(\d+)$/);
  if (!start || !end) return { rows: 1, cols: 1 };
  return { rows: Math.abs(Number(end[2]) - Number(start[2])) + 1, cols: Math.abs(columnIndex(end[1]) - columnIndex(start[1])) + 1 };
}

function normalizeRange(input) {
  if (typeof input !== 'string') return '';
  const parts = input.trim().split(':');
  if (parts.length === 1) return normalizeCell(parts[0]);
  if (parts.length !== 2) return '';
  const start = normalizeCell(parts[0]);
  const end = normalizeCell(parts[1]);
  return start && end ? `${start}:${end}` : '';
}

function normalizeColumns(input) {
  if (typeof input !== 'string') return '';
  const match = input.trim().match(/^\$?([A-Za-z]{1,3})(?::\$?([A-Za-z]{1,3}))?$/);
  if (!match) return '';
  const start = match[1].toUpperCase();
  const end = (match[2] || start).toUpperCase();
  if (columnIndex(start) < 1 || columnIndex(start) > 16384 || columnIndex(end) < 1 || columnIndex(end) > 16384) return '';
  return columnIndex(start) <= columnIndex(end) ? `${start}:${end}` : `${end}:${start}`;
}

function boundedRange(input, limit = 1000) {
  const range = normalizeRange(input);
  if (!range) return '';
  const size = rangeSize(range);
  return size.rows * size.cols <= limit ? range : '';
}

function cellValueOk(value) {
  return (typeof value === 'string' || (typeof value === 'number' && Number.isFinite(value))) && String(value).length <= 5000;
}

function normalizeGrid(raw) {
  if (Array.isArray(raw.values)) {
    if (raw.values.length < 1 || raw.values.length > 30) return null;
    const grid = [];
    for (const row of raw.values) {
      if (!Array.isArray(row) || row.length < 1 || row.length > 15) return null;
      const clean = [];
      for (const value of row) {
        if (value == null) { clean.push(''); continue; }
        if (!cellValueOk(value)) return null;
        clean.push(value);
      }
      grid.push(clean);
    }
    return grid;
  }
  if (!cellValueOk(raw.value)) return null;
  return [[raw.value]];
}

function positiveInt(value, max) {
  const n = Number(value);
  return Number.isInteger(n) && n >= 1 && n <= max ? n : 0;
}

function colorOk(value) {
  return typeof value === 'string' && /^#[0-9A-Fa-f]{6}$/.test(value);
}

function normalizeAction(raw) {
  const op = typeof raw?.op === 'string' ? raw.op : '';
  if (op === 'replaceSelection' || op === 'insertText' || op === 'formatSelection' || op === 'styleSelection') {
    if (op === 'replaceSelection' || op === 'insertText') {
      const text = typeof raw.text === 'string' && raw.text.length <= 8000 ? raw.text : null;
      if (text == null || !text.length) return null;
      if (op === 'insertText' && raw.where != null && !['start', 'end'].includes(raw.where)) return null;
      return { host: 'word', op, text, where: raw.where === 'start' ? 'start' : 'end' };
    }
    if (op === 'formatSelection') {
      if (typeof raw.bold !== 'boolean' && typeof raw.italic !== 'boolean') return null;
      return { host: 'word', op, bold: raw.bold, italic: raw.italic };
    }
    const style = WORD_STYLES[String(raw.style || '').trim().toLowerCase()];
    return style ? { host: 'word', op, style } : null;
  }
  if (op === 'replaceAll' && !raw.sheet) {
    const text = typeof raw.text === 'string' ? raw.text : '';
    const replacement = typeof raw.replacement === 'string' ? raw.replacement : '';
    if (!text || text.length > 500 || replacement.length > 500) return null;
    return { host: 'word', op, text, replacement, matchCase: raw.matchCase === true, matchWholeWord: raw.matchWholeWord === true, bulk: true };
  }
  if (op === 'addSheet') {
    const name = normalizeSheet(raw.name);
    return name ? { host: 'excel', op, name } : null;
  }
  const sheet = normalizeSheet(raw.sheet);
  if (!sheet && op !== 'addSheet') return null;
  if (op === 'renameSheet') {
    const name = normalizeSheet(raw.name);
    return name && name !== sheet ? { host: 'excel', op, sheet, name } : null;
  }
  if (op === 'set') {
    const range = boundedRange(raw.range, 450);
    const grid = normalizeGrid(raw);
    if (!range || !grid) return null;
    const size = rangeSize(range);
    if (grid.length > size.rows || grid.some(row => row.length > size.cols)) return null;
    return { host: 'excel', op, sheet, range, grid };
  }
  if (op === 'format') {
    const range = boundedRange(raw.range, 1000);
    if (!range) return null;
    const action = { host: 'excel', op, sheet, range };
    if (typeof raw.bold === 'boolean') action.bold = raw.bold;
    if (typeof raw.italic === 'boolean') action.italic = raw.italic;
    if (raw.fill != null) { if (!colorOk(raw.fill)) return null; action.fill = raw.fill; }
    if (raw.fontColor != null) { if (!colorOk(raw.fontColor)) return null; action.fontColor = raw.fontColor; }
    if (raw.numberFormat != null) {
      if (typeof raw.numberFormat !== 'string' || raw.numberFormat.length > 40 || /[\r\n]/.test(raw.numberFormat)) return null;
      action.numberFormat = raw.numberFormat;
    }
    if (raw.align != null) {
      const align = { left: 'Left', center: 'Center', right: 'Right' }[String(raw.align).toLowerCase()];
      if (!align) return null;
      action.align = align;
    }
    if (typeof raw.wrap === 'boolean') action.wrap = raw.wrap;
    const keys = ['bold', 'italic', 'fill', 'fontColor', 'numberFormat', 'align', 'wrap'];
    return keys.some(key => action[key] != null) ? action : null;
  }
  if (op === 'clear' || op === 'sort' || op === 'chart' || op === 'table') {
    const range = boundedRange(raw.range, op === 'clear' ? 1000 : 1000);
    if (!range) return null;
    if (op === 'clear') return { host: 'excel', op, sheet, range };
    if (op === 'sort') {
      const start = columnIndex(range.split(':')[0].match(/^([A-Z]+)/)[1]);
      let key = -1;
      if (typeof raw.column === 'number' && raw.column >= 1 && raw.column <= 50) key = raw.column - 1;
      else if (typeof raw.column === 'string' && /^[A-Za-z]{1,3}$/.test(raw.column)) key = columnIndex(raw.column.toUpperCase()) - start;
      if (key < 0 || key >= rangeSize(range).cols) return null;
      return { host: 'excel', op, sheet, range, key, ascending: raw.ascending !== false, hasHeaders: raw.hasHeaders !== false };
    }
    if (op === 'chart') {
      const type = CHART_TYPES[String(raw.type || 'column').toLowerCase()];
      const title = raw.title == null ? '' : String(raw.title);
      if (!type || title.length > 120) return null;
      return { host: 'excel', op, sheet, range, type, title };
    }
    const name = typeof raw.name === 'string' ? raw.name.trim() : '';
    if (!/^[A-Za-z_][A-Za-z0-9_]{0,40}$/.test(name) || /^[A-Za-z]{1,3}[0-9]+$/.test(name)) return null;
    return { host: 'excel', op, sheet, range, name };
  }
  if (op === 'insertRows' || op === 'deleteRows') {
    const row = positiveInt(raw.row, 1048576);
    const count = positiveInt(raw.count || 1, 20);
    if (!row || !count || row + count - 1 > 1048576) return null;
    return { host: 'excel', op, sheet, row, count };
  }
  if (op === 'insertColumns') {
    const column = typeof raw.column === 'string' ? normalizeCell(`${raw.column}1`).replace(/1$/, '') : '';
    const count = positiveInt(raw.count || 1, 10);
    if (!/^[A-Z]{1,3}$/.test(column) || !count || columnIndex(column) + count - 1 > 16384) return null;
    return { host: 'excel', op, sheet, column, count };
  }
  if (op === 'autofit') {
    const range = /[0-9]/.test(String(raw.range || '')) ? boundedRange(raw.range, 2000) : normalizeColumns(raw.range);
    return range ? { host: 'excel', op, sheet, range } : null;
  }
  if (op === 'pivot') {
    // The source range is handed to Excel rather than read into the pane, so it may
    // be far larger than a normal read or write range.
    const range = boundedRange(raw.range, 1000000);
    if (!range) return null;
    const name = typeof raw.name === 'string' ? raw.name.trim() : '';
    if (!name || name.length > 60 || /[\r\n]/.test(name)) return null;
    const destinationSheet = normalizeSheet(raw.destinationSheet || name);
    const destinationCell = normalizeCell(raw.destinationCell || 'A3');
    if (!destinationSheet || !destinationCell) return null;
    const fieldList = value => {
      if (!Array.isArray(value) || value.length > 8) return null;
      const list = value.map(item => typeof item === 'string' ? item.trim() : '');
      return list.every(item => item && item.length <= 100) ? list : null;
    };
    const rows = fieldList(raw.rows || []);
    const columns = fieldList(raw.columns || []);
    const filters = fieldList(raw.filters || []);
    if (!rows || !columns || !filters || rows.length + columns.length < 1) return null;
    if (!Array.isArray(raw.values) || raw.values.length < 1 || raw.values.length > 8) return null;
    const values = [];
    for (const item of raw.values) {
      const field = typeof item?.field === 'string' ? item.field.trim() : '';
      if (!field || field.length > 100) return null;
      const summarizeBy = item.summarizeBy == null ? 'Sum' : PIVOT_AGGREGATIONS[String(item.summarizeBy).toLowerCase()];
      if (!summarizeBy) return null;
      const label = item.name == null ? field : String(item.name).trim();
      if (!label || label.length > 100) return null;
      values.push({ field, summarizeBy, name: label });
    }
    const layout = PIVOT_LAYOUTS[String(raw.layout || 'tabular').toLowerCase()];
    if (!layout) return null;
    return {
      host: 'excel', op, sheet, range, name, destinationSheet, destinationCell,
      rows, columns, filters, values, layout, grandTotals: raw.grandTotals !== false
    };
  }
  if (op === 'filter') {
    const range = boundedRange(raw.range, 100000);
    if (!range) return null;
    if (raw.clear === true) return { host: 'excel', op, sheet, range, clear: true, bulk: true };
    const start = columnIndex(range.split(':')[0].match(/^([A-Z]+)/)[1]);
    let key = -1;
    if (typeof raw.column === 'number' && raw.column >= 1 && raw.column <= 16384) key = raw.column - 1;
    else if (typeof raw.column === 'string' && /^[A-Za-z]{1,3}$/.test(raw.column.trim())) key = columnIndex(raw.column.trim().toUpperCase()) - start;
    if (key < 0 || key >= rangeSize(range).cols) return null;
    const filterOn = FILTER_ON[String(raw.filterOn || 'values').toLowerCase()];
    if (!filterOn) return null;
    const criteria = { filterOn };
    if (filterOn === 'Values') {
      if (!Array.isArray(raw.values) || raw.values.length < 1 || raw.values.length > 200) return null;
      const values = raw.values.map(value => String(value));
      if (values.some(value => !value || value.length > 200)) return null;
      criteria.values = values;
    } else if (filterOn === 'Custom') {
      const c1 = typeof raw.criterion1 === 'string' ? raw.criterion1.trim() : '';
      if (!c1 || c1.length > 100) return null;
      criteria.criterion1 = c1;
      if (raw.criterion2 != null && String(raw.criterion2).trim()) {
        const c2 = String(raw.criterion2).trim();
        if (c2.length > 100) return null;
        criteria.criterion2 = c2;
        criteria.operator = String(raw.operator || 'And').toLowerCase() === 'or' ? 'Or' : 'And';
      }
    } else if (filterOn === 'CellColor' || filterOn === 'FontColor') {
      if (!colorOk(raw.color)) return null;
      criteria.color = raw.color;
    } else {
      const count = positiveInt(raw.count, 1000);
      if (!count) return null;
      criteria.criterion1 = String(count);
    }
    return { host: 'excel', op, sheet, range, key, criteria };
  }
  if (op === 'conditionalFormat') {
    const range = boundedRange(raw.range, 100000);
    if (!range) return null;
    if (raw.clear === true) return { host: 'excel', op, sheet, range, clear: true, bulk: true };
    const operator = CF_OPERATORS[String(raw.operator || '').toLowerCase()];
    if (!operator) return null;
    const formula1 = raw.formula1 == null ? '' : String(raw.formula1).trim();
    if (!formula1 || formula1.length > 100) return null;
    const needsSecond = operator === 'Between' || operator === 'NotBetween';
    const formula2 = raw.formula2 == null ? '' : String(raw.formula2).trim();
    if (needsSecond && (!formula2 || formula2.length > 100)) return null;
    if (!needsSecond && formula2) return null;
    const action = { host: 'excel', op, sheet, range, operator, formula1 };
    if (needsSecond) action.formula2 = formula2;
    if (raw.fill != null) { if (!colorOk(raw.fill)) return null; action.fill = raw.fill; }
    if (raw.fontColor != null) { if (!colorOk(raw.fontColor)) return null; action.fontColor = raw.fontColor; }
    if (typeof raw.bold === 'boolean') action.bold = raw.bold;
    return action.fill || action.fontColor || action.bold != null ? action : null;
  }
  if (op === 'borders') {
    const range = boundedRange(raw.range, 100000);
    if (!range) return null;
    const edge = BORDER_EDGES[String(raw.edge || 'all').toLowerCase()];
    const style = BORDER_STYLES[String(raw.style || 'continuous').toLowerCase()];
    const weight = BORDER_WEIGHTS[String(raw.weight || 'thin').toLowerCase()];
    if (!edge || !style || !weight) return null;
    const action = { host: 'excel', op, sheet, range, edge, style, weight };
    if (raw.color != null) { if (!colorOk(raw.color)) return null; action.color = raw.color; }
    return action;
  }
  if (op === 'freeze') {
    if (raw.unfreeze === true) return { host: 'excel', op, sheet, unfreeze: true };
    const rows = raw.rows == null ? 0 : Number(raw.rows);
    const columns = raw.columns == null ? 0 : Number(raw.columns);
    if (!Number.isInteger(rows) || !Number.isInteger(columns) || rows < 0 || columns < 0 || rows > 1048576 || columns > 16384) return null;
    if (!rows && !columns) return null;
    return { host: 'excel', op, sheet, rows, columns };
  }
  if (op === 'replaceAll') {
    const text = typeof raw.text === 'string' ? raw.text : '';
    const replacement = typeof raw.replacement === 'string' ? raw.replacement : '';
    if (!text || text.length > 500 || replacement.length > 500) return null;
    return { host: 'excel', op, sheet, text, replacement, completeMatch: raw.completeMatch === true, matchCase: raw.matchCase === true, bulk: true };
  }
  if (op === 'dataValidation') {
    const range = boundedRange(raw.range, 100000);
    if (!range) return null;
    if (raw.clear === true) return { host: 'excel', op, sheet, range, clear: true, bulk: true };
    if (!Array.isArray(raw.list) || raw.list.length < 1 || raw.list.length > 200) return null;
    const items = raw.list.map(item => String(item).trim());
    if (items.some(item => !item || item.length > 100 || item.includes(','))) return null;
    return { host: 'excel', op, sheet, range, list: items, ignoreBlanks: raw.ignoreBlanks !== false };
  }
  if (op === 'namedRange') {
    const name = typeof raw.name === 'string' ? raw.name.trim() : '';
    if (!/^[A-Za-z_][A-Za-z0-9_.]{0,60}$/.test(name) || /^[A-Za-z]{1,3}[0-9]+$/.test(name)) return null;
    const range = normalizeRange(raw.range);
    if (!range) return null;
    return { host: 'excel', op, sheet, name, range };
  }
  if (op === 'hide') {
    const target = String(raw.target || '').toLowerCase();
    if (!['rows', 'columns', 'sheet'].includes(target)) return null;
    const hidden = raw.hidden !== false;
    if (target === 'sheet') return { host: 'excel', op, sheet, target, hidden };
    // Whole-row ("5:10") and whole-column ("D:E") addresses are the natural way to hide.
    const text = String(raw.range || '').trim();
    const range = target === 'rows' && /^\d+:\d+$/.test(text)
      ? text
      : /[0-9]/.test(text) ? boundedRange(text, 100000) : normalizeColumns(text);
    if (!range) return null;
    return { host: 'excel', op, sheet, target, range, hidden };
  }
  if (op === 'tableColumn') {
    const table = typeof raw.table === 'string' ? raw.table.trim() : '';
    const name = typeof raw.name === 'string' ? raw.name.trim() : '';
    if (!table || table.length > 60 || !name || name.length > 60) return null;
    const formula = typeof raw.formula === 'string' ? raw.formula.trim() : '';
    if (formula && (!formula.startsWith('=') || formula.length > 500)) return null;
    if (!formula) {
      if (!Array.isArray(raw.values) || raw.values.length < 1 || raw.values.length > 5000) return null;
      if (raw.values.some(value => !cellValueOk(value))) return null;
      return { host: 'excel', op, sheet, table, name, formula: '', values: raw.values.map(value => [value]) };
    }
    return { host: 'excel', op, sheet, table, name, formula, values: null };
  }
  return null;
}

function normalizeActions(list) {
  if (!Array.isArray(list) || list.length < 1 || list.length > 40) return { actions: [], skipped: Array.isArray(list) ? list.length : 0 };
  const actions = [];
  let skipped = 0;
  for (const raw of list) {
    const action = normalizeAction(raw);
    if (action) actions.push(action);
    else skipped += 1;
  }
  return { actions, skipped };
}

function parseRevision(text) {
  const blocks = [...text.matchAll(/```(?:json)?\s*([\s\S]*?)```/gi)];
  for (let i = blocks.length - 1; i >= 0; i--) {
    try {
      const parsed = JSON.parse(blocks[i][1]);
      if (!Array.isArray(parsed?.actions)) continue;
      const { actions, skipped } = normalizeActions(parsed.actions);
      if (!actions.length) continue;
      const prose = (text.slice(0, blocks[i].index) + text.slice(blocks[i].index + blocks[i][0].length)).trim();
      return { prose, actions, skipped, edits: null };
    } catch { /* not an actions block */ }
  }
  const edits = parseEdits(text);
  return { prose: edits.prose, edits: edits.edits, skipped: edits.skipped, actions: null };
}

function actionIsRisky(action) {
  if (action.op === 'set') return action.grid.some(row => row.some(riskyEdit));
  if (action.op === 'tableColumn') return Boolean(action.formula && riskyEdit(action.formula));
  return false;
}

function describeAction(action) {
  if (action.op === 'set') {
    const preview = action.grid.length === 1 && action.grid[0].length === 1 ? String(action.grid[0][0]) : `${action.grid.length}×${action.grid[0].length} values`;
    return `${action.sheet}!${action.range} → ${preview}`;
  }
  if (action.op === 'format') return `Format ${action.sheet}!${action.range}`;
  if (action.op === 'clear') return `Clear ${action.sheet}!${action.range}`;
  if (action.op === 'insertRows') return `Insert ${action.count} row${action.count === 1 ? '' : 's'} at row ${action.row} on ${action.sheet}`;
  if (action.op === 'deleteRows') return `Delete ${action.count} row${action.count === 1 ? '' : 's'} at row ${action.row} on ${action.sheet}`;
  if (action.op === 'insertColumns') return `Insert ${action.count} column${action.count === 1 ? '' : 's'} at ${action.column} on ${action.sheet}`;
  if (action.op === 'addSheet') return `Add sheet ${action.name}`;
  if (action.op === 'renameSheet') return `Rename ${action.sheet} to ${action.name}`;
  if (action.op === 'autofit') return `Autofit ${action.sheet}!${action.range}`;
  if (action.op === 'sort') return `Sort ${action.sheet}!${action.range}`;
  if (action.op === 'chart') return `Chart ${action.sheet}!${action.range} (${action.type})`;
  if (action.op === 'table') return `Table ${action.name} on ${action.sheet}!${action.range}`;
  if (action.op === 'pivot') {
    const values = action.values.map(value => `${value.summarizeBy} of ${value.field}`).join(', ');
    const rows = action.rows.join(', ') || 'none';
    const columns = action.columns.join(', ') || 'none';
    return `PivotTable ${action.name} at ${action.destinationSheet}!${action.destinationCell} from ${action.sheet}!${action.range} · rows ${rows} · columns ${columns} · ${values}`;
  }
  if (action.op === 'filter') {
    if (action.clear) return `Clear all filters on ${action.sheet} (not reversible by Undo)`;
    const detail = action.criteria.filterOn === 'Values'
      ? `values ${action.criteria.values.join(', ')}`
      : action.criteria.filterOn === 'Custom'
        ? `${action.criteria.criterion1}${action.criteria.criterion2 ? ` ${action.criteria.operator} ${action.criteria.criterion2}` : ''}`
        : action.criteria.color || action.criteria.criterion1 || '';
    return `Filter ${action.sheet}!${action.range} column ${indexToColumn(columnIndex(action.range.split(':')[0].match(/^([A-Z]+)/)[1]) + action.key)} · ${action.criteria.filterOn} ${detail}`;
  }
  if (action.op === 'conditionalFormat') {
    if (action.clear) return `Clear conditional formats on ${action.sheet}!${action.range} (not reversible by Undo)`;
    return `Highlight ${action.sheet}!${action.range} where ${action.operator} ${action.formula1}${action.formula2 ? ` and ${action.formula2}` : ''}`;
  }
  if (action.op === 'borders') return `Borders ${action.edge} on ${action.sheet}!${action.range} (${action.style}, ${action.weight})`;
  if (action.op === 'freeze') {
    if (action.unfreeze) return `Unfreeze panes on ${action.sheet}`;
    return `Freeze ${action.rows} row(s) and ${action.columns} column(s) on ${action.sheet}`;
  }
  if (action.op === 'replaceAll') {
    const where = action.host === 'word' ? 'in the document' : `on ${action.sheet}`;
    return `Replace every "${action.text}" with "${action.replacement}" ${where} (not reversible by Undo)`;
  }
  if (action.op === 'dataValidation') {
    if (action.clear) return `Clear data validation on ${action.sheet}!${action.range} (not reversible by Undo)`;
    return `Dropdown on ${action.sheet}!${action.range}: ${action.list.join(', ')}`;
  }
  if (action.op === 'namedRange') return `Name ${action.sheet}!${action.range} as ${action.name}`;
  if (action.op === 'hide') {
    const verb = action.hidden ? 'Hide' : 'Unhide';
    return action.target === 'sheet' ? `${verb} sheet ${action.sheet}` : `${verb} ${action.target} ${action.sheet}!${action.range}`;
  }
  if (action.op === 'tableColumn') return `Add column ${action.name} to table ${action.table}${action.formula ? ` = ${action.formula}` : ''}`;
  if (action.op === 'replaceSelection') return 'Replace the Word selection';
  if (action.op === 'insertText') return `Insert text at the ${action.where} of the document`;
  if (action.op === 'formatSelection') return 'Format the Word selection';
  if (action.op === 'styleSelection') return `Apply Word style ${action.style}`;
  return action.op;
}

function writeCell(cell, value) {
  if (typeof value === 'string' && value.startsWith('=')) cell.formulas = [[value]];
  else cell.values = [[value]];
}

function sheetMissing(sheet) {
  return !!(sheet && sheet.isNullObject);
}

async function applyExcelActions(actions) {
  if (typeof Excel === 'undefined') throw new Error('These changes need Excel.');
  return Excel.run(async context => {
    const wanted = new Map();
    for (const action of actions) {
      if (!action.sheet || wanted.has(action.sheet)) continue;
      const sheet = context.workbook.worksheets.getItemOrNullObject(action.sheet);
      sheet.load('isNullObject');
      wanted.set(action.sheet, sheet);
    }
    await context.sync();
    const createdNames = new Set(actions.filter(action => action.op === 'addSheet').map(action => action.name));
    for (const [name, sheet] of wanted) {
      if (sheetMissing(sheet) && !createdNames.has(name)) throw new Error(`Sheet “${name}” is not in this workbook. Nothing was written.`);
    }
    const undo = [];
    for (const action of actions) {
      if (action.op === 'addSheet') {
        const created = context.workbook.worksheets.add(action.name);
        wanted.set(action.name, created);
        undo.push({ host: 'excel', op: 'addSheet', name: action.name });
        continue;
      }
      const sheet = action.sheet ? wanted.get(action.sheet) : null;
      if (action.sheet && (!sheet || sheetMissing(sheet))) throw new Error(`Sheet “${action.sheet}” is not in this workbook. Nothing was written.`);
      if (action.op === 'renameSheet') {
        sheet.name = action.name;
        undo.push({ host: 'excel', op: 'renameSheet', sheet: action.name, name: action.sheet });
        continue;
      }
      if (action.op === 'set') {
        const range = sheet.getRange(action.range);
        const cells = [];
        for (let r = 0; r < action.grid.length; r++) {
          for (let c = 0; c < action.grid[r].length; c++) {
            const cell = action.grid.length === 1 && action.grid[0].length === 1 ? range : range.getCell(r, c);
            cell.load(['values', 'formulas']);
            cells.push({ r, c, cell });
          }
        }
        await context.sync();
        const before = cells.map(({ r, c, cell }) => ({
          address: action.grid.length === 1 && action.grid[0].length === 1 ? action.range : cellAddress(action.range, r, c),
          formula: cell.formulas?.[0]?.[0],
          value: cell.values?.[0]?.[0]
        }));
        cells.forEach(({ r, c, cell }) => writeCell(cell, action.grid[r][c]));
        undo.push({ host: 'excel', op: 'set', sheet: action.sheet, cells: before });
        continue;
      }
      if (action.op === 'format') {
        const range = sheet.getRange(action.range);
        range.format?.font?.load?.(['bold', 'italic', 'color']);
        range.format?.fill?.load?.('color');
        range.format?.load?.(['horizontalAlignment', 'wrapText']);
        range.load('numberFormat');
        await context.sync();
        undo.push({
          host: 'excel', op: 'format', sheet: action.sheet, range: action.range,
          bold: range.format?.font?.bold, italic: range.format?.font?.italic, fontColor: range.format?.font?.color,
          fill: range.format?.fill?.color, align: range.format?.horizontalAlignment, wrap: range.format?.wrapText,
          numberFormat: range.numberFormat
        });
        if (action.bold != null) range.format.font.bold = action.bold;
        if (action.italic != null) range.format.font.italic = action.italic;
        if (action.fontColor) range.format.font.color = action.fontColor;
        if (action.fill) range.format.fill.color = action.fill;
        if (action.align) range.format.horizontalAlignment = action.align;
        if (action.wrap != null) range.format.wrapText = action.wrap;
        if (action.numberFormat) {
          const size = rangeSize(action.range);
          range.numberFormat = Array.from({ length: size.rows }, () => Array.from({ length: size.cols }, () => action.numberFormat));
        }
        continue;
      }
      if (action.op === 'clear') {
        const range = sheet.getRange(action.range);
        range.load(['values', 'formulas', 'rowCount', 'columnCount']);
        await context.sync();
        undo.push({ host: 'excel', op: 'clear', sheet: action.sheet, range: action.range, formulas: range.formulas, values: range.values });
        range.clear();
        continue;
      }
      if (action.op === 'insertRows') {
        sheet.getRange(`${action.row}:${action.row + action.count - 1}`).insert('Down');
        undo.push({ host: 'excel', op: 'deleteRows', sheet: action.sheet, row: action.row, count: action.count });
        continue;
      }
      if (action.op === 'deleteRows') {
        // Capture the full used width, not just A:H, so Undo can restore every column.
        const used = sheet.getUsedRangeOrNullObject();
        used.load(['isNullObject', 'columnCount']);
        await context.sync();
        const width = Math.max(8, Math.min(used.isNullObject ? 8 : used.columnCount, 200));
        const addr = `A${action.row}:${indexToColumn(width)}${action.row + action.count - 1}`;
        const range = sheet.getRange(addr);
        range.load(['values', 'formulas']);
        await context.sync();
        undo.push({ host: 'excel', op: 'restoreRows', sheet: action.sheet, row: action.row, count: action.count, formulas: range.formulas, values: range.values });
        sheet.getRange(`${action.row}:${action.row + action.count - 1}`).delete('Up');
        continue;
      }
      if (action.op === 'insertColumns') {
        const end = indexToColumn(columnIndex(action.column) + action.count - 1);
        sheet.getRange(`${action.column}:${end}`).insert('Right');
        undo.push({ host: 'excel', op: 'deleteColumns', sheet: action.sheet, column: action.column, count: action.count });
        continue;
      }
      if (action.op === 'autofit') {
        const range = sheet.getRange(action.range);
        // Capture each width and height first: autofit itself reports nothing to undo.
        const clean = action.range.replace(/\$/g, '').toUpperCase();
        const columnOnly = /^[A-Z]+:[A-Z]+$/.test(clean);
        const match = clean.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
        const firstCol = columnOnly ? columnIndex(clean.split(':')[0]) : match ? columnIndex(match[1]) : 0;
        const lastCol = columnOnly ? columnIndex(clean.split(':')[1]) : match ? columnIndex(match[3] || match[1]) : 0;
        const firstRow = !columnOnly && match ? Number(match[2]) : 0;
        const lastRow = !columnOnly && match ? Number(match[4] || match[2]) : 0;
        const widths = [];
        const heights = [];
        for (let c = firstCol; c > 0 && c <= lastCol && c - firstCol < 200; c++) {
          const one = sheet.getRange(indexToColumn(c) + '1');
          one.format.load('columnWidth');
          widths.push({ column: indexToColumn(c), one });
        }
        for (let r = firstRow; r > 0 && r <= lastRow && r - firstRow < 500; r++) {
          const one = sheet.getRange('A' + r);
          one.format.load('rowHeight');
          heights.push({ row: r, one });
        }
        await context.sync();
        undo.push({
          host: 'excel', op: 'autofit', sheet: action.sheet,
          widths: widths.map(item => ({ column: item.column, width: item.one.format.columnWidth })),
          heights: heights.map(item => ({ row: item.row, height: item.one.format.rowHeight }))
        });
        range.format.autofitColumns();
        range.format.autofitRows();
        continue;
      }
      if (action.op === 'sort') {
        const range = sheet.getRange(action.range);
        range.load(['values', 'formulas']);
        await context.sync();
        undo.push({ host: 'excel', op: 'restoreGrid', sheet: action.sheet, range: action.range, formulas: range.formulas, values: range.values });
        range.sort.apply([{ key: action.key, ascending: action.ascending }], false, action.hasHeaders);
        continue;
      }
      if (action.op === 'chart') {
        const name = `Revise${Math.random().toString(36).slice(2, 8)}`;
        const chart = sheet.charts.add(action.type, sheet.getRange(action.range), 'Auto');
        chart.name = name;
        if (action.title) chart.title.text = action.title;
        undo.push({ host: 'excel', op: 'chart', sheet: action.sheet, name });
        continue;
      }
      if (action.op === 'table') {
        const table = sheet.tables.add(sheet.getRange(action.range), true);
        table.name = action.name;
        undo.push({ host: 'excel', op: 'table', sheet: action.sheet, name: action.name });
        continue;
      }
      if (action.op === 'pivot') {
        let dest = wanted.get(action.destinationSheet);
        if (!dest) {
          dest = context.workbook.worksheets.getItemOrNullObject(action.destinationSheet);
          dest.load('isNullObject');
          await context.sync();
        }
        let createdSheet = false;
        if (!dest || sheetMissing(dest)) {
          dest = context.workbook.worksheets.add(action.destinationSheet);
          wanted.set(action.destinationSheet, dest);
          createdSheet = true;
        }
        const pivot = sheet.pivotTables.add(action.name, sheet.getRange(action.range), dest.getRange(action.destinationCell));
        pivot.layout.layoutType = action.layout;
        pivot.layout.showRowGrandTotals = action.grandTotals;
        pivot.layout.showColumnGrandTotals = action.grandTotals;
        await context.sync();
        for (const field of action.rows) pivot.rowHierarchies.add(pivot.hierarchies.getItem(field));
        for (const field of action.columns) pivot.columnHierarchies.add(pivot.hierarchies.getItem(field));
        for (const field of action.filters) pivot.filterHierarchies.add(pivot.hierarchies.getItem(field));
        for (const value of action.values) {
          const data = pivot.dataHierarchies.add(pivot.hierarchies.getItem(value.field));
          data.summarizeBy = value.summarizeBy;
          data.name = value.name;
        }
        pivot.refresh();
        undo.push({ host: 'excel', op: 'pivot', sheet: action.destinationSheet, name: action.name, createdSheet });
        continue;
      }
      if (action.op === 'filter') {
        if (action.clear) sheet.autoFilter.clearCriteria();
        else {
          sheet.autoFilter.apply(sheet.getRange(action.range), action.key, action.criteria);
          undo.push({ host: 'excel', op: 'filter', sheet: action.sheet, key: action.key });
        }
        continue;
      }
      if (action.op === 'conditionalFormat') {
        const range = sheet.getRange(action.range);
        if (action.clear) { range.conditionalFormats.clearAll(); continue; }
        const cf = range.conditionalFormats.add('CellValue');
        cf.cellValue.rule = action.formula2
          ? { operator: action.operator, formula1: action.formula1, formula2: action.formula2 }
          : { operator: action.operator, formula1: action.formula1 };
        if (action.fill) cf.cellValue.format.fill.color = action.fill;
        if (action.fontColor) cf.cellValue.format.font.color = action.fontColor;
        if (action.bold != null) cf.cellValue.format.font.bold = action.bold;
        cf.load('id');
        await context.sync();
        undo.push({ host: 'excel', op: 'conditionalFormat', sheet: action.sheet, range: action.range, id: cf.id });
        continue;
      }
      if (action.op === 'borders') {
        const range = sheet.getRange(action.range);
        const edges = action.edge === 'all'
          ? ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideVertical', 'InsideHorizontal']
          : action.edge === 'outline' ? ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] : [action.edge];
        const targets = edges.map(edge => {
          const border = range.format.borders.getItem(edge);
          border.load(['style', 'weight', 'color']);
          return { edge, border };
        });
        await context.sync();
        undo.push({
          host: 'excel', op: 'borders', sheet: action.sheet, range: action.range,
          edges: targets.map(({ edge, border }) => ({ edge, style: border.style, weight: border.weight, color: border.color }))
        });
        for (const { border } of targets) {
          border.style = action.style;
          border.weight = action.weight;
          if (action.color) border.color = action.color;
        }
        continue;
      }
      if (action.op === 'freeze') {
        const location = sheet.freezePanes.getLocationOrNullObject();
        location.load(['isNullObject', 'address']);
        await context.sync();
        undo.push({ host: 'excel', op: 'freeze', sheet: action.sheet, address: location.isNullObject ? '' : location.address });
        if (action.unfreeze) sheet.freezePanes.unfreeze();
        else {
          if (action.rows) sheet.freezePanes.freezeRows(action.rows);
          if (action.columns) sheet.freezePanes.freezeColumns(action.columns);
        }
        continue;
      }
      if (action.op === 'replaceAll') {
        sheet.replaceAll(action.text, action.replacement, { completeMatch: action.completeMatch, matchCase: action.matchCase });
        continue;
      }
      if (action.op === 'dataValidation') {
        const range = sheet.getRange(action.range);
        if (action.clear) { range.dataValidation.clear(); continue; }
        range.dataValidation.rule = { list: { inCellDropDown: true, source: action.list.join(',') } };
        range.dataValidation.ignoreBlanks = action.ignoreBlanks;
        undo.push({ host: 'excel', op: 'dataValidation', sheet: action.sheet, range: action.range });
        continue;
      }
      if (action.op === 'namedRange') {
        context.workbook.names.add(action.name, sheet.getRange(action.range));
        undo.push({ host: 'excel', op: 'namedRange', name: action.name });
        continue;
      }
      if (action.op === 'hide') {
        if (action.target === 'sheet') {
          sheet.load('visibility');
          await context.sync();
          undo.push({ host: 'excel', op: 'hideSheet', sheet: action.sheet, visibility: sheet.visibility });
          sheet.visibility = action.hidden ? 'Hidden' : 'Visible';
          continue;
        }
        const range = sheet.getRange(action.range);
        const property = action.target === 'rows' ? 'rowHidden' : 'columnHidden';
        range.load(property);
        await context.sync();
        undo.push({ host: 'excel', op: 'hide', sheet: action.sheet, range: action.range, property, previous: range[property] });
        range[property] = action.hidden;
        continue;
      }
      if (action.op === 'tableColumn') {
        const table = sheet.tables.getItem(action.table);
        const values = action.formula ? [[action.formula]] : action.values;
        table.columns.add(undefined, values, action.name);
        undo.push({ host: 'excel', op: 'tableColumn', sheet: action.sheet, table: action.table, name: action.name });
      }
    }
    await context.sync();
    return undo;
  });
}

function cellAddress(range, rowOffset, colOffset) {
  const top = range.split(':')[0].match(/^([A-Z]+)(\d+)$/);
  return indexToColumn(columnIndex(top[1]) + colOffset) + (Number(top[2]) + rowOffset);
}

async function undoExcel(records) {
  await Excel.run(async context => {
    for (const rec of [...records].reverse()) {
      const sheet = rec.sheet ? context.workbook.worksheets.getItem(rec.sheet) : null;
      if (rec.op === 'set') {
        for (const cell of rec.cells) {
          const range = context.workbook.worksheets.getItem(rec.sheet).getRange(cell.address);
          if (String(cell.formula || '').startsWith('=')) range.formulas = [[cell.formula]];
          else if (cell.value === '' || cell.value == null) range.clear();
          else range.values = [[cell.value]];
        }
      } else if (rec.op === 'format') {
        const range = sheet.getRange(rec.range);
        if (rec.bold != null) range.format.font.bold = rec.bold;
        if (rec.italic != null) range.format.font.italic = rec.italic;
        if (rec.fontColor) range.format.font.color = rec.fontColor;
        if (rec.fill) range.format.fill.color = rec.fill;
        if (rec.align) range.format.horizontalAlignment = rec.align;
        if (rec.wrap != null) range.format.wrapText = rec.wrap;
        if (rec.numberFormat) range.numberFormat = rec.numberFormat;
      } else if (rec.op === 'clear' || rec.op === 'restoreGrid') {
        const range = sheet.getRange(rec.range);
        if (rec.formulas) range.formulas = rec.formulas;
        else if (rec.values) range.values = rec.values;
      } else if (rec.op === 'deleteRows') {
        sheet.getRange(`${rec.row}:${rec.row + rec.count - 1}`).delete('Up');
      } else if (rec.op === 'restoreRows') {
        sheet.getRange(`${rec.row}:${rec.row + rec.count - 1}`).insert('Down');
        const range = sheet.getRange(`A${rec.row}:${indexToColumn(8)}${rec.row + rec.count - 1}`);
        if (rec.formulas) range.formulas = rec.formulas;
      } else if (rec.op === 'deleteColumns') {
        const end = indexToColumn(columnIndex(rec.column) + rec.count - 1);
        sheet.getRange(`${rec.column}:${end}`).delete('Left');
      } else if (rec.op === 'addSheet') {
        context.workbook.worksheets.getItem(rec.name).delete();
      } else if (rec.op === 'renameSheet') {
        context.workbook.worksheets.getItem(rec.sheet).name = rec.name;
      } else if (rec.op === 'chart') {
        sheet.charts.getItem(rec.name).delete();
      } else if (rec.op === 'table') {
        sheet.tables.getItem(rec.name).delete();
      } else if (rec.op === 'pivot') {
        if (rec.createdSheet) context.workbook.worksheets.getItem(rec.sheet).delete();
        else sheet.pivotTables.getItem(rec.name).delete();
      } else if (rec.op === 'filter') {
        sheet.autoFilter.clearColumnCriteria(rec.key);
      } else if (rec.op === 'conditionalFormat') {
        sheet.getRange(rec.range).conditionalFormats.getItem(rec.id).delete();
      } else if (rec.op === 'borders') {
        const range = sheet.getRange(rec.range);
        for (const edge of rec.edges) {
          const border = range.format.borders.getItem(edge.edge);
          border.style = edge.style;
          border.weight = edge.weight;
          border.color = edge.color;
        }
      } else if (rec.op === 'freeze') {
        if (rec.address) sheet.freezePanes.freezeAt(rec.address);
        else sheet.freezePanes.unfreeze();
      } else if (rec.op === 'dataValidation') {
        sheet.getRange(rec.range).dataValidation.clear();
      } else if (rec.op === 'namedRange') {
        context.workbook.names.getItem(rec.name).delete();
      } else if (rec.op === 'hideSheet') {
        context.workbook.worksheets.getItem(rec.sheet).visibility = rec.visibility || 'Visible';
      } else if (rec.op === 'hide') {
        sheet.getRange(rec.range)[rec.property] = rec.previous;
      } else if (rec.op === 'tableColumn') {
        sheet.tables.getItem(rec.table).columns.getItem(rec.name).delete();
      } else if (rec.op === 'autofit') {
        for (const item of rec.widths || []) {
          if (item.width != null) sheet.getRange(item.column + '1').format.columnWidth = item.width;
        }
        for (const item of rec.heights || []) {
          if (item.height != null) sheet.getRange('A' + item.row).format.rowHeight = item.height;
        }
      }
    }
    await context.sync();
  });
}

async function applyWordActions(actions) {
  if (typeof Word === 'undefined') throw new Error('These changes need Word.');
  return Word.run(async context => {
    const undo = [];
    for (const action of actions) {
      const selection = context.document.getSelection();
      if (action.op !== 'insertText') selection.load('text');
      await context.sync();
      if (action.op === 'replaceSelection') {
        undo.push({ host: 'word', op: 'replaceSelection', text: selection.text || '' });
        selection.insertText(action.text, 'Replace');
      } else if (action.op === 'insertText') {
        context.document.body.insertText(action.text, action.where === 'start' ? 'Start' : 'End');
        undo.push({ host: 'word', op: 'removeText', text: action.text });
      } else if (action.op === 'formatSelection') {
        selection.font?.load?.(['bold', 'italic']);
        await context.sync();
        undo.push({ host: 'word', op: 'formatSelection', bold: selection.font?.bold, italic: selection.font?.italic });
        if (typeof action.bold === 'boolean') selection.font.bold = action.bold;
        if (typeof action.italic === 'boolean') selection.font.italic = action.italic;
      } else if (action.op === 'styleSelection') {
        undo.push({ host: 'word', op: 'styleSelection', style: selection.style || 'Normal' });
        selection.style = action.style;
      } else if (action.op === 'replaceAll') {
        const found = context.document.body.search(action.text, { matchCase: action.matchCase, matchWholeWord: action.matchWholeWord });
        found.load('items');
        await context.sync();
        for (const item of found.items || []) item.insertText(action.replacement, 'Replace');
      }
    }
    await context.sync();
    return undo;
  });
}

async function undoWord(records) {
  await Word.run(async context => {
    for (const rec of [...records].reverse()) {
      const selection = context.document.getSelection();
      if (rec.op === 'replaceSelection') selection.insertText(rec.text, 'Replace');
      else if (rec.op === 'formatSelection') {
        if (rec.bold != null) selection.font.bold = rec.bold;
        if (rec.italic != null) selection.font.italic = rec.italic;
      } else if (rec.op === 'styleSelection') selection.style = rec.style;
      else if (rec.op === 'removeText' && typeof context.document.body.search === 'function') {
        const found = context.document.body.search(rec.text, { matchCase: true });
        found.load('items');
        await context.sync();
        found.items?.[0]?.insertText('', 'Replace');
      }
    }
    await context.sync();
  });
}

async function applyPlan(actions) {
  const undo = [];
  const excel = actions.filter(action => action.host === 'excel');
  const word = actions.filter(action => action.host === 'word');
  if (excel.length) undo.push(...await applyExcelActions(excel));
  if (word.length) undo.push(...await applyWordActions(word));
  return undo;
}

async function undoPlan(records) {
  const excel = records.filter(rec => rec.host === 'excel');
  const word = records.filter(rec => rec.host === 'word');
  if (excel.length) await undoExcel(excel);
  if (word.length) await undoWord(word);
}

function attachUndo(actionsEl, label, records, noun) {
  const undoBtn = document.createElement('button');
  undoBtn.className = 'secondary';
  undoBtn.textContent = 'Undo last apply';
  undoBtn.addEventListener('click', async () => {
    undoBtn.disabled = true;
    try {
      await undoPlan(records);
      label.textContent = `UNDONE — RESTORED ${records.length} CHANGE${records.length === 1 ? '' : 'S'}`;
      toast(`Restored the previous ${noun}.`);
    } catch (error) {
      undoBtn.disabled = false;
      toast(`Could not undo: ${error.message}`);
    }
  });
  actionsEl.append(undoBtn);
}

async function renderPlanCard(actions, skipped, options = {}) {
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant edits-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  const direct = !options.manualOnly && Boolean($('directRevise')?.checked);
  label.textContent = direct ? `DIRECT REVISE · ${actions.length}` : `PROPOSED REVISION · ${actions.length} · APPLY WITH YOUR CLICK`;
  card.append(label);
  actions.forEach(action => {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.textContent = describeAction(action) + (actionIsRisky(action) ? ' · waits for your click' : '');
    card.append(row);
  });
  if (skipped) {
    const note = document.createElement('div');
    note.className = 'edit-note';
    note.textContent = `${skipped} change${skipped === 1 ? '' : 's'} skipped because the address or value was not usable.`;
    card.append(note);
  }
  if (actions.some(actionIsRisky)) {
    const warning = document.createElement('div');
    warning.className = 'edit-warning';
    warning.textContent = 'Some formulas can reach the network or run legacy calls. Direct revise leaves those for your click.';
    card.append(warning);
  }
  const buttonRow = document.createElement('div');
  buttonRow.className = 'button-row';
  const risky = actions.filter(actionIsRisky);
  const safe = actions.filter(action => !actionIsRisky(action));
  const manual = direct ? risky : actions;
  const noun = actions.some(action => action.host === 'word') && !actions.some(action => action.host === 'excel') ? 'document' : 'workbook';
  if (direct && safe.length) {
    try {
      const records = await applyPlan(safe);
      label.textContent = manual.length ? `REVISED ${safe.length} · ${manual.length} STILL NEED A CLICK` : `REVISED ${safe.length} CHANGE${safe.length === 1 ? '' : 'S'}`;
      attachUndo(buttonRow, label, records, noun);
      toast(noun === 'document' ? 'Revised the document.' : 'Revised the workbook.');
    } catch (error) {
      toast(`Could not revise: ${error.message}`);
      manual.splice(0, manual.length, ...actions);
    }
  }
  if (manual.length) {
    const apply = document.createElement('button');
    apply.className = 'primary';
    apply.textContent = direct && risky.length && manual.length === risky.length ? 'Apply risky changes' : 'Apply edits';
    const dismiss = document.createElement('button');
    dismiss.className = 'secondary';
    dismiss.textContent = 'Dismiss';
    apply.addEventListener('click', async () => {
      apply.disabled = true;
      dismiss.disabled = true;
      try {
        const records = await applyPlan(manual);
        label.textContent = `APPLIED ${manual.length} CHANGE${manual.length === 1 ? '' : 'S'}`;
        attachUndo(buttonRow, label, records, noun);
        toast(noun === 'document' ? 'Revised the document.' : 'Revised the workbook.');
      } catch (error) {
        apply.disabled = false;
        dismiss.disabled = false;
        toast(`Could not apply edits: ${error.message}`);
      }
    });
    dismiss.addEventListener('click', () => card.remove());
    buttonRow.append(apply, dismiss);
  } else {
    const dismiss = document.createElement('button');
    dismiss.className = 'secondary';
    dismiss.textContent = 'Dismiss';
    dismiss.addEventListener('click', () => card.remove());
    buttonRow.append(dismiss);
  }
  card.append(buttonRow);
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

async function withCells(edits, fn) {
  return Excel.run(async context => {
    const sheets = edits.map(edit => context.workbook.worksheets.getItemOrNullObject(edit.sheet));
    sheets.forEach(sheet => sheet.load('isNullObject'));
    await context.sync();
    const ranges = edits.map((edit, i) => sheets[i].isNullObject ? null : sheets[i].getRange(edit.cell));
    ranges.forEach(range => range && range.load(['values', 'formulas']));
    await context.sync();
    const result = await fn(ranges);
    await context.sync();
    return result;
  });
}

function oldCellText(old) {
  if (!old) return '(sheet missing)';
  if (String(old.formula || '').startsWith('=')) return old.formula;
  if (old.value === '' || old.value == null) return '(empty)';
  return String(old.value);
}

async function renderEditsCard(edits, skipped) {
  const conversation = $('conversation');
  const card = document.createElement('div'); card.className = 'message assistant edits-card';
  const label = document.createElement('div'); label.className = 'message-label'; label.textContent = `PROPOSED EDITS · ${edits.length} · REVIEW OLD → NEW · APPLY ONLY WITH YOUR CLICK`;
  card.append(label);
  let olds = null;
  try { olds = await withCells(edits, ranges => ranges.map(range => range ? { value: range.values?.[0]?.[0], formula: range.formulas?.[0]?.[0] } : null)); }
  catch { olds = null; }
  edits.forEach((edit, i) => {
    const row = document.createElement('div'); row.className = 'edit-row';
    const oldText = olds == null ? '(current unreadable)' : oldCellText(olds[i]);
    row.textContent = `${edit.sheet}!${edit.cell}: ${oldText} → ${String(edit.value)}`;
    card.append(row);
  });
  if (skipped) {
    const note = document.createElement('div'); note.className = 'edit-note';
    note.textContent = `${skipped} edit${skipped === 1 ? '' : 's'} skipped — each change must be one cell on a sheet in this workbook.`;
    card.append(note);
  }
  if (edits.some(edit => riskyEdit(edit.value))) {
    const warning = document.createElement('div'); warning.className = 'edit-warning';
    warning.textContent = 'Some formulas can reach the network or run legacy calls (WEBSERVICE, FILTERXML, RTD, HYPERLINK, and similar). Apply only if you trust them.';
    card.append(warning);
  }
  const actions = document.createElement('div'); actions.className = 'button-row';
  const apply = document.createElement('button'); apply.className = 'primary'; apply.textContent = 'Apply edits';
  const dismiss = document.createElement('button'); dismiss.className = 'secondary'; dismiss.textContent = 'Dismiss';
  apply.addEventListener('click', async () => {
    apply.disabled = true; dismiss.disabled = true;
    try {
      const recs = await withCells(edits, async ranges => {
        if (ranges.some(range => !range)) throw new Error('A proposed sheet is not in this workbook. Nothing was written.');
        const before = ranges.map((range, i) => ({ sheet: edits[i].sheet, cell: edits[i].cell, value: range.values?.[0]?.[0], formula: range.formulas?.[0]?.[0] }));
        ranges.forEach((range, i) => {
          const value = edits[i].value;
          if (typeof value === 'string' && value.startsWith('=')) range.formulas = [[value]];
          else range.values = [[value]];
        });
        return before;
      });
      label.textContent = `APPLIED ${edits.length} EDITS`;
      const undoBtn = document.createElement('button'); undoBtn.className = 'secondary'; undoBtn.textContent = 'Undo last apply';
      undoBtn.addEventListener('click', async () => {
        undoBtn.disabled = true;
        try {
          await Excel.run(async context => {
            for (const rec of recs) {
              const range = context.workbook.worksheets.getItem(rec.sheet).getRange(rec.cell);
              if (String(rec.formula || '').startsWith('=')) range.formulas = [[rec.formula]];
              else if (rec.value === '' || rec.value == null) range.clear();
              else range.values = [[rec.value]];
            }
            await context.sync();
          });
          label.textContent = `UNDONE — RESTORED ${recs.length} CELLS`;
          toast('Restored previous cell values.');
        } catch (error) { undoBtn.disabled = false; toast(`Could not undo: ${error.message}`); }
      });
      actions.append(undoBtn);
      toast(`Applied ${edits.length} edits to the workbook.`);
    } catch (error) {
      apply.disabled = false; dismiss.disabled = false;
      toast(`Could not apply edits: ${error.message}`);
    }
  });
  dismiss.addEventListener('click', () => card.remove());
  actions.append(apply, dismiss); card.append(actions);
  conversation.append(card); conversation.scrollTop = conversation.scrollHeight;
}

function dedupContext(messages) {
  // Older turns keep their prose only; re-sending every past 45k context block would multiply input tokens each turn.
  return messages.map((m, i) => (m.role === 'user' && i < messages.length - 1)
    ? { ...m, content: m.content.replace(/\n\n\[(?:Excel (?:selection|used range|workbook metadata|range)|Word (?:selection|document|document outline))[^\]]*\]\n[\s\S]*$/, '') }
    : m);
}

function consumeSseFrames(buf, onFrame) {
  let rest = buf;
  let idx;
  while ((idx = rest.indexOf('\n\n')) >= 0) {
    onFrame(rest.slice(0, idx));
    rest = rest.slice(idx + 2);
  }
  return rest;
}

async function streamChat(payload) {
  const controller = new AbortController();
  abortChat = controller;
  let bubble = null;
  let text = '';
  try {
    const headers = { 'Content-Type': 'application/json' };
    const token = relayToken();
    if (token) headers.Authorization = `Bearer ${token}`;
    const response = await fetch(`${API}/api/chat-stream`, { method: 'POST', headers, body: JSON.stringify(payload), signal: controller.signal });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data.error || `Request failed (${response.status})`);
    }
    bubble = renderMessage('assistant', '');
    const bodyEl = bubble.children[bubble.children.length - 1]; // HTMLCollection has no .at()
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buf = '', model = '';
    const onFrame = frame => {
    for (const line of frame.split('\n')) {
      if (!line.startsWith('data: ')) continue;
      const raw = line.slice(6);
      if (raw === '[DONE]') continue;
      let event; try { event = JSON.parse(raw); } catch { continue; }
      if (event.error) throw new Error(event.error);
      if (event.model) model = event.model;
      if (event.status === 'thinking' && !text) { bodyEl.textContent = 'Thinking…'; continue; }
      const delta = event.choices?.[0]?.delta?.content;
      if (delta) {
        text += delta; bodyEl.textContent = text;
        $('conversation').scrollTop = $('conversation').scrollHeight;
      }
    }
  };
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      buf = consumeSseFrames(buf, onFrame);
    }
    buf += decoder.decode();
    if (buf.trim()) onFrame(buf);
    if (!text.trim()) {
      const error = new Error('The provider returned no text answer.');
      error.bubble = bubble;
      throw error;
    }
    return { text, model, bubble };
  } catch (error) {
    if (bubble) { error.bubble = bubble; error.partial = text; }
    throw error;
  } finally {
    if (abortChat === controller) abortChat = null;
  }
}

function resetComposer() {
  state.busy = false;
  $('send').textContent = '↑';
  $('send').setAttribute?.('aria-label', 'Send message');
}

function stopGeneration() {
  state.turn += 1;
  abortChat?.abort();
  resetComposer();
}

function operatingMode() {
  const value = $('workMode')?.value;
  return value === 'chat' || value === 'plan' || value === 'agent' ? value : 'legacy';
}

async function readWorkbookMeta() {
  if (!state.officeReady) return null;
  if (state.host === 'Word') {
    const text = 'host: Word\nAsk for the selection before quoting the document.';
    return { header: '[Word document outline]', text, address: 'Document', kind: 'metadata', truncated: false, chars: text.length };
  }
  return Excel.run(async context => {
    const book = context.workbook;
    book.load?.('name');
    const sheets = book.worksheets;
    sheets.load?.('items/name');
    await context.sync();
    const lines = [`workbook: ${book.name || '(unsaved)'}`];
    for (const sheet of (sheets.items || []).slice(0, 30)) {
      let extra = '';
      try {
        const used = sheet.getUsedRangeOrNullObject?.();
        if (used) {
          used.load?.(['isNullObject', 'address', 'rowCount', 'columnCount']);
          await context.sync();
          if (!used.isNullObject) extra += ` used ${used.address} rows ${used.rowCount} cols ${used.columnCount}`;
        }
      } catch { /* this sheet has no used-range helper */ }
      try {
        const tables = sheet.tables;
        tables?.load?.('items/name');
        await context.sync();
        const names = (tables?.items || []).map(table => table.name).filter(Boolean);
        if (names.length) extra += ` tables ${names.join(', ')}`;
      } catch { /* tables are not available here */ }
      lines.push(`sheet: ${sheet.name || 'Sheet'}${extra}`);
    }
    const selected = book.getSelectedRange?.();
    if (selected) {
      selected.load?.(['address', 'rowCount', 'columnCount']);
      await context.sync();
      if (selected.address) lines.push(`selection: ${selected.address} rows ${selected.rowCount || ''} cols ${selected.columnCount || ''}`);
    }
    const text = lines.join('\n').slice(0, 4000);
    return { header: '[Excel workbook metadata]', text, address: 'Workbook', kind: 'metadata', truncated: false, chars: text.length };
  });
}

async function contextForMode(mode) {
  if (!state.officeReady || !$('includeContext').checked) return null;
  if (mode === 'chat' || state.forceSheet) return buildContext();
  try { return await readWorkbookMeta(); }
  catch { return buildContext(); }
}

async function readGrid(sheet, range) {
  const size = rangeSize(range);
  // Keep the full column width. Trimming to 8 columns hid the bucket and amount
  // columns of wide sheets, so the model never saw the fields it was asked about.
  const cols = Math.min(size.cols, READ_MAX_CELLS);
  const rows = Math.max(1, Math.min(size.rows, Math.floor(READ_MAX_CELLS / cols)));
  return Excel.run(async context => {
    const worksheet = context.workbook.worksheets.getItem(sheet);
    const cells = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const cell = worksheet.getRange(cellAddress(range, r, c));
        cell.load(['values', 'formulas']);
        row.push(cell);
      }
      cells.push(row);
    }
    await context.sync();
    return {
      values: cells.map(row => row.map(cell => cell.values?.[0]?.[0] ?? '')),
      formulas: cells.map(row => row.map(cell => cell.formulas?.[0]?.[0] ?? ''))
    };
  });
}

function plainRange(address) {
  return normalizeRange(String(address || '').split('!').pop());
}

function clampRange(addr, maxRows, maxCols) {
  const size = rangeSize(addr);
  const rows = Math.min(size.rows, maxRows);
  const cols = Math.min(size.cols, maxCols);
  if (rows === size.rows && cols === size.cols) return addr;
  const top = addr.split(':')[0].match(/^([A-Z]+)(\d+)$/);
  return `${top[1]}${top[2]}:${indexToColumn(columnIndex(top[1]) + cols - 1)}${Number(top[2]) + rows - 1}`;
}

// Column statistics rather than raw cells, so a sheet far larger than the context
// window can still be understood well enough to choose pivot fields.
async function profileSheet(sheetName, requested) {
  return Excel.run(async context => {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const asked = normalizeRange(requested || '');
    let addr = asked ? clampRange(asked, PROFILE_MAX_ROWS, PROFILE_MAX_COLS) : '';
    let totalRows = 0;
    if (!addr) {
      const used = worksheet.getUsedRangeOrNullObject();
      used.load(['isNullObject', 'address', 'rowCount', 'columnCount']);
      await context.sync();
      if (used.isNullObject) return null;
      const full = plainRange(used.address);
      if (!full) return null;
      totalRows = rangeSize(full).rows;
      addr = clampRange(full, PROFILE_MAX_ROWS, PROFILE_MAX_COLS);
    }
    const target = worksheet.getRange(addr);
    target.load('values');
    await context.sync();
    const values = target.values || [];
    if (!values.length) return null;
    const header = (values[0] || []).map(value => String(value ?? '').trim());
    const body = values.slice(1);
    const columns = header.map((name, c) => {
      const seen = new Map();
      let numeric = 0, sum = 0, min = null, max = null, blank = 0;
      for (const row of body) {
        const value = row?.[c];
        if (value == null || value === '') { blank += 1; continue; }
        const key = String(value);
        seen.set(key, (seen.get(key) || 0) + 1);
        const n = typeof value === 'number' ? value : Number(key);
        if (Number.isFinite(n) && key.trim() !== '') {
          numeric += 1; sum += n;
          min = min == null ? n : Math.min(min, n);
          max = max == null ? n : Math.max(max, n);
        }
      }
      const ranked = [...seen.entries()].sort((a, b) => b[1] - a[1]);
      return {
        letter: indexToColumn(c + 1), header: name || '(blank header)',
        filled: body.length - blank, blank, distinct: seen.size, numeric,
        sum: numeric ? Math.round(sum * 100) / 100 : null, min, max,
        samples: ranked.slice(0, PROFILE_DISTINCT).map(([value, count]) => `${value} (${count})`)
      };
    });
    return { address: addr, columns, sampledRows: body.length, totalRows: totalRows || body.length };
  });
}

function formatProfile(sheetName, profile) {
  const head = `[Excel profile ${sheetName}!${profile.address} — ${profile.sampledRows} data rows sampled${profile.totalRows > profile.sampledRows ? ` of ${profile.totalRows}` : ''}, ${profile.columns.length} columns]`;
  const lines = profile.columns.map(column => {
    const parts = [`${column.letter} | ${column.header}`, `filled ${column.filled}`, `distinct ${column.distinct}`];
    if (column.blank) parts.push(`blank ${column.blank}`);
    if (column.numeric) parts.push(`numeric ${column.numeric}`, `sum ${column.sum}`, `min ${column.min}`, `max ${column.max}`);
    if (column.samples.length) parts.push(`e.g. ${column.samples.join(', ')}`);
    return parts.join(' | ');
  });
  return `${head}\n${lines.join('\n')}`;
}

async function gridsForChecks(checks) {
  const grids = {};
  for (const check of checks) {
    try { grids[check.id] = await readGrid(check.sheet, check.range); }
    catch { grids[check.id] = { values: [], formulas: [] }; }
  }
  return grids;
}

function usableChecks(plan) {
  return (plan?.finalValidation || []).map(check => {
    const sheet = normalizeSheet(check.sheet);
    const range = boundedRange(check.range, 160);
    return sheet && range ? { ...check, sheet, range } : null;
  }).filter(Boolean);
}

async function routeForPrompt(text, mode) {
  let configured = { qwen: false, grok: false };
  try {
    const health = await api('health');
    configured = health.configured || configured;
  } catch { /* the send path reports a down relay */ }
  const selected = $('activeProvider').value === 'auto' ? 'auto' : ($('activeProvider').value || 'qwen');
  const routed = AgentCore.routeProvider({ selected, mode, fast: Boolean($('fastMode').checked), text, ready: configured });
  if (routed.error) return routed;
  return {
    ...routed,
    model: routed.provider === 'grok' ? $('grokModel').value.trim() : $('qwenModel').value.trim(),
    endpoint: $('qwenEndpoint').value.trim()
  };
}

function renderAuditCard() {
  if (!state.audit.length) return;
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant audit-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  label.textContent = 'AUDIT';
  const body = document.createElement('div');
  body.className = 'edit-row';
  body.textContent = AgentCore.summarize(state.audit);
  card.append(label, body);
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

function renderValidation(results) {
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant edits-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  const failed = results.filter(result => result.status === 'FAIL');
  label.textContent = failed.length ? `VALIDATION FAILED · ${failed.length}` : `VALIDATION · ${results.length}`;
  card.append(label);
  results.forEach(result => {
    const row = document.createElement('div');
    row.className = result.status === 'FAIL' ? 'edit-warning' : 'edit-note';
    row.textContent = `${result.status} ${result.check}: before ${result.before} after ${result.after} difference ${result.difference}`;
    card.append(row);
    if (result.status === 'FAIL' || result.status === 'NEEDS_USER') {
      const explained = AgentCore.explainFailure(result);
      const note = document.createElement('div');
      note.className = 'edit-warning';
      note.textContent = `${explained.issue} — ${explained.possibleCause} ${explained.recommendedAction}`;
      card.append(note);
    }
  });
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

async function executeClassified(actions, { direct }) {
  const blocked = [];
  const safe = [];
  for (const action of actions) {
    if (!direct || actionIsRisky(action) || AgentCore.isDestructive(action)) blocked.push(action);
    else safe.push(action);
  }
  let status = blocked.length ? 'BLOCKED' : 'SUCCESS';
  if (safe.length) {
    try {
      const records = await applyPlan(safe);
      status = blocked.length ? 'WARNING' : 'SUCCESS';
      // Without this card the agent loop applied changes with no Undo offered at all.
      renderAppliedCard(safe, records, blocked.length);
    } catch (error) {
      status = 'FAILED';
      toast(`Could not revise: ${error.message}`);
    }
  }
  if (blocked.length) await renderPlanCard(blocked, 0, { manualOnly: true });
  return { status, blocked: blocked.length, applied: safe.length };
}

function renderAppliedCard(actions, records, blockedCount) {
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant edits-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  label.textContent = blockedCount
    ? `REVISED ${actions.length} · ${blockedCount} STILL NEED A CLICK`
    : `REVISED ${actions.length} CHANGE${actions.length === 1 ? '' : 'S'}`;
  card.append(label);
  actions.forEach(action => {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.textContent = describeAction(action);
    card.append(row);
  });
  const buttonRow = document.createElement('div');
  buttonRow.className = 'button-row';
  const noun = actions.some(action => action.host === 'word') && !actions.some(action => action.host === 'excel') ? 'document' : 'workbook';
  if (records.length) attachUndo(buttonRow, label, records, noun);
  card.append(buttonRow);
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

function renderModePlan(plan, actions, skipped) {
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant edits-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  label.textContent = plan ? `PLAN · ${plan.goal}` : 'PLAN · APPROVAL REQUIRED';
  card.append(label);
  (plan?.steps || []).forEach(step => {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.textContent = `${step.id}. ${step.objective}${step.requiresApproval ? ' · approval' : ''}`;
    card.append(row);
  });
  for (const [title, items] of [['Assumptions', plan?.assumptions], ['Inputs', plan?.requiredInputs], ['Risks', plan?.risks]]) {
    if (!items?.length) continue;
    const row = document.createElement('div');
    row.className = 'edit-note';
    row.textContent = `${title}: ${items.join('; ')}`;
    card.append(row);
  }
  actions.forEach(action => {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.textContent = describeAction(action);
    card.append(row);
  });
  if (skipped) {
    const note = document.createElement('div');
    note.className = 'edit-note';
    note.textContent = `${skipped} change${skipped === 1 ? '' : 's'} skipped.`;
    card.append(note);
  }
  const buttonRow = document.createElement('div');
  buttonRow.className = 'button-row';
  const approve = document.createElement('button');
  approve.className = 'primary';
  approve.textContent = 'Approve';
  const dismiss = document.createElement('button');
  dismiss.className = 'secondary';
  dismiss.textContent = 'Dismiss';
  approve.addEventListener('click', async () => {
    approve.disabled = true;
    dismiss.disabled = true;
    if (!actions.length) { toast('The plan has no actions to run.'); return; }
    const result = await executeClassified(actions, { direct: Boolean($('directRevise').checked) });
    state.audit.push(AgentCore.audit({ kind: 'execute', status: result.status, count: actions.length, blocked: result.blocked }));
    label.textContent = result.status === 'SUCCESS' ? 'PLAN APPROVED · APPLIED' : `PLAN APPROVED · ${result.status}`;
    renderAuditCard();
  });
  dismiss.addEventListener('click', () => card.remove());
  buttonRow.append(approve, dismiss);
  card.append(buttonRow);
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

async function sendAgentPrompt(prompt) {
  const clean = prompt.trim();
  if (!clean || state.busy) return;
  const turn = ++state.turn;
  state.busy = true;
  $('send').textContent = '■';
  $('send').setAttribute?.('aria-label', 'Stop');
  try {
    await ensureKeys();
    if (turn !== state.turn) return;
    const mode = operatingMode();
    let routed = await routeForPrompt(clean, mode);
    if (routed.error) { toast(routed.error); return; }
    state.audit.push(AgentCore.audit({ kind: 'request', mode, provider: routed.provider, reason: routed.reason, text: clean.slice(0, 500) }));
    let ctx = null;
    try { ctx = await contextForMode(mode); } catch (error) { if (turn === state.turn) toast(`Could not read workbook context: ${error.message}`); }
    if (turn !== state.turn) return;
    const content = `${clean}${ctx ? `\n\n${ctx.header}\n${ctx.text}` : ''}\n\n[Mode: ${mode}]`;
    if (content.length > 49_500) { toast('Message plus workbook context is too large — shorten the message or turn context off.'); return; }
    state.forceSheet = false;
    state.messages.push({ role: 'user', content });
    renderMessage('user', clean + (ctx ? `\n\nContext: ${ctx.address}${ctx.truncated ? ' (truncated)' : ''}` : ''));
    state.audit.push(AgentCore.audit({ kind: 'context', address: ctx?.address || 'none', chars: ctx?.chars || 0 }));
    let guard = 0;
    let retried = false;
    while (guard < AgentCore.MAX_LOOPS) {
      if (turn !== state.turn) return;
      let result;
      try {
        result = await streamChat({ provider: routed.provider, model: routed.model, endpoint: routed.endpoint, messages: dedupContext(state.messages.slice(-24)), fast: $('fastMode').checked, host: state.host === 'Word' ? 'word' : 'excel', mode });
      } catch (error) {
        if (error.name === 'AbortError' || retried || !routed.fallback || turn !== state.turn) throw error;
        retried = true;
        state.audit.push(AgentCore.audit({ kind: 'fallback', provider: routed.fallback, reason: error.message }));
        routed = { ...routed, provider: routed.fallback, model: routed.fallback === 'grok' ? $('grokModel').value.trim() : $('qwenModel').value.trim(), fallback: '' };
        continue;
      }
      if (turn !== state.turn) return;
      const envelope = AgentCore.parseEnvelope(result.text);
      const revision = parseRevision(result.text);
      state.messages.push({ role: 'assistant', content: result.text });
      state.lastAnswer = envelope.prose || revision.prose;
      $('insertAnswer').disabled = !state.lastAnswer;
      $('lastAnswerState').textContent = result.model ? `via ${result.model}` : '';
      result.bubble.children[result.bubble.children.length - 1].textContent = state.lastAnswer;
      state.audit.push(AgentCore.audit({ kind: 'model', provider: routed.provider, model: result.model, decision: envelope.decision || 'COMPLETE', plan: envelope.plan?.goal || '' }));
      const normalized = envelope.actions ? normalizeActions(envelope.actions) : { actions: revision.actions || [], skipped: revision.skipped || 0 };
      if (mode === 'chat') {
        if (normalized.actions.length || revision.edits) {
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: 'chat mode does not write' }));
          toast('Chat mode does not change the file.');
        }
        break;
      }
      if (mode === 'plan' || envelope.decision === 'NEEDS_USER' || (mode === 'agent' && !$('directRevise').checked && normalized.actions.length)) {
        renderModePlan(envelope.plan, normalized.actions, normalized.skipped);
        break;
      }
      if (envelope.profileRequest && !normalized.actions.length && guard < AgentCore.MAX_LOOPS - 1) {
        const sheet = normalizeSheet(envelope.profileRequest.sheet);
        let profile = null;
        let failure = '';
        if (!sheet) failure = `sheet "${envelope.profileRequest.sheet}" is not a usable sheet name`;
        else {
          try { profile = await profileSheet(sheet, envelope.profileRequest.range); }
          catch (error) { failure = error.message; }
          if (!profile && !failure) failure = 'the sheet or range is empty';
        }
        if (profile) {
          const text = formatProfile(sheet, profile).slice(0, 8000);
          state.messages.push({ role: 'user', content: `Requested profile.\n\n${text}` });
          state.audit.push(AgentCore.audit({ kind: 'context', address: `${sheet}!${profile.address} (profile)`, chars: text.length }));
        } else {
          state.messages.push({ role: 'user', content: `No profile was produced for "${envelope.profileRequest.sheet}": ${failure}. Nothing was read, so this is not a successful read.` });
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: `profile rejected: ${failure}` }));
        }
        guard += 1;
        continue;
      }
      if (envelope.contextRequest && !normalized.actions.length && guard < AgentCore.MAX_LOOPS - 1) {
        const sheet = normalizeSheet(envelope.contextRequest.sheet);
        const asked = normalizeRange(envelope.contextRequest.range);
        const range = boundedRange(envelope.contextRequest.range, READ_MAX_CELLS);
        if (sheet && range) {
          const grid = await readGrid(sheet, range);
          const full = annotateRows(grid.values, grid.formulas);
          const text = full.slice(0, 8000);
          const cut = full.length > text.length ? '\n[truncated: the read exceeded the context budget]' : '';
          const note = `Requested context.\n\n[Excel range ${sheet}!${range}]\n${text}${cut}`;
          state.messages.push({ role: 'user', content: note });
          state.audit.push(AgentCore.audit({ kind: 'context', address: `${sheet}!${range}`, chars: text.length }));
          guard += 1;
          continue;
        }
        // State why nothing was read. Falling through used to report "Observed: SUCCESS",
        // which the model took as a successful but empty read, then refused to guess columns.
        const size = asked ? rangeSize(asked) : null;
        const reason = !sheet
          ? `sheet "${envelope.contextRequest.sheet}" is not a usable sheet name`
          : !asked
            ? `range "${envelope.contextRequest.range}" is not a valid A1 range`
            : `${sheet}!${asked} is ${size.rows * size.cols} cells, over the ${READ_MAX_CELLS}-cell read limit`;
        state.messages.push({ role: 'user', content: `No range was read: ${reason}. Nothing was returned, so this is not a successful read. Request a smaller range and ask again.` });
        state.audit.push(AgentCore.audit({ kind: 'blocked', reason: `context request rejected: ${reason}` }));
        guard += 1;
        continue;
      }
      const checks = usableChecks(envelope.plan);
      const profile = readFinanceProfile();
      const beforeGrids = await gridsForChecks(checks);
      let exec = { status: 'SUCCESS', blocked: 0, applied: 0 };
      if (normalized.actions.length) exec = await executeClassified(normalized.actions, { direct: Boolean($('directRevise').checked) });
      state.audit.push(AgentCore.audit({ kind: 'execute', status: exec.status, count: normalized.actions.length, blocked: exec.blocked }));
      const afterGrids = await gridsForChecks(checks);
      const validations = checks.map(check => AgentCore.evaluateGridCheck({ ...check, profile }, afterGrids[check.id], beforeGrids[check.id]));
      validations.forEach(result => state.audit.push(AgentCore.audit({ kind: 'validation', status: result.status, check: result.check, before: result.before, after: result.after, difference: result.difference })));
      if (validations.length) renderValidation(validations);
      const next = AgentCore.decide({ decision: envelope.decision, validations, iteration: guard, maxLoops: AgentCore.MAX_LOOPS });
      if (next === 'FAILED') {
        state.audit.push(AgentCore.audit({ kind: 'stopped', reason: 'validation failed' }));
        toast('Stopped. A validation failed. Nothing further will run.');
        break;
      }
      if (next === 'NEEDS_USER') {
        state.audit.push(AgentCore.audit({ kind: 'stopped', reason: 'a check needs a column' }));
        toast('Stopped. A finance check needs a column before it can pass or fail.');
        break;
      }
      if (next === 'CONTINUE' || next === 'REPLAN') {
        const observed = validations.map(result => `${result.check} ${result.status}`).join('; ') || exec.status;
        state.messages.push({ role: 'user', content: `Observed: ${observed}. Continue only if every check passed.` });
        guard += 1;
        continue;
      }
      break;
    }
    if (guard >= AgentCore.MAX_LOOPS) {
      state.audit.push(AgentCore.audit({ kind: 'stopped', reason: 'loop limit' }));
      toast('Stopped at the loop limit.');
    }
    renderAuditCard();
    refreshContextCard();
  } catch (error) {
    const bodyEl = error.bubble?.parent ? error.bubble.children[error.bubble.children.length - 1] : null;
    if (error.name === 'AbortError') {
      if (bodyEl) bodyEl.textContent = error.partial ? `${error.partial}\n\n(stopped)` : '(stopped)';
      if (error.bubble?.parent) toast('Stopped.');
      return;
    }
    if (turn !== state.turn) return;
    if (!state.lastAnswer) $('insertAnswer').disabled = true;
    if (bodyEl) {
      bodyEl.textContent = error.partial ? `${error.partial}\n\n(${error.message})` : `Error: ${error.message}`;
      toast(error.message);
    } else {
      renderMessage('assistant', `Error: ${error.message}`);
      toast(error.message);
    }
  } finally {
    if (turn === state.turn) resetComposer();
  }
}

async function sendPrompt(prompt) {
  if (operatingMode() !== 'legacy') return sendAgentPrompt(prompt);
  const clean = prompt.trim();
  if (!clean || state.busy) return;
  const turn = ++state.turn;
  state.busy = true;
  $('send').textContent = '■';
  $('send').setAttribute?.('aria-label', 'Stop');
  try {
    await ensureKeys();
    if (turn !== state.turn) return;
    const provider = $('activeProvider').value;
    let ctx = null;
    try { ctx = await buildContext(); } catch (error) { if (turn === state.turn) toast(`Could not read workbook context: ${error.message}`); }
    if (turn !== state.turn) return;
    const content = clean + (ctx ? `\n\n${ctx.header}\n${ctx.text}` : '');
    if (content.length > 49_500) { toast('Message plus workbook context is too large — shorten the message or turn context off.'); return; }
    state.forceSheet = false;
    state.messages.push({ role: 'user', content }); renderMessage('user', clean + (ctx ? `\n\nContext: ${ctx.address}${ctx.truncated ? ' (truncated)' : ''}` : ''));
    const result = await streamChat({ provider, model: provider === 'qwen' ? $('qwenModel').value.trim() : $('grokModel').value.trim(), endpoint: $('qwenEndpoint').value.trim(), messages: dedupContext(state.messages.slice(-24)), fast: $('fastMode').checked, host: state.host === 'Word' ? 'word' : 'excel' });
    if (turn !== state.turn) return;
    const revision = parseRevision(result.text);
    state.messages.push({ role: 'assistant', content: result.text });
    state.lastAnswer = revision.prose; $('insertAnswer').disabled = !revision.prose; $('lastAnswerState').textContent = result.model ? `via ${result.model}` : '';
    result.bubble.children[result.bubble.children.length - 1].textContent = revision.prose;
    if (revision.actions) await renderPlanCard(revision.actions, revision.skipped);
    else if (revision.edits) await renderEditsCard(revision.edits, revision.skipped);
    refreshContextCard();
  } catch (error) {
    const bodyEl = error.bubble?.parent ? error.bubble.children[error.bubble.children.length - 1] : null;
    if (error.name === 'AbortError') {
      if (bodyEl) bodyEl.textContent = error.partial ? `${error.partial}\n\n(stopped)` : '(stopped)';
      if (error.bubble?.parent) toast('Stopped.');
      return;
    }
    if (turn !== state.turn) return;
    if (!state.lastAnswer) $('insertAnswer').disabled = true;
    if (bodyEl) {
      bodyEl.textContent = error.partial ? `${error.partial}\n\n(${error.message})` : `Error: ${error.message}`;
      toast(error.message);
    } else {
      renderMessage('assistant', `Error: ${error.message}`);
      toast(error.message);
    }
  } finally {
    if (turn === state.turn) resetComposer();
  }
}

function clearConversation() {
  const conversation = $('conversation');
  while (conversation.children.length) conversation.children[0].remove();
  const welcome = document.createElement('div'); welcome.className = 'welcome';
  const icon = document.createElement('div'); icon.className = 'welcome-icon'; icon.textContent = '✦';
  const title = document.createElement('h1'); title.textContent = 'What should I revise?';
  const copy = document.createElement('p'); copy.textContent = 'Ask about the selection, or tell me to change the sheet or the document.';
  welcome.append(icon, title, copy);
  conversation.append(welcome);
}

function newChat() {
  stopGeneration();
  state.messages = [];
  state.lastAnswer = '';
  state.forceSheet = false;
  state.audit = [];
  $('insertAnswer').disabled = true;
  $('lastAnswerState').textContent = '';
  clearConversation();
}

function showExecutionPath() {
  const el = $('executionPath');
  if (!el) return;
  el.textContent = hostedRelay()
    ? 'Cloud · questions and workbook text go through Cloudflare to the selected provider'
    : 'Local · the relay stays on this PC';
}

function splitColumnList(value) {
  return String(value || '').split(',').map(part => part.trim()).filter(Boolean);
}

function readFinanceProfile() {
  const tolerance = $('checkTolerance')?.value;
  return AgentCore.normalizeFinanceProfile({
    debitColumn: $('debitColumn')?.value,
    creditColumn: $('creditColumn')?.value,
    amountColumn: $('amountColumn')?.value,
    keyColumns: splitColumnList($('keyColumns')?.value),
    requiredColumns: splitColumnList($('requiredColumns')?.value),
    entityColumn: $('entityColumn')?.value,
    currencyColumn: $('currencyColumn')?.value,
    dateColumn: $('dateColumn')?.value,
    tolerance: tolerance == null || tolerance === '' ? null : Number(tolerance)
  });
}

function loadFinanceProfile() {
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem('eqgc.finance.v1') || '{}'); } catch { saved = {}; }
  const profile = AgentCore.normalizeFinanceProfile(saved);
  const fields = {
    debitColumn: profile.debitColumn,
    creditColumn: profile.creditColumn,
    amountColumn: profile.amountColumn,
    keyColumns: profile.keyColumns.join(', '),
    requiredColumns: profile.requiredColumns.join(', '),
    entityColumn: profile.entityColumn,
    currencyColumn: profile.currencyColumn,
    dateColumn: profile.dateColumn,
    checkTolerance: profile.tolerance
  };
  for (const [id, value] of Object.entries(fields)) {
    const el = $(id);
    if (el) el.value = value == null ? '' : String(value);
  }
}

function saveFinanceProfile() {
  try { localStorage.setItem('eqgc.finance.v1', JSON.stringify(readFinanceProfile())); } catch { /* storage can be unavailable */ }
}

function applyHostedChrome() {
  showExecutionPath();
  if (!hostedRelay()) return;
  $('relayFields')?.classList.remove('hidden');
  $('localKeyFields')?.classList.add('hidden');
  const hint = $('keyHint');
  if (hint) hint.textContent = 'This pane is served by Cloudflare. The relay token unlocks it. Qwen and Grok keys stay in Cloudflare secrets and are not typed here. Remember keeps the token in this pane only.';
  const saved = relayToken();
  if (saved) sessionRelayToken = saved;
}

async function refreshSettings() {
  if (hostedRelay() && !relayToken()) {
    const status = $('keyStatusHosted');
    if (status) { status.textContent = 'Add the Cloudflare relay token'; status.style.color = '#777'; }
    setRelayDot('warn');
    return;
  }
  try {
    const data = await api('health');
    if (hostedRelay()) {
      const provider = $('activeProvider').value === 'grok' ? 'grok' : 'qwen';
      const active = Boolean(data.configured?.[provider]);
      const status = $('keyStatusHosted');
      if (status) {
        status.textContent = active ? 'Relay token accepted · provider key is on Cloudflare' : 'Relay token accepted · add the provider key in Cloudflare';
        status.style.color = active ? '#168251' : '#777';
      }
      setRelayDot(active ? 'ok' : 'warn');
      return;
    }
    const provider = $('provider').value;
    const remembered = Boolean(storedKeys()[provider]);
    const active = Boolean(data.configured[provider]);
    $('keyStatus').textContent = active ? (remembered ? 'Key active · remembered on this PC' : 'Key active for this session') : '';
    $('keyStatus').style.color = active ? '#168251' : '#777';
    setRelayDot(active ? 'ok' : 'warn');
  } catch { $('keyStatus').textContent = 'Relay unavailable'; $('keyStatus').style.color = '#b23b36'; setRelayDot('offline'); }
}

function toggleProviderFields() {
  const isQwen = $('provider').value === 'qwen';
  $('qwenFields').classList.toggle('hidden', !isQwen); $('grokFields').classList.toggle('hidden', isQwen);
  $('activeProvider').value = $('provider').value;
  refreshSettings();
}

async function insertAnswer() {
  if (!state.lastAnswer || !state.officeReady) return;
  const answer = state.lastAnswer;
  if (answer.length > CELL_CHAR_LIMIT) { toast('That answer is longer than Excel’s 32,767 character cell limit.'); return; }
  const formula = answer.trim();
  const asFormula = formula.startsWith('=') && !formula.includes('\n') && formula.length <= 8192;
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange().getCell(0, 0);
      if (asFormula) range.formulas = [[formula]];
      else range.values = [[answer]];
      await context.sync();
    });
    toast(asFormula ? 'Formula inserted into the selected cell.' : 'Answer inserted into the selected cell.');
  } catch (error) { toast(`Could not insert answer: ${error.message}`); }
}

$('settingsToggle').addEventListener('click', () => { $('settings').classList.toggle('hidden'); refreshSettings(); });
$('closeSettings').addEventListener('click', () => $('settings').classList.add('hidden'));
$('provider').addEventListener('change', toggleProviderFields);
$('activeProvider').addEventListener('change', () => {
  const value = $('activeProvider').value;
  if (value !== 'qwen' && value !== 'grok') return;
  $('provider').value = value;
  toggleProviderFields();
});
$('refreshSelection').addEventListener('click', refreshContextCard);
$('includeContext').addEventListener('change', refreshContextCard);
$('addUsedRange').addEventListener('click', () => { state.forceSheet = true; refreshContextCard(); toast('Next question will use the whole active sheet.'); });
$('saveKey').addEventListener('click', async () => {
  if (hostedRelay()) {
    const token = $('relayToken').value.trim();
    if (token.length < 16) return toast('Enter the Cloudflare relay token.');
    rememberRelay($('rememberKey').checked ? token : '');
    if (!$('rememberKey').checked) sessionRelayToken = token;
    $('relayToken').value = '';
    await refreshSettings();
    toast($('rememberKey').checked ? 'Relay token remembered in this pane.' : 'Relay token kept for this session.');
    return;
  }
  const apiKey = $('apiKey').value.trim(); if (!apiKey) return toast('Enter an API key first.');
  const provider = $('provider').value;
  try {
    await api('configure', { provider, apiKey });
    rememberKey(provider, $('rememberKey').checked ? apiKey : '');
    keysReady = null;
    $('apiKey').value = ''; await refreshSettings();
    toast($('rememberKey').checked ? 'Key loaded and remembered on this PC.' : 'Key loaded into local relay memory.');
  }
  catch (error) { toast(error.message); }
});
$('clearKey').addEventListener('click', async () => {
  if (hostedRelay()) {
    rememberRelay('');
    await refreshSettings();
    toast('Relay token cleared from this pane.');
    return;
  }
  const provider = $('provider').value;
  try { await api('clear-key', { provider }); rememberKey(provider, ''); keysReady = null; await refreshSettings(); toast('Key cleared.'); }
  catch (error) { toast(error.message); }
});
$('newChat').addEventListener('click', newChat);
$('composer').addEventListener('submit', async event => {
  event.preventDefault();
  if (state.busy) { stopGeneration(); return; }
  const value = $('prompt').value;
  if (!value.trim()) return;
  $('prompt').value = '';
  await sendPrompt(value);
});
$('prompt').addEventListener('keydown', event => { if (event.key === 'Enter' && !event.shiftKey) { event.preventDefault(); $('composer').requestSubmit(); } });
$('insertAnswer').addEventListener('click', insertAnswer);
document.querySelectorAll('.action').forEach(button => button.addEventListener('click', () => { $('prompt').value = button.dataset.prompt; $('prompt').focus(); }));
loadFinanceProfile();
['debitColumn', 'creditColumn', 'amountColumn', 'keyColumns', 'requiredColumns', 'entityColumn', 'currencyColumn', 'dateColumn', 'checkTolerance'].forEach(id => {
  $(id).addEventListener('change', saveFinanceProfile);
});

Office.onReady(info => {
  if (info.host === Office.HostType.Excel || info.host === Office.HostType.Word) {
    state.host = info.host === Office.HostType.Word ? 'Word' : 'Excel';
    state.officeReady = true;
    const eyebrow = $('contextEyebrow');
    if (eyebrow) eyebrow.textContent = state.host === 'Word' ? 'DOCUMENT CONTEXT · AUTO' : 'WORKBOOK CONTEXT · AUTO';
    const brand = $('brandHost');
    if (brand) brand.textContent = state.host === 'Word' ? 'Revising this document' : 'Revising this workbook';
    refreshContextCard();
    watchSelection();
  }
  applyHostedChrome();
  refreshSettings();
  ensureKeys().then(refreshSettings);
});
applyHostedChrome();
