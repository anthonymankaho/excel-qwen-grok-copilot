import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const source = [
  fs.readFileSync(path.join(root, 'public', 'agent-core.js'), 'utf8'),
  fs.readFileSync(path.join(root, 'public', 'taskpane.js'), 'utf8')
].join('\n');

function makeElement(id) {
  const el = {
    id, value: '', textContent: '', placeholder: '', disabled: false, checked: false,
    style: {}, dataset: {}, className: '', children: [], parent: null, listeners: {},
    scrollTop: 0, scrollHeight: 0,
    classList: {
      set: new Set(),
      toggle(c, force) { const on = force === undefined ? !this.set.has(c) : force; on ? this.set.add(c) : this.set.delete(c); },
      add(c) { this.set.add(c); }, remove(c) { this.set.delete(c); }, contains(c) { return this.set.has(c); }
    },
    addEventListener(type, fn) { (this.listeners[type] ||= []).push(fn); },
    append(...nodes) { for (const n of nodes) { n.parent = this; this.children.push(n); } },
    querySelector(sel) { return this.children.find(c => String(c.className).split(' ').includes(sel.slice(1))) || null; },
    remove() { if (this.parent) this.parent.children = this.parent.children.filter(c => c !== this); this.parent = null; },
    focus() {}, requestSubmit() { return fire(el, 'submit'); }
  };
  // Real DOM children is an HTMLCollection without .at(); keep it undefined so regressions fail here.
  Object.defineProperty(el.children, 'at', { value: undefined });
  return el;
}
function fire(el, type, ev = {}) {
  return Promise.all((el.listeners[type] || []).map(fn => fn({ preventDefault() {}, ...ev })));
}
const flush = () => new Promise(resolve => setImmediate(resolve));

function makeSandbox({ selection, used, reply, seedKeys, cells, missingSheets, location }) {
  const els = new Map();
  const defaults = { provider: 'qwen', activeProvider: 'qwen', qwenModel: 'qwen4-max', qwenEndpoint: 'https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions' };
  const get = id => { if (!els.has(id)) { const el = makeElement(id); if (defaults[id]) el.value = defaults[id]; if (id === 'includeContext' || id === 'rememberKey') el.checked = true; els.set(id, el); } return els.get(id); };
  const store = seedKeys ? { 'eqgc.keys.v1': JSON.stringify(seedKeys) } : {};
  const writes = [];
  const calls = [];
  let streamTurn = 0;
  const cellStore = {};
  for (const [k, v] of Object.entries(cells || {})) cellStore[k] = { value: v.value ?? '', formula: v.formula ?? '' };
  const missing = new Set(missingSheets || []);
  const sheetVisibility = {};
  const range = (obj) => ({
    load() {}, isNullObject: false,
    getCell() {
      const cell = target('Selection', 'A1');
      cell.getResizedRange = (rowDelta, colDelta) => {
        const rows = rowDelta + 1, cols = colDelta + 1;
        return range({
          address: obj.address,
          text: (obj.text || []).slice(0, rows).map(row => (row || []).slice(0, cols)),
          formulas: (obj.formulas || []).slice(0, rows).map(row => (row || []).slice(0, cols)),
          rowCount: rows,
          columnCount: cols
        });
      };
      return cell;
    },
    ...obj
  });
  const target = (sheet, cell) => {
    const key = `${sheet}!${cell}`;
    const cur = () => (cellStore[key] ||= { value: '', formula: '' });
    return {
      load() {},
      get values() { return [[cur().value]]; },
      get formulas() { return [[cur().formula]]; },
      set formulas(v) { cur().formula = v[0][0]; cur().value = v[0][0]; writes.push({ sheet, cell, kind: 'formula', value: v[0][0] }); },
      set values(v) { cur().value = v[0][0]; cur().formula = ''; writes.push({ sheet, cell, kind: 'value', value: v[0][0] }); },
      clear() { cur().value = ''; cur().formula = ''; writes.push({ sheet, cell, kind: 'clear' }); }
    };
  };
  function shiftAddress(addr, rowOffset, colOffset) {
    const top = String(addr).split(':')[0];
    const match = top.match(/^([A-Z]+)(\d+)$/i);
    if (!match) return top;
    let n = 0;
    for (const ch of match[1].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    n += colOffset;
    let letters = '';
    while (n > 0) { const m = (n - 1) % 26; letters = String.fromCharCode(65 + m) + letters; n = Math.floor((n - 1) / 26); }
    return letters + (Number(match[2]) + rowOffset);
  }
  function colNum(letters) {
    let n = 0;
    for (const ch of String(letters).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function colName(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  // A multi-cell range must return a real grid, not a 1x1, or whole-range reads
  // (profiling) cannot be exercised.
  function gridFor(sheetName, addr) {
    const m = String(addr).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m) return null;
    const values = [], formulas = [];
    for (let r = Number(m[2]); r <= Number(m[4]); r++) {
      const vr = [], fr = [];
      for (let c = colNum(m[1]); c <= colNum(m[3]); c++) {
        const cur = cellStore[`${sheetName}!${colName(c)}${r}`] || { value: '', formula: '' };
        vr.push(cur.value); fr.push(cur.formula);
      }
      values.push(vr); formulas.push(fr);
    }
    return { values, formulas };
  }
  function makeWritableRange(sheetName, addr) {
    const cell = target(sheetName, addr);
    const grid = gridFor(sheetName, addr);
    cell.address = `${sheetName}!${addr}`;
    if (grid) {
      Object.defineProperty(cell, 'values', { get: () => grid.values, configurable: true });
      Object.defineProperty(cell, 'formulas', { get: () => grid.formulas, configurable: true });
      cell.rowCount = grid.values.length;
      cell.columnCount = grid.values[0].length;
    }
    cell.getCell = (r, c) => makeWritableRange(sheetName, shiftAddress(addr, r, c));
    const font = { load() {} };
    for (const key of ['bold', 'italic', 'color']) {
      let current;
      Object.defineProperty(font, key, { get() { return current; }, set(v) { current = v; writes.push({ sheet: sheetName, cell: addr, kind: 'font', key, value: v }); } });
    }
    const fill = { load() {} };
    let fillColor;
    Object.defineProperty(fill, 'color', { get() { return fillColor; }, set(v) { fillColor = v; writes.push({ sheet: sheetName, cell: addr, kind: 'fill', value: v }); } });
    const borderState = {};
    const sizeState = { columnWidth: 64, rowHeight: 15 };
    let rowHidden = false;
    let columnHidden = false;
    Object.defineProperty(cell, 'rowHidden', {
      get: () => rowHidden,
      set: v => { rowHidden = v; writes.push({ sheet: sheetName, cell: addr, kind: 'rowHidden', value: v }); }
    });
    Object.defineProperty(cell, 'columnHidden', {
      get: () => columnHidden,
      set: v => { columnHidden = v; writes.push({ sheet: sheetName, cell: addr, kind: 'columnHidden', value: v }); }
    });
    cell.format = {
      font, fill, load() {},
      get columnWidth() { return sizeState.columnWidth; },
      set columnWidth(v) { sizeState.columnWidth = v; writes.push({ sheet: sheetName, cell: addr, kind: 'columnWidth', value: v }); },
      get rowHeight() { return sizeState.rowHeight; },
      set rowHeight(v) { sizeState.rowHeight = v; writes.push({ sheet: sheetName, cell: addr, kind: 'rowHeight', value: v }); },
      autofitColumns() { writes.push({ sheet: sheetName, cell: addr, kind: 'autofitColumns' }); },
      autofitRows() { writes.push({ sheet: sheetName, cell: addr, kind: 'autofitRows' }); },
      borders: {
        getItem(edge) {
          const st = (borderState[edge] ||= { style: 'None', weight: 'Hairline', color: '#000000' });
          const border = { load() {} };
          for (const key of ['style', 'weight', 'color']) {
            Object.defineProperty(border, key, {
              get: () => st[key],
              set: v => { st[key] = v; writes.push({ sheet: sheetName, cell: addr, kind: 'border', edge, key, value: v }); }
            });
          }
          return border;
        }
      }
    };
    let cfSeq = 0;
    cell.conditionalFormats = {
      add(type) {
        const id = 'cf' + (++cfSeq);
        const rec = { sheet: sheetName, cell: addr, kind: 'conditionalFormat', id, type, rule: null, fill: null, fontColor: null, bold: null };
        writes.push(rec);
        return {
          load() {}, get id() { return id; },
          cellValue: {
            set rule(v) { rec.rule = v; },
            format: { fill: { set color(v) { rec.fill = v; } }, font: { set color(v) { rec.fontColor = v; }, set bold(v) { rec.bold = v; } } }
          },
          delete() { writes.push({ kind: 'deleteConditionalFormat', id }); }
        };
      },
      clearAll() { writes.push({ sheet: sheetName, cell: addr, kind: 'clearConditionalFormats' }); },
      getItem(id) { return { delete() { writes.push({ kind: 'deleteConditionalFormat', id }); } }; }
    };
    cell.dataValidation = {
      set rule(v) { writes.push({ sheet: sheetName, cell: addr, kind: 'dataValidation', rule: v }); },
      set ignoreBlanks(v) { writes.push({ sheet: sheetName, cell: addr, kind: 'dataValidationIgnoreBlanks', value: v }); },
      clear() { writes.push({ sheet: sheetName, cell: addr, kind: 'clearDataValidation' }); }
    };
    let numberFormat = [['']];
    Object.defineProperty(cell, 'numberFormat', {
      get() { return numberFormat; },
      set(v) { numberFormat = v; writes.push({ sheet: sheetName, cell: addr, kind: 'numberFormat', value: v }); }
    });
    let align = 'General';
    Object.defineProperty(cell.format, 'horizontalAlignment', {
      get() { return align; },
      set(v) { align = v; writes.push({ sheet: sheetName, cell: addr, kind: 'align', value: v }); }
    });
    cell.insert = direction => { writes.push({ sheet: sheetName, cell: addr, kind: 'insert', direction }); };
    cell.delete = direction => { writes.push({ sheet: sheetName, cell: addr, kind: 'delete', direction }); };
    cell.sort = { apply(fields, _match, hasHeaders) { writes.push({ sheet: sheetName, cell: addr, kind: 'sort', fields, hasHeaders }); } };
    return cell;
  }
  function makeSheet(name) {
    let currentName = name;
    return {
      isNullObject: missing.has(name),
      load() {},
      get name() { return currentName; },
      set name(v) { writes.push({ kind: 'renameSheet', from: currentName, name: v }); currentName = v; },
      delete() { writes.push({ sheet: name, kind: 'deleteSheet' }); },
      get visibility() { return sheetVisibility[name] || 'Visible'; },
      set visibility(v) { sheetVisibility[name] = v; writes.push({ sheet: name, kind: 'visibility', value: v }); },
      getRange: addr => makeWritableRange(name, addr),
      getUsedRangeOrNullObject: () => ({ isNullObject: true, load() {} }),
      charts: {
        add(type) {
          const rec = { sheet: name, kind: 'chart', type, name: '', title: '' };
          writes.push(rec);
          return { set name(v) { rec.name = v; }, get name() { return rec.name; }, title: { set text(v) { rec.title = v; } }, delete() { writes.push({ kind: 'deleteChart', name: rec.name }); } };
        },
        getItem(chartName) { return { delete() { writes.push({ kind: 'deleteChart', name: chartName }); } }; }
      },
      tables: {
        add() { writes.push({ sheet: name, kind: 'table' }); return { set name(v) { writes.push({ sheet: name, kind: 'tableName', name: v }); }, delete() { writes.push({ kind: 'deleteTable' }); } }; },
        getItem(tableName) {
          return {
            delete() { writes.push({ kind: 'deleteTable', name: tableName }); },
            columns: {
              add(index, values, columnName) { writes.push({ sheet: name, kind: 'tableColumn', table: tableName, name: columnName, values }); },
              getItem(columnName) { return { delete() { writes.push({ kind: 'deleteTableColumn', table: tableName, name: columnName }); } }; }
            }
          };
        }
      },
      pivotTables: {
        add(pivotName, source, destination) {
          const rec = {
            sheet: name, kind: 'pivot', name: pivotName,
            source: source.address, destination: destination.address,
            rows: [], columns: [], filters: [], values: [], layout: null, grandTotals: null, refreshed: false
          };
          writes.push(rec);
          const hierarchy = field => ({ field });
          const srcMatch = String(source.address || '').match(/^(.*)!([A-Z]+\d+:[A-Z]+\d+)$/);
          const headerNames = srcMatch
            ? (gridFor(srcMatch[1], srcMatch[2])?.values[0] || []).map(v => String(v ?? '').trim()).filter(Boolean)
            : [];
          return {
            layout: {
              set layoutType(v) { rec.layout = v; },
              set showRowGrandTotals(v) { rec.grandTotals = v; },
              set showColumnGrandTotals(v) { rec.grandTotals = v; }
            },
            hierarchies: { load() {}, items: headerNames.map(name => ({ name })), getItem: hierarchy },
            rowHierarchies: { add(h) { rec.rows.push(h.field); } },
            columnHierarchies: { add(h) { rec.columns.push(h.field); } },
            filterHierarchies: { add(h) { rec.filters.push(h.field); } },
            dataHierarchies: { add(h) { const d = { field: h.field, summarizeBy: '', name: '' }; rec.values.push(d); return d; } },
            refresh() { rec.refreshed = true; },
            delete() { writes.push({ kind: 'deletePivot', name: pivotName }); }
          };
        },
        getItem(pivotName) { return { delete() { writes.push({ kind: 'deletePivot', name: pivotName }); } }; }
      },
      autoFilter: {
        apply(range, columnIndex, criteria) { writes.push({ sheet: name, kind: 'filter', range: range.address, columnIndex, criteria }); },
        clearCriteria() { writes.push({ sheet: name, kind: 'clearFilters' }); },
        clearColumnCriteria(columnIndex) { writes.push({ sheet: name, kind: 'clearFilterColumn', columnIndex }); },
        reapply() { writes.push({ sheet: name, kind: 'reapplyFilter' }); }
      },
      freezePanes: {
        freezeRows(count) { writes.push({ sheet: name, kind: 'freezeRows', count }); },
        freezeColumns(count) { writes.push({ sheet: name, kind: 'freezeColumns', count }); },
        unfreeze() { writes.push({ sheet: name, kind: 'unfreeze' }); },
        freezeAt(addr) { writes.push({ sheet: name, kind: 'freezeAt', address: addr }); },
        getLocationOrNullObject() { return { isNullObject: true, address: '', load() {} }; }
      },
      replaceAll(text, replacement, criteria) {
        writes.push({ sheet: name, kind: 'replaceAll', text, replacement, criteria });
        return { value: 1 };
      }
    };
  }
  const wordFont = { load() {} };
  for (const key of ['bold', 'italic']) {
    let current = false;
    Object.defineProperty(wordFont, key, { get() { return current; }, set(v) { current = v; writes.push({ kind: 'wordFont', key, value: v }); } });
  }
  const wordDoc = { selectionText: 'Hello team', bodyText: 'Hello team. More text.', style: 'Normal', font: wordFont };
  const sandbox = {
    console, JSON, Math, String, Array, Object, Promise, RegExp, Number, TextEncoder, TextDecoder, AbortController,
    setTimeout: () => 0, clearTimeout: () => {},
    localStorage: {
      getItem: k => (k in store ? store[k] : null),
      setItem: (k, v) => { store[k] = String(v); },
      removeItem: k => { delete store[k]; }
    },
    document: { getElementById: get, createElement: tag => makeElement(tag), querySelectorAll: () => [] },
    Office: { HostType: { Excel: 'Excel', Word: 'Word' }, onReady(cb) { sandbox.officeCb = cb; } },
    Excel: {
      run: fn => fn({
        workbook: {
          getSelectedRange: () => range(selection),
          name: 'Book1',
          load() {},
          names: {
            add(name, reference) { writes.push({ kind: 'namedRange', name, reference: reference.address }); return { delete() { writes.push({ kind: 'deleteNamedRange', name }); } }; },
            getItem(name) { return { delete() { writes.push({ kind: 'deleteNamedRange', name }); } }; }
          },
          worksheets: {
            load() {},
            items: [{ name: 'Sheet1', load() {} }],
            getActiveWorksheet: () => ({ getUsedRangeOrNullObject: () => range(used) }),
            add(sheetName) { writes.push({ kind: 'addSheet', name: sheetName }); return makeSheet(sheetName); },
            getItem: sheet => makeSheet(sheet),
            getItemOrNullObject: sheet => makeSheet(sheet)
          }
        },
        sync: async () => {}
      })
    },
    Word: {
      run: fn => fn({
        document: {
          getSelection: () => ({
            get text() { return wordDoc.selectionText; },
            load() {},
            insertText(text, where) { writes.push({ kind: 'wordInsert', where, text }); if (where === 'Replace') wordDoc.selectionText = text; },
            font: wordDoc.font,
            get style() { return wordDoc.style; },
            set style(v) { wordDoc.style = v; writes.push({ kind: 'wordStyle', style: v }); }
          }),
          body: {
            get text() { return wordDoc.bodyText; },
            load() {},
            insertText(text, where) { writes.push({ kind: 'wordBody', where, text }); },
            search(term, options) {
              writes.push({ kind: 'wordSearch', term, options });
              return { load() {}, items: [{ insertText(text, where) { writes.push({ kind: 'wordReplace', term, text, where }); } }] };
            }
          }
        },
        sync: async () => {}
      })
    },
    fetch: async (url, opts) => {
      const body = opts?.body ? JSON.parse(opts.body) : null;
      calls.push({ url, body, headers: opts?.headers || {} });
      if (url.endsWith('/api/health')) return { ok: true, status: 200, json: async () => ({ ok: true, configured: { qwen: true, grok: false } }) };
      const turnText = Array.isArray(reply?.turns) ? reply.turns[Math.min(streamTurn, reply.turns.length - 1)] : null;
      const data = turnText != null ? { text: turnText, model: reply.model || 'test-model' } : (reply || { text: 'mock answer', model: 'test-model' });
      if (url.endsWith('/api/chat-stream')) {
        streamTurn += 1;
        const frames = `data: ${JSON.stringify({ model: data.model || 'test-model', choices: [{ delta: { content: data.text } }] })}\n\ndata: [DONE]\n\n`;
        let sent = false;
        return {
          ok: true, status: 200,
          body: { getReader: () => ({ read: async () => sent ? { done: true } : (sent = true, { done: false, value: new TextEncoder().encode(frames) }) }) }
        };
      }
      return { ok: true, status: 200, json: async () => data };
    }
  };
  if (location) sandbox.location = location;
  sandbox.writes = writes; sandbox.calls = calls; sandbox.els = get; sandbox.store = store;
  vm.createContext(sandbox);
  vm.runInContext(source, sandbox);
  return sandbox;
}

const singleEmpty = { address: "'Sheet1'!A1", text: [['']], formulas: [['']], rowCount: 1, columnCount: 1 };
const usedSmall = { address: "'Sheet1'!A1:B2", text: [['1', '2'], ['3', '4']], formulas: [['=1+0', '=A1*2'], ['=3', '=4']] };
const selMulti = { address: "'Sheet1'!A1:B2", text: [['1', '2'], ['3', '4']], formulas: [['=1+0', '=A1*2'], ['=3', '=4']], rowCount: 2, columnCount: 2 };
const chatCalls = sb => sb.calls.filter(c => c.url.endsWith('/api/chat-stream') || c.url.endsWith('/api/chat'));
const lastContent = sb => chatCalls(sb).at(-1).body.messages.at(-1).content;
const visible = el => el ? [el.textContent || '', ...(el.children || []).map(visible)].join('\n') : '';
async function boot(sb) { sb.officeCb({ host: 'Excel' }); await flush(); await flush(); }
async function ask(sb, text) { sb.els('prompt').value = text; await fire(sb.els('composer'), 'submit'); await flush(); await flush(); }

test('auto context: empty single cell falls back to used range, formulas included', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb); await ask(sb, 'review');
  assert.match(lastContent(sb), /\[Excel used range 'Sheet1'!A1:B2\]/);
  assert.match(lastContent(sb), /\[=A1\*2\]/);
});

test('auto context: multi-cell selection wins over used range', async () => {
  const sb = makeSandbox({ selection: selMulti, used: usedSmall });
  await boot(sb); await ask(sb, 'review');
  assert.match(lastContent(sb), /\[Excel selection 'Sheet1'!A1:B2\]/);
});

test('Send context off sends no workbook data', async () => {
  const sb = makeSandbox({ selection: selMulti, used: usedSmall });
  await boot(sb);
  sb.els('includeContext').checked = false;
  await fire(sb.els('includeContext'), 'change'); await flush();
  await ask(sb, 'review');
  assert.equal(lastContent(sb), 'review');
});

test('Whole sheet chip forces used range for one question only', async () => {
  const sb = makeSandbox({ selection: selMulti, used: usedSmall });
  await boot(sb);
  await fire(sb.els('addUsedRange'), 'click'); await flush();
  await ask(sb, 'first');
  assert.match(lastContent(sb), /\[Excel used range/);
  await ask(sb, 'second');
  assert.match(lastContent(sb), /\[Excel selection/);
});

test('oversize context is truncated with an explicit marker inside the cap', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: { address: "'Sheet1'!A1", text: [['x'.repeat(60_000)]], formulas: [[null]] } });
  await boot(sb); await ask(sb, 'review');
  const content = lastContent(sb);
  assert.ok(content.length <= 49_500);
  assert.match(content, /\[\.\.\.truncated: /);
  assert.match(sb.els('contextHint').textContent, /truncated/);
});

test('edits block renders a proposal; Apply writes formulas/values only on click', async () => {
  const reply = { text: 'Fix D5.\n```json\n{"edits":[{"sheet":"Sheet1","cell":"D5","value":"=SUM(A1:A4)"},{"sheet":"Sheet1","cell":"B2","value":42}]}\n```', model: 'test-model' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb); await ask(sb, 'fix D5');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  assert.ok(card, 'proposal card rendered');
  assert.equal(sb.writes.length, 0, 'nothing written before click');
  const prose = sb.els('conversation').children.find(c => String(c.className).includes('message') && c.children.some(b => b.textContent === 'Fix D5.'));
  assert.ok(prose, 'json block stripped from chat prose');
  const actions = card.children.find(c => c.children.length === 2);
  await fire(actions.children[0], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes, [
    { sheet: 'Sheet1', cell: 'D5', kind: 'formula', value: '=SUM(A1:A4)' },
    { sheet: 'Sheet1', cell: 'B2', kind: 'value', value: 42 }
  ]);
  assert.match(card.children[0].textContent, /APPLIED 2 EDITS/);
});

test('Dismiss removes the proposal without writing', async () => {
  const reply = { text: 'x\n```json\n{"edits":[{"sheet":"Sheet1","cell":"A1","value":1}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb); await ask(sb, 'change A1');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  const actions = card.children.find(c => c.children.length === 2);
  await fire(actions.children[1], 'click'); await flush();
  assert.equal(card.parent, null);
  assert.equal(sb.writes.length, 0);
});

test('invalid edits payload is ignored and prose kept intact', async () => {
  const reply = { text: 'Hello\n```json\n{"edits":[{"sheet":"S","cell":"not-a-cell","value":1}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb); await ask(sb, 'hi');
  assert.equal(sb.els('conversation').children.some(c => String(c.className).includes('edits-card')), false);
  const bubble = sb.els('conversation').children.find(c => String(c.className).includes('assistant') && c.children.some(b => String(b.textContent).includes('Hello')));
  assert.ok(bubble, 'assistant prose rendered');
  assert.match(bubble.children[bubble.children.length - 1].textContent, /```json/);
});

test('oversize prompt is blocked before any provider call', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb);
  const before = chatCalls(sb).length;
  await ask(sb, 'x'.repeat(49_600));
  assert.equal(chatCalls(sb).length, before);
  assert.match(sb.els('toast').textContent, /too large/);
});

test('remembered key is auto-configured into the relay on pane load', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, seedKeys: { qwen: 'sk-sp-remembered' } });
  await boot(sb); await flush();
  const cfg = sb.calls.find(c => c.url.endsWith('/api/configure'));
  assert.ok(cfg, 'configure called without user action');
  assert.deepEqual(cfg.body, { provider: 'qwen', apiKey: 'sk-sp-remembered' });
});

test('Use key persists only while Remember is ticked; Clear key removes storage', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb);
  sb.els('apiKey').value = 'sk-sp-once';
  await fire(sb.els('saveKey'), 'click'); await flush();
  assert.equal(JSON.parse(sb.store['eqgc.keys.v1']).qwen, 'sk-sp-once');
  sb.els('rememberKey').checked = false;
  sb.els('apiKey').value = 'sk-sp-session';
  await fire(sb.els('saveKey'), 'click'); await flush();
  assert.equal(JSON.parse(sb.store['eqgc.keys.v1']).qwen, undefined);
  sb.els('rememberKey').checked = true;
  sb.els('apiKey').value = 'sk-sp-again';
  await fire(sb.els('saveKey'), 'click'); await flush();
  await fire(sb.els('clearKey'), 'click'); await flush();
  assert.deepEqual(JSON.parse(sb.store['eqgc.keys.v1'] || '{}'), {});
  assert.ok(sb.calls.some(c => c.url.endsWith('/api/clear-key')));
});

test('history context dedup: only the latest user message carries the Excel block', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb);
  await ask(sb, 'first');
  await ask(sb, 'second');
  const msgs = chatCalls(sb).at(-1).body.messages;
  const users = msgs.filter(m => m.role === 'user');
  assert.equal(users.length, 2);
  assert.ok(!users[0].content.includes('[Excel'), 'older context stripped');
  assert.match(users[1].content, /\[Excel used range/);
});

test('Fast toggle is sent to the relay as fast flag', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb);
  await ask(sb, 'quick');
  assert.equal(chatCalls(sb).at(-1).body.fast, false);
  sb.els('fastMode').checked = true;
  await ask(sb, 'quicker');
  assert.equal(chatCalls(sb).at(-1).body.fast, true);
});

test('edit card shows old → new; Apply then Undo restores originals', async () => {
  const reply = { text: 'Fix.\n```json\n{"edits":[{"sheet":"Sheet1","cell":"D5","value":"=SUM(A1:A4)"},{"sheet":"Sheet1","cell":"B2","value":42}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells: { 'Sheet1!D5': { value: 123 }, 'Sheet1!B2': { value: 2 } } });
  await boot(sb); await ask(sb, 'fix');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  const rows = card.children.filter(c => c.className === 'edit-row');
  assert.match(rows[0].textContent, /123 → =SUM\(A1:A4\)/);
  assert.match(rows[1].textContent, /2 → 42/);
  const actions = card.children.find(c => c.children.length === 2);
  await fire(actions.children[0], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-2), [
    { sheet: 'Sheet1', cell: 'D5', kind: 'formula', value: '=SUM(A1:A4)' },
    { sheet: 'Sheet1', cell: 'B2', kind: 'value', value: 42 }
  ]);
  const undoBtn = actions.children[2];
  assert.ok(undoBtn, 'undo button appears after apply');
  await fire(undoBtn, 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-2), [
    { sheet: 'Sheet1', cell: 'D5', kind: 'value', value: 123 },
    { sheet: 'Sheet1', cell: 'B2', kind: 'value', value: 2 }
  ]);
  assert.match(card.children[0].textContent, /UNDONE — RESTORED 2 CELLS/);
});

test('undo of an originally empty cell clears it', async () => {
  const reply = { text: 'x\n```json\n{"edits":[{"sheet":"Sheet1","cell":"A1","value":7}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb); await ask(sb, 'set A1');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  const rows = card.children.filter(c => c.className === 'edit-row');
  assert.match(rows[0].textContent, /\(empty\) → 7/);
  const actions = card.children.find(c => c.children.length === 2);
  await fire(actions.children[0], 'click'); await flush(); await flush();
  await fire(actions.children[2], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', cell: 'A1', kind: 'clear' }]);
});

test('absolute cell refs are normalized and invalid edits are skipped', async () => {
  const reply = { text: 'Set two.\n```json\n{"edits":[{"sheet":"Sheet1","cell":"$D$5","value":"=A1+1"},{"sheet":"Sheet1","cell":"not-a-cell","value":1}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells: { 'Sheet1!D5': { value: 4 } } });
  await boot(sb); await ask(sb, 'set');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  const rows = card.children.filter(c => c.className === 'edit-row');
  assert.equal(rows.length, 1);
  assert.match(rows[0].textContent, /Sheet1!D5: 4 → =A1\+1/);
  assert.match(card.children.find(c => c.className === 'edit-note').textContent, /1 edit skipped/);
  const actions = card.children.find(c => c.className === 'button-row');
  await fire(actions.children[0], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', cell: 'D5', kind: 'formula', value: '=A1+1' }]);
});

test('risky formulas stay click-gated and are called out on the card', async () => {
  const reply = { text: 'Pull.\n```json\n{"edits":[{"sheet":"Sheet1","cell":"A1","value":"=WEBSERVICE(\\"https://example.com\\")"}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb); await ask(sb, 'pull');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  assert.match(card.children.find(c => c.className === 'edit-warning').textContent, /WEBSERVICE/);
  assert.equal(sb.writes.length, 0);
  const actions = card.children.find(c => c.className === 'button-row');
  await fire(actions.children[0], 'click'); await flush(); await flush();
  assert.equal(sb.writes.at(-1).kind, 'formula');
  assert.match(sb.writes.at(-1).value, /WEBSERVICE/);
});

test('a missing sheet is shown and nothing is written', async () => {
  const reply = { text: 'x\n```json\n{"edits":[{"sheet":"Gone","cell":"A1","value":1}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, missingSheets: ['Gone'] });
  await boot(sb); await ask(sb, 'set');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  assert.match(card.children.find(c => c.className === 'edit-row').textContent, /sheet missing/);
  const actions = card.children.find(c => c.className === 'button-row');
  await fire(actions.children[0], 'click'); await flush(); await flush();
  assert.equal(sb.writes.length, 0);
  assert.match(sb.els('toast').textContent, /not in this workbook/);
});

test('insert writes a single formula as a formula', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: '=SUM(A1:A4)', model: 'm' } });
  await boot(sb); await ask(sb, 'formula');
  await fire(sb.els('insertAnswer'), 'click'); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Selection', cell: 'A1', kind: 'formula', value: '=SUM(A1:A4)' }]);
});

test('oversized sheets are read from the top-left grid only', async () => {
  const text = Array.from({ length: 200 }, (_, i) => [i === 180 ? 'UNIQUE_TAIL' : 'a', 'b']);
  const formulas = text.map(() => ['', '']);
  const sb = makeSandbox({ selection: singleEmpty, used: { address: "'Sheet1'!A1:B200", text, formulas, rowCount: 200, columnCount: 2 } });
  await boot(sb); await ask(sb, 'review');
  const content = lastContent(sb);
  assert.equal(content.includes('UNIQUE_TAIL'), false);
  assert.match(content, /150 of 200 rows/);
  assert.match(sb.els('contextHint').textContent, /truncated/);
});

test('cell tabs and newlines stay inside one context row', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: { address: "'Sheet1'!A1:B1", text: [['a\nb\tc', 'ok']], formulas: [['', '']] } });
  await boot(sb); await ask(sb, 'review');
  assert.match(lastContent(sb), /a b c\tok/);
  assert.equal(lastContent(sb).includes('\nb'), false);
});

test('an empty provider stream is shown as an error', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: '', model: 'm' } });
  await boot(sb); await ask(sb, 'hi');
  const bubble = sb.els('conversation').children.find(c => String(c.className).includes('assistant'));
  assert.match(bubble.children[bubble.children.length - 1].textContent, /no text answer/);
  assert.equal(sb.els('insertAnswer').disabled, true);
});

test('New chat drops previous turns from the next request', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb);
  await ask(sb, 'first');
  await fire(sb.els('newChat'), 'click');
  await ask(sb, 'second');
  const msgs = chatCalls(sb).at(-1).body.messages;
  assert.equal(msgs.length, 1);
  assert.match(msgs[0].content, /^second/);
});

test('Direct revise writes a range immediately and Undo clears it', async () => {
  const reply = { text: 'Fill the column.\n```json\n{"actions":[{"op":"set","sheet":"Sheet1","range":"B2:B3","values":[[10],[20]]}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('directRevise').checked = true;
  await ask(sb, 'fill');
  assert.equal(chatCalls(sb).at(-1).body.host, 'excel');
  assert.ok(sb.writes.some(w => w.cell === 'B2' && w.kind === 'value' && w.value === 10));
  assert.ok(sb.writes.some(w => w.cell === 'B3' && w.kind === 'value' && w.value === 20));
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  assert.match(card.children[0].textContent, /REVISED 1/);
  const undo = card.children.find(c => c.className === 'button-row').children.find(c => /Undo/.test(c.textContent));
  await fire(undo, 'click'); await flush(); await flush();
  assert.ok(sb.writes.filter(w => w.kind === 'clear' && (w.cell === 'B2' || w.cell === 'B3')).length >= 2);
});

test('Direct off waits for Apply before writing a revision', async () => {
  const reply = { text: 'Set.\n```json\n{"actions":[{"op":"set","sheet":"Sheet1","range":"C3","value":"=A1+1"}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('directRevise').checked = false;
  await ask(sb, 'set');
  assert.equal(sb.writes.length, 0);
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  const apply = card.children.find(c => c.className === 'button-row').children[0];
  await fire(apply, 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', cell: 'C3', kind: 'formula', value: '=A1+1' }]);
});

test('Direct revise holds a network formula for a click', async () => {
  const reply = { text: 'Pull.\n```json\n{"actions":[{"op":"set","sheet":"Sheet1","range":"A1","value":"=WEBSERVICE(\\"https://example.com\\")"}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('directRevise').checked = true;
  await ask(sb, 'pull');
  assert.equal(sb.writes.length, 0);
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  assert.match(card.children.find(c => c.className === 'edit-warning').textContent, /network/);
});

test('Direct revise formats, inserts, charts, and adds a sheet', async () => {
  const reply = { text: 'Shape it.\n```json\n{"actions":[{"op":"format","sheet":"Sheet1","range":"A1:B1","bold":true,"fill":"#1F4E79","fontColor":"#FFFFFF","numberFormat":"#,##0","align":"center"},{"op":"insertRows","sheet":"Sheet1","row":5,"count":1},{"op":"addSheet","name":"Summary"},{"op":"chart","sheet":"Sheet1","range":"A1:B2","type":"column","title":"Sales"},{"op":"renameSheet","sheet":"Sheet1","name":"Data"}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('directRevise').checked = true;
  await ask(sb, 'shape');
  assert.ok(sb.writes.some(w => w.kind === 'font' && w.key === 'bold' && w.value === true));
  assert.ok(sb.writes.some(w => w.kind === 'fill' && w.value === '#1F4E79'));
  assert.ok(sb.writes.some(w => w.kind === 'numberFormat'));
  assert.ok(sb.writes.some(w => w.kind === 'insert' && w.direction === 'Down'));
  assert.ok(sb.writes.some(w => w.kind === 'addSheet' && w.name === 'Summary'));
  assert.ok(sb.writes.some(w => w.kind === 'chart' && w.type === 'ColumnClustered' && w.title === 'Sales'));
  assert.ok(sb.writes.some(w => w.kind === 'renameSheet' && w.name === 'Data'));
});

test('Direct revise replaces and styles the Word selection', async () => {
  const reply = { text: 'Done.\n```json\n{"actions":[{"op":"replaceSelection","text":"Revised paragraph."},{"op":"formatSelection","bold":true},{"op":"styleSelection","style":"Heading 1"},{"op":"insertText","text":" Added.","where":"end"}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  sb.officeCb({ host: 'Word' });
  await flush(); await flush();
  sb.els('directRevise').checked = true;
  await ask(sb, 'revise');
  assert.equal(chatCalls(sb).at(-1).body.host, 'word');
  assert.match(lastContent(sb), /\[Word selection\]/);
  assert.ok(sb.writes.some(w => w.kind === 'wordInsert' && w.where === 'Replace' && w.text === 'Revised paragraph.'));
  assert.ok(sb.writes.some(w => w.kind === 'wordFont' && w.key === 'bold' && w.value === true));
  assert.ok(sb.writes.some(w => w.kind === 'wordStyle' && w.style === 'Heading 1'));
  assert.ok(sb.writes.some(w => w.kind === 'wordBody' && w.where === 'End'));
});

test('Chat mode answers and does not write, even when Direct is on', async () => {
  const reply = { text: 'Here is the total.\n```json\n{"actions":[{"op":"set","sheet":"Sheet1","range":"B2","value":5}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'chat';
  sb.els('directRevise').checked = true;
  await ask(sb, 'what is the total');
  assert.equal(sb.writes.length, 0);
  assert.equal(sb.els('conversation').children.some(c => String(c.className).includes('edits-card')), false);
  assert.match(`${sb.els('toast').textContent}\n${visible(sb.els('conversation'))}`, /Chat mode does not change the file|AUDIT/);
});

test('Plan mode shows the plan and writes only after Approve', async () => {
  const reply = { text: 'Proposal.\n```json\n{"plan":{"goal":"Fill total","mode":"plan","steps":[{"id":1,"objective":"Write the total","actionType":"edit","requiresApproval":true,"validation":"value is 99"}],"assumptions":["Header is row 1"],"requiredInputs":[],"risks":[],"finalValidation":[]},"actions":[{"op":"set","sheet":"Sheet1","range":"B2","value":99}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'plan';
  sb.els('directRevise').checked = true;
  await ask(sb, 'fill the total');
  assert.equal(sb.writes.length, 0);
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  assert.match(visible(card), /Fill total/);
  const approve = card.children.find(c => c.className === 'button-row').children.find(c => c.textContent === 'Approve');
  await fire(approve, 'click'); await flush(); await flush();
  assert.ok(sb.writes.some(w => w.cell === 'B2' && w.kind === 'value' && w.value === 99));
});

test('Agent mode sends workbook metadata and holds a delete for approval', async () => {
  const reply = { text: 'I would remove that row.\n```json\n{"plan":{"goal":"Remove a blank row","mode":"agent","steps":[{"id":1,"objective":"Delete row 5","actionType":"edit","requiresApproval":true,"validation":"row count"}],"assumptions":[],"requiredInputs":[],"risks":["Deletes data"],"finalValidation":[]},"actions":[{"op":"deleteRows","sheet":"Sheet1","row":5,"count":1}],"decision":"COMPLETE"}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'delete the blank row');
  assert.match(lastContent(sb), /\[Excel workbook metadata\]/);
  assert.equal(lastContent(sb).includes('[=A1*2]'), false);
  assert.equal(sb.writes.some(w => w.kind === 'delete'), false);
  assert.match(visible(sb.els('conversation')), /Delete 1 row/);
});

test('Agent mode stops when debit does not equal credit and does not continue', async () => {
  const reply = {
    model: 'm',
    turns: [
      'Draft.\n```json\n{"plan":{"goal":"Draft the journal","mode":"agent","steps":[{"id":1,"objective":"Check the balance","actionType":"validate","requiresApproval":false,"validation":"debit equals credit"}],"assumptions":["This is a draft"],"requiredInputs":[],"risks":["Do not post"],"finalValidation":[{"check":"debitEqualsCredit","sheet":"Sheet1","range":"C1:D2"}]},"actions":[{"op":"set","sheet":"Sheet1","range":"E1","value":1}],"decision":"CONTINUE"}\n```',
      'Again.\n```json\n{"actions":[{"op":"set","sheet":"Sheet1","range":"F1","value":2}],"decision":"COMPLETE"}\n```'
    ]
  };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells: { 'Sheet1!C1': { value: 'Debit' }, 'Sheet1!D1': { value: 'Credit' }, 'Sheet1!C2': { value: 100 }, 'Sheet1!D2': { value: 40 } } });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'draft the accrual journal');
  assert.equal(chatCalls(sb).length, 1);
  assert.ok(sb.writes.some(w => w.cell === 'E1' && w.value === 1));
  assert.equal(sb.writes.some(w => w.cell === 'F1'), false);
  assert.match(visible(sb.els('conversation')), /FAIL debitEqualsCredit/);
  assert.match(visible(sb.els('conversation')), /draft/i);
});

test('Agent mode asks for columns when debit and credit are not identified', async () => {
  const reply = {
    model: 'm',
    turns: [
      'Draft.\n```json\n{"plan":{"goal":"Draft the journal","mode":"agent","steps":[{"id":1,"objective":"Check the balance","actionType":"validate","requiresApproval":false,"validation":"debit equals credit"}],"assumptions":[],"requiredInputs":[],"risks":[],"finalValidation":[{"check":"debitEqualsCredit","sheet":"Sheet1","range":"C2:D2"}]},"actions":[{"op":"set","sheet":"Sheet1","range":"E1","value":1}],"decision":"CONTINUE"}\n```',
      'Again.\n```json\n{"actions":[{"op":"set","sheet":"Sheet1","range":"F1","value":2}],"decision":"COMPLETE"}\n```'
    ]
  };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells: { 'Sheet1!C2': { value: 100 }, 'Sheet1!D2': { value: 40 } } });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'draft the accrual journal');
  assert.equal(chatCalls(sb).length, 1);
  assert.equal(sb.writes.some(w => w.cell === 'F1'), false);
  assert.match(visible(sb.els('conversation')), /NEEDS_USER debitEqualsCredit/);
  assert.equal(visible(sb.els('conversation')).includes('FAIL debitEqualsCredit'), false);
});

test('Agent mode with Direct off waits for Approve before writing', async () => {
  const reply = { text: 'Ready.\n```json\n{"plan":{"goal":"Write the total","mode":"agent","steps":[{"id":1,"objective":"Set the total","actionType":"edit","requiresApproval":true,"validation":"value is 7"}],"assumptions":[],"requiredInputs":[],"risks":[],"finalValidation":[]},"actions":[{"op":"set","sheet":"Sheet1","range":"B2","value":7}],"decision":"COMPLETE"}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = false;
  await ask(sb, 'write the total');
  assert.equal(sb.writes.length, 0);
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  assert.match(visible(card), /Write the total/);
  const approve = card.children.find(c => c.className === 'button-row').children.find(c => c.textContent === 'Approve');
  await fire(approve, 'click'); await flush();
  assert.equal(sb.writes.length, 0);
  const applyCard = sb.els('conversation').children.filter(c => String(c.className).includes('edits-card')).at(-1);
  const apply = applyCard.children.find(c => c.className === 'button-row').children.find(c => c.textContent === 'Apply edits');
  await fire(apply, 'click'); await flush();
  assert.ok(sb.writes.some(w => w.cell === 'B2' && w.value === 7));
});

test('Auto stays on Qwen when Grok has no key', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.', model: 'qwen4-max' } });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('activeProvider').value = 'auto';
  await ask(sb, 'Reconcile the ledger');
  assert.equal(chatCalls(sb).at(-1).body.provider, 'qwen');
  assert.equal(chatCalls(sb).at(-1).body.mode, 'agent');
  assert.match(visible(sb.els('conversation')), /via qwen/);
});

test('A hosted pane sends the relay token and does not upload a provider key', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.', model: 'qwen4-max' }, location: { hostname: 'excel-qwen-grok.example.workers.dev' } });
  await boot(sb);
  sb.els('relayToken').value = 'relay-token-0123456789';
  sb.els('rememberKey').checked = true;
  await fire(sb.els('saveKey'), 'click');
  await flush();
  assert.equal(sb.store['eqgc.relay.v1'], 'relay-token-0123456789');
  assert.equal(sb.calls.some(call => String(call.url).includes('/api/configure')), false);
  await ask(sb, 'hello from cloud');
  assert.equal(chatCalls(sb).at(-1).headers.Authorization, 'Bearer relay-token-0123456789');
});

test('Agent read-back delivers every column of a wide sheet, not just A to H', async () => {
  const names = ['Entity', 'Customer', 'Invoice', 'Currency', 'Rate', 'Current', '31-60', '61-90', '91-120', '121-150', '151-180', '181-270', '271-365', 'Over 365', 'Total Local', 'FX Adj', 'Total USD'];
  const cells = {};
  names.forEach((name, i) => { cells[`Aging!${String.fromCharCode(65 + i)}1`] = { value: name }; });
  const reply = {
    model: 'm',
    turns: [
      'Need the layout.\n```json\n{"plan":"Read the header row first.","actions":[],"contextRequest":{"sheet":"Aging","range":"A1:Q2"},"decision":"CONTINUE"}\n```',
      'Got it.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```'
    ]
  };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'draft a pivot style summary for Aging');
  assert.equal(chatCalls(sb).length, 2);
  const fed = lastContent(sb);
  assert.match(fed, /\[Excel range Aging!A1:Q2\]/);
  assert.match(fed, /Over 365/);
  assert.match(fed, /Total USD/);
  assert.match(visible(sb.els('conversation')), /Context Aging!A1:Q2/);
});

test('An over-limit read request is reported as not read, never as Observed SUCCESS', async () => {
  const reply = {
    model: 'm',
    turns: [
      'Reading.\n```json\n{"actions":[],"contextRequest":{"sheet":"Aging","range":"A1:Z200"},"decision":"CONTINUE"}\n```',
      'Understood.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```'
    ]
  };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'review the whole ageing sheet');
  assert.equal(chatCalls(sb).length, 2);
  const fed = lastContent(sb);
  assert.match(fed, /No range was read/);
  assert.match(fed, /5200 cells, over the 2000-cell read limit/);
  assert.equal(/Observed: SUCCESS/.test(fed), false);
  assert.equal(sb.writes.length, 0);
});

test('A profile request returns column statistics, not raw cells', async () => {
  const cells = {
    'Aging!A1': { value: 'Entity' }, 'Aging!B1': { value: 'Bucket' }, 'Aging!C1': { value: 'Total USD' },
    'Aging!A2': { value: 'MUK' }, 'Aging!B2': { value: 'Current' }, 'Aging!C2': { value: 100 },
    'Aging!A3': { value: 'MUK' }, 'Aging!B3': { value: '31-60' }, 'Aging!C3': { value: 200 },
    'Aging!A4': { value: 'MCA' }, 'Aging!B4': { value: 'Current' }, 'Aging!C4': { value: 50 }
  };
  const reply = {
    model: 'm',
    turns: [
      'Profiling first.\n```json\n{"actions":[],"profileRequest":{"sheet":"Aging","range":"A1:C4"},"decision":"CONTINUE"}\n```',
      'Now I know the fields.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```'
    ]
  };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'build an ageing pivot');
  assert.equal(chatCalls(sb).length, 2);
  const fed = lastContent(sb);
  assert.match(fed, /\[Excel profile Aging!A1:C4 — 3 data rows sampled, 3 columns\]/);
  assert.match(fed, /A \| Entity \| filled 3 \| distinct 2/);
  assert.match(fed, /MUK \(2\)/);
  assert.match(fed, /C \| Total USD \| filled 3 \| distinct 3 \| numeric 3 \| sum 350 \| min 50 \| max 200/);
  assert.match(visible(sb.els('conversation')), /Context Aging!A1:C4 \(profile\)/);
});

test('A pivot action creates a real native PivotTable with fields and aggregation', async () => {
  const action = {
    op: 'pivot', sheet: 'Aging', range: 'A1:C4', name: 'Aging Pivot',
    destinationSheet: 'Pivot - Aging', destinationCell: 'A3',
    rows: ['Entity'], columns: ['Bucket'], filters: [],
    values: [{ field: 'Total USD', summarizeBy: 'Sum' }], layout: 'Tabular', grandTotals: true
  };
  const reply = { model: 'm', text: 'Building it.\n```json\n{"plan":"Add the pivot.","actions":[' + JSON.stringify(action) + '],"decision":"COMPLETE"}\n```' };
  const cells = { 'Aging!A1': { value: 'Entity' }, 'Aging!B1': { value: 'Bucket' }, 'Aging!C1': { value: 'Total USD' } };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells, missingSheets: ['Pivot - Aging'] });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'draft a pivot table for Aging');
  const pivot = sb.writes.find(w => w.kind === 'pivot');
  assert.ok(pivot, 'a native PivotTable was created');
  assert.equal(pivot.name, 'Aging Pivot');
  assert.equal(pivot.source, 'Aging!A1:C4');
  assert.equal(pivot.destination, 'Pivot - Aging!A3');
  assert.equal(pivot.rows.join(','), 'Entity');
  assert.equal(pivot.columns.join(','), 'Bucket');
  assert.equal(pivot.values.length, 1);
  assert.equal(pivot.values[0].field, 'Total USD');
  assert.equal(pivot.values[0].summarizeBy, 'Sum');
  assert.equal(pivot.layout, 'Tabular');
  assert.equal(pivot.grandTotals, true);
  assert.equal(pivot.refreshed, true);
  assert.ok(sb.writes.some(w => w.kind === 'addSheet' || w.name === 'Pivot - Aging'), 'destination sheet was created');
});

test('A pivot action without value fields or with a bad aggregation is rejected', async () => {
  const base = { op: 'pivot', sheet: 'Aging', range: 'A1:C4', name: 'P', rows: ['Entity'], values: [{ field: 'Total USD' }] };
  const bad = [
    { ...base, values: [] },
    { ...base, values: [{ field: 'Total USD', summarizeBy: 'Median' }] },
    { ...base, rows: [], columns: [] },
    { ...base, rows: 'Entity' },
    { ...base, values: [{ field: '' }] }
  ];
  for (const action of bad) {
    const reply = { model: 'm', text: 'Try.\n```json\n{"actions":[' + JSON.stringify(action) + '],"decision":"COMPLETE"}\n```' };
    const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
    await boot(sb);
    sb.els('workMode').value = 'agent';
    sb.els('directRevise').checked = true;
    await ask(sb, 'draft a pivot');
    assert.equal(sb.writes.some(w => w.kind === 'pivot'), false, JSON.stringify(action));
  }
});

async function runActions(actions, prompt) {
  const reply = { model: 'm', text: 'Applying.\n```json\n{"actions":' + JSON.stringify(actions) + ',"decision":"COMPLETE"}\n```' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, prompt);
  return sb;
}

test('Filter, conditional format, borders, freeze and dropdown all apply', async () => {
  const sb = await runActions([
    { op: 'filter', sheet: 'Sheet1', range: 'A1:C10', column: 'A', filterOn: 'Values', values: ['MUK'] },
    { op: 'conditionalFormat', sheet: 'Sheet1', range: 'C2:C10', operator: 'GreaterThan', formula1: '100', fill: '#FFC7CE' },
    { op: 'borders', sheet: 'Sheet1', range: 'A1:C10', edge: 'all', style: 'Continuous', weight: 'Thin', color: '#1F4E79' },
    { op: 'freeze', sheet: 'Sheet1', rows: 1, columns: 0 },
    { op: 'dataValidation', sheet: 'Sheet1', range: 'B2:B10', list: ['MUK', 'MCA'] }
  ], 'format and filter the table');
  const filter = sb.writes.find(w => w.kind === 'filter');
  assert.ok(filter, 'filter applied');
  assert.equal(filter.columnIndex, 0);
  assert.equal(filter.criteria.filterOn, 'Values');
  assert.equal(filter.criteria.values.join(','), 'MUK');
  const cf = sb.writes.find(w => w.kind === 'conditionalFormat');
  assert.ok(cf, 'conditional format applied');
  assert.equal(cf.type, 'CellValue');
  assert.equal(cf.rule.operator, 'GreaterThan');
  assert.equal(cf.rule.formula1, '100');
  assert.equal(cf.fill, '#FFC7CE');
  const borders = sb.writes.filter(w => w.kind === 'border');
  assert.equal(borders.length, 18);
  assert.ok(borders.some(b => b.edge === 'EdgeTop' && b.key === 'style' && b.value === 'Continuous'));
  assert.ok(borders.some(b => b.edge === 'InsideHorizontal' && b.key === 'weight' && b.value === 'Thin'));
  assert.ok(sb.writes.some(w => w.kind === 'freezeRows' && w.count === 1));
  const dv = sb.writes.find(w => w.kind === 'dataValidation');
  assert.ok(dv, 'dropdown applied');
  assert.equal(dv.rule.list.source, 'MUK,MCA');
  assert.equal(dv.rule.list.inCellDropDown, true);
});

test('Bulk changes wait for a click even with Direct on', async () => {
  const sb = await runActions([
    { op: 'replaceAll', sheet: 'Sheet1', text: 'OLD', replacement: 'NEW' },
    { op: 'filter', sheet: 'Sheet1', range: 'A1:C10', clear: true },
    { op: 'conditionalFormat', sheet: 'Sheet1', range: 'A1:C10', clear: true },
    { op: 'dataValidation', sheet: 'Sheet1', range: 'B2:B10', clear: true }
  ], 'clean everything up');
  for (const kind of ['replaceAll', 'clearFilters', 'clearConditionalFormats', 'clearDataValidation']) {
    assert.equal(sb.writes.some(w => w.kind === kind), false, kind + ' must wait for approval');
  }
  assert.match(visible(sb.els('conversation')), /not reversible by Undo/);
});

test('Malformed filter, highlight, border, freeze, replace and dropdown payloads are rejected', async () => {
  const bad = [
    { op: 'filter', sheet: 'Sheet1', range: 'A1:C10', column: 'Z', filterOn: 'Values', values: ['x'] },
    { op: 'filter', sheet: 'Sheet1', range: 'A1:C10', column: 'A', filterOn: 'Values', values: [] },
    { op: 'filter', sheet: 'Sheet1', range: 'A1:C10', column: 'A', filterOn: 'Nonsense' },
    { op: 'conditionalFormat', sheet: 'Sheet1', range: 'A1:C10', operator: 'Between', formula1: '1' },
    { op: 'conditionalFormat', sheet: 'Sheet1', range: 'A1:C10', operator: 'GreaterThan', formula1: '1' },
    { op: 'conditionalFormat', sheet: 'Sheet1', range: 'A1:C10', operator: 'Sideways', formula1: '1', fill: '#FFC7CE' },
    { op: 'borders', sheet: 'Sheet1', range: 'A1:C10', edge: 'sideways' },
    { op: 'borders', sheet: 'Sheet1', range: 'A1:C10', weight: 'enormous' },
    { op: 'freeze', sheet: 'Sheet1' },
    { op: 'freeze', sheet: 'Sheet1', rows: -1 },
    { op: 'replaceAll', sheet: 'Sheet1', text: '', replacement: 'x' },
    { op: 'dataValidation', sheet: 'Sheet1', range: 'A1:C10', list: ['a,b'] },
    { op: 'dataValidation', sheet: 'Sheet1', range: 'A1:C10', list: [] }
  ];
  const kinds = ['filter', 'conditionalFormat', 'border', 'freezeRows', 'freezeColumns', 'replaceAll', 'dataValidation'];
  for (const action of bad) {
    const sb = await runActions([action], 'try this');
    assert.equal(sb.writes.some(w => kinds.includes(w.kind)), false, JSON.stringify(action));
  }
});

test('Named range, hide, calculated table column and an extra chart type all apply', async () => {
  const sb = await runActions([
    { op: 'namedRange', sheet: 'Sheet1', range: 'C2:C10', name: 'AgingUSD' },
    { op: 'hide', sheet: 'Sheet1', target: 'rows', range: 'A5:A20', hidden: true },
    { op: 'hide', sheet: 'Sheet1', target: 'columns', range: 'D:E', hidden: true },
    { op: 'hide', sheet: 'Sheet1', target: 'sheet', hidden: true },
    { op: 'tableColumn', sheet: 'Sheet1', table: 'SalesTable', name: 'Margin', formula: '=[@Revenue]-[@Cost]' },
    { op: 'chart', sheet: 'Sheet1', range: 'A1:B10', type: 'doughnut', title: 'Mix' }
  ], 'tidy and extend the table');
  const named = sb.writes.find(w => w.kind === 'namedRange');
  assert.ok(named, 'named range created');
  assert.equal(named.name, 'AgingUSD');
  assert.equal(named.reference, 'Sheet1!C2:C10');
  assert.ok(sb.writes.some(w => w.kind === 'rowHidden' && w.value === true));
  assert.ok(sb.writes.some(w => w.kind === 'columnHidden' && w.value === true));
  assert.ok(sb.writes.some(w => w.kind === 'visibility' && w.value === 'Hidden'));
  const column = sb.writes.find(w => w.kind === 'tableColumn');
  assert.ok(column, 'calculated column added');
  assert.equal(column.table, 'SalesTable');
  assert.equal(column.name, 'Margin');
  assert.equal(column.values[0][0], '=[@Revenue]-[@Cost]');
  const chart = sb.writes.find(w => w.kind === 'chart');
  assert.ok(chart);
  assert.equal(chart.type, 'Doughnut');
});

test('Autofit captures widths and heights so Undo can restore them', async () => {
  const sb = await runActions([{ op: 'autofit', sheet: 'Sheet1', range: 'A1:C3' }], 'autofit please');
  assert.ok(sb.writes.some(w => w.kind === 'autofitColumns'));
  assert.ok(sb.writes.some(w => w.kind === 'autofitRows'));
  const cards = sb.els('conversation').children;
  const card = cards.find(c => String(c.className).includes('edits-card'));
  assert.ok(card, 'a revision card was rendered');
  const row = (card.children || []).find(c => c.className === 'button-row');
  const undo = row && (row.children || []).find(c => /Undo/.test(c.textContent));
  assert.ok(undo, 'an Undo button is offered');
  await fire(undo, 'click'); await flush(); await flush();
  assert.ok(sb.writes.some(w => w.kind === 'columnWidth' && w.value === 64), 'column width restored');
  assert.ok(sb.writes.some(w => w.kind === 'rowHeight' && w.value === 15), 'row height restored');
});

test('Malformed named range, hide and table column payloads are rejected', async () => {
  const bad = [
    { op: 'namedRange', sheet: 'Sheet1', range: 'A1:B2', name: 'A1' },
    { op: 'namedRange', sheet: 'Sheet1', range: 'A1:B2', name: 'has space' },
    { op: 'namedRange', sheet: 'Sheet1', range: 'nonsense', name: 'Good' },
    { op: 'hide', sheet: 'Sheet1', target: 'diagonal', range: 'A1:B2' },
    { op: 'hide', sheet: 'Sheet1', target: 'rows' },
    { op: 'tableColumn', sheet: 'Sheet1', table: 'T', name: 'C', formula: 'no equals sign' },
    { op: 'tableColumn', sheet: 'Sheet1', table: 'T', name: 'C' },
    { op: 'chart', sheet: 'Sheet1', range: 'A1:B2', type: 'treemap' }
  ];
  const kinds = ['namedRange', 'rowHidden', 'columnHidden', 'visibility', 'tableColumn', 'chart'];
  for (const action of bad) {
    const sb = await runActions([action], 'try this');
    assert.equal(sb.writes.some(w => kinds.includes(w.kind)), false, JSON.stringify(action));
  }
});

test('Word replace-all is click-gated and then replaces every match', async () => {
  const reply = { model: 'm', text: 'Replacing.\n```json\n{"actions":[{"op":"replaceAll","text":"team","replacement":"group"}],"decision":"COMPLETE"}\n```' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  sb.officeCb({ host: 'Word' });
  await flush(); await flush();
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'replace team with group everywhere');
  assert.equal(sb.writes.some(w => w.kind === 'wordSearch'), false, 'bulk replace waits for a click');
  assert.match(visible(sb.els('conversation')), /not reversible by Undo/);
  const risky = sb.els('conversation').children.flatMap(card => card.children || []).find(el => el.className === 'button-row');
  assert.ok(risky, 'an approval card was rendered');
});

test('An ageing pivot with 11 value fields is accepted, not silently dropped', async () => {
  const headers = ['Entity', 'Customer', 'Currency', 'Lines', 'FX to USD', 'Local O/S', 'USD O/S', 'Current', '1-30 Days', '31-60 Days', '61-90 Days', '91-180 Days', '181-365 Days', '1-1.5 Years', '1.5-2 Years', 'Over 2 Years', 'Credits (USD)'];
  const cells = {};
  headers.forEach((name, i) => { cells[`Aging!${String.fromCharCode(65 + i)}5`] = { value: name }; });
  const valueFields = ['USD O/S', 'Current', '1-30 Days', '31-60 Days', '61-90 Days', '91-180 Days', '181-365 Days', '1-1.5 Years', '1.5-2 Years', 'Over 2 Years', 'Credits (USD)'];
  const action = {
    op: 'pivot', sheet: 'Aging', range: 'A5:Q752', name: 'Pivot - Combined Aging',
    destinationSheet: 'Pivot - Combined Aging', destinationCell: 'A3',
    rows: ['Entity'], columns: [], filters: ['Currency'],
    values: valueFields.map(field => ({ field, summarizeBy: 'Sum' })), layout: 'Tabular', grandTotals: true
  };
  const reply = { model: 'm', text: 'Building.\n```json\n{"actions":[' + JSON.stringify(action) + '],"decision":"COMPLETE"}\n```' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells, missingSheets: ['Pivot - Combined Aging'] });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'draft a pivot for Combined Aging (USD)');
  const pivot = sb.writes.find(w => w.kind === 'pivot');
  assert.ok(pivot, 'the pivot was created rather than rejected');
  assert.equal(pivot.values.length, 11);
  assert.equal(pivot.source, 'Aging!A5:Q752');
  assert.equal(pivot.filters.join(','), 'Currency');
});

test('An unknown pivot field names the offending header and leaves no partial pivot', async () => {
  const cells = { 'Aging!A1': { value: 'Entity' }, 'Aging!B1': { value: 'Bucket' }, 'Aging!C1': { value: 'Total USD' } };
  const action = {
    op: 'pivot', sheet: 'Aging', range: 'A1:C4', name: 'P',
    destinationSheet: 'Pivot - X', destinationCell: 'A3',
    rows: ['Entity'], columns: [], filters: [],
    values: [{ field: 'USD O/S', summarizeBy: 'Sum' }], layout: 'Tabular'
  };
  const reply = {
    model: 'm',
    turns: [
      'Try.\n```json\n{"actions":[' + JSON.stringify(action) + '],"decision":"COMPLETE"}\n```',
      'Understood.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```'
    ]
  };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells, missingSheets: ['Pivot - X'] });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'build it');
  assert.ok(sb.writes.some(w => w.kind === 'deletePivot'), 'the partial pivot was removed');
  assert.ok(sb.writes.some(w => w.kind === 'deleteSheet'), 'the sheet it created was removed');
  const fed = lastContent(sb);
  assert.match(fed, /Pivot field\(s\) not in Aging!A1:C4: USD O\/S/);
  assert.match(fed, /Headers are: Entity, Bucket, Total USD/);
  assert.match(visible(sb.els('conversation')), /Execute FAILED/);
  assert.match(visible(sb.els('conversation')), /USD O\/S/);
});
