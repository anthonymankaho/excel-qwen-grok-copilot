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
    replaceChildren(...nodes) {
      for (const child of this.children) child.parent = null;
      this.children.length = 0;
      this.append(...nodes);
    },
    querySelector(sel) { return this.children.find(c => String(c.className).split(' ').includes(sel.slice(1))) || null; },
    remove() { if (this.parent) this.parent.children = this.parent.children.filter(c => c !== this); this.parent = null; },
    setAttribute(k, v) { (this.attributes ||= {})[k] = String(v); },
    getAttribute(k) { return this.attributes?.[k] ?? null; },
    focus() {}, click() { return fire(el, 'click'); }, requestSubmit() { return fire(el, 'submit'); }
  };
  // Real DOM children is an HTMLCollection without .at(); keep it undefined so regressions fail here.
  Object.defineProperty(el.children, 'at', { value: undefined });
  return el;
}
function fire(el, type, ev = {}) {
  return Promise.all((el.listeners[type] || []).map(fn => fn({ preventDefault() {}, ...ev })));
}
const flush = () => new Promise(resolve => setImmediate(resolve));

function makeSandbox({ selection, used, reply, seedKeys, cells, missingSheets, location, sheetNames, usedRanges, healthBuild, blockStorage, allowStorageAccess, officeStore, streamFailures = 0, prefs, extraStore, wordSelection, wordBody }) {
  const els = new Map();
  const defaults = { provider: 'qwen', qwenModel: 'qwen4-max', qwenEndpoint: 'https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions' };
  const get = id => { if (!els.has(id)) { const el = makeElement(id); if (defaults[id]) el.value = defaults[id]; if (id === 'includeContext' || id === 'rememberKey') el.checked = true; els.set(id, el); } return els.get(id); };
  const store = seedKeys ? { 'eqgc.keys.v1': JSON.stringify(seedKeys) } : {};
  if (extraStore) Object.assign(store, extraStore);
  if (prefs) store['eqgc.chat.v1'] = JSON.stringify(prefs);
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
      getUsedRangeOrNullObject: () => (usedRanges?.[name] ? { isNullObject: false, load() {}, ...usedRanges[name] } : { isNullObject: true, load() {} }),
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
  const wordDoc = { selectionText: wordSelection === undefined ? 'Hello team' : wordSelection, bodyText: wordBody === undefined ? 'Hello team. More text.' : wordBody, style: 'Normal', font: wordFont };
  let storageBlocked = !!blockStorage;
  const officeValues = officeStore || null;
  const sandbox = {
    console, JSON, Math, String, Array, Object, Promise, RegExp, Number, TextEncoder, TextDecoder, AbortController,
    btoa: s => Buffer.from(s, 'binary').toString('base64'),
    atob: s => Buffer.from(s, 'base64').toString('binary'),
    setTimeout: (fn, delay) => { if (delay === 700) queueMicrotask(fn); return 0; }, clearTimeout: () => {},
    localStorage: {
      getItem: k => (k in store ? store[k] : null),
      setItem: (k, v) => { if (storageBlocked) throw new Error('SecurityError: third-party storage blocked'); store[k] = String(v); },
      removeItem: k => { delete store[k]; }
    },
    document: { getElementById: get, createElement: tag => makeElement(tag), querySelectorAll: () => [], addEventListener() {}, requestStorageAccess: allowStorageAccess ? async () => { storageBlocked = false; } : undefined },
    Office: { HostType: { Excel: 'Excel', Word: 'Word' }, onReady(cb) { sandbox.officeCb = cb; } },
    OfficeRuntime: officeValues ? { storage: {
      getItem: async key => officeValues[key] ?? null,
      setItem: async (key, value) => { officeValues[key] = value; },
      removeItem: async key => { delete officeValues[key]; }
    } } : undefined,
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
            items: (sheetNames || ['Sheet1']).map(n => ({ name: n, load() {} })),
            getActiveWorksheet: () => ({ getUsedRangeOrNullObject: () => range(used) }),
            add(sheetName) {
              writes.push({ kind: 'addSheet', name: sheetName });
              const sheet = makeSheet(sheetName);
              // Real Office.js: a freshly added object throws PropertyNotLoaded when an
              // unloaded property is read. Reproduce that so mocks cannot hide the bug.
              let nullLoaded = false;
              sheet.load = props => { if (String(props).includes('isNullObject')) nullLoaded = true; };
              Object.defineProperty(sheet, 'isNullObject', {
                get() {
                  if (!nullLoaded) throw new Error("The property 'isNullObject' is not available. Before reading the property value, call the load method on the containing object and call context.sync() on the relevant request context.");
                  return false;
                }
              });
              return sheet;
            },
            getItem: sheet => makeSheet(sheet),
            getItemOrNullObject: sheet => makeSheet(sheet)
          }
        },
        sync: async () => {}
      })
    },
    Word: {
      run: fn => {
        const tracked = [];
        return fn({
        document: {
          getSelection: () => {
            const state = { pending: new Set(), loaded: new Set() };
            tracked.push(state);
            return {
            get text() { return wordDoc.selectionText; },
            load(props) {
              const list = Array.isArray(props) ? props : String(props || '').split(',');
              for (const prop of list) {
                const name = String(prop || '').trim();
                if (name) state.pending.add(name);
              }
            },
            insertText(text, where) { writes.push({ kind: 'wordInsert', where, text }); if (where === 'Replace') wordDoc.selectionText = text; },
            font: wordDoc.font,
            get style() {
              if (!state.loaded.has('style')) throw new Error("The property 'style' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.");
              return wordDoc.style;
            },
            set style(v) { wordDoc.style = v; writes.push({ kind: 'wordStyle', style: v }); }
            };
          },
          body: {
            get text() { return wordDoc.bodyText; },
            load() {},
            insertText(text, where) { writes.push({ kind: 'wordBody', where, text }); if (where === 'Replace') wordDoc.bodyText = text; },
            search(term, options) {
              writes.push({ kind: 'wordSearch', term, options });
              return { load() {}, items: [{ insertText(text, where) { writes.push({ kind: 'wordReplace', term, text, where }); } }] };
            }
          }
        },
        sync: async () => {
          for (const state of tracked) {
            for (const prop of state.pending) state.loaded.add(prop);
            state.pending.clear();
          }
        }
      });
      }
    },
    fetch: async (url, opts) => {
      const body = opts?.body ? JSON.parse(opts.body) : null;
      calls.push({ url, body, headers: opts?.headers || {} });
      if (url.endsWith('/api/health')) return { ok: true, status: 200, json: async () => ({ ok: true, build: healthBuild || '2026-11-26.3', configured: { qwen: true } }) };
      if (url.endsWith('/api/connectors/github/account')) return { ok: true, status: 200, json: async () => ({ connected: true, login: 'anthonymankaho', email: '' }) };
      if (url.endsWith('/api/connectors/github/connect')) return { ok: true, status: 200, json: async () => ({ connected: true, login: 'anthonymankaho', email: '' }) };
      if (url.endsWith('/api/connectors/github/repos')) return { ok: true, status: 200, json: async () => ({ items: [{ fullName: 'owner/repo', private: false, defaultBranch: 'main', description: '' }] }) };
      if (url.endsWith('/api/connectors/github/read')) return { ok: true, status: 200, json: async () => ({ repo: body.repo, path: body.path, text: 'hello from GitHub', truncated: false }) };
      if (url.endsWith('/api/connectors/github/write')) return { ok: true, status: 200, json: async () => ({ ok: true, target: body.target, repo: body.repo, base: 'main', branch: body.target === 'default' ? 'main' : 'eqgc/test', commit: 'abc', prUrl: body.target === 'default' ? '' : 'https://github.com/owner/repo/pull/7', files: body.files.map(file => file.path) }) };
      if (url.endsWith('/api/git/open')) return { ok: true, status: 200, json: async () => ({ ok: true, repo: body.repo || 'C:\\repo', path: 'src/app.js', text: 'function hello() {}\n' }) };
      if (url.endsWith('/api/git/pull')) return { ok: true, status: 200, json: async () => ({ ok: true, repo: body.repo || 'C:\\repo', path: body.path, text: 'pulled line\n', summary: 'Already up to date.' }) };
      if (url.endsWith('/api/git/commit')) return { ok: true, status: 200, json: async () => ({ ok: true, repo: body.repo, path: body.path, commit: 'abc1234', message: body.message }) };
      if (url.endsWith('/api/git/push')) return { ok: true, status: 200, json: async () => ({ ok: true, repo: body.repo, summary: 'main -> main' }) };
      if (url.endsWith('/api/connectors/search')) return { ok: true, status: 200, json: async () => ({ items: [{ id: 'item-1', name: 'aging.xlsx', detail: 'sheet' }] }) };
      if (url.endsWith('/api/connectors/read')) return { ok: true, status: 200, json: async () => ({ name: body.item.name, kind: 'XLSX', text: 'Coverage: COMPLETE\nGJ001 | 100' }) };
      if (url.endsWith('/api/extract')) return { ok: true, status: 200, json: async () => ({ text: '[Sheet: Data]\nL1: Voucher | Debit\nL2: GJ001 | 100', kind: 'XLSX' }) };
      if (url.endsWith('/api/sessions/delete')) return { ok: true, status: 200, json: async () => ({ ok: true }) };
      if (url.includes('/api/sessions/read')) return { ok: true, status: 200, json: async () => ({ id: 's-old-1', started: '2026-11-24T09:00:00Z', prompt: 'review whole file', messages: [{ role: 'user', content: 'review whole file' }, { role: 'assistant', content: 'Old report.' }], audit: [], checkpoint: { prompt: 'review whole file', progress: [], observation: '' } }) };
      if (url.endsWith('/api/sessions/save')) return { ok: true, status: 200, json: async () => ({ ok: true }) };
      if (url.endsWith('/api/sessions')) return { ok: true, status: 200, json: async () => ({ sessions: [{ id: 's-old-1', updated: '2026-11-24T09:00:00Z', workbook: 'MPL', prompt: 'review whole file', outcome: 'Completed', turns: 2 }] }) };
      if (url.endsWith('/api/memory/add')) return { ok: true, status: 200, json: async () => ({ ok: true }) };
      if (url.endsWith('/api/memory')) return { ok: true, status: 200, json: async () => ({ text: '- [2026-11-24] MPL truth is 30_GJ_2026' }) };
      if (url.includes('/api/skills/context')) return { ok: true, status: 200, json: async () => ({ selected: ['monthly-report-review'], text: 'SAVED SKILLS test block' }) };
      if (url.endsWith('/api/skills/github')) return { ok: true, status: 200, json: async () => ({ ok: true, id: 'demo-skill', path: 'C:\\skills\\demo-skill\\SKILL.md', replaced: false }) };
      if (url.endsWith('/api/skills/save')) return { ok: true, status: 200, json: async () => ({ ok: true, path: 'C:\\skills\\new-skill\\SKILL.md' }) };
      if (url.endsWith('/api/skills')) return { ok: true, status: 200, json: async () => ({ root: 'C:\\skills', skills: [{ id: 'monthly-report-review', description: 'Review the whole report', chars: 100 }] }) };
      const turnText = Array.isArray(reply?.turns) ? reply.turns[Math.min(streamTurn, reply.turns.length - 1)] : null;
      const data = turnText != null ? { text: turnText, model: reply.model || 'test-model' } : (reply || { text: 'mock answer', model: 'test-model' });
      if (url.endsWith('/api/chat-stream')) {
        streamTurn += 1;
        if (streamTurn <= streamFailures) return { ok: false, status: 500, json: async () => { throw new Error('non-JSON error'); } };
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
function savedAuditText(sb) {
  return sb.calls.filter(call => String(call.url).endsWith('/api/sessions/save')).map(call => JSON.stringify(call.body.audit || [])).join('\n');
}
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

test('composer Model and Effort are sent with the question', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb);
  sb.els('modelPick').value = 'qwen3.8-flash';
  await fire(sb.els('modelPick'), 'change');
  sb.els('effortMode').value = 'low';
  await ask(sb, 'draft the journal');
  const body = chatCalls(sb).at(-1).body;
  assert.equal(body.model, 'qwen3.8-flash');
  assert.equal(body.effort, 'low');
  assert.equal(body.fast, false);
  assert.equal(sb.els('qwenModel').value, 'qwen3.8-flash');
});

test('Agent sends the selected Effort', async () => {
  const reply = { text: 'Done.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```', model: 'qwen4-max' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('effortMode').value = 'xhigh';
  await ask(sb, 'review the sheet');
  assert.equal(chatCalls(sb).at(-1).body.effort, 'xhigh');
  assert.equal(chatCalls(sb).at(-1).body.mode, 'agent');
});

test('only auto, qwen4-max, qwen3.8-max, and qwen3.8-flash can be selected', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb);
  sb.els('modelPick').value = 'qwen3.7-plus';
  await fire(sb.els('modelPick'), 'change');
  assert.equal(sb.els('qwenModel').value, 'qwen4-max');
  sb.els('modelPick').value = 'auto';
  await fire(sb.els('modelPick'), 'change');
  sb.els('effortMode').value = 'medium';
  await fire(sb.els('effortMode'), 'change');
  await ask(sb, 'hi');
  assert.equal(chatCalls(sb).at(-1).body.model, 'auto');
  assert.equal(chatCalls(sb).at(-1).body.effort, 'medium');
  assert.equal(JSON.parse(sb.store['eqgc.chat.v1']).model, 'auto');
  assert.equal(JSON.parse(sb.store['eqgc.chat.v1']).effort, 'medium');
});

test('a remembered Model and Effort are restored into the composer', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, prefs: { model: 'qwen3.8-max', effort: 'xhigh' } });
  await boot(sb);
  assert.equal(sb.els('qwenModel').value, 'qwen3.8-max');
  assert.equal(sb.els('modelPick').value, 'qwen3.8-max');
  assert.equal(sb.els('effortMode').value, 'xhigh');
});

test('a remembered model outside the supported list falls back to qwen4-max', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, prefs: { model: 'qwen3.7-plus', effort: 'low' } });
  await boot(sb);
  assert.equal(sb.els('qwenModel').value, 'qwen4-max');
  assert.equal(sb.els('modelPick').value, 'qwen4-max');
  assert.equal(sb.els('effortMode').value, 'low');
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
  assert.match(lastContent(sb), /\[Word document\]/);
  assert.match(lastContent(sb), /More text/);
  assert.ok(sb.writes.some(w => w.kind === 'wordInsert' && w.where === 'Replace' && w.text === 'Revised paragraph.'));
  assert.ok(sb.writes.some(w => w.kind === 'wordFont' && w.key === 'bold' && w.value === true));
  assert.ok(sb.writes.some(w => w.kind === 'wordStyle' && w.style === 'Heading 1'));
  assert.ok(sb.writes.some(w => w.kind === 'wordBody' && w.where === 'End'));
});

test('a finished answer does not show an audit card', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Report.', model: 'm' } });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  await ask(sb, 'review the note');
  assert.equal(sb.els('conversation').children.some(c => String(c.className).includes('audit-card')), false);
  assert.equal(visible(sb.els('conversation')).includes('AUDIT'), false);
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

test('Agent pauses at the round limit and continues the task with that result', async () => {
  const read = 'Reading.\n```json\n{"contextRequest":{"sheet":"Sheet1","range":"B2"},"decision":"CONTINUE"}\n```';
  const reply = { turns: [] };
  const sb = makeSandbox({
    selection: singleEmpty, used: usedSmall,
    cells: { 'Sheet1!B2': { value: 73 } },
    reply
  });
  const limit = sb.AgentCore.MAX_LOOPS;
  reply.turns = [...Array(limit).fill(read), 'Finished reading.'];
  await boot(sb);
  sb.els('workMode').value = 'agent';
  await ask(sb, 'Inspect the selected worksheet');
  assert.equal(chatCalls(sb).length, limit);
  assert.match(sb.els('agentStatus').textContent, new RegExp(`Agent ${limit}\\/${limit}.*Paused`));
  assert.equal(sb.els('continueAgent').classList.contains('hidden'), false);
  await fire(sb.els('continueAgent'), 'click');
  assert.equal(chatCalls(sb).length, limit + 1);
  assert.match(chatCalls(sb).at(-1).body.messages.at(-2).content, /Inspect the selected worksheet/);
  assert.match(lastContent(sb), /Latest pane observation:[\s\S]+Sheet1!B2[\s\S]+Row 2:\s+73/);
  assert.match(sb.els('agentStatus').textContent, new RegExp(`Agent 1\\/${limit}.*Completed`));
  assert.equal(sb.els('continueAgent').classList.contains('hidden'), true);
});

test('Stop during check input read prevents the planned Agent edit', async () => {
  const reply = { text: 'Ready.\n```json\n{"plan":{"goal":"Fill the value","steps":[{"id":1,"objective":"Fill B2"}],"finalValidation":[{"check":"formulaErrors","sheet":"Sheet1","range":"B2"}]},"actions":[{"op":"set","sheet":"Sheet1","range":"B2","value":7}],"decision":"COMPLETE"}\n```' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  const officeRun = sb.Excel.run;
  let finishRead;
  sb.Excel.run = callback => /Reading check inputs/.test(sb.els('agentStatus').textContent)
    ? new Promise(resolve => { finishRead = () => resolve(officeRun(callback)); })
    : officeRun(callback);
  const request = ask(sb, 'Fill the value');
  await flush();
  assert.equal(typeof finishRead, 'function');
  await fire(sb.els('composer'), 'submit');
  assert.match(sb.els('agentStatus').textContent, /Stopping/);
  finishRead();
  await request;
  assert.equal(sb.writes.length, 0);
  assert.match(sb.els('agentStatus').textContent, /Stopped by you/);
  assert.equal(sb.els('continueAgent').classList.contains('hidden'), true);
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

test('Agent and chat route to Qwen only', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.', model: 'qwen4-max' } });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  await ask(sb, 'Reconcile the ledger');
  assert.equal(chatCalls(sb).at(-1).body.provider, 'qwen');
  assert.equal(chatCalls(sb).at(-1).body.mode, 'agent');
  assert.match(savedAuditText(sb), /"provider":"qwen"/);
});

test('A hosted pane sends the user own key per request (BYO) and never configures the server', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.', model: 'qwen4-max' }, location: { hostname: 'excel-qwen-grok.example.workers.dev' } });
  await boot(sb);
  sb.els('apiKey').value = 'sk-sp-byo-user-key-001';
  sb.els('rememberKey').checked = true;
  await fire(sb.els('saveKey'), 'click');
  await flush();
  assert.match(sb.store['eqgc.keys.v1'], /sk-sp-byo-user-key-001/);
  assert.equal(sb.calls.some(call => String(call.url).includes('/api/configure')), false);
  await ask(sb, 'hello from cloud');
  const wire = chatCalls(sb).at(-1);
  assert.equal(wire.body.apiKey, 'sk-sp-byo-user-key-001');
  assert.equal(wire.headers.Authorization, undefined);
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
  assert.match(savedAuditText(sb), /Aging!A1:Q2/);
});

test('An over-limit read is paged with offset until fully read, never reported as not read', async () => {
  const reply = {
    model: 'm',
    turns: [
      'Reading.\n```json\n{"actions":[],"contextRequest":{"sheet":"Aging","range":"A1:Z200"},"decision":"CONTINUE"}\n```',
      'Next chunk.\n```json\n{"actions":[],"contextRequest":{"sheet":"Aging","range":"A1:Z200","offset":76},"decision":"CONTINUE"}\n```',
      'Last chunk.\n```json\n{"actions":[],"contextRequest":{"sheet":"Aging","range":"A1:Z200","offset":152},"decision":"CONTINUE"}\n```',
      'Done.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```'
    ]
  };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells: { 'Aging!A1': { value: 'Entity' } } });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'review the whole ageing sheet');
  assert.equal(chatCalls(sb).length, 4);
  const first = chatCalls(sb)[1].body.messages.at(-1).content;
  assert.match(first, /\[Excel range Aging!A1:Z76\]/);
  assert.match(first, /76\/200 rows/);
  assert.match(first, /offset=76/);
  const second = chatCalls(sb)[2].body.messages.at(-1).content;
  assert.match(second, /\[Excel range Aging!A77:Z152\]/);
  assert.match(second, /152\/200 rows/);
  assert.match(second, /offset=152/);
  const fed = lastContent(sb);
  assert.match(fed, /\[Excel range Aging!A153:Z200\]/);
  assert.match(fed, /200\/200 rows/);
  assert.match(fed, /fully read/);
  assert.equal(/No range was read/.test(fed), false);
  assert.equal(sb.writes.length, 0);
});

test('A finance check aggregates the full requested range, not a 160-cell sample', async () => {
  const cells = { 'Sheet1!A1': { value: 'Debit' }, 'Sheet1!B1': { value: 'Credit' } };
  for (let r = 2; r <= 200; r++) {
    cells[`Sheet1!A${r}`] = { value: 1 };
    cells[`Sheet1!B${r}`] = { value: r === 200 ? 0 : 1 };
  }
  const reply = { text: 'Check.\n```json\n{"plan":{"goal":"Verify the balance","finalValidation":[{"check":"debitEqualsCredit","sheet":"Sheet1","range":"A1:B200"}]},"actions":[],"decision":"COMPLETE"}\n```' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  await ask(sb, 'verify the journal balances');
  const shown = visible(sb.els('conversation'));
  assert.match(shown, /FAIL debitEqualsCredit/);
  assert.match(shown, /before 199 after 198 difference 1/);
});

test('A check range that cannot be fully read is NEEDS_USER, never PASS', async () => {
  const reply = { text: 'Check.\n```json\n{"plan":{"goal":"Scan errors","finalValidation":[{"check":"formulaErrors","sheet":"Ledger","range":"A1:Z2000"}]},"actions":[],"decision":"COMPLETE"}\n```' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells: { 'Ledger!A1': { value: 'Date' } } });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  const run = sb.Excel.run;
  let runs = 0;
  sb.Excel.run = callback => {
    if (/Reading check inputs|Checking results/.test(sb.els('agentStatus').textContent)) {
      runs += 1;
      if (runs % 2 === 0) return Promise.reject(new Error('The requested resource does not exist.'));
    }
    return run(callback);
  };
  await ask(sb, 'scan the ledger for errors');
  const shown = visible(sb.els('conversation'));
  assert.match(shown, /NEEDS_USER formulaErrors/);
  assert.match(shown, /could not be fully read/);
  assert.equal(/PASS formulaErrors/.test(shown), false);
});

test('Chat reviews the whole file read-only: scan, group balance and cross-check, never writes', async () => {
  const cells = {
    '30_GJ_2026!A1': { value: 'Voucher' }, '30_GJ_2026!B1': { value: 'Debit' }, '30_GJ_2026!C1': { value: 'Credit' },
    '30_GJ_2026!A2': { value: 'GJ001' }, '30_GJ_2026!B2': { value: 100 }, '30_GJ_2026!C2': { value: 100 },
    '30_GJ_2026!A3': { value: 'GJ002' }, '30_GJ_2026!B3': { value: 50 }, '30_GJ_2026!C3': { value: 50 },
    '40_CSV!A1': { value: 'Voucher' }, '40_CSV!B1': { value: 'Debit' }, '40_CSV!C1': { value: 'Credit' },
    '40_CSV!A2': { value: 'GJ001' }, '40_CSV!B2': { value: 100 }, '40_CSV!C2': { value: 100 },
    '40_CSV!A3': { value: 'GJ002' }, '40_CSV!B3': { value: 50 }, '40_CSV!C3': { value: 50 },
    '40_CSV!A4': { value: 'GJ002' }, '40_CSV!B4': { value: 50 }, '40_CSV!C4': { value: 50 }
  };
  const reply = { turns: [
    'Scanning.\n```json\n{"scanRequest":{},"decision":"CONTINUE"}\n```',
    'Reviewed.\n```json\n{"plan":{"goal":"Whole-file review","finalValidation":[{"check":"groupBalance","sheet":"30_GJ_2026","range":"A1:C3"},{"check":"crossCheck","sheet":"40_CSV","range":"A1:C4","otherSheet":"30_GJ_2026","otherRange":"A1:C3","columns":[2]}]},"actions":[],"decision":"COMPLETE"}\n```'
  ] };
  const sb = makeSandbox({
    selection: singleEmpty, used: usedSmall, reply, cells,
    sheetNames: ['30_GJ_2026', '40_CSV'],
    usedRanges: {
      '30_GJ_2026': { address: "'30_GJ_2026'!A1:C3", rowCount: 3, columnCount: 3 },
      '40_CSV': { address: "'40_CSV'!A1:C4", rowCount: 4, columnCount: 3 }
    }
  });
  await boot(sb);
  sb.els('workMode').value = 'chat';
  await ask(sb, 'Please review the whole workbook journal.');
  assert.equal(sb.writes.length, 0);
  assert.equal(chatCalls(sb).length, 2);
  const scanFed = chatCalls(sb)[1].body.messages.at(-1).content;
  assert.match(scanFed, /Workbook scan: 2 sheets, 0 error values/);
  assert.match(scanFed, /sheet: 30_GJ_2026 .*scanned 3\/3/);
  assert.match(scanFed, /sheet: 40_CSV .*scanned 4\/4/);
  const shown = visible(sb.els('conversation'));
  assert.match(shown, /PASS groupBalance/);
  assert.match(shown, /FAIL crossCheck/);
  assert.match(shown, /rows 3 vs 2/);
  assert.match(shown, /total 200 vs 150/);
});

test('Saved skills context is attached to the question and audited', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.', model: 'm' } });
  await boot(sb);
  sb.els('workMode').value = 'chat';
  await ask(sb, 'review the whole report');
  assert.match(lastContent(sb), /\[Saved skills\]\nSAVED SKILLS test block/);
  assert.match(visible(sb.els('conversation')), /Skills: monthly-report-review/);
  assert.match(savedAuditText(sb), /skills monthly-report-review/);
});

test('GitHub skill install posts the repo and token without putting the token in the question', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb);
  sb.els('githubRepo').value = 'https://github.com/kgraph57/demo-skill';
  sb.els('githubToken').value = 'ghp_pane_token';
  sb.els('githubRemember').checked = true;
  sb.els('githubReplace').checked = false;
  await fire(sb.els('installGithubSkill'), 'click'); await flush();
  const call = sb.calls.find(c => String(c.url).endsWith('/api/skills/github'));
  assert.ok(call);
  assert.equal(call.body.repo, 'https://github.com/kgraph57/demo-skill');
  assert.equal(call.body.token, 'ghp_pane_token');
  assert.equal(call.body.replace, false);
  assert.equal(JSON.parse(sb.store['eqgc.github.v1']).token, 'ghp_pane_token');
  assert.equal(String(sb.els('prompt').value).includes('ghp_pane_token'), false);
});

test('Save current task as skill posts a new folder with trigger and procedure', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Procedure.', model: 'm' } });
  await boot(sb);
  await ask(sb, 'review the whole report');
  sb.els('skillId').value = 'oct-close-review';
  sb.els('skillDescription').value = 'October close review';
  await fire(sb.els('saveSkill'), 'click'); await flush();
  const save = sb.calls.find(c => String(c.url).endsWith('/api/skills/save'));
  assert.ok(save);
  assert.equal(save.body.id, 'oct-close-review');
  assert.match(save.body.body, /## Trigger[\s\S]+review the whole report/);
  assert.match(save.body.body, /Procedure\./);
});

test('Memory notes are attached to questions while Share is on', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.', model: 'm' } });
  await boot(sb);
  sb.els('workMode').value = 'chat';
  await ask(sb, 'review the whole report');
  assert.match(lastContent(sb), /\[Memory\]\n- \[2026-11-24\] MPL truth is 30_GJ_2026/);
});

test('A finished run saves its session for later resume', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Report.', model: 'm' } });
  await boot(sb);
  sb.els('workMode').value = 'chat';
  await ask(sb, 'review the whole report');
  await flush();
  const save = sb.calls.find(c => String(c.url).endsWith('/api/sessions/save'));
  assert.ok(save);
  assert.match(save.body.id, /^s-[a-z0-9-]+$/);
  assert.equal(save.body.prompt, 'review the whole report');
  assert.ok(save.body.messages.some(m => m.role === 'assistant' && m.content === 'Report.'));
});

test('History resume restores the conversation and the Agent checkpoint', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Continued.', model: 'm' } });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  await flush();
  const row = sb.els('historyList').children[0];
  const resume = row.children.find(c => c.textContent === 'Resume');
  await fire(resume, 'click'); await flush();
  assert.match(visible(sb.els('conversation')), /Old report\./);
  assert.equal(sb.els('continueAgent').classList.contains('hidden'), false);
  await fire(sb.els('continueAgent'), 'click'); await flush();
  assert.match(lastContent(sb), /Continue the original task/);
});

test('Workbook access opens from the bottom button and stays out of the sidebar', async () => {
  const html = fs.readFileSync(path.join(root, 'public', 'index.html'), 'utf8');
  const aside = html.slice(html.indexOf('<aside id="sidebar"'), html.indexOf('</aside>'));
  assert.equal(aside.includes('id="workbookCard"'), false);
  assert.equal(aside.includes('id="contextEyebrow"'), false);
  assert.match(html, /id="workbookCard" class="context-card card hidden"/);
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(sb);
  sb.els('workbookCard').classList.add('hidden');
  sb.els('sidebar').classList.add('hidden');
  await fire(sb.els('workbookToggle'), 'click');
  assert.equal(sb.els('workbookCard').classList.contains('hidden'), false);
  assert.equal(sb.els('sidebar').classList.contains('hidden'), true);
  assert.equal(sb.els('workbookToggle').getAttribute('aria-expanded'), 'true');
  await fire(sb.els('workbookToggle'), 'click');
  assert.equal(sb.els('workbookCard').classList.contains('hidden'), true);
});

test('Copilot-style sidebar hides the five browse cards and Share quick toggle mirrors the consent gate', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Ok.', model: 'm' } });
  await boot(sb);
  // The sandbox does not parse index.html classes, so drive the toggle from JS both ways.
  const initial = sb.els('sidebar').classList.contains('hidden');
  await fire(sb.els('sidebarToggle'), 'click');
  assert.equal(sb.els('sidebar').classList.contains('hidden'), !initial);
  if (sb.els('sidebar').classList.contains('hidden')) await fire(sb.els('sidebarToggle'), 'click');
  assert.equal(sb.els('sidebar').classList.contains('hidden'), false);
  const html = fs.readFileSync(path.join(root, 'public', 'index.html'), 'utf8');
  const aside = html.slice(html.indexOf('<aside id="sidebar"'), html.indexOf('</aside>'));
  for (const id of ['cloudSources', 'skillsCard', 'historyCard', 'memoryCard']) {
    assert.ok(aside.includes(`id="${id}"`), `${id} lives in the sidebar`);
  }
  assert.equal(/<details id="guideCard"[^>]*\sopen[\s>]/.test(aside), false);
  assert.match(aside, /<summary><strong>MCP<\/strong>/);
  assert.match(aside, /<summary><strong>Skills<\/strong>/);
  assert.equal(aside.includes('<strong>Cloud sources</strong>'), false);
  assert.equal(aside.includes('<strong>Saved skills</strong>'), false);
  assert.equal(html.indexOf('id="cloudSources"'), html.lastIndexOf('id="cloudSources"'), 'cards are not duplicated in main');
  sb.els('shareQuick').checked = false;
  await fire(sb.els('shareQuick'), 'change');
  assert.equal(sb.els('includeContext').checked, false);
  sb.els('includeContext').checked = true;
  await fire(sb.els('includeContext'), 'change');
  assert.equal(sb.els('shareQuick').checked, true);
  await fire(sb.els('sidebarClose'), 'click');
  assert.equal(sb.els('sidebar').classList.contains('hidden'), true);
  await fire(sb.els('sidebarToggle'), 'click');
  assert.equal(sb.els('sidebar').classList.contains('hidden'), false);
});

function findButton(el, label) {
  if (el?.textContent === label && el.listeners?.click) return el;
  for (const child of el?.children || []) {
    const found = findButton(child, label);
    if (found) return found;
  }
  return null;
}

test('git commit in Word sends the document text and does not call the model', async () => {
  const sb = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    extraStore: { 'eqgc.git.v1': JSON.stringify({ repo: 'C:\\repo', file: 'notes.txt' }) }
  });
  await boot(sb);
  sb.officeCb({ host: 'Word' });
  await ask(sb, 'git commit Update the note');
  assert.equal(sb.calls.some(call => String(call.url).endsWith('/api/chat-stream')), false);
  const call = sb.calls.find(item => String(item.url).endsWith('/api/git/commit'));
  assert.equal(call.body.path, 'notes.txt');
  assert.equal(call.body.content, 'Hello team. More text.');
  assert.equal(call.body.message, 'Update the note');
  assert.match(visible(sb.els('conversation')), /abc1234/);
});

test('git push in Word does not call the model or use --force', async () => {
  const sb = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    extraStore: { 'eqgc.git.v1': JSON.stringify({ repo: 'C:\\repo', file: 'notes.txt' }) }
  });
  await boot(sb);
  sb.officeCb({ host: 'Word' });
  await ask(sb, 'git push');
  assert.equal(sb.calls.some(call => String(call.url).endsWith('/api/chat-stream')), false);
  const call = sb.calls.find(item => String(item.url).endsWith('/api/git/push'));
  assert.equal(call.body.repo, 'C:\\repo');
  assert.equal(JSON.stringify(call.body).includes('force'), false);
  assert.match(visible(sb.els('conversation')), /main -> main/);
  await ask(sb, 'git push --force');
  assert.equal(sb.calls.filter(item => String(item.url).endsWith('/api/git/push')).length, 1);
  assert.match(visible(sb.els('conversation')), /--force/);
  const excel = makeSandbox({ selection: singleEmpty, used: usedSmall, extraStore: { 'eqgc.git.v1': JSON.stringify({ repo: 'C:\\repo', file: 'notes.txt' }) } });
  await boot(excel);
  await ask(excel, 'git push');
  assert.equal(excel.calls.some(call => String(call.url).includes('/api/git/')), false);
});

test('git commit and push writes the document and then pushes', async () => {
  const sb = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    extraStore: { 'eqgc.git.v1': JSON.stringify({ repo: 'C:\\repo', file: 'notes.txt' }) }
  });
  await boot(sb);
  sb.officeCb({ host: 'Word' });
  await ask(sb, 'git commit and push Update the note');
  assert.equal(sb.calls.some(call => String(call.url).endsWith('/api/chat-stream')), false);
  const commit = sb.calls.find(item => String(item.url).endsWith('/api/git/commit'));
  const push = sb.calls.find(item => String(item.url).endsWith('/api/git/push'));
  assert.equal(commit.body.content, 'Hello team. More text.');
  assert.equal(commit.body.message, 'Update the note');
  assert.ok(sb.calls.indexOf(commit) < sb.calls.indexOf(push));
  assert.equal(sb.writes.some(item => item.kind === 'wordBody' && item.where === 'End'), false);
});

test('a finished Word answer is appended and an Excel answer is not', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Review the function.', model: 'm' } });
  await boot(sb);
  sb.officeCb({ host: 'Word' });
  await ask(sb, 'what does this file do');
  const appended = sb.writes.filter(item => item.kind === 'wordBody' && item.where === 'End');
  assert.equal(appended.length, 1);
  assert.match(appended[0].text, /Qwen/);
  assert.match(appended[0].text, /Review the function/);
  const excel = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Review the function.', model: 'm' } });
  await boot(excel);
  await ask(excel, 'what does this file do');
  assert.equal(excel.writes.some(item => item.kind === 'wordBody'), false);
});

test('Agent in Word reads a named document passage and does not ask for a selection', async () => {
  const body = `${'Overview line.\n'.repeat(2000)}Revision Analysis (2026-11-26)\nOpus 5.5 advice: keep the local relay.\n`;
  const sb = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    reply: { text: 'Read.', model: 'm' },
    wordSelection: '',
    wordBody: body
  });
  await boot(sb);
  sb.officeCb({ host: 'Word' });
  sb.els('workMode').value = 'agent';
  await ask(sb, 'please read from "Revision Analysis (2026-11-26)"');
  const sent = chatCalls(sb).map(call => JSON.stringify(call.body)).join('\n');
  assert.match(sent, /Opus 5\.5 advice: keep the local relay/);
  assert.equal(sent.includes('Ask for the selection'), false);
  assert.match(sent, /\[Word document\]/);
});

test('a selected Word heading still includes the text under it', async () => {
  const body = 'Overview\nRevision Analysis (2026-11-26)\nOpus 5.5 advice: keep the local relay.\nSecond advice line.\n';
  const sb = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    reply: { text: 'Read.', model: 'm' },
    wordSelection: 'Revision Analysis (2026-11-26)',
    wordBody: body
  });
  await boot(sb);
  sb.officeCb({ host: 'Word' });
  sb.els('workMode').value = 'agent';
  await ask(sb, 'please read from "Revision Analysis (2026-11-26)"');
  const sent = chatCalls(sb).map(call => JSON.stringify(call.body)).join('\n');
  assert.match(sent, /Opus 5\.5 advice: keep the local relay/);
  assert.match(sent, /Second advice line/);
  assert.match(sent, /\[Word document\]/);
  assert.equal(sent.includes('[Word selection]'), false);
});

test('open in Word replaces the document and git pull in Excel does not call Git', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, extraStore: { 'eqgc.git.v1': JSON.stringify({ repo: 'C:\\repo', file: '' }) } });
  await boot(sb);
  sb.officeCb({ host: 'Word' });
  await ask(sb, 'open src/app.js');
  const opened = sb.writes.find(item => item.kind === 'wordBody' && item.where === 'Replace');
  assert.equal(opened.text, 'function hello() {}\n');
  assert.match(sb.store['eqgc.git.v1'], /src\/app\.js/);
  await ask(sb, 'git pull');
  assert.equal(sb.calls.some(call => String(call.url).endsWith('/api/chat-stream')), false);
  const excel = makeSandbox({ selection: singleEmpty, used: usedSmall });
  await boot(excel);
  await ask(excel, 'git pull');
  assert.equal(excel.calls.some(call => String(call.url).includes('/api/git/')), false);
  assert.match(visible(excel.els('conversation')), /Word/);
});

test('a GitHub write stays on the card until Write to GitHub is clicked', async () => {
  const token = 'ghp_testtoken';
  const sb = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    reply: { text: 'Draft ready.\n```json\n{"githubWrite":{"repo":"owner/repo","title":"Update notes","body":"Explain the change","files":[{"path":"README.md","content":"Hello from the pane"}]}}\n```', model: 'm' }
  });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('githubMcpToken').value = token;
  await ask(sb, 'update the readme');
  assert.equal(sb.calls.some(call => String(call.url).endsWith('/api/connectors/github/write')), false);
  assert.match(visible(sb.els('conversation')), /Hello from the pane/);
  assert.equal(visible(sb.els('conversation')).includes(token), false);
  const button = findButton(sb.els('conversation'), 'Write to GitHub');
  assert.ok(button);
  await fire(button, 'click');
  await flush();
  const call = sb.calls.find(item => String(item.url).endsWith('/api/connectors/github/write'));
  assert.equal(call.body.target, 'pull-request');
  assert.equal(call.body.files[0].content, 'Hello from the pane');
  assert.equal(call.body.token, token);
  assert.equal(call.body.useAccount, undefined);
  assert.match(visible(sb.els('conversation')), /pull\/7/);
  const sentToModel = sb.calls.filter(item => String(item.url).endsWith('/api/chat-stream')).map(item => JSON.stringify(item.body)).join('\n');
  assert.equal(sentToModel.includes(token), false);
});

test('Github Account uses the signed-in login and a write does not put a token in the chat', async () => {
  const sb = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    reply: { text: 'Draft ready.\n```json\n{"githubWrite":{"repo":"owner/repo","title":"Update notes","body":"Explain the change","files":[{"path":"README.md","content":"Hello from the pane"}]}}\n```', model: 'm' }
  });
  await boot(sb);
  assert.match(sb.els('githubMcpStatus').textContent, /Connected as anthonymankaho/);
  assert.equal(sb.els('connectGitHub').classList.contains('hidden'), true);
  assert.equal(sb.els('disconnectGitHub').classList.contains('hidden'), false);
  sb.els('workMode').value = 'agent';
  await ask(sb, 'update the readme');
  const button = findButton(sb.els('conversation'), 'Write to GitHub');
  await fire(button, 'click');
  await flush();
  const call = sb.calls.find(item => String(item.url).endsWith('/api/connectors/github/write'));
  assert.equal(call.body.useAccount, true);
  assert.equal(call.body.token, undefined);
  const sentToModel = sb.calls.filter(item => String(item.url).endsWith('/api/chat-stream')).map(item => JSON.stringify(item.body)).join('\n');
  assert.equal(sentToModel.includes('useAccount'), false);
  assert.equal(sentToModel.includes('gho_'), false);
  assert.equal(sentToModel.includes('anthonymankaho'), false);
  await fire(sb.els('disconnectGitHub'), 'click');
  assert.equal(sb.store['eqgc.github.account'], 'off');
  assert.match(sb.els('githubMcpStatus').textContent, /Not connected/);
  const optedOut = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    extraStore: { 'eqgc.github.account': 'off' },
    reply: { turns: [
      '```json\n{"githubRequest":{"action":"repos","query":"demo"},"decision":"CONTINUE"}\n```',
      'Done.\n```json\n{"decision":"COMPLETE"}\n```'
    ] }
  });
  await boot(optedOut);
  assert.equal(optedOut.calls.some(item => String(item.url).endsWith('/api/connectors/github/account')), false);
  assert.match(optedOut.els('githubMcpStatus').textContent, /Not connected/);
  optedOut.els('workMode').value = 'agent';
  await ask(optedOut, 'list my repos');
  const repos = optedOut.calls.find(item => String(item.url).endsWith('/api/connectors/github/repos'));
  assert.equal(repos.body.useAccount, undefined);
  assert.equal(repos.body.token, undefined);
  await fire(optedOut.els('connectGitHub'), 'click');
  await flush();
  assert.equal(optedOut.store['eqgc.github.account'], undefined);
  assert.match(optedOut.els('githubMcpStatus').textContent, /Connected as anthonymankaho/);
  assert.ok(optedOut.calls.some(item => String(item.url).endsWith('/api/connectors/github/connect')));
});

test('the chat box can select a GitHub repository and sends that name, not a token', async () => {
  const sb = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    reply: { text: 'Noted.\n```json\n{"decision":"COMPLETE"}\n```', model: 'm' }
  });
  await boot(sb);
  assert.equal(sb.els('githubRepoPick').textContent, 'No Repository');
  assert.equal(sb.els('githubRepoMenu').classList.contains('hidden'), true);
  sb.els('workMode').value = 'agent';
  await fire(sb.els('githubRepoPick'), 'click');
  await flush();
  assert.equal(sb.els('githubRepoMenu').classList.contains('hidden'), false);
  const listed = sb.calls.find(item => String(item.url).endsWith('/api/connectors/github/repos'));
  assert.equal(listed.body.useAccount, true);
  assert.equal(listed.body.token, undefined);
  const option = sb.els('githubRepoList').children[0];
  await fire(option, 'click');
  assert.equal(sb.els('githubRepoPick').textContent, 'repo');
  assert.match(sb.store['eqgc.github.repo'], /owner\/repo/);
  await ask(sb, 'read the readme');
  assert.match(lastContent(sb), /\[GitHub repository\]\nowner\/repo/);
  const sent = JSON.stringify(chatCalls(sb).at(-1).body);
  assert.equal(sent.includes('gho_'), false);
  assert.equal(sent.includes('ghp_'), false);
  await fire(sb.els('githubRepoPick'), 'click');
  await fire(sb.els('githubRepoNone'), 'click');
  assert.equal(sb.els('githubRepoPick').textContent, 'No Repository');
  assert.equal(sb.store['eqgc.github.repo'], undefined);
});

test('DOCX attachments are extracted by the local relay', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Ok.', model: 'm' } });
  await boot(sb);
  const picker = sb.els('attachmentPicker');
  picker.files = [{ name: 'memo.docx', size: 1000, arrayBuffer: async () => new Uint8Array([80, 75]).buffer }];
  await fire(picker, 'change', { target: picker }); await flush();
  const extract = sb.calls.find(c => String(c.url).endsWith('/api/extract'));
  assert.ok(extract);
  assert.equal(extract.body.name, 'memo.docx');
});

test('Agent reads one OneDrive file when the model asks and Share is on', async () => {
  const sb = makeSandbox({
    selection: singleEmpty,
    used: usedSmall,
    reply: { turns: [
      '```json\n{"sourceRequest":{"action":"read","id":"item-1","name":"aging.xlsx"},"decision":"CONTINUE"}\n```',
      'The file is in.\n```json\n{"decision":"COMPLETE"}\n```'
    ] }
  });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  await ask(sb, 'Read the October aging from OneDrive');
  const read = sb.calls.find(c => String(c.url).endsWith('/api/connectors/read'));
  assert.ok(read);
  assert.equal(read.body.item.id, 'item-1');
  assert.equal(read.body.item.name, 'aging.xlsx');
  assert.match(lastContent(sb), /GJ001 \| 100/);
});

test('XLSX attachments are extracted by the local relay and attached as evidence', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Ok.', model: 'm' } });
  await boot(sb);
  const picker = sb.els('attachmentPicker');
  picker.files = [{ name: 'book.xlsx', size: 1000, arrayBuffer: async () => new Uint8Array([80, 75]).buffer }];
  await fire(picker, 'change', { target: picker }); await flush();
  const extract = sb.calls.find(c => String(c.url).endsWith('/api/extract'));
  assert.ok(extract);
  assert.equal(extract.body.name, 'book.xlsx');
  assert.match(sb.els('attachmentTextPreview').textContent, /Voucher \| Debit/);
});

test('History delete arms once, then deletes the session on confirm', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Ok.', model: 'm' } });
  await boot(sb); await flush();
  const row = sb.els('historyList').children[0];
  const del = row.children.find(c => c.textContent === 'Delete');
  await fire(del, 'click');
  assert.equal(del.textContent, 'Confirm delete?');
  assert.equal(sb.calls.some(c => String(c.url).endsWith('/api/sessions/delete')), false);
  await fire(del, 'click'); await flush();
  const call = sb.calls.find(c => String(c.url).endsWith('/api/sessions/delete'));
  assert.ok(call);
  assert.equal(call.body.id, 's-old-1');
});

test('Hosted pane hides local-data cards but keeps the guide', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Ok.', model: 'm' }, location: { hostname: 'excel-qwen-grok.example.workers.dev' } });
  await boot(sb); await flush();
  for (const id of ['skillsCard', 'historyCard', 'memoryCard', 'cloudSources', 'githubRepoSlot']) {
    assert.equal(sb.els(id).classList.contains('hidden'), true, `${id} hidden on hosted`);
  }
  assert.equal(sb.els('guideCard').classList.contains('hidden'), false);
});

test('Relay-extracted attachments allow 20 MB while text files stay at 5 MB', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Ok.', model: 'm' } });
  await boot(sb);
  const picker = sb.els('attachmentPicker');
  picker.files = [{ name: 'big.xlsx', size: 6 * 1024 * 1024, arrayBuffer: async () => new Uint8Array([80, 75]).buffer }];
  await fire(picker, 'change', { target: picker }); await flush();
  assert.equal(sb.calls.filter(c => String(c.url).endsWith('/api/extract')).length, 1);
  picker.files = [{ name: 'huge.xlsx', size: 21 * 1024 * 1024, arrayBuffer: async () => new Uint8Array([80, 75]).buffer }];
  await fire(picker, 'change', { target: picker }); await flush();
  assert.equal(sb.calls.filter(c => String(c.url).endsWith('/api/extract')).length, 1);
  const inline = sb.els('attachmentList').children.map(c => c.textContent).join(' ');
  assert.match(inline, /over the 20 MB limit/);
  assert.match(inline, /→ .*split\/compress/);
});

test('Copilot-style attach menu opens and attached files render as removable chips', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Ok.', model: 'm' } });
  await boot(sb);
  const initial = sb.els('attachMenu').classList.contains('hidden');
  await fire(sb.els('attachFiles'), 'click');
  assert.equal(sb.els('attachMenu').classList.contains('hidden'), !initial);
  if (sb.els('attachMenu').classList.contains('hidden')) await fire(sb.els('attachFiles'), 'click');
  assert.equal(sb.els('attachMenu').classList.contains('hidden'), false);
  await fire(sb.els('menuUpload'), 'click');
  assert.equal(sb.els('attachMenu').classList.contains('hidden'), true);
  const picker = sb.els('attachmentPicker');
  picker.files = [{ name: 'book.xlsx', size: 1000, arrayBuffer: async () => new Uint8Array([80, 75]).buffer }];
  await fire(picker, 'change', { target: picker }); await flush();
  const chips = sb.els('attachmentChips').children;
  assert.equal(chips.length, 1);
  assert.match(chips[0].children.map(c => c.textContent).join(''), /book\.xlsx/);
  const remove = chips[0].children.find(c => c.textContent === '×');
  await fire(remove, 'click'); await flush();
  assert.equal(sb.els('attachmentChips').children.length, 0);
});

test('When the browser blocks add-in storage, the hosted key stays session-only and says so', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.', model: 'm' }, location: { hostname: 'excel-qwen-grok.example.workers.dev' }, blockStorage: true });
  await boot(sb);
  sb.els('apiKey').value = 'sk-sp-byo-user-key-001';
  sb.els('rememberKey').checked = true;
  await fire(sb.els('saveKey'), 'click'); await flush();
  assert.match(sb.els('toast').textContent, /Office and browser storage are unavailable/);
  await ask(sb, 'hello from cloud');
  assert.equal(chatCalls(sb).at(-1).body.apiKey, 'sk-sp-byo-user-key-001');
});

test('Hosted routing uses the in-browser key, not the server configured flag', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.', model: 'm' }, location: { hostname: 'excel-qwen-grok.example.workers.dev' } });
  await boot(sb);
  sb.els('workMode').value = 'chat';
  await ask(sb, 'hello before key');
  assert.match(sb.els('toast').textContent, /Add your Qwen API key/);
  assert.equal(sb.els('settings').classList.contains('hidden'), false);
  sb.els('apiKey').value = 'sk-sp-byo-user-key-001';
  await fire(sb.els('saveKey'), 'click'); await flush();
  await ask(sb, 'hello after key');
  assert.equal(chatCalls(sb).at(-1).body.apiKey, 'sk-sp-byo-user-key-001');
  assert.match(visible(sb.els('conversation')), /Noted\./);
});

test('Save key in hosted pane requests storage access and persists when granted', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.', model: 'm' }, location: { hostname: 'excel-qwen-grok.example.workers.dev' }, blockStorage: true, allowStorageAccess: true });
  await boot(sb);
  assert.match(sb.els('saveKey').textContent, /Save key · 儲存金鑰/);
  sb.els('apiKey').value = 'sk-sp-byo-user-key-001';
  sb.els('rememberKey').checked = true;
  await fire(sb.els('saveKey'), 'click'); await flush();
  assert.match(sb.store['eqgc.keys.v1'], /sk-sp-byo-user-key-001/);
  assert.match(sb.els('toast').textContent, /saved in this browser/);
});

test('Hosted key survives a blocked browser store using Office storage and Clear removes it', async () => {
  const officeStore = {};
  const options = { selection: singleEmpty, used: usedSmall, location: { hostname: 'excel-qwen-grok.example.workers.dev' }, blockStorage: true, officeStore };
  const first = makeSandbox(options);
  await boot(first);
  first.els('apiKey').value = 'sk-sp-byo-user-key-001';
  await fire(first.els('saveKey'), 'click'); await flush();
  assert.equal(officeStore['eqgc.hosted.qwen-key.v1'], 'sk-sp-byo-user-key-001');
  assert.match(first.els('keyStatus').textContent, /remembered in Office/);
  const reopened = makeSandbox(options);
  await boot(reopened);
  await ask(reopened, 'hi');
  assert.equal(chatCalls(reopened).at(-1).body.apiKey, 'sk-sp-byo-user-key-001');
  await fire(reopened.els('clearKey'), 'click'); await flush();
  assert.equal(officeStore['eqgc.hosted.qwen-key.v1'], undefined);
});

test('Hosted stream retries one transient HTTP 500 before a successful answer', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, location: { hostname: 'excel-qwen-grok.example.workers.dev' }, streamFailures: 1, reply: { text: 'Recovered.', model: 'm' } });
  await boot(sb);
  sb.els('apiKey').value = 'sk-sp-byo-user-key-001';
  await fire(sb.els('saveKey'), 'click'); await flush();
  await ask(sb, 'hi');
  assert.equal(chatCalls(sb).length, 2);
  assert.match(visible(sb.els('conversation')), /Recovered\./);
});

test('every pane element id exists in index.html (removed controls cannot break the pane silently)', () => {
  const html = fs.readFileSync(path.join(root, 'public', 'index.html'), 'utf8');
  const ids = new Set([...html.matchAll(/id="([^"]+)"/g)].map(m => m[1]));
  const js = fs.readFileSync(path.join(root, 'public', 'taskpane.js'), 'utf8');
  const used = [...new Set([...js.matchAll(/\$\('([A-Za-z0-9_-]+)'\)/g)].map(m => m[1]))];
  const missing = used.filter(id => !ids.has(id));
  assert.deepEqual(missing, []);
});

test('A stale relay build is reported and blocks the send instead of a silent failure', async () => {
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply: { text: 'Noted.' }, healthBuild: '2026-11-01.1' });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  await ask(sb, 'review the whole workbook');
  assert.equal(chatCalls(sb).length, 0);
  assert.match(sb.els('agentStatus').textContent, /relay is running build 2026-11-01\.1/);
  assert.match(sb.els('toast').textContent, /Restart the relay/);
});

test('A failed sheet scan is reported INCOMPLETE, never as a clean zero-error sheet', async () => {
  const reply = { turns: [
    'Scanning.\n```json\n{"scanRequest":{},"decision":"CONTINUE"}\n```',
    'Noted.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```'
  ] };
  const sb = makeSandbox({
    selection: singleEmpty, used: usedSmall, reply,
    cells: { 'Sheet1!A1': { value: 'Date' }, 'Bad!A1': { value: '#REF!' } },
    sheetNames: ['Sheet1', 'Bad'],
    usedRanges: {
      Sheet1: { address: "'Sheet1'!A1:B2", rowCount: 2, columnCount: 2 },
      Bad: { address: "'Bad'!A1:B2", rowCount: 2, columnCount: 2 }
    }
  });
  await boot(sb);
  sb.els('workMode').value = 'chat';
  const run = sb.Excel.run;
  let syncs = 0;
  sb.Excel.run = callback => run(context => {
    if (/Scanning/.test(sb.els('agentStatus').textContent)) {
      const sync = context.sync;
      context.sync = async () => { syncs += 1; if (syncs === 3) throw new Error('The requested resource does not exist.'); return sync(); };
    }
    return callback(context);
  });
  await ask(sb, 'review the whole workbook');
  const fed = chatCalls(sb)[1].body.messages.at(-1).content;
  assert.match(fed, /sheet: Sheet1 .*INCOMPLETE: The requested resource does not exist/);
  assert.match(fed, /sheet: Bad .*errors 1 \| at A1/);
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
  assert.match(fed, /\[Excel profile Aging!A1:C4 — 3 data rows profiled; 4\/4 requested rows, 3\/3 columns; complete/);
  assert.match(fed, /A \| Entity \| filled 3 \| distinct 2/);
  assert.match(fed, /MUK \(2\)/);
  assert.match(fed, /C \| Total USD \| filled 3 \| distinct 3 \| numeric 3 \| sum 350 \| min 50 \| max 200/);
  assert.match(savedAuditText(sb), /Aging!A1:C4 \(profile\)/);
});

test('a large profile scans a later batch and reports Excel error values with full coverage', async () => {
  const cells = {
    'Ledger!A1': { value: 'Date' }, 'Ledger!B1': { value: 'Account' }, 'Ledger!C1': { value: 'Net' },
    'Ledger!C2': { value: 100 }, 'Ledger!C399': { value: '#VALUE!' }
  };
  const reply = { turns: [
    'Read the ledger.\n```json\n{"actions":[],"profileRequest":{"sheet":"Ledger","range":"A1:BH400"},"decision":"CONTINUE"}\n```',
    'Reviewed.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```'
  ] };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  await ask(sb, 'review the whole ledger');
  const fed = lastContent(sb);
  assert.match(fed, /400\/400 requested rows, 60\/60 columns; complete/);
  assert.match(fed, /C \| Net .*errors 1 \| at C399/);
  assert.equal(chatCalls(sb).length, 2);
});

test('a failed later profile batch reports partial coverage', async () => {
  const reply = { turns: [
    'Profile it.\n```json\n{"actions":[],"profileRequest":{"sheet":"Ledger","range":"A1:BH400"},"decision":"CONTINUE"}\n```',
    'The profile is incomplete.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```'
  ] };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells: { 'Ledger!A1': { value: 'Date' } } });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  const run = sb.Excel.run;
  sb.Excel.run = callback => run(context => {
    if (/Reading sheet profile/.test(sb.els('agentStatus').textContent)) {
      const sync = context.sync;
      let batches = 0;
      context.sync = async () => {
        batches += 1;
        if (batches === 2) throw new Error('The requested resource does not exist.');
        return sync();
      };
    }
    return callback(context);
  });
  await ask(sb, 'review the ledger');
  assert.match(lastContent(sb), /333\/400 requested rows, 60\/60 columns; INCOMPLETE/);
  assert.match(lastContent(sb), /Excel error: The requested resource does not exist/);
});

test('Agent cannot send a requested workbook range while file sharing is off', async () => {
  const reply = { text: 'Read it.\n```json\n{"actions":[],"contextRequest":{"sheet":"Sheet1","range":"A1:B2"},"decision":"CONTINUE"}\n```' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('includeContext').checked = false;
  let reads = 0;
  sb.Excel.run = async () => { reads += 1; throw new Error('Workbook read must be blocked'); };
  await ask(sb, 'review this workbook');
  assert.equal(reads, 0);
  assert.equal(chatCalls(sb).length, 1);
  assert.match(sb.els('agentStatus').textContent, /File sharing off/);
});

test('Agent reports a failed range read and continues without claiming success', async () => {
  const reply = { turns: [
    'Read it.\n```json\n{"actions":[],"contextRequest":{"sheet":"Workbook Review","range":"A40:C69"},"decision":"CONTINUE"}\n```',
    'The range remains unverified.\n```json\n{"actions":[],"decision":"COMPLETE"}\n```'
  ] };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  const run = sb.Excel.run;
  let failed = false;
  sb.Excel.run = callback => {
    if (!failed && /Reading workbook/.test(sb.els('agentStatus').textContent)) {
      failed = true;
      return Promise.reject(new Error('The requested resource does not exist.'));
    }
    return run(callback);
  };
  await ask(sb, 'review Workbook Review');
  assert.equal(failed, true);
  assert.equal(chatCalls(sb).length, 2);
  assert.match(lastContent(sb), /Workbook Review!A40:C69 could not be read/);
  assert.match(lastContent(sb), /Nothing was read/);
  assert.equal(/Observed: SUCCESS/.test(lastContent(sb)), false);
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
  assert.match(savedAuditText(sb), /FAILED/);
  assert.match(savedAuditText(sb), /USD O\/S/);
});

test('addSheet then set/format on the new sheet in one batch applies (real Office.js isNullObject semantics)', async () => {
  const reply = {
    text: 'Build.\n```json\n{"actions":[{"op":"addSheet","name":"Sales Invoice"},{"op":"set","sheet":"Sales Invoice","range":"A1:B2","values":[["Item","Amount"],["Widget",1200]]},{"op":"set","sheet":"Sales Invoice","range":"B3","values":[["=B2*1.05"]]},{"op":"format","sheet":"Sales Invoice","range":"A1:B1","bold":true}],"decision":"COMPLETE"}\n```',
    model: 'm'
  };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'create a demo sales invoice in a new tab');
  assert.ok(sb.writes.some(w => w.kind === 'addSheet' && w.name === 'Sales Invoice'), 'the sheet was created');
  assert.ok(sb.writes.some(w => w.kind === 'value' && w.sheet === 'Sales Invoice'), 'values were written to the new sheet');
  assert.ok(sb.writes.some(w => w.kind === 'formula' && w.sheet === 'Sales Invoice'), 'the formula was written to the new sheet');
  assert.doesNotMatch(savedAuditText(sb), /FAILED/);
  assert.match(visible(sb.els('conversation')), /REVISED 4 CHANGES/);
});

test('a set on a genuinely missing sheet still fails closed and writes nothing', async () => {
  const reply = { text: 'Write.\n```json\n{"actions":[{"op":"set","sheet":"Ghost","range":"A1","values":[["x"]]}],"decision":"COMPLETE"}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, missingSheets: ['Ghost'] });
  await boot(sb);
  sb.els('workMode').value = 'agent';
  sb.els('directRevise').checked = true;
  await ask(sb, 'write to ghost');
  assert.ok(!sb.writes.some(w => w.kind === 'value'), 'nothing was written');
  assert.match(savedAuditText(sb), /FAILED/);
});
