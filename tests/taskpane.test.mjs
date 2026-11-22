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
  function makeWritableRange(sheetName, addr) {
    const cell = target(sheetName, addr);
    cell.getCell = (r, c) => makeWritableRange(sheetName, shiftAddress(addr, r, c));
    const font = { load() {} };
    for (const key of ['bold', 'italic', 'color']) {
      let current;
      Object.defineProperty(font, key, { get() { return current; }, set(v) { current = v; writes.push({ sheet: sheetName, cell: addr, kind: 'font', key, value: v }); } });
    }
    const fill = { load() {} };
    let fillColor;
    Object.defineProperty(fill, 'color', { get() { return fillColor; }, set(v) { fillColor = v; writes.push({ sheet: sheetName, cell: addr, kind: 'fill', value: v }); } });
    cell.format = {
      font, fill, load() {},
      autofitColumns() { writes.push({ sheet: sheetName, cell: addr, kind: 'autofitColumns' }); },
      autofitRows() { writes.push({ sheet: sheetName, cell: addr, kind: 'autofitRows' }); }
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
        getItem(tableName) { return { delete() { writes.push({ kind: 'deleteTable', name: tableName }); } }; }
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
            search() { return { load() {}, items: [{ insertText() { writes.push({ kind: 'wordRemove' }); } }] }; }
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
