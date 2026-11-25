import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const source = fs.readFileSync(path.join(root, 'public', 'taskpane.js'), 'utf8');

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

function makeSandbox({ selection, used, reply, seedKeys, cells }) {
  const els = new Map();
  const defaults = { provider: 'qwen', activeProvider: 'qwen', qwenModel: 'qwen4-max', qwenEndpoint: 'https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions' };
  const get = id => { if (!els.has(id)) { const el = makeElement(id); if (defaults[id]) el.value = defaults[id]; if (id === 'includeContext' || id === 'rememberKey') el.checked = true; els.set(id, el); } return els.get(id); };
  const store = seedKeys ? { 'eqgc.keys.v1': JSON.stringify(seedKeys) } : {};
  const writes = [];
  const calls = [];
  const cellStore = {};
  for (const [k, v] of Object.entries(cells || {})) cellStore[k] = { value: v.value ?? '', formula: v.formula ?? '', nf: v.nf ?? 'General' };
  const range = (obj) => ({ load() {}, isNullObject: false, getCell: () => target('Selection', 'A1'), ...obj });
  const cur = key => (cellStore[key] ||= { value: '', formula: '', nf: 'General' });
  const target = (sheet, cell) => {
    const key = `${sheet}!${cell}`;
    return {
      load() {},
      get values() { return [[cur(key).value]]; },
      get formulas() { return [[cur(key).formula]]; },
      get numberFormat() { return [[cur(key).nf]]; },
      set numberFormat(v) { cur(key).nf = v[0][0]; writes.push({ sheet, cell, kind: 'numberFormat', format: v[0][0] }); },
      set formulas(v) { cur(key).formula = v[0][0]; cur(key).value = v[0][0]; writes.push({ sheet, cell, kind: 'formula', value: v[0][0] }); },
      set values(v) {
        const x = v[0][0];
        if (typeof x === 'string' && x.startsWith('=')) { cur(key).formula = x; cur(key).value = x; writes.push({ sheet, cell, kind: 'formula', value: x }); }
        else { cur(key).value = x == null ? '' : x; cur(key).formula = ''; writes.push({ sheet, cell, kind: x == null ? 'clear' : 'value', value: x }); }
      },
      clear() { cur(key).value = ''; cur(key).formula = ''; writes.push({ sheet, cell, kind: 'clear' }); }
    };
  };
  const colIdx = s => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
  const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const dimsOf = addr => {
    const A = addr.toUpperCase();
    if (/^[A-Z]+[0-9]+$/.test(A)) { const m = A.match(/^([A-Z]+)([0-9]+)$/); return { r1: +m[2], r2: +m[2], c1: colIdx(m[1]), c2: colIdx(m[1]) }; }
    if (/^[A-Z]+:[A-Z]+$/.test(A)) { const [a, b] = A.split(':'); const rows = Object.keys(cellStore).map(k => +k.split('!')[1].match(/[0-9]+$/)[0]); return { r1: 1, r2: Math.max(3, ...rows), c1: colIdx(a), c2: colIdx(b) }; }
    if (/^[0-9]+:[0-9]+$/.test(A)) { const [a, b] = A.split(':').map(Number); const cols = Object.keys(cellStore).map(k => colIdx(k.split('!')[1].match(/^[A-Z]+/)[0])); return { r1: a, r2: b, c1: 1, c2: Math.max(3, ...cols) }; }
    const [a, b] = A.split(':'); const p = s => { const m = s.match(/^([A-Z]+)([0-9]+)$/); return { r: +m[2], c: colIdx(m[1]) }; };
    const p1 = p(a), p2 = p(b);
    return { r1: Math.min(p1.r, p2.r), r2: Math.max(p1.r, p2.r), c1: Math.min(p1.c, p2.c), c2: Math.max(p1.c, p2.c) };
  };
  const gridTarget = (sheet, addr) => {
    const d = dimsOf(addr);
    const keys = [];
    for (let r = d.r1; r <= d.r2; r++) for (let c = d.c1; c <= d.c2; c++) keys.push(`${sheet}!${colName(c)}${r}`);
    const width = d.c2 - d.c1 + 1;
    const rowsOf = get => { const rows = []; for (let r = d.r1; r <= d.r2; r++) rows.push(keys.filter(k => k.endsWith(`${r}`)).map(get)); return rows; };
    return {
      load() {},
      get values() { return rowsOf(k => cur(k).value); },
      get formulas() { return rowsOf(k => cur(k).formula); },
      get numberFormat() { return rowsOf(k => cur(k).nf); },
      set values(v) { v.forEach((row, i) => row.forEach((x, j) => { const k = keys[i * width + j]; cur(k).value = x == null ? '' : x; cur(k).formula = typeof x === 'string' && x.startsWith('=') ? x : ''; })); writes.push({ sheet, addr, kind: 'range', value: v }); },
      set numberFormat(v) { v.forEach((row, i) => row.forEach((x, j) => { cur(keys[i * width + j]).nf = x; })); writes.push({ sheet, addr, kind: 'numberFormat', format: v[0][0] }); },
      insert(dir) { writes.push({ sheet, addr, kind: dir === 'Down' ? 'insertRows' : 'insertColumns' }); },
      delete(dir) { writes.push({ sheet, addr, kind: dir === 'Up' ? 'deleteRows' : 'deleteColumns' }); },
      clear() { keys.forEach(k => { cur(k).value = ''; cur(k).formula = ''; }); writes.push({ sheet, addr, kind: 'clear' }); }
    };
  };
  const getRange = (sheet, addr) => (/^[A-Z]{1,3}[1-9][0-9]{0,6}$/i.test(addr) ? target(sheet, addr) : gridTarget(sheet, addr));
  const sandbox = {
    console, JSON, Math, String, Array, Object, Promise, RegExp, TextEncoder, TextDecoder,
    setTimeout: () => 0, clearTimeout: () => {},
    localStorage: {
      getItem: k => (k in store ? store[k] : null),
      setItem: (k, v) => { store[k] = String(v); },
      removeItem: k => { delete store[k]; }
    },
    document: { getElementById: get, createElement: tag => makeElement(tag), querySelectorAll: () => [] },
    Office: { HostType: { Excel: 'Excel' }, onReady(cb) { sandbox.officeCb = cb; } },
    Excel: {
      InsertShiftDirection: { down: 'Down', right: 'Right' },
      DeleteShiftDirection: { up: 'Up', left: 'Left' },
      run: fn => fn({
        workbook: {
          getSelectedRange: () => range(selection),
          worksheets: {
            getActiveWorksheet: () => ({ getUsedRangeOrNullObject: () => range(used) }),
            getItem: sheet => ({ getRange: addr => getRange(sheet, addr) })
          }
        },
        sync: async () => {}
      })
    },
    fetch: async (url, opts) => {
      const body = opts?.body ? JSON.parse(opts.body) : null;
      calls.push({ url, body });
      if (url.endsWith('/api/health')) return { ok: true, status: 200, json: async () => ({ ok: true, configured: { qwen: true, grok: false } }) };
      const data = reply || { text: 'mock answer', model: 'test-model' };
      if (url.endsWith('/api/chat-stream')) {
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
    { sheet: 'Sheet1', cell: 'B2', kind: 'value', value: 2 },
    { sheet: 'Sheet1', cell: 'D5', kind: 'value', value: 123 }
  ]);
  assert.match(card.children[0].textContent, /UNDONE — RESTORED 2 OPERATIONS/);
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

test('range write shows grid preview; undo restores the old grid', async () => {
  const reply = { text: 'x\n```json\n{"edits":[{"sheet":"Sheet1","op":"writeRange","range":"A1:B1","values":[[10,20]]}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells: { 'Sheet1!A1': { value: 1 }, 'Sheet1!B1': { value: 2 } } });
  await boot(sb); await ask(sb, 'revise');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  const rows = card.children.filter(c => c.className === 'edit-row');
  assert.match(rows[0].textContent, /write 1×2/);
  assert.match(rows[0].textContent, /10\t20/);
  const actions = card.children.find(c => c.children.length === 2);
  await fire(actions.children[0], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', addr: 'A1:B1', kind: 'range', value: [[10, 20]] }]);
  await fire(actions.children[2], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', addr: 'A1:B1', kind: 'range', value: [[1, 2]] }]);
});

test('insert rows undo deletes the inserted rows', async () => {
  const reply = { text: 'x\n```json\n{"edits":[{"sheet":"Sheet1","op":"insertRows","at":2,"count":2}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb); await ask(sb, 'add rows');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  const actions = card.children.find(c => c.children.length === 2);
  await fire(actions.children[0], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', addr: '2:3', kind: 'insertRows' }]);
  await fire(actions.children[2], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', addr: '2:3', kind: 'deleteRows' }]);
});

test('delete rows undo re-inserts and restores content', async () => {
  const reply = { text: 'x\n```json\n{"edits":[{"sheet":"Sheet1","op":"deleteRows","at":5,"count":2}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply, cells: { 'Sheet1!A5': { value: 'x' }, 'Sheet1!A6': { value: 'y' } } });
  await boot(sb); await ask(sb, 'remove rows');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  const actions = card.children.find(c => c.children.length === 2);
  await fire(actions.children[0], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', addr: '5:6', kind: 'deleteRows' }]);
  await fire(actions.children[2], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-3), [
    { sheet: 'Sheet1', addr: '5:6', kind: 'insertRows' },
    { sheet: 'Sheet1', addr: '5:6', kind: 'range', value: [['x', null, null], ['y', null, null]] },
    { sheet: 'Sheet1', addr: '5:6', kind: 'numberFormat', format: 'General' }
  ]);
});

test('number format undo restores the previous format', async () => {
  const reply = { text: 'x\n```json\n{"edits":[{"sheet":"Sheet1","op":"numberFormat","range":"B1:B2","format":"0.0%"}]}\n```', model: 'm' };
  const sb = makeSandbox({ selection: singleEmpty, used: usedSmall, reply });
  await boot(sb); await ask(sb, 'format percent');
  const card = sb.els('conversation').children.find(c => String(c.className).includes('edits-card'));
  const actions = card.children.find(c => c.children.length === 2);
  await fire(actions.children[0], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', addr: 'B1:B2', kind: 'numberFormat', format: '0.0%' }]);
  await fire(actions.children[2], 'click'); await flush(); await flush();
  assert.deepEqual(sb.writes.slice(-1), [{ sheet: 'Sheet1', addr: 'B1:B2', kind: 'numberFormat', format: 'General' }]);
});
