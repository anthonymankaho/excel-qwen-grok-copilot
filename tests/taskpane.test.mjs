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
  return el;
}
function fire(el, type, ev = {}) {
  return Promise.all((el.listeners[type] || []).map(fn => fn({ preventDefault() {}, ...ev })));
}
const flush = () => new Promise(resolve => setImmediate(resolve));

function makeSandbox({ selection, used, reply, seedKeys }) {
  const els = new Map();
  const defaults = { provider: 'qwen', activeProvider: 'qwen', qwenModel: 'qwen4-max', qwenEndpoint: 'https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions' };
  const get = id => { if (!els.has(id)) { const el = makeElement(id); if (defaults[id]) el.value = defaults[id]; if (id === 'includeContext' || id === 'rememberKey') el.checked = true; els.set(id, el); } return els.get(id); };
  const store = seedKeys ? { 'eqgc.keys.v1': JSON.stringify(seedKeys) } : {};
  const writes = [];
  const calls = [];
  const range = (obj) => ({ load() {}, isNullObject: false, getCell: () => target('Selection', 'A1'), ...obj });
  const target = (sheet, cell) => ({
    set formulas(v) { writes.push({ sheet, cell, kind: 'formula', value: v[0][0] }); },
    set values(v) { writes.push({ sheet, cell, kind: 'value', value: v[0][0] }); }
  });
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
      run: fn => fn({
        workbook: {
          getSelectedRange: () => range(selection),
          worksheets: {
            getActiveWorksheet: () => ({ getUsedRangeOrNullObject: () => range(used) }),
            getItem: sheet => ({ getRange: cell => target(sheet, cell) })
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
  assert.match(bubble.children.at(-1).textContent, /```json/);
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
