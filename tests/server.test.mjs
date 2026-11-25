import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import http from 'node:http';
import { createRelay, originAllowed, prepareChat, providerDelta, qwenEndpointAllowed, requestHostAllowed, takeSseDataLines } from '../server.mjs';

const qwenUrl = 'https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions';
const secrets = { qwen: 'sk-sp-testkey' };
const messages = [{ role: 'user', content: 'hi' }];

test('Qwen endpoints stay on QwenCloud chat completions', () => {
  assert.ok(qwenEndpointAllowed(qwenUrl));
  assert.ok(qwenEndpointAllowed('https://ap-southeast-1.maas.qwencloudapi.com/compatible-mode/v1/chat/completions'));
  assert.ok(qwenEndpointAllowed('https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions'));
  assert.equal(qwenEndpointAllowed('http://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions'), null);
  assert.equal(qwenEndpointAllowed('https://user:pass@token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions'), null);
  assert.equal(qwenEndpointAllowed('https://evil.example/compatible-mode/v1/chat/completions'), null);
  assert.equal(qwenEndpointAllowed('https://token-plan.maas.qwencloudapi.com.evil.example/compatible-mode/v1/chat/completions'), null);
  assert.equal(qwenEndpointAllowed('https://token-plan.maas.qwencloudapi.com/v1/other'), null);
});

test('relay accepts only the loopback host and same origin', () => {
  assert.equal(requestHostAllowed('127.0.0.1:3000', 3000), true);
  assert.equal(requestHostAllowed('localhost:3000', 3000), true);
  assert.equal(requestHostAllowed('evil.example', 3000), false);
  assert.equal(requestHostAllowed('evil.example:3000', 3000), false);
  assert.equal(originAllowed(undefined, 3000), true);
  assert.equal(originAllowed('http://127.0.0.1:3000', 3000), true);
  assert.equal(originAllowed('http://evil.example', 3000), false);
});

test('Fast disables Qwen thinking; bad models and non-Qwen providers are rejected', () => {
  const qwen = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, fast: true }, secrets);
  assert.equal(qwen.payload.enable_thinking, false);
  assert.equal(qwen.payload.stream, false);
  const grok = prepareChat({ provider: 'grok', messages, fast: true, model: 'grok-4.7' }, secrets);
  assert.equal(grok.status, 400);
  const slow = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, fast: false }, secrets);
  assert.equal(slow.payload.enable_thinking, undefined);
  assert.match(prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, host: 'word' }, secrets).payload.messages[0].content, /Word document/);
  assert.match(prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, mode: 'agent' }, secrets).payload.messages[0].content, /MODE AGENT/);
  assert.equal(prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages }, secrets).payload.messages[0].content.includes('MODE AGENT'), false);
  assert.match(slow.payload.messages[0].content, /Excel workbook/);
  const bad = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, model: '../etc/passwd' }, secrets);
  assert.equal(bad.status, 400);
  const evil = prepareChat({ provider: 'qwen', endpoint: 'https://evil.example/steal', messages }, secrets);
  assert.equal(evil.status, 400);
});

test('Effort sets reasoning_effort on models that support it and never sends a thinking budget', () => {
  for (const effort of ['low', 'medium', 'high', 'xhigh']) {
    const prepared = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, effort }, secrets);
    assert.equal(prepared.payload.enable_thinking, true, effort);
    assert.equal(prepared.payload.reasoning_effort, effort);
    assert.equal(prepared.payload.thinking_budget, undefined);
  }
  const flash = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, model: 'qwen3.8-flash', effort: 'MEDIUM' }, secrets);
  assert.equal(flash.payload.model, 'qwen3.8-flash');
  assert.equal(flash.payload.reasoning_effort, 'medium');
  assert.equal(flash.payload.enable_thinking, true);
  const auto = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, model: 'auto', effort: 'high' }, secrets);
  assert.equal(auto.payload.model, 'auto');
  assert.equal(auto.payload.enable_thinking, true);
  assert.equal(auto.payload.reasoning_effort, 'high');
  assert.equal(auto.payload.thinking_budget, undefined);
  const unsupported = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, model: 'qwen3.7-plus', effort: 'high' }, secrets);
  assert.equal(unsupported.status, 400);
});

test('Effort Off, none, and Fast turn thinking off without reasoning_effort', () => {
  for (const effort of ['off', 'none', 'OFF']) {
    const prepared = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, effort }, secrets);
    assert.equal(prepared.payload.enable_thinking, false, effort);
    assert.equal(prepared.payload.reasoning_effort, undefined);
  }
  const fastWins = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, fast: true, effort: 'xhigh' }, secrets);
  assert.equal(fastWins.payload.enable_thinking, false);
  assert.equal(fastWins.payload.reasoning_effort, undefined);
  const omitted = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages }, secrets);
  assert.equal(omitted.payload.enable_thinking, undefined);
  assert.equal(omitted.payload.reasoning_effort, undefined);
});

test('an unknown Effort is rejected', () => {
  const bad = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, effort: 'max' }, secrets);
  assert.equal(bad.status, 400);
  assert.match(bad.error, /effort/i);
});

test('SSE tails and provider failures are not dropped', () => {
  const partial = takeSseDataLines('data: {"model":"qwen4-max","choices":[{"delta":{"content":"Hi"}}]}');
  assert.equal(partial.data.length, 0);
  const flushed = takeSseDataLines(partial.rest + '\n');
  assert.equal(providerDelta('qwen', JSON.parse(flushed.data[0])).content, 'Hi');
  assert.equal(providerDelta('qwen', { model: 'qwen4-max', choices: [{ delta: { content: 'Yo' } }] }).content, 'Yo');
  assert.match(providerDelta('qwen', { error: { message: 'bad key' } }).error, /bad key/);
});

test('skills store lists, selects, and saves Copilot-format skills without overwriting', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-skills-'));
  fs.mkdirSync(path.join(dir, 'monthly-report-review'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'monthly-report-review', 'SKILL.md'), '---\nname: monthly-report-review\ndescription: Review the whole monthly report workbook and reconciliations.\n---\n\n# Review\n\nScan every sheet.\n');
  fs.mkdirSync(path.join(dir, 'wrong-name'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'wrong-name', 'SKILL.md'), '---\nname: other\ndescription: x\n---\nbody');
  process.env.SKILLS_ROOT = dir;
  const { server } = createRelay();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  try {
    const list = await rawRequest(port, { path: '/api/skills' });
    assert.equal(list.status, 200);
    assert.match(list.body, /monthly-report-review/);
    assert.equal(list.body.includes('wrong-name'), false);
    const ctx = await rawRequest(port, { path: '/api/skills/context?question=' + encodeURIComponent('Please review the whole report workbook') });
    assert.match(ctx.body, /monthly-report-review/);
    const forced = await rawRequest(port, { path: '/api/skills/context?question=' + encodeURIComponent('hello @monthly-report-review') });
    assert.match(forced.body, /monthly-report-review/);
    const save = await rawRequest(port, { method: 'POST', path: '/api/skills/save', body: { id: 'oct-close-review', description: 'October close review', body: '# Oct\n\nSteps.' } });
    assert.equal(save.status, 200);
    assert.ok(fs.existsSync(path.join(dir, 'oct-close-review', 'SKILL.md')));
    const dup = await rawRequest(port, { method: 'POST', path: '/api/skills/save', body: { id: 'oct-close-review', description: 'x', body: 'y' } });
    assert.equal(dup.status, 400);
    const bad = await rawRequest(port, { method: 'POST', path: '/api/skills/save', body: { id: 'Bad_Name', description: 'x', body: 'y' } });
    assert.equal(bad.status, 400);
  } finally {
    await new Promise(resolve => server.close(resolve));
    delete process.env.SKILLS_ROOT;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('sessions and memory persist in per-user app data and round-trip for resume', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-appdata-'));
  process.env.LOCALAPPDATA = dir;
  const { server } = createRelay();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  try {
    const save = await rawRequest(port, { method: 'POST', path: '/api/sessions/save', body: { id: 's-test-1', prompt: 'review whole file', workbook: 'MPL', outcome: 'Completed', messages: [{ role: 'user', content: 'review whole file' }, { role: 'assistant', content: 'Report.' }], checkpoint: { prompt: 'review whole file', progress: [], observation: '' } } });
    assert.equal(save.status, 200);
    const list = await rawRequest(port, { path: '/api/sessions' });
    assert.match(list.body, /s-test-1/);
    const read = await rawRequest(port, { path: '/api/sessions/read?id=s-test-1' });
    assert.equal(read.status, 200);
    assert.match(read.body, /Report\./);
    assert.match(read.body, /checkpoint/);
    const bad = await rawRequest(port, { method: 'POST', path: '/api/sessions/save', body: { id: '../evil', messages: [] } });
    assert.equal(bad.status, 400);
    const missing = await rawRequest(port, { path: '/api/sessions/read?id=s-nope-9' });
    assert.equal(missing.status, 404);
    const del = await rawRequest(port, { method: 'POST', path: '/api/sessions/delete', body: { id: 's-test-1' } });
    assert.equal(del.status, 200);
    const gone = await rawRequest(port, { path: '/api/sessions/read?id=s-test-1' });
    assert.equal(gone.status, 404);
    const badDel = await rawRequest(port, { method: 'POST', path: '/api/sessions/delete', body: { id: '../evil' } });
    assert.equal(badDel.status, 400);
    const note = await rawRequest(port, { method: 'POST', path: '/api/memory/add', body: { note: 'MPL source of truth is 30_GJ_2026' } });
    assert.equal(note.status, 200);
    const memory = await rawRequest(port, { path: '/api/memory' });
    assert.match(memory.body, /30_GJ_2026/);
  } finally {
    await new Promise(resolve => server.close(resolve));
    delete process.env.LOCALAPPDATA;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

function minimalPdf(text) {
  const objs = [];
  objs[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>';
  objs[3] = '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>';
  const stream = text ? `BT /F1 24 Tf 72 720 Td (${text}) Tj ET` : '';
  objs[4] = `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`;
  objs[5] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';
  let pdf = '%PDF-1.4\n';
  const offsets = [0];
  for (let i = 1; i <= 5; i++) { offsets[i] = pdf.length; pdf += `${i} 0 obj\n${objs[i]}\nendobj\n`; }
  const xref = pdf.length;
  pdf += 'xref\n0 6\n0000000000 65535 f \n' + offsets.slice(1).map(o => String(o).padStart(10, '0') + ' 00000 n \n').join('');
  pdf += `trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return Buffer.from(pdf, 'binary');
}

test('extract turns XLSX into sheet text and rejects unsupported types', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-extract-'));
  process.env.LOCALAPPDATA = dir;
  const ExcelJS = (await import('exceljs')).default;
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.addRow(['Voucher', 'Debit']);
  sheet.addRow(['GJ001', 100]);
  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  const { server } = createRelay();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  try {
    const ok = await rawRequest(port, { method: 'POST', path: '/api/extract', body: { name: 'book.xlsx', data: buffer.toString('base64') } });
    assert.equal(ok.status, 200);
    assert.match(ok.body, /\[Sheet: Data\]/);
    assert.match(ok.body, /GJ001 \| 100/);
    const bad = await rawRequest(port, { method: 'POST', path: '/api/extract', body: { name: 'tool.exe', data: 'AA==' } });
    assert.equal(bad.status, 400);
    const img = await rawRequest(port, { method: 'POST', path: '/api/extract', body: { name: 'scan.png', data: 'iVBORw0KGgo=' } });
    assert.equal(img.status, 400);
    assert.match(img.body, /Qwen key/);
    const empty = await rawRequest(port, { method: 'POST', path: '/api/extract', body: { name: 'book.xlsx', data: '' } });
    assert.equal(empty.status, 400);
    const { spawnSync } = await import('node:child_process');
    const hasPoppler = !spawnSync('pdftotext', ['-v'], { windowsHide: true }).error;
    if (hasPoppler) {
      const scan = await rawRequest(port, { method: 'POST', path: '/api/extract', body: { name: 'scan.pdf', data: minimalPdf('').toString('base64') } });
      assert.equal(scan.status, 400);
      assert.match(scan.body, /no extractable text layer/);
      const texted = await rawRequest(port, { method: 'POST', path: '/api/extract', body: { name: 'text.pdf', data: minimalPdf('Hello PDF attach test').toString('base64') } });
      assert.equal(texted.status, 200);
      assert.match(texted.body, /Hello PDF attach test/);
    }
  } finally {
    await new Promise(resolve => server.close(resolve));
    delete process.env.LOCALAPPDATA;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

function rawRequest(port, { path = '/', method = 'GET', host, origin, body }) {
  const payload = body ? JSON.stringify(body) : null;
  const headers = { Host: host || `127.0.0.1:${port}` };
  if (origin) headers.Origin = origin;
  if (payload) {
    headers['Content-Type'] = 'application/json';
    headers['Content-Length'] = Buffer.byteLength(payload);
  }
  return new Promise((resolve, reject) => {
    const req = http.request({ hostname: '127.0.0.1', port, path, method, headers }, res => {
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString('utf8') }));
    });
    req.on('error', reject);
    if (payload) req.write(payload);
    req.end();
  });
}

test('HTTP relay rejects foreign hosts, cross-origin calls, path escape, and non-Qwen endpoints', async () => {
  const { server } = createRelay();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  try {
    const health = await rawRequest(port, { path: '/api/health' });
    assert.equal(health.status, 200);
    assert.match(health.body, /"ok":true/);
    assert.equal((await rawRequest(port, { path: '/api/health', host: 'evil.example' })).status, 403);
    assert.equal((await rawRequest(port, { path: '/api/health', origin: 'http://evil.example' })).status, 403);
    const escaped = await rawRequest(port, { path: '/%2e%2e/%2e%2e/package.json' });
    assert.notEqual(escaped.status, 200);
    assert.equal(escaped.body.includes('excel-qwen-grok-copilot'), false);
    const configured = await rawRequest(port, { method: 'POST', path: '/api/configure', body: { provider: 'qwen', apiKey: 'sk-sp-testkey' } });
    assert.equal(configured.status, 200);
    const evil = await rawRequest(port, { method: 'POST', path: '/api/chat', body: { provider: 'qwen', endpoint: 'https://evil.example/compatible-mode/v1/chat/completions', messages } });
    assert.equal(evil.status, 400);
    assert.match(evil.body, /QwenCloud/);
    const pane = await rawRequest(port, { path: '/taskpane.js' });
    assert.equal(pane.status, 200);
    assert.match(pane.body, /sendPrompt/);
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});
