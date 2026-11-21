import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { createRelay, originAllowed, prepareChat, providerDelta, qwenEndpointAllowed, requestHostAllowed, takeSseDataLines } from '../server.mjs';

const qwenUrl = 'https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions';
const secrets = { qwen: 'sk-sp-testkey', grok: 'xai-test-key' };
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

test('Fast disables Qwen thinking and lowers Grok effort; bad models are rejected', () => {
  const qwen = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, fast: true }, secrets);
  assert.equal(qwen.payload.enable_thinking, false);
  assert.equal(qwen.payload.stream, false);
  const grok = prepareChat({ provider: 'grok', messages, fast: true, model: 'grok-4.7' }, secrets);
  assert.deepEqual(grok.payload.reasoning, { effort: 'low' });
  assert.equal(grok.url.hostname, 'api.x.ai');
  const slow = prepareChat({ provider: 'grok', messages, fast: false }, secrets);
  assert.equal(slow.payload.reasoning, undefined);
  const bad = prepareChat({ provider: 'qwen', endpoint: qwenUrl, messages, model: '../etc/passwd' }, secrets);
  assert.equal(bad.status, 400);
  const evil = prepareChat({ provider: 'qwen', endpoint: 'https://evil.example/steal', messages }, secrets);
  assert.equal(evil.status, 400);
});

test('SSE tails and provider failures are not dropped', () => {
  const partial = takeSseDataLines('data: {"type":"response.output_text.delta","delta":"Hi"}');
  assert.equal(partial.data.length, 0);
  const flushed = takeSseDataLines(partial.rest + '\n');
  assert.equal(providerDelta('grok', JSON.parse(flushed.data[0])).content, 'Hi');
  assert.equal(providerDelta('grok', { type: 'response.completed', response: { model: 'grok-4.7' } }).model, 'grok-4.7');
  assert.match(providerDelta('grok', { type: 'response.failed', error: { message: 'overloaded' } }).error, /overloaded/);
  assert.match(providerDelta('grok', { type: 'response.incomplete', response: { incomplete_details: { reason: 'max_output_tokens' } } }).error, /max_output_tokens/);
  assert.equal(providerDelta('qwen', { model: 'qwen4-max', choices: [{ delta: { content: 'Yo' } }] }).content, 'Yo');
  assert.match(providerDelta('qwen', { error: { message: 'bad key' } }).error, /bad key/);
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
