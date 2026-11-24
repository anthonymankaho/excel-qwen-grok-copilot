import test from 'node:test';
import assert from 'node:assert/strict';
import { handleRelayRequest } from '../worker.mjs';

// Marketplace v1: BYO-key pass-through. No relay token, no owner key in the serving path.
const env = { RELAY_TOKEN: 'legacy-should-be-ignored', QWEN_API_KEY: 'sk-sp-owner-key-must-not-be-used' };
const OWN_KEY = 'sk-sp-byo-user-key-001';
const ENDPOINT = 'https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions';

function call(path, { method = 'GET', origin, body } = {}) {
  const headers = {};
  if (origin) headers.Origin = origin;
  return handleRelayRequest(new Request(`https://excel.example${path}`, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined
  }), env);
}

function stubFetch(handler) {
  const original = globalThis.fetch;
  const seen = [];
  globalThis.fetch = async (url, options) => {
    seen.push({ url: String(url), headers: options?.headers || {} });
    return handler(url, options);
  };
  return { seen, restore: () => { globalThis.fetch = original; } };
}

const okJson = data => new Response(JSON.stringify(data), { status: 200, headers: { 'Content-Type': 'application/json' } });

test('hosted health is open, BYO-flagged, and never reports an owner key', async () => {
  const health = await call('/api/health');
  assert.equal(health.status, 200);
  const data = await health.json();
  assert.equal(data.hosted, true);
  assert.equal(data.byo, true);
  assert.equal(data.configured.qwen, false);
  assert.equal((await call('/api/health', { origin: 'https://evil.example' })).status, 403);
});

test('chat without a valid own key is 401 and configure never stores keys', async () => {
  const missing = await call('/api/chat', { method: 'POST', body: { provider: 'qwen', messages: [{ role: 'user', content: 'hi' }] } });
  assert.equal(missing.status, 401);
  const bad = await call('/api/chat', { method: 'POST', body: { provider: 'qwen', apiKey: 'not-a-key', messages: [{ role: 'user', content: 'hi' }] } });
  assert.equal(bad.status, 401);
  const configure = await call('/api/configure', { method: 'POST', body: { provider: 'qwen', apiKey: OWN_KEY } });
  assert.equal(configure.status, 400);
  assert.match(await configure.text(), /per request/);
});

test('chat forwards the caller key verbatim and never the owner key', async () => {
  const stub = stubFetch(() => okJson({ model: 'qwen4-max', choices: [{ message: { content: 'Answer.' } }] }));
  try {
    const res = await call('/api/chat', { method: 'POST', body: { provider: 'qwen', apiKey: OWN_KEY, endpoint: ENDPOINT, messages: [{ role: 'user', content: 'hi' }] } });
    assert.equal(res.status, 200);
    assert.match(await res.text(), /Answer\./);
    assert.equal(stub.seen.length, 1);
    assert.equal(stub.seen[0].headers.Authorization, `Bearer ${OWN_KEY}`);
    assert.equal(JSON.stringify(stub.seen).includes('owner-key'), false);
  } finally { stub.restore(); }
});

test('chat-stream forwards the caller key and streams deltas', async () => {
  const stub = stubFetch(() => new Response('data: {"choices":[{"delta":{"content":"Yo"}}]}\n\ndata: [DONE]\n\n', { status: 200, headers: { 'Content-Type': 'text/event-stream' } }));
  try {
    const res = await call('/api/chat-stream', { method: 'POST', body: { provider: 'qwen', apiKey: OWN_KEY, endpoint: ENDPOINT, messages: [{ role: 'user', content: 'hi' }] } });
    assert.equal(res.status, 200);
    const text = await res.text();
    assert.match(text, /Yo/);
    assert.equal(stub.seen[0].headers.Authorization, `Bearer ${OWN_KEY}`);
  } finally { stub.restore(); }
});

test('a non-Qwen endpoint is still rejected even with a valid own key', async () => {
  const evil = await call('/api/chat', { method: 'POST', body: { provider: 'qwen', apiKey: OWN_KEY, endpoint: 'https://evil.example/compatible-mode/v1/chat/completions', messages: [{ role: 'user', content: 'hi' }] } });
  assert.equal(evil.status, 400);
  assert.match(await evil.text(), /QwenCloud/);
});

test('per-key rate limit returns 429 after 60 requests in a minute', async () => {
  const stub = stubFetch(() => okJson({ model: 'qwen4-max', choices: [{ message: { content: 'ok' } }] }));
  const key = 'sk-sp-rate-limit-probe-002';
  try {
    let last;
    for (let i = 0; i < 61; i++) {
      last = await call('/api/chat', { method: 'POST', body: { provider: 'qwen', apiKey: key, endpoint: ENDPOINT, messages: [{ role: 'user', content: 'hi' }] } });
    }
    assert.equal(last.status, 429);
    assert.match(await last.text(), /60 requests per minute/);
  } finally { stub.restore(); }
});