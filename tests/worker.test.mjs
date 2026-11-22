import test from 'node:test';
import assert from 'node:assert/strict';
import { handleRelayRequest } from '../worker.mjs';

const token = 'relay-token-0123456789';
const env = { RELAY_TOKEN: token, QWEN_API_KEY: 'sk-sp-testkey', GROK_API_KEY: '' };

function call(path, { method = 'GET', token: presented, origin, body } = {}) {
  const headers = {};
  if (presented) headers.Authorization = `Bearer ${presented}`;
  if (origin) headers.Origin = origin;
  return handleRelayRequest(new Request(`https://excel.example${path}`, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined
  }), env);
}

test('the hosted relay rejects a missing or wrong token and a foreign origin', async () => {
  assert.equal((await call('/api/health')).status, 401);
  assert.equal((await call('/api/health', { token: 'relay-token-0000000000' })).status, 401);
  assert.equal((await call('/api/health', { token, origin: 'https://evil.example' })).status, 403);
  const health = await call('/api/health', { token });
  assert.equal(health.status, 200);
  const data = await health.json();
  assert.equal(data.hosted, true);
  assert.equal(data.configured.qwen, true);
  assert.equal(data.configured.grok, false);
});

test('the hosted relay does not accept a pasted provider key and still blocks a non-Qwen URL', async () => {
  const configure = await call('/api/configure', { method: 'POST', token, body: { provider: 'grok', apiKey: 'xai-live-key-value' } });
  assert.equal(configure.status, 400);
  assert.match(await configure.text(), /Worker secrets/);
  const evil = await call('/api/chat', {
    method: 'POST',
    token,
    body: { provider: 'qwen', endpoint: 'https://evil.example/compatible-mode/v1/chat/completions', messages: [{ role: 'user', content: 'hi' }] }
  });
  assert.equal(evil.status, 400);
  assert.match(await evil.text(), /QwenCloud/);
});
