import assert from 'node:assert/strict';
import { once } from 'node:events';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import { createRelay } from '../server.mjs';

const SINGAPORE_PAYG_CHAT = 'https://ap-southeast-1.maas.qwencloudapi.com/compatible-mode/v1/chat/completions';
const FAKE_QWEN_KEY = 'sk-ws-test-only-not-a-real-key';
const FAKE_GROK_KEY = 'xai-test-key-not-real';

function jsonResponse(body, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

async function start(fetchImpl) {
  const relay = createRelay({ fetchImpl });
  relay.server.listen(0, '127.0.0.1');
  await once(relay.server, 'listening');
  return { relay, base: `http://127.0.0.1:${relay.server.address().port}` };
}

async function stop(relay) {
  relay.server.close();
  await once(relay.server, 'close');
}

async function post(base, path, body) {
  const response = await fetch(`${base}${path}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(body)
  });
  const data = await response.json();
  return { status: response.status, data, raw: JSON.stringify(data) };
}

test('settings default to the documented Singapore pay-as-you-go chat completions URL', async () => {
  const html = readFileSync(new URL('../public/index.html', import.meta.url), 'utf8');
  const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8');
  assert.match(html, new RegExp(`value="${SINGAPORE_PAYG_CHAT.replaceAll('/', '\\/')}"`));
  assert.match(html, /Qwen · QwenCloud/);
  assert.doesNotMatch(html, /Alibaba Model Studio|dashscope|aliyuncs\.com|token-plan\.maas/i);
  assert.match(readme, new RegExp(SINGAPORE_PAYG_CHAT.replaceAll('/', '\\/')));
  assert.match(readme, /https:\/\/token-plan\.maas\.qwencloudapi\.com\/compatible-mode\/v1/);
  assert.match(readme, /https:\/\/api\.x\.ai\/v1\/responses/);

  const { relay, base } = await start(async () => { throw new Error('provider fetch should not run'); });
  try {
    const page = await fetch(`${base}/`);
    const served = await page.text();
    assert.equal(page.status, 200);
    assert.match(served, new RegExp(`value="${SINGAPORE_PAYG_CHAT.replaceAll('/', '\\/')}"`));
    assert.match(served, /value="qwen3\.8-max"/);
    assert.match(served, /value="grok-4\.7"/);
  } finally {
    await stop(relay);
  }
});

test('qwen chat posts to the Singapore pay-as-you-go URL and does not return the key', async () => {
  let captured;
  const fetchImpl = async (url, init) => {
    captured = { url: String(url), init };
    return jsonResponse({ choices: [{ message: { content: 'mocked answer' } }], model: 'qwen3.8-max' });
  };
  const { relay, base } = await start(fetchImpl);
  try {
    const configured = await post(base, '/api/configure', { provider: 'qwen', apiKey: FAKE_QWEN_KEY });
    assert.equal(configured.status, 200);
    assert.equal(configured.data.configured.qwen, true);
    assert.equal(configured.raw.includes(FAKE_QWEN_KEY), false);

    const chat = await post(base, '/api/chat', {
      provider: 'qwen',
      model: 'qwen3.8-max',
      endpoint: SINGAPORE_PAYG_CHAT,
      messages: [{ role: 'user', content: 'Hello' }]
    });
    assert.equal(chat.status, 200);
    assert.equal(chat.data.text, 'mocked answer');
    assert.equal(chat.raw.includes(FAKE_QWEN_KEY), false);
    assert.equal(captured.url, SINGAPORE_PAYG_CHAT);
    assert.equal(captured.init.method, 'POST');
    assert.equal(captured.init.headers.Authorization, `Bearer ${FAKE_QWEN_KEY}`);
    const payload = JSON.parse(captured.init.body);
    assert.equal(payload.model, 'qwen3.8-max');
    assert.equal(payload.stream, false);
    assert.equal(payload.messages.at(-1).content, 'Hello');
    assert.equal(payload.messages[0].role, 'system');
  } finally {
    await stop(relay);
  }
});

test('the relay does not rewrite a caller-supplied Qwen URL, including another documented pay-as-you-go region', async () => {
  const virginia = 'https://us-east-1.maas.qwencloudapi.com/compatible-mode/v1/chat/completions';
  let capturedUrl = '';
  const fetchImpl = async url => {
    capturedUrl = String(url);
    return jsonResponse({ choices: [{ message: { content: 'region ok' } }] });
  };
  const { relay, base } = await start(fetchImpl);
  try {
    assert.equal((await post(base, '/api/configure', { provider: 'qwen', apiKey: FAKE_QWEN_KEY })).status, 200);
    const chat = await post(base, '/api/chat', {
      provider: 'qwen',
      model: 'qwen3.8-max',
      endpoint: virginia,
      messages: [{ role: 'user', content: 'Hello' }]
    });
    assert.equal(chat.status, 200);
    assert.equal(capturedUrl, virginia);
    assert.notEqual(capturedUrl, SINGAPORE_PAYG_CHAT);
  } finally {
    await stop(relay);
  }
});

test('grok still calls the xAI responses API and ignores the Qwen endpoint', async () => {
  let captured;
  const fetchImpl = async (url, init) => {
    captured = { url: String(url), init };
    return jsonResponse({
      model: 'grok-4.7',
      output: [{ type: 'message', content: [{ type: 'output_text', text: 'grok answer' }] }]
    });
  };
  const { relay, base } = await start(fetchImpl);
  try {
    assert.equal((await post(base, '/api/configure', { provider: 'grok', apiKey: FAKE_GROK_KEY })).status, 200);
    const chat = await post(base, '/api/chat', {
      provider: 'grok',
      model: 'grok-4.7',
      endpoint: 'https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions',
      messages: [{ role: 'user', content: 'Hello' }]
    });
    assert.equal(chat.status, 200);
    assert.equal(chat.data.text, 'grok answer');
    assert.equal(chat.raw.includes(FAKE_GROK_KEY), false);
    assert.equal(captured.url, 'https://api.x.ai/v1/responses');
    const payload = JSON.parse(captured.init.body);
    assert.equal(payload.model, 'grok-4.7');
    assert.equal(payload.store, false);
    assert.deepEqual(payload.input, [{ role: 'user', content: 'Hello' }]);
    assert.equal(captured.init.headers.Authorization, `Bearer ${FAKE_GROK_KEY}`);
  } finally {
    await stop(relay);
  }
});

test('qwen rejects a non-HTTPS endpoint before calling the provider', async () => {
  let called = false;
  const { relay, base } = await start(async () => { called = true; throw new Error('should not fetch'); });
  try {
    assert.equal((await post(base, '/api/configure', { provider: 'qwen', apiKey: FAKE_QWEN_KEY })).status, 200);
    const chat = await post(base, '/api/chat', {
      provider: 'qwen',
      model: 'qwen3.8-max',
      endpoint: 'http://ap-southeast-1.maas.qwencloudapi.com/compatible-mode/v1/chat/completions',
      messages: [{ role: 'user', content: 'Hello' }]
    });
    assert.equal(chat.status, 400);
    assert.equal(called, false);
    assert.equal(chat.raw.includes(FAKE_QWEN_KEY), false);
  } finally {
    await stop(relay);
  }
});

test('clearing a key drops it from memory and does not persist it in the response', async () => {
  const { relay, base } = await start(async () => jsonResponse({ choices: [{ message: { content: 'no' } }] }));
  try {
    assert.equal((await post(base, '/api/configure', { provider: 'qwen', apiKey: FAKE_QWEN_KEY })).status, 200);
    const cleared = await post(base, '/api/clear-key', { provider: 'qwen' });
    assert.equal(cleared.status, 200);
    assert.equal(cleared.raw.includes(FAKE_QWEN_KEY), false);
    const health = await fetch(`${base}/api/health`);
    const healthBody = await health.json();
    assert.equal(healthBody.configured.qwen, false);
    assert.equal(JSON.stringify(healthBody).includes(FAKE_QWEN_KEY), false);
    const chat = await post(base, '/api/chat', {
      provider: 'qwen',
      model: 'qwen3.8-max',
      endpoint: SINGAPORE_PAYG_CHAT,
      messages: [{ role: 'user', content: 'Hello' }]
    });
    assert.equal(chat.status, 401);
  } finally {
    await stop(relay);
  }
});
