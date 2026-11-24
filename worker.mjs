import { RELAY_BUILD, prepareChat, providerDelta, providerErrorMessage, takeSseDataLines } from './relay-core.mjs';

// Marketplace v1 is BYO-key: the caller's own Qwen key travels per request, is forwarded to QwenCloud,
// and is never stored, logged, or substituted with an owner key. Rate limits are per isolate (best effort).
const RATE_WINDOW_MS = 60_000;
const RATE_PER_KEY = 60;
const RATE_GLOBAL = 240;
const rateBuckets = new Map();

function rateLimited(key) {
  const now = Date.now();
  const prune = bucket => { while (bucket.length && now - bucket[0] > RATE_WINDOW_MS) bucket.shift(); };
  const global = rateBuckets.get('*') || [];
  prune(global);
  if (global.length >= RATE_GLOBAL) return true;
  const bucket = rateBuckets.get(key) || [];
  prune(bucket);
  if (bucket.length >= RATE_PER_KEY) return true;
  bucket.push(now);
  global.push(now);
  rateBuckets.set(key, bucket);
  rateBuckets.set('*', global);
  return false;
}

function byoKey(body) {
  const key = typeof body?.apiKey === 'string' ? body.apiKey.trim() : '';
  if (!key || key.length < 8 || key.length > 500 || !key.startsWith('sk-')) return '';
  return key;
}

const JSON_HEADERS = { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' };

function json(status, data) {
  return new Response(JSON.stringify(data), { status, headers: JSON_HEADERS });
}

function sameOrigin(request) {
  const origin = request.headers.get('origin');
  if (!origin) return true;
  try { return new URL(origin).host === new URL(request.url).host; }
  catch { return false; }
}

async function readJson(request) {
  const text = await request.text();
  if (text.length > 2_000_000) {
    const error = new Error('Request too large');
    error.status = 400;
    throw error;
  }
  try { return JSON.parse(text || '{}'); }
  catch {
    const error = new Error('Invalid JSON');
    error.status = 400;
    throw error;
  }
}

async function relaySse(prep, secret) {
  const payload = { ...prep.payload, stream: true };
  const response = await fetch(prep.url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${secret}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
    redirect: 'error',
    signal: AbortSignal.timeout(180_000)
  });
  if (!response.ok || !response.body) {
    const data = await response.json().catch(() => ({}));
    return json(response.ok ? 502 : response.status, { error: response.ok ? 'The provider returned an empty stream.' : providerErrorMessage(data, response.status) });
  }
  const encoder = new TextEncoder();
  const decoder = new TextDecoder();
  const stream = new ReadableStream({
    async start(controller) {
      const send = obj => controller.enqueue(encoder.encode(`data: ${JSON.stringify(obj)}\n\n`));
      let buf = '';
      const apply = flush => {
        let pending = buf;
        if (flush && pending.trim()) pending += '\n';
        const taken = takeSseDataLines(pending);
        buf = taken.rest;
        for (const raw of taken.data) {
          let event;
          try { event = JSON.parse(raw); } catch { continue; }
          const delta = providerDelta(prep.provider, event);
          if (delta.error) throw new Error(delta.error);
          if (delta.model) send({ model: delta.model });
          if (delta.content) send({ choices: [{ delta: { content: delta.content } }] });
        }
      };
      try {
        send({ status: 'thinking' });
        const reader = response.body.getReader();
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          buf += decoder.decode(value, { stream: true });
          apply(false);
        }
        buf += decoder.decode();
        apply(true);
        controller.enqueue(encoder.encode('data: [DONE]\n\n'));
        controller.close();
      } catch (error) {
        const message = error.name === 'TimeoutError' || error.name === 'AbortError' ? 'Provider request timed out.' : error.message;
        try { send({ error: message }); } catch { /* the client is already gone */ }
        try { controller.close(); } catch { /* already closed */ }
      }
    }
  });
  return new Response(stream, { headers: { 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' } });
}

export async function handleRelayRequest(request, env) {
  const url = new URL(request.url);
  if (!url.pathname.startsWith('/api/')) {
    if (!env.ASSETS?.fetch) return new Response('Not found', { status: 404 });
    return env.ASSETS.fetch(request);
  }
  if (!sameOrigin(request)) return json(403, { error: 'Relay only accepts requests from its own site.' });
  if (request.method === 'GET' && url.pathname === '/api/health') {
    return json(200, { ok: true, hosted: true, byo: true, build: RELAY_BUILD, configured: { qwen: false } });
  }
  if (request.method === 'POST' && (url.pathname === '/api/configure' || url.pathname === '/api/clear-key')) {
    return json(400, { error: 'On the hosted pane your Qwen key stays in this browser and is sent per request; the server never stores it.' });
  }
  if (request.method === 'POST' && (url.pathname === '/api/chat' || url.pathname === '/api/chat-stream')) {
    try {
      const body = await readJson(request);
      const key = byoKey(body);
      if (!key) return json(401, { error: 'Add your Qwen key in Settings. The hosted pane sends it per request and never stores it server-side.' });
      if (rateLimited(key)) return json(429, { error: 'Rate limit: 60 requests per minute per key. Wait a minute and retry.' });
      const secrets = { qwen: key };
      const prep = prepareChat(body, secrets);
      if (prep.error) return json(prep.status, { error: prep.error });
      if (url.pathname === '/api/chat-stream') return relaySse(prep, key);
      const response = await fetch(prep.url, {
        method: 'POST',
        headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(prep.payload),
        redirect: 'error',
        signal: AbortSignal.timeout(180_000)
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) return json(response.status, { error: providerErrorMessage(data, response.status) });
      const text = data?.choices?.[0]?.message?.content;
      if (typeof text !== 'string' || !text.trim()) return json(502, { error: 'The provider returned no text answer.' });
      return json(200, { text, model: data.model || prep.model });
    } catch (error) {
      const message = error.name === 'TimeoutError' ? 'Provider request timed out.' : error.message;
      return json(error.status || 502, { error: message });
    }
  }
  return json(405, { error: 'Method not allowed.' });
}

export default {
  fetch(request, env) { return handleRelayRequest(request, env); }
};
