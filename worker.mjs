import { RELAY_BUILD, prepareChat, providerDelta, providerErrorMessage, relayTokenOk, takeSseDataLines } from './relay-core.mjs';

const JSON_HEADERS = { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' };

function json(status, data) {
  return new Response(JSON.stringify(data), { status, headers: JSON_HEADERS });
}

function bearer(request) {
  const header = request.headers.get('authorization') || '';
  const match = header.match(/^Bearer\s+(\S+)\s*$/i);
  return match ? match[1] : '';
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
  if (!relayTokenOk(bearer(request), env.RELAY_TOKEN || '')) return json(401, { error: 'Enter the Cloudflare relay token in Settings.' });
  const secrets = { qwen: env.QWEN_API_KEY || '' };
  if (request.method === 'GET' && url.pathname === '/api/health') {
    return json(200, { ok: true, hosted: true, build: RELAY_BUILD, configured: { qwen: Boolean(secrets.qwen) } });
  }
  if (request.method === 'POST' && (url.pathname === '/api/configure' || url.pathname === '/api/clear-key')) {
    return json(400, { error: 'On Cloudflare, the Qwen key is a Worker secret. This pane stores only the relay token.' });
  }
  if (request.method === 'POST' && (url.pathname === '/api/chat' || url.pathname === '/api/chat-stream')) {
    try {
      const body = await readJson(request);
      const prep = prepareChat(body, secrets);
      if (prep.error) return json(prep.status, { error: prep.error });
      if (url.pathname === '/api/chat-stream') return relaySse(prep, secrets[prep.provider]);
      const response = await fetch(prep.url, {
        method: 'POST',
        headers: { Authorization: `Bearer ${secrets[prep.provider]}`, 'Content-Type': 'application/json' },
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
