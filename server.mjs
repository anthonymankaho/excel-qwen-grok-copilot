import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { extractGrokText, originAllowed, prepareChat, providerDelta, providerErrorMessage, requestHostAllowed, takeSseDataLines } from './relay-core.mjs';

export { extractGrokText, originAllowed, prepareChat, providerDelta, providerErrorMessage, qwenEndpointAllowed, requestHostAllowed, takeSseDataLines } from './relay-core.mjs';

const root = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(root, 'public');
const host = '127.0.0.1';
const port = Number(process.env.PORT || 3000);
const BODY_LIMIT = 2_000_000;
const mime = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.xml': 'application/xml; charset=utf-8', '.png': 'image/png', '.json': 'application/json; charset=utf-8' };

function send(res, status, data, type = 'application/json; charset=utf-8') {
  res.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
  const body = Buffer.isBuffer(data) ? data : typeof data === 'string' ? data : JSON.stringify(data);
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    let settled = false;
    const fail = error => { if (settled) return; settled = true; reject(error); };
    req.on('data', part => {
      body += part;
      if (body.length > BODY_LIMIT) { req.destroy(); fail(new Error('Request too large')); }
    });
    req.on('end', () => {
      if (settled) return;
      settled = true;
      try { resolve(JSON.parse(body || '{}')); }
      catch { reject(new Error('Invalid JSON')); }
    });
    req.on('error', fail);
  });
}


async function readProviderStream(response, provider, res) {
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buf = '';
  const emit = (payload) => res.write(`data: ${JSON.stringify(payload)}\n\n`);
  const applyBuf = (flushTail) => {
    let pending = buf;
    if (flushTail && pending.trim()) pending += '\n';
    const taken = takeSseDataLines(pending);
    buf = taken.rest;
    for (const raw of taken.data) {
      let event;
      try { event = JSON.parse(raw); } catch { continue; }
      const delta = providerDelta(provider, event);
      if (delta.error) throw new Error(delta.error);
      if (delta.model) emit({ model: delta.model });
      if (delta.content) emit({ choices: [{ delta: { content: delta.content } }] });
    }
  };
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    applyBuf(false);
  }
  buf += decoder.decode();
  applyBuf(true);
}

export function createRelay() {
  const secrets = { qwen: '', grok: '' };
  const server = http.createServer(async (req, res) => {
    const addr = server.address();
    const boundPort = addr && typeof addr === 'object' ? addr.port : port;
    if (!requestHostAllowed(req.headers.host, boundPort) || !originAllowed(req.headers.origin, boundPort)) {
      return send(res, 403, { error: 'Relay only accepts local requests.' });
    }
    let url;
    try { url = new URL(req.url, `http://${host}:${boundPort}`); }
    catch { return send(res, 400, { error: 'Bad request.' }); }

    if (req.method === 'GET' && url.pathname === '/api/health') return send(res, 200, { ok: true, configured: { qwen: Boolean(secrets.qwen), grok: Boolean(secrets.grok) } });

    if (req.method === 'POST' && url.pathname === '/api/configure') {
      try {
        const body = await readBody(req);
        if (!['qwen', 'grok'].includes(body.provider) || typeof body.apiKey !== 'string' || body.apiKey.trim().length < 8 || body.apiKey.length > 500) return send(res, 400, { error: 'Enter a valid API key.' });
        secrets[body.provider] = body.apiKey.trim();
        return send(res, 200, { ok: true, configured: { qwen: Boolean(secrets.qwen), grok: Boolean(secrets.grok) } });
      } catch (error) { return send(res, 400, { error: error.message }); }
    }

    if (req.method === 'POST' && url.pathname === '/api/clear-key') {
      try {
        const body = await readBody(req);
        if (!['qwen', 'grok'].includes(body.provider)) return send(res, 400, { error: 'Unknown provider.' });
        secrets[body.provider] = '';
        return send(res, 200, { ok: true });
      } catch (error) { return send(res, 400, { error: error.message }); }
    }

    if (req.method === 'POST' && url.pathname === '/api/chat') {
      try {
        const body = await readBody(req);
        const prep = prepareChat(body, secrets);
        if (prep.error) return send(res, prep.status, { error: prep.error });
        const { provider, model, url: target, payload } = prep;
        const response = await fetch(target, { method: 'POST', headers: { Authorization: `Bearer ${secrets[provider]}`, 'Content-Type': 'application/json' }, body: JSON.stringify(payload), redirect: 'error', signal: AbortSignal.timeout(180_000) });
        const data = await response.json().catch(() => ({}));
        if (!response.ok) return send(res, response.status, { error: providerErrorMessage(data, response.status) });
        const text = provider === 'qwen' ? data?.choices?.[0]?.message?.content : extractGrokText(data);
        if (typeof text !== 'string' || !text.trim()) return send(res, 502, { error: 'The provider returned no text answer.' });
        return send(res, 200, { text, model: data.model || model });
      } catch (error) {
        const message = error.name === 'TimeoutError' ? 'Provider request timed out.' : error.message;
        const status = message === 'Invalid JSON' || message === 'Request too large' ? 400 : 502;
        return send(res, status, { error: message });
      }
    }

    if (req.method === 'POST' && url.pathname === '/api/chat-stream') {
      const ac = new AbortController();
      const timer = setTimeout(() => ac.abort(), 240_000);
      let settled = false;
      let clientGone = false;
      // res 'close' = client gone mid-stream; req 'close' fires as soon as the body is consumed and would abort every call.
      res.on('close', () => { clientGone = true; clearTimeout(timer); if (!settled) ac.abort(); });
      try {
        const body = await readBody(req);
        const prep = prepareChat(body, secrets);
        if (prep.error) { settled = true; clearTimeout(timer); return send(res, prep.status, { error: prep.error }); }
        const { provider, url: target, payload } = prep;
        payload.stream = true;
        const response = await fetch(target, { method: 'POST', headers: { Authorization: `Bearer ${secrets[provider]}`, 'Content-Type': 'application/json' }, body: JSON.stringify(payload), redirect: 'error', signal: ac.signal });
        if (!response.ok || !response.body) {
          const data = await response.json().catch(() => ({}));
          settled = true;
          clearTimeout(timer);
          return send(res, response.ok ? 502 : response.status, { error: response.ok ? 'The provider returned an empty stream.' : providerErrorMessage(data, response.status) });
        }
        res.writeHead(200, { 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-store', 'Connection': 'keep-alive', 'X-Content-Type-Options': 'nosniff' });
        res.write(`data: ${JSON.stringify({ status: 'thinking' })}\n\n`);
        await readProviderStream(response, provider, res);
        if (clientGone) return;
        res.write('data: [DONE]\n\n');
        settled = true;
        clearTimeout(timer);
        return res.end();
      } catch (error) {
        settled = true;
        clearTimeout(timer);
        if (clientGone || res.destroyed) return;
        const message = error.name === 'AbortError' ? 'Provider request timed out.' : error.message;
        if (!res.headersSent && (message === 'Invalid JSON' || message === 'Request too large')) return send(res, 400, { error: message });
        if (res.headersSent) {
          try { res.write(`data: ${JSON.stringify({ error: message })}\n\n`); res.end(); } catch { /* client already gone */ }
          return;
        }
        return send(res, 502, { error: message });
      }
    }

    if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, { error: 'Method not allowed.' });
    let relative;
    try { relative = url.pathname === '/' ? 'index.html' : decodeURIComponent(url.pathname.replace(/^\/+/, '')); }
    catch { return send(res, 400, 'Bad request', 'text/plain; charset=utf-8'); }
    const file = path.resolve(publicDir, relative);
    const fromPublic = path.relative(publicDir, file);
    if (fromPublic.startsWith('..') || path.isAbsolute(fromPublic)) return send(res, 403, 'Forbidden', 'text/plain; charset=utf-8');
    const type = mime[path.extname(file)] || 'application/octet-stream';
    fs.readFile(file, (error, contents) => {
      if (error) return send(res, 404, 'Not found', 'text/plain; charset=utf-8');
      if (req.method === 'HEAD') {
        res.writeHead(200, { 'Content-Type': type, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
        return res.end();
      }
      return send(res, 200, contents, type);
    });
  });
  return { server, secrets };
}

function isDirectRun() {
  const entry = process.argv[1];
  if (!entry) return false;
  return path.resolve(entry) === path.resolve(fileURLToPath(import.meta.url));
}

if (isDirectRun()) {
  const { server } = createRelay();
  server.on('error', error => {
    if (error.code === 'EADDRINUSE') console.error(`Port ${port} is already in use. Stop the other relay first (npm run relay-stop).`);
    else console.error(error);
    process.exitCode = 1;
  });
  server.listen(port, host, () => console.log(`Excel Qwen/Grok Copilot ready at http://${host}:${port}`));
}
