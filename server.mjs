import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(root, 'public');
const host = '127.0.0.1';
const port = Number(process.env.PORT || 3000);
const BODY_LIMIT = 2_000_000;
const QWEN_PATH = '/compatible-mode/v1/chat/completions';
const MODEL_RE = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,79}$/;
const SYSTEM_PROMPT = `You are an assistant inside Microsoft Office, revising the open Excel workbook or Word document the way Copilot and Claude for Excel do. Be accurate and concise. Distinguish values from formulas. The user message may include an [Excel selection ...], [Excel used range ...], [Word selection], or [Word document] block.

When the user asks to change the file, reply with a short explanation plus one fenced json block. Do not claim the change is already saved; the pane applies the actions. Prefer {"actions":[...]} with at most 40 actions. A legacy {"edits":[{"sheet":"Sheet1","cell":"D5","value":"=SUM(A1:A4)"}]} block is still accepted for single cells.

Excel actions:
- {"op":"set","sheet":"Sheet1","range":"A1:B2","values":[["Name","Total"],["Ada",100]]} — a formula is a string that starts with =. A single cell may use "value" instead of "values".
- {"op":"format","sheet":"Sheet1","range":"A1:B1","bold":true,"italic":false,"fill":"#1F4E79","fontColor":"#FFFFFF","numberFormat":"#,##0","align":"center","wrap":true}
- {"op":"clear","sheet":"Sheet1","range":"C2:C20"}
- {"op":"insertRows","sheet":"Sheet1","row":5,"count":1}
- {"op":"insertColumns","sheet":"Sheet1","column":"C","count":1}
- {"op":"deleteRows","sheet":"Sheet1","row":5,"count":1}
- {"op":"addSheet","name":"Summary"}
- {"op":"renameSheet","sheet":"Sheet1","name":"Data"}
- {"op":"autofit","sheet":"Sheet1","range":"A:D"}
- {"op":"sort","sheet":"Sheet1","range":"A1:B20","column":"B","ascending":false,"hasHeaders":true}
- {"op":"chart","sheet":"Sheet1","range":"A1:B20","type":"column","title":"Sales"} — type is column, bar, line, or pie
- {"op":"table","sheet":"Sheet1","range":"A1:B20","name":"SalesTable"}

Word actions:
- {"op":"replaceSelection","text":"Revised paragraph."}
- {"op":"insertText","text":"Added at the end.","where":"end"} — where is start or end
- {"op":"formatSelection","bold":true,"italic":false}
- {"op":"styleSelection","style":"Heading 1"} — Heading 1, Heading 2, Heading 3, Title, or Normal

Use only actions for the open host. Do not propose WEBSERVICE, FILTERXML, or RTD unless the user explicitly asks. Do not invent a sheet or range you have not seen, except addSheet.`;
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

export function qwenEndpointAllowed(endpoint) {
  if (typeof endpoint !== 'string' || !endpoint.trim()) return null;
  let url;
  try { url = new URL(endpoint.trim()); } catch { return null; }
  if (url.protocol !== 'https:' || url.username || url.password) return null;
  const hostname = url.hostname.toLowerCase().replace(/\.$/, '');
  const trusted = hostname === 'qwencloudapi.com' || hostname.endsWith('.qwencloudapi.com') || hostname === 'aliyuncs.com' || hostname.endsWith('.aliyuncs.com');
  if (!trusted) return null;
  const pathname = url.pathname.replace(/\/+$/, '') || '/';
  if (pathname !== QWEN_PATH) return null;
  return url;
}

export function requestHostAllowed(header, boundPort) {
  if (typeof header !== 'string' || !header.trim()) return false;
  const value = header.trim().toLowerCase();
  const suffix = `:${boundPort}`;
  return value === `127.0.0.1${suffix}` || value === `localhost${suffix}`;
}

export function originAllowed(origin, boundPort) {
  if (origin == null || origin === '') return true;
  return origin === `http://127.0.0.1:${boundPort}` || origin === `http://localhost:${boundPort}`;
}

export function takeSseDataLines(buffer) {
  const data = [];
  let rest = buffer;
  let idx;
  while ((idx = rest.indexOf('\n')) >= 0) {
    const line = rest.slice(0, idx).replace(/\r$/, '').trim();
    rest = rest.slice(idx + 1);
    if (!line.startsWith('data:')) continue;
    const raw = line.slice(5).trim();
    if (!raw || raw === '[DONE]') continue;
    data.push(raw);
  }
  return { data, rest };
}

export function providerDelta(provider, event) {
  if (!event || typeof event !== 'object') return { content: '', model: '', error: '' };
  if (provider === 'qwen') {
    if (event.error) {
      const message = event.error.message || event.error;
      return { content: '', model: '', error: typeof message === 'string' ? message : 'The provider failed while streaming.' };
    }
    const delta = event.choices?.[0]?.delta?.content;
    return { content: typeof delta === 'string' ? delta : '', model: typeof event.model === 'string' ? event.model : '', error: '' };
  }
  if (event.type === 'response.output_text.delta' && typeof event.delta === 'string') return { content: event.delta, model: '', error: '' };
  if (event.type === 'response.completed') return { content: '', model: typeof event.response?.model === 'string' ? event.response.model : '', error: '' };
  if (event.type === 'response.incomplete') {
    const reason = event.response?.incomplete_details?.reason || 'incomplete';
    return { content: '', model: typeof event.response?.model === 'string' ? event.response.model : '', error: `The provider stopped early (${reason}).` };
  }
  if (event.type === 'response.failed' || event.type === 'error') {
    const message = event.error?.message || event.response?.error?.message || event.message || 'The provider failed while streaming.';
    return { content: '', model: '', error: typeof message === 'string' ? message : 'The provider failed while streaming.' };
  }
  return { content: '', model: '', error: '' };
}

function hostNote(host) {
  return host === 'word'
    ? '\nThe open file is a Word document. Use only Word actions.'
    : '\nThe open file is an Excel workbook. Use only Excel actions.';
}

function resolveModel(model, fallback) {
  if (model == null || (typeof model === 'string' && model.trim() === '')) return { model: fallback };
  if (typeof model !== 'string' || !MODEL_RE.test(model.trim())) return { error: 'Invalid model name.' };
  return { model: model.trim() };
}

export function prepareChat(body, secrets) {
  const provider = body?.provider;
  const messages = body?.messages;
  if (!['qwen', 'grok'].includes(provider)) return { status: 400, error: 'Choose Qwen or Grok.' };
  if (!secrets[provider]) return { status: 401, error: `Add your ${provider === 'qwen' ? 'Qwen' : 'Grok'} API key in Settings first.` };
  if (!Array.isArray(messages) || messages.length < 1 || messages.length > 24) return { status: 400, error: 'Conversation must contain 1 to 24 messages.' };
  if (messages.some(m => !m || !['user', 'assistant'].includes(m.role) || typeof m.content !== 'string' || m.content.length > 50_000)) return { status: 400, error: 'Invalid conversation message.' };
  let url;
  let payload;
  if (provider === 'qwen') {
    url = qwenEndpointAllowed(body.endpoint);
    if (!url) return { status: 400, error: 'Qwen endpoint must be an HTTPS QwenCloud chat completions URL.' };
    const model = resolveModel(body.model, 'qwen4-max');
    if (model.error) return { status: 400, error: model.error };
    payload = { model: model.model, messages: [{ role: 'system', content: SYSTEM_PROMPT + hostNote(body.host) }, ...messages], stream: false };
    if (body.fast) payload.enable_thinking = false;
  } else {
    url = new URL('https://api.x.ai/v1/responses');
    const model = resolveModel(body.model, 'grok-4.7');
    if (model.error) return { status: 400, error: model.error };
    payload = { model: model.model, instructions: SYSTEM_PROMPT + hostNote(body.host), input: messages.map(({ role, content }) => ({ role, content })), store: false };
    if (body.fast) payload.reasoning = { effort: 'low' };
  }
  return { provider, model: payload.model, url, payload };
}

function extractGrokText(data) {
  if (typeof data?.output_text === 'string' && data.output_text.trim()) return data.output_text;
  for (const item of data?.output || []) {
    if (item.type === 'message') {
      const text = (item.content || []).filter(part => part.type === 'output_text').map(part => part.text).join('');
      if (text) return text;
    }
  }
  return '';
}

function providerErrorMessage(data, status) {
  return data?.error?.message || data?.message || `Provider returned HTTP ${status}.`;
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
