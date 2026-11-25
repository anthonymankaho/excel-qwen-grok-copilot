import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { extractDocx, extractPdf, extractWorkbook } from './extract-files.mjs';
import { RELAY_BUILD, appendMemory, buildSkillContext, deleteSession, fetchGitHubSkill, installSkillFile, listSessions, listSkills, originAllowed, parseGitHubSkillSource, prepareChat, providerDelta, providerErrorMessage, qwenEndpointAllowed, readMemory, readSession, requestHostAllowed, resolveSkillsRoot, saveSkill, takeSseDataLines, writeSession } from './relay-core.mjs';
import { LocalSourceMcp } from './mcp-connectors.mjs';

export { originAllowed, prepareChat, providerDelta, providerErrorMessage, qwenEndpointAllowed, requestHostAllowed, takeSseDataLines } from './relay-core.mjs';

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

// Scanned PDFs/images: transcribe with the user's OWN Qwen key (vision). Images leave this PC only to QwenCloud.
async function visionTranscribe(content, apiKey, endpoint) {
  const response = await fetch(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({ model: 'qwen4-max', messages: [{ role: 'user', content }], stream: false, enable_thinking: false }),
    signal: AbortSignal.timeout(180000)
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data?.error?.message || `Vision transcription failed (${response.status}).`);
  const text = data?.choices?.[0]?.message?.content;
  if (typeof text !== 'string' || !text.trim()) throw new Error('Vision transcription returned no text.');
  return text;
}

async function transcribePdfPage(png, pageNumber, apiKey, endpoint) {
  const content = [
    { type: 'text', text: `Transcribe page ${pageNumber} verbatim. Preserve tables as pipe-separated rows and numbers exactly. Output only this page.` },
    { type: 'image_url', image_url: { url: `data:image/png;base64,${png.toString('base64')}` } }
  ];
  return `[Vision transcription of page ${pageNumber} via your Qwen key]\n${await visionTranscribe(content, apiKey, endpoint)}`;
}

async function transcribeImage(buffer, mime, apiKey, endpoint) {
  const content = [
    { type: 'text', text: 'Transcribe this image verbatim, preserving tables as pipe-separated rows and numbers exactly. Output only the transcription.' },
    { type: 'image_url', image_url: { url: `data:${mime};base64,${buffer.toString('base64')}` } }
  ];
  return `[Vision transcription of an image via your Qwen key]\n${await visionTranscribe(content, apiKey, endpoint)}`;
}

function cloudOAuthConfig() {
  try {
    const config = JSON.parse(fs.readFileSync(path.join(root, 'cloud-oauth.json'), 'utf8'));
    return {
      microsoftClientId: typeof config.microsoftClientId === 'string' ? config.microsoftClientId : ''
    };
  } catch { return {}; }
}

function readBody(req, limit = BODY_LIMIT) {
  return new Promise((resolve, reject) => {
    let body = '';
    let settled = false;
    const fail = error => { if (settled) return; settled = true; reject(error); };
    req.on('data', part => {
      body += part;
      if (body.length > limit) { req.destroy(); fail(new Error('Request too large')); }
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
  const secrets = { qwen: '' };
  const sources = new LocalSourceMcp(cloudOAuthConfig());
  const server = http.createServer(async (req, res) => {
    const addr = server.address();
    const boundPort = addr && typeof addr === 'object' ? addr.port : port;
    if (!requestHostAllowed(req.headers.host, boundPort) || !originAllowed(req.headers.origin, boundPort)) {
      return send(res, 403, { error: 'Relay only accepts local requests.' });
    }
    let url;
    try { url = new URL(req.url, `http://${host}:${boundPort}`); }
    catch { return send(res, 400, { error: 'Bad request.' }); }

    if (req.method === 'GET' && url.pathname === '/api/health') return send(res, 200, { ok: true, build: RELAY_BUILD, configured: { qwen: Boolean(secrets.qwen) } });

    if (req.method === 'GET' && url.pathname === '/api/connectors/status') return send(res, 200, sources.status());

    if (req.method === 'POST' && url.pathname === '/api/connectors/onedrive/authorize') {
      try {
        await readBody(req);
        return send(res, 200, await sources.startOneDriveInteractiveFlow());
      } catch (error) {
        return send(res, 400, { error: error.message });
      }
    }

    if (req.method === 'POST' && url.pathname === '/api/connectors/search') {
      try {
        const body = await readBody(req);
        return send(res, 200, await sources.search(body.source, body.query));
      } catch (error) {
        const clientError = /query|choose|valid search result/i.test(error.message);
        return send(res, clientError ? 400 : 502, { error: error.message });
      }
    }

    if (req.method === 'POST' && url.pathname === '/api/connectors/read') {
      try {
        const body = await readBody(req);
        const endpoint = qwenEndpointAllowed(body.endpoint);
        const transcribePage = secrets.qwen && endpoint
          ? (png, page) => transcribePdfPage(png, page, secrets.qwen, endpoint)
          : null;
        return send(res, 200, await sources.read(body.source, body.item, { transcribePage }));
      } catch (error) {
        const clientError = /choose|valid|file name/i.test(error.message);
        return send(res, clientError ? 400 : 502, { error: error.message });
      }
    }

    if (req.method === 'POST' && url.pathname === '/api/connectors/disconnect') {
      try {
        const body = await readBody(req);
        await sources.disconnect(body.source);
        return send(res, 200, sources.status());
      } catch (error) {
        return send(res, 400, { error: error.message });
      }
    }

    if (req.method === 'GET' && url.pathname === '/api/skills') {
      const root = resolveSkillsRoot(process.env.SKILLS_ROOT || '');
      return send(res, 200, { root, skills: listSkills(root) });
    }

    if (req.method === 'GET' && url.pathname === '/api/skills/context') {
      const root = resolveSkillsRoot(process.env.SKILLS_ROOT || '');
      return send(res, 200, buildSkillContext(url.searchParams.get('question') || '', root));
    }

    if (req.method === 'POST' && url.pathname === '/api/skills/github') {
      try {
        const body = await readBody(req);
        const parsed = parseGitHubSkillSource(body.repo || body.url);
        if (parsed.error) return send(res, 400, { error: parsed.error });
        const token = body.token == null || body.token === '' ? '' : body.token;
        if (token !== '' && (typeof token !== 'string' || token.length > 300 || /\s/.test(token))) return send(res, 400, { error: 'Invalid GitHub token.' });
        const fetched = await fetchGitHubSkill({ ...parsed, token });
        if (fetched.error) return send(res, fetched.status || 400, { error: fetched.error });
        const root = resolveSkillsRoot(process.env.SKILLS_ROOT || '');
        const result = installSkillFile(root, fetched.text, { replace: body.replace === true });
        if (result.error) return send(res, 400, { error: result.error });
        return send(res, 200, { ok: true, id: result.id, path: result.path, replaced: result.replaced });
      } catch (error) { return send(res, 400, { error: error.message }); }
    }

    if (req.method === 'POST' && url.pathname === '/api/skills/save') {
      try {
        const body = await readBody(req);
        const root = resolveSkillsRoot(process.env.SKILLS_ROOT || '');
        const result = saveSkill(root, body);
        if (result.error) return send(res, 400, { error: result.error });
        return send(res, 200, { ok: true, path: result.path });
      } catch (error) { return send(res, 400, { error: error.message }); }
    }

    if (req.method === 'GET' && url.pathname === '/api/sessions') return send(res, 200, { sessions: listSessions() });

    if (req.method === 'GET' && url.pathname === '/api/sessions/read') {
      const session = readSession(url.searchParams.get('id'));
      if (!session) return send(res, 404, { error: 'Session not found.' });
      return send(res, 200, session);
    }

    if (req.method === 'POST' && url.pathname === '/api/sessions/save') {
      try {
        const result = writeSession(await readBody(req));
        if (result.error) return send(res, 400, { error: result.error });
        return send(res, 200, { ok: true });
      } catch (error) { return send(res, 400, { error: error.message }); }
    }

    if (req.method === 'POST' && url.pathname === '/api/sessions/delete') {
      try {
        const result = deleteSession((await readBody(req)).id);
        if (result.error) return send(res, 400, { error: result.error });
        return send(res, 200, { ok: true });
      } catch (error) { return send(res, 400, { error: error.message }); }
    }

    if (req.method === 'POST' && url.pathname === '/api/extract') {
      try {
        const body = await readBody(req, 30_000_000);
        const name = String(body.name || '');
        const ext = name.split('.').pop().toLowerCase();
        const buffer = Buffer.from(String(body.data || ''), 'base64');
        if (!buffer.length) return send(res, 400, { error: 'Empty file.' });
        const endpoint = qwenEndpointAllowed(body.endpoint) || 'https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions';
        const vision = secrets.qwen ? { apiKey: secrets.qwen, endpoint } : null;
        let text = '';
        const transcribePage = vision ? (png, page) => transcribePdfPage(png, page, vision.apiKey, vision.endpoint) : null;
        if (ext === 'xlsx' || ext === 'xlsm') text = (await extractWorkbook(buffer)).text;
        else if (ext === 'docx') text = (await extractDocx(buffer)).text;
        else if (ext === 'pdf') text = (await extractPdf(buffer, { transcribePage })).text;
        else if (['png', 'jpg', 'jpeg', 'webp'].includes(ext)) {
          if (!vision) return send(res, 400, { error: 'Image transcription needs your Qwen key in Settings.' });
          text = await transcribeImage(buffer, ext === 'jpg' ? 'image/jpeg' : `image/${ext}`, vision.apiKey, vision.endpoint);
        }
        else return send(res, 400, { error: 'Extraction supports PDF, DOCX, XLSX, XLSM, PNG, JPG and WEBP.' });
        if (!text.trim()) return send(res, 400, { error: `${name}: no readable text was found.` });
        return send(res, 200, { text: text.slice(0, 200000), kind: ext.toUpperCase() });
      } catch (error) { return send(res, 400, { error: error.message }); }
    }

    if (req.method === 'GET' && url.pathname === '/api/memory') return send(res, 200, { text: readMemory() });

    if (req.method === 'POST' && url.pathname === '/api/memory/add') {
      try {
        const result = appendMemory((await readBody(req)).note);
        if (result.error) return send(res, 400, { error: result.error });
        return send(res, 200, { ok: true });
      } catch (error) { return send(res, 400, { error: error.message }); }
    }

    if (req.method === 'POST' && url.pathname === '/api/configure') {
      try {
        const body = await readBody(req);
        if (body.provider !== 'qwen' || typeof body.apiKey !== 'string' || body.apiKey.trim().length < 8 || body.apiKey.length > 500) return send(res, 400, { error: 'Enter a valid API key.' });
        secrets[body.provider] = body.apiKey.trim();
        return send(res, 200, { ok: true, configured: { qwen: Boolean(secrets.qwen) } });
      } catch (error) { return send(res, 400, { error: error.message }); }
    }

    if (req.method === 'POST' && url.pathname === '/api/clear-key') {
      try {
        const body = await readBody(req);
        if (body.provider !== 'qwen') return send(res, 400, { error: 'Unknown provider.' });
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
        const text = data?.choices?.[0]?.message?.content;
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
  server.listen(port, host, () => console.log(`Excel Qwen Copilot ready at http://${host}:${port}`));
}
