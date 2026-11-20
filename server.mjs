import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(root, 'public');
const host = '127.0.0.1';
const port = Number(process.env.PORT || 3000);
const secrets = { qwen: '', grok: '' }; // Process memory only; cleared when the relay stops.
const SYSTEM_PROMPT = 'You are an assistant inside Microsoft Excel. Be accurate, concise, and clear. When analyzing spreadsheet data, state assumptions and distinguish values from formulas. The user message may include an [Excel selection ...] or [Excel used range ...] block: tab-separated rows where formulas appear as value [=formula]. When the user asks to change, fix, or create cell content, reply with your explanation plus one fenced ```json block containing {"edits":[{"sheet":"<sheet name>","cell":"<A1 address>","value":"<number, text, or =formula>"}]} (at most 200 edits). You cannot change cells yourself; edits apply only after the user clicks Apply in the pane, so never claim a change is already done.';
const mime = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.xml': 'application/xml; charset=utf-8', '.png': 'image/png', '.json': 'application/json; charset=utf-8' };

function send(res, status, data, type = 'application/json; charset=utf-8') {
  res.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
  const body = Buffer.isBuffer(data) ? data : typeof data === 'string' ? data : JSON.stringify(data);
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', part => { body += part; if (body.length > 1_000_000) reject(new Error('Request too large')); });
    req.on('end', () => { try { resolve(JSON.parse(body || '{}')); } catch { reject(new Error('Invalid JSON')); } });
    req.on('error', reject);
  });
}

function extractGrokText(data) {
  for (const item of data.output || []) {
    if (item.type === 'message') {
      const text = (item.content || []).filter(part => part.type === 'output_text').map(part => part.text).join('');
      if (text) return text;
    }
  }
  return '';
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${host}:${port}`);
  if (req.method === 'GET' && url.pathname === '/api/health') return send(res, 200, { ok: true, configured: { qwen: Boolean(secrets.qwen), grok: Boolean(secrets.grok) } });

  if (req.method === 'POST' && url.pathname === '/api/configure') {
    try {
      const body = await readBody(req);
      if (!['qwen', 'grok'].includes(body.provider) || typeof body.apiKey !== 'string' || body.apiKey.length < 8 || body.apiKey.length > 500) return send(res, 400, { error: 'Enter a valid API key.' });
      secrets[body.provider] = body.apiKey.trim();
      return send(res, 200, { ok: true, configured: { qwen: Boolean(secrets.qwen), grok: Boolean(secrets.grok) } });
    } catch (error) { return send(res, 400, { error: error.message }); }
  }

  if (req.method === 'POST' && url.pathname === '/api/clear-key') {
    try { const body = await readBody(req); if (!['qwen', 'grok'].includes(body.provider)) return send(res, 400, { error: 'Unknown provider.' }); secrets[body.provider] = ''; return send(res, 200, { ok: true }); }
    catch (error) { return send(res, 400, { error: error.message }); }
  }

  if (req.method === 'POST' && url.pathname === '/api/chat') {
    try {
      const body = await readBody(req);
      const { provider, model, endpoint, messages } = body;
      if (!['qwen', 'grok'].includes(provider)) return send(res, 400, { error: 'Choose Qwen or Grok.' });
      if (!secrets[provider]) return send(res, 401, { error: `Add your ${provider === 'qwen' ? 'Qwen' : 'Grok'} API key in Settings first.` });
      if (!Array.isArray(messages) || messages.length < 1 || messages.length > 24) return send(res, 400, { error: 'Conversation must contain 1 to 24 messages.' });
      if (messages.some(m => !['user', 'assistant'].includes(m.role) || typeof m.content !== 'string' || m.content.length > 50_000)) return send(res, 400, { error: 'Invalid conversation message.' });

      let url, payload;
      if (provider === 'qwen') {
        url = new URL(endpoint);
        if (url.protocol !== 'https:' || url.username || url.password) return send(res, 400, { error: 'Qwen endpoint must be an HTTPS URL.' });
        payload = { model: model || 'qwen4-max', messages: [{ role: 'system', content: SYSTEM_PROMPT }, ...messages], stream: false };
      } else {
        url = new URL('https://api.x.ai/v1/responses');
        payload = { model: model || 'grok-4.7', instructions: SYSTEM_PROMPT, input: messages.map(({ role, content }) => ({ role, content })), store: false };
      }
      const response = await fetch(url, { method: 'POST', headers: { Authorization: `Bearer ${secrets[provider]}`, 'Content-Type': 'application/json' }, body: JSON.stringify(payload), signal: AbortSignal.timeout(90_000) });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) return send(res, response.status, { error: data?.error?.message || data?.message || `Provider returned HTTP ${response.status}.` });
      const text = provider === 'qwen' ? data?.choices?.[0]?.message?.content : extractGrokText(data);
      if (typeof text !== 'string' || !text.trim()) return send(res, 502, { error: 'The provider returned no text answer.' });
      return send(res, 200, { text, model: data.model || model });
    } catch (error) {
      const message = error.name === 'TimeoutError' ? 'Provider request timed out.' : error.message;
      return send(res, 502, { error: message });
    }
  }

  const relative = url.pathname === '/' ? 'index.html' : decodeURIComponent(url.pathname.slice(1));
  const file = path.resolve(publicDir, relative);
  if (!file.startsWith(publicDir + path.sep)) return send(res, 403, 'Forbidden', 'text/plain; charset=utf-8');
  fs.readFile(file, (error, contents) => error ? send(res, 404, 'Not found', 'text/plain; charset=utf-8') : send(res, 200, contents, mime[path.extname(file)] || 'application/octet-stream'));
});

server.listen(port, host, () => console.log(`Excel Qwen/Grok Copilot ready at http://${host}:${port}`));
