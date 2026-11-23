import { randomBytes, createHash } from 'node:crypto';
import http from 'node:http';
import { spawn } from 'node:child_process';
import { Client, InMemoryTransport, StreamableHTTPClientTransport } from '@modelcontextprotocol/client';
import { McpServer } from '@modelcontextprotocol/server';
import { PublicClientApplication } from '@azure/msal-node';
import ExcelJS from 'exceljs';
import mammoth from 'mammoth';
import { z } from 'zod';

const GOOGLE_DRIVE_URL = 'https://drivemcp.googleapis.com/mcp/v1';
const GOOGLE_GMAIL_URL = 'https://gmailmcp.googleapis.com/mcp/v1';
const GOOGLE_SCOPES = [
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/gmail.readonly'
];
const GRAPH_SCOPES = ['Files.Read'];
const DRIVE_DOWNLOAD_HOSTS = ['.1drv.com', '.1drv.ms', '.sharepoint.com'];
const MAX_SOURCE_BYTES = 12 * 1024 * 1024;
const MAX_SOURCE_CHARS = 70_000;

function textResult(value) {
  return { content: [{ type: 'text', text: JSON.stringify(value) }] };
}

function parseToolResult(result) {
  if (result?.isError) {
    const message = result.content?.filter(part => part.type === 'text').map(part => part.text).join('\n');
    throw new Error(message || 'The source connector returned an error.');
  }
  if (result?.structuredContent && typeof result.structuredContent === 'object') return result.structuredContent;
  const raw = result?.content?.filter(part => part.type === 'text').map(part => part.text).join('\n') || '';
  try { return JSON.parse(raw); } catch { return { text: raw }; }
}

function htmlToText(value) {
  return String(value || '')
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<(script|style|iframe|object|template|svg|noscript)\b[^>]*>[\s\S]*?<\/\s*\1\s*>/gi, ' ')
    .replace(/<(?:img|source|video|audio|link|meta|base|form|input)\b[^>]*>/gi, ' ')
    .replace(/<br\s*\/?\s*>/gi, '\n')
    .replace(/<\/(?:p|div|tr|li|h[1-6]|section|article|table)\s*>/gi, '\n')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/[\t\u00a0 ]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function cellText(value, displayText = '') {
  if (value == null) return '';
  if (value instanceof Date) return displayText || value.toISOString();
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) return value.richText.map(part => part.text || '').join('');
    if (value.formula) return `${displayText || value.result || ''} [=${value.formula}]`;
    if ('text' in value) return String(value.text ?? '');
    if ('result' in value) return String(value.result ?? '');
    if ('hyperlink' in value) return String(value.hyperlink ?? '');
    return '';
  }
  return displayText || String(value);
}

function truncateText(value, max = MAX_SOURCE_CHARS) {
  const text = String(value || '').replace(/\u0000/g, '').trim();
  if (text.length <= max) return { text, truncated: false };
  const omitted = text.length - max;
  const marker = `\n[...truncated: ${omitted} characters omitted]`;
  return { text: `${text.slice(0, max - marker.length)}${marker}`, truncated: true };
}

function safeError(error, fallback) {
  const message = String(error?.message || '').slice(0, 300);
  return message || fallback;
}

function openSystemBrowser(rawUrl) {
  const url = new URL(rawUrl);
  if (url.protocol !== 'https:' || !['accounts.google.com', 'login.microsoftonline.com'].includes(url.hostname)) {
    throw new Error('The sign-in provider returned an unsupported browser URL.');
  }
  const command = process.platform === 'win32' ? 'rundll32.exe' : process.platform === 'darwin' ? 'open' : 'xdg-open';
  const args = process.platform === 'win32' ? ['url.dll,FileProtocolHandler', url.href] : [url.href];
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { detached: true, stdio: 'ignore', windowsHide: true });
    child.once('error', reject);
    child.once('spawn', () => { child.unref(); resolve(); });
  });
}

const GOOGLE_CLIENT_ID_RE = /^[A-Za-z0-9-]+\.apps\.googleusercontent\.com$/;
const MICROSOFT_CLIENT_ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export class LocalSourceMcp {
  constructor(config = {}) {
    this.config = {
      googleClientId: typeof config.googleClientId === 'string' ? config.googleClientId.trim().slice(0, 500) : '',
      microsoftClientId: typeof config.microsoftClientId === 'string' ? config.microsoftClientId.trim().slice(0, 100) : ''
    };
    this.googleToken = null;
    this.googleFlow = null;
    this.graphApp = null;
    this.graphClientId = '';
    this.graphAccount = null;
    this.graphFlow = null;

    this.server = new McpServer({ name: 'excel-qwen-grok-local-sources', version: '1.0.0' }, {
      instructions: 'Read-only search and review tools for the signed-in user. Never modify cloud files or mail.'
    });
    this.client = new Client({ name: 'excel-qwen-grok-local-source-client', version: '1.0.0' });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();

    this.server.registerTool('google_drive_search', {
      description: 'Search the signed-in Google Drive account for matching files. Read only.',
      inputSchema: { query: z.string().min(1).max(300) }
    }, async ({ query }) => textResult(await this.searchGoogleDrive(query)));
    this.server.registerTool('google_drive_read', {
      description: 'Read text from one selected Google Drive file. Read only.',
      inputSchema: { fileId: z.string().min(1).max(500), name: z.string().min(1).max(500) }
    }, async ({ fileId, name }) => textResult(await this.readGoogleDrive(fileId, name)));
    this.server.registerTool('gmail_search', {
      description: 'Search the signed-in Gmail account with Gmail search syntax. Read only.',
      inputSchema: { query: z.string().min(1).max(300) }
    }, async ({ query }) => textResult(await this.searchGmail(query)));
    this.server.registerTool('gmail_read_thread', {
      description: 'Read the full text of one selected Gmail thread. Read only.',
      inputSchema: { threadId: z.string().min(1).max(500) }
    }, async ({ threadId }) => textResult(await this.readGmailThread(threadId)));
    this.server.registerTool('onedrive_search', {
      description: 'Search the signed-in personal OneDrive account. Read only.',
      inputSchema: { query: z.string().min(1).max(300) }
    }, async ({ query }) => textResult(await this.searchOneDrive(query)));
    this.server.registerTool('onedrive_read', {
      description: 'Read text from one selected personal OneDrive file. Read only.',
      inputSchema: { fileId: z.string().min(1).max(500), name: z.string().min(1).max(500) }
    }, async ({ fileId, name }) => textResult(await this.readOneDrive(fileId, name)));

    this.ready = Promise.all([
      this.server.connect(serverTransport),
      this.client.connect(clientTransport)
    ]);
  }

  status() {
    if (this.googleFlow && Date.now() - this.googleFlow.createdAt > 10 * 60_000) this.clearGoogleFlow();
    const waiting = this.graphFlow && ['starting', 'waiting'].includes(this.graphFlow.status);
    return {
      google: {
        configured: GOOGLE_CLIENT_ID_RE.test(this.config.googleClientId),
        connected: Boolean(this.googleToken?.refreshToken || (this.googleToken?.accessToken && this.googleToken.expiresAt > Date.now())),
        connecting: Boolean(this.googleFlow)
      },
      onedrive: {
        configured: MICROSOFT_CLIENT_ID_RE.test(this.config.microsoftClientId),
        connected: Boolean(this.graphAccount),
        connecting: Boolean(waiting),
        error: this.graphFlow?.status === 'error' ? this.graphFlow.error : ''
      }
    };
  }

  async invoke(name, args) {
    await this.ready;
    return parseToolResult(await this.client.callTool({ name, arguments: args }));
  }

  clearGoogleFlow() {
    const flow = this.googleFlow;
    this.googleFlow = null;
    if (flow?.timer) clearTimeout(flow.timer);
    if (flow?.callbackServer?.listening) flow.callbackServer.close();
  }

  async startGoogleOAuth() {
    if (!GOOGLE_CLIENT_ID_RE.test(this.config.googleClientId)) throw new Error('Google cloud sign-in is not configured by the app owner yet.');
    if (this.googleFlow) throw new Error('Google sign-in is already open. Finish it in your browser or wait for it to expire.');
    const state = randomBytes(32).toString('base64url');
    const verifier = randomBytes(48).toString('base64url');
    const challenge = createHash('sha256').update(verifier).digest('base64url');
    const callbackServer = http.createServer(async (req, res) => {
      const callbackUrl = new URL(req.url || '/', 'http://127.0.0.1');
      if (req.method !== 'GET' || callbackUrl.pathname !== '/') {
        res.writeHead(callbackUrl.pathname === '/favicon.ico' ? 204 : 404, { 'Cache-Control': 'no-store' });
        return res.end();
      }
      const result = await this.completeGoogleOAuth(callbackUrl);
      res.writeHead(result.status, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
      res.end(result.html);
    });
    await new Promise((resolve, reject) => {
      callbackServer.once('error', reject);
      callbackServer.listen(0, '127.0.0.1', resolve);
    });
    const port = callbackServer.address().port;
    const redirectUri = `http://127.0.0.1:${port}`;
    const authorizationUrl = new URL('https://accounts.google.com/o/oauth2/v2/auth');
    authorizationUrl.searchParams.set('client_id', this.config.googleClientId);
    authorizationUrl.searchParams.set('redirect_uri', redirectUri);
    authorizationUrl.searchParams.set('response_type', 'code');
    authorizationUrl.searchParams.set('scope', GOOGLE_SCOPES.join(' '));
    authorizationUrl.searchParams.set('access_type', 'offline');
    authorizationUrl.searchParams.set('prompt', 'select_account consent');
    authorizationUrl.searchParams.set('state', state);
    authorizationUrl.searchParams.set('code_challenge', challenge);
    authorizationUrl.searchParams.set('code_challenge_method', 'S256');
    const flow = { state, verifier, redirectUri, createdAt: Date.now(), callbackServer };
    this.googleFlow = flow;
    flow.timer = setTimeout(() => { if (this.googleFlow === flow) this.clearGoogleFlow(); }, 10 * 60_000);
    flow.timer.unref();
    try {
      await openSystemBrowser(authorizationUrl.href);
      return { connecting: true };
    } catch (error) {
      this.clearGoogleFlow();
      throw new Error(safeError(error, 'Could not open the system browser.'));
    }
  }

  async completeGoogleOAuth(url) {
    const flow = this.googleFlow;
    if (!flow || Date.now() - flow.createdAt > 10 * 60_000) {
      this.clearGoogleFlow();
      return { status: 400, html: this.callbackPage(false, 'This sign-in request expired. Close this page and start again in the add-in.') };
    }
    if (url.searchParams.get('state') !== flow.state) {
      this.clearGoogleFlow();
      return { status: 400, html: this.callbackPage(false, 'The sign-in response could not be verified. Start again in the add-in.') };
    }
    if (url.searchParams.has('error')) {
      this.clearGoogleFlow();
      return { status: 400, html: this.callbackPage(false, 'Google sign-in was cancelled or declined.') };
    }
    const code = url.searchParams.get('code');
    if (!code) {
      this.clearGoogleFlow();
      return { status: 400, html: this.callbackPage(false, 'Google did not return an authorization code.') };
    }
    try {
      const response = await fetch('https://oauth2.googleapis.com/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: this.config.googleClientId,
          code,
          code_verifier: flow.verifier,
          grant_type: 'authorization_code',
          redirect_uri: flow.redirectUri
        }),
        redirect: 'error',
        signal: AbortSignal.timeout(20_000)
      });
      const tokenData = await response.json().catch(() => ({}));
      if (!response.ok || !tokenData.access_token) {
        const errorCode = typeof tokenData.error === 'string' && /^[a-z0-9_]{1,80}$/i.test(tokenData.error)
          ? tokenData.error
          : 'unknown_error';
        throw new Error(`Google token exchange failed: ${errorCode} (HTTP ${response.status}).`);
      }
      this.googleToken = {
        accessToken: tokenData.access_token,
        expiresAt: Date.now() + Math.max(60, Number(tokenData.expires_in) || 3600) * 1000,
        refreshToken: tokenData.refresh_token || this.googleToken?.refreshToken || ''
      };
      this.clearGoogleFlow();
      return { status: 200, html: this.callbackPage(true, 'Google Drive and Gmail are connected for this Windows user. You can close this page and return to the add-in.') };
    } catch (error) {
      this.clearGoogleFlow();
      return { status: 400, html: this.callbackPage(false, safeError(error, 'Google sign-in failed. Check the OAuth setup and try again.')) };
    }
  }

  callbackPage(ok, message) {
    const safe = String(message).replace(/[&<>"']/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[character]);
    return `<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Source sign-in</title><body style="font:16px Segoe UI,Arial,sans-serif;max-width:620px;margin:15vh auto;padding:24px;color:#242424"><h2>${ok ? 'Connected' : 'Sign-in not completed'}</h2><p>${safe}</p></body></html>`;
  }

  async googleAccessToken() {
    if (!this.googleToken) throw new Error('Connect Google Drive and Gmail first.');
    if (this.googleToken.accessToken && this.googleToken.expiresAt > Date.now() + 60_000) return this.googleToken.accessToken;
    if (!this.googleToken.refreshToken) throw new Error('Google sign-in expired. Reconnect Google Drive and Gmail.');
    const response = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: this.config.googleClientId,
        refresh_token: this.googleToken.refreshToken,
        grant_type: 'refresh_token'
      }),
      redirect: 'error',
      signal: AbortSignal.timeout(20_000)
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok || !data.access_token) {
      this.googleToken = null;
      throw new Error('Google sign-in expired. Reconnect Google Drive and Gmail.');
    }
    this.googleToken = {
      accessToken: data.access_token,
      expiresAt: Date.now() + Math.max(60, Number(data.expires_in) || 3600) * 1000,
      refreshToken: data.refresh_token || this.googleToken.refreshToken
    };
    return this.googleToken.accessToken;
  }

  async remoteGoogleTool(serviceUrl, name, args) {
    const token = await this.googleAccessToken();
    const transport = new StreamableHTTPClientTransport(new URL(serviceUrl), {
      requestInit: { headers: { Authorization: `Bearer ${token}` } },
      redirectPolicy: 'same-origin'
    });
    const client = new Client({ name: 'excel-qwen-grok-google-mcp', version: '1.0.0' });
    try {
      await client.connect(transport);
      const result = await client.callTool({ name, arguments: args });
      return parseToolResult(result);
    } finally {
      await client.close().catch(() => {});
    }
  }

  async searchGoogleDrive(query) {
    const escaped = query.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    const response = await this.remoteGoogleTool(GOOGLE_DRIVE_URL, 'search_files', {
      query: `fullText contains '${escaped}'`, pageSize: 10, excludeContentSnippets: false
    });
    const files = Array.isArray(response.files) ? response.files : [];
    return {
      source: 'google_drive',
      items: files.map(file => ({
        id: String(file.id || ''), name: String(file.title || 'Untitled').slice(0, 500),
        detail: [file.mimeType, file.modifiedTime].filter(Boolean).join(' · '),
        snippet: String(file.contentSnippet || '').slice(0, 1200), url: String(file.viewUrl || ''),
        mimeType: String(file.mimeType || ''), kind: 'Google Drive'
      })).filter(file => file.id)
    };
  }

  async readGoogleDrive(fileId, name = 'Google Drive file') {
    const response = await this.remoteGoogleTool(GOOGLE_DRIVE_URL, 'read_file_content', { fileId, includeComments: false });
    const text = String(response.fileContent || response.text || '').trim();
    if (!text) throw new Error('Google Drive did not return readable text for this file.');
    const clipped = truncateText(text);
    return { name, kind: 'Google Drive', ...clipped };
  }

  async searchGmail(query) {
    const response = await this.remoteGoogleTool(GOOGLE_GMAIL_URL, 'search_threads', { query, pageSize: 10 });
    const threads = Array.isArray(response.threads) ? response.threads : [];
    return {
      source: 'gmail',
      items: threads.map(thread => {
        const message = thread.messages?.[thread.messages.length - 1] || thread.messages?.[0] || {};
        return {
          id: String(thread.id || ''), name: String(message.subject || 'Email thread').slice(0, 500),
          detail: [message.sender, message.date].filter(Boolean).join(' · '),
          snippet: String(message.snippet || '').slice(0, 1200), kind: 'Gmail thread'
        };
      }).filter(item => item.id)
    };
  }

  async readGmailThread(threadId) {
    const response = await this.remoteGoogleTool(GOOGLE_GMAIL_URL, 'get_thread', { threadId, messageFormat: 'FULL_CONTENT' });
    const messages = Array.isArray(response.messages) ? response.messages : [];
    const body = messages.map(message => {
      const content = String(message.plaintextBody || htmlToText(message.htmlBody) || message.snippet || '').trim();
      return [
        message.subject ? `Subject: ${message.subject}` : '',
        message.sender ? `From: ${message.sender}` : '',
        message.date ? `Date: ${message.date}` : '',
        content
      ].filter(Boolean).join('\n');
    }).filter(Boolean).join('\n\n---\n\n');
    if (!body) throw new Error('Gmail did not return readable text for this thread.');
    return { name: messages[0]?.subject || 'Gmail thread', kind: 'Gmail thread', ...truncateText(body) };
  }

  async startOneDriveInteractiveFlow() {
    const clientId = this.config.microsoftClientId;
    if (!MICROSOFT_CLIENT_ID_RE.test(clientId)) throw new Error('OneDrive sign-in is not configured by the app owner yet.');
    if (this.graphFlow && ['starting', 'waiting'].includes(this.graphFlow.status)) return { connecting: true };
    if (!this.graphApp || this.graphClientId !== clientId) {
      this.graphApp = new PublicClientApplication({ auth: { clientId, authority: 'https://login.microsoftonline.com/consumers/' } });
      this.graphClientId = clientId;
      this.graphAccount = null;
    }
    const flow = { status: 'starting', error: '', promise: null };
    this.graphFlow = flow;
    flow.promise = this.graphApp.acquireTokenInteractive({
      scopes: GRAPH_SCOPES,
      prompt: 'select_account',
      openBrowser: openSystemBrowser,
      successTemplate: '<!doctype html><meta charset="utf-8"><title>OneDrive connected</title><h2>OneDrive connected</h2><p>Close this window and return to Excel or Word.</p>',
      errorTemplate: '<!doctype html><meta charset="utf-8"><title>Sign-in not completed</title><h2>Sign-in not completed</h2><p>Close this window and retry Connect in Excel or Word.</p>'
    }).then(async result => {
      if (!result?.account) throw new Error('Microsoft sign-in did not return an account.');
      if (this.graphFlow !== flow) {
        await this.graphApp.getTokenCache().removeAccount(result.account).catch(() => {});
        return false;
      }
      this.graphAccount = result.account;
      flow.status = 'connected';
      return true;
    }).catch(error => {
      flow.status = 'error';
      flow.error = safeError(error, 'Microsoft sign-in failed.');
      return false;
    });
    flow.status = 'waiting';
    return { connecting: true };
  }

  async graphAccessToken() {
    if (!this.graphApp || !this.graphAccount) throw new Error('Connect OneDrive first.');
    try {
      const result = await this.graphApp.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: this.graphAccount });
      if (!result?.accessToken) throw new Error('Microsoft sign-in expired.');
      return result.accessToken;
    } catch {
      this.graphAccount = null;
      throw new Error('OneDrive sign-in expired. Reconnect OneDrive.');
    }
  }

  async graphFetch(url, options = {}) {
    const token = await this.graphAccessToken();
    const response = await fetch(url, {
      ...options,
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', ...(options.headers || {}) },
      redirect: 'error',
      signal: AbortSignal.timeout(45_000)
    });
    if (!response.ok) {
      if (response.status === 401 || response.status === 403) throw new Error('Microsoft Graph denied the read. Reconnect OneDrive or check the Files.Read permission.');
      throw new Error(`OneDrive request failed (${response.status}).`);
    }
    return response;
  }

  async searchOneDrive(query) {
    const path = `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${encodeURIComponent(query.replace(/'/g, "''"))}')`;
    const url = new URL(path);
    url.searchParams.set('$top', '10');
    url.searchParams.set('$select', 'id,name,size,lastModifiedDateTime,webUrl,file,folder,parentReference');
    const response = await this.graphFetch(url.href);
    const data = await response.json();
    return {
      source: 'onedrive',
      items: (Array.isArray(data.value) ? data.value : []).map(file => ({
        id: String(file.id || ''), name: String(file.name || 'Untitled').slice(0, 500),
        detail: [file.file?.mimeType, file.lastModifiedDateTime].filter(Boolean).join(' · '),
        snippet: '', url: String(file.webUrl || ''), mimeType: String(file.file?.mimeType || ''), kind: 'OneDrive'
      })).filter(file => file.id && file.mimeType)
    };
  }

  async downloadOneDriveFile(fileId) {
    const endpoint = `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(fileId)}/content`;
    const token = await this.graphAccessToken();
    let response = await fetch(endpoint, {
      headers: { Authorization: `Bearer ${token}`, Accept: '*/*' },
      redirect: 'manual',
      signal: AbortSignal.timeout(45_000)
    });
    for (let redirects = 0; [301, 302, 303, 307, 308].includes(response.status) && redirects < 3; redirects++) {
      const location = response.headers.get('location');
      let downloadUrl;
      try { downloadUrl = new URL(location); } catch { throw new Error('OneDrive returned an invalid download link.'); }
      const host = downloadUrl.hostname.toLowerCase();
      const allowed = downloadUrl.protocol === 'https:' && (host === 'onedrive.live.com' || DRIVE_DOWNLOAD_HOSTS.some(suffix => host.endsWith(suffix)));
      if (!allowed) throw new Error('OneDrive returned an unsupported download host.');
      response = await fetch(downloadUrl, { redirect: 'manual', signal: AbortSignal.timeout(45_000) });
    }
    if ([301, 302, 303, 307, 308].includes(response.status)) throw new Error('OneDrive returned too many download redirects.');
    if (!response.ok) {
      if (response.status === 401 || response.status === 403) throw new Error('Microsoft Graph denied the file read. Reconnect OneDrive or check the Files.Read permission.');
      throw new Error(`OneDrive file download failed (${response.status}).`);
    }
    const contentLength = Number(response.headers.get('content-length') || 0);
    if (contentLength > MAX_SOURCE_BYTES) throw new Error('This file is over the 12 MB review limit.');
    const chunks = [];
    let size = 0;
    const reader = response.body?.getReader();
    if (!reader) throw new Error('OneDrive returned an empty download.');
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      size += value.byteLength;
      if (size > MAX_SOURCE_BYTES) { await reader.cancel().catch(() => {}); throw new Error('This file is over the 12 MB review limit.'); }
      chunks.push(Buffer.from(value));
    }
    const buffer = Buffer.concat(chunks, size);
    return buffer;
  }

  async readOneDrive(fileId, name) {
    const ext = String(name).split('.').pop().toLowerCase();
    const buffer = await this.downloadOneDriveFile(fileId);
    let extracted = '';
    if (['xlsx', 'xlsm'].includes(ext)) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const rows = [];
      let charCount = 0;
      for (const sheet of workbook.worksheets) {
        rows.push(`## Sheet: ${sheet.name}`);
        for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 1500); rowNumber++) {
          const row = sheet.getRow(rowNumber);
          const cells = [];
          for (let column = 1; column <= Math.min(sheet.columnCount, 60); column++) {
            const cell = row.getCell(column);
            cells.push(cellText(cell.value, cell.text));
          }
          const line = cells.join('\t').replace(/[\t ]+$/g, '');
          if (line.trim()) rows.push(line);
          charCount += line.length + 1;
          if (charCount >= MAX_SOURCE_CHARS) break;
        }
        if (charCount >= MAX_SOURCE_CHARS) break;
      }
      extracted = rows.join('\n');
    } else if (ext === 'docx') {
      const result = await mammoth.extractRawText({ buffer });
      extracted = result.value;
    } else if (['txt', 'md', 'csv', 'tsv', 'json', 'xml', 'html', 'htm'].includes(ext)) {
      extracted = buffer.toString('utf8');
      if (ext === 'html' || ext === 'htm' || ext === 'xml') extracted = htmlToText(extracted);
    } else {
      throw new Error('OneDrive review supports XLSX/XLSM, DOCX, HTML, TXT, MD, CSV, TSV, JSON and XML files. Use Google Drive MCP for PDF review.');
    }
    if (!String(extracted).trim()) throw new Error('OneDrive did not return readable text for this file.');
    return { name, kind: `OneDrive ${ext.toUpperCase()}`, ...truncateText(extracted) };
  }

  async search(source, query) {
    if (typeof query !== 'string' || !query.trim() || query.length > 300) throw new Error('Enter a search query of 1 to 300 characters.');
    const names = { google_drive: 'google_drive_search', gmail: 'gmail_search', onedrive: 'onedrive_search' };
    if (!names[source]) throw new Error('Choose Google Drive, Gmail or OneDrive.');
    return this.invoke(names[source], { query: query.trim() });
  }

  async read(source, input = {}) {
    const names = {
      google_drive: ['google_drive_read', { fileId: input.id, name: input.name || 'Google Drive file' }],
      gmail: ['gmail_read_thread', { threadId: input.id }],
      onedrive: ['onedrive_read', { fileId: input.id, name: input.name }]
    };
    if (!names[source] || typeof input.id !== 'string' || !input.id || input.id.length > 500) throw new Error('Choose a valid search result to read.');
    if (source === 'onedrive' && (typeof input.name !== 'string' || !input.name || input.name.length > 500)) throw new Error('The OneDrive file name is invalid.');
    return this.invoke(names[source][0], names[source][1]);
  }

  async disconnect(source) {
    if (source === 'google') {
      this.googleToken = null;
      this.clearGoogleFlow();
      return;
    }
    if (source === 'onedrive') {
      const account = this.graphAccount;
      if (account && this.graphApp) await this.graphApp.getTokenCache().removeAccount(account).catch(() => {});
      this.graphAccount = null;
      this.graphFlow = null;
      return;
    }
    throw new Error('Choose Google or OneDrive.');
  }
}
