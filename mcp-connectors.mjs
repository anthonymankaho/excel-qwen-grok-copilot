import { spawn } from 'node:child_process';
import { Client, InMemoryTransport } from '@modelcontextprotocol/client';
import { McpServer } from '@modelcontextprotocol/server';
import { PublicClientApplication } from '@azure/msal-node';
import { extractDocx, extractPdf, extractWorkbook } from './extract-files.mjs';
import { z } from 'zod';

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
  if (url.protocol !== 'https:' || url.hostname !== 'login.microsoftonline.com') {
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

const MICROSOFT_CLIENT_ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export class LocalSourceMcp {
  constructor(config = {}) {
    this.config = {
      microsoftClientId: typeof config.microsoftClientId === 'string' ? config.microsoftClientId.trim().slice(0, 100) : ''
    };
    this.graphApp = null;
    this.graphClientId = '';
    this.graphAccount = null;
    this.graphFlow = null;

    this.server = new McpServer({ name: 'excel-qwen-grok-local-sources', version: '1.0.0' }, {
      instructions: 'Read-only search and review tools for the signed-in user. Never modify cloud files or mail.'
    });
    this.client = new Client({ name: 'excel-qwen-grok-local-source-client', version: '1.0.0' });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();

    this.server.registerTool('onedrive_search', {
      description: 'Search the signed-in personal OneDrive account. Read only.',
      inputSchema: { query: z.string().min(1).max(300) }
    }, async ({ query }) => textResult(await this.searchOneDrive(query)));
    this.server.registerTool('onedrive_read', {
      description: 'Read text from one selected personal OneDrive file. Read only. PDF, Word, and Excel include a coverage line.',
      inputSchema: { fileId: z.string().min(1).max(500), name: z.string().min(1).max(500) }
    }, async ({ fileId, name }) => textResult(await this.readOneDrive(fileId, name, { transcribePage: this._transcribePage || null })));

    this.ready = Promise.all([
      this.server.connect(serverTransport),
      this.client.connect(clientTransport)
    ]);
  }

  status() {
    const waiting = this.graphFlow && ['starting', 'waiting'].includes(this.graphFlow.status);
    return {
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

  async readOneDrive(fileId, name, hooks = {}) {
    const ext = String(name).split('.').pop().toLowerCase();
    const buffer = await this.downloadOneDriveFile(fileId);
    let extracted = '';
    if (['xlsx', 'xlsm'].includes(ext)) {
      extracted = (await extractWorkbook(buffer)).text;
    } else if (ext === 'docx') {
      extracted = (await extractDocx(buffer)).text;
    } else if (ext === 'pdf') {
      extracted = (await extractPdf(buffer, { transcribePage: hooks.transcribePage || null })).text;
    } else if (['txt', 'md', 'csv', 'tsv', 'json', 'xml', 'html', 'htm'].includes(ext)) {
      extracted = buffer.toString('utf8');
      if (ext === 'html' || ext === 'htm' || ext === 'xml') extracted = htmlToText(extracted);
    } else {
      throw new Error('OneDrive review supports PDF, DOCX, XLSX/XLSM, HTML, TXT, MD, CSV, TSV, JSON and XML files.');
    }
    if (!String(extracted).trim()) throw new Error('OneDrive did not return readable text for this file.');
    return { name, kind: `OneDrive ${ext.toUpperCase()}`, ...truncateText(extracted) };
  }

  async search(source, query) {
    if (typeof query !== 'string' || !query.trim() || query.length > 300) throw new Error('Enter a search query of 1 to 300 characters.');
    if (source !== 'onedrive') throw new Error('Choose OneDrive.');
    return this.invoke('onedrive_search', { query: query.trim() });
  }

  async read(source, input = {}, hooks = {}) {
    if (source !== 'onedrive' || typeof input.id !== 'string' || !input.id || input.id.length > 500) throw new Error('Choose a valid OneDrive search result to read.');
    if (typeof input.name !== 'string' || !input.name || input.name.length > 500) throw new Error('The OneDrive file name is invalid.');
    this._transcribePage = hooks.transcribePage || null;
    try { return await this.invoke('onedrive_read', { fileId: input.id, name: input.name }); }
    finally { this._transcribePage = null; }
  }

  async disconnect(source) {
    if (source === 'onedrive') {
      const account = this.graphAccount;
      if (account && this.graphApp) await this.graphApp.getTokenCache().removeAccount(account).catch(() => {});
      this.graphAccount = null;
      this.graphFlow = null;
      return;
    }
    throw new Error('Choose OneDrive.');
  }
}
