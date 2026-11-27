const API = '';
const PANE_BUILD = '2026-11-26.4'; // Must match relay-core RELAY_BUILD; a stale relay process is detected and reported.
const CHAT_MODELS = ['auto', 'qwen4-max', 'qwen3.8-max', 'qwen3.8-flash'];
const CHAT_EFFORTS = ['off', 'low', 'medium', 'high', 'xhigh'];
const CONTEXT_CAP = 45_000; // Relay rejects messages over 50,000 chars; leave room for the prompt itself.
const ATTACHMENT_CAP = 26_000;
const ATTACHMENT_FILE_CAP = 18_000;
const ATTACHMENT_MAX_BYTES = 5 * 1024 * 1024;
const ATTACHMENT_RELAY_MAX_BYTES = 20 * 1024 * 1024; // binary types: only extracted text enters context, so the raw file may be larger
const ATTACHMENT_MAX_FILES = 3;
const MAX_CONTEXT_ROWS = 150;
const MAX_CONTEXT_COLS = 30;
const READ_MAX_CELLS = 2000; // One read-back batch; larger requests are paged with contextRequest.offset, never silently trimmed.
const PROFILE_BATCH_CELLS = 20_000; // Profiles and finance checks scan the full requested range in batches; never claim more coverage than was read.
const PROFILE_DISTINCT = 12; // Sample values shown per column; enough to identify a grouping field.
const PROFILE_DISTINCT_CAP = 5000;
const CELL_CHAR_LIMIT = 32767;
const FINANCE_EVIDENCE_INSTRUCTION = `\n\n[Finance evidence rules — apply only when the request or supplied data is finance/accounting related. Keep the response in English and Traditional Chinese, with aligned concise sections: Conclusion / 結論; Verified facts / 已核實資料; Calculation checks / 計算檢查; Assumptions and missing information / 假設及缺漏。 Cite each material claim and amount to an exact supplied location: Excel [Sheet!A1] or [Sheet!A1:C8], Word [Selection:L0001-L0005], or report [filename:L0001-L0005]. Excel context identifies column letters and row numbers; attached text identifies line numbers and preserves HTML headings. Never invent a cell, line, period, currency, or value. If only workbook metadata was supplied, request the exact range before claiming a value is verified. Show calculation inputs, formula, result, and rounding/tolerance. Say when a calculation is model-generated and has not been independently recalculated by the add-in. Only call an add-in check PASS when its validation card says PASS; report FAIL, WARNING, or NEEDS_USER as shown. Keep assumptions separate from source facts and state when evidence is insufficient. For a whole-workbook review, structure the answer as: Conclusion / 結論; Passed checks / 已通過 (table); Exceptions by priority / 需處理 (issue, likely cause, recommended action, exact cells); Earlier open items / 原有未結事項; Suggested formulas / 建議公式; Next steps / 下一步 (offer fixes only after the user confirms). State which sheets and rows were covered. Do not change the required Agent JSON envelope.`;
const KEY_STORE = 'eqgc.keys.v1';
const HOSTED_KEY_STORE = 'eqgc.hosted.qwen-key.v1';
const RISKY_FORMULA = /^=.*\b(?:WEBSERVICE|FILTERXML|RTD|REGISTER\.ID|CALL|DDE|INFO|HYPERLINK)\s*\(/i;
const state = { messages: [], lastAnswer: '', lastPrompt: '', busy: false, officeReady: false, forceSheet: false, turn: 0, host: '', audit: [], attachments: [], attachmentErrors: [], sessionId: null, sessionStarted: null, workbookName: '', githubAccount: null };
let activeAgent = null;
let agentCheckpoint = null;
const $ = id => document.getElementById(id);
let abortChat = null;
let sourcePollTimer = null;
try { localStorage.removeItem('eqgc.cloud-sources.v1'); } catch { /* obsolete connector secrets are not retained */ }
// Excel web embeds the pane cross-site; tracking prevention can block iframe storage entirely.
let storageOk = true;
try { localStorage.setItem('eqgc.probe', '1'); localStorage.removeItem('eqgc.probe'); } catch { storageOk = false; }

function storedKeys() {
  try { return JSON.parse(localStorage.getItem(KEY_STORE) || '{}'); } catch { return {}; }
}
function rememberKey(provider, key) {
  const all = storedKeys();
  if (key) all[provider] = key; else delete all[provider];
  try {
    localStorage.setItem(KEY_STORE, JSON.stringify(all));
    return localStorage.getItem(KEY_STORE) === JSON.stringify(all);
  }
  catch { storageOk = false; return false; }
}
let keysReady = null;
let sessionByoKey = '';
let hostedKeyReady = null;
let hostedKeyStore = 'session';
function hostedRelay() {
  const host = globalThis.location?.hostname || '';
  return host !== '' && host !== '127.0.0.1' && host !== 'localhost';
}
// Hosted (Marketplace) pane: the user's OWN key, kept in this browser only, sent per request.
function byoKey() {
  if (sessionByoKey) return sessionByoKey;
  return storedKeys().qwen || '';
}
function officeKeyStorage() {
  const storage = globalThis.OfficeRuntime?.storage;
  return storage && typeof storage.getItem === 'function' && typeof storage.setItem === 'function' && typeof storage.removeItem === 'function' ? storage : null;
}
async function loadHostedKey() {
  if (!hostedRelay()) return;
  if (!hostedKeyReady) hostedKeyReady = (async () => {
    const storage = officeKeyStorage();
    if (storage) {
      try {
        const key = await storage.getItem(HOSTED_KEY_STORE);
        if (typeof key === 'string' && key.startsWith('sk-')) {
          sessionByoKey = key;
          hostedKeyStore = 'office';
          return;
        }
      } catch { /* Fall back to browser storage. */ }
    }
    const key = storedKeys().qwen;
    if (typeof key === 'string' && key.startsWith('sk-')) {
      sessionByoKey = key;
      hostedKeyStore = 'browser';
    }
  })();
  return hostedKeyReady;
}
async function persistHostedKey(key) {
  const storage = officeKeyStorage();
  if (storage) {
    try {
      await storage.setItem(HOSTED_KEY_STORE, key);
      if (await storage.getItem(HOSTED_KEY_STORE) === key) {
        rememberKey('qwen', ''); // Remove any older browser copy when accessible.
        hostedKeyStore = 'office';
        return true;
      }
    } catch { /* Fall back to browser storage. */ }
  }
  if (rememberKey('qwen', key)) { hostedKeyStore = 'browser'; return true; }
  hostedKeyStore = 'session';
  return false;
}
async function clearHostedKey() {
  let cleared = true;
  const storage = officeKeyStorage();
  if (storage) {
    try { await storage.removeItem(HOSTED_KEY_STORE); }
    catch { cleared = false; }
  }
  const oldBrowserKey = storedKeys().qwen;
  if (!rememberKey('qwen', '') && oldBrowserKey) cleared = false;
  sessionByoKey = '';
  hostedKeyStore = 'session';
  return cleared;
}
function ensureKeys() {
  if (hostedRelay()) return loadHostedKey();
  if (!keysReady) keysReady = (async () => {
    const stored = storedKeys();
    let allOk = true;
    for (const provider of ['qwen']) {
      if (stored[provider]) { try { await api('configure', { provider, apiKey: stored[provider] }); } catch { allOk = false; } }
    }
    if (!allOk) keysReady = null; // relay was offline; retry on next question
  })();
  return keysReady;
}

async function api(path, body) {
  const headers = {};
  if (body) headers['Content-Type'] = 'application/json';
  const response = await fetch(`${API}/api/${path}`, { method: body ? 'POST' : 'GET', headers, body: body ? JSON.stringify(body) : undefined });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `Request failed (${response.status})`);
  return data;
}

function toast(message) {
  const el = $('toast'); el.textContent = message; el.classList.remove('hidden');
  clearTimeout(toast.timer); toast.timer = setTimeout(() => el.classList.add('hidden'), 3200);
}

function setSourceStatus(id, value, connected = false) {
  const el = $(id);
  if (!el) return;
  el.textContent = value;
  el.style.color = connected ? '#168251' : '#777d89';
}

async function refreshCloudSources() {
  $('cloudSources').classList.toggle('hidden', hostedRelay());
  if (hostedRelay()) return;
  try {
    const status = await api('connectors/status');
    const onedrive = status.onedrive || {};
    setSourceStatus('onedriveSourceStatus', onedrive.connected ? 'Connected · this PC' : onedrive.connecting ? 'Finish sign-in in your browser' : onedrive.error ? onedrive.error : onedrive.configured ? 'Not connected · sign in with your account' : 'Not set up by app owner', onedrive.connected);
    $('connectOneDrive').disabled = !onedrive.configured || onedrive.connected || onedrive.connecting;
    $('connectOneDrive').textContent = onedrive.connected ? 'Connected' : onedrive.connecting ? 'Signing in…' : 'Connect';
    $('disconnectOneDrive').classList.toggle('hidden', !onedrive.connected);
    if (onedrive.connecting) startSourcePolling();
    else stopSourcePolling();
  } catch (error) {
    setSourceStatus('onedriveSourceStatus', explainError(error.message || 'Local relay unavailable'));
  }
}

function startSourcePolling() {
  if (sourcePollTimer) return;
  sourcePollTimer = setInterval(refreshCloudSources, 2500);
}

function stopSourcePolling() {
  if (sourcePollTimer) clearInterval(sourcePollTimer);
  sourcePollTimer = null;
}

function setSidebar(open) {
  $('sidebar').classList.toggle('hidden', !open);
  $('sidebarToggle').setAttribute('aria-expanded', String(open));
  try { localStorage.setItem('eqgc.sidebar.v1', open ? '1' : '0'); } catch { /* storage optional */ }
}

function persistSession(outcome) {
  if (!state.sessionId || hostedRelay()) return;
  api('sessions/save', { id: state.sessionId, started: state.sessionStarted, workbook: state.workbookName, prompt: state.lastPrompt, outcome, messages: state.messages, audit: state.audit, checkpoint: agentCheckpoint }).catch(() => {});
}

async function refreshHistory() {
  const card = $('historyCard');
  if (!card) return;
  if (hostedRelay()) { card.classList.add('hidden'); return; }
  try {
    const data = await api('sessions');
    $('historyCount').textContent = `${data.sessions.length} sessions`;
    const host = $('historyList');
    host.replaceChildren();
    if (!data.sessions.length) {
      const empty = document.createElement('p'); empty.className = 'hint'; empty.textContent = 'No saved sessions yet on this PC.'; host.append(empty); return;
    }
    for (const s of data.sessions) {
      const row = document.createElement('div'); row.className = 'source-connect';
      const info = document.createElement('div');
      const strong = document.createElement('strong'); strong.textContent = `${String(s.updated || '').slice(0, 16).replace('T', ' ')} · ${s.workbook || 'workbook'}`;
      const small = document.createElement('small'); small.textContent = `${s.prompt} — ${s.outcome}`;
      info.append(strong, small);
      const resume = document.createElement('button'); resume.className = 'text-button'; resume.type = 'button'; resume.textContent = 'Resume';
      resume.addEventListener('click', () => resumeSession(s.id));
      const del = document.createElement('button'); del.className = 'text-button'; del.type = 'button'; del.textContent = 'Delete';
      del.addEventListener('click', async () => {
        if (del.dataset.armed !== '1') { del.dataset.armed = '1'; del.textContent = 'Confirm delete?'; return; }
        try {
          await api('sessions/delete', { id: s.id });
          await refreshHistory();
          toast('Session deleted.');
        } catch (error) { toast(explainError(error.message)); }
      });
      row.append(info, resume, del);
      host.append(row);
    }
  } catch (error) { $('historyCount').textContent = explainError(error.message || 'Local relay unavailable'); }
}

async function resumeSession(id) {
  if (state.busy) return toast('Stop the current run before resuming.');
  try {
    const s = await api(`sessions/read?id=${encodeURIComponent(id)}`);
    if (!s || !Array.isArray(s.messages)) return toast('Session not found.');
    state.sessionId = s.id;
    state.sessionStarted = s.started;
    state.lastPrompt = s.prompt || '';
    state.messages = s.messages.map(m => ({ role: m.role === 'assistant' ? 'assistant' : 'user', content: String(m.content || '') }));
    state.audit = Array.isArray(s.audit) ? s.audit : [];
    const conversation = $('conversation');
    conversation.replaceChildren();
    for (const m of state.messages) renderMessage(m.role, m.content);
    const lastAssistant = [...state.messages].reverse().find(m => m.role === 'assistant');
    state.lastAnswer = lastAssistant ? lastAssistant.content : '';
    $('insertAnswer').disabled = !state.lastAnswer;
    agentCheckpoint = s.checkpoint || { prompt: state.lastPrompt, progress: [], observation: '', observationSource: '' };
    showAgentStatus(0, 'Resumed · 已恢復');
    syncAgentControls();
    toast(`Resumed session ${s.id}`);
  } catch (error) { toast(explainError(error.message)); }
}

async function refreshMemory() {
  const card = $('memoryCard');
  if (!card) return;
  if (hostedRelay()) { card.classList.add('hidden'); return; }
  try {
    const data = await api('memory');
    const text = String(data.text || '');
    $('memoryText').textContent = text || '(empty)';
    $('memoryCount').textContent = `${text.trim() ? text.trim().split('\n').length : 0} notes`;
  } catch (error) { $('memoryCount').textContent = explainError(error.message || 'Local relay unavailable'); }
}

async function addMemoryNote() {
  const note = $('memoryNote').value.trim();
  if (!note) return toast('Type a note to remember.');
  try {
    await api('memory/add', { note });
    $('memoryNote').value = '';
    await refreshMemory();
    toast('Memory note saved on this PC.');
  } catch (error) { toast(explainError(error.message)); }
}

async function refreshSkills() {
  const card = $('skillsCard');
  if (!card) return;
  if (hostedRelay()) { card.classList.add('hidden'); return; }
  try {
    const data = await api('skills');
    $('skillsCount').textContent = data.root ? `${data.skills.length} · OneDrive` : 'No local skills folder';
    const host = $('skillsList');
    host.replaceChildren();
    if (!data.skills.length) {
      const empty = document.createElement('p'); empty.className = 'hint'; empty.textContent = 'No skills found in the local OneDrive Copilot skills folder.'; host.append(empty); return;
    }
    for (const skill of data.skills) {
      const row = document.createElement('div'); row.className = 'source-connect';
      const info = document.createElement('div');
      const strong = document.createElement('strong'); strong.textContent = skill.id;
      const small = document.createElement('small'); small.textContent = skill.description.slice(0, 140);
      info.append(strong, small);
      const use = document.createElement('button'); use.className = 'text-button'; use.type = 'button'; use.textContent = '@use';
      use.addEventListener('click', () => { $('prompt').value = `@${skill.id} ${$('prompt').value}`; $('prompt').focus(); });
      row.append(info, use);
      host.append(row);
    }
  } catch (error) { $('skillsCount').textContent = explainError(error.message || 'Local relay unavailable'); }
}

function githubTokenValue() {
  return ($('githubMcpToken')?.value || '').trim();
}

function githubAccountOptedOut() {
  return localStorage.getItem('eqgc.github.account') === 'off';
}

function showGitHubAccount(connected, login) {
  setSourceStatus('githubMcpStatus', connected ? `Connected as ${login}` : 'Not connected', connected);
  $('connectGitHub')?.classList.toggle('hidden', connected);
  $('disconnectGitHub')?.classList.toggle('hidden', !connected);
}

function githubAuthBody(extra = {}) {
  const token = githubTokenValue();
  if (token) return { ...extra, token };
  if (!githubAccountOptedOut() && state.githubAccount?.connected) return { ...extra, useAccount: true };
  return { ...extra };
}

const GITHUB_REPO_STORE = 'eqgc.github.repo';

function selectedGitHubRepo() {
  try {
    const saved = JSON.parse(localStorage.getItem(GITHUB_REPO_STORE) || '{}');
    const fullName = String(saved.fullName || '');
    if (!/^[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/.test(fullName)) return '';
    return fullName;
  } catch { return ''; }
}

function repoShortName(fullName) {
  const parts = String(fullName || '').split('/');
  return parts[1] || 'No Repository';
}

function refreshGitHubRepoPick() {
  const fullName = selectedGitHubRepo();
  const button = $('githubRepoPick');
  if (!button) return;
  button.textContent = fullName ? repoShortName(fullName) : 'No Repository';
  button.title = fullName ? `${fullName} · GitHub repository` : 'No Repository · 空白工作區';
  $('githubRepoNone')?.setAttribute('aria-checked', String(!fullName));
}

function setGitHubRepo(fullName) {
  if (!fullName) localStorage.removeItem(GITHUB_REPO_STORE);
  else localStorage.setItem(GITHUB_REPO_STORE, JSON.stringify({ fullName }));
  refreshGitHubRepoPick();
  setGitHubRepoMenu(false);
}

function githubRepoNote() {
  const fullName = selectedGitHubRepo();
  if (!fullName || hostedRelay() || !$('includeContext').checked) return '';
  return `\n\n[GitHub repository]\n${fullName}\nUse this repository for githubRequest and githubWrite unless the user names another repository.`;
}

function setGitHubRepoMenu(open) {
  $('githubRepoMenu')?.classList.toggle('hidden', !open);
  $('githubRepoPick')?.setAttribute('aria-expanded', String(open));
  if (open) {
    setWorkbookCard(false);
    setAttachMenu(false);
    loadGitHubRepoMenu();
  }
}

async function loadGitHubRepoMenu() {
  const list = $('githubRepoList');
  const empty = $('githubRepoEmpty');
  const connect = $('githubRepoConnect');
  if (!list) return;
  list.replaceChildren();
  const auth = githubAuthBody();
  const signedIn = Boolean(auth.token || auth.useAccount);
  connect?.classList.toggle('hidden', signedIn);
  if (!signedIn) {
    if (empty) {
      empty.classList.remove('hidden');
      empty.textContent = 'Connect to Github in MCP to list repositories. 請先在 MCP 按 Connect to Github。';
    }
    return;
  }
  if (empty) {
    empty.classList.remove('hidden');
    empty.textContent = 'Loading repositories… 正在讀取 repo…';
  }
  try {
    const result = await api('connectors/github/repos', githubAuthBody({ query: '' }));
    const items = Array.isArray(result.items) ? result.items : [];
    if (empty) empty.classList.toggle('hidden', items.length > 0);
    if (empty && !items.length) empty.textContent = 'No repositories were returned. 沒有讀到 repo。';
    const selected = selectedGitHubRepo();
    for (const item of items) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'repo-option';
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(item.fullName === selected));
      const strong = document.createElement('strong');
      strong.textContent = repoShortName(item.fullName);
      const small = document.createElement('small');
      small.textContent = item.fullName;
      button.append(strong, small);
      button.addEventListener('click', () => setGitHubRepo(item.fullName));
      list.append(button);
    }
  } catch (error) {
    if (empty) {
      empty.classList.remove('hidden');
      empty.textContent = explainError(error.message);
    }
  }
}

async function refreshGitHubAccount() {
  if (hostedRelay()) return;
  if (githubAccountOptedOut()) {
    state.githubAccount = { connected: false, login: '' };
    showGitHubAccount(false, '');
    return;
  }
  try {
    const account = await api('connectors/github/account');
    const login = String(account?.login || '');
    state.githubAccount = account?.connected && login ? { connected: true, login } : { connected: false, login: '' };
  } catch {
    state.githubAccount = state.githubAccount?.connected ? state.githubAccount : { connected: false, login: '' };
  }
  showGitHubAccount(Boolean(state.githubAccount?.connected), state.githubAccount?.login || '');
}

async function connectGitHubAccount() {
  localStorage.removeItem('eqgc.github.account');
  try {
    const result = await api('connectors/github/connect', {});
    if (result?.pending) toast('The browser opened. Approve GitHub, then click Connect to Github again. 瀏覽器已打開。批准後再按一次 Connect to Github。');
  } catch (error) {
    toast(explainError(error.message));
  }
  await refreshGitHubAccount();
}

function disconnectGitHubAccount() {
  localStorage.setItem('eqgc.github.account', 'off');
  state.githubAccount = { connected: false, login: '' };
  showGitHubAccount(false, '');
  toast('This pane will not use the GitHub login. The PC stays signed in. 這個窗格不再使用 GitHub 登入。這部電腦仍然保持登入。');
}

function refreshGitHubMcpStatus() {
  if (!githubAccountOptedOut() && state.githubAccount?.connected) {
    showGitHubAccount(true, state.githubAccount.login || '');
    return;
  }
  showGitHubAccount(false, '');
}

function restoreGitHubToken() {
  try {
    const saved = JSON.parse(localStorage.getItem('eqgc.github.v1') || '{}');
    if (typeof saved.token === 'string' && saved.token) {
      $('githubToken').value = saved.token;
      $('githubRemember').checked = true;
    }
  } catch { /* storage optional */ }
  refreshGitHubMcpStatus();
}

function saveGitHubMcpToken() {
  const token = $('githubMcpToken').value.trim();
  if ($('githubMcpRemember').checked) {
    if (!token) return toast('Enter a GitHub token.');
    localStorage.setItem('eqgc.github.v1', JSON.stringify({ token }));
    $('githubToken').value = token;
    $('githubRemember').checked = true;
  } else {
    localStorage.removeItem('eqgc.github.v1');
  }
  refreshGitHubMcpStatus();
  toast($('githubMcpRemember').checked ? 'GitHub token saved in this browser.' : 'GitHub token kept for this open pane only.');
}

const GIT_STORE = 'eqgc.git.v1';

function readGitWorkspace() {
  try {
    const saved = JSON.parse(localStorage.getItem(GIT_STORE) || '{}');
    return {
      repo: typeof saved.repo === 'string' ? saved.repo : '',
      file: typeof saved.file === 'string' ? saved.file : ''
    };
  } catch { return { repo: '', file: '' }; }
}

function writeGitWorkspace(next) {
  const current = readGitWorkspace();
  const saved = { repo: next.repo ?? current.repo, file: next.file ?? current.file };
  localStorage.setItem(GIT_STORE, JSON.stringify(saved));
  refreshGitWorkspaceStatus();
  return saved;
}

function refreshGitWorkspaceStatus() {
  const saved = readGitWorkspace();
  if ($('gitRepo') && saved.repo && !$('gitRepo').value) $('gitRepo').value = saved.repo;
  const label = saved.file ? `${saved.file}` : 'No file open';
  setSourceStatus('gitWorkspaceStatus', saved.repo ? `${label} · local Git` : label, Boolean(saved.file));
}

function saveGitRepoFolder() {
  const repo = $('gitRepo').value.trim();
  if (!repo) return toast('Enter the local repo folder.');
  writeGitWorkspace({ repo });
  toast('Repo folder saved on this PC.');
}

function restoreGitWorkspace() {
  const saved = readGitWorkspace();
  if ($('gitRepo') && saved.repo) $('gitRepo').value = saved.repo;
  refreshGitWorkspaceStatus();
}

async function readWordBody() {
  return Word.run(async context => {
    const body = context.document.body;
    body.load('text');
    await context.sync();
    return body.text || '';
  });
}

async function replaceWordBody(text) {
  await Word.run(async context => {
    context.document.body.insertText(text, 'Replace');
    await context.sync();
  });
}

async function appendWordAnswer(text) {
  const body = String(text || '').trim();
  if (!body || state.host !== 'Word') return;
  await Word.run(async context => {
    context.document.body.insertText(`\n\nQwen\n${body}\n`, 'End');
    await context.sync();
  });
}

async function runGitWorkspace(prompt) {
  const command = AgentCore.parseGitCommand(prompt);
  if (!command) return false;
  renderMessage('user', prompt.trim());
  state.busy = true;
  $('send').textContent = '■';
  try {
    if (hostedRelay()) {
      renderMessage('assistant', 'Opening a code file uses the local relay on this PC. 開啟程式檔要用這部電腦的本機 relay。');
      return true;
    }
    if (state.host !== 'Word') {
      renderMessage('assistant', 'Open the code file in Word. Excel stays on the workbook. 請在 Word 開啟程式檔。Excel 仍是活頁簿。');
      return true;
    }
    const saved = readGitWorkspace();
    const repo = ($('gitRepo')?.value || saved.repo || '').trim();
    const file = command.path || saved.file;
    if (command.rejected) {
      renderMessage('assistant', 'git push updates the current branch and does not use --force. Type git push. git push 會更新目前分支，不會使用 --force。請輸入 git push。');
      return true;
    }
    if (command.action === 'push') {
      if (!repo) {
        renderMessage('assistant', 'Save the repo folder under MCP first. 請先在 MCP 儲存 repo 資料夾。');
        return true;
      }
      const pushed = await api('git/push', { repo });
      renderMessage('assistant', `Pushed the current branch. ${pushed.summary || ''} 已 push 目前分支。`);
      return true;
    }
    if (!file) {
      renderMessage('assistant', 'Open one file first. Example: open src/app.js. 請先開一個檔，例如 open src/app.js。');
      return true;
    }
    if (command.action === 'commit' || command.action === 'commit-push') {
      const content = await readWordBody();
      const result = await api('git/commit', { repo, path: file, content, message: command.message });
      writeGitWorkspace({ repo: result.repo || repo, file: result.path || file });
      if (command.action === 'commit') {
        renderMessage('assistant', result.unchanged
          ? `No change to commit in ${result.path}. ${result.path} 沒有變更。`
          : `Committed ${result.path} as ${result.commit}. 已提交 ${result.path}，版本 ${result.commit}。`);
        return true;
      }
      try {
        const pushed = await api('git/push', { repo: result.repo || repo });
        const committed = result.unchanged
          ? `No change to commit in ${result.path}.`
          : `Committed ${result.path} as ${result.commit}.`;
        renderMessage('assistant', `${committed} Pushed. ${pushed.summary || ''} ${result.unchanged ? `${result.path} 沒有變更。` : `已提交 ${result.path}，版本 ${result.commit}。`} 已 push。`);
      } catch (error) {
        const committed = result.unchanged
          ? `No change to commit in ${result.path}.`
          : `Committed ${result.path} as ${result.commit}.`;
        renderMessage('assistant', `${committed} Push did not run. ${explainError(error.message)} 已處理提交。push 沒有執行。`);
      }
      return true;
    }
    const result = command.action === 'pull'
      ? await api('git/pull', { repo, path: file })
      : await api('git/open', { repo, path: file });
    await replaceWordBody(result.text);
    writeGitWorkspace({ repo: result.repo || repo, file: result.path || file });
    renderMessage('assistant', command.action === 'pull'
      ? `Pulled and opened ${result.path}. ${result.summary || ''} 已 pull 並打開 ${result.path}。`
      : `Opened ${result.path} from the repo. 已從 repo 打開 ${result.path}。`);
    return true;
  } catch (error) {
    renderMessage('assistant', explainError(error.message));
    return true;
  } finally {
    resetComposer();
  }
}

function clearGitHubMcpToken() {
  localStorage.removeItem('eqgc.github.v1');
  $('githubMcpToken').value = '';
  $('githubToken').value = '';
  $('githubRemember').checked = false;
  refreshGitHubMcpStatus();
  toast('GitHub token cleared.');
}

async function installGitHubSkill() {
  const repo = $('githubRepo').value.trim();
  const token = $('githubToken').value.trim();
  if (!repo) return toast('Enter owner/repo or a GitHub link to SKILL.md.');
  try {
    const result = await api('skills/github', { repo, token, replace: $('githubReplace').checked });
    if ($('githubRemember').checked) localStorage.setItem('eqgc.github.v1', JSON.stringify({ token }));
    else localStorage.removeItem('eqgc.github.v1');
    refreshGitHubMcpStatus();
    await refreshSkills();
    toast(result.replaced ? `Skill updated: ${result.id}` : `Skill installed: ${result.id}`);
  } catch (error) { toast(explainError(error.message)); }
}

async function saveCurrentSkill() {
  const id = $('skillId').value.trim();
  const description = $('skillDescription').value.trim();
  const task = state.lastPrompt || $('prompt').value.trim();
  if (!id) return toast('Enter a new skill name (lowercase letters, digits, hyphens).');
  if (!description) return toast('Describe when the skill should be used.');
  if (!task) return toast('Send a task first, then save it as a skill.');
  const body = `# ${id}\n\n## Trigger\n\n${task}\n\n## Procedure applied\n\n${state.lastAnswer || '(no answer recorded yet)'}\n\n## Guardrails\n\nRead-only unless the user approves edits. Cite exact sheets and cells. State coverage and limitations. Never overwrite source data.`;
  try {
    const result = await api('skills/save', { id, description, body });
    $('skillId').value = ''; $('skillDescription').value = '';
    await refreshSkills();
    toast(`Skill saved: ${result.path}`);
  } catch (error) { toast(explainError(error.message)); }
}

function renderCloudResults(items) {
  const host = $('sourceResults');
  host.replaceChildren();
  if (!items.length) {
    const empty = document.createElement('p'); empty.className = 'hint'; empty.textContent = 'No matching results.'; host.append(empty); return;
  }
  items.forEach(item => {
    const row = document.createElement('article'); row.className = 'source-result';
    const title = document.createElement('strong'); title.textContent = item.name || 'Untitled'; row.append(title);
    if (item.detail) { const detail = document.createElement('small'); detail.textContent = item.detail; row.append(detail); }
    if (item.snippet) { const snippet = document.createElement('p'); snippet.textContent = item.snippet; row.append(snippet); }
    const add = document.createElement('button'); add.type = 'button'; add.className = 'primary'; add.textContent = 'Review and add to chat';
    add.addEventListener('click', async () => {
      add.disabled = true; add.textContent = 'Reading…';
      try {
        const result = await api('connectors/read', { source: 'onedrive', item: { id: item.id, name: item.name } });
        addSourceAttachment(result.name || item.name, result.kind || item.kind || 'OneDrive', result.text || '');
        add.textContent = 'Added';
      } catch (error) { add.disabled = false; add.textContent = 'Review and add to chat'; toast(explainError(error.message)); }
    });
    row.append(add); host.append(row);
  });
}

async function searchCloudSource() {
  const query = $('sourceQuery').value.trim();
  if (!query) { toast('Enter a search query first.'); return; }
  $('searchSources').disabled = true;
  $('searchSources').textContent = 'Searching…';
  $('sourceResults').replaceChildren();
  try {
    const result = await api('connectors/search', { source: 'onedrive', query });
    renderCloudResults(Array.isArray(result.items) ? result.items : []);
  } catch (error) {
    const message = document.createElement('p'); message.className = 'hint'; message.textContent = error.message;
    $('sourceResults').replaceChildren(message);
  } finally {
    $('searchSources').disabled = false;
    $('searchSources').textContent = 'Search';
  }
}

function addSourceAttachment(name, kind, sourceText) {
  if (state.attachments.length >= ATTACHMENT_MAX_FILES) { toast(`Attach up to ${ATTACHMENT_MAX_FILES} files per chat.`); return; }
  const text = String(sourceText || '').trim();
  if (!text) { toast('The selected source returned no readable text.'); return; }
  const clipped = trimContext(annotateEvidenceLines(text), ATTACHMENT_FILE_CAP);
  state.attachments.push({ name: String(name || 'Cloud source').slice(0, 240), kind: String(kind || 'Cloud source').slice(0, 60), size: 0, text: clipped.text, chars: clipped.text.length, truncated: clipped.truncated });
  renderAttachments();
  refreshContextCard();
  toast('Selected source text is ready; it will be sent with your next question.');
}

function setRelayDot(mode) {
  const dot = $('statusDot');
  if (!dot) return;
  dot.classList.toggle('warn', mode === 'warn');
  dot.classList.toggle('offline', mode === 'offline');
}

async function loadLimited(context, range, nullable) {
  if (nullable) {
    range.load('isNullObject');
    await context.sync();
    if (range.isNullObject) return null;
  }
  range.load(['rowCount', 'columnCount', 'address']);
  await context.sync();
  if (range.rowCount === 0 || range.columnCount === 0) return null;
  let view = range;
  let gridTruncated = false;
  const fullRows = range.rowCount;
  const fullCols = range.columnCount;
  if (Number.isFinite(fullRows) && Number.isFinite(fullCols) && (fullRows > MAX_CONTEXT_ROWS || fullCols > MAX_CONTEXT_COLS)) {
    const rows = Math.min(fullRows, MAX_CONTEXT_ROWS);
    const cols = Math.min(fullCols, MAX_CONTEXT_COLS);
    view = range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    gridTruncated = true;
  }
  view.load(['address', 'text', 'formulas', 'rowCount', 'columnCount']);
  await context.sync();
  return { address: view.address, text: view.text, formulas: view.formulas, rowCount: view.rowCount ?? fullRows, columnCount: view.columnCount ?? fullCols, gridTruncated, fullRows, fullCols };
}

async function readSelection() {
  return Excel.run(async context => loadLimited(context, context.workbook.getSelectedRange(), false));
}

async function readUsedRange() {
  return Excel.run(async context => loadLimited(context, context.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject(), true));
}

function flatCell(value) {
  return String(value ?? '').replace(/[\t\r\n]+/g, ' ');
}

function annotateRows(rows, formulas, address = '') {
  const range = String(address || '').split('!').pop().replace(/\$/g, '').split(':')[0];
  const origin = range.match(/^([A-Z]{1,3})([1-9]\d{0,6})$/i);
  const firstColumn = origin ? columnIndex(origin[1]) : 0;
  const firstRow = origin ? Number(origin[2]) : 0;
  const data = (rows || []).map((row, i) => {
    const cells = (row || []).map((value, j) => {
      const formula = formulas?.[i]?.[j];
      const shown = flatCell(value);
      return formula && String(formula).startsWith('=') ? `${shown} [${flatCell(formula)}]` : shown;
    }).join('\t');
    return origin ? `Row ${firstRow + i}:\t${cells}` : cells;
  });
  if (origin && rows?.length) data.unshift(`Columns: ${(rows[0] || []).map((_, i) => indexToColumn(firstColumn + i)).join('\t')}`);
  return data.join('\n');
}

function annotateEvidenceLines(text) {
  return String(text || '').replace(/\r\n?/g, '\n').split('\n')
    .map((line, index) => `[L${String(index + 1).padStart(4, '0')}] ${line}`).join('\n');
}

function trimContext(text, limit) {
  if (text.length <= limit) return { text, truncated: false };
  const marker = `\n[...truncated: ${text.length - limit} characters omitted; select a smaller range or attach fewer files]`;
  if (limit <= marker.length) return { text: text.slice(0, Math.max(0, limit)), truncated: true };
  return { text: text.slice(0, Math.max(0, limit - marker.length)) + marker, truncated: true };
}

function quotedPassages(question) {
  const found = [];
  for (const match of String(question || '').matchAll(/["“「]([^"”」\n]{4,160})["”」]/g)) {
    const value = match[1].trim();
    if (value) found.push(value);
  }
  return found;
}

function focusWordExcerpt(raw, needles, limit) {
  const annotated = annotateEvidenceLines(raw);
  const needle = needles.find(item => raw.toLowerCase().includes(item.toLowerCase()));
  if (!needle || annotated.length <= limit) return trimContext(annotated, limit);
  const at = raw.toLowerCase().indexOf(needle.toLowerCase());
  const lineIndex = raw.slice(0, at).split('\n').length - 1;
  const lines = annotated.split('\n');
  const startLine = Math.max(0, lineIndex - 2);
  const lead = startLine > 0 ? '[...earlier document text omitted; this excerpt contains the requested passage]\n' : '';
  return trimContext(lead + lines.slice(startLine).join('\n'), limit);
}

function wordContextSource(bodyText, selectionText) {
  const body = String(bodyText || '').replace(/\r\n?/g, '\n').trim();
  const selected = String(selectionText || '').replace(/\r\n?/g, '\n').trim();
  const text = body || selected;
  if (!text) return null;
  return { kind: 'document', text };
}

async function buildWordContext(limit = CONTEXT_CAP, question = '') {
  const needles = quotedPassages(question);
  const source = await Word.run(async context => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    body.load('text');
    if (paragraphs) paragraphs.load('items/text');
    await context.sync();
    const fromBody = String(body.text || '');
    const fromParagraphs = (paragraphs?.items || []).map(item => item.text || '').join('\n');
    return wordContextSource(fromParagraphs.length > fromBody.length ? fromParagraphs : fromBody, '');
  });
  if (!source) return null;
  const clipped = focusWordExcerpt(source.text, needles, limit);
  const note = 'This is the Word document text, the same as an Excel sheet.\n';
  const text = note + clipped.text;
  return { header: '[Word document]', text, address: 'Document', kind: 'document', truncated: clipped.truncated, chars: text.length };
}

async function buildContext(limit = CONTEXT_CAP, question = '') {
  if (!state.officeReady || !$('includeContext').checked) return null;
  if (state.host === 'Word') return buildWordContext(limit, question);
  let source = null, kind = 'selection';
  if (!state.forceSheet) {
    const selection = await readSelection();
    if (selection && selection.rowCount * selection.columnCount > 1) source = selection;
  }
  if (!source) { source = await readUsedRange(); kind = 'used range'; }
  if (!source) return null;
  let text = annotateRows(source.text, source.formulas, source.address);
  if (!text) return null;
  let truncated = false;
  if (source.gridTruncated) {
    const rows = Math.min(source.fullRows || MAX_CONTEXT_ROWS, MAX_CONTEXT_ROWS);
    const cols = Math.min(source.fullCols || MAX_CONTEXT_COLS, MAX_CONTEXT_COLS);
    text += `\n[...truncated: read ${rows} of ${source.fullRows} rows and ${cols} of ${source.fullCols} columns — select a smaller range for the rest]`;
    truncated = true;
  }
  const clipped = trimContext(text, limit);
  text = clipped.text;
  truncated ||= clipped.truncated;
  return { header: `[Excel ${kind} ${source.address}]`, text, address: source.address, kind, truncated, chars: text.length };
}

function attachmentContext() {
  let out = '';
  for (const file of state.attachments) {
    const label = file.name.replace(/[\[\]\r\n]/g, '_');
    const header = `\n\n[Attached file START: ${label} · ${file.kind}]\n`;
    const footer = '\n[Attached file END]';
    const reserve = state.attachments.length > 1 ? 110 : 0;
    const room = ATTACHMENT_CAP - reserve - out.length - header.length - footer.length;
    if (room <= 0) break;
    let text = file.text;
    if (text.length > room) {
      const marker = `\n[truncated: ${text.length - room} characters omitted from this file]`;
      text = room > marker.length ? text.slice(0, room - marker.length) + marker : marker.slice(0, room);
    }
    out += header + text + footer;
  }
  if (state.attachments.length && state.attachments.some(file => !out.includes(`[Attached file START: ${file.name.replace(/[\[\]\r\n]/g, '_')} · ${file.kind}]`))) out = out.slice(0, ATTACHMENT_CAP - 110) + '\n[Other selected files omitted at the attachment context limit.]';
  return out;
}

function collectFinanceSources(ctx, attachmentText) {
  const sources = [];
  if (ctx?.address) {
    if (ctx.kind === 'metadata') sources.push('Workbook metadata only · no cell values were read');
    else if (state.host === 'Word') {
      const lines = [...String(ctx.text || '').matchAll(/^\[L(\d+)\]/gm)].map(match => Number(match[1]));
      const lineRef = lines.length ? ` · L${String(lines[0]).padStart(4, '0')}-L${String(lines.at(-1)).padStart(4, '0')}` : '';
      sources.push(`Word ${ctx.address}${lineRef}${ctx.truncated ? ' · truncated' : ''}`);
    } else sources.push(`Excel ${ctx.address}${ctx.truncated ? ' · truncated' : ''}`);
  }
  for (const file of state.attachments) {
    const label = file.name.replace(/[\[\]\r\n]/g, '_');
    const marker = `[Attached file START: ${label} · ${file.kind}]\n`;
    const start = attachmentText.indexOf(marker);
    if (start < 0) continue;
    const contentStart = start + marker.length;
    const end = attachmentText.indexOf('\n[Attached file END]', contentStart);
    const excerpt = attachmentText.slice(contentStart, end < 0 ? attachmentText.length : end);
    const lineNumbers = [...excerpt.matchAll(/^\[L(\d+)\]/gm)].map(match => Number(match[1]));
    const lineRef = lineNumbers.length
      ? `L${String(lineNumbers[0]).padStart(4, '0')}-L${String(lineNumbers.at(-1)).padStart(4, '0')}`
      : 'included text';
    sources.push(`${file.name} · ${lineRef}${file.truncated ? ' · source truncated' : ''}`);
  }
  return [...new Set(sources)];
}

function a1Bounds(range) {
  const pieces = String(range || '').replace(/\$/g, '').split(':');
  const parse = cell => {
    const match = cell.match(/^([A-Z]{1,3})([1-9]\d{0,6})$/i);
    return match ? { column: columnIndex(match[1]), row: Number(match[2]) } : null;
  };
  const first = parse(pieces[0] || '');
  const last = parse(pieces[1] || pieces[0] || '');
  return first && last ? { top: first.row, left: first.column, bottom: last.row, right: last.column } : null;
}

function financeCitationCheck(answer, sources) {
  const citations = [...String(answer || '').matchAll(/\[([^\]\r\n]+)\]/g)]
    .map(match => match[1].trim())
    .filter(value => /!\$?[A-Z]{1,3}\$?\d/i.test(value) || /:L\d+/i.test(value));
  return [...new Set(citations)].map(citation => {
    const excel = citation.match(/^(.+)!([^!]+)$/);
    if (excel) {
      const citationSheet = normalizeSheet(excel[1]);
      const citationBounds = a1Bounds(excel[2]);
      const matched = citationBounds && sources.some(source => {
        const address = source.replace(/^Excel\s+/i, '').split(' · ')[0];
        const parts = address.split('!');
        if (parts.length < 2 || normalizeSheet(parts.at(-2))?.toLowerCase() !== citationSheet?.toLowerCase()) return false;
        const bounds = a1Bounds(parts.at(-1));
        return bounds && citationBounds.top >= bounds.top && citationBounds.left >= bounds.left
          && citationBounds.bottom <= bounds.bottom && citationBounds.right <= bounds.right;
      });
      return { citation, matched: Boolean(matched) };
    }
    const line = citation.match(/^(.+):L(\d+)(?:-L(\d+))?$/i);
    if (line) {
      const citedName = line[1].replace(/^Word\s+/i, '').toLowerCase();
      const citedStart = Number(line[2]);
      const citedEnd = Number(line[3] || line[2]);
      const matched = sources.some(source => {
        const ref = source.match(/^(.*) · L(\d+)-L(\d+)/);
        if (!ref || ref[1].replace(/^Word\s+/i, '').toLowerCase() !== citedName) return false;
        return citedStart >= Number(ref[2]) && citedEnd <= Number(ref[3]);
      });
      return { citation, matched: Boolean(matched) };
    }
    return { citation, matched: false };
  });
}

function renderAttachments() {
  const host = $('attachmentList');
  host.replaceChildren();
  state.attachments.forEach((file, index) => {
    const row = document.createElement('div'); row.className = 'attachment-item';
    const name = document.createElement('span'); name.textContent = `${file.name} · ${file.chars.toLocaleString()} chars${file.truncated ? ' · truncated' : ''}`;
    const remove = document.createElement('button'); remove.type = 'button'; remove.className = 'text-button'; remove.textContent = 'Remove'; remove.setAttribute('aria-label', `Remove ${file.name}`);
    remove.addEventListener('click', () => { state.attachments.splice(index, 1); renderAttachments(); refreshContextCard(); });
    row.append(name, remove); host.append(row);
  });
  for (const message of state.attachmentErrors || []) {
    const note = document.createElement('div'); note.className = 'edit-note'; note.textContent = message; host.append(note);
  }
  const chips = $('attachmentChips');
  chips.replaceChildren();
  state.attachments.forEach((file, index) => {
    const chip = document.createElement('div'); chip.className = 'attachment-chip';
    const icon = document.createElement('span'); icon.textContent = ({ PDF: '📄', XLSX: '📊', XLSM: '📊', PNG: '🖼', JPG: '🖼', JPEG: '🖼', WEBP: '🖼' })[file.kind] || '📃';
    const name = document.createElement('span'); name.textContent = file.name; name.title = file.name;
    const remove = document.createElement('button'); remove.type = 'button'; remove.textContent = '×'; remove.setAttribute('aria-label', `Remove ${file.name}`);
    remove.addEventListener('click', () => { state.attachments.splice(index, 1); renderAttachments(); refreshContextCard(); });
    chip.append(icon, name, remove); chips.append(chip);
  });
  const preview = $('attachmentPreview');
  preview.classList.toggle('hidden', !state.attachments.length);
  $('attachmentTextPreview').textContent = attachmentContext().trim();
}

function decodeAttachment(bytes, extension) {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return new TextDecoder('utf-8').decode(bytes.subarray(3));
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  const ascii = new TextDecoder('latin1').decode(bytes.subarray(0, Math.min(bytes.length, 8192)));
  const charset = extension === 'html' || extension === 'htm' ? ascii.match(/<meta\b[^>]*charset\s*=\s*["']?([\w-]+)/i)?.[1] : null;
  if (charset) { try { return new TextDecoder(charset).decode(bytes); } catch { /* fall through to UTF-8 */ } }
  try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
  catch { try { return new TextDecoder('big5').decode(bytes); } catch { return new TextDecoder('windows-1252').decode(bytes); } }
}

function htmlAttachmentText(source) {
  // Strip executable and resource-bearing content before parsing. Never insert the uploaded HTML into the live pane.
  const safe = source
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<script\b[^>]*>[\s\S]*?(?:<\/script\s*>|$)/gi, ' ')
    .replace(/<(script|style|iframe|object|embed|template|svg|canvas|noscript)\b[^>]*>[\s\S]*?<\/\s*\1\s*>/gi, ' ')
    .replace(/<(?:iframe|object|embed|link|img|source|video|audio|meta|base|form|input)\b[^>]*>/gi, ' ')
    .replace(/\s+(?:src|srcset|href|action|formaction|poster|srcdoc|style|on[\w:-]+)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, ' ');
  if (typeof DOMParser === 'undefined') return safe.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&quot;/gi, '"').replace(/&#39;/gi, "'").replace(/\s+/g, ' ').trim();
  const doc = new DOMParser().parseFromString(safe, 'text/html');
  doc.querySelectorAll('script,style,iframe,object,embed,template,svg,canvas,noscript,link,meta,base,form,input,button').forEach(node => node.remove());
  doc.querySelectorAll('[hidden],[aria-hidden="true"],[style*="display:none" i],[style*="visibility:hidden" i]').forEach(node => node.remove());
  const tables = [...doc.querySelectorAll('table')].filter(table => !table.parentElement?.closest('table'));
  for (const table of tables) {
    const lines = [...table.querySelectorAll('tr')].map(row => [...row.querySelectorAll('th,td')].filter(cell => cell.closest('table') === table).map(cell => cell.textContent.replace(/\s+/g, ' ').trim()).join(' | ')).filter(Boolean);
    const caption = table.querySelector('caption')?.textContent.replace(/\s+/g, ' ').trim();
    if (caption) lines.unshift(caption);
    table.replaceWith(doc.createTextNode(`\n${lines.join('\n')}\n`));
  }
  doc.querySelectorAll('h1,h2,h3,h4,h5,h6').forEach(node => {
    const level = Number(node.tagName.slice(1)) || 1;
    node.insertBefore(doc.createTextNode(`${'#'.repeat(level)} `), node.firstChild);
  });
  doc.querySelectorAll('br').forEach(node => node.replaceWith(doc.createTextNode('\n')));
  doc.querySelectorAll('h1,h2,h3,h4,h5,h6,p,div,section,article,main,header,footer,li,tr').forEach(node => { node.before(doc.createTextNode('\n')); node.after(doc.createTextNode('\n')); });
  const title = (doc.querySelector('title')?.textContent || '').replace(/\s+/g, ' ').trim();
  const body = (doc.body?.textContent || doc.documentElement?.textContent || '').replace(/[\t\u00a0 ]+/g, ' ').replace(/ *\n */g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  return `${title ? `Title: ${title}\n\n` : ''}${body}`.trim();
}

async function readAttachment(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const local = new Set(['html', 'htm', 'txt', 'md', 'csv', 'tsv', 'json', 'xml']);
  const relayed = new Set(['pdf', 'docx', 'xlsx', 'xlsm', 'png', 'jpg', 'jpeg', 'webp']);
  if (!local.has(ext) && !relayed.has(ext)) throw new Error(`${file.name}: supported files are HTML, TXT, MD, CSV, TSV, JSON, XML, PDF, DOCX, XLSX, XLSM, PNG, JPG and WEBP.`);
  const cap = relayed.has(ext) ? ATTACHMENT_RELAY_MAX_BYTES : ATTACHMENT_MAX_BYTES;
  if (file.size > cap) throw new Error(`${file.name}: file is over the ${Math.round(cap / 1024 / 1024)} MB limit.`);
  let extracted;
  if (relayed.has(ext)) {
    // Binary formats are extracted by the local relay; the pane never parses them.
    if (hostedRelay()) throw new Error(`${file.name}: PDF, Word, and Excel extraction needs the local relay on this PC.`);
    const bytes = new Uint8Array(await file.arrayBuffer());
    let binary = '';
    for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    const result = await api('extract', { name: file.name, data: btoa(binary) });
    extracted = annotateEvidenceLines(String(result.text || ''));
  } else {
    const text = decodeAttachment(new Uint8Array(await file.arrayBuffer()), ext);
    extracted = annotateEvidenceLines(ext === 'html' || ext === 'htm' ? htmlAttachmentText(text) : text.replace(/\u0000/g, ''));
  }
  if (!extracted.trim()) throw new Error(`${file.name}: no readable text was found.`);
  const clipped = trimContext(extracted, ATTACHMENT_FILE_CAP);
  return { name: file.name, kind: ext.toUpperCase(), size: file.size, text: clipped.text, chars: clipped.text.length, truncated: clipped.truncated };
}

async function addAttachments(files) {
  const incoming = [...files];
  if (state.attachments.length + incoming.length > ATTACHMENT_MAX_FILES) { toast(`Attach up to ${ATTACHMENT_MAX_FILES} files per chat.`); return; }
  const added = [];
  state.attachmentErrors = [];
  for (const file of incoming) {
    try { added.push(await readAttachment(file)); }
    catch (error) {
      toast(explainError(error.message));
      state.attachmentErrors.push(explainError(error.message));
    }
  }
  state.attachments.push(...added);
  renderAttachments();
  refreshContextCard();
  if (added.length) toast(`${added.length} file(s) ready. They will be sent with your next question.`);
}

async function refreshContextCard() {
  renderAttachments();
  try {
    const ctx = state.officeReady && $('includeContext').checked ? await contextForMode(operatingMode(), Math.min(CONTEXT_CAP, Math.max(1, CONTEXT_CAP - attachmentContext().length))) : null;
    const fileNames = state.attachments.map(file => file.name);
    const fileNote = fileNames.length ? ` · ${fileNames.length} file(s): ${fileNames.join(', ')}` : '';
    if (!state.officeReady) {
      $('selectionName').textContent = fileNames.length ? `${fileNames.length} attached file(s) ready` : 'Open Excel or Word to read context';
      $('selectionPreview').textContent = fileNames.length ? 'Selected file or cloud source text will be sent with your next question.' : 'Choose a local HTML/text file or review a cloud source.';
      $('contextHint').textContent = `Up to ${ATTACHMENT_CAP.toLocaleString()} attachment characters per question${fileNote}`;
      return;
    }
    if (!$('includeContext').checked) {
      const fileKind = state.host === 'Word' ? 'Document' : 'Workbook';
      $('selectionName').textContent = fileNames.length ? `${fileNames.length} attached file(s) ready` : `${fileKind} sharing off`;
      $('selectionPreview').textContent = `${fileKind} data will not be sent to AI. Selected attachment text is controlled separately.`;
      $('contextHint').textContent = `${fileKind} sharing off${fileNote}`;
      return;
    }
    $('contextEyebrow').textContent = operatingMode() === 'agent' ? 'WORKBOOK ACCESS · AGENT AUTO' : 'WORKBOOK CONTEXT · AUTO';
    if (!ctx) {
      $('selectionName').textContent = fileNames.length ? `${fileNames.length} attached file(s) ready` : 'No workbook data';
      $('selectionPreview').textContent = fileNames.length ? 'Selected file text will be sent with your next question.' : '(The active sheet and selection are empty)';
      $('contextHint').textContent = fileNames.length ? `No workbook data${fileNote}` : 'Nothing will be sent';
      return;
    }
    $('selectionName').textContent = ctx.address;
    $('selectionPreview').textContent = ctx.text;
    $('contextHint').textContent = `Next question sends ${ctx.kind} · ${ctx.chars.toLocaleString()} chars${ctx.truncated ? ' · truncated' : ''}${ctx.kind === 'metadata' ? ' · Agent chooses ranges as needed' : ' · formulas included'}${fileNote}`;
  } catch (error) { toast(`Could not read workbook context: ${error.message}`); }
}

function watchSelection() {
  const doc = Office.context?.document;
  if (!doc || typeof doc.addHandlerAsync !== 'function' || !Office.EventType?.DocumentSelectionChanged) return;
  doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    clearTimeout(watchSelection.timer);
    watchSelection.timer = setTimeout(refreshContextCard, 400);
  });
}

function renderGitHubWrite(proposal) {
  const conversation = $('conversation');
  conversation.querySelector('.welcome')?.remove();
  const card = document.createElement('div');
  card.className = 'message assistant';
  const label = document.createElement('div');
  label.className = 'message-label';
  label.textContent = 'GITHUB WRITE';
  const body = document.createElement('div');
  const title = document.createElement('strong');
  title.textContent = `${proposal.repo} · ${proposal.title}`;
  body.append(title);
  if (proposal.body) {
    const why = document.createElement('p');
    why.textContent = proposal.body;
    body.append(why);
  }
  for (const file of proposal.files) {
    const name = document.createElement('div');
    name.textContent = file.path;
    const pre = document.createElement('pre');
    pre.className = 'github-draft';
    pre.textContent = file.content;
    body.append(name, pre);
  }
  const status = document.createElement('p');
  status.className = 'hint';
  status.textContent = 'Nothing has been sent to GitHub. Write to GitHub opens a pull request on a new branch. The default branch needs its own click. 尚未送出。Write to GitHub 會開新分支和 pull request。預設分支要另外再按一次。';
  const actions = document.createElement('div');
  actions.className = 'github-write';
  const writeBtn = document.createElement('button');
  writeBtn.type = 'button';
  writeBtn.className = 'primary';
  writeBtn.textContent = 'Write to GitHub';
  const defaultBtn = document.createElement('button');
  defaultBtn.type = 'button';
  defaultBtn.className = 'secondary';
  defaultBtn.textContent = 'Write to default branch';
  const sendWrite = async target => {
    writeBtn.disabled = true;
    defaultBtn.disabled = true;
    try {
      const auth = githubAuthBody();
      if (!auth.token && !auth.useAccount) throw new Error('Connect to Github in MCP first. 請先在 MCP 按 Connect to Github。');
      const result = await api('connectors/github/write', {
        ...auth, repo: proposal.repo, title: proposal.title, body: proposal.body, files: proposal.files, target
      });
      status.textContent = result.prUrl
        ? `Pull request opened: ${result.prUrl}`
        : `Committed ${result.commit} on ${result.base}.`;
      state.audit.push(AgentCore.audit({ kind: 'context', address: result.prUrl || result.commit || proposal.repo, chars: proposal.files.length }));
    } catch (error) {
      status.textContent = explainError(error.message);
      writeBtn.disabled = false;
      defaultBtn.disabled = false;
    }
  };
  writeBtn.addEventListener('click', () => sendWrite('pull-request'));
  defaultBtn.addEventListener('click', () => sendWrite('default'));
  actions.append(writeBtn, defaultBtn);
  body.append(status, actions);
  card.append(label, body);
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

function renderMessage(role, content) {
  const conversation = $('conversation');
  conversation.querySelector('.welcome')?.remove();
  const bubble = document.createElement('div'); bubble.className = `message ${role}`;
  if (role === 'assistant') {
    const label = document.createElement('div'); label.className = 'message-label'; label.textContent = 'QWEN'; bubble.append(label);
  }
  const body = document.createElement('div'); body.textContent = content; bubble.append(body); conversation.append(bubble); conversation.scrollTop = conversation.scrollHeight;
  return bubble;
}

function columnIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function normalizeCell(cell) {
  if (typeof cell !== 'string') return '';
  const match = cell.trim().match(/^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/);
  if (!match) return '';
  const col = match[1].toUpperCase();
  const row = Number(match[2]);
  if (columnIndex(col) > 16384 || row > 1048576) return '';
  return col + String(row);
}

function normalizeSheet(name) {
  if (typeof name !== 'string') return '';
  let sheet = name.trim();
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  if (!sheet || sheet.length > 31 || /[:\\/?*[\]]/.test(sheet) || sheet.startsWith("'") || sheet.endsWith("'")) return '';
  return sheet;
}

function parseEdits(text) {
  const blocks = [...text.matchAll(/```(?:json)?\s*([\s\S]*?)```/gi)];
  for (let i = blocks.length - 1; i >= 0; i--) {
    try {
      const parsed = JSON.parse(blocks[i][1]);
      const edits = parsed?.edits;
      if (!Array.isArray(edits) || edits.length < 1 || edits.length > 200) continue;
      const clean = [];
      let skipped = 0;
      for (const edit of edits) {
        const sheet = normalizeSheet(edit?.sheet);
        const cell = normalizeCell(edit?.cell);
        const value = edit?.value;
        const valueOk = (typeof value === 'string' || (typeof value === 'number' && Number.isFinite(value))) && String(value).length <= 5000;
        if (!sheet || !cell || !valueOk) { skipped += 1; continue; }
        clean.push({ sheet, cell, value });
      }
      if (!clean.length) continue;
      const prose = (text.slice(0, blocks[i].index) + text.slice(blocks[i].index + blocks[i][0].length)).trim();
      return { prose, edits: clean, skipped };
    } catch { /* not an edits block; keep looking */ }
  }
  return { prose: text, edits: null, skipped: 0 };
}

function riskyEdit(value) {
  return typeof value === 'string' && RISKY_FORMULA.test(value.trim());
}

const CHART_TYPES = {
  column: 'ColumnClustered', bar: 'BarClustered', line: 'Line', pie: 'Pie',
  columnstacked: 'ColumnStacked', columnstacked100: 'ColumnStacked100',
  barstacked: 'BarStacked', barstacked100: 'BarStacked100',
  linestacked: 'LineStacked', linemarkers: 'LineMarkers', linemarkersstacked: 'LineMarkersStacked',
  area: 'Area', areastacked: 'AreaStacked', areastacked100: 'AreaStacked100',
  doughnut: 'Doughnut', doughnutexploded: 'DoughnutExploded', pieexploded: 'PieExploded',
  pieofpie: 'PieOfPie', barofpie: 'BarOfPie',
  scatter: 'XYScatter', xyscatter: 'XYScatter', xyscattersmooth: 'XYScatterSmooth',
  xyscatterlines: 'XYScatterLines', xyscattersmoothnomarkers: 'XYScatterSmoothNoMarkers',
  radar: 'Radar', radarmarkers: 'RadarMarkers', radarfilled: 'RadarFilled',
  bubble: 'Bubble', surface: 'Surface', stockhlc: 'StockHLC', stockohlc: 'StockOHLC'
};
const PIVOT_AGGREGATIONS = {
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min',
  countnumbers: 'CountNumbers', product: 'Product', stddev: 'StandardDeviation', variance: 'Variance'
};
const PIVOT_LAYOUTS = { tabular: 'Tabular', compact: 'Compact', outline: 'Outline' };
const BORDER_EDGES = {
  top: 'EdgeTop', bottom: 'EdgeBottom', left: 'EdgeLeft', right: 'EdgeRight',
  insidevertical: 'InsideVertical', insidehorizontal: 'InsideHorizontal',
  diagonaldown: 'DiagonalDown', diagonalup: 'DiagonalUp',
  edgetop: 'EdgeTop', edgebottom: 'EdgeBottom', edgeleft: 'EdgeLeft', edgeright: 'EdgeRight',
  all: 'all', outline: 'outline'
};
const BORDER_STYLES = {
  none: 'None', continuous: 'Continuous', solid: 'Continuous', dash: 'Dash', dashed: 'Dash',
  dashdot: 'DashDot', dashdotdot: 'DashDotDot', dot: 'Dot', dotted: 'Dot', double: 'Double', slantdashdot: 'SlantDashDot'
};
const BORDER_WEIGHTS = { hairline: 'Hairline', thin: 'Thin', medium: 'Medium', thick: 'Thick' };
const CF_OPERATORS = {
  between: 'Between', notbetween: 'NotBetween', equalto: 'EqualTo', equal: 'EqualTo', '==': 'EqualTo',
  notequalto: 'NotEqualTo', notequal: 'NotEqualTo', '!=': 'NotEqualTo',
  greaterthan: 'GreaterThan', '>': 'GreaterThan', lessthan: 'LessThan', '<': 'LessThan',
  greaterthanorequal: 'GreaterThanOrEqual', '>=': 'GreaterThanOrEqual',
  lessthanorequal: 'LessThanOrEqual', '<=': 'LessThanOrEqual'
};
const FILTER_ON = {
  values: 'Values', custom: 'Custom', topitems: 'TopItems', bottomitems: 'BottomItems',
  toppercent: 'TopPercent', bottompercent: 'BottomPercent', cellcolor: 'CellColor', fontcolor: 'FontColor'
};
const WORD_STYLES = { 'heading 1': 'Heading 1', 'heading 2': 'Heading 2', 'heading 3': 'Heading 3', title: 'Title', normal: 'Normal' };

function indexToColumn(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function rangeSize(addr) {
  const parts = String(addr).split(':');
  const start = parts[0].match(/^([A-Z]+)(\d+)$/);
  const end = (parts[1] || parts[0]).match(/^([A-Z]+)(\d+)$/);
  if (!start || !end) return { rows: 1, cols: 1 };
  return { rows: Math.abs(Number(end[2]) - Number(start[2])) + 1, cols: Math.abs(columnIndex(end[1]) - columnIndex(start[1])) + 1 };
}

function normalizeRange(input) {
  if (typeof input !== 'string') return '';
  const parts = input.trim().split(':');
  if (parts.length === 1) return normalizeCell(parts[0]);
  if (parts.length !== 2) return '';
  const start = normalizeCell(parts[0]);
  const end = normalizeCell(parts[1]);
  return start && end ? `${start}:${end}` : '';
}

function normalizeColumns(input) {
  if (typeof input !== 'string') return '';
  const match = input.trim().match(/^\$?([A-Za-z]{1,3})(?::\$?([A-Za-z]{1,3}))?$/);
  if (!match) return '';
  const start = match[1].toUpperCase();
  const end = (match[2] || start).toUpperCase();
  if (columnIndex(start) < 1 || columnIndex(start) > 16384 || columnIndex(end) < 1 || columnIndex(end) > 16384) return '';
  return columnIndex(start) <= columnIndex(end) ? `${start}:${end}` : `${end}:${start}`;
}

function boundedRange(input, limit = 1000) {
  const range = normalizeRange(input);
  if (!range) return '';
  const size = rangeSize(range);
  return size.rows * size.cols <= limit ? range : '';
}

// Any range is walkable: one READ_MAX_CELLS batch per request, resumed with offset.
function pagedRange(addr, offset = 0) {
  const size = rangeSize(addr);
  const skip = Math.min(Math.max(0, Number.isInteger(offset) ? offset : 0), size.rows - 1);
  const rows = Math.min(size.rows - skip, Math.max(1, Math.floor(READ_MAX_CELLS / size.cols)));
  const start = addr.split(':')[0].match(/^([A-Z]+)(\d+)$/);
  const end = (addr.split(':')[1] || addr.split(':')[0]).match(/^([A-Z]+)(\d+)$/);
  const top = Number(start[2]) + skip;
  return { range: `${start[1]}${top}:${end[1]}${top + rows - 1}`, skip, rows, totalRows: size.rows, firstRow: Number(start[2]) };
}

function cellValueOk(value) {
  return (typeof value === 'string' || (typeof value === 'number' && Number.isFinite(value))) && String(value).length <= 5000;
}

function normalizeGrid(raw) {
  if (Array.isArray(raw.values)) {
    if (raw.values.length < 1 || raw.values.length > 30) return null;
    const grid = [];
    for (const row of raw.values) {
      if (!Array.isArray(row) || row.length < 1 || row.length > 15) return null;
      const clean = [];
      for (const value of row) {
        if (value == null) { clean.push(''); continue; }
        if (!cellValueOk(value)) return null;
        clean.push(value);
      }
      grid.push(clean);
    }
    return grid;
  }
  if (!cellValueOk(raw.value)) return null;
  return [[raw.value]];
}

function positiveInt(value, max) {
  const n = Number(value);
  return Number.isInteger(n) && n >= 1 && n <= max ? n : 0;
}

function colorOk(value) {
  return typeof value === 'string' && /^#[0-9A-Fa-f]{6}$/.test(value);
}

function normalizeAction(raw) {
  const op = typeof raw?.op === 'string' ? raw.op : '';
  if (op === 'replaceSelection' || op === 'insertText' || op === 'formatSelection' || op === 'styleSelection') {
    if (op === 'replaceSelection' || op === 'insertText') {
      const text = typeof raw.text === 'string' && raw.text.length <= 8000 ? raw.text : null;
      if (text == null || !text.length) return null;
      if (op === 'insertText' && raw.where != null && !['start', 'end'].includes(raw.where)) return null;
      return { host: 'word', op, text, where: raw.where === 'start' ? 'start' : 'end' };
    }
    if (op === 'formatSelection') {
      if (typeof raw.bold !== 'boolean' && typeof raw.italic !== 'boolean') return null;
      return { host: 'word', op, bold: raw.bold, italic: raw.italic };
    }
    const style = WORD_STYLES[String(raw.style || '').trim().toLowerCase()];
    return style ? { host: 'word', op, style } : null;
  }
  if (op === 'replaceAll' && !raw.sheet) {
    const text = typeof raw.text === 'string' ? raw.text : '';
    const replacement = typeof raw.replacement === 'string' ? raw.replacement : '';
    if (!text || text.length > 500 || replacement.length > 500) return null;
    return { host: 'word', op, text, replacement, matchCase: raw.matchCase === true, matchWholeWord: raw.matchWholeWord === true, bulk: true };
  }
  if (op === 'addSheet') {
    const name = normalizeSheet(raw.name);
    return name ? { host: 'excel', op, name } : null;
  }
  const sheet = normalizeSheet(raw.sheet);
  if (!sheet && op !== 'addSheet') return null;
  if (op === 'renameSheet') {
    const name = normalizeSheet(raw.name);
    return name && name !== sheet ? { host: 'excel', op, sheet, name } : null;
  }
  if (op === 'set') {
    const range = boundedRange(raw.range, 450);
    const grid = normalizeGrid(raw);
    if (!range || !grid) return null;
    const size = rangeSize(range);
    if (grid.length > size.rows || grid.some(row => row.length > size.cols)) return null;
    return { host: 'excel', op, sheet, range, grid };
  }
  if (op === 'format') {
    const range = boundedRange(raw.range, 1000);
    if (!range) return null;
    const action = { host: 'excel', op, sheet, range };
    if (typeof raw.bold === 'boolean') action.bold = raw.bold;
    if (typeof raw.italic === 'boolean') action.italic = raw.italic;
    if (raw.fill != null) { if (!colorOk(raw.fill)) return null; action.fill = raw.fill; }
    if (raw.fontColor != null) { if (!colorOk(raw.fontColor)) return null; action.fontColor = raw.fontColor; }
    if (raw.numberFormat != null) {
      if (typeof raw.numberFormat !== 'string' || raw.numberFormat.length > 40 || /[\r\n]/.test(raw.numberFormat)) return null;
      action.numberFormat = raw.numberFormat;
    }
    if (raw.align != null) {
      const align = { left: 'Left', center: 'Center', right: 'Right' }[String(raw.align).toLowerCase()];
      if (!align) return null;
      action.align = align;
    }
    if (typeof raw.wrap === 'boolean') action.wrap = raw.wrap;
    const keys = ['bold', 'italic', 'fill', 'fontColor', 'numberFormat', 'align', 'wrap'];
    return keys.some(key => action[key] != null) ? action : null;
  }
  if (op === 'clear' || op === 'sort' || op === 'chart' || op === 'table') {
    const range = boundedRange(raw.range, op === 'clear' ? 1000 : 1000);
    if (!range) return null;
    if (op === 'clear') return { host: 'excel', op, sheet, range };
    if (op === 'sort') {
      const start = columnIndex(range.split(':')[0].match(/^([A-Z]+)/)[1]);
      let key = -1;
      if (typeof raw.column === 'number' && raw.column >= 1 && raw.column <= 50) key = raw.column - 1;
      else if (typeof raw.column === 'string' && /^[A-Za-z]{1,3}$/.test(raw.column)) key = columnIndex(raw.column.toUpperCase()) - start;
      if (key < 0 || key >= rangeSize(range).cols) return null;
      return { host: 'excel', op, sheet, range, key, ascending: raw.ascending !== false, hasHeaders: raw.hasHeaders !== false };
    }
    if (op === 'chart') {
      const type = CHART_TYPES[String(raw.type || 'column').toLowerCase()];
      const title = raw.title == null ? '' : String(raw.title);
      if (!type || title.length > 120) return null;
      return { host: 'excel', op, sheet, range, type, title };
    }
    const name = typeof raw.name === 'string' ? raw.name.trim() : '';
    if (!/^[A-Za-z_][A-Za-z0-9_]{0,40}$/.test(name) || /^[A-Za-z]{1,3}[0-9]+$/.test(name)) return null;
    return { host: 'excel', op, sheet, range, name };
  }
  if (op === 'insertRows' || op === 'deleteRows') {
    const row = positiveInt(raw.row, 1048576);
    const count = positiveInt(raw.count || 1, 20);
    if (!row || !count || row + count - 1 > 1048576) return null;
    return { host: 'excel', op, sheet, row, count };
  }
  if (op === 'insertColumns') {
    const column = typeof raw.column === 'string' ? normalizeCell(`${raw.column}1`).replace(/1$/, '') : '';
    const count = positiveInt(raw.count || 1, 10);
    if (!/^[A-Z]{1,3}$/.test(column) || !count || columnIndex(column) + count - 1 > 16384) return null;
    return { host: 'excel', op, sheet, column, count };
  }
  if (op === 'autofit') {
    const range = /[0-9]/.test(String(raw.range || '')) ? boundedRange(raw.range, 2000) : normalizeColumns(raw.range);
    return range ? { host: 'excel', op, sheet, range } : null;
  }
  if (op === 'pivot') {
    // The source range is handed to Excel rather than read into the pane, so it may
    // be far larger than a normal read or write range.
    const range = boundedRange(raw.range, 1000000);
    if (!range) return null;
    const name = typeof raw.name === 'string' ? raw.name.trim() : '';
    if (!name || name.length > 60 || /[\r\n]/.test(name)) return null;
    const destinationSheet = normalizeSheet(raw.destinationSheet || name);
    const destinationCell = normalizeCell(raw.destinationCell || 'A3');
    if (!destinationSheet || !destinationCell) return null;
    const fieldList = value => {
      if (!Array.isArray(value) || value.length > 8) return null;
      const list = value.map(item => typeof item === 'string' ? item.trim() : '');
      return list.every(item => item && item.length <= 100) ? list : null;
    };
    const rows = fieldList(raw.rows || []);
    const columns = fieldList(raw.columns || []);
    const filters = fieldList(raw.filters || []);
    if (!rows || !columns || !filters || rows.length + columns.length < 1) return null;
    if (!Array.isArray(raw.values) || raw.values.length < 1 || raw.values.length > 16) return null;
    const values = [];
    for (const item of raw.values) {
      const field = typeof item?.field === 'string' ? item.field.trim() : '';
      if (!field || field.length > 100) return null;
      const summarizeBy = item.summarizeBy == null ? 'Sum' : PIVOT_AGGREGATIONS[String(item.summarizeBy).toLowerCase()];
      if (!summarizeBy) return null;
      const label = item.name == null ? field : String(item.name).trim();
      if (!label || label.length > 100) return null;
      values.push({ field, summarizeBy, name: label });
    }
    const layout = PIVOT_LAYOUTS[String(raw.layout || 'tabular').toLowerCase()];
    if (!layout) return null;
    return {
      host: 'excel', op, sheet, range, name, destinationSheet, destinationCell,
      rows, columns, filters, values, layout, grandTotals: raw.grandTotals !== false
    };
  }
  if (op === 'filter') {
    const range = boundedRange(raw.range, 100000);
    if (!range) return null;
    if (raw.clear === true) return { host: 'excel', op, sheet, range, clear: true, bulk: true };
    const start = columnIndex(range.split(':')[0].match(/^([A-Z]+)/)[1]);
    let key = -1;
    if (typeof raw.column === 'number' && raw.column >= 1 && raw.column <= 16384) key = raw.column - 1;
    else if (typeof raw.column === 'string' && /^[A-Za-z]{1,3}$/.test(raw.column.trim())) key = columnIndex(raw.column.trim().toUpperCase()) - start;
    if (key < 0 || key >= rangeSize(range).cols) return null;
    const filterOn = FILTER_ON[String(raw.filterOn || 'values').toLowerCase()];
    if (!filterOn) return null;
    const criteria = { filterOn };
    if (filterOn === 'Values') {
      if (!Array.isArray(raw.values) || raw.values.length < 1 || raw.values.length > 200) return null;
      const values = raw.values.map(value => String(value));
      if (values.some(value => !value || value.length > 200)) return null;
      criteria.values = values;
    } else if (filterOn === 'Custom') {
      const c1 = typeof raw.criterion1 === 'string' ? raw.criterion1.trim() : '';
      if (!c1 || c1.length > 100) return null;
      criteria.criterion1 = c1;
      if (raw.criterion2 != null && String(raw.criterion2).trim()) {
        const c2 = String(raw.criterion2).trim();
        if (c2.length > 100) return null;
        criteria.criterion2 = c2;
        criteria.operator = String(raw.operator || 'And').toLowerCase() === 'or' ? 'Or' : 'And';
      }
    } else if (filterOn === 'CellColor' || filterOn === 'FontColor') {
      if (!colorOk(raw.color)) return null;
      criteria.color = raw.color;
    } else {
      const count = positiveInt(raw.count, 1000);
      if (!count) return null;
      criteria.criterion1 = String(count);
    }
    return { host: 'excel', op, sheet, range, key, criteria };
  }
  if (op === 'conditionalFormat') {
    const range = boundedRange(raw.range, 100000);
    if (!range) return null;
    if (raw.clear === true) return { host: 'excel', op, sheet, range, clear: true, bulk: true };
    const operator = CF_OPERATORS[String(raw.operator || '').toLowerCase()];
    if (!operator) return null;
    const formula1 = raw.formula1 == null ? '' : String(raw.formula1).trim();
    if (!formula1 || formula1.length > 100) return null;
    const needsSecond = operator === 'Between' || operator === 'NotBetween';
    const formula2 = raw.formula2 == null ? '' : String(raw.formula2).trim();
    if (needsSecond && (!formula2 || formula2.length > 100)) return null;
    if (!needsSecond && formula2) return null;
    const action = { host: 'excel', op, sheet, range, operator, formula1 };
    if (needsSecond) action.formula2 = formula2;
    if (raw.fill != null) { if (!colorOk(raw.fill)) return null; action.fill = raw.fill; }
    if (raw.fontColor != null) { if (!colorOk(raw.fontColor)) return null; action.fontColor = raw.fontColor; }
    if (typeof raw.bold === 'boolean') action.bold = raw.bold;
    return action.fill || action.fontColor || action.bold != null ? action : null;
  }
  if (op === 'borders') {
    const range = boundedRange(raw.range, 100000);
    if (!range) return null;
    const edge = BORDER_EDGES[String(raw.edge || 'all').toLowerCase()];
    const style = BORDER_STYLES[String(raw.style || 'continuous').toLowerCase()];
    const weight = BORDER_WEIGHTS[String(raw.weight || 'thin').toLowerCase()];
    if (!edge || !style || !weight) return null;
    const action = { host: 'excel', op, sheet, range, edge, style, weight };
    if (raw.color != null) { if (!colorOk(raw.color)) return null; action.color = raw.color; }
    return action;
  }
  if (op === 'freeze') {
    if (raw.unfreeze === true) return { host: 'excel', op, sheet, unfreeze: true };
    const rows = raw.rows == null ? 0 : Number(raw.rows);
    const columns = raw.columns == null ? 0 : Number(raw.columns);
    if (!Number.isInteger(rows) || !Number.isInteger(columns) || rows < 0 || columns < 0 || rows > 1048576 || columns > 16384) return null;
    if (!rows && !columns) return null;
    return { host: 'excel', op, sheet, rows, columns };
  }
  if (op === 'replaceAll') {
    const text = typeof raw.text === 'string' ? raw.text : '';
    const replacement = typeof raw.replacement === 'string' ? raw.replacement : '';
    if (!text || text.length > 500 || replacement.length > 500) return null;
    return { host: 'excel', op, sheet, text, replacement, completeMatch: raw.completeMatch === true, matchCase: raw.matchCase === true, bulk: true };
  }
  if (op === 'dataValidation') {
    const range = boundedRange(raw.range, 100000);
    if (!range) return null;
    if (raw.clear === true) return { host: 'excel', op, sheet, range, clear: true, bulk: true };
    if (!Array.isArray(raw.list) || raw.list.length < 1 || raw.list.length > 200) return null;
    const items = raw.list.map(item => String(item).trim());
    if (items.some(item => !item || item.length > 100 || item.includes(','))) return null;
    return { host: 'excel', op, sheet, range, list: items, ignoreBlanks: raw.ignoreBlanks !== false };
  }
  if (op === 'namedRange') {
    const name = typeof raw.name === 'string' ? raw.name.trim() : '';
    if (!/^[A-Za-z_][A-Za-z0-9_.]{0,60}$/.test(name) || /^[A-Za-z]{1,3}[0-9]+$/.test(name)) return null;
    const range = normalizeRange(raw.range);
    if (!range) return null;
    return { host: 'excel', op, sheet, name, range };
  }
  if (op === 'hide') {
    const target = String(raw.target || '').toLowerCase();
    if (!['rows', 'columns', 'sheet'].includes(target)) return null;
    const hidden = raw.hidden !== false;
    if (target === 'sheet') return { host: 'excel', op, sheet, target, hidden };
    // Whole-row ("5:10") and whole-column ("D:E") addresses are the natural way to hide.
    const text = String(raw.range || '').trim();
    const range = target === 'rows' && /^\d+:\d+$/.test(text)
      ? text
      : /[0-9]/.test(text) ? boundedRange(text, 100000) : normalizeColumns(text);
    if (!range) return null;
    return { host: 'excel', op, sheet, target, range, hidden };
  }
  if (op === 'tableColumn') {
    const table = typeof raw.table === 'string' ? raw.table.trim() : '';
    const name = typeof raw.name === 'string' ? raw.name.trim() : '';
    if (!table || table.length > 60 || !name || name.length > 60) return null;
    const formula = typeof raw.formula === 'string' ? raw.formula.trim() : '';
    if (formula && (!formula.startsWith('=') || formula.length > 500)) return null;
    if (!formula) {
      if (!Array.isArray(raw.values) || raw.values.length < 1 || raw.values.length > 5000) return null;
      if (raw.values.some(value => !cellValueOk(value))) return null;
      return { host: 'excel', op, sheet, table, name, formula: '', values: raw.values.map(value => [value]) };
    }
    return { host: 'excel', op, sheet, table, name, formula, values: null };
  }
  return null;
}

function normalizeActions(list) {
  if (!Array.isArray(list) || list.length < 1 || list.length > 40) return { actions: [], skipped: Array.isArray(list) ? list.length : 0 };
  const actions = [];
  let skipped = 0;
  for (const raw of list) {
    const action = normalizeAction(raw);
    if (action) actions.push(action);
    else skipped += 1;
  }
  return { actions, skipped };
}

function parseRevision(text) {
  const blocks = [...text.matchAll(/```(?:json)?\s*([\s\S]*?)```/gi)];
  for (let i = blocks.length - 1; i >= 0; i--) {
    try {
      const parsed = JSON.parse(blocks[i][1]);
      if (!Array.isArray(parsed?.actions)) continue;
      const { actions, skipped } = normalizeActions(parsed.actions);
      if (!actions.length) continue;
      const prose = (text.slice(0, blocks[i].index) + text.slice(blocks[i].index + blocks[i][0].length)).trim();
      return { prose, actions, skipped, edits: null };
    } catch { /* not an actions block */ }
  }
  const edits = parseEdits(text);
  return { prose: edits.prose, edits: edits.edits, skipped: edits.skipped, actions: null };
}

function actionIsRisky(action) {
  if (action.op === 'set') return action.grid.some(row => row.some(riskyEdit));
  if (action.op === 'tableColumn') return Boolean(action.formula && riskyEdit(action.formula));
  return false;
}

function describeAction(action) {
  if (action.op === 'set') {
    const preview = action.grid.length === 1 && action.grid[0].length === 1 ? String(action.grid[0][0]) : `${action.grid.length}×${action.grid[0].length} values`;
    return `${action.sheet}!${action.range} → ${preview}`;
  }
  if (action.op === 'format') return `Format ${action.sheet}!${action.range}`;
  if (action.op === 'clear') return `Clear ${action.sheet}!${action.range}`;
  if (action.op === 'insertRows') return `Insert ${action.count} row${action.count === 1 ? '' : 's'} at row ${action.row} on ${action.sheet}`;
  if (action.op === 'deleteRows') return `Delete ${action.count} row${action.count === 1 ? '' : 's'} at row ${action.row} on ${action.sheet}`;
  if (action.op === 'insertColumns') return `Insert ${action.count} column${action.count === 1 ? '' : 's'} at ${action.column} on ${action.sheet}`;
  if (action.op === 'addSheet') return `Add sheet ${action.name}`;
  if (action.op === 'renameSheet') return `Rename ${action.sheet} to ${action.name}`;
  if (action.op === 'autofit') return `Autofit ${action.sheet}!${action.range}`;
  if (action.op === 'sort') return `Sort ${action.sheet}!${action.range}`;
  if (action.op === 'chart') return `Chart ${action.sheet}!${action.range} (${action.type})`;
  if (action.op === 'table') return `Table ${action.name} on ${action.sheet}!${action.range}`;
  if (action.op === 'pivot') {
    const values = action.values.map(value => `${value.summarizeBy} of ${value.field}`).join(', ');
    const rows = action.rows.join(', ') || 'none';
    const columns = action.columns.join(', ') || 'none';
    return `PivotTable ${action.name} at ${action.destinationSheet}!${action.destinationCell} from ${action.sheet}!${action.range} · rows ${rows} · columns ${columns} · ${values}`;
  }
  if (action.op === 'filter') {
    if (action.clear) return `Clear all filters on ${action.sheet} (not reversible by Undo)`;
    const detail = action.criteria.filterOn === 'Values'
      ? `values ${action.criteria.values.join(', ')}`
      : action.criteria.filterOn === 'Custom'
        ? `${action.criteria.criterion1}${action.criteria.criterion2 ? ` ${action.criteria.operator} ${action.criteria.criterion2}` : ''}`
        : action.criteria.color || action.criteria.criterion1 || '';
    return `Filter ${action.sheet}!${action.range} column ${indexToColumn(columnIndex(action.range.split(':')[0].match(/^([A-Z]+)/)[1]) + action.key)} · ${action.criteria.filterOn} ${detail}`;
  }
  if (action.op === 'conditionalFormat') {
    if (action.clear) return `Clear conditional formats on ${action.sheet}!${action.range} (not reversible by Undo)`;
    return `Highlight ${action.sheet}!${action.range} where ${action.operator} ${action.formula1}${action.formula2 ? ` and ${action.formula2}` : ''}`;
  }
  if (action.op === 'borders') return `Borders ${action.edge} on ${action.sheet}!${action.range} (${action.style}, ${action.weight})`;
  if (action.op === 'freeze') {
    if (action.unfreeze) return `Unfreeze panes on ${action.sheet}`;
    return `Freeze ${action.rows} row(s) and ${action.columns} column(s) on ${action.sheet}`;
  }
  if (action.op === 'replaceAll') {
    const where = action.host === 'word' ? 'in the document' : `on ${action.sheet}`;
    return `Replace every "${action.text}" with "${action.replacement}" ${where} (not reversible by Undo)`;
  }
  if (action.op === 'dataValidation') {
    if (action.clear) return `Clear data validation on ${action.sheet}!${action.range} (not reversible by Undo)`;
    return `Dropdown on ${action.sheet}!${action.range}: ${action.list.join(', ')}`;
  }
  if (action.op === 'namedRange') return `Name ${action.sheet}!${action.range} as ${action.name}`;
  if (action.op === 'hide') {
    const verb = action.hidden ? 'Hide' : 'Unhide';
    return action.target === 'sheet' ? `${verb} sheet ${action.sheet}` : `${verb} ${action.target} ${action.sheet}!${action.range}`;
  }
  if (action.op === 'tableColumn') return `Add column ${action.name} to table ${action.table}${action.formula ? ` = ${action.formula}` : ''}`;
  if (action.op === 'replaceSelection') return 'Replace the Word selection';
  if (action.op === 'insertText') return `Insert text at the ${action.where} of the document`;
  if (action.op === 'formatSelection') return 'Format the Word selection';
  if (action.op === 'styleSelection') return `Apply Word style ${action.style}`;
  return action.op;
}

function writeCell(cell, value) {
  if (typeof value === 'string' && value.startsWith('=')) cell.formulas = [[value]];
  else cell.values = [[value]];
}

function sheetMissing(sheet) {
  return !!(sheet && sheet.isNullObject);
}

async function applyExcelActions(actions) {
  if (typeof Excel === 'undefined') throw new Error('These changes need Excel.');
  return Excel.run(async context => {
    const wanted = new Map();
    for (const action of actions) {
      if (!action.sheet || wanted.has(action.sheet)) continue;
      const sheet = context.workbook.worksheets.getItemOrNullObject(action.sheet);
      sheet.load('isNullObject');
      wanted.set(action.sheet, sheet);
    }
    await context.sync();
    const createdNames = new Set(actions.filter(action => action.op === 'addSheet').map(action => action.name));
    for (const [name, sheet] of wanted) {
      if (sheetMissing(sheet) && !createdNames.has(name)) throw new Error(`Sheet “${name}” is not in this workbook. Nothing was written.`);
    }
    const undo = [];
    for (const action of actions) {
      if (action.op === 'addSheet') {
        const created = context.workbook.worksheets.add(action.name);
        wanted.set(action.name, created);
        undo.push({ host: 'excel', op: 'addSheet', name: action.name });
        continue;
      }
      const sheet = action.sheet ? wanted.get(action.sheet) : null;
      if (action.sheet && (!sheet || sheetMissing(sheet))) throw new Error(`Sheet “${action.sheet}” is not in this workbook. Nothing was written.`);
      if (action.op === 'renameSheet') {
        sheet.name = action.name;
        undo.push({ host: 'excel', op: 'renameSheet', sheet: action.name, name: action.sheet });
        continue;
      }
      if (action.op === 'set') {
        const range = sheet.getRange(action.range);
        const cells = [];
        for (let r = 0; r < action.grid.length; r++) {
          for (let c = 0; c < action.grid[r].length; c++) {
            const cell = action.grid.length === 1 && action.grid[0].length === 1 ? range : range.getCell(r, c);
            cell.load(['values', 'formulas']);
            cells.push({ r, c, cell });
          }
        }
        await context.sync();
        const before = cells.map(({ r, c, cell }) => ({
          address: action.grid.length === 1 && action.grid[0].length === 1 ? action.range : cellAddress(action.range, r, c),
          formula: cell.formulas?.[0]?.[0],
          value: cell.values?.[0]?.[0]
        }));
        cells.forEach(({ r, c, cell }) => writeCell(cell, action.grid[r][c]));
        undo.push({ host: 'excel', op: 'set', sheet: action.sheet, cells: before });
        continue;
      }
      if (action.op === 'format') {
        const range = sheet.getRange(action.range);
        range.format?.font?.load?.(['bold', 'italic', 'color']);
        range.format?.fill?.load?.('color');
        range.format?.load?.(['horizontalAlignment', 'wrapText']);
        range.load('numberFormat');
        await context.sync();
        undo.push({
          host: 'excel', op: 'format', sheet: action.sheet, range: action.range,
          bold: range.format?.font?.bold, italic: range.format?.font?.italic, fontColor: range.format?.font?.color,
          fill: range.format?.fill?.color, align: range.format?.horizontalAlignment, wrap: range.format?.wrapText,
          numberFormat: range.numberFormat
        });
        if (action.bold != null) range.format.font.bold = action.bold;
        if (action.italic != null) range.format.font.italic = action.italic;
        if (action.fontColor) range.format.font.color = action.fontColor;
        if (action.fill) range.format.fill.color = action.fill;
        if (action.align) range.format.horizontalAlignment = action.align;
        if (action.wrap != null) range.format.wrapText = action.wrap;
        if (action.numberFormat) {
          const size = rangeSize(action.range);
          range.numberFormat = Array.from({ length: size.rows }, () => Array.from({ length: size.cols }, () => action.numberFormat));
        }
        continue;
      }
      if (action.op === 'clear') {
        const range = sheet.getRange(action.range);
        range.load(['values', 'formulas', 'rowCount', 'columnCount']);
        await context.sync();
        undo.push({ host: 'excel', op: 'clear', sheet: action.sheet, range: action.range, formulas: range.formulas, values: range.values });
        range.clear();
        continue;
      }
      if (action.op === 'insertRows') {
        sheet.getRange(`${action.row}:${action.row + action.count - 1}`).insert('Down');
        undo.push({ host: 'excel', op: 'deleteRows', sheet: action.sheet, row: action.row, count: action.count });
        continue;
      }
      if (action.op === 'deleteRows') {
        // Capture the full used width, not just A:H, so Undo can restore every column.
        const used = sheet.getUsedRangeOrNullObject();
        used.load(['isNullObject', 'columnCount']);
        await context.sync();
        const width = Math.max(8, Math.min(used.isNullObject ? 8 : used.columnCount, 200));
        const addr = `A${action.row}:${indexToColumn(width)}${action.row + action.count - 1}`;
        const range = sheet.getRange(addr);
        range.load(['values', 'formulas']);
        await context.sync();
        undo.push({ host: 'excel', op: 'restoreRows', sheet: action.sheet, row: action.row, count: action.count, formulas: range.formulas, values: range.values });
        sheet.getRange(`${action.row}:${action.row + action.count - 1}`).delete('Up');
        continue;
      }
      if (action.op === 'insertColumns') {
        const end = indexToColumn(columnIndex(action.column) + action.count - 1);
        sheet.getRange(`${action.column}:${end}`).insert('Right');
        undo.push({ host: 'excel', op: 'deleteColumns', sheet: action.sheet, column: action.column, count: action.count });
        continue;
      }
      if (action.op === 'autofit') {
        const range = sheet.getRange(action.range);
        // Capture each width and height first: autofit itself reports nothing to undo.
        const clean = action.range.replace(/\$/g, '').toUpperCase();
        const columnOnly = /^[A-Z]+:[A-Z]+$/.test(clean);
        const match = clean.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
        const firstCol = columnOnly ? columnIndex(clean.split(':')[0]) : match ? columnIndex(match[1]) : 0;
        const lastCol = columnOnly ? columnIndex(clean.split(':')[1]) : match ? columnIndex(match[3] || match[1]) : 0;
        const firstRow = !columnOnly && match ? Number(match[2]) : 0;
        const lastRow = !columnOnly && match ? Number(match[4] || match[2]) : 0;
        const widths = [];
        const heights = [];
        for (let c = firstCol; c > 0 && c <= lastCol && c - firstCol < 200; c++) {
          const one = sheet.getRange(indexToColumn(c) + '1');
          one.format.load('columnWidth');
          widths.push({ column: indexToColumn(c), one });
        }
        for (let r = firstRow; r > 0 && r <= lastRow && r - firstRow < 500; r++) {
          const one = sheet.getRange('A' + r);
          one.format.load('rowHeight');
          heights.push({ row: r, one });
        }
        await context.sync();
        undo.push({
          host: 'excel', op: 'autofit', sheet: action.sheet,
          widths: widths.map(item => ({ column: item.column, width: item.one.format.columnWidth })),
          heights: heights.map(item => ({ row: item.row, height: item.one.format.rowHeight }))
        });
        range.format.autofitColumns();
        range.format.autofitRows();
        continue;
      }
      if (action.op === 'sort') {
        const range = sheet.getRange(action.range);
        range.load(['values', 'formulas']);
        await context.sync();
        undo.push({ host: 'excel', op: 'restoreGrid', sheet: action.sheet, range: action.range, formulas: range.formulas, values: range.values });
        range.sort.apply([{ key: action.key, ascending: action.ascending }], false, action.hasHeaders);
        continue;
      }
      if (action.op === 'chart') {
        const name = `Revise${Math.random().toString(36).slice(2, 8)}`;
        const chart = sheet.charts.add(action.type, sheet.getRange(action.range), 'Auto');
        chart.name = name;
        if (action.title) chart.title.text = action.title;
        undo.push({ host: 'excel', op: 'chart', sheet: action.sheet, name });
        continue;
      }
      if (action.op === 'table') {
        const table = sheet.tables.add(sheet.getRange(action.range), true);
        table.name = action.name;
        undo.push({ host: 'excel', op: 'table', sheet: action.sheet, name: action.name });
        continue;
      }
      if (action.op === 'pivot') {
        let dest = wanted.get(action.destinationSheet);
        if (!dest) {
          dest = context.workbook.worksheets.getItemOrNullObject(action.destinationSheet);
          dest.load('isNullObject');
          await context.sync();
        }
        let createdSheet = false;
        if (!dest || sheetMissing(dest)) {
          dest = context.workbook.worksheets.add(action.destinationSheet);
          wanted.set(action.destinationSheet, dest);
          createdSheet = true;
          await context.sync();
        }
        const pivot = sheet.pivotTables.add(action.name, sheet.getRange(action.range), dest.getRange(action.destinationCell));
        pivot.layout.layoutType = action.layout;
        pivot.layout.showRowGrandTotals = action.grandTotals;
        pivot.layout.showColumnGrandTotals = action.grandTotals;
        pivot.hierarchies.load('items/name');
        await context.sync();
        // Name the offending field instead of letting Office.js fail opaquely later.
        const available = (pivot.hierarchies.items || []).map(item => item.name);
        const requested = [...action.rows, ...action.columns, ...action.filters, ...action.values.map(value => value.field)];
        const missing = requested.filter(field => !available.includes(field));
        if (missing.length) {
          pivot.delete();
          if (createdSheet) dest.delete();
          await context.sync();
          throw new Error(`Pivot field(s) not in ${action.sheet}!${action.range}: ${missing.join(', ')}. Headers are: ${available.join(', ')}`);
        }
        for (const field of action.rows) pivot.rowHierarchies.add(pivot.hierarchies.getItem(field));
        for (const field of action.columns) pivot.columnHierarchies.add(pivot.hierarchies.getItem(field));
        for (const field of action.filters) pivot.filterHierarchies.add(pivot.hierarchies.getItem(field));
        for (const value of action.values) {
          const data = pivot.dataHierarchies.add(pivot.hierarchies.getItem(value.field));
          data.summarizeBy = value.summarizeBy;
          data.name = value.name;
        }
        pivot.refresh();
        undo.push({ host: 'excel', op: 'pivot', sheet: action.destinationSheet, name: action.name, createdSheet });
        continue;
      }
      if (action.op === 'filter') {
        if (action.clear) sheet.autoFilter.clearCriteria();
        else {
          sheet.autoFilter.apply(sheet.getRange(action.range), action.key, action.criteria);
          undo.push({ host: 'excel', op: 'filter', sheet: action.sheet, key: action.key });
        }
        continue;
      }
      if (action.op === 'conditionalFormat') {
        const range = sheet.getRange(action.range);
        if (action.clear) { range.conditionalFormats.clearAll(); continue; }
        const cf = range.conditionalFormats.add('CellValue');
        cf.cellValue.rule = action.formula2
          ? { operator: action.operator, formula1: action.formula1, formula2: action.formula2 }
          : { operator: action.operator, formula1: action.formula1 };
        if (action.fill) cf.cellValue.format.fill.color = action.fill;
        if (action.fontColor) cf.cellValue.format.font.color = action.fontColor;
        if (action.bold != null) cf.cellValue.format.font.bold = action.bold;
        cf.load('id');
        await context.sync();
        undo.push({ host: 'excel', op: 'conditionalFormat', sheet: action.sheet, range: action.range, id: cf.id });
        continue;
      }
      if (action.op === 'borders') {
        const range = sheet.getRange(action.range);
        const edges = action.edge === 'all'
          ? ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideVertical', 'InsideHorizontal']
          : action.edge === 'outline' ? ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] : [action.edge];
        const targets = edges.map(edge => {
          const border = range.format.borders.getItem(edge);
          border.load(['style', 'weight', 'color']);
          return { edge, border };
        });
        await context.sync();
        undo.push({
          host: 'excel', op: 'borders', sheet: action.sheet, range: action.range,
          edges: targets.map(({ edge, border }) => ({ edge, style: border.style, weight: border.weight, color: border.color }))
        });
        for (const { border } of targets) {
          border.style = action.style;
          border.weight = action.weight;
          if (action.color) border.color = action.color;
        }
        continue;
      }
      if (action.op === 'freeze') {
        const location = sheet.freezePanes.getLocationOrNullObject();
        location.load(['isNullObject', 'address']);
        await context.sync();
        undo.push({ host: 'excel', op: 'freeze', sheet: action.sheet, address: location.isNullObject ? '' : location.address });
        if (action.unfreeze) sheet.freezePanes.unfreeze();
        else {
          if (action.rows) sheet.freezePanes.freezeRows(action.rows);
          if (action.columns) sheet.freezePanes.freezeColumns(action.columns);
        }
        continue;
      }
      if (action.op === 'replaceAll') {
        sheet.replaceAll(action.text, action.replacement, { completeMatch: action.completeMatch, matchCase: action.matchCase });
        continue;
      }
      if (action.op === 'dataValidation') {
        const range = sheet.getRange(action.range);
        if (action.clear) { range.dataValidation.clear(); continue; }
        range.dataValidation.rule = { list: { inCellDropDown: true, source: action.list.join(',') } };
        range.dataValidation.ignoreBlanks = action.ignoreBlanks;
        undo.push({ host: 'excel', op: 'dataValidation', sheet: action.sheet, range: action.range });
        continue;
      }
      if (action.op === 'namedRange') {
        context.workbook.names.add(action.name, sheet.getRange(action.range));
        undo.push({ host: 'excel', op: 'namedRange', name: action.name });
        continue;
      }
      if (action.op === 'hide') {
        if (action.target === 'sheet') {
          sheet.load('visibility');
          await context.sync();
          undo.push({ host: 'excel', op: 'hideSheet', sheet: action.sheet, visibility: sheet.visibility });
          sheet.visibility = action.hidden ? 'Hidden' : 'Visible';
          continue;
        }
        const range = sheet.getRange(action.range);
        const property = action.target === 'rows' ? 'rowHidden' : 'columnHidden';
        range.load(property);
        await context.sync();
        undo.push({ host: 'excel', op: 'hide', sheet: action.sheet, range: action.range, property, previous: range[property] });
        range[property] = action.hidden;
        continue;
      }
      if (action.op === 'tableColumn') {
        const table = sheet.tables.getItem(action.table);
        const values = action.formula ? [[action.formula]] : action.values;
        table.columns.add(undefined, values, action.name);
        undo.push({ host: 'excel', op: 'tableColumn', sheet: action.sheet, table: action.table, name: action.name });
      }
    }
    await context.sync();
    return undo;
  });
}

function cellAddress(range, rowOffset, colOffset) {
  const top = range.split(':')[0].match(/^([A-Z]+)(\d+)$/);
  return indexToColumn(columnIndex(top[1]) + colOffset) + (Number(top[2]) + rowOffset);
}

async function undoExcel(records) {
  await Excel.run(async context => {
    for (const rec of [...records].reverse()) {
      const sheet = rec.sheet ? context.workbook.worksheets.getItem(rec.sheet) : null;
      if (rec.op === 'set') {
        for (const cell of rec.cells) {
          const range = context.workbook.worksheets.getItem(rec.sheet).getRange(cell.address);
          if (String(cell.formula || '').startsWith('=')) range.formulas = [[cell.formula]];
          else if (cell.value === '' || cell.value == null) range.clear();
          else range.values = [[cell.value]];
        }
      } else if (rec.op === 'format') {
        const range = sheet.getRange(rec.range);
        if (rec.bold != null) range.format.font.bold = rec.bold;
        if (rec.italic != null) range.format.font.italic = rec.italic;
        if (rec.fontColor) range.format.font.color = rec.fontColor;
        if (rec.fill) range.format.fill.color = rec.fill;
        if (rec.align) range.format.horizontalAlignment = rec.align;
        if (rec.wrap != null) range.format.wrapText = rec.wrap;
        if (rec.numberFormat) range.numberFormat = rec.numberFormat;
      } else if (rec.op === 'clear' || rec.op === 'restoreGrid') {
        const range = sheet.getRange(rec.range);
        if (rec.formulas) range.formulas = rec.formulas;
        else if (rec.values) range.values = rec.values;
      } else if (rec.op === 'deleteRows') {
        sheet.getRange(`${rec.row}:${rec.row + rec.count - 1}`).delete('Up');
      } else if (rec.op === 'restoreRows') {
        sheet.getRange(`${rec.row}:${rec.row + rec.count - 1}`).insert('Down');
        const range = sheet.getRange(`A${rec.row}:${indexToColumn(8)}${rec.row + rec.count - 1}`);
        if (rec.formulas) range.formulas = rec.formulas;
      } else if (rec.op === 'deleteColumns') {
        const end = indexToColumn(columnIndex(rec.column) + rec.count - 1);
        sheet.getRange(`${rec.column}:${end}`).delete('Left');
      } else if (rec.op === 'addSheet') {
        context.workbook.worksheets.getItem(rec.name).delete();
      } else if (rec.op === 'renameSheet') {
        context.workbook.worksheets.getItem(rec.sheet).name = rec.name;
      } else if (rec.op === 'chart') {
        sheet.charts.getItem(rec.name).delete();
      } else if (rec.op === 'table') {
        sheet.tables.getItem(rec.name).delete();
      } else if (rec.op === 'pivot') {
        if (rec.createdSheet) context.workbook.worksheets.getItem(rec.sheet).delete();
        else sheet.pivotTables.getItem(rec.name).delete();
      } else if (rec.op === 'filter') {
        sheet.autoFilter.clearColumnCriteria(rec.key);
      } else if (rec.op === 'conditionalFormat') {
        sheet.getRange(rec.range).conditionalFormats.getItem(rec.id).delete();
      } else if (rec.op === 'borders') {
        const range = sheet.getRange(rec.range);
        for (const edge of rec.edges) {
          const border = range.format.borders.getItem(edge.edge);
          border.style = edge.style;
          border.weight = edge.weight;
          border.color = edge.color;
        }
      } else if (rec.op === 'freeze') {
        if (rec.address) sheet.freezePanes.freezeAt(rec.address);
        else sheet.freezePanes.unfreeze();
      } else if (rec.op === 'dataValidation') {
        sheet.getRange(rec.range).dataValidation.clear();
      } else if (rec.op === 'namedRange') {
        context.workbook.names.getItem(rec.name).delete();
      } else if (rec.op === 'hideSheet') {
        context.workbook.worksheets.getItem(rec.sheet).visibility = rec.visibility || 'Visible';
      } else if (rec.op === 'hide') {
        sheet.getRange(rec.range)[rec.property] = rec.previous;
      } else if (rec.op === 'tableColumn') {
        sheet.tables.getItem(rec.table).columns.getItem(rec.name).delete();
      } else if (rec.op === 'autofit') {
        for (const item of rec.widths || []) {
          if (item.width != null) sheet.getRange(item.column + '1').format.columnWidth = item.width;
        }
        for (const item of rec.heights || []) {
          if (item.height != null) sheet.getRange('A' + item.row).format.rowHeight = item.height;
        }
      }
    }
    await context.sync();
  });
}

async function applyWordActions(actions) {
  if (typeof Word === 'undefined') throw new Error('These changes need Word.');
  return Word.run(async context => {
    const undo = [];
    for (const action of actions) {
      const selection = context.document.getSelection();
      if (action.op !== 'insertText') selection.load('text');
      await context.sync();
      if (action.op === 'replaceSelection') {
        undo.push({ host: 'word', op: 'replaceSelection', text: selection.text || '' });
        selection.insertText(action.text, 'Replace');
      } else if (action.op === 'insertText') {
        context.document.body.insertText(action.text, action.where === 'start' ? 'Start' : 'End');
        undo.push({ host: 'word', op: 'removeText', text: action.text });
      } else if (action.op === 'formatSelection') {
        selection.font?.load?.(['bold', 'italic']);
        await context.sync();
        undo.push({ host: 'word', op: 'formatSelection', bold: selection.font?.bold, italic: selection.font?.italic });
        if (typeof action.bold === 'boolean') selection.font.bold = action.bold;
        if (typeof action.italic === 'boolean') selection.font.italic = action.italic;
      } else if (action.op === 'styleSelection') {
        selection.load('style');
        await context.sync();
        undo.push({ host: 'word', op: 'styleSelection', style: selection.style || 'Normal' });
        selection.style = action.style;
      } else if (action.op === 'replaceAll') {
        const found = context.document.body.search(action.text, { matchCase: action.matchCase, matchWholeWord: action.matchWholeWord });
        found.load('items');
        await context.sync();
        for (const item of found.items || []) item.insertText(action.replacement, 'Replace');
      }
    }
    await context.sync();
    return undo;
  });
}

async function undoWord(records) {
  await Word.run(async context => {
    for (const rec of [...records].reverse()) {
      const selection = context.document.getSelection();
      if (rec.op === 'replaceSelection') selection.insertText(rec.text, 'Replace');
      else if (rec.op === 'formatSelection') {
        if (rec.bold != null) selection.font.bold = rec.bold;
        if (rec.italic != null) selection.font.italic = rec.italic;
      } else if (rec.op === 'styleSelection') selection.style = rec.style;
      else if (rec.op === 'removeText' && typeof context.document.body.search === 'function') {
        const found = context.document.body.search(rec.text, { matchCase: true });
        found.load('items');
        await context.sync();
        found.items?.[0]?.insertText('', 'Replace');
      }
    }
    await context.sync();
  });
}

async function applyPlan(actions) {
  const undo = [];
  const excel = actions.filter(action => action.host === 'excel');
  const word = actions.filter(action => action.host === 'word');
  if (excel.length) undo.push(...await applyExcelActions(excel));
  if (word.length) undo.push(...await applyWordActions(word));
  return undo;
}

async function undoPlan(records) {
  const excel = records.filter(rec => rec.host === 'excel');
  const word = records.filter(rec => rec.host === 'word');
  if (excel.length) await undoExcel(excel);
  if (word.length) await undoWord(word);
}

function attachUndo(actionsEl, label, records, noun) {
  const undoBtn = document.createElement('button');
  undoBtn.className = 'secondary';
  undoBtn.textContent = 'Undo last apply';
  undoBtn.addEventListener('click', async () => {
    undoBtn.disabled = true;
    try {
      await undoPlan(records);
      label.textContent = `UNDONE — RESTORED ${records.length} CHANGE${records.length === 1 ? '' : 'S'}`;
      toast(`Restored the previous ${noun}.`);
    } catch (error) {
      undoBtn.disabled = false;
      toast(`Could not undo: ${error.message}`);
    }
  });
  actionsEl.append(undoBtn);
}

async function renderPlanCard(actions, skipped, options = {}) {
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant edits-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  const direct = !options.manualOnly && Boolean($('directRevise')?.checked);
  label.textContent = direct ? `DIRECT REVISE · ${actions.length}` : `PROPOSED REVISION · ${actions.length} · APPLY WITH YOUR CLICK`;
  card.append(label);
  actions.forEach(action => {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.textContent = describeAction(action) + (actionIsRisky(action) ? ' · waits for your click' : '');
    card.append(row);
  });
  if (skipped) {
    const note = document.createElement('div');
    note.className = 'edit-note';
    note.textContent = `${skipped} change${skipped === 1 ? '' : 's'} skipped because the address or value was not usable.`;
    card.append(note);
  }
  if (actions.some(actionIsRisky)) {
    const warning = document.createElement('div');
    warning.className = 'edit-warning';
    warning.textContent = 'Some formulas can reach the network or run legacy calls. Direct revise leaves those for your click.';
    card.append(warning);
  }
  const buttonRow = document.createElement('div');
  buttonRow.className = 'button-row';
  const risky = actions.filter(actionIsRisky);
  const safe = actions.filter(action => !actionIsRisky(action));
  const manual = direct ? risky : actions;
  const noun = actions.some(action => action.host === 'word') && !actions.some(action => action.host === 'excel') ? 'document' : 'workbook';
  if (direct && safe.length) {
    try {
      const records = await applyPlan(safe);
      label.textContent = manual.length ? `REVISED ${safe.length} · ${manual.length} STILL NEED A CLICK` : `REVISED ${safe.length} CHANGE${safe.length === 1 ? '' : 'S'}`;
      attachUndo(buttonRow, label, records, noun);
      toast(noun === 'document' ? 'Revised the document.' : 'Revised the workbook.');
    } catch (error) {
      toast(`Could not revise: ${error.message}`);
      manual.splice(0, manual.length, ...actions);
    }
  }
  if (manual.length) {
    const apply = document.createElement('button');
    apply.className = 'primary';
    apply.textContent = direct && risky.length && manual.length === risky.length ? 'Apply risky changes' : 'Apply edits';
    const dismiss = document.createElement('button');
    dismiss.className = 'secondary';
    dismiss.textContent = 'Dismiss';
    apply.addEventListener('click', async () => {
      apply.disabled = true;
      dismiss.disabled = true;
      try {
        const records = await applyPlan(manual);
        label.textContent = `APPLIED ${manual.length} CHANGE${manual.length === 1 ? '' : 'S'}`;
        attachUndo(buttonRow, label, records, noun);
        toast(noun === 'document' ? 'Revised the document.' : 'Revised the workbook.');
      } catch (error) {
        apply.disabled = false;
        dismiss.disabled = false;
        toast(`Could not apply edits: ${error.message}`);
      }
    });
    dismiss.addEventListener('click', () => card.remove());
    buttonRow.append(apply, dismiss);
  } else {
    const dismiss = document.createElement('button');
    dismiss.className = 'secondary';
    dismiss.textContent = 'Dismiss';
    dismiss.addEventListener('click', () => card.remove());
    buttonRow.append(dismiss);
  }
  card.append(buttonRow);
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

async function withCells(edits, fn) {
  return Excel.run(async context => {
    const sheets = edits.map(edit => context.workbook.worksheets.getItemOrNullObject(edit.sheet));
    sheets.forEach(sheet => sheet.load('isNullObject'));
    await context.sync();
    const ranges = edits.map((edit, i) => sheets[i].isNullObject ? null : sheets[i].getRange(edit.cell));
    ranges.forEach(range => range && range.load(['values', 'formulas']));
    await context.sync();
    const result = await fn(ranges);
    await context.sync();
    return result;
  });
}

function oldCellText(old) {
  if (!old) return '(sheet missing)';
  if (String(old.formula || '').startsWith('=')) return old.formula;
  if (old.value === '' || old.value == null) return '(empty)';
  return String(old.value);
}

async function renderEditsCard(edits, skipped) {
  const conversation = $('conversation');
  const card = document.createElement('div'); card.className = 'message assistant edits-card';
  const label = document.createElement('div'); label.className = 'message-label'; label.textContent = `PROPOSED EDITS · ${edits.length} · REVIEW OLD → NEW · APPLY ONLY WITH YOUR CLICK`;
  card.append(label);
  let olds = null;
  try { olds = await withCells(edits, ranges => ranges.map(range => range ? { value: range.values?.[0]?.[0], formula: range.formulas?.[0]?.[0] } : null)); }
  catch { olds = null; }
  edits.forEach((edit, i) => {
    const row = document.createElement('div'); row.className = 'edit-row';
    const oldText = olds == null ? '(current unreadable)' : oldCellText(olds[i]);
    row.textContent = `${edit.sheet}!${edit.cell}: ${oldText} → ${String(edit.value)}`;
    card.append(row);
  });
  if (skipped) {
    const note = document.createElement('div'); note.className = 'edit-note';
    note.textContent = `${skipped} edit${skipped === 1 ? '' : 's'} skipped — each change must be one cell on a sheet in this workbook.`;
    card.append(note);
  }
  if (edits.some(edit => riskyEdit(edit.value))) {
    const warning = document.createElement('div'); warning.className = 'edit-warning';
    warning.textContent = 'Some formulas can reach the network or run legacy calls (WEBSERVICE, FILTERXML, RTD, HYPERLINK, and similar). Apply only if you trust them.';
    card.append(warning);
  }
  const actions = document.createElement('div'); actions.className = 'button-row';
  const apply = document.createElement('button'); apply.className = 'primary'; apply.textContent = 'Apply edits';
  const dismiss = document.createElement('button'); dismiss.className = 'secondary'; dismiss.textContent = 'Dismiss';
  apply.addEventListener('click', async () => {
    apply.disabled = true; dismiss.disabled = true;
    try {
      const recs = await withCells(edits, async ranges => {
        if (ranges.some(range => !range)) throw new Error('A proposed sheet is not in this workbook. Nothing was written.');
        const before = ranges.map((range, i) => ({ sheet: edits[i].sheet, cell: edits[i].cell, value: range.values?.[0]?.[0], formula: range.formulas?.[0]?.[0] }));
        ranges.forEach((range, i) => {
          const value = edits[i].value;
          if (typeof value === 'string' && value.startsWith('=')) range.formulas = [[value]];
          else range.values = [[value]];
        });
        return before;
      });
      label.textContent = `APPLIED ${edits.length} EDITS`;
      const undoBtn = document.createElement('button'); undoBtn.className = 'secondary'; undoBtn.textContent = 'Undo last apply';
      undoBtn.addEventListener('click', async () => {
        undoBtn.disabled = true;
        try {
          await Excel.run(async context => {
            for (const rec of recs) {
              const range = context.workbook.worksheets.getItem(rec.sheet).getRange(rec.cell);
              if (String(rec.formula || '').startsWith('=')) range.formulas = [[rec.formula]];
              else if (rec.value === '' || rec.value == null) range.clear();
              else range.values = [[rec.value]];
            }
            await context.sync();
          });
          label.textContent = `UNDONE — RESTORED ${recs.length} CELLS`;
          toast('Restored previous cell values.');
        } catch (error) { undoBtn.disabled = false; toast(`Could not undo: ${error.message}`); }
      });
      actions.append(undoBtn);
      toast(`Applied ${edits.length} edits to the workbook.`);
    } catch (error) {
      apply.disabled = false; dismiss.disabled = false;
      toast(`Could not apply edits: ${error.message}`);
    }
  });
  dismiss.addEventListener('click', () => card.remove());
  actions.append(apply, dismiss); card.append(actions);
  conversation.append(card); conversation.scrollTop = conversation.scrollHeight;
}

function dedupContext(messages, freshCount = 1) {
  // Older turns keep their prose only; re-sending past workbook/file context multiplies input tokens each turn.
  return messages.map((m, i) => (m.role === 'user' && i < messages.length - freshCount)
    ? { ...m, content: m.content.replace(/\n\n\[Attached file START:[^\]]+\][\s\S]*?\n\[Attached file END\]/g, '').replace(/\n\n\[(?:Excel (?:selection|used range|workbook metadata|range)|Word (?:selection|document|document outline))[^\]]*\]\n[\s\S]*$/, '') }
    : m);
}

function consumeSseFrames(buf, onFrame) {
  let rest = buf;
  let idx;
  while ((idx = rest.indexOf('\n\n')) >= 0) {
    onFrame(rest.slice(0, idx));
    rest = rest.slice(idx + 2);
  }
  return rest;
}

async function streamChat(payload) {
  const controller = new AbortController();
  abortChat = controller;
  let bubble = null;
  let text = '';
  try {
    const headers = { 'Content-Type': 'application/json' };
    if (hostedRelay() && !byoKey()) throw new Error('Add your Qwen key in Settings. The hosted pane sends it per request and never stores it server-side.');
    const wire = hostedRelay() ? { ...payload, apiKey: byoKey() } : payload;
    let response;
    for (let attempt = 0; attempt < 2; attempt++) {
      response = await fetch(`${API}/api/chat-stream`, { method: 'POST', headers, body: JSON.stringify(wire), signal: controller.signal });
      if (![500, 502, 503, 504].includes(response.status) || attempt === 1) break;
      await new Promise(resolve => setTimeout(resolve, 700));
    }
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      const detail = typeof data.error === 'string' ? data.error : data.error?.message;
      throw new Error(detail || `Cloud relay returned HTTP ${response.status}. Please retry; if it repeats, report the time and status to support. 雲端中繼回傳 HTTP ${response.status}；若持續出現，請提供時間和狀態碼。`);
    }
    bubble = renderMessage('assistant', '');
    const bodyEl = bubble.children[bubble.children.length - 1]; // HTMLCollection has no .at()
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buf = '', model = '';
    const onFrame = frame => {
    for (const line of frame.split('\n')) {
      if (!line.startsWith('data: ')) continue;
      const raw = line.slice(6);
      if (raw === '[DONE]') continue;
      let event; try { event = JSON.parse(raw); } catch { continue; }
      if (event.error) throw new Error(event.error);
      if (event.model) model = event.model;
      if (event.status === 'thinking' && !text) { bodyEl.textContent = 'Thinking…'; continue; }
      const delta = event.choices?.[0]?.delta?.content;
      if (delta) {
        text += delta; bodyEl.textContent = text;
        $('conversation').scrollTop = $('conversation').scrollHeight;
      }
    }
  };
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      buf = consumeSseFrames(buf, onFrame);
    }
    buf += decoder.decode();
    if (buf.trim()) onFrame(buf);
    if (!text.trim()) {
      const error = new Error('The provider returned no text answer.');
      error.bubble = bubble;
      throw error;
    }
    return { text, model, bubble };
  } catch (error) {
    if (bubble) { error.bubble = bubble; error.partial = text; }
    throw error;
  } finally {
    if (abortChat === controller) abortChat = null;
  }
}

function resetComposer() {
  state.busy = false;
  $('send').textContent = '↑';
  $('send').setAttribute?.('aria-label', 'Send message');
}

function showAgentStatus(round, detail) {
  const status = $('agentStatus');
  if (!status) return;
  status.textContent = detail ? `Agent ${round}/${AgentCore.MAX_LOOPS} · ${detail}` : '';
  syncAgentControls();
}

function syncAgentControls() {
  const controls = $('agentControls');
  if (!controls) return;
  controls.classList.toggle('hidden', !$('agentStatus').textContent || (!activeAgent && operatingMode() !== 'agent'));
  $('continueAgent').classList.toggle('hidden', !agentCheckpoint || state.busy);
  $('addUsedRange').classList.toggle('hidden', operatingMode() === 'agent' || operatingMode() === 'plan');
}

function stopGeneration() {
  if (activeAgent) {
    activeAgent.stopRequested = true;
    abortChat?.abort();
    showAgentStatus(activeAgent.round, 'Stopping… · 正在停止…');
    return;
  }
  state.turn += 1;
  abortChat?.abort();
  resetComposer();
}

function operatingMode() {
  const value = $('workMode')?.value;
  return value === 'chat' || value === 'plan' || value === 'agent' ? value : 'legacy';
}

function selectedModel() {
  const value = $('qwenModel').value.trim();
  return CHAT_MODELS.includes(value) ? value : 'qwen4-max';
}

function selectedEffort() {
  const value = String($('effortMode')?.value || '').trim().toLowerCase();
  return CHAT_EFFORTS.includes(value) ? value : '';
}

function chatFields() {
  const effort = selectedEffort();
  return { model: selectedModel(), fast: Boolean($('fastMode').checked), ...(effort ? { effort } : {}) };
}

function saveChatPrefs() {
  try { localStorage.setItem('eqgc.chat.v1', JSON.stringify({ model: selectedModel(), effort: selectedEffort() })); } catch { /* storage optional */ }
}

function syncModel(sourceId) {
  const picked = $(sourceId).value.trim();
  const model = CHAT_MODELS.includes(picked) ? picked : 'qwen4-max';
  $('qwenModel').value = model;
  $('modelPick').value = model;
  saveChatPrefs();
}

function restoreChatPrefs() {
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem('eqgc.chat.v1') || '{}'); } catch { saved = {}; }
  const model = CHAT_MODELS.includes(saved.model) ? saved.model : selectedModel();
  $('qwenModel').value = model;
  $('modelPick').value = model;
  if (CHAT_EFFORTS.includes(saved.effort)) $('effortMode').value = saved.effort;
}

async function readWorkbookMeta() {
  if (!state.officeReady) return null;
  if (state.host === 'Word') return buildWordContext(8000);
  return Excel.run(async context => {
    const book = context.workbook;
    book.load?.('name');
    const sheets = book.worksheets;
    sheets.load?.('items/name');
    await context.sync();
    state.workbookName = book.name || '';
    const sheetItems = sheets.items || [];
    const lines = [`workbook: ${book.name || '(unsaved)'}`, `sheets: ${sheetItems.length}`];
    for (const sheet of sheetItems.slice(0, 30)) {
      let extra = '';
      try {
        const used = sheet.getUsedRangeOrNullObject?.();
        if (used) {
          used.load?.(['isNullObject', 'address', 'rowCount', 'columnCount']);
          await context.sync();
          if (!used.isNullObject) extra += ` used ${used.address} rows ${used.rowCount} cols ${used.columnCount}`;
        }
      } catch { /* this sheet has no used-range helper */ }
      try {
        const tables = sheet.tables;
        tables?.load?.('items/name');
        await context.sync();
        const names = (tables?.items || []).map(table => table.name).filter(Boolean);
        if (names.length) extra += ` tables ${names.join(', ')}`;
      } catch { /* tables are not available here */ }
      lines.push(`sheet: ${sheet.name || 'Sheet'}${extra}`);
    }
    for (const sheet of sheetItems.slice(30)) lines.push(`sheet: ${sheet.name || 'Sheet'} (size not inventoried; request a profile)`);
    const selected = book.getSelectedRange?.();
    if (selected) {
      selected.load?.(['address', 'rowCount', 'columnCount']);
      await context.sync();
      if (selected.address) lines.push(`selection: ${selected.address} rows ${selected.rowCount || ''} cols ${selected.columnCount || ''}`);
    }
    const clipped = trimContext(lines.join('\n'), 8000);
    return { header: '[Excel workbook metadata]', text: clipped.text, address: 'Workbook', kind: 'metadata', truncated: clipped.truncated, chars: clipped.text.length };
  });
}

async function contextForMode(mode, limit = CONTEXT_CAP, question = '') {
  if (!state.officeReady || !$('includeContext').checked) return null;
  if (state.host === 'Word') return buildWordContext(limit, question);
  if (mode === 'chat' || mode === 'legacy' || state.forceSheet) return buildContext(limit);
  let ctx;
  try { ctx = await readWorkbookMeta(); }
  catch { ctx = await buildContext(limit); }
  if (ctx && ctx.text.length > limit) {
    const clipped = trimContext(ctx.text, limit);
    return { ...ctx, text: clipped.text, chars: clipped.text.length, truncated: Boolean(ctx.truncated || clipped.truncated) };
  }
  return ctx;
}

async function readGrid(sheet, range) {
  const size = rangeSize(range);
  const cols = Math.min(size.cols, READ_MAX_CELLS);
  const rows = Math.max(1, Math.min(size.rows, Math.floor(READ_MAX_CELLS / cols)));
  return Excel.run(async context => {
    const worksheet = context.workbook.worksheets.getItem(sheet);
    const cells = worksheet.getRange(clampRange(range, rows, cols));
    cells.load(['values', 'formulas']);
    await context.sync();
    return {
      values: cells.values || [],
      formulas: cells.formulas || []
    };
  });
}

function plainRange(address) {
  return normalizeRange(String(address || '').split('!').pop());
}

function clampRange(addr, maxRows, maxCols) {
  const size = rangeSize(addr);
  const rows = Math.min(size.rows, maxRows);
  const cols = Math.min(size.cols, maxCols);
  if (rows === size.rows && cols === size.cols) return addr;
  const top = addr.split(':')[0].match(/^([A-Z]+)(\d+)$/);
  return `${top[1]}${top[2]}:${indexToColumn(columnIndex(top[1]) + cols - 1)}${Number(top[2]) + rows - 1}`;
}

// Read large ranges in local batches and return compact statistics to the model.
async function profileSheet(sheetName, requested, onProgress = () => {}) {
  return Excel.run(async context => {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    let addr = normalizeRange(requested || '');
    if (!addr) {
      const used = worksheet.getUsedRangeOrNullObject();
      used.load(['isNullObject', 'address', 'rowCount', 'columnCount']);
      await context.sync();
      if (used.isNullObject) return null;
      addr = plainRange(used.address);
      if (!addr) return null;
    }
    const size = rangeSize(addr);
    const start = addr.split(':')[0].match(/^([A-Z]+)(\d+)$/);
    const firstColumn = columnIndex(start[1]);
    const firstRow = Number(start[2]);
    const cols = size.cols;
    const rows = size.rows;
    const firstLetter = indexToColumn(firstColumn);
    const lastLetter = indexToColumn(firstColumn + cols - 1);
    const batchRows = Math.max(1, Math.floor(PROFILE_BATCH_CELLS / cols));
    let header = null;
    let headerSource = '';
    if (firstRow > 1) {
      const titleRow = worksheet.getRange(`${firstLetter}1:${lastLetter}1`);
      titleRow.load('values');
      await context.sync();
      const candidate = titleRow.values?.[0] || [];
      if (candidate.filter(value => typeof value === 'string' && value.trim()).length >= Math.min(2, cols)) {
        header = candidate.map(value => String(value ?? '').trim());
        headerSource = 'row 1';
      }
    }
    const stats = Array.from({ length: cols }, () => ({ seen: new Map(), distinctCapped: false, filled: 0, blank: 0, numeric: 0, sum: 0, correction: 0, min: null, max: null, errors: 0, errorCells: [] }));
    let sampledRows = 0;
    let scannedRows = 0;
    let scanError = '';
    for (let offset = 0; offset < rows; offset += batchRows) {
      const count = Math.min(batchRows, rows - offset);
      const top = firstRow + offset;
      const target = worksheet.getRange(`${firstLetter}${top}:${lastLetter}${top + count - 1}`);
      target.load('values');
      try { await context.sync(); }
      catch (error) {
        if (!scannedRows) throw error;
        scanError = String(error?.message || error).slice(0, 300);
        break;
      }
      const values = target.values || [];
      if (values.length !== count) {
        if (!scannedRows) throw new Error('Excel returned an incomplete profile batch.');
        scanError = `Excel returned ${values.length} of ${count} rows in a profile batch`;
        break;
      }
      if (!header) {
        header = (values[0] || []).map(value => String(value ?? '').trim());
        headerSource = `row ${top}`;
      }
      for (let r = headerSource === `row ${top}` ? 1 : 0; r < values.length; r++) {
        sampledRows += 1;
        for (let c = 0; c < cols; c++) {
          const value = values[r]?.[c];
          const stat = stats[c];
          if (value == null || value === '') { stat.blank += 1; continue; }
          stat.filled += 1;
          const key = String(value);
          if (stat.seen.has(key)) stat.seen.set(key, stat.seen.get(key) + 1);
          else if (stat.seen.size < PROFILE_DISTINCT_CAP) stat.seen.set(key, 1);
          else stat.distinctCapped = true;
          if (AgentCore.isExcelError(value)) {
            stat.errors += 1;
            if (stat.errorCells.length < PROFILE_DISTINCT) stat.errorCells.push(`${indexToColumn(firstColumn + c)}${top + r}`);
          }
          const n = typeof value === 'number' ? value : Number(key);
          if (Number.isFinite(n) && key.trim() !== '') {
            stat.numeric += 1;
            const y = n - stat.correction;
            const next = stat.sum + y;
            stat.correction = (next - stat.sum) - y;
            stat.sum = next;
            stat.min = stat.min == null ? n : Math.min(stat.min, n);
            stat.max = stat.max == null ? n : Math.max(stat.max, n);
          }
        }
      }
      scannedRows += count;
      onProgress(scannedRows, size.rows);
    }
    const columns = stats.map((stat, c) => {
      const ranked = [...stat.seen.entries()].sort((a, b) => b[1] - a[1]);
      return {
        letter: indexToColumn(firstColumn + c), header: header?.[c] || '(blank header)',
        filled: stat.filled, blank: stat.blank,
        distinct: stat.distinctCapped ? `at least ${stat.seen.size}` : stat.seen.size,
        numeric: stat.numeric, sum: stat.numeric ? stat.sum : null,
        min: stat.min, max: stat.max, errors: stat.errors, errorCells: stat.errorCells,
        samples: ranked.slice(0, PROFILE_DISTINCT).map(([value, count]) => `${value} (${count})`)
      };
    });
    return {
      address: `${firstLetter}${firstRow}:${lastLetter}${firstRow + scannedRows - 1}`,
      requestedAddress: addr, columns, headerSource, sampledRows,
      scannedRows, totalRows: size.rows, scannedCols: cols, totalCols: size.cols,
      complete: !scanError && scannedRows === size.rows && cols === size.cols,
      scanError
    };
  });
}

function formatProfile(sheetName, profile) {
  const coverage = `${profile.scannedRows}/${profile.totalRows} requested rows, ${profile.scannedCols}/${profile.totalCols} columns`;
  const head = `[Excel profile ${sheetName}!${profile.address} — ${profile.sampledRows} data rows profiled; ${coverage}; ${profile.complete ? 'complete' : 'INCOMPLETE'}; headers from ${profile.headerSource}]`;
  const lines = profile.columns.map(column => {
    const parts = [`${column.letter} | ${column.header}`, `filled ${column.filled}`, `distinct ${column.distinct}`];
    if (column.blank) parts.push(`blank ${column.blank}`);
    if (column.numeric) parts.push(`numeric ${column.numeric}`, `sum ${column.sum}`, `min ${column.min}`, `max ${column.max}`);
    if (column.errors) parts.push(`errors ${column.errors}`, `at ${column.errorCells.join(', ')}`);
    if (column.samples.length) parts.push(`e.g. ${column.samples.join(', ')}`);
    return parts.join(' | ');
  });
  return `${head}\n${lines.join('\n')}${profile.complete ? '' : `\n[INCOMPLETE: ${sheetName}!${profile.requestedAddress} was requested; only the displayed profile range was read.${profile.scanError ? ` Excel error: ${profile.scanError}.` : ''}]`}`;
}

// Whole-workbook scan: every sheet's used range, batched, error values counted with addresses.
async function scanWorkbook(onProgress = () => {}) {
  return Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const results = [];
    for (const item of sheets.items) {
      const worksheet = context.workbook.worksheets.getItem(item.name);
      const used = worksheet.getUsedRangeOrNullObject();
      used.load(['isNullObject', 'address', 'rowCount', 'columnCount']);
      await context.sync();
      if (used.isNullObject) { results.push({ sheet: item.name, rows: 0, cols: 0, scanned: 0, errors: 0, errorCells: [], scanError: '' }); continue; }
      const addr = plainRange(used.address);
      const size = rangeSize(addr);
      const start = addr.split(':')[0].match(/^([A-Z]+)(\d+)$/);
      const end = (addr.split(':')[1] || addr.split(':')[0]).match(/^([A-Z]+)(\d+)$/);
      const batchRows = Math.max(1, Math.floor(PROFILE_BATCH_CELLS / size.cols));
      let errors = 0, scanned = 0, scanError = '';
      const errorCells = [];
      for (let offset = 0; offset < size.rows; offset += batchRows) {
        const count = Math.min(batchRows, size.rows - offset);
        const top = Number(start[2]) + offset;
        const target = worksheet.getRange(`${start[1]}${top}:${end[1]}${top + count - 1}`);
        target.load('values');
        try { await context.sync(); }
        catch (error) { scanError = String(error?.message || error).slice(0, 200); break; }
        const values = target.values || [];
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < (values[r] || []).length; c++) {
            if (AgentCore.isExcelError(values[r][c])) {
              errors += 1;
              if (errorCells.length < 12) errorCells.push(`${indexToColumn(columnIndex(start[1]) + c)}${top + r}`);
            }
          }
        }
        scanned += count;
        onProgress(item.name, scanned, size.rows);
      }
      results.push({ sheet: item.name, rows: size.rows, cols: size.cols, scanned, errors, errorCells, scanError });
    }
    return results;
  });
}

function formatScan(results) {
  const lines = results.map(r => {
    const parts = [`sheet: ${r.sheet}`, `rows ${r.rows} × cols ${r.cols}`, `scanned ${r.scanned}/${r.rows}`, `errors ${r.errors}`];
    if (r.errorCells.length) parts.push(`at ${r.errorCells.join(', ')}`);
    if (r.scanError || r.scanned < r.rows) parts.push(`INCOMPLETE${r.scanError ? `: ${r.scanError}` : ''}`);
    return parts.join(' | ');
  });
  const totalErrors = results.reduce((n, r) => n + r.errors, 0);
  return `[Workbook scan: ${results.length} sheets, ${totalErrors} error values]\n${lines.join('\n')}`;
}

async function readFullGrid(sheet, range, wantFormulas) {
  const size = rangeSize(range);
  const batchRows = Math.max(1, Math.floor(PROFILE_BATCH_CELLS / size.cols));
  const start = range.split(':')[0].match(/^([A-Z]+)(\d+)$/);
  const end = (range.split(':')[1] || range.split(':')[0]).match(/^([A-Z]+)(\d+)$/);
  const values = [], formulas = [];
  for (let offset = 0; offset < size.rows; offset += batchRows) {
    const count = Math.min(batchRows, size.rows - offset);
    const top = Number(start[2]) + offset;
    const batch = await Excel.run(async context => {
      const cells = context.workbook.worksheets.getItem(sheet).getRange(`${start[1]}${top}:${end[1]}${top + count - 1}`);
      cells.load(wantFormulas ? ['values', 'formulas'] : ['values']);
      await context.sync();
      return { values: cells.values || [], formulas: cells.formulas || [] };
    });
    if (batch.values.length !== count) throw new Error(`Excel returned ${batch.values.length} of ${count} check rows`);
    values.push(...batch.values);
    if (wantFormulas) formulas.push(...batch.formulas);
  }
  return { values, formulas };
}

async function gridsForChecks(checks, phase = 'after') {
  const grids = {};
  for (const check of checks) {
    const target = check.check === 'crossCheck' && phase === 'before'
      ? { sheet: check.otherSheet, range: check.otherRange }
      : { sheet: check.sheet, range: check.range };
    if (!target.sheet || !target.range) { grids[check.id] = { values: [], formulas: [], readError: 'the comparison range is missing' }; continue; }
    try { grids[check.id] = await readFullGrid(target.sheet, target.range, check.check === 'formulaErrors'); }
    catch (error) { grids[check.id] = { values: [], formulas: [], readError: String(error?.message || error).slice(0, 300) }; }
  }
  return grids;
}

function usableChecks(plan) {
  return (plan?.finalValidation || []).map(check => {
    const sheet = normalizeSheet(check.sheet);
    const range = normalizeRange(check.range);
    return sheet && range ? { ...check, sheet, range, otherSheet: normalizeSheet(check.otherSheet), otherRange: normalizeRange(check.otherRange) } : null;
  }).filter(Boolean);
}

async function routeForPrompt(text, mode) {
  let configured = { qwen: false };
  let relayBuild = '';
  try {
    const health = await api('health');
    configured = health.configured || configured;
    relayBuild = typeof health.build === 'string' ? health.build : '';
  } catch { /* the send path reports a down relay */ }
  // Hosted is BYO: readiness is the in-browser key, not the server's (always empty) configured flag.
  if (hostedRelay()) configured = { qwen: Boolean(byoKey()) };
  const routed = AgentCore.routeProvider({ selected: 'qwen', ready: configured });
  if (routed.error) return routed;
  return { ...routed, model: selectedModel(), endpoint: $('qwenEndpoint').value.trim(), relayBuild, buildMismatch: Boolean(relayBuild) && relayBuild !== PANE_BUILD };
}

function renderAuditCard() {
  // The trail stays in the saved session. The pane does not show an AUDIT card.
}

function validationEvidence(result) {
  return (result.details || []).map(detail => {
    const relative = String(detail).match(/^R(\d+)C(\d+)$/);
    if (!relative || !result.sheet || !result.range) return detail;
    return `${result.sheet}!${cellAddress(result.range, Number(relative[1]) - 1, Number(relative[2]) - 1)}`;
  });
}

function renderValidation(results) {
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant edits-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  const failed = results.filter(result => result.status === 'FAIL');
  label.textContent = failed.length ? `VALIDATION FAILED · ${failed.length}` : `VALIDATION · ${results.length}`;
  card.append(label);
  results.forEach(result => {
    const row = document.createElement('div');
    row.className = result.status === 'FAIL' ? 'edit-warning' : 'edit-note';
    const location = result.sheet && result.range ? ` · ${result.sheet}!${result.range}` : '';
    const tolerance = Number.isFinite(result.tolerance) ? ` · tolerance ${result.tolerance}` : '';
    row.textContent = `${result.status} ${result.check}${location}: before ${result.before} after ${result.after} difference ${result.difference}${tolerance}`;
    if (result.details?.length) {
      const detail = document.createElement('div');
      detail.className = 'edit-note';
      const evidence = validationEvidence(result);
      detail.textContent = `Evidence: ${evidence.slice(0, 5).join('; ')}${evidence.length > 5 ? `; +${evidence.length - 5} more` : ''}`;
      card.append(row, detail);
    } else card.append(row);
    if (result.status === 'FAIL' || result.status === 'NEEDS_USER') {
      const explained = AgentCore.explainFailure(result);
      const note = document.createElement('div');
      note.className = 'edit-warning';
      note.textContent = `${explained.issue} — ${explained.possibleCause} ${explained.recommendedAction}`;
      card.append(note);
    }
  });
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

function renderFinanceEvidenceCard(sources, validations = [], answer = '') {
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant edits-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  label.textContent = 'FINANCE EVIDENCE · 財務證據';
  card.append(label);

  const sourceLabel = document.createElement('div');
  sourceLabel.className = 'edit-note';
  sourceLabel.textContent = 'Sources sent to the model · 傳給模型的來源';
  card.append(sourceLabel);
  if (sources.length) {
    sources.forEach(source => {
      const row = document.createElement('div');
      row.className = 'edit-row';
      row.textContent = source;
      card.append(row);
    });
  } else {
    const row = document.createElement('div');
    row.className = 'edit-warning';
    row.textContent = 'No workbook or report data was read; this answer is not source-verified. · 未讀取工作簿或報表資料；此答案未經來源核實。';
    card.append(row);
  }

  const citationLabel = document.createElement('div');
  citationLabel.className = 'edit-note';
  citationLabel.textContent = 'Answer citation locations · 答案中的來源位置';
  card.append(citationLabel);
  const citations = financeCitationCheck(answer, sources);
  if (!citations.length) {
    const row = document.createElement('div');
    row.className = 'edit-warning';
    row.textContent = 'No supported source citations were detected in the answer. · 答案中未偵測到可比對的來源標記。';
    card.append(row);
  } else {
    citations.forEach(item => {
      const row = document.createElement('div');
      row.className = item.matched ? 'edit-row' : 'edit-warning';
      row.textContent = `${item.matched ? 'Location matched' : 'Location not found in supplied context'} · ${item.citation}`;
      card.append(row);
    });
    const note = document.createElement('div');
    note.className = 'edit-note';
    note.textContent = 'Location matching does not confirm that the cited source supports the claim. · 位置比對不代表該來源必然支持該結論。';
    card.append(note);
  }

  if (validations.length) {
    const checkLabel = document.createElement('div');
    checkLabel.className = 'edit-note';
    checkLabel.textContent = 'Add-in calculation checks · 增益集計算檢查';
    card.append(checkLabel);
    validations.forEach(result => {
      const row = document.createElement('div');
      row.className = result.status === 'FAIL' ? 'edit-warning' : 'edit-row';
      const location = result.sheet && result.range ? ` · ${result.sheet}!${result.range}` : '';
      row.textContent = `${result.status} ${result.check}${location} · before ${result.before} · after ${result.after} · difference ${result.difference}`;
      card.append(row);
    });
  } else {
    const row = document.createElement('div');
    row.className = 'edit-note';
    row.textContent = 'The add-in did not independently validate this answer’s arithmetic. · 本增益集沒有獨立驗證此答案的計算。';
    card.append(row);
  }
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

async function executeClassified(actions, { direct }) {
  const blocked = [];
  const safe = [];
  for (const action of actions) {
    if (!direct || actionIsRisky(action) || AgentCore.isDestructive(action)) blocked.push(action);
    else safe.push(action);
  }
  let status = blocked.length ? 'BLOCKED' : 'SUCCESS';
  let failure = '';
  let appliedActions = [];
  if (safe.length) {
    try {
      const records = await applyPlan(safe);
      appliedActions = safe;
      status = blocked.length ? 'WARNING' : 'SUCCESS';
      // Without this card the agent loop applied changes with no Undo offered at all.
      renderAppliedCard(safe, records, blocked.length);
    } catch (error) {
      status = 'FAILED';
      failure = String(error?.message || error);
      toast(`Could not revise: ${failure}`);
    }
  }
  if (blocked.length) await renderPlanCard(blocked, 0, { manualOnly: true });
  return { status, blocked: blocked.length, appliedActions, failure };
}

function renderAppliedCard(actions, records, blockedCount) {
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant edits-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  label.textContent = blockedCount
    ? `REVISED ${actions.length} · ${blockedCount} STILL NEED A CLICK`
    : `REVISED ${actions.length} CHANGE${actions.length === 1 ? '' : 'S'}`;
  card.append(label);
  actions.forEach(action => {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.textContent = describeAction(action);
    card.append(row);
  });
  const buttonRow = document.createElement('div');
  buttonRow.className = 'button-row';
  const noun = actions.some(action => action.host === 'word') && !actions.some(action => action.host === 'excel') ? 'document' : 'workbook';
  if (records.length) attachUndo(buttonRow, label, records, noun);
  card.append(buttonRow);
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

function renderModePlan(plan, actions, skipped) {
  const conversation = $('conversation');
  const card = document.createElement('div');
  card.className = 'message assistant edits-card';
  const label = document.createElement('div');
  label.className = 'message-label';
  label.textContent = plan ? `PLAN · ${plan.goal}` : 'PLAN · APPROVAL REQUIRED';
  card.append(label);
  (plan?.steps || []).forEach(step => {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.textContent = `${step.id}. ${step.objective}${step.requiresApproval ? ' · approval' : ''}`;
    card.append(row);
  });
  for (const [title, items] of [['Assumptions', plan?.assumptions], ['Inputs', plan?.requiredInputs], ['Risks', plan?.risks]]) {
    if (!items?.length) continue;
    const row = document.createElement('div');
    row.className = 'edit-note';
    row.textContent = `${title}: ${items.join('; ')}`;
    card.append(row);
  }
  actions.forEach(action => {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.textContent = describeAction(action);
    card.append(row);
  });
  if (skipped) {
    const note = document.createElement('div');
    note.className = 'edit-note';
    note.textContent = `${skipped} change${skipped === 1 ? '' : 's'} skipped.`;
    card.append(note);
  }
  const buttonRow = document.createElement('div');
  buttonRow.className = 'button-row';
  const approve = document.createElement('button');
  approve.className = 'primary';
  approve.textContent = 'Approve';
  const dismiss = document.createElement('button');
  dismiss.className = 'secondary';
  dismiss.textContent = 'Dismiss';
  approve.addEventListener('click', async () => {
    approve.disabled = true;
    dismiss.disabled = true;
    if (!actions.length) { toast('The plan has no actions to run.'); return; }
    const result = await executeClassified(actions, { direct: Boolean($('directRevise').checked) });
    state.audit.push(AgentCore.audit({ kind: 'execute', status: result.status, count: actions.length, blocked: result.blocked }));
    label.textContent = result.status === 'SUCCESS' ? 'PLAN APPROVED · APPLIED' : `PLAN APPROVED · ${result.status}`;
    renderAuditCard();
  });
  dismiss.addEventListener('click', () => card.remove());
  buttonRow.append(approve, dismiss);
  card.append(buttonRow);
  conversation.append(card);
  conversation.scrollTop = conversation.scrollHeight;
}

async function sendAgentPrompt(prompt, checkpoint = null) {
  const clean = prompt.trim();
  if (!clean || state.busy) return;
  if (clean.length > 49_500) { toast('Message is too large. Shorten the prompt and try again.'); return; }
  const turn = ++state.turn;
  const mode = operatingMode();
  const run = mode === 'agent' ? { round: 0, stopRequested: false } : null;
  const task = checkpoint || { prompt: clean, progress: [], observation: '', observationSource: '' };
  activeAgent = run;
  if (!state.sessionId) { state.sessionId = `s-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`; state.sessionStarted = new Date().toISOString(); }
  if (!checkpoint) agentCheckpoint = null;
  state.busy = true;
  if (run) $('newChat').disabled = true;
  showAgentStatus(0, run ? 'Preparing · 準備中' : '');
  $('send').textContent = '■';
  $('send').setAttribute?.('aria-label', 'Stop');
  const interrupted = () => turn !== state.turn || run?.stopRequested;
  let outcome = 'Could not finish · 未能完成';
  let answerForWord = '';
  const observe = (content, source = '') => {
    task.observation = content;
    task.observationSource = source;
    state.messages.push({ role: 'user', content });
  };
  const failEarly = reason => {
    outcome = reason;
    toast(explainError(reason));
    state.audit.push(AgentCore.audit({ kind: 'blocked', reason }));
    renderAuditCard();
    if (hostedRelay() && /Qwen API key|Qwen key/.test(reason)) {
      $('settings').classList.remove('hidden');
      refreshSettings();
      $('apiKey').focus();
    }
  };
  try {
    await ensureKeys();
    if (interrupted()) return;
    let routed = await routeForPrompt(clean, mode);
    if (interrupted()) return;
    if (routed.error) { failEarly(routed.error); return; }
    if (routed.buildMismatch) { failEarly(`The relay is running build ${routed.relayBuild} but this pane is ${PANE_BUILD}. Restart the relay (npm run relay-stop, then npm start) and send again.`); return; }
    const displayPrompt = checkpoint ? 'Continue Agent · 繼續 Agent' : clean;
    state.audit.push(AgentCore.audit({ kind: 'request', mode, provider: routed.provider, reason: checkpoint ? 'continued after round limit' : routed.reason, text: displayPrompt.slice(0, 500) }));
    const files = attachmentContext();
    state.lastPrompt = clean;
    let skillCtx = null;
    try { if ($('includeContext').checked) skillCtx = await api(`skills/context?question=${encodeURIComponent(clean)}`); } catch { /* saved skills are optional */ }
    if (interrupted()) return;
    const skillText = skillCtx?.text || '';
    let memoryText = '';
    try { if ($('includeContext').checked) memoryText = String((await api('memory')).text || '').slice(-4000); } catch { /* memory is optional */ }
    if (interrupted()) return;
    const resumeContext = checkpoint ? [
      '[Agent continuation] Continue the original task from the preceding user message. Use the conversation and the pane observations below. Re-read the affected workbook range or Word selection before editing if the current state is uncertain. Avoid repeating edits already applied. Review pending decisions and report any remaining work.',
      task.progress.length ? `Recent edits confirmed by the pane at the time of application (verify if the user has since edited or undone them):\n${task.progress.join('\n')}` : '',
      task.observation ? `Latest pane observation:\n${task.observation}` : ''
    ].filter(Boolean).join('\n\n') : '';
    const contextLimit = Math.max(0, Math.min(CONTEXT_CAP, 49_000 - clean.length - files.length - skillText.length - memoryText.length - FINANCE_EVIDENCE_INSTRUCTION.length - 300));
    if (contextLimit < 300 && $('includeContext').checked) { failEarly('The prompt and attachments leave too little room for workbook context. Remove a file or shorten the prompt.'); return; }
    let ctx = null;
    try { ctx = await contextForMode(mode, contextLimit, clean); } catch (error) { if (turn === state.turn) toast(`Could not read workbook context: ${error.message}`); }
    if (interrupted()) return;
    const financeTask = AgentCore.isFinancePrompt(`${clean}\n${files}\n${ctx?.text || ''}`);
    const evidenceSources = collectFinanceSources(ctx, files);
    if (checkpoint?.observationSource) evidenceSources.push(checkpoint.observationSource);
    if (skillCtx?.selected?.length) evidenceSources.push(`Saved skills: ${skillCtx.selected.join(', ')}`);
    const financeValidations = [];
    let financeAnswerText = '';
    const userPrompt = financeTask ? `${clean}${FINANCE_EVIDENCE_INSTRUCTION}` : clean;
    const repoNote = githubRepoNote();
    const content = `${userPrompt}${files}${ctx ? `\n\n${ctx.header}\n${ctx.text}` : ''}${skillText ? `\n\n[Saved skills]\n${skillText}` : ''}${memoryText ? `\n\n[Memory]\n${memoryText}` : ''}${repoNote}\n\n[Mode: ${mode}]`;
    if (content.length > 49_500) { failEarly('Message plus workbook context is too large — shorten the message or turn context off.'); return; }
    agentCheckpoint = null;
    state.forceSheet = false;
    state.messages.push({ role: 'user', content });
    if (resumeContext) state.messages.push({ role: 'user', content: resumeContext });
    const fileNames = state.attachments.map(file => file.name);
    const skillNames = skillCtx?.selected?.length ? `\n\nSkills: ${skillCtx.selected.join(', ')}` : '';
    renderMessage('user', displayPrompt + (fileNames.length ? `\n\nAttachments: ${fileNames.join(', ')}` : '') + skillNames + (selectedGitHubRepo() && repoNote ? `\n\nGitHub: ${selectedGitHubRepo()}` : '') + (ctx ? `\n\nContext: ${ctx.address}${ctx.truncated ? ' (truncated)' : ''}` : ''));
    state.audit.push(AgentCore.audit({ kind: 'context', address: ctx?.address || 'none', chars: ctx?.chars || 0 }));
    if (skillText) state.audit.push(AgentCore.audit({ kind: 'context', address: `skills ${skillCtx.selected.join(', ')}`, chars: skillText.length }));
    if (skillCtx?.selected?.length) renderAuditCard();
    if (memoryText) state.audit.push(AgentCore.audit({ kind: 'context', address: 'memory', chars: memoryText.length }));
    if (state.attachments.length) state.audit.push(AgentCore.audit({ kind: 'context', address: fileNames.join(', '), chars: files.length }));
    let guard = 0;
    while (guard < AgentCore.MAX_LOOPS) {
      if (interrupted()) return;
      if (run) {
        run.round = guard + 1;
        showAgentStatus(run.round, 'Waiting for model · 等候模型');
      }
      const messages = dedupContext(state.messages.slice(-24), resumeContext && guard === 0 ? 2 : 1);
      const fields = chatFields();
      const result = await streamChat({ provider: routed.provider, model: fields.model, endpoint: routed.endpoint, messages, fast: fields.fast, effort: fields.effort, host: state.host === 'Word' ? 'word' : 'excel', mode });
      if (interrupted()) return;
      const envelope = AgentCore.parseEnvelope(result.text);
      const revision = parseRevision(result.text);
      financeAnswerText = envelope.prose || revision.prose;
      state.messages.push({ role: 'assistant', content: result.text });
      state.lastAnswer = envelope.prose || revision.prose;
      if (state.lastAnswer) answerForWord = state.lastAnswer;
      $('insertAnswer').disabled = !state.lastAnswer;
      $('lastAnswerState').textContent = result.model ? `via ${result.model}` : '';
      result.bubble.children[result.bubble.children.length - 1].textContent = state.lastAnswer || '(model returned only a JSON envelope; see the audit card)';
      state.audit.push(AgentCore.audit({ kind: 'model', round: run?.round, provider: routed.provider, model: result.model, decision: envelope.decision || 'COMPLETE', plan: envelope.plan?.goal || '' }));
      const normalized = envelope.actions ? normalizeActions(envelope.actions) : { actions: revision.actions || [], skipped: revision.skipped || 0 };
      if (!$('includeContext').checked && (envelope.profileRequest || envelope.contextRequest || envelope.scanRequest || envelope.sourceRequest || envelope.githubRequest || envelope.githubWrite || usableChecks(envelope.plan).length)) {
        observe('Office file sharing is off. No workbook or document range was read or sent to AI. Turn on Share Office file data with AI to review the open file.');
        state.audit.push(AgentCore.audit({ kind: 'blocked', reason: 'Office file sharing is off' }));
        outcome = 'File sharing off · 檔案分享已關閉';
        break;
      }
      if (envelope.githubWrite) {
        if (hostedRelay()) {
          observe('GitHub MCP is available on the local relay on this PC. Nothing was written.');
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: 'GitHub MCP needs the local relay' }));
          outcome = 'Local relay required · 需要本機 relay';
          break;
        }
        renderGitHubWrite(envelope.githubWrite);
        state.audit.push(AgentCore.audit({ kind: 'stopped', reason: 'GitHub write is waiting for a click' }));
        outcome = 'Review required · 等候審批';
        break;
      }
      if (mode === 'plan' || envelope.decision === 'NEEDS_USER' || (mode === 'agent' && !$('directRevise').checked && normalized.actions.length)) {
        renderModePlan(envelope.plan, normalized.actions, normalized.skipped);
        outcome = normalized.actions.length ? 'Review required · 等候審批' : 'Input required · 等候輸入';
        break;
      }
      if (envelope.sourceRequest && !normalized.actions.length) {
        if (hostedRelay()) {
          observe('OneDrive search and file reading are available on the local relay on this PC. Nothing was read.');
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: 'OneDrive needs the local relay' }));
          outcome = 'Local relay required · 需要本機 relay';
          break;
        }
        showAgentStatus(run?.round || 1, envelope.sourceRequest.action === 'search' ? 'Searching OneDrive · 搜尋 OneDrive' : 'Reading OneDrive file · 讀取 OneDrive 檔案');
        try {
          if (envelope.sourceRequest.action === 'search') {
            const result = await api('connectors/search', { source: 'onedrive', query: envelope.sourceRequest.query });
            const items = Array.isArray(result.items) ? result.items : [];
            const lines = items.map((item, index) => `${index + 1}. ${item.name} | id: ${item.id}${item.detail ? ` | ${item.detail}` : ''}`);
            observe(`OneDrive search for "${envelope.sourceRequest.query}".\n${lines.join('\n') || 'No files matched.'}\nNothing was read yet. To read one file, set sourceRequest to {"action":"read","id":"...","name":"file.xlsx"}.`, 'OneDrive search');
            state.audit.push(AgentCore.audit({ kind: 'context', address: `OneDrive search ${envelope.sourceRequest.query}`, chars: lines.join('\n').length }));
          } else {
            const result = await api('connectors/read', { source: 'onedrive', item: { id: envelope.sourceRequest.id, name: envelope.sourceRequest.name }, endpoint: $('qwenEndpoint').value.trim() });
            const clipped = trimContext(String(result.text || ''), 8000);
            const label = result.name || envelope.sourceRequest.name;
            observe(`Requested OneDrive file ${label}.\n\n${clipped.text}`, label);
            evidenceSources.push(label);
            state.audit.push(AgentCore.audit({ kind: 'context', address: label, chars: clipped.text.length }));
          }
        } catch (error) {
          observe(`No OneDrive file was read: ${error.message}. Nothing was returned, so this is not a successful read.`);
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: `onedrive rejected: ${String(error.message || error).slice(0, 240)}` }));
        }
        if (interrupted()) return;
        guard += 1;
        continue;
      }
      if (envelope.githubRequest && !normalized.actions.length) {
        if (hostedRelay()) {
          observe('GitHub MCP is available on the local relay on this PC. Nothing was read or written.');
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: 'GitHub MCP needs the local relay' }));
          outcome = 'Local relay required · 需要本機 relay';
          break;
        }
        showAgentStatus(run?.round || 1, envelope.githubRequest.action === 'repos' ? 'Listing GitHub repositories · 列出 GitHub repo' : 'Reading GitHub file · 讀取 GitHub 檔案');
        try {
          if (envelope.githubRequest.action === 'repos') {
            const result = await api('connectors/github/repos', githubAuthBody({ query: envelope.githubRequest.query }));
            const items = Array.isArray(result.items) ? result.items : [];
            const lines = items.map((item, index) => `${index + 1}. ${item.fullName}${item.private ? ' (private)' : ''}${item.defaultBranch ? ` | default: ${item.defaultBranch}` : ''}`);
            const label = envelope.githubRequest.query ? `GitHub repos matching "${envelope.githubRequest.query}"` : 'GitHub repositories';
            observe(`${label}.\n${lines.join('\n') || 'No repositories matched.'}\nNothing was written. To read one file, set githubRequest to {"action":"read","repo":"owner/repo","path":"README.md"}.`, label);
            state.audit.push(AgentCore.audit({ kind: 'context', address: label, chars: lines.join('\n').length }));
          } else {
            const result = await api('connectors/github/read', githubAuthBody({ repo: envelope.githubRequest.repo, path: envelope.githubRequest.path }));
            const clipped = trimContext(String(result.text || ''), 8000);
            const label = `${envelope.githubRequest.repo}/${envelope.githubRequest.path}`;
            const note = result.truncated || clipped.truncated ? '\nThe file was truncated.' : '';
            observe(`GitHub file ${label}.\n\n${clipped.text}${note}`, label);
            state.audit.push(AgentCore.audit({ kind: 'context', address: label, chars: clipped.text.length }));
          }
        } catch (error) {
          observe(`No GitHub file was read: ${error.message}. Nothing was returned, so this is not a successful read.`);
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: `github rejected: ${String(error.message || error).slice(0, 240)}` }));
        }
        if (interrupted()) return;
        guard += 1;
        continue;
      }
      if (envelope.scanRequest && !normalized.actions.length) {
        showAgentStatus(run?.round || 1, 'Scanning workbook · 掃描活頁簿');
        let results = null;
        let failure = '';
        try {
          results = await scanWorkbook((sheet, done, total) => {
            if (interrupted()) { const stopped = new Error('Stopped'); stopped.name = 'AbortError'; throw stopped; }
            showAgentStatus(run?.round || 1, `Scanning ${sheet} ${done}/${total} · 掃描工作表`);
          });
        } catch (error) { failure = String(error?.message || error).slice(0, 300); }
        if (interrupted()) return;
        if (results) {
          const clipped = trimContext(formatScan(results), 8000);
          observe(`Requested workbook scan.\n\n${clipped.text}`, 'Workbook scan · all sheets');
          evidenceSources.push('Workbook scan · all sheets');
          state.audit.push(AgentCore.audit({ kind: 'context', address: 'workbook scan', chars: clipped.text.length }));
        } else {
          observe(`The workbook scan failed: ${failure}. Nothing was read, so this is not a successful read.`);
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: `workbook scan failed: ${failure}` }));
        }
        guard += 1;
        continue;
      }
      if (envelope.profileRequest && !normalized.actions.length) {
        showAgentStatus(run?.round || 1, 'Reading sheet profile · 讀取工作表摘要');
        const sheet = normalizeSheet(envelope.profileRequest.sheet);
        let profile = null;
        let failure = '';
        if (!sheet) failure = `sheet "${envelope.profileRequest.sheet}" is not a usable sheet name`;
        else {
          try {
            profile = await profileSheet(sheet, envelope.profileRequest.range, (done, total) => {
              if (interrupted()) { const stopped = new Error('Stopped'); stopped.name = 'AbortError'; throw stopped; }
              showAgentStatus(run?.round || 1, `Profiling ${done}/${total} rows · 掃描工作表`);
            });
          }
          catch (error) { failure = error.message; }
          if (!profile && !failure) failure = 'the sheet or range is empty';
        }
        if (interrupted()) return;
        if (profile) {
          const clipped = trimContext(formatProfile(sheet, profile), 8000);
          const text = clipped.text;
          const address = `${sheet}!${profile.address}`;
          const source = `${address} · profile statistics${profile.complete ? '' : ' · incomplete'}${clipped.truncated ? ' · output truncated' : ''}`;
          observe(`Requested profile.\n\n${text}`, source);
          evidenceSources.push(source);
          state.audit.push(AgentCore.audit({ kind: 'context', address: `${address} (profile)`, chars: text.length }));
        } else {
          observe(`No profile was produced for "${envelope.profileRequest.sheet}": ${failure}. Nothing was read, so this is not a successful read.`);
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: `profile rejected: ${failure}` }));
        }
        guard += 1;
        continue;
      }
      if (envelope.contextRequest && !normalized.actions.length) {
        showAgentStatus(run?.round || 1, 'Reading workbook · 讀取活頁簿');
        const sheet = normalizeSheet(envelope.contextRequest.sheet);
        const asked = normalizeRange(envelope.contextRequest.range);
        const page = asked ? pagedRange(asked, envelope.contextRequest.offset) : null;
        const range = page ? page.range : '';
        if (sheet && range) {
          let grid;
          try { grid = await readGrid(sheet, range); }
          catch (error) {
            if (interrupted()) return;
            const reason = String(error?.message || error).slice(0, 300);
            observe(`The requested range ${sheet}!${range} could not be read: ${reason}. Nothing was read. Retry the range once or use a smaller range; report any remaining gap explicitly.`);
            state.audit.push(AgentCore.audit({ kind: 'blocked', reason: `context read failed: ${reason}` }));
            guard += 1;
            continue;
          }
          if (interrupted()) return;
          const full = annotateRows(grid.values, grid.formulas, `${sheet}!${range}`);
          const text = full.slice(0, 8000);
          const cut = full.length > text.length ? '\n[truncated: the read exceeded the context budget]' : '';
          const consumed = page.skip + page.rows;
          const coverage = `\n[Coverage: ${consumed}/${page.totalRows} rows of ${sheet}!${asked} read (rows ${page.firstRow + page.skip}-${page.firstRow + consumed - 1}). ${consumed >= page.totalRows ? 'The requested range is fully read.' : `Send contextRequest with offset=${consumed} for the next chunk.`}]`;
          const note = `Requested context.\n\n[Excel range ${sheet}!${range}]\n${text}${cut}${coverage}`;
          const source = `${sheet}!${range}${cut ? ' · truncated' : ''}`;
          observe(note, source);
          evidenceSources.push(source);
          state.audit.push(AgentCore.audit({ kind: 'context', address: `${sheet}!${range}`, chars: text.length }));
          guard += 1;
          continue;
        }
        // State why nothing was read. Falling through used to report "Observed: SUCCESS",
        // which the model took as a successful but empty read, then refused to guess columns.
        const reason = !sheet
          ? `sheet "${envelope.contextRequest.sheet}" is not a usable sheet name`
          : `range "${envelope.contextRequest.range}" is not a valid A1 range`;
        observe(`No range was read: ${reason}. Nothing was returned, so this is not a successful read. Request a smaller range and ask again.`);
        state.audit.push(AgentCore.audit({ kind: 'blocked', reason: `context request rejected: ${reason}` }));
        guard += 1;
        continue;
      }
      const checks = usableChecks(envelope.plan);
      const profile = readFinanceProfile();
      if (checks.length) showAgentStatus(run?.round || 1, 'Reading check inputs · 讀取檢查資料');
      const beforeGrids = await gridsForChecks(checks, 'before');
      if (interrupted()) return;
      let exec = { status: 'SUCCESS', blocked: 0, appliedActions: [], failure: '' };
      if (normalized.actions.length || revision.edits) {
        if (mode === 'chat') {
          state.audit.push(AgentCore.audit({ kind: 'blocked', reason: 'chat mode does not write' }));
          toast('Chat mode does not change the file.');
          outcome = 'Chat read-only · Chat 唯讀';
          break;
        }
        if (normalized.actions.length) {
          showAgentStatus(run?.round || 1, 'Applying edits · 套用修改');
          exec = await executeClassified(normalized.actions, { direct: Boolean($('directRevise').checked) });
        }
      }
      state.audit.push(AgentCore.audit({ kind: 'execute', status: exec.status, count: normalized.actions.length, blocked: exec.blocked, skipped: normalized.skipped, reason: exec.failure }));
      for (const action of exec.appliedActions) task.progress.push(describeAction(action).slice(0, 240));
      task.progress = task.progress.slice(-12);
      if (interrupted()) return;
      if (normalized.skipped) {
        state.messages.push({ role: 'user', content: `${normalized.skipped} proposed action(s) were rejected by the pane as unusable and were NOT applied. Re-check field names, counts and ranges: a pivot accepts at most 16 value fields and 8 fields each for rows, columns and filters.` });
      }
      if (exec.failure) {
        // The model must see the real Office.js error, not a bare "Observed: FAILED".
        observe(`The pane reported an Office error applying that batch: ${exec.failure}. Re-read affected locations to determine what actually changed before proposing a correction. Do not repeat the same action unchanged.`);
        if (exec.blocked) {
          outcome = 'Edit failed · 修改失敗';
          break;
        }
        guard += 1;
        continue;
      }
      if (checks.length) showAgentStatus(run?.round || 1, 'Checking results · 檢查結果');
      const afterGrids = await gridsForChecks(checks, 'after');
      if (interrupted()) return;
      const validations = checks.map(check => ({
        ...AgentCore.evaluateGridCheck({ ...check, profile }, afterGrids[check.id], beforeGrids[check.id]),
        sheet: check.sheet,
        range: check.range,
        tolerance: Number.isFinite(check.tolerance) ? check.tolerance : profile.tolerance
      }));
      if (financeTask) financeValidations.push(...validations);
      validations.forEach(result => state.audit.push(AgentCore.audit({ kind: 'validation', status: result.status, check: result.check, address: `${result.sheet}!${result.range}`, before: result.before, after: result.after, difference: result.difference })));
      if (validations.length) renderValidation(validations);
      const next = AgentCore.decide({ decision: envelope.decision, validations, iteration: guard, maxLoops: AgentCore.MAX_LOOPS });
      if (exec.blocked && !validations.some(result => ['FAIL', 'NEEDS_USER'].includes(result.status)) && envelope.decision !== 'FAILED') {
        observe(`${exec.blocked} proposed edit(s) are waiting for the user to review and have not been applied. Check the current state before proposing further edits.`);
        state.audit.push(AgentCore.audit({ kind: 'stopped', reason: 'edits are waiting for review' }));
        outcome = 'Review required · 等候審批';
        break;
      }
      const observed = validations.map(result => `${result.check} ${result.status} at ${result.sheet}!${result.range}`).join('; ') || exec.status;
      if (next === 'FAILED') {
        const validationFailed = validations.some(result => result.status === 'FAIL');
        const limitReached = ['CONTINUE', 'REPLAN'].includes(envelope.decision)
          && guard + 1 >= AgentCore.MAX_LOOPS;
        if (limitReached && !validationFailed) {
          observe(`Observed: ${observed}. Continue only if every check passed.`);
          guard += 1;
          break;
        }
        const reason = validationFailed
          ? 'validation failed'
          : 'the model reported that it could not complete this step';
        state.audit.push(AgentCore.audit({ kind: 'stopped', reason }));
        outcome = validationFailed ? 'Check failed · 檢查失敗' : 'Model reported a failure · 模型回報失敗';
        toast(`Stopped: ${reason}.`);
        break;
      }
      if (next === 'NEEDS_USER') {
        state.audit.push(AgentCore.audit({ kind: 'stopped', reason: 'a check needs a column' }));
        toast('Stopped. A finance check needs a column before it can pass or fail.');
        outcome = 'Input required · 等候輸入';
        break;
      }
      if (next === 'CONTINUE' || next === 'REPLAN') {
        if (mode === 'chat' && !checks.length && !envelope.contextRequest && !envelope.profileRequest && !envelope.scanRequest && !envelope.sourceRequest) {
          outcome = 'Completed · 已完成';
          break;
        }
        observe(`Observed: ${observed}. Continue only if every check passed.`);
        guard += 1;
        continue;
      }
      outcome = 'Completed · 已完成';
      break;
    }
    if (guard >= AgentCore.MAX_LOOPS) {
      agentCheckpoint = task;
      outcome = 'Paused at round limit · 已達輪數上限';
      state.audit.push(AgentCore.audit({ kind: 'stopped', reason: `${AgentCore.MAX_LOOPS}-round limit reached` }));
      toast('Round limit reached. Review the audit card, then choose Continue Agent.');
    }
    renderAuditCard();
    if (financeTask) renderFinanceEvidenceCard([...new Set(evidenceSources)], financeValidations, financeAnswerText);
    refreshContextCard();
  } catch (error) {
    const bodyEl = error.bubble?.parent ? error.bubble.children[error.bubble.children.length - 1] : null;
    if (error.name === 'AbortError') {
      if (bodyEl) bodyEl.textContent = error.partial ? `${error.partial}\n\n(stopped)` : '(stopped)';
      if (error.bubble?.parent) toast('Stopped.');
      return;
    }
    if (turn !== state.turn) return;
    if (!state.lastAnswer) $('insertAnswer').disabled = true;
    const explained = explainError(error.message);
    if (bodyEl) {
      bodyEl.textContent = error.partial ? `${error.partial}\n\n(${explained})` : `Error: ${explained}`;
      toast(explained);
    } else {
      renderMessage('assistant', `Error: ${explained}`);
      toast(explained);
    }
    outcome = 'Error · 錯誤';
    state.audit.push(AgentCore.audit({ kind: 'stopped', reason: String(error.message || error).slice(0, 300) }));
    renderAuditCard();
  } finally {
    if (turn === state.turn) {
      if (run?.stopRequested) {
        agentCheckpoint = null;
        outcome = 'Stopped by you · 已手動停止';
        state.audit.push(AgentCore.audit({ kind: 'stopped', reason: 'user pressed Stop' }));
        renderAuditCard();
        refreshContextCard();
      }
      activeAgent = null;
      if (run) $('newChat').disabled = false;
      resetComposer();
      if (answerForWord && state.host === 'Word' && !run?.stopRequested) {
        try { await appendWordAnswer(answerForWord); }
        catch (error) { toast(`Could not write the answer into Word: ${error.message}`); }
      }
      state.audit.push(AgentCore.audit({ kind: 'stopped', reason: `run ended at round ${run?.round ?? 0}: ${outcome}` }));
      renderAuditCard();
      persistSession(outcome);
      refreshHistory();
      showAgentStatus(run?.round ?? 0, run ? outcome : '');
    }
  }
}

async function sendPrompt(prompt) {
  if (operatingMode() !== 'legacy') return sendAgentPrompt(prompt);
  const clean = prompt.trim();
  if (!clean || state.busy) return;
  if (clean.length > 49_500) { toast('Message is too large. Shorten the prompt and try again.'); return; }
  const turn = ++state.turn;
  if (!state.sessionId) { state.sessionId = `s-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`; state.sessionStarted = new Date().toISOString(); }
  state.busy = true;
  $('send').textContent = '■';
  $('send').setAttribute?.('aria-label', 'Stop');
  let answerForWord = '';
  try {
    await ensureKeys();
    if (turn !== state.turn) return;
    const provider = 'qwen';
    const files = attachmentContext();
    state.lastPrompt = clean;
    let skillCtx = null;
    try { if ($('includeContext').checked) skillCtx = await api(`skills/context?question=${encodeURIComponent(clean)}`); } catch { /* saved skills are optional */ }
    if (turn !== state.turn) return;
    const skillText = skillCtx?.text || '';
    let memoryText = '';
    try { if ($('includeContext').checked) memoryText = String((await api('memory')).text || '').slice(-4000); } catch { /* memory is optional */ }
    if (turn !== state.turn) return;
    const contextLimit = Math.max(0, Math.min(CONTEXT_CAP, 49_000 - clean.length - files.length - skillText.length - memoryText.length - FINANCE_EVIDENCE_INSTRUCTION.length - 100));
    if (contextLimit < 300 && $('includeContext').checked) { toast('The prompt and attachments leave too little room for workbook context. Remove a file or shorten the prompt.'); return; }
    let ctx = null;
    try { ctx = await buildContext(contextLimit, clean); } catch (error) { if (turn === state.turn) toast(`Could not read workbook context: ${error.message}`); }
    if (turn !== state.turn) return;
    const financeTask = AgentCore.isFinancePrompt(`${clean}\n${files}\n${ctx?.text || ''}`);
    const evidenceSources = collectFinanceSources(ctx, files);
    const repoNote = githubRepoNote();
    const content = `${clean}${financeTask ? FINANCE_EVIDENCE_INSTRUCTION : ''}${files}${ctx ? `\n\n${ctx.header}\n${ctx.text}` : ''}${skillText ? `\n\n[Saved skills]\n${skillText}` : ''}${memoryText ? `\n\n[Memory]\n${memoryText}` : ''}${repoNote}`;
    if (content.length > 49_500) { toast('Message plus workbook context is too large — shorten the message or turn context off.'); return; }
    state.forceSheet = false;
    state.messages.push({ role: 'user', content });
    const fileNames = state.attachments.map(file => file.name);
    const skillNames = skillCtx?.selected?.length ? `\n\nSkills: ${skillCtx.selected.join(', ')}` : '';
    renderMessage('user', clean + (fileNames.length ? `\n\nAttachments: ${fileNames.join(', ')}` : '') + skillNames + (selectedGitHubRepo() && repoNote ? `\n\nGitHub: ${selectedGitHubRepo()}` : '') + (ctx ? `\n\nContext: ${ctx.address}${ctx.truncated ? ' (truncated)' : ''}` : ''));
    const fields = chatFields();
    const result = await streamChat({ provider, model: fields.model, endpoint: $('qwenEndpoint').value.trim(), messages: dedupContext(state.messages.slice(-24)), fast: fields.fast, effort: fields.effort, host: state.host === 'Word' ? 'word' : 'excel' });
    if (turn !== state.turn) return;
    const revision = parseRevision(result.text);
    state.messages.push({ role: 'assistant', content: result.text });
    state.lastAnswer = revision.prose;
    if (state.lastAnswer) answerForWord = state.lastAnswer;
    $('insertAnswer').disabled = !revision.prose; $('lastAnswerState').textContent = result.model ? `via ${result.model}` : '';
    result.bubble.children[result.bubble.children.length - 1].textContent = revision.prose;
    if (revision.actions) await renderPlanCard(revision.actions, revision.skipped);
    else if (revision.edits) await renderEditsCard(revision.edits, revision.skipped);
    if (financeTask) renderFinanceEvidenceCard(evidenceSources, [], revision.prose);
    refreshContextCard();
  } catch (error) {
    const bodyEl = error.bubble?.parent ? error.bubble.children[error.bubble.children.length - 1] : null;
    if (error.name === 'AbortError') {
      if (bodyEl) bodyEl.textContent = error.partial ? `${error.partial}\n\n(stopped)` : '(stopped)';
      if (error.bubble?.parent) toast('Stopped.');
      return;
    }
    if (turn !== state.turn) return;
    if (!state.lastAnswer) $('insertAnswer').disabled = true;
    if (bodyEl) {
      bodyEl.textContent = error.partial ? `${error.partial}\n\n(${explainError(error.message)})` : `Error: ${explainError(error.message)}`;
      toast(explainError(error.message));
    } else {
      renderMessage('assistant', `Error: ${explainError(error.message)}`);
      toast(explainError(error.message));
    }
  } finally {
    if (turn === state.turn) {
      if (answerForWord && state.host === 'Word') {
        try { await appendWordAnswer(answerForWord); }
        catch (error) { toast(`Could not write the answer into Word: ${error.message}`); }
      }
      resetComposer();
      persistSession('Completed · 已完成');
      refreshHistory();
    }
  }
}

function clearConversation() {
  const conversation = $('conversation');
  while (conversation.children.length) conversation.children[0].remove();
  const welcome = document.createElement('div'); welcome.className = 'welcome';
  const icon = document.createElement('div'); icon.className = 'welcome-icon'; icon.textContent = '✦';
  const title = document.createElement('h1'); title.textContent = 'What should I revise?';
  const copy = document.createElement('p'); copy.textContent = 'Ask about the selection, or tell me to change the sheet or the document.';
  welcome.append(icon, title, copy);
  conversation.append(welcome);
}

function newChat() {
  if (activeAgent) return;
  stopGeneration();
  agentCheckpoint = null;
  showAgentStatus(0, '');
  state.messages = [];
  state.attachments = [];
  state.attachmentErrors = [];
  $('attachmentPicker').value = '';
  renderAttachments();
  refreshContextCard();
  state.lastAnswer = '';
  state.forceSheet = false;
  state.audit = [];
  state.sessionId = null;
  state.sessionStarted = null;
  $('insertAnswer').disabled = true;
  $('lastAnswerState').textContent = '';
  clearConversation();
}

function showExecutionPath() {
  const el = $('executionPath');
  if (!el) return;
  el.textContent = hostedRelay()
    ? 'Cloud · questions and workbook text go through Cloudflare to the selected provider'
    : 'Local · the relay stays on this PC';
}

function splitColumnList(value) {
  return String(value || '').split(',').map(part => part.trim()).filter(Boolean);
}

function readFinanceProfile() {
  const tolerance = $('checkTolerance')?.value;
  return AgentCore.normalizeFinanceProfile({
    debitColumn: $('debitColumn')?.value,
    creditColumn: $('creditColumn')?.value,
    amountColumn: $('amountColumn')?.value,
    keyColumns: splitColumnList($('keyColumns')?.value),
    requiredColumns: splitColumnList($('requiredColumns')?.value),
    entityColumn: $('entityColumn')?.value,
    currencyColumn: $('currencyColumn')?.value,
    dateColumn: $('dateColumn')?.value,
    tolerance: tolerance == null || tolerance === '' ? null : Number(tolerance)
  });
}

function loadFinanceProfile() {
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem('eqgc.finance.v1') || '{}'); } catch { saved = {}; }
  const profile = AgentCore.normalizeFinanceProfile(saved);
  const fields = {
    debitColumn: profile.debitColumn,
    creditColumn: profile.creditColumn,
    amountColumn: profile.amountColumn,
    keyColumns: profile.keyColumns.join(', '),
    requiredColumns: profile.requiredColumns.join(', '),
    entityColumn: profile.entityColumn,
    currencyColumn: profile.currencyColumn,
    dateColumn: profile.dateColumn,
    checkTolerance: profile.tolerance
  };
  for (const [id, value] of Object.entries(fields)) {
    const el = $(id);
    if (el) el.value = value == null ? '' : String(value);
  }
}

function saveFinanceProfile() {
  try { localStorage.setItem('eqgc.finance.v1', JSON.stringify(readFinanceProfile())); } catch { /* storage can be unavailable */ }
}

function applyHostedChrome() {
  showExecutionPath();
  if (!hostedRelay()) return;
  $('cloudSources')?.classList.add('hidden');
  $('githubRepoSlot')?.classList.add('hidden');
  $('relayFields')?.classList.add('hidden');
  $('localKeyFields')?.classList.remove('hidden');
  const hint = $('keyHint');
  if (hint) hint.textContent = 'Hosted pane (BYO): your Qwen key is sent per request to QwenCloud through this Worker. The Worker never saves it. Remember uses Office storage when available, or this browser’s storage. If both are blocked, the key works for this session only. 託管版（自備金鑰）：每次請求會經 Worker 將您的 key 傳給 QwenCloud；Worker 不會儲存。記住功能優先使用 Office 儲存，其次是瀏覽器儲存；兩者均不可用時，只限本次使用。';
  $('saveKey').textContent = 'Save key · 儲存金鑰';
  $('clearKey').textContent = 'Clear key · 清除';
}

async function refreshSettings() {
  if (hostedRelay()) {
    await loadHostedKey();
    const active = Boolean(byoKey());
    const status = $('keyStatus');
    if (status) {
      const place = hostedKeyStore === 'office' ? 'remembered in Office' : hostedKeyStore === 'browser' ? 'remembered in this browser' : 'this session only';
      status.textContent = active ? `Your key · ${place} · sent per request` : 'Add your own Qwen key (BYO)';
      status.style.color = active ? '#168251' : '#777';
    }
    setRelayDot(active ? 'ok' : 'warn');
    return;
  }
  try {
    const data = await api('health');
    const provider = $('provider').value;
    const remembered = Boolean(storedKeys()[provider]);
    const active = Boolean(data.configured[provider]);
    $('keyStatus').textContent = active ? (remembered ? 'Key active · remembered on this PC' : 'Key active for this session') : '';
    $('keyStatus').style.color = active ? '#168251' : '#777';
    setRelayDot(active ? 'ok' : 'warn');
  } catch { $('keyStatus').textContent = 'Relay unavailable'; $('keyStatus').style.color = '#b23b36'; setRelayDot('offline'); }
}

function toggleProviderFields() {
  refreshSettings();
}

async function connectOneDrive() {
  try {
    await api('connectors/onedrive/authorize', {});
    toast('Complete Microsoft sign-in in the browser window.');
    startSourcePolling();
    await refreshCloudSources();
  } catch (error) { toast(explainError(error.message)); }
}

async function disconnectSource(source) {
  try {
    await api('connectors/disconnect', { source });
    $('sourceResults').replaceChildren();
    await refreshCloudSources();
    toast('OneDrive disconnected from this relay.');
  } catch (error) { toast(explainError(error.message)); }
}

async function insertAnswer() {
  if (!state.lastAnswer || !state.officeReady) return;
  const answer = state.lastAnswer;
  if (answer.length > CELL_CHAR_LIMIT) { toast('That answer is longer than Excel’s 32,767 character cell limit.'); return; }
  const formula = answer.trim();
  const asFormula = formula.startsWith('=') && !formula.includes('\n') && formula.length <= 8192;
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange().getCell(0, 0);
      if (asFormula) range.formulas = [[formula]];
      else range.values = [[answer]];
      await context.sync();
    });
    toast(asFormula ? 'Formula inserted into the selected cell.' : 'Answer inserted into the selected cell.');
  } catch (error) { toast(`Could not insert answer: ${error.message}`); }
}

$('settingsToggle').addEventListener('click', () => { $('settings').classList.toggle('hidden'); refreshSettings(); });
$('closeSettings').addEventListener('click', () => $('settings').classList.add('hidden'));
$('provider').addEventListener('change', toggleProviderFields);
$('refreshSelection').addEventListener('click', refreshContextCard);
$('includeContext').addEventListener('change', () => { $('shareQuick').checked = $('includeContext').checked; refreshContextCard(); });
$('shareQuick').addEventListener('change', () => { $('includeContext').checked = $('shareQuick').checked; refreshContextCard(); });
function setWorkbookCard(open) {
  $('workbookCard').classList.toggle('hidden', !open);
  $('workbookToggle').setAttribute('aria-expanded', String(open));
}
$('workbookToggle').addEventListener('click', () => { const open = $('workbookCard').classList.contains('hidden'); setWorkbookCard(open); if (open) refreshContextCard(); });
$('sidebarToggle').addEventListener('click', () => setSidebar($('sidebar').classList.contains('hidden')));
$('sidebarClose').addEventListener('click', () => setSidebar(false));
$('newChatSide').addEventListener('click', () => { newChat(); setSidebar(false); });
$('addUsedRange').addEventListener('click', () => { state.forceSheet = true; refreshContextCard(); toast('Next question will use the whole active sheet.'); });
$('saveKey').addEventListener('click', async () => {
  if (hostedRelay()) {
    const key = $('apiKey').value.trim();
    if (!key.startsWith('sk-') || key.length < 8) return toast(explainError('Add your Qwen key in Settings. It must be your own sk-… key.'));
    // Ask during the click gesture; awaiting Office storage first can consume it.
    if ($('rememberKey').checked && !storageOk && typeof document.requestStorageAccess === 'function') {
      try { await document.requestStorageAccess(); storageOk = true; } catch { /* Office storage may still work. */ }
    }
    await loadHostedKey();
    sessionByoKey = key;
    const persisted = $('rememberKey').checked ? await persistHostedKey(key) : false;
    const cleared = $('rememberKey').checked ? true : await clearHostedKey();
    if (!$('rememberKey').checked) sessionByoKey = key;
    $('apiKey').value = '';
    await refreshSettings();
    if ($('rememberKey').checked && !persisted) toast(explainError('Office and browser storage are unavailable. Key works for this session only. Office 及瀏覽器儲存均不可用；金鑰只限本次使用。'));
    else if (!cleared) toast('Key works now, but an older saved copy could not be cleared. Try Clear key again.');
    else toast($('rememberKey').checked ? `Key saved in ${hostedKeyStore === 'office' ? 'Office storage' : 'this browser'}.` : 'Key kept for this session only.');
    return;
  }
  const apiKey = $('apiKey').value.trim(); if (!apiKey) return toast('Enter an API key first.');
  const provider = $('provider').value;
  try {
    await api('configure', { provider, apiKey });
    rememberKey(provider, $('rememberKey').checked ? apiKey : '');
    keysReady = null;
    $('apiKey').value = ''; await refreshSettings();
    toast($('rememberKey').checked ? 'Key loaded and remembered on this PC.' : 'Key loaded into local relay memory.');
  }
  catch (error) { toast(explainError(error.message)); }
});
$('clearKey').addEventListener('click', async () => {
  if (hostedRelay()) {
    await loadHostedKey();
    const cleared = await clearHostedKey();
    await refreshSettings();
    toast(cleared ? 'Your key cleared from this pane.' : 'Key cleared from this session; saved storage may still contain it. Try Clear again.');
    return;
  }
  const provider = $('provider').value;
  try { await api('clear-key', { provider }); rememberKey(provider, ''); keysReady = null; await refreshSettings(); toast('Key cleared.'); }
  catch (error) { toast(explainError(error.message)); }
});
$('connectOneDrive').addEventListener('click', connectOneDrive);
$('disconnectOneDrive').addEventListener('click', () => disconnectSource('onedrive'));
$('searchSources').addEventListener('click', searchCloudSource);
$('saveSkill').addEventListener('click', saveCurrentSkill);
$('addMemory').addEventListener('click', addMemoryNote);
$('sourceQuery').addEventListener('keydown', event => { if (event.key === 'Enter') { event.preventDefault(); searchCloudSource(); } });
$('newChat').addEventListener('click', newChat);
$('workMode')?.addEventListener('change', () => { state.forceSheet = false; syncAgentControls(); refreshContextCard(); });
$('continueAgent')?.addEventListener('click', async () => {
  if (state.busy || operatingMode() !== 'agent' || !agentCheckpoint) return;
  await sendAgentPrompt(agentCheckpoint.prompt, agentCheckpoint);
});
// Errors must be readable and problem-solving: keep the raw reason, append a bilingual cause + next step.
const ERROR_GUIDE = [
  [/relay is running build (\S+) but this pane is (\S+)/, m => `版本不符：relay ${m[1]}、pane ${m[2]}。解法：npm run relay-stop → npm run relay-ensure，再重載 pane。 Build mismatch: restart the relay, then reload the pane.`],
  [/no extractable text layer/, () => '掃描 PDF（無文字層）。解法：Settings 輸入您的 Qwen key 启用視覺轉錄，或改附文字版 PDF 匯出。 Scanned PDF: add your Qwen key in Settings for vision transcription, or attach a text-based export.'],
  [/needs your Qwen key|Add your Qwen API key/, () => '未載入 Qwen key。解法：Settings → 貼上您的 key → Use key。 No Qwen key loaded: Settings → paste your key → Use key.'],
  [/relay unavailable|Failed to fetch|ECONNREFUSED|fetch failed/, () => '本機 relay 未啟動或斷線。解法：npm start（或 npm run relay-ensure）。 Local relay down: run npm start or npm run relay-ensure.'],
  [/401|Incorrect API key|invalid.*api key/i, () => 'key 被拒（401）。解法：確認 key 正確且未輪換；Settings 重新 Use key。 Key rejected (401): verify the key is current and re-apply it in Settings.'],
  [/429|quota|rate limit/i, () => '配額或速率限制（429）。解法：稍後再試；Token Plan 餘額只可在 console 查看。 Quota/rate limit (429): retry later; Token Plan balance is console-only.'],
  [/Model not exist|model_not_found/i, () => '模型名稱不存在。解法：Settings 改用 qwen4-max。 Model not found: switch to qwen4-max in Settings.'],
  [/over the (\d+) MB limit/, m => `檔案超過 ${m[1]} MB。解法：分頁／壓縮，或改匯出文字版。 File over ${m[1]} MB: split/compress it or export a text version.`],
  [/no readable text was found/, () => '檔內無可讀文字（純圖片或空檔）。解法：掃描 PDF 會自动試視覺轉錄（需 key）；影像附件需 key。 No readable text (image-only or empty): scanned PDFs auto-try vision (key required).'],
  [/pdftotext is not available|pdftoppm is not available/, () => '缺 poppler 工具。解法：安裝 poppler（winget install oschwartz10612.Poppler），或改附文字版。 poppler missing: install it via winget or attach a text export.'],
  [/Session not found/, () => 'session 已刪或不存在。解法：重新整理 History 清單。 Session deleted or missing: refresh the History list.'],
  [/Office file sharing is off/, () => 'Share 開關關閉。解法：footer 或側欄勾選 Share。 Share gate is off: tick Share in the footer or sidebar.']
];

function explainError(message) {
  const text = String(message || 'Unknown error');
  for (const [pattern, remedy] of ERROR_GUIDE) {
    const match = text.match(pattern);
    if (match) return `${text}\n→ ${remedy(match)}`;
  }
  return text;
}

function setAttachMenu(open) {
  $('attachMenu').classList.toggle('hidden', !open);
  $('attachFiles').setAttribute('aria-expanded', String(open));
}
$('attachFiles').addEventListener('click', () => { setGitHubRepoMenu(false); setAttachMenu($('attachMenu').classList.contains('hidden')); });
$('menuUpload').addEventListener('click', () => { setAttachMenu(false); $('attachmentPicker').click(); });
$('menuCloud').addEventListener('click', () => { setAttachMenu(false); setSidebar(true); $('cloudSources').open = true; $('sourceQuery').focus(); });
$('menuSkill').addEventListener('click', () => { setAttachMenu(false); setSidebar(true); $('skillsCard').open = true; $('skillId').focus(); });
$('menuSource').addEventListener('click', () => { setAttachMenu(false); setSidebar(true); });
document.addEventListener('click', event => {
  if (!event.target.closest?.('.attach-menu-wrap')) setAttachMenu(false);
  if (!event.target.closest?.('.repo-slot')) setGitHubRepoMenu(false);
});
$('attachmentPicker').addEventListener('change', async event => {
  await addAttachments(event.target.files || []);
  event.target.value = '';
});
$('composer').addEventListener('submit', async event => {
  event.preventDefault();
  if (state.busy) { stopGeneration(); return; }
  const value = $('prompt').value;
  if (!value.trim()) return;
  $('prompt').value = '';
  if (await runGitWorkspace(value)) return;
  await sendPrompt(value);
});
$('prompt').addEventListener('keydown', event => { if (event.key === 'Enter' && !event.shiftKey) { event.preventDefault(); $('composer').requestSubmit(); } });
$('insertAnswer').addEventListener('click', insertAnswer);

loadFinanceProfile();
['debitColumn', 'creditColumn', 'amountColumn', 'keyColumns', 'requiredColumns', 'entityColumn', 'currencyColumn', 'dateColumn', 'checkTolerance'].forEach(id => {
  $(id).addEventListener('change', saveFinanceProfile);
});

$('installGithubSkill').addEventListener('click', installGitHubSkill);
$('saveGitHubToken').addEventListener('click', saveGitHubMcpToken);
$('clearGitHubToken').addEventListener('click', clearGitHubMcpToken);
$('connectGitHub').addEventListener('click', connectGitHubAccount);
$('disconnectGitHub').addEventListener('click', disconnectGitHubAccount);
$('githubRepoPick').addEventListener('click', () => setGitHubRepoMenu($('githubRepoMenu').classList.contains('hidden')));
$('githubRepoNone').addEventListener('click', () => setGitHubRepo(''));
$('githubRepoConnect').addEventListener('click', () => { setGitHubRepoMenu(false); setSidebar(true); $('cloudSources').open = true; });
refreshGitHubRepoPick();
setGitHubRepoMenu(false);
$('saveGitRepo').addEventListener('click', saveGitRepoFolder);
restoreGitHubToken();
refreshGitHubAccount();
restoreGitWorkspace();
$('modelPick').addEventListener('change', () => syncModel('modelPick'));
$('qwenModel').addEventListener('change', () => syncModel('qwenModel'));
$('effortMode').addEventListener('change', saveChatPrefs);
restoreChatPrefs();

Office.onReady(info => {
  if (info.host === Office.HostType.Excel || info.host === Office.HostType.Word) {
    state.host = info.host === Office.HostType.Word ? 'Word' : 'Excel';
    state.officeReady = true;
    const eyebrow = $('contextEyebrow');
    if (eyebrow) eyebrow.textContent = state.host === 'Word' ? 'DOCUMENT CONTEXT · AUTO' : 'WORKBOOK CONTEXT · AUTO';
    const brand = $('brandHost');
    if (brand) brand.textContent = state.host === 'Word' ? 'Revising this document' : 'Revising this workbook';
    refreshContextCard();
    watchSelection();
  }
  applyHostedChrome();
  refreshSettings();
  refreshCloudSources();
  refreshSkills();
  refreshHistory();
  refreshMemory();
  ensureKeys().then(refreshSettings);
});
applyHostedChrome();
refreshSkills();
refreshHistory();
refreshMemory();
try { if (localStorage.getItem('eqgc.sidebar.v1') === '1') setSidebar(true); } catch { /* storage optional */ }
