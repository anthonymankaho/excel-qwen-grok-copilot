// Local GitHub connector. The token is supplied per request and is never logged or returned.
// Reads list repositories or one text file. Writes create one commit: a new branch plus a
// pull request, or, only when asked, a non-force update of the default branch.

const NAME_RE = /^[A-Za-z0-9_.-]{1,100}$/;
const SECRET_RE = /(?:sk-[A-Za-z0-9]{8,}|ghp_[A-Za-z0-9]{8,}|github_pat_[A-Za-z0-9_]{8,}|AKIA[0-9A-Z]{8,}|-----BEGIN [A-Z ]*PRIVATE KEY-----|xox[baprs]-[A-Za-z0-9-]{8,})/;
const SECRET_PATHS = new Set(['.env', '.env.local', '.env.production', 'id_rsa', 'id_dsa', 'id_ecdsa', 'id_ed25519']);
const MAX_FILE_BYTES = 64 * 1024;
const MAX_FILES = 5;
const MAX_READ_CHARS = 70_000;

export function githubTokenError(token) {
  if (token == null || token === '') return '';
  if (typeof token !== 'string' || token.length > 300 || /\s/.test(token)) return 'Invalid GitHub token.';
  return '';
}

export function parseGitHubRepo(raw) {
  const text = String(raw || '').trim();
  if (!text || text.includes('..') || text.includes('@') || text.includes('\\')) return { error: 'Use owner/repo.' };
  const match = text.match(/^([A-Za-z0-9_.-]+)\/([A-Za-z0-9_.-]+)$/);
  if (!match || !NAME_RE.test(match[1]) || !NAME_RE.test(match[2])) return { error: 'Use owner/repo.' };
  return { owner: match[1], repo: match[2] };
}

export function parseGitHubPath(raw) {
  const text = String(raw || '').trim().replace(/^\/+/, '');
  if (!text || text.length > 400 || text.includes('\\') || text.includes('..') || /[\u0000-\u001f]/.test(text)) {
    return { error: 'Invalid GitHub file path.' };
  }
  const parts = text.split('/');
  if (parts.some(part => !part || part === '.' || part === '..')) return { error: 'Invalid GitHub file path.' };
  const base = parts.at(-1).toLowerCase();
  if (SECRET_PATHS.has(base) || base.endsWith('.pem') || base.endsWith('.key')) {
    return { error: 'This path looks like a secret file. It was not sent to GitHub.' };
  }
  return { path: text };
}

export function githubSecretError(text) {
  if (SECRET_RE.test(String(text || ''))) return 'This draft looks like a secret. It was not sent to GitHub.';
  return '';
}

function cleanQuery(query) {
  return String(query || '').trim().slice(0, 200);
}

export function normalizeGitHubWrite(raw) {
  if (!raw || typeof raw !== 'object') return { error: 'The GitHub draft is incomplete.' };
  const repo = parseGitHubRepo(raw.repo);
  if (repo.error) return repo;
  const title = String(raw.title || 'Update from Qwen for Excel and Word').trim().slice(0, 120);
  const body = String(raw.body || '').trim().slice(0, 4000);
  if (!title) return { error: 'The GitHub draft needs a title.' };
  const secret = githubSecretError(`${title}\n${body}`);
  if (secret) return { error: secret };
  const target = raw.target === 'default' ? 'default' : 'pull-request';
  if (raw.target != null && raw.target !== 'default' && raw.target !== 'pull-request') {
    return { error: 'Choose a pull request or the default branch.' };
  }
  if (!Array.isArray(raw.files) || !raw.files.length) return { error: 'The GitHub draft has no files.' };
  if (raw.files.length > MAX_FILES) return { error: `One approval can include at most ${MAX_FILES} files.` };
  const files = [];
  for (const file of raw.files) {
    const parsed = parseGitHubPath(file?.path);
    if (parsed.error) return parsed;
    if (typeof file?.content !== 'string' || !file.content) return { error: `The draft for ${parsed.path} is empty.` };
    if (Buffer.byteLength(file.content) > MAX_FILE_BYTES) return { error: `${parsed.path} is over 64 KB.` };
    const fileSecret = githubSecretError(file.content);
    if (fileSecret) return { error: fileSecret };
    files.push({ path: parsed.path, content: file.content });
  }
  return { owner: repo.owner, repo: repo.repo, fullName: `${repo.owner}/${repo.repo}`, title, body, files, target };
}

async function githubRequest(url, { token = '', method = 'GET', body, accept, fetchImpl } = {}) {
  let parsed;
  try { parsed = new URL(url); } catch { return { error: 'GitHub requests stay on api.github.com.', status: 400 }; }
  if (parsed.protocol !== 'https:' || parsed.hostname !== 'api.github.com' || parsed.username || parsed.password) {
    return { error: 'GitHub requests stay on api.github.com.', status: 400 };
  }
  const headers = {
    Accept: accept || 'application/vnd.github+json',
    'User-Agent': 'excel-qwen-grok-copilot',
    'X-GitHub-Api-Version': '2022-11-28'
  };
  if (token) headers.Authorization = `Bearer ${token}`;
  if (body !== undefined) headers['Content-Type'] = 'application/json';
  const fetchFn = fetchImpl || globalThis.fetch;
  let response;
  try {
    response = await fetchFn(parsed.href, {
      method,
      redirect: 'manual',
      headers,
      body: body === undefined ? undefined : JSON.stringify(body)
    });
  } catch {
    return { error: 'GitHub could not be reached.', status: 502 };
  }
  if (response.status >= 300 && response.status < 400) {
    return { error: 'GitHub redirected this request. The redirect was not followed.', status: 502 };
  }
  const text = typeof response.text === 'function' ? await response.text() : '';
  if (response.status >= 400) {
    const remaining = response.headers?.get?.('x-ratelimit-remaining');
    if (response.status === 429 || remaining === '0') return { error: 'GitHub rate limit. Wait a minute and retry.', status: response.status };
    if (response.status === 401 || response.status === 403) return { error: 'GitHub refused this token or this repository is private.', status: response.status };
    if (response.status === 404) return { error: 'GitHub could not find that repository or file.', status: 404 };
    if (response.status === 422) return { error: 'GitHub rejected this change. The branch may have moved.', status: 422 };
    return { error: `GitHub returned HTTP ${response.status}.`, status: response.status };
  }
  if (accept === 'application/vnd.github.raw') return { status: response.status, text };
  try { return { status: response.status, data: text ? JSON.parse(text) : {} }; }
  catch { return { error: 'GitHub returned an unreadable response.', status: 502 }; }
}

function repoItem(repo) {
  return {
    fullName: String(repo.full_name || ''),
    private: Boolean(repo.private),
    defaultBranch: String(repo.default_branch || ''),
    description: String(repo.description || '').slice(0, 200)
  };
}

export async function listGitHubRepos({ token = '', query = '', fetchImpl } = {}) {
  const tokenError = githubTokenError(token);
  if (tokenError) return { error: tokenError, status: 400 };
  const term = cleanQuery(query);
  if (!token && !term) return { error: 'A GitHub token is required to list your repositories. Public search needs a name.', status: 400 };
  if (!token) {
    const url = new URL('https://api.github.com/search/repositories');
    url.searchParams.set('q', term);
    url.searchParams.set('per_page', '10');
    const result = await githubRequest(url.href, { fetchImpl });
    if (result.error) return result;
    const items = (Array.isArray(result.data?.items) ? result.data.items : []).slice(0, 10).map(repoItem).filter(item => item.fullName);
    return { items };
  }
  const url = new URL('https://api.github.com/user/repos');
  url.searchParams.set('per_page', '30');
  url.searchParams.set('sort', 'updated');
  url.searchParams.set('affiliation', 'owner,collaborator,organization_member');
  const result = await githubRequest(url.href, { token, fetchImpl });
  if (result.error) return result;
  const needle = term.toLowerCase();
  const items = (Array.isArray(result.data) ? result.data : [])
    .map(repoItem)
    .filter(item => item.fullName && (!needle || item.fullName.toLowerCase().includes(needle)))
    .slice(0, 10);
  return { items };
}

export async function readGitHubFile({ token = '', repo, path, fetchImpl } = {}) {
  const tokenError = githubTokenError(token);
  if (tokenError) return { error: tokenError, status: 400 };
  const parsedRepo = parseGitHubRepo(repo);
  if (parsedRepo.error) return { error: parsedRepo.error, status: 400 };
  const parsedPath = parseGitHubPath(path);
  if (parsedPath.error) return { error: parsedPath.error, status: 400 };
  const encoded = parsedPath.path.split('/').map(encodeURIComponent).join('/');
  const url = `https://api.github.com/repos/${encodeURIComponent(parsedRepo.owner)}/${encodeURIComponent(parsedRepo.repo)}/contents/${encoded}`;
  const result = await githubRequest(url, { token, accept: 'application/vnd.github.raw', fetchImpl });
  if (result.error) return result;
  if (result.text.includes('\u0000')) return { error: 'GitHub returned a file that is not text.', status: 400 };
  if (Buffer.byteLength(result.text) > MAX_FILE_BYTES) return { error: 'This file is over 64 KB.', status: 400 };
  let text = result.text;
  let truncated = false;
  if (text.length > MAX_READ_CHARS) {
    const omitted = text.length - MAX_READ_CHARS;
    text = `${text.slice(0, MAX_READ_CHARS)}\n[...truncated: ${omitted} characters omitted]`;
    truncated = true;
  }
  return { repo: `${parsedRepo.owner}/${parsedRepo.repo}`, path: parsedPath.path, text, truncated };
}

function branchName(now) {
  const date = now instanceof Date ? now : new Date();
  const pad = value => String(value).padStart(2, '0');
  return `eqgc/${date.getUTCFullYear()}${pad(date.getUTCMonth() + 1)}${pad(date.getUTCDate())}-${pad(date.getUTCHours())}${pad(date.getUTCMinutes())}${pad(date.getUTCSeconds())}`;
}

export async function writeGitHubFiles({ token = '', repo, files, title, body, target, fetchImpl, now } = {}) {
  const tokenError = githubTokenError(token);
  if (tokenError) return { error: tokenError, status: 400 };
  if (!token) return { error: 'A write needs your own GitHub token.', status: 400 };
  const draft = normalizeGitHubWrite({ repo, files, title, body, target });
  if (draft.error) return { error: draft.error, status: 400 };
  const baseUrl = `https://api.github.com/repos/${encodeURIComponent(draft.owner)}/${encodeURIComponent(draft.repo)}`;
  const repoInfo = await githubRequest(baseUrl, { token, fetchImpl });
  if (repoInfo.error) return repoInfo;
  const base = String(repoInfo.data?.default_branch || '');
  if (!base || !/^[A-Za-z0-9._/-]{1,200}$/.test(base)) return { error: 'GitHub did not name a default branch.', status: 502 };
  const ref = await githubRequest(`${baseUrl}/git/ref/heads/${encodeURIComponent(base).replace(/%2F/g, '/')}`, { token, fetchImpl });
  if (ref.error) return ref;
  const parent = String(ref.data?.object?.sha || '');
  if (!/^[0-9a-f]{40}$/i.test(parent)) return { error: 'GitHub did not return the default branch commit.', status: 502 };
  const commitInfo = await githubRequest(`${baseUrl}/git/commits/${parent}`, { token, fetchImpl });
  if (commitInfo.error) return commitInfo;
  const baseTree = String(commitInfo.data?.tree?.sha || '');
  if (!/^[0-9a-f]{40}$/i.test(baseTree)) return { error: 'GitHub did not return the current tree.', status: 502 };
  const tree = [];
  for (const file of draft.files) {
    const blob = await githubRequest(`${baseUrl}/git/blobs`, {
      token, method: 'POST', fetchImpl, body: { content: file.content, encoding: 'utf-8' }
    });
    if (blob.error) return blob;
    const sha = String(blob.data?.sha || '');
    if (!/^[0-9a-f]{40}$/i.test(sha)) return { error: 'GitHub did not store the file draft.', status: 502 };
    tree.push({ path: file.path, mode: '100644', type: 'blob', sha });
  }
  const nextTree = await githubRequest(`${baseUrl}/git/trees`, {
    token, method: 'POST', fetchImpl, body: { base_tree: baseTree, tree }
  });
  if (nextTree.error) return nextTree;
  const treeSha = String(nextTree.data?.sha || '');
  const message = draft.body ? `${draft.title}\n\n${draft.body}` : draft.title;
  const commit = await githubRequest(`${baseUrl}/git/commits`, {
    token, method: 'POST', fetchImpl, body: { message, tree: treeSha, parents: [parent] }
  });
  if (commit.error) return commit;
  const commitSha = String(commit.data?.sha || '');
  if (!/^[0-9a-f]{40}$/i.test(commitSha)) return { error: 'GitHub did not create the commit.', status: 502 };
  const paths = draft.files.map(file => file.path);
  if (draft.target === 'default') {
    const updated = await githubRequest(`${baseUrl}/git/refs/heads/${encodeURIComponent(base).replace(/%2F/g, '/')}`, {
      token, method: 'PATCH', fetchImpl, body: { sha: commitSha, force: false }
    });
    if (updated.error) return { error: 'GitHub rejected the default-branch update. The default branch was not changed.', status: updated.status || 422 };
    return { ok: true, target: 'default', repo: draft.fullName, base, branch: base, commit: commitSha, files: paths };
  }
  const branch = branchName(now);
  const created = await githubRequest(`${baseUrl}/git/refs`, {
    token, method: 'POST', fetchImpl, body: { ref: `refs/heads/${branch}`, sha: commitSha }
  });
  if (created.error) return created;
  const pull = await githubRequest(`${baseUrl}/pulls`, {
    token, method: 'POST', fetchImpl, body: { title: draft.title, head: branch, base, body: draft.body }
  });
  if (pull.error) {
    return { error: `The branch ${branch} was created, and the pull request was not opened.`, status: pull.status || 502, branch };
  }
  return {
    ok: true,
    target: 'pull-request',
    repo: draft.fullName,
    base,
    branch,
    commit: commitSha,
    prUrl: String(pull.data?.html_url || ''),
    files: paths
  };
}
