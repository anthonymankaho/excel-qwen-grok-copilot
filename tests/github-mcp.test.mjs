import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import { fileURLToPath } from 'node:url';
import { createRelay } from '../server.mjs';
import { listGitHubRepos, readGitHubFile, writeGitHubFiles } from '../github-mcp.mjs';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const PARENT = 'a'.repeat(40);
const TREE = 'b'.repeat(40);
const BLOB = 'c'.repeat(40);
const COMMIT = 'd'.repeat(40);

function mockFetch() {
  const calls = [];
  const fetchImpl = async (url, options) => {
    const body = options.body ? JSON.parse(options.body) : null;
    calls.push({ url, method: options.method, redirect: options.redirect, body, authorization: options.headers.Authorization });
    assert.equal(new URL(url).hostname, 'api.github.com');
    assert.equal(options.redirect, 'manual');
    const json = (value, status = 200) => ({ status, headers: { get: () => null }, text: async () => JSON.stringify(value) });
    if (options.headers.Accept === 'application/vnd.github.raw') {
      return { status: 200, headers: { get: () => null }, text: async () => 'hello file' };
    }
    if (options.method === 'POST' && url.endsWith('/git/blobs')) return json({ sha: BLOB });
    if (options.method === 'POST' && url.endsWith('/git/trees')) return json({ sha: TREE });
    if (options.method === 'POST' && url.endsWith('/git/commits')) return json({ sha: COMMIT });
    if (options.method === 'POST' && url.endsWith('/git/refs')) return json({ ref: body.ref });
    if (options.method === 'POST' && url.endsWith('/pulls')) return json({ html_url: 'https://github.com/owner/repo/pull/4' });
    if (options.method === 'PATCH') return json({ ref: 'refs/heads/main' });
    if (url.includes('/git/ref/heads/')) return json({ object: { sha: PARENT } });
    if (url.includes('/git/commits/')) return json({ tree: { sha: TREE } });
    if (options.method === 'GET') return json({ default_branch: 'main', full_name: 'owner/repo' });
    throw new Error(`unexpected ${options.method} ${url}`);
  };
  return { calls, fetchImpl };
}

test('GitHub reads stay on api.github.com and do not follow a redirect', async () => {
  const token = 'ghp_testtoken';
  const listed = await listGitHubRepos({
    token, query: 'repo',
    fetchImpl: async (url, options) => {
      assert.equal(url.includes('/user/repos'), true);
      assert.equal(options.headers.Authorization, `Bearer ${token}`);
      return { status: 200, headers: { get: () => null }, text: async () => JSON.stringify([{ full_name: 'owner/repo', private: true, default_branch: 'main', description: 'notes' }, { full_name: 'other/skip', private: false, default_branch: 'main' }]) };
    }
  });
  assert.equal(listed.items.length, 1);
  assert.equal(listed.items[0].fullName, 'owner/repo');
  const read = await readGitHubFile({ token, repo: 'owner/repo', path: 'docs/README.md', fetchImpl: async (url, options) => {
    assert.equal(url, 'https://api.github.com/repos/owner/repo/contents/docs/README.md');
    assert.equal(options.redirect, 'manual');
    return { status: 200, headers: { get: () => null }, text: async () => 'hello file' };
  } });
  assert.equal(read.text, 'hello file');
  const redirected = await readGitHubFile({
    repo: 'owner/repo', path: 'README.md',
    fetchImpl: async () => ({ status: 302, headers: { get: () => 'https://evil.example' }, text: async () => 'no' })
  });
  assert.match(redirected.error, /redirect/i);
  assert.equal(redirected.error.includes(token), false);
});

test('a GitHub write opens a pull request and does not force the default branch', async () => {
  const token = 'ghp_testtoken';
  const { calls, fetchImpl } = mockFetch();
  const result = await writeGitHubFiles({
    token, repo: 'owner/repo', title: 'Update notes', body: 'Why',
    files: [{ path: 'README.md', content: 'Hello' }], target: 'pull-request',
    fetchImpl, now: new Date('2026-11-26T01:02:03Z')
  });
  assert.equal(result.prUrl, 'https://github.com/owner/repo/pull/4');
  assert.equal(result.branch, 'eqgc/20261126-010203');
  assert.equal(result.files[0], 'README.md');
  assert.equal(JSON.stringify(result).includes(token), false);
  const pulls = calls.find(call => call.url.endsWith('/pulls'));
  assert.equal(pulls.body.base, 'main');
  assert.equal(pulls.body.head, 'eqgc/20261126-010203');
  assert.equal(calls.some(call => call.body && call.body.force === true), false);
  const direct = mockFetch();
  const updated = await writeGitHubFiles({
    token, repo: 'owner/repo', title: 'Update notes',
    files: [{ path: 'README.md', content: 'Hello' }], target: 'default',
    fetchImpl: direct.fetchImpl, now: new Date('2026-11-26T01:02:03Z')
  });
  assert.equal(updated.base, 'main');
  assert.equal(updated.prUrl, undefined);
  const patch = direct.calls.find(call => call.method === 'PATCH');
  assert.deepEqual(patch.body, { sha: COMMIT, force: false });
});

test('a GitHub draft that looks like a secret is rejected before any request', async () => {
  let calls = 0;
  const result = await writeGitHubFiles({
    token: 'ghp_testtoken', repo: 'owner/repo', title: 'keys',
    files: [{ path: 'notes.txt', content: 'token ghp_secretvalue' }],
    fetchImpl: async () => { calls += 1; return { status: 200, headers: { get: () => null }, text: async () => '{}' }; }
  });
  assert.match(result.error, /secret/i);
  assert.equal(calls, 0);
  const env = await readGitHubFile({ repo: 'owner/repo', path: '.env', fetchImpl: async () => { calls += 1; } });
  assert.match(env.error, /secret/i);
  assert.equal(calls, 0);
});

test('the local relay writes to GitHub only through the click route and omits the token', async () => {
  const previous = global.fetch;
  const token = 'github_pat_testtoken';
  const { fetchImpl } = mockFetch();
  global.fetch = fetchImpl;
  const { server } = createRelay();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  try {
    const response = await post(port, '/api/connectors/github/write', {
      token, repo: 'owner/repo', title: 'Update notes', body: 'Why',
      files: [{ path: 'README.md', content: 'Hello' }], target: 'pull-request'
    });
    assert.equal(response.status, 200);
    assert.equal(response.body.includes(token), false);
    assert.match(response.body, /pull\/4/);
    const secret = await post(port, '/api/connectors/github/write', {
      token, repo: 'owner/repo', title: 'keys', files: [{ path: 'notes.txt', content: 'ghp_secretvalue' }], target: 'pull-request'
    });
    assert.equal(secret.status, 400);
    assert.equal(secret.body.includes(token), false);
  } finally {
    global.fetch = previous;
    await new Promise(resolve => server.close(resolve));
  }
});

test('the hosted worker does not import the GitHub connector', () => {
  const worker = fs.readFileSync(path.join(root, 'worker.mjs'), 'utf8');
  assert.equal(worker.includes('github-mcp'), false);
  assert.equal(worker.includes('github-account'), false);
  assert.equal(worker.includes('git-workspace'), false);
});

function post(port, path, body) {
  const payload = JSON.stringify(body);
  return new Promise((resolve, reject) => {
    const req = http.request({ hostname: '127.0.0.1', port, path, method: 'POST', headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) } }, res => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.end(payload);
  });
}
