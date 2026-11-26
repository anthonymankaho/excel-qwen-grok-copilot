import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { createRelay } from '../server.mjs';
import { githubAccountStatus, setGithubAccountHooksForTests, startGithubWebLogin } from '../github-account.mjs';

function request(port, method, path, body) {
  const payload = body == null ? null : JSON.stringify(body);
  return new Promise((resolve, reject) => {
    const req = http.request({
      hostname: '127.0.0.1', port, path, method,
      headers: payload ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) } : {}
    }, res => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    if (payload) req.end(payload);
    else req.end();
  });
}

test('GitHub account status never returns the token, and connect does not open a login when already connected', async () => {
  const secret = 'gho_testtokenvalue';
  let spawned = 0;
  let mode = 'connected';
  setGithubAccountHooksForTests({
    exec: async args => {
      if (mode !== 'connected') return { code: 1, stdout: '', stderr: 'not logged in' };
      if (args[0] === 'auth' && args[1] === 'token') return { code: 0, stdout: `${secret}\n`, stderr: '' };
      return { code: 0, stdout: JSON.stringify({ login: 'anthonymankaho', email: null, token: secret }), stderr: '' };
    },
    spawn: () => { spawned += 1; }
  });
  const previous = global.fetch;
  global.fetch = async (url, options) => {
    assert.equal(options.headers.Authorization, `Bearer ${secret}`);
    assert.equal(String(url).includes('/user/repos'), true);
    return { status: 200, headers: { get: () => null }, text: async () => JSON.stringify([{ full_name: 'owner/repo', private: false, default_branch: 'main', description: '' }]) };
  };
  const { server } = createRelay();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  try {
    const status = await githubAccountStatus();
    assert.equal(status.connected, true);
    assert.equal(status.login, 'anthonymankaho');
    assert.equal(status.email, '');
    assert.equal(Object.hasOwn(status, 'token'), false);
    assert.equal(JSON.stringify(status).includes(secret), false);

    const account = await request(port, 'GET', '/api/connectors/github/account');
    assert.equal(account.status, 200);
    assert.equal(account.body.includes(secret), false);
    assert.match(account.body, /anthonymankaho/);
    const connect = await request(port, 'POST', '/api/connectors/github/connect', {});
    assert.equal(spawned, 0);
    assert.equal(connect.body.includes('pending'), false);
    assert.equal(connect.body.includes(secret), false);

    const repos = await request(port, 'POST', '/api/connectors/github/repos', { useAccount: true, query: 'repo' });
    assert.equal(repos.status, 200);
    assert.equal(repos.body.includes(secret), false);
    assert.match(repos.body, /owner\/repo/);

    mode = 'signed-out';
    const pending = await request(port, 'POST', '/api/connectors/github/connect', {});
    assert.equal(spawned, 1);
    assert.equal(JSON.parse(pending.body).pending, true);
    startGithubWebLogin();
    assert.equal(spawned, 2);
  } finally {
    setGithubAccountHooksForTests(null);
    global.fetch = previous;
    await new Promise(resolve => server.close(resolve));
  }
});
