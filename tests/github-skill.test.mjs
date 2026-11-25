import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import http from 'node:http';
import { createRelay } from '../server.mjs';
import { fetchGitHubSkill, installSkillFile, parseGitHubSkillSource } from '../relay-core.mjs';

const skill = `---
name: demo-skill
description: Use for a demo chart.
triggers: chart, 圖表
---

# Demo

Keep this body.
`;

test('a GitHub skill source accepts owner/repo and one SKILL.md link', () => {
  assert.deepEqual(parseGitHubSkillSource('kgraph57/mckinsey-style-visualization-skill'), {
    owner: 'kgraph57', repo: 'mckinsey-style-visualization-skill', ref: '', path: 'SKILL.md'
  });
  assert.deepEqual(parseGitHubSkillSource('https://github.com/kgraph57/demo-skill/blob/main/skills/demo-skill/SKILL.md'), {
    owner: 'kgraph57', repo: 'demo-skill', ref: 'main', path: 'skills/demo-skill/SKILL.md'
  });
  for (const bad of [
    'https://user:ghp_secret@github.com/owner/repo',
    'https://gist.github.com/owner/repo',
    'https://github.com/owner/repo/archive/main.zip',
    'https://github.com/owner/repo/blob/main/../../SKILL.md',
    'https://github.com/owner/repo/blob/main/README.md',
    'http://github.com/owner/repo'
  ]) {
    assert.equal(typeof parseGitHubSkillSource(bad).error, 'string', bad);
  }
});

test('an installed GitHub skill keeps its original file and replaces only when asked', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-gh-skill-'));
  try {
    const saved = installSkillFile(root, skill, {});
    assert.equal(saved.id, 'demo-skill');
    const file = path.join(root, 'demo-skill', 'SKILL.md');
    assert.match(fs.readFileSync(file, 'utf8'), /triggers: chart, 圖表/);
    const again = installSkillFile(root, skill, {});
    assert.match(again.error, /already exists/i);
    fs.writeFileSync(path.join(root, 'demo-skill', 'notes.txt'), 'keep me');
    const replaced = installSkillFile(root, skill.replace('Keep this body.', 'Updated body.'), { replace: true });
    assert.equal(replaced.replaced, true);
    assert.match(fs.readFileSync(file, 'utf8'), /Updated body/);
    assert.equal(fs.readFileSync(path.join(root, 'demo-skill', 'notes.txt'), 'utf8'), 'keep me');
    assert.match(installSkillFile(root, '---\nname: Bad_Name\ndescription: x\n---\nbody').error, /name/i);
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('GitHub skill fetch stays on api.github.com and does not follow a redirect', async () => {
  const token = 'ghp_testtoken';
  let calls = 0;
  const fetched = await fetchGitHubSkill({
    owner: 'kgraph57', repo: 'demo-skill', ref: 'main', path: 'SKILL.md', token,
    fetchImpl: async (url, options) => {
      calls += 1;
      assert.equal(url, 'https://api.github.com/repos/kgraph57/demo-skill/contents/SKILL.md?ref=main');
      assert.equal(options.redirect, 'manual');
      assert.equal(options.headers.Authorization, `Bearer ${token}`);
      return { status: 200, headers: { get: () => null }, text: async () => skill };
    }
  });
  assert.equal(calls, 1);
  assert.match(fetched.text, /demo-skill/);
  const redirected = await fetchGitHubSkill({
    owner: 'kgraph57', repo: 'demo-skill', ref: '', path: 'SKILL.md', token,
    fetchImpl: async () => ({ status: 302, headers: { get: () => 'https://evil.example/skill' }, text: async () => 'no' })
  });
  assert.match(redirected.error, /redirect/i);
  assert.equal(redirected.error.includes(token), false);
});

test('the local relay installs a GitHub skill and leaves the token out of the reply', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-gh-http-'));
  process.env.SKILLS_ROOT = root;
  const previous = global.fetch;
  const token = 'github_pat_testtoken';
  global.fetch = async () => ({ status: 200, headers: { get: () => String(skill.length) }, text: async () => skill });
  const { server } = createRelay();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  try {
    const response = await request(port, { repo: 'kgraph57/demo-skill', token, replace: false });
    assert.equal(response.status, 200);
    const body = JSON.parse(response.body);
    assert.equal(body.id, 'demo-skill');
    assert.equal(response.body.includes(token), false);
    assert.match(fs.readFileSync(path.join(root, 'demo-skill', 'SKILL.md'), 'utf8'), /triggers: chart/);
    const duplicate = await request(port, { repo: 'kgraph57/demo-skill', token, replace: false });
    assert.equal(duplicate.status, 400);
    assert.equal(duplicate.body.includes(token), false);
  } finally {
    global.fetch = previous;
    await new Promise(resolve => server.close(resolve));
    delete process.env.SKILLS_ROOT;
    fs.rmSync(root, { recursive: true, force: true });
  }
});

function request(port, body) {
  const payload = JSON.stringify(body);
  return new Promise((resolve, reject) => {
    const req = http.request({ hostname: '127.0.0.1', port, path: '/api/skills/github', method: 'POST', headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) } }, res => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.end(payload);
  });
}
