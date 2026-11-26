import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { commitGitFile, openGitFile, pullGitFile, pushGitBranch } from '../git-workspace.mjs';

function git(cwd, args) {
  const result = spawnSync('git', args, { cwd, encoding: 'utf8' });
  if (result.status !== 0) throw new Error(result.stderr || result.stdout || `git ${args.join(' ')} failed`);
}

function initRepo() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-git-'));
  git(dir, ['init']);
  git(dir, ['config', 'user.email', 'test@example.com']);
  git(dir, ['config', 'user.name', 'Test']);
  fs.writeFileSync(path.join(dir, 'notes.txt'), 'hello\n');
  fs.writeFileSync(path.join(dir, 'keep.txt'), 'leave me\n');
  git(dir, ['add', '--', 'notes.txt', 'keep.txt']);
  git(dir, ['commit', '-m', 'start']);
  return dir;
}

test('open and commit change one file inside the repo', async () => {
  const dir = initRepo();
  try {
    const opened = await openGitFile({ repo: dir, path: 'notes.txt' });
    assert.equal(opened.text, 'hello\n');
    const committed = await commitGitFile({ repo: dir, path: 'notes.txt', content: 'hello world', message: 'Update the note' });
    assert.equal(committed.ok, true);
    assert.match(fs.readFileSync(path.join(dir, 'notes.txt'), 'utf8'), /hello world\n/);
    assert.equal(fs.readFileSync(path.join(dir, 'keep.txt'), 'utf8'), 'leave me\n');
    const same = await commitGitFile({ repo: dir, path: 'notes.txt', content: 'hello world\n', message: 'again' });
    assert.equal(same.unchanged, true);
    const outside = await openGitFile({ repo: dir, path: '../outside.txt' });
    assert.match(outside.error, /inside the repo/i);
    const secret = await commitGitFile({ repo: dir, path: 'notes.txt', content: 'token ghp_secretvalue', message: 'no' });
    assert.match(secret.error, /secret/i);
    assert.equal(fs.readFileSync(path.join(dir, 'notes.txt'), 'utf8').includes('ghp_'), false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('git pull uses fast-forward only and a secret path is refused before git runs', async () => {
  const calls = [];
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-git-pull-'));
  fs.mkdirSync(path.join(dir, '.git'));
  fs.writeFileSync(path.join(dir, 'notes.txt'), 'hello\n');
  const spawnImpl = async (args, cwd) => {
    calls.push({ args, cwd });
    if (args[0] === 'rev-parse' && args[1] === '--show-toplevel') return { code: 0, stdout: dir, stderr: '' };
    if (args[0] === 'pull') return { code: 0, stdout: 'Already up to date.', stderr: '' };
    return { code: 0, stdout: '', stderr: '' };
  };
  try {
    const pulled = await pullGitFile({ repo: dir, path: 'notes.txt', spawnImpl });
    assert.equal(pulled.text, 'hello\n');
    assert.deepEqual(calls.find(call => call.args[0] === 'pull').args, ['pull', '--ff-only']);
    calls.length = 0;
    const env = await openGitFile({ repo: dir, path: '.env', spawnImpl });
    assert.match(env.error, /secret/i);
    assert.equal(calls.some(call => call.args[0] === 'pull'), false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('git push sends only git push and does not add --force or -u', async () => {
  const calls = [];
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-git-push-'));
  fs.mkdirSync(path.join(dir, '.git'));
  const spawnImpl = async (args) => {
    calls.push(args);
    if (args[0] === 'rev-parse' && args[1] === '--show-toplevel') return { code: 0, stdout: dir, stderr: '' };
    if (args[0] === 'push') return { code: 0, stdout: '', stderr: 'main -> main' };
    return { code: 1, stdout: '', stderr: 'unexpected' };
  };
  try {
    const pushed = await pushGitBranch({ repo: dir, spawnImpl });
    assert.equal(pushed.ok, true);
    assert.deepEqual(calls.find(args => args[0] === 'push'), ['push']);
    calls.length = 0;
    const spawnFail = async (args) => {
      calls.push(args);
      if (args[0] === 'rev-parse') return { code: 0, stdout: dir, stderr: '' };
      return { code: 1, stdout: '', stderr: 'fatal: The current branch main has no upstream branch. git push --set-upstream origin main' };
    };
    const missing = await pushGitBranch({ repo: dir, spawnImpl: spawnFail });
    assert.match(missing.error, /no upstream/i);
    assert.equal(calls.some(args => args.includes('--force') || args.includes('-u') || args.includes('--set-upstream')), false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
