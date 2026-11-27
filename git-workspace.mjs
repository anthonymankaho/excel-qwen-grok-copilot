// Local git for one code file opened in Word. Uses the Git login already on this PC.
// The relay runs `git pull --ff-only`, `git add`, `git commit` for that file, and `git push` with no --force.
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const MAX_FILE_BYTES = 64 * 1024;
const GIT_TIMEOUT_MS = 90_000;
const GIT_OUTPUT_CAP = 8000;
const SECRET_RE = /(?:sk-[A-Za-z0-9]{8,}|ghp_[A-Za-z0-9]{8,}|github_pat_[A-Za-z0-9_]{8,}|AKIA[0-9A-Z]{8,}|-----BEGIN [A-Z ]*PRIVATE KEY-----|xox[baprs]-[A-Za-z0-9-]{8,})/;
const SECRET_NAMES = new Set(['.env', '.env.local', '.env.production', 'id_rsa', 'id_dsa', 'id_ecdsa', 'id_ed25519']);

export function gitSecretError(text) {
  if (SECRET_RE.test(String(text || ''))) return 'This file looks like a secret. It was not opened or committed.';
  return '';
}

function cleanGitText(value) {
  const text = String(value || '').replace(/\u0000/g, '').trim();
  return text.length > 500 ? text.slice(-500) : text;
}

function appendCapped(current, chunk) {
  const next = current + String(chunk);
  return next.length > GIT_OUTPUT_CAP ? next.slice(-GIT_OUTPUT_CAP) : next;
}

function runGit(cwd, args, spawnImpl) {
  if (spawnImpl) return spawnImpl(args, cwd);
  return new Promise(resolve => {
    let stdout = '';
    let stderr = '';
    let settled = false;
    const finish = result => {
      if (settled) return;
      settled = true;
      resolve(result);
    };
    let child;
    try {
      child = spawn('git', args, { cwd, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
    } catch (error) {
      finish({ code: 127, stdout: '', stderr: cleanGitText(error.message) });
      return;
    }
    const timer = setTimeout(() => {
      child.kill();
      finish({ code: 1, stdout, stderr: 'Git did not finish. If it is waiting for a login, sign in and try again.' });
    }, GIT_TIMEOUT_MS);
    child.stdout.on('data', chunk => { stdout = appendCapped(stdout, chunk); });
    child.stderr.on('data', chunk => { stderr = appendCapped(stderr, chunk); });
    child.once('error', error => {
      clearTimeout(timer);
      finish({ code: 127, stdout, stderr: cleanGitText(error.message) });
    });
    child.once('close', code => {
      clearTimeout(timer);
      finish({ code: code ?? 1, stdout, stderr });
    });
  });
}

export async function resolveGitRepo(folder, spawnImpl) {
  const root = path.resolve(String(folder || '').trim());
  if (!root || !fs.existsSync(root)) return { error: 'Set a local repo folder that exists.' };
  let stat;
  try { stat = fs.statSync(root); } catch { return { error: 'Set a local repo folder that exists.' }; }
  if (!stat.isDirectory()) return { error: 'Set a local repo folder that exists.' };
  const top = await runGit(root, ['rev-parse', '--show-toplevel'], spawnImpl);
  if (top.code !== 0) return { error: 'That folder is not a Git repository.' };
  return { root: path.resolve(cleanGitText(top.stdout)) };
}

export function resolveGitFile(root, filePath) {
  const raw = String(filePath || '').trim();
  if (!raw || raw.includes('\0') || raw.includes('\n')) return { error: 'Name one file. Example: open src/app.js' };
  const absolute = path.resolve(root, raw);
  const relative = path.relative(root, absolute);
  if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) return { error: 'The file must stay inside the repo.' };
  const base = path.basename(absolute).toLowerCase();
  if (SECRET_NAMES.has(base) || base.endsWith('.pem') || base.endsWith('.key')) {
    return { error: 'This file looks like a secret. It was not opened or committed.' };
  }
  return { absolute, relative: relative.split(path.sep).join('/') };
}

function readTextFile(absolute) {
  if (!fs.existsSync(absolute) || !fs.statSync(absolute).isFile()) return { error: 'That file is not in the repo.' };
  if (fs.statSync(absolute).size > MAX_FILE_BYTES) return { error: 'This file is over 64 KB.' };
  const original = fs.readFileSync(absolute);
  if (original.includes(0)) return { error: 'This file is not text.' };
  const text = original.toString('utf8');
  const secret = gitSecretError(text);
  if (secret) return { error: secret };
  return { text, original: text, bytes: original };
}

async function restoreFile(root, relative, absolute, bytes, added, spawnImpl) {
  try { fs.writeFileSync(absolute, bytes); } catch { /* the git error is the one shown */ }
  if (added) await runGit(root, ['reset', '-q', '--', relative], spawnImpl);
}

function withRepoNewlines(original, incoming) {
  const normalized = String(incoming || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const newline = original.includes('\r\n') ? '\r\n' : '\n';
  let body = normalized.replace(/\n/g, newline);
  const originalEnds = /\r?\n$/.test(original);
  const bodyEnds = body.endsWith(newline);
  if (originalEnds && !bodyEnds) body += newline;
  if (!originalEnds && bodyEnds) body = body.slice(0, -newline.length);
  return body;
}

async function locate(repo, filePath, spawnImpl) {
  const raw = String(filePath || '').trim();
  if (path.isAbsolute(raw)) {
    const found = await resolveGitRepo(path.dirname(raw), spawnImpl);
    if (found.error) return found;
    const file = resolveGitFile(found.root, path.relative(found.root, path.resolve(raw)));
    if (file.error) return file;
    return { root: found.root, ...file };
  }
  const found = await resolveGitRepo(repo, spawnImpl);
  if (found.error) return found;
  const file = resolveGitFile(found.root, raw);
  if (file.error) return file;
  return { root: found.root, ...file };
}

export async function openGitFile({ repo, path: filePath, spawnImpl } = {}) {
  const located = await locate(repo, filePath, spawnImpl);
  if (located.error) return { error: located.error, status: 400 };
  const read = readTextFile(located.absolute);
  if (read.error) return { error: read.error, status: 400 };
  return { ok: true, repo: located.root, path: located.relative, text: read.text };
}

export async function pullGitFile({ repo, path: filePath, spawnImpl } = {}) {
  const located = await locate(repo, filePath, spawnImpl);
  if (located.error) return { error: located.error, status: 400 };
  const pulled = await runGit(located.root, ['pull', '--ff-only'], spawnImpl);
  if (pulled.code === 127) return { error: 'Git is not available on this PC.', status: 400 };
  if (pulled.code !== 0) return { error: cleanGitText(pulled.stderr) || cleanGitText(pulled.stdout) || 'git pull failed.', status: 400 };
  const read = readTextFile(located.absolute);
  if (read.error) return { error: read.error, status: 400 };
  return { ok: true, repo: located.root, path: located.relative, text: read.text, summary: cleanGitText(pulled.stdout) || 'git pull completed.' };
}

export async function commitGitFile({ repo, path: filePath, content, message, spawnImpl } = {}) {
  if (typeof content !== 'string') return { error: 'The Word document was empty.', status: 400 };
  if (Buffer.byteLength(content) > MAX_FILE_BYTES) return { error: 'This file is over 64 KB.', status: 400 };
  const secret = gitSecretError(content);
  if (secret) return { error: secret, status: 400 };
  const located = await locate(repo, filePath, spawnImpl);
  if (located.error) return { error: located.error, status: 400 };
  const existing = readTextFile(located.absolute);
  if (existing.error) return { error: existing.error, status: 400 };
  const next = withRepoNewlines(existing.original, content);
  if (Buffer.byteLength(next) > MAX_FILE_BYTES) return { error: 'This file is over 64 KB.', status: 400 };
  if (next === existing.original) return { ok: true, unchanged: true, repo: located.root, path: located.relative };
  fs.writeFileSync(located.absolute, next);
  const added = await runGit(located.root, ['add', '--', located.relative], spawnImpl);
  if (added.code !== 0) {
    await restoreFile(located.root, located.relative, located.absolute, existing.bytes, false, spawnImpl);
    return { error: cleanGitText(added.stderr) || 'git add failed.', status: 400 };
  }
  const title = String(message || 'Update from Word').trim().slice(0, 500) || 'Update from Word';
  const committed = await runGit(located.root, ['commit', '-m', title, '--', located.relative], spawnImpl);
  if (committed.code !== 0) {
    await restoreFile(located.root, located.relative, located.absolute, existing.bytes, true, spawnImpl);
    return { error: cleanGitText(committed.stderr) || cleanGitText(committed.stdout) || 'git commit failed.', status: 400 };
  }
  const shown = await runGit(located.root, ['rev-parse', '--short', 'HEAD'], spawnImpl);
  return { ok: true, repo: located.root, path: located.relative, commit: cleanGitText(shown.stdout), message: title };
}

export async function pushGitBranch({ repo, spawnImpl } = {}) {
  const found = await resolveGitRepo(repo, spawnImpl);
  if (found.error) return { error: found.error, status: 400 };
  const pushed = await runGit(found.root, ['push'], spawnImpl);
  if (pushed.code === 127) return { error: 'Git is not available on this PC.', status: 400 };
  const detail = cleanGitText(pushed.stderr) || cleanGitText(pushed.stdout);
  if (pushed.code !== 0) {
    if (/no upstream|set-upstream|has no upstream/i.test(`${pushed.stderr}\n${pushed.stdout}`)) {
      return { error: 'This branch has no upstream. Push was not run with --force. Set an upstream in Git, then type git push again.', status: 400 };
    }
    return { error: detail || 'git push failed.', status: 400 };
  }
  return { ok: true, repo: found.root, summary: detail || 'git push completed.' };
}
