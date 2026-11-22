import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import { makeDownloadZip, shouldSkip, zipBytes } from '../scripts/make-download-zip.mjs';

const projectRoot = path.resolve(import.meta.dirname, '..');

test('the download zip skips keys, git, and installed packages', () => {
  assert.equal(shouldSkip('.env'), true);
  assert.equal(shouldSkip('.env.local'), true);
  assert.equal(shouldSkip('node_modules/left-pad/index.js'), true);
  assert.equal(shouldSkip('.git/config'), true);
  assert.equal(shouldSkip('.wrangler/state.json'), true);
  assert.equal(shouldSkip('download/Qwen-Grok-for-Excel-and-Word.zip'), true);
  assert.equal(shouldSkip('public/index.html'), false);
  assert.equal(shouldSkip('manifest.cloudflare.xml'), false);
});

test('a packed file round-trips through the zip', async () => {
  const bytes = zipBytes([{ name: 'Qwen-Grok-for-Excel-and-Word/hello.txt', data: 'ready' }]);
  const dir = await mkdtemp(path.join(tmpdir(), 'eqgc-zip-'));
  const file = path.join(dir, 'sample.zip');
  await writeFile(file, bytes);
  const { stdout } = await promisify(execFile)('unzip', ['-p', file, 'Qwen-Grok-for-Excel-and-Word/hello.txt']);
  assert.equal(stdout, 'ready');
  await rm(dir, { recursive: true, force: true });
});

test('the project zip has the installer and leaves secrets out', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'eqgc-pack-'));
  const outFile = path.join(dir, 'Qwen-Grok-for-Excel-and-Word.zip');
  try {
    const names = await makeDownloadZip(projectRoot, outFile);
    assert.ok(names.includes('Qwen-Grok-for-Excel-and-Word/START-HERE.txt'));
    assert.ok(names.includes('Qwen-Grok-for-Excel-and-Word/Setup and Start.bat'));
    assert.ok(names.includes('Qwen-Grok-for-Excel-and-Word/manifest.xml'));
    assert.ok(names.includes('Qwen-Grok-for-Excel-and-Word/manifest.cloudflare.xml'));
    assert.ok(names.includes('Qwen-Grok-for-Excel-and-Word/public/index.html'));
    assert.equal(names.some((name) => name.includes('.env') || name.includes('node_modules') || name.includes('.git/')), false);
    const saved = await readFile(outFile);
    assert.equal(saved.readUInt32LE(0), 0x04034b50);
    const { stdout } = await promisify(execFile)('unzip', ['-p', outFile, 'Qwen-Grok-for-Excel-and-Word/START-HERE.txt']);
    assert.match(stdout, /This folder has no API keys/);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
