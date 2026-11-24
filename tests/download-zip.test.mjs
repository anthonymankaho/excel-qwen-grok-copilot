import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { listZipEntries, makeDownloadZip, readZipEntry, shouldSkip, zipBytes } from '../scripts/make-download-zip.mjs';

const projectRoot = path.resolve(import.meta.dirname, '..');

test('the download zip skips keys, git, and installed packages', () => {
  assert.equal(shouldSkip('.env'), true);
  assert.equal(shouldSkip('.env.local'), true);
  assert.equal(shouldSkip('node_modules/left-pad/index.js'), true);
  assert.equal(shouldSkip('.git/config'), true);
  assert.equal(shouldSkip('.wrangler/state.json'), true);
  assert.equal(shouldSkip('download/Qwen-for-Excel-and-Word.zip'), true);
  assert.equal(shouldSkip('public/index.html'), false);
  assert.equal(shouldSkip('manifest.cloudflare.xml'), false);
});

test('a packed file round-trips through the zip', async () => {
  const bytes = zipBytes([{ name: 'Qwen-for-Excel-and-Word/hello.txt', data: 'ready' }]);
  const dir = await mkdtemp(path.join(tmpdir(), 'eqgc-zip-'));
  const file = path.join(dir, 'sample.zip');
  await writeFile(file, bytes);
  const saved = await readFile(file);
  assert.equal(readZipEntry(saved, 'Qwen-for-Excel-and-Word/hello.txt'), 'ready');
  assert.deepEqual(listZipEntries(saved), ['Qwen-for-Excel-and-Word/hello.txt']);
  assert.equal(readZipEntry(saved, 'Qwen-for-Excel-and-Word/absent.txt'), null);
  await rm(dir, { recursive: true, force: true });
});

test('store assets exist as PNGs and stay out of the install zip', async () => {
  assert.equal(shouldSkip('store-assets/logo-96.png'), true);
  for (const name of ['logo-96.png', 'logo-64.png', 'logo-32.png', 'screenshot-1-sidebar-guide.png', 'screenshot-2-attach-menu.png', 'screenshot-3-settings-byo.png']) {
    const bytes = await readFile(path.join(projectRoot, 'store-assets', name));
    assert.equal(bytes.readUInt32BE(0), 0x89504e47, `${name} is a PNG`);
  }
});

test('store pack: legal pages disclose BYO and the manifest points at them', async () => {
  const privacy = await readFile(path.join(projectRoot, 'public', 'legal', 'privacy.html'), 'utf8');
  assert.match(privacy, /never stores or logs it/);
  assert.match(privacy, /Share/);
  const eula = await readFile(path.join(projectRoot, 'public', 'legal', 'eula.html'), 'utf8');
  assert.match(eula, /your own/);
  const support = await readFile(path.join(projectRoot, 'public', 'legal', 'support.html'), 'utf8');
  assert.match(support, /GitHub Issues/);
  const manifest = await readFile(path.join(projectRoot, 'manifest.cloudflare.xml'), 'utf8');
  assert.match(manifest, /<Version>1\.5\.0\.0<\/Version>/);
  assert.match(manifest, /<SupportUrl DefaultValue="https:\/\/excel-qwen-grok\.anthonym-354\.workers\.dev\/legal\/support\.html"\/>/);
  assert.match(manifest, /<PrivacyUrl DefaultValue="https:\/\/excel-qwen-grok\.anthonym-354\.workers\.dev\/legal\/privacy\.html"\/>/);
  assert.match(manifest, /<TermsOfUseUrl DefaultValue="https:\/\/excel-qwen-grok\.anthonym-354\.workers\.dev\/legal\/eula\.html"\/>/);
});

test('the project zip has the installer and leaves secrets out', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'eqgc-pack-'));
  const outFile = path.join(dir, 'Qwen-for-Excel-and-Word.zip');
  try {
    const names = await makeDownloadZip(projectRoot, outFile);
    for (const expected of [
      'START-HERE.txt', 'Setup and Start.bat', 'manifest.xml', 'manifest.cloudflare.xml',
      'public/index.html', 'package.json', 'SOP-INSTALL.md', 'SOP-USER-GUIDE.md', 'README.md',
      'public/legal/privacy.html', 'public/legal/eula.html', 'public/legal/support.html'
    ]) {
      assert.ok(names.includes(`Qwen-for-Excel-and-Word/${expected}`), `${expected} is packed`);
    }
    assert.equal(names.some((name) => name.includes('.env') || name.includes('node_modules') || name.includes('.git/')), false);
    const saved = await readFile(outFile);
    assert.equal(saved.readUInt32LE(0), 0x04034b50);
    assert.deepEqual(listZipEntries(saved), names);
    const startHere = readZipEntry(saved, 'Qwen-for-Excel-and-Word/START-HERE.txt');
    assert.match(startHere, /This folder has no API keys/);
    assert.match(startHere, /npm install/);
    assert.match(startHere, /npm start/);
    assert.match(startHere, /SOP-INSTALL\.md/);
    assert.match(startHere, /SOP-USER-GUIDE\.md/);
    const pkg = JSON.parse(readZipEntry(saved, 'Qwen-for-Excel-and-Word/package.json'));
    assert.ok(pkg.scripts.start, 'the packed package.json can launch the add-in');
    assert.ok(pkg.scripts.setup, 'the packed package.json has the CLI setup entry point');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
