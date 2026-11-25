import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateRawSync, inflateRawSync, crc32 } from 'node:zlib';

const ROOT_NAME = 'Qwen-for-Excel-and-Word';
const SKIP_DIRS = new Set(['node_modules', '.git', '.wrangler', 'download', 'store-assets']);

export const START_HERE = `Qwen for Excel and Word

This folder has no API keys. Each person uses their own key.

INSTALL FROM A TERMINAL  (recommended)
1. Install Node.js 20 or newer from https://nodejs.org
2. Close Excel and Word.
3. Open PowerShell in this folder and run:

     npm run setup
     npm start

   If the folder path contains spaces, quote it:
     cd "C:\\Users\\Your Name\\Qwen-for-Excel-and-Word"

   npm run setup checks Node 20+, runs npm install, and prints the next steps.
   npm start sideloads the add-in and opens Excel. Leave that terminal window
   open - it IS the relay.

   Plain npm works the same:  npm install  then  npm start.
   Neither route installs auto-start or writes any key.

OR DOUBLE-CLICK
   Setup and Start.bat does the same and ALSO installs a logon auto-start
   plus a per-minute keep-alive task. Skip it if you do not want anything
   running at sign-in. npm start alone is enough to use the add-in.

FIRST USE
4. In Excel: Home > Add-ins > Qwen for Excel and Word.
5. Settings > paste YOUR OWN Qwen key > leave Remember
   ticked > Use key. The dot turns green.

OPTIONAL CLOUD SOURCES
   The app owner sets the public OAuth client IDs once in cloud-oauth.json
   before sharing the package. Each colleague then clicks Connect and signs
   in to their own Microsoft account. No account address or token
   is built into the app.

WORD
   Close Word, then in this folder run:  npm run start:word

MODEL AND EFFORT
   The composer offers auto, qwen4-max, qwen3.8-max, and qwen3.8-flash.
   Effort is Off, Low, Medium, High, or Extra. Fast skips thinking.

EXCEL ON THE WEB
   Upload manifest.cloudflare.xml from Home > Add-ins > Upload My Add-in.
   Pane address: https://excel-qwen-grok.anthonym-354.workers.dev
   Each person pastes their own Qwen key. Do not use a Token Plan key
   with this custom add-in unless Alibaba confirms permission.
   No provider key is in this zip. The hosted Worker was not updated
   with the 2026-11-25.2 local build.

DOCS
   SOP-INSTALL.md     installation, CLI steps, troubleshooting
   SOP-USER-GUIDE.md  how to use it / 說明書
   README.md          full reference

STOP
   npm run relay-stop   stops the relay
   npm run stop         removes the sideload from Excel
`;

export function shouldSkip(relativePath) {
  const parts = relativePath.split(/[\\/]/).filter(Boolean);
  if (parts.some((part) => SKIP_DIRS.has(part))) return true;
  const name = parts.at(-1) || '';
  if (name === '.env' || name.startsWith('.env.')) return true;
  if (name.endsWith('.log') || name === 'Thumbs.db' || name === '.DS_Store') return true;
  return false;
}

async function collect(dir, prefix = '') {
  const entries = [];
  for (const name of await readdir(dir)) {
    const relative = prefix ? `${prefix}/${name}` : name;
    if (shouldSkip(relative)) continue;
    const full = path.join(dir, name);
    const info = await stat(full);
    if (info.isDirectory()) entries.push(...await collect(full, relative));
    else if (info.isFile()) entries.push({ relative: relative.replaceAll('\\', '/'), full });
  }
  return entries;
}

function dosTime(date) {
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const day = ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { time, day };
}

export function zipBytes(files, stamp = new Date('2026-11-22T12:00:00Z')) {
  const { time, day } = dosTime(stamp);
  const locals = [];
  const central = [];
  let offset = 0;
  for (const file of files) {
    const name = Buffer.from(file.name);
    const data = Buffer.isBuffer(file.data) ? file.data : Buffer.from(file.data);
    const compressed = deflateRawSync(data);
    const sum = crc32(data) >>> 0;
    const local = Buffer.alloc(30 + name.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(time, 10);
    local.writeUInt16LE(day, 12);
    local.writeUInt32LE(sum, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    name.copy(local, 30);
    locals.push(local, compressed);
    const header = Buffer.alloc(46 + name.length);
    header.writeUInt32LE(0x02014b50, 0);
    header.writeUInt16LE(20, 4);
    header.writeUInt16LE(20, 6);
    header.writeUInt16LE(0x0800, 8);
    header.writeUInt16LE(8, 10);
    header.writeUInt16LE(time, 12);
    header.writeUInt16LE(day, 14);
    header.writeUInt32LE(sum, 16);
    header.writeUInt32LE(compressed.length, 20);
    header.writeUInt32LE(data.length, 24);
    header.writeUInt16LE(name.length, 28);
    header.writeUInt32LE(offset, 42);
    name.copy(header, 46);
    central.push(header);
    offset += local.length + compressed.length;
  }
  const directory = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, directory, end]);
}

// Pure-Node zip reading, so the pack can be verified on Windows where the Unix
// unzip binary does not exist. zipBytes writes no data descriptors, so the local
// headers carry the compressed size.
function walkEntries(bytes) {
  const entries = [];
  let offset = 0;
  while (offset + 30 <= bytes.length && bytes.readUInt32LE(offset) === 0x04034b50) {
    const method = bytes.readUInt16LE(offset + 8);
    const compressedSize = bytes.readUInt32LE(offset + 18);
    const nameLength = bytes.readUInt16LE(offset + 26);
    const extraLength = bytes.readUInt16LE(offset + 28);
    const name = bytes.toString('utf8', offset + 30, offset + 30 + nameLength);
    const start = offset + 30 + nameLength + extraLength;
    entries.push({ name, method, data: bytes.subarray(start, start + compressedSize) });
    offset = start + compressedSize;
  }
  return entries;
}

export function listZipEntries(bytes) {
  return walkEntries(bytes).map((entry) => entry.name);
}

export function readZipEntry(bytes, name) {
  const entry = walkEntries(bytes).find((item) => item.name === name);
  if (!entry) return null;
  if (entry.method === 0) return entry.data.toString('utf8');
  if (entry.method === 8) return inflateRawSync(entry.data).toString('utf8');
  throw new Error(`Unsupported zip compression method ${entry.method} for ${name}`);
}

export async function makeDownloadZip(projectRoot, outFile) {
  const files = await collect(projectRoot);
  files.sort((a, b) => a.relative.localeCompare(b.relative));
  const packed = [
    { name: `${ROOT_NAME}/START-HERE.txt`, data: START_HERE },
    ...await Promise.all(files.map(async (file) => ({
      name: `${ROOT_NAME}/${file.relative}`,
      data: await readFile(file.full)
    })))
  ];
  packed.sort((a, b) => a.name.localeCompare(b.name));
  await mkdir(path.dirname(outFile), { recursive: true });
  await writeFile(outFile, zipBytes(packed));
  return packed.map((file) => file.name);
}

const here = path.dirname(fileURLToPath(import.meta.url));
const isDirectRun = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isDirectRun) {
  const projectRoot = path.resolve(here, '..');
  const outFile = path.resolve(projectRoot, process.argv[2] || 'download/Qwen-for-Excel-and-Word.zip');
  const names = await makeDownloadZip(projectRoot, outFile);
  console.log(`${outFile} (${names.length} files)`);
}
