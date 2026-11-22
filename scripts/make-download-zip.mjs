import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateRawSync, crc32 } from 'node:zlib';

const ROOT_NAME = 'Qwen-Grok-for-Excel-and-Word';
const SKIP_DIRS = new Set(['node_modules', '.git', '.wrangler', 'download']);

export const START_HERE = `Qwen + Grok for Excel and Word

This folder has no API keys.

On this Windows PC
1. Install Node.js 20 or newer from https://nodejs.org
2. Close Excel and Word.
3. Double-click Setup and Start.bat.
4. In Excel, open Home > Add-ins. The name is Qwen + Grok for Excel and Word.
5. Open Settings, paste your own Qwen or Grok key, leave Remember ticked, and press Use key.

Word needs one extra step: close Word, open a terminal in this folder, and run npm run start:word.

Excel on the web
Upload manifest.cloudflare.xml from Home > Add-ins > Upload My Add-in.
The pane address is https://excel-qwen-grok.anthonym-354.workers.dev
Paste the relay token in Settings. Qwen and Grok keys are Worker secrets, set with wrangler, and they are not in this zip.

Read README.md for the full guide.
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
  const outFile = path.resolve(projectRoot, process.argv[2] || 'download/Qwen-Grok-for-Excel-and-Word.zip');
  const names = await makeDownloadZip(projectRoot, outFile);
  console.log(`${outFile} (${names.length} files)`);
}
