import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));

async function healthy() {
  try {
    const r = await fetch('http://127.0.0.1:3000/api/health', { signal: AbortSignal.timeout(2500) });
    return r.ok;
  } catch { return false; }
}

if (await healthy()) process.exit(0);
spawn(process.execPath, [path.join(root, 'server.mjs')], {
  cwd: root, detached: true, stdio: 'ignore', windowsHide: true
}).unref();
console.log('Relay was down; started hidden.');
