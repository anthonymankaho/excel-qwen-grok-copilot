import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
if (!process.env.APPDATA) {
  console.error('Auto-start is only supported on Windows (APPDATA is not set).');
  process.exit(1);
}
const startupDir = path.join(process.env.APPDATA, 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup');
const item = path.join(startupDir, 'ExcelQwenGrokRelay.bat');
const TASK = 'ExcelQwenGrokRelayEnsure';
const ensure = path.join(root, 'scripts', 'relay-ensure.mjs');

if (process.argv.includes('--remove')) {
  if (fs.existsSync(item)) fs.unlinkSync(item);
  try { execSync(`schtasks /delete /tn "${TASK}" /f`, { stdio: 'inherit' }); } catch { /* task not installed */ }
  console.log('Auto-start and self-heal removed. Stop a running relay with: npm run relay-stop');
  process.exit(0);
}

// Per-user Startup item: starts the hidden relay at logon; its console closes immediately.
fs.writeFileSync(item, `@echo off\r\ncd /d "${root}"\r\nstart "ExcelQwenGrokRelay" /min node scripts\\relay-ensure.mjs\r\n`, 'utf8');
console.log(`Startup item installed: ${item}`);

// Self-heal backstop: per-minute ensure task (time trigger, current user, no elevation).
try {
  execSync(`schtasks /create /tn "${TASK}" /tr "\\"${process.execPath}\\" \\"${ensure}\\"" /sc minute /mo 1 /f`, { stdio: 'inherit' });
  console.log(`Self-heal task installed: "${TASK}" (every minute, starts the relay hidden if down).`);
} catch {
  console.log('NOTE: could not create the per-minute self-heal task (permissions). Logon auto-start still works; restart manually with Start Local Relay.bat if the relay dies mid-session.');
}
console.log('Start the relay now with: npm run relay-ensure');
console.log('Remove later with: npm run autostart-remove');
