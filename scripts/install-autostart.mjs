import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const startupDir = path.join(process.env.APPDATA, 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup');
const item = path.join(startupDir, 'ExcelQwenGrokRelay.bat');

if (process.argv.includes('--remove')) {
  if (fs.existsSync(item)) fs.unlinkSync(item);
  console.log('Auto-start removed. The relay can still be started manually with Start Local Relay.bat.');
  process.exit(0);
}

// Per-user Startup item: no elevation, no script engine, no PowerShell.
// The relay runs in a minimized console so it stays stoppable with Ctrl+C / close.
fs.writeFileSync(item, `@echo off\r\ncd /d "${root}"\r\nstart "ExcelQwenGrokRelay" /min node server.mjs\r\n`, 'utf8');
console.log(`Auto-start installed: ${item}`);
console.log('The relay starts minimized at every logon. Start it now by double-clicking that file or Start Local Relay.bat.');
console.log('Remove later with: npm run autostart-remove');
