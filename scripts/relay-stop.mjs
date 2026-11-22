import { execSync } from 'node:child_process';

let out = '';
try {
  out = execSync('netstat -ano | findstr ":3000" | findstr "LISTENING"', { encoding: 'utf8' });
} catch {
  // findstr exits 1 when nothing matches; that means no relay is running, not an error.
}
const pids = [...new Set(out.trim().split(/\r?\n/).map(line => line.trim().split(/\s+/).pop()).filter(Boolean))];
if (!pids.length) { console.log('No relay listening on port 3000.'); process.exit(0); }
for (const pid of pids) {
  execSync(`taskkill /pid ${pid} /f`, { stdio: 'inherit' });
}
console.log('Relay stopped.');
try {
  execSync('schtasks /query /tn "ExcelQwenGrokRelayEnsure"', { stdio: 'ignore' });
  console.log('Note: the per-minute ensure task will restart it within a minute unless auto-start is removed (npm run autostart-remove).');
} catch {
  // No ensure task installed, so nothing will restart it.
}
