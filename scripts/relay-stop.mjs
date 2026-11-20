import { execSync } from 'node:child_process';

const out = execSync('netstat -ano | findstr ":3000" | findstr "LISTENING"', { encoding: 'utf8' });
const pids = [...new Set(out.trim().split(/\r?\n/).map(line => line.trim().split(/\s+/).pop()))];
if (!pids.length) { console.log('No relay listening on port 3000.'); process.exit(0); }
for (const pid of pids) {
  execSync(`taskkill /pid ${pid} /f`, { stdio: 'inherit' });
}
console.log('Relay stopped. Note: the per-minute ensure task will restart it within a minute unless auto-start is removed (npm run autostart-remove).');
