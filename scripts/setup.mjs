// CLI installer: checks Node, installs dependencies, then prints the exact next command.
// Works from an extracted zip or a git clone, on any OS.
import { execSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const major = Number(process.versions.node.split('.')[0]);

console.log('Qwen + Grok for Excel and Word - setup');
console.log(`Folder: "${root}"`);

if (!Number.isFinite(major) || major < 20) {
  console.error(`\nNode.js 20 or newer is required; this is ${process.versions.node}.`);
  console.error('Install it from https://nodejs.org and run this again.');
  process.exit(1);
}
console.log(`Node ${process.versions.node} - OK`);

console.log('\nInstalling dependencies...');
execSync('npm install --no-fund --no-audit', { cwd: root, stdio: 'inherit' });

console.log('\nDependencies installed. No API key was written anywhere.');
console.log('\nNEXT STEPS');
console.log('  1. Close Excel and Word.');
console.log('  2. In this folder run:  npm start');
console.log('     That sideloads the add-in and opens Excel. Leave the window open - it is the relay.');
console.log('  3. In Excel: Home > Add-ins > Qwen + Grok for Excel and Word.');
console.log('  4. Settings > pick Qwen or Grok > paste YOUR OWN key > Remember ticked > Use key.');
console.log('  5. Optional cloud sources: app owner sets shared client IDs once in cloud-oauth.json; each person clicks Connect and signs in to their own account. See README.md.');
console.log('\n  Word instead of Excel:  npm run start:word');
console.log('  Stop the relay:         npm run relay-stop');
console.log('  Remove the sideload:    npm run stop');
console.log('\n  Auto-start at sign-in is NOT installed by this script.');
console.log('  Only run "npm run autostart" if you want it, and "npm run autostart-remove" to undo.');
console.log('\n  Docs: SOP-INSTALL.md (install) and SOP-USER-GUIDE.md (how to use / 說明書).');
