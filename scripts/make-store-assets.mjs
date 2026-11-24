// Generates Marketplace listing assets: pane screenshots (1280x800) and store logos.
// Uses the owner PC's playwright-core install; override with PLAYWRIGHT_ROOT if moved.
// Serves public/ locally with stub /api responses so the pane renders cleanly without Office.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outDir = path.join(root, 'store-assets');
fs.mkdirSync(outDir, { recursive: true });
const require = createRequire(process.env.PLAYWRIGHT_ROOT || 'C:/Users/User/playwright-verify/package.json');
const { chromium } = require('playwright-core');
const SHELL = process.env.CHROME_HEADLESS_SHELL || 'C:/Users/User/AppData/Local/ms-playwright/chromium_headless_shell-1243/chrome-headless-shell-win64/chrome-headless-shell.exe';

const MIME = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.json': 'application/json' };
const API_STUBS = {
  '/api/health': { ok: true, build: '2026-11-24.9', configured: { qwen: true } },
  '/api/skills': { root: '', skills: [] },
  '/api/skills/context': { selected: [], text: '' },
  '/api/sessions': { sessions: [] },
  '/api/memory': { text: '' },
  '/api/connectors/status': { onedrive: { configured: false, connected: false } }
};
const server = http.createServer((req, res) => {
  const urlPath = decodeURIComponent(req.url.split('?')[0]);
  if (API_STUBS[urlPath]) {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(API_STUBS[urlPath]));
    return;
  }
  const file = path.join(root, 'public', urlPath === '/' ? 'index.html' : urlPath);
  if (!file.startsWith(path.join(root, 'public')) || !fs.existsSync(file) || !fs.statSync(file).isFile()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
  fs.createReadStream(file).pipe(res);
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const base = `http://127.0.0.1:${server.address().port}`;

const browser = await chromium.launch({ executablePath: SHELL, args: ['--no-sandbox'] });
const errors = [];
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  page.on('pageerror', error => errors.push(String(error)));
  await page.goto(base, { waitUntil: 'networkidle' });
  await page.click('#sidebarToggle');
  await page.waitForTimeout(300);
  await page.screenshot({ path: path.join(outDir, 'screenshot-1-sidebar-guide.png') });
  await page.click('#sidebarClose');
  await page.click('#attachFiles');
  await page.waitForTimeout(200);
  await page.screenshot({ path: path.join(outDir, 'screenshot-2-attach-menu.png') });
  await page.click('#attachFiles');
  await page.click('#settingsToggle');
  await page.waitForTimeout(200);
  await page.screenshot({ path: path.join(outDir, 'screenshot-3-settings-byo.png') });

  for (const size of [96, 64, 32]) {
    const logo = await browser.newPage({ viewport: { width: size, height: size } });
    await logo.setContent(`<body style="margin:0"><div style="width:${size}px;height:${size}px;border-radius:${Math.round(size * 0.22)}px;background:#4f46e5;display:flex;align-items:center;justify-content:center;color:#fff;font:${Math.round(size * 0.55)}px Segoe UI,Arial,sans-serif">✦</div></body>`);
    await logo.screenshot({ path: path.join(outDir, `logo-${size}.png`) });
    await logo.close();
  }
} finally {
  await browser.close();
  server.close();
}
if (errors.length) { console.log('PAGEERRORS', errors); process.exit(1); }
console.log('ASSETS', fs.readdirSync(outDir).join(', '));