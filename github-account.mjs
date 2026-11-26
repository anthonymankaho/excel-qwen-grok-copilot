// GitHub account for the local pane. The token stays in the GitHub CLI credential store.
// Status responses never include the token.
import { execFile, spawn } from 'node:child_process';

let execOverride = null;
let spawnOverride = null;

export function setGithubAccountHooksForTests(hooks) {
  execOverride = hooks?.exec || null;
  spawnOverride = hooks?.spawn || null;
}

function execGh(args) {
  if (execOverride) return execOverride(args);
  return new Promise(resolve => {
    execFile('gh', args, { windowsHide: true, timeout: 20000 }, (error, stdout, stderr) => {
      resolve({
        code: error ? 1 : 0,
        stdout: String(stdout || ''),
        stderr: String(stderr || '').slice(0, 300)
      });
    });
  });
}

export async function githubAccountStatus() {
  const result = await execGh(['api', 'user', '--jq', '{login:.login,email:.email}']);
  if (result.code !== 0) return { connected: false, login: '', email: '' };
  try {
    const data = JSON.parse(result.stdout);
    const login = String(data.login || '');
    const email = data.email ? String(data.email) : '';
    return { connected: Boolean(login), login, email };
  } catch {
    return { connected: false, login: '', email: '' };
  }
}

export async function githubAccountToken() {
  const result = await execGh(['auth', 'token']);
  if (result.code !== 0) return '';
  return result.stdout.trim();
}

export function startGithubWebLogin() {
  const args = ['auth', 'login', '--hostname', 'github.com', '--git-protocol', 'https', '--web', '--skip-ssh-key'];
  if (spawnOverride) {
    spawnOverride(args);
    return { pending: true };
  }
  const child = spawn('gh', args, { detached: true, stdio: 'ignore', windowsHide: false });
  child.unref();
  return { pending: true };
}
