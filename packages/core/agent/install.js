/** Installing an agent's SDK into the workspace with the workspace's own package manager. */
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import path from 'node:path';

const LOCKFILES = [
  ['pnpm-lock.yaml', 'pnpm'],
  ['yarn.lock', 'yarn'],
  ['bun.lockb', 'bun'],
  ['bun.lock', 'bun'],
  ['package-lock.json', 'npm'],
];

const INSTALL_TIMEOUT_MS = 10 * 60 * 1000;

/** @returns {'npm' | 'pnpm' | 'yarn' | 'bun'} */
export function detectPackageManager(cwd, env = process.env) {
  const ua = env.npm_config_user_agent ?? '';
  for (const pm of ['pnpm', 'yarn', 'bun', 'npm']) if (ua.startsWith(pm)) return pm;
  // Workspace members have no lockfile of their own; it lives at the root.
  for (let dir = path.resolve(cwd); ; dir = path.dirname(dir)) {
    for (const [lockfile, pm] of LOCKFILES) if (existsSync(path.join(dir, lockfile))) return pm;
    if (path.dirname(dir) === dir) return 'npm';
  }
}

export function installCommand(packageManager, pkg) {
  switch (packageManager) {
    case 'pnpm':
      return { cmd: 'pnpm', args: ['add', '-D', pkg] };
    case 'yarn':
      return { cmd: 'yarn', args: ['add', '-D', pkg] };
    case 'bun':
      return { cmd: 'bun', args: ['add', '-d', pkg] };
    default:
      return { cmd: 'npm', args: ['install', '-D', pkg] };
  }
}

export function runCommand({ cmd, args }, cwd, timeoutMs = INSTALL_TIMEOUT_MS) {
  const label = [cmd, ...args].join(' ');
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, { cwd, env: process.env, shell: process.platform === 'win32', stdio: ['ignore', 'ignore', 'pipe'], windowsHide: true });
    let stderr = '';
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error(`${label} timed out`));
    }, timeoutMs);
    child.stderr.on('data', (c) => (stderr = (stderr + c).slice(-2000)));
    child.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });
    child.on('close', (code) => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(new Error(stderr.trim() || `${label} exited with code ${code ?? 'unknown'}`));
    });
  });
}

export const installSdk = (packageManager, pkg, cwd) => runCommand(installCommand(packageManager, pkg), cwd);
