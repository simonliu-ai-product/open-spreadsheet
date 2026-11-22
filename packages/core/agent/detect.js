/**
 * Finding an agent on this machine: its CLI on PATH and its SDK in the
 * workspace's node_modules. The SDKs are never dependencies of core; the
 * person installs one into their workspace from the panel.
 */
import { spawn } from 'node:child_process';
import { existsSync, readFileSync, statSync } from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const IS_WINDOWS = process.platform === 'win32';
const VERSION_TIMEOUT_MS = 5000;

export function findOnPath(command, env = process.env) {
  const dirs = (env.PATH ?? env.Path ?? '').split(path.delimiter).filter(Boolean);
  const exts = IS_WINDOWS ? (env.PATHEXT ?? '.EXE;.CMD;.BAT;.COM').split(';') : [''];
  for (const dir of dirs) {
    for (const ext of exts) {
      const candidate = path.join(dir, command + ext.toLowerCase());
      try {
        if (statSync(candidate, { throwIfNoEntry: false })?.isFile()) return candidate;
      } catch {
        // unreadable: keep looking
      }
    }
  }
  return null;
}

/** The SDKs spawn the CLI without a shell, which cannot run a Windows .cmd shim, so only a real executable overrides their bundled one. */
export function findSpawnableCli(command, env = process.env) {
  const found = findOnPath(command, env);
  if (!found || !IS_WINDOWS) return found;
  return found.toLowerCase().endsWith('.exe') ? found : null;
}

const versionCache = new Map();

/** `<cli> --version`, once per path. */
export function commandVersion(file) {
  let cached = versionCache.get(file);
  if (!cached) {
    cached = readVersion(file);
    versionCache.set(file, cached);
  }
  return cached;
}

function readVersion(file) {
  return new Promise((resolve) => {
    let out = '';
    let child;
    try {
      child = spawn(file, ['--version'], { stdio: ['ignore', 'pipe', 'pipe'], shell: IS_WINDOWS, windowsHide: true });
    } catch {
      resolve(null);
      return;
    }
    const timer = setTimeout(() => {
      child.kill();
      resolve(null);
    }, VERSION_TIMEOUT_MS);
    child.stdout?.on('data', (c) => (out += c));
    child.stderr?.on('data', (c) => (out += c));
    child.on('error', () => {
      clearTimeout(timer);
      resolve(null);
    });
    child.on('close', () => {
      clearTimeout(timer);
      resolve(parseVersion(out));
    });
  });
}

export function parseVersion(output) {
  const m = /(\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?)/.exec(output);
  return m ? m[1] : null;
}

/** The file `import` would load from a package.json: its root export (import, node, default, require), else module or main. */
export function entryFromPackageJson(pkg) {
  return pickExport(pkg.exports) ?? pkg.module ?? pkg.main ?? null;
}

function pickExport(field) {
  if (!field) return null;
  if (typeof field === 'string') return field;
  if (typeof field !== 'object') return null;
  const isSubpathMap = Object.keys(field).some((k) => k.startsWith('.'));
  return pickCondition(isSubpathMap ? field['.'] : field);
}

function pickCondition(target) {
  if (!target) return null;
  if (typeof target === 'string') return target;
  if (Array.isArray(target)) {
    for (const t of target) {
      const hit = pickCondition(t);
      if (hit) return hit;
    }
    return null;
  }
  if (typeof target !== 'object') return null;
  for (const key of ['import', 'node', 'default', 'require']) {
    if (key in target) {
      const hit = pickCondition(target[key]);
      if (hit) return hit;
    }
  }
  return null;
}

/**
 * Walks node_modules up from the workspace as Node does, but reads
 * package.json directly: the SDKs' `exports` maps don't expose it.
 * @returns {{ dir: string, version: string, entry: string } | null}
 */
export function resolveWorkspacePackage(cwd, name) {
  for (let dir = path.resolve(cwd); ; dir = path.dirname(dir)) {
    const file = path.join(dir, 'node_modules', name, 'package.json');
    if (existsSync(file)) {
      let pkg;
      try {
        pkg = JSON.parse(readFileSync(file, 'utf8'));
      } catch {
        return null;
      }
      const entry = entryFromPackageJson(pkg);
      if (!entry) return null;
      const pkgDir = path.dirname(file);
      return { dir: pkgDir, version: pkg.version ?? '0.0.0', entry: path.resolve(pkgDir, entry) };
    }
    if (path.dirname(dir) === dir) return null;
  }
}

export async function importWorkspacePackage(cwd, name) {
  const resolved = resolveWorkspacePackage(cwd, name);
  if (!resolved) return null;
  return import(pathToFileURL(resolved.entry).href);
}
