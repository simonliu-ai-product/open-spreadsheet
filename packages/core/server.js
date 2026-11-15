import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, realpathSync, renameSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { currentFile, hashOf, ID, listWorkbooks, OpsError, readWorkbook, trashWorkbook, workbookFile, writeWorkbook } from './ops.js';

const CORE = path.dirname(fileURLToPath(import.meta.url));
// Wherever the package manager put the engine (beside core in this repo, under node_modules once installed); Vite may serve its files.
const ENGINE = realpathSync(path.dirname(fileURLToPath(import.meta.resolve('@open-spreadsheet/engine'))));
export const VERSION = JSON.parse(readFileSync(path.join(CORE, 'package.json'), 'utf8')).version;

/**
 * A workspace is a folder with open-spreadsheet.config.json and a sheets/
 * folder; each workbook is sheets/<id>/workbook.json, the cells exactly as
 * typed, so git can diff it and an agent can edit it.
 */
const blankWorkbook = (title) => ({ title, sheets: [{ name: 'Sheet1', cells: {} }] });

/** Copies the bundled agent skills into the workspace, for Claude Code (.claude) and other agents (.agents). */
export function syncSkills(root) {
  const from = path.join(CORE, 'skills');
  const names = readdirSync(from).filter((n) => existsSync(path.join(from, n, 'SKILL.md')));
  for (const base of ['.claude/skills', '.agents/skills']) {
    for (const n of names) {
      const to = path.join(root, base, n);
      mkdirSync(to, { recursive: true });
      cpSync(path.join(from, n), to, { recursive: true });
    }
  }
  return names;
}

export function initWorkspace(dir) {
  const root = path.resolve(dir);
  if (existsSync(root) && readdirSync(root).length > 0) throw new Error(`${root} is not empty`);
  mkdirSync(path.join(root, 'sheets', 'getting-started'), { recursive: true });
  const name = path.basename(root);
  writeFileSync(
    path.join(root, 'package.json'),
    `${JSON.stringify(
      {
        name,
        private: true,
        type: 'module',
        scripts: { dev: 'open-spreadsheet dev', check: 'open-spreadsheet check' },
        dependencies: { '@open-spreadsheet/core': `^${VERSION}` },
      },
      null,
      2,
    )}\n`,
  );
  writeFileSync(path.join(root, 'open-spreadsheet.config.json'), '{}\n');
  writeFileSync(path.join(root, '.gitignore'), 'node_modules\n.DS_Store\n# workbooks moved to the trash from the viewer or an agent\nsheets/.trash\n');
  const agents = readFileSync(path.join(CORE, 'templates', 'AGENTS.md'), 'utf8');
  writeFileSync(path.join(root, 'AGENTS.md'), agents);
  writeFileSync(path.join(root, 'CLAUDE.md'), agents);
  writeFileSync(
    path.join(root, '.mcp.json'),
    `${JSON.stringify({ mcpServers: { 'open-spreadsheet': { command: 'npx', args: ['open-spreadsheet-mcp'] } } }, null, 2)}\n`,
  );
  writeFileSync(
    workbookFile(root, 'getting-started'),
    `${JSON.stringify(
      {
        title: 'Getting started',
        sheets: [
          {
            name: 'Sheet1',
            cells: { A1: 'Item', B1: 'Amount', A2: 'Coffee', B2: '120', A3: 'Lunch', B3: '180', A4: 'Total', B4: '=SUM(B2:B3)' },
            styles: { A1: { b: true }, B1: { b: true }, A4: { b: true }, B4: { b: true } },
          },
        ],
      },
      null,
      2,
    )}\n`,
  );
  syncSkills(root);
  return root;
}

const json = (res, status, body, headers = {}) => {
  res.statusCode = status;
  res.setHeader('content-type', 'application/json; charset=utf-8');
  for (const [k, v] of Object.entries(headers)) res.setHeader(k, v);
  res.end(JSON.stringify(body));
};

const readBody = (req) =>
  new Promise((resolve, reject) => {
    let data = '';
    req.on('data', (c) => {
      data += c;
      if (data.length > 50_000_000) reject(new Error('too large'));
    });
    req.on('end', () => resolve(data));
    req.on('error', reject);
  });

/**
 * Writes come only from this machine's own pages. A proxy in front (a tunnel,
 * a cloud IDE) may forward the page from another host; then the Origin must
 * match the Host the proxy forwarded, scheme or not.
 */
function allowedWrite(req) {
  const origin = req.headers.origin;
  if (!origin) return true;
  try {
    const o = new URL(origin);
    if (['localhost', '127.0.0.1', '[::1]'].includes(o.hostname)) return true;
    const host = String(req.headers['x-forwarded-host'] ?? req.headers.host ?? '').split(',')[0].trim();
    return Boolean(host) && o.host === host;
  } catch {
    return false;
  }
}

const STATUS = { bad_id: 400, bad_data: 400, bad_json: 422, not_found: 404, exists: 409, conflict: 412 };

function apiPlugin(root) {
  // Bodies this server just wrote, so its own write does not come back as an outside change.
  const ownWrites = new Map();
  return {
    name: 'open-spreadsheet:api',
    configureServer(server) {
      server.watcher.add(path.join(root, 'sheets'));
      server.watcher.on('change', (file) => {
        const m = /sheets[\\/]([^\\/]+)[\\/]workbook\.json$/.exec(file);
        if (!m) return;
        let body = '';
        try {
          body = readFileSync(file, 'utf8');
        } catch {
          return;
        }
        if (ownWrites.get(m[1]) === body) return;
        server.ws.send({ type: 'custom', event: 'open-spreadsheet:changed', data: { id: m[1], hash: hashOf(body) } });
      });

      // What the person is looking at, for agents: node_modules/.open-spreadsheet/current.json.
      let writing = Promise.resolve();
      server.ws.on('open-spreadsheet:current', (raw) => {
        if (!raw || typeof raw.workbook !== 'string' || !ID.test(raw.workbook)) return;
        const pick = (v, max = 200) => (typeof v === 'string' ? v.slice(0, max) : undefined);
        const current = {
          workbook: raw.workbook,
          title: pick(raw.title),
          sheet: pick(raw.sheet, 64),
          active: pick(raw.active, 16),
          selection: pick(raw.selection, 32),
          input: pick(raw.input, 2000),
          value: typeof raw.value === 'number' || typeof raw.value === 'boolean' ? raw.value : pick(raw.value, 2000),
          note: pick(raw.note, 4000),
          url: `/#/w/${raw.workbook}`,
          file: `sheets/${raw.workbook}/workbook.json`,
          updatedAt: new Date().toISOString(),
        };
        const file = currentFile(root);
        writing = writing.then(() => {
          try {
            mkdirSync(path.dirname(file), { recursive: true });
            writeFileSync(`${file}.tmp`, `${JSON.stringify(current, null, 2)}\n`);
            renameSync(`${file}.tmp`, file);
          } catch {
            // Best effort: an agent can still ask.
          }
        });
      });

      server.middlewares.use(async (req, res, next) => {
        const url = new URL(req.url ?? '/', 'http://x');
        if (!url.pathname.startsWith('/__api/')) return next();
        try {
          if (url.pathname === '/__api/workbooks' && req.method === 'GET') {
            return json(res, 200, { name: path.basename(root), version: VERSION, workbooks: listWorkbooks(root) });
          }
          if (url.pathname === '/__api/workbooks' && req.method === 'POST') {
            if (!allowedWrite(req)) return json(res, 403, { error: 'forbidden' });
            // { id, title }: a blank workbook; { data }: an imported or templated one; { from }: a copy.
            const { id, title, data, from } = JSON.parse(await readBody(req));
            if (!ID.test(id ?? '')) return json(res, 400, { error: 'Use lowercase letters, numbers and dashes.' });
            if (existsSync(workbookFile(root, id))) return json(res, 409, { error: 'That name is taken.' });
            let book = blankWorkbook(title || id);
            if (from !== undefined) book = { ...readWorkbook(root, from).data, ...(title ? { title } : {}) };
            else if (data !== undefined) book = { ...data, title: title || data.title || id };
            const { text } = await writeWorkbook(root, id, book, { create: true });
            ownWrites.set(id, text);
            return json(res, 201, { id });
          }
          const m = /^\/__api\/workbooks\/([^/]+)$/.exec(url.pathname);
          if (m && ID.test(m[1])) {
            const id = m[1];
            if (req.method === 'GET') {
              const { data, hash } = readWorkbook(root, id);
              return json(res, 200, data, { etag: `"${hash}"` });
            }
            if (req.method === 'PUT') {
              if (!allowedWrite(req)) return json(res, 403, { error: 'forbidden' });
              const data = JSON.parse(await readBody(req));
              const ifMatch = String(req.headers['if-match'] ?? '').replaceAll('"', '') || undefined;
              const { hash, text } = await writeWorkbook(root, id, data, { expectedHash: ifMatch });
              ownWrites.set(id, text);
              return json(res, 200, { ok: true }, { etag: `"${hash}"` });
            }
            if (req.method === 'DELETE') {
              if (!allowedWrite(req)) return json(res, 403, { error: 'forbidden' });
              trashWorkbook(root, id);
              return json(res, 200, { ok: true });
            }
          }
          return json(res, 404, { error: 'not found' });
        } catch (e) {
          if (e instanceof OpsError) return json(res, STATUS[e.code] ?? 400, { error: e.message, code: e.code });
          return json(res, 500, { error: String(e?.message ?? e) });
        }
      });
    },
  };
}

export async function startDev({ root, port, open, host }) {
  if (!existsSync(path.join(root, 'open-spreadsheet.config.json'))) {
    process.stderr.write(`No open-spreadsheet.config.json in ${root}. Run open-spreadsheet init <dir> first.\n`);
    process.exit(1);
  }
  const { createServer } = await import('vite');
  const server = await createServer({
    configFile: false,
    root: path.join(CORE, 'app'),
    // Vite's cache lives in the workspace, not next to the framework, so read-only installs work.
    cacheDir: path.join(root, 'node_modules', '.vite', 'open-spreadsheet'),
    logLevel: 'info',
    plugins: [apiPlugin(root)],
    server: { port, strictPort: true, open, host, fs: { allow: [CORE, ENGINE, root] } },
  });
  await server.listen();
  server.printUrls();
  for (const signal of ['SIGINT', 'SIGTERM']) process.on(signal, () => server.close().then(() => process.exit(0)));
}
