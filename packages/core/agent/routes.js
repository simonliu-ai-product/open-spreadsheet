/**
 * The agent panel's endpoints on the dev server.
 *
 * GET    /__agent/status                 → AgentStatus
 * POST   /__agent/install                { provider } → { ok }
 * GET    /__agent/sessions               → { sessions: AgentSessionSummary[] }
 * POST   /__agent/sessions               { provider, workbook? } → AgentSession
 * GET    /__agent/sessions/:id           → AgentSession
 * DELETE /__agent/sessions/:id           → { ok }
 * POST   /__agent/sessions/:id/messages  { text, context?, autoApprove?, model?, effort?, edit? } → { ok }
 * POST   /__agent/sessions/:id/stop      → { ok }
 * POST   /__agent/sessions/:id/respond   { itemId, answers? | allow? | cancel?, context?, autoApprove?, model?, effort? } → { ok }
 * GET    /__agent/events                 → server-sent turn progress
 *
 * The agent runs commands in the workspace, so only this machine may drive it,
 * even when the dev server listens on the network.
 *
 * @typedef {import('./types.d.ts').AgentTurnContext} AgentTurnContext
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { ID } from '../ops.js';
import { changedWorkbooks, checkItemResult, checkWorkbooks, shouldCheckAfterTurn, snapshotWorkbooks } from './check-after-turn.js';
import { detectPackageManager, installSdk } from './install.js';
import { createProviders } from './providers/index.js';
import { createAgentRunner } from './runner.js';
import { agentStatus } from './status.js';
import { createSessionStore, newId, sessionsFile } from './store.js';
import { isAgentEffort, isAgentProviderId } from './types.js';

const CORE = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const BODY_LIMIT = 1024 * 1024;
const MODEL_RE = /^[\w.:[\]-]{1,80}$/;

// ---------- request parsing (exported for tests) ----------

const pick = (v, max) => (typeof v === 'string' && v ? v.slice(0, max) : undefined);

/** @returns {AgentTurnContext | undefined} */
export function parseContext(raw) {
  if (!raw || typeof raw !== 'object' || typeof raw.workbook !== 'string' || !ID.test(raw.workbook)) return undefined;
  /** @type {AgentTurnContext} */
  const ctx = { workbook: raw.workbook };
  const fields = { title: pick(raw.title, 200), sheet: pick(raw.sheet, 64), input: pick(raw.input, 2000), note: pick(raw.note, 4000) };
  // Addresses are A1 text only: they go into the prompt bare.
  const a1 = (v) => (typeof v === 'string' && /^\$?[A-Z]{1,3}\$?\d{1,7}(:\$?[A-Z]{1,3}\$?\d{1,7})?$/i.test(v) ? v.toUpperCase() : undefined);
  Object.assign(ctx, Object.fromEntries(Object.entries({ ...fields, active: a1(raw.active), selection: a1(raw.selection) }).filter(([, v]) => v !== undefined)));
  return ctx;
}

export const parseModel = (raw) => (typeof raw === 'string' && MODEL_RE.test(raw) ? raw : null);

/** @returns {import('./types.d.ts').RespondPayload | null} */
export function parseRespond(raw) {
  if (!raw || typeof raw !== 'object') return null;
  if (raw.cancel === true) return { cancel: true };
  if (typeof raw.allow === 'boolean') return { allow: raw.allow };
  if (raw.answers && typeof raw.answers === 'object' && !Array.isArray(raw.answers)) {
    return { answers: Object.fromEntries(Object.entries(raw.answers).filter(([, v]) => typeof v === 'string')) };
  }
  return null;
}

/** @returns {import('./runner.js').TurnSettings} */
export function parseTurnSettings(raw) {
  const settings = {};
  const context = parseContext(raw.context);
  if (context) settings.context = context;
  if (typeof raw.autoApprove === 'boolean') settings.autoApprove = raw.autoApprove;
  if ('model' in raw) settings.model = parseModel(raw.model);
  if ('effort' in raw) settings.effort = isAgentEffort(raw.effort) ? raw.effort : null;
  return settings;
}

/** Same machine only: the agent can run commands here. */
export function isLoopback(req) {
  const address = req.socket?.remoteAddress ?? '';
  return address === '::1' || address.startsWith('127.') || address.startsWith('::ffff:127.');
}

/** `"agent": false` in open-spreadsheet.config.json turns the panel off. */
export function agentDisabledByConfig(root) {
  try {
    return JSON.parse(readFileSync(path.join(root, 'open-spreadsheet.config.json'), 'utf8'))?.agent === false;
  } catch {
    return false;
  }
}

class HttpError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

const readJson = (req) =>
  new Promise((resolve, reject) => {
    let data = '';
    req.on('data', (c) => {
      data += c;
      if (data.length > BODY_LIMIT) reject(new HttpError(413, 'request too large'));
    });
    req.on('end', () => {
      try {
        const body = data ? JSON.parse(data) : {};
        resolve(body && typeof body === 'object' ? body : {});
      } catch {
        reject(new HttpError(400, 'invalid JSON'));
      }
    });
    req.on('error', reject);
  });

const json = (res, status, body) => {
  res.statusCode = status;
  res.setHeader('content-type', 'application/json; charset=utf-8');
  res.setHeader('cache-control', 'no-store');
  res.end(JSON.stringify(body));
};

// ---------- routes ----------

/**
 * @param {import('vite').ViteDevServer} server
 * @param {{ root: string, allowedWrite: (req: import('node:http').IncomingMessage) => boolean, providers?: Record<string, import('./provider.js').AgentProvider> }} opts
 */
export function registerAgentRoutes(server, opts) {
  const { root, allowedWrite } = opts;
  const providers = opts.providers ?? createProviders({ mcpServer: { command: process.execPath, args: [path.join(CORE, 'mcp.js'), '--cwd', root] } });
  const store = createSessionStore({ file: sessionsFile(root) });
  const checkCommand = `"${process.execPath}" "${path.join(CORE, 'bin.js')}" check`;

  // Agent output goes over its own stream rather than the HMR socket, which
  // every network client can subscribe to when the server runs with --host.
  const streams = new Set();
  const broadcast = (event) => {
    const frame = `data: ${JSON.stringify(event)}\n\n`;
    for (const s of streams) s.write(frame);
  };
  const heartbeat = setInterval(() => {
    for (const s of streams) s.write(': ping\n\n');
  }, 25_000);
  heartbeat.unref();

  /** Workbook hashes when each running turn started, to tell what it changed. */
  const snapshots = new Map();

  const checkAfterTurn = (turn) => {
    const before = snapshots.get(turn.sessionId);
    snapshots.delete(turn.sessionId);
    const session = store.get(turn.sessionId);
    if (!session || !before) return;
    const changed = changedWorkbooks(before, snapshotWorkbooks(root));
    if (!shouldCheckAfterTurn({ session, userItemId: turn.userItemId, changed })) return;
    const item = store.appendItem(session.id, { kind: 'check', id: newId('c'), ts: new Date().toISOString(), workbooks: changed, status: 'running', findings: [], total: 0, error: null });
    try {
      store.updateItem(session.id, item.id, 'check', { status: 'done', ...checkItemResult(checkWorkbooks(root, changed)) });
    } catch (err) {
      store.updateItem(session.id, item.id, 'check', { status: 'failed', error: err instanceof Error ? err.message : String(err) });
    }
  };

  const runner = createAgentRunner({
    cwd: root,
    writableRoots: [path.join(root, 'sheets'), path.join(root, 'assets')],
    store,
    providers,
    promptOptions: () => ({ checkCommand, mcp: true }),
    onTurnStart: (turn) => snapshots.set(turn.sessionId, snapshotWorkbooks(root)),
    onTurnEnd: checkAfterTurn,
  });

  store.subscribe(broadcast);
  server.httpServer?.on('close', () => {
    clearInterval(heartbeat);
    for (const s of streams) s.end();
    runner.stopAll();
    for (const p of Object.values(providers)) p.dispose?.();
    void store.flush();
  });

  const installs = new Map();

  /** A write from a page on this server, sent as JSON when it has a body. */
  const guard = (req, { body = true } = {}) => {
    if (String(req.headers['sec-fetch-site'] ?? '').toLowerCase() === 'cross-site' || !allowedWrite(req)) throw new HttpError(403, 'forbidden');
    if (body && !String(req.headers['content-type'] ?? '').toLowerCase().startsWith('application/json')) throw new HttpError(415, 'content-type must be application/json');
  };

  server.middlewares.use('/__agent', async (req, res, next) => {
    const url = new URL(req.url ?? '/', 'http://x');
    const method = req.method ?? 'GET';
    const segments = url.pathname.split('/').filter(Boolean);
    const offReason = agentDisabledByConfig(root) ? 'config' : isLoopback(req) ? null : 'remote';
    try {
      if (method === 'GET' && url.pathname === '/status') {
        if (offReason) return json(res, 200, { enabled: false, reason: offReason, packageManager: '', providers: [] });
        return json(res, 200, await agentStatus(root, providers, detectPackageManager(root)));
      }
      if (offReason) return json(res, 404, { error: 'the agent panel is off' });

      if (method === 'GET' && url.pathname === '/events') {
        res.writeHead(200, { 'content-type': 'text/event-stream', 'cache-control': 'no-store', connection: 'keep-alive' });
        res.write(': connected\n\n');
        streams.add(res);
        const drop = () => streams.delete(res);
        res.on('close', drop);
        res.on('error', drop);
        return;
      }

      if (method === 'POST' && url.pathname === '/install') {
        guard(req);
        const body = await readJson(req);
        if (!isAgentProviderId(body.provider)) return json(res, 400, { error: 'invalid provider' });
        const pkg = providers[body.provider].sdkPackage;
        let install = installs.get(pkg);
        if (!install) {
          install = installSdk(detectPackageManager(root), pkg, root).finally(() => installs.delete(pkg));
          installs.set(pkg, install);
        }
        await install;
        return json(res, 200, { ok: true });
      }

      if (method === 'GET' && url.pathname === '/sessions') return json(res, 200, { sessions: store.list() });

      if (method === 'POST' && url.pathname === '/sessions') {
        guard(req);
        const body = await readJson(req);
        if (!isAgentProviderId(body.provider)) return json(res, 400, { error: 'invalid provider' });
        return json(res, 200, store.create(body.provider, typeof body.workbook === 'string' && ID.test(body.workbook) ? body.workbook : null));
      }

      if (segments[0] === 'sessions' && segments[1]) {
        const id = segments[1];
        const action = segments[2];
        const session = store.get(id);
        if (!session) return json(res, 404, { error: 'session not found' });

        if (method === 'GET' && !action) return json(res, 200, session);

        if (method === 'DELETE' && !action) {
          guard(req, { body: false });
          runner.stop(id);
          store.remove(id);
          return json(res, 200, { ok: true });
        }

        if (method === 'POST' && action === 'messages') {
          guard(req);
          const body = await readJson(req);
          if (typeof body.text !== 'string' || !body.text.trim()) return json(res, 400, { error: 'missing text' });
          if (runner.isRunning(id)) return json(res, 409, { error: 'session is already running' });
          if (body.edit !== undefined && !session.items.some((i) => i.id === body.edit && i.kind === 'user')) return json(res, 400, { error: 'unknown message' });
          await runner.send(id, {
            text: body.text,
            context: parseContext(body.context),
            autoApprove: body.autoApprove === true,
            model: parseModel(body.model),
            effort: isAgentEffort(body.effort) ? body.effort : null,
            edit: typeof body.edit === 'string' ? body.edit : undefined,
          });
          return json(res, 202, { ok: true });
        }

        if (method === 'POST' && action === 'stop') {
          guard(req, { body: false });
          return json(res, 200, { ok: runner.stop(id) });
        }

        if (method === 'POST' && action === 'respond') {
          guard(req);
          const body = await readJson(req);
          const payload = parseRespond(body);
          if (typeof body.itemId !== 'string' || !payload) return json(res, 400, { error: 'invalid response' });
          const ok = runner.respond(id, body.itemId, payload, parseTurnSettings(body));
          return json(res, ok ? 200 : 409, ok ? { ok } : { error: 'nothing pending' });
        }
      }
      next();
    } catch (err) {
      json(res, err instanceof HttpError ? err.status : 500, { error: String(err?.message ?? err) });
    }
  });
}
