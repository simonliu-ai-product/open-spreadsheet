/**
 * The panel's conversations, in memory and saved (debounced) to
 * node_modules/.open-spreadsheet/agent-sessions.json: personal, never tracked.
 * Every change is also an event, which the dev server streams to the panel.
 *
 * @typedef {import('./types.d.ts').AgentSession} AgentSession
 * @typedef {import('./types.d.ts').AgentItem} AgentItem
 * @typedef {import('./types.d.ts').AgentStreamEvent} AgentStreamEvent
 */
import { randomBytes } from 'node:crypto';
import { existsSync, readFileSync } from 'node:fs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { isAgentProviderId, sessionSummary } from './types.js';

const MAX_SESSIONS = 30;
const MAX_ITEMS_PER_SESSION = 400;
const SAVE_DEBOUNCE_MS = 150;
const TITLE_MAX = 60;

export const sessionsFile = (root) => path.join(root, 'node_modules', '.open-spreadsheet', 'agent-sessions.json');

export const newId = (prefix) => `${prefix}-${randomBytes(6).toString('hex')}`;

/** A conversation's name: the first line of its first message. */
export function titleFromPrompt(text) {
  const line = text
    .split('\n')
    .map((l) => l.trim())
    .find((l) => l.length > 0);
  if (!line) return 'Untitled';
  return line.length > TITLE_MAX ? `${line.slice(0, TITLE_MAX - 1)}…` : line;
}

/** @returns {AgentItem} */
function closePending(item) {
  if (item.kind === 'assistant' && item.streaming) return { ...item, streaming: false };
  if (item.kind === 'tool' && item.status === 'running') return { ...item, status: 'error' };
  if (item.kind === 'question' && item.status === 'pending') return { ...item, status: 'cancelled' };
  if (item.kind === 'permission' && item.status === 'pending') return { ...item, status: 'denied' };
  if (item.kind === 'check' && item.status === 'running') return { ...item, status: 'failed', error: 'interrupted' };
  return item;
}

/** @returns {AgentSession[]} */
function readPersisted(file) {
  if (!existsSync(file)) return [];
  try {
    const parsed = JSON.parse(readFileSync(file, 'utf8'));
    if (parsed?.version !== 1 || !Array.isArray(parsed.sessions)) return [];
    // A dev-server restart mid-turn leaves no process to finish it: every
    // reloaded session starts idle and its pending prompts are closed out.
    return parsed.sessions.filter((s) => isAgentProviderId(s?.provider) && Array.isArray(s.items)).map((s) => ({ ...s, status: 'idle', items: s.items.map(closePending) }));
  } catch {
    return [];
  }
}

/** @param {{ file: string | null }} opts  null keeps everything in memory (tests). */
export function createSessionStore(opts) {
  /** @type {Map<string, AgentSession>} */
  const sessions = new Map();
  /** @type {Set<(event: AgentStreamEvent) => void>} */
  const listeners = new Set();
  let saveTimer = null;
  // Touch order breaks ties between sessions updated in the same millisecond.
  let seq = 0;
  const order = new Map();

  if (opts.file) {
    for (const s of readPersisted(opts.file)) {
      sessions.set(s.id, s);
      order.set(s.id, seq++);
    }
  }

  const emit = (event) => {
    for (const l of listeners) l(event);
  };

  const writeSnapshot = async (file) => {
    const body = { version: 1, sessions: list().flatMap((s) => sessions.get(s.id) ?? []) };
    const tmp = `${file}.tmp`;
    try {
      await fs.mkdir(path.dirname(file), { recursive: true });
      await fs.writeFile(tmp, JSON.stringify(body), { encoding: 'utf8', mode: 0o600 });
      await fs.rename(tmp, file);
    } catch (err) {
      console.error('[open-spreadsheet] could not save agent sessions:', err);
      await fs.rm(tmp, { force: true }).catch(() => {});
    }
  };

  // Saves share one temp file, so a flush racing a debounced save waits for it.
  let saving = Promise.resolve();
  const save = () => {
    if (opts.file) saving = saving.then(() => writeSnapshot(opts.file));
    return saving;
  };
  const scheduleSave = () => {
    if (!opts.file) return;
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      saveTimer = null;
      void save();
    }, SAVE_DEBOUNCE_MS);
  };

  const touch = (session) => {
    session.updatedAt = new Date().toISOString();
    order.set(session.id, seq++);
    emit({ type: 'session', session: sessionSummary(session) });
    scheduleSave();
  };

  const need = (id) => {
    const session = sessions.get(id);
    if (!session) throw new Error(`unknown agent session: ${id}`);
    return session;
  };

  function list() {
    return Array.from(sessions.values())
      .sort((a, b) => (a.updatedAt !== b.updatedAt ? (a.updatedAt < b.updatedAt ? 1 : -1) : (order.get(b.id) ?? 0) - (order.get(a.id) ?? 0)))
      .map(sessionSummary);
  }

  const get = (id) => sessions.get(id) ?? null;

  function create(provider, workbook) {
    const now = new Date().toISOString();
    /** @type {AgentSession} */
    const session = { id: newId('s'), provider, title: 'New chat', createdAt: now, updatedAt: now, status: 'idle', workbook: workbook ?? null, providerSessionId: null, items: [] };
    sessions.set(session.id, session);
    order.set(session.id, seq++);
    for (const stale of list().filter((s) => s.status === 'idle').slice(MAX_SESSIONS)) {
      sessions.delete(stale.id);
      order.delete(stale.id);
      emit({ type: 'session-removed', sessionId: stale.id });
    }
    touch(session);
    return session;
  }

  function remove(id) {
    if (!sessions.delete(id)) return false;
    order.delete(id);
    emit({ type: 'session-removed', sessionId: id });
    scheduleSave();
    return true;
  }

  /** @param {Partial<Pick<AgentSession, 'title' | 'status' | 'providerSessionId' | 'workbook'>>} changes */
  function patch(id, changes) {
    const session = need(id);
    Object.assign(session, changes);
    touch(session);
    return session;
  }

  /** @template {AgentItem} T @param {T} item @returns {T} */
  function appendItem(sessionId, item) {
    const session = need(sessionId);
    session.items.push(item);
    if (session.items.length > MAX_ITEMS_PER_SESSION) {
      session.items.splice(0, session.items.length - MAX_ITEMS_PER_SESSION);
      session.trimmed = true;
    }
    emit({ type: 'item', sessionId, item });
    touch(session);
    return item;
  }

  function updateItem(sessionId, itemId, kind, changes) {
    const session = need(sessionId);
    const index = session.items.findIndex((i) => i.id === itemId && i.kind === kind);
    if (index === -1) return null;
    const next = { ...session.items[index], ...changes };
    session.items[index] = next;
    emit({ type: 'item', sessionId, item: next });
    touch(session);
    return next;
  }

  /** Drops an item and everything after it. */
  function truncateFrom(sessionId, itemId) {
    const session = need(sessionId);
    const index = session.items.findIndex((i) => i.id === itemId);
    if (index === -1) return false;
    session.items.splice(index);
    emit({ type: 'truncate', sessionId, itemId });
    touch(session);
    return true;
  }

  function appendDelta(sessionId, itemId, text) {
    if (!text) return;
    const item = need(sessionId).items.find((i) => i.id === itemId);
    if (item?.kind !== 'assistant') return;
    item.text += text;
    emit({ type: 'delta', sessionId, itemId, text });
    scheduleSave();
  }

  function subscribe(listener) {
    listeners.add(listener);
    return () => listeners.delete(listener);
  }

  async function flush() {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = null;
    await save();
  }

  return { list, get, create, remove, patch, appendItem, updateItem, truncateFrom, appendDelta, subscribe, flush };
}
