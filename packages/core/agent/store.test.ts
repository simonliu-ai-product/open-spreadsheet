import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
// @ts-expect-error plain JS module
import { createSessionStore, sessionsFile, titleFromPrompt } from './store.js';

const ws = () => mkdtempSync(path.join(tmpdir(), 'oss-agent-store-'));

describe('agent session store', () => {
  it('keeps sessions under node_modules/.open-spreadsheet, never in tracked files', () => {
    expect(sessionsFile('/ws')).toBe(path.join('/ws', 'node_modules', '.open-spreadsheet', 'agent-sessions.json'));
  });

  it('streams every change as an event', () => {
    const store = createSessionStore({ file: null });
    const events: any[] = [];
    store.subscribe((e: any) => events.push(e));
    const s = store.create('codex', 'budget');
    store.appendItem(s.id, { kind: 'assistant', id: 'a', ts: '', text: '', streaming: true });
    store.appendDelta(s.id, 'a', 'Hi');
    store.updateItem(s.id, 'a', 'assistant', { streaming: false });
    store.truncateFrom(s.id, 'a');
    store.remove(s.id);
    expect(events.map((e) => e.type)).toEqual(['session', 'item', 'session', 'delta', 'item', 'session', 'truncate', 'session', 'session-removed']);
    expect(events[0].session).not.toHaveProperty('items');
    expect(events[0].session.workbook).toBe('budget');
  });

  it('lists the newest first and keeps at most 30 idle conversations', () => {
    const store = createSessionStore({ file: null });
    const ids = Array.from({ length: 32 }, () => store.create('codex', null).id);
    const list = store.list();
    expect(list).toHaveLength(30);
    expect(list[0].id).toBe(ids[31]);
    expect(store.get(ids[0])).toBeNull();
  });

  it('caps a transcript and marks it trimmed', () => {
    const store = createSessionStore({ file: null });
    const s = store.create('codex', null);
    for (let n = 0; n < 405; n++) store.appendItem(s.id, { kind: 'error', id: `e${n}`, ts: '', text: String(n) });
    expect(store.get(s.id).items).toHaveLength(400);
    expect(store.get(s.id).items[0].id).toBe('e5');
    expect(store.get(s.id).trimmed).toBe(true);
  });

  it('saves to disk and reloads with interrupted turns closed out', async () => {
    const file = sessionsFile(ws());
    const store = createSessionStore({ file });
    const s = store.create('codex', 'budget');
    store.patch(s.id, { status: 'running' });
    store.appendItem(s.id, { kind: 'tool', id: 't', ts: '', name: 'command', title: 'Run command', detail: null, status: 'running' });
    store.appendItem(s.id, { kind: 'question', id: 'q', ts: '', questions: [], answers: null, status: 'pending' });
    store.appendItem(s.id, { kind: 'check', id: 'c', ts: '', workbooks: ['budget'], status: 'running', findings: [], total: 0, error: null });
    await store.flush();
    expect(existsSync(file)).toBe(true);
    const again = createSessionStore({ file }).get(s.id);
    expect(again.status).toBe('idle');
    expect(again.items.map((i: any) => i.status)).toEqual(['error', 'cancelled', 'failed']);
  });

  it('ignores a file it cannot read, and sessions of agents it no longer offers', () => {
    const dir = ws();
    const bad = path.join(dir, 'bad.json');
    writeFileSync(bad, '{nope');
    expect(createSessionStore({ file: bad }).list()).toEqual([]);
    const old = path.join(dir, 'old.json');
    writeFileSync(old, JSON.stringify({ version: 1, sessions: [{ id: 's-1', provider: 'claude', items: [], updatedAt: '' }] }));
    expect(createSessionStore({ file: old }).list()).toEqual([]);
    expect(readFileSync(old, 'utf8')).toContain('claude');
  });

  it('names a conversation after its first line', () => {
    expect(titleFromPrompt('\n  Add a total row \nthen chart it')).toBe('Add a total row');
    expect(titleFromPrompt('x'.repeat(80))).toHaveLength(60);
    expect(titleFromPrompt('   ')).toBe('Untitled');
  });
});
