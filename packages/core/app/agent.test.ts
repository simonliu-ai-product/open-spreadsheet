import { describe, expect, it } from 'vitest';
import { applyEvent, checkFixPrompt, getState, loadSession, resolveModelChoice } from './agent.ts';

describe('agent panel client', () => {
  it('keeps transcripts in step with the event stream', async () => {
    const session = { id: 's-1', provider: 'codex' as const, title: 'T', createdAt: '', updatedAt: '2026-01-01', status: 'running' as const, workbook: 'budget', providerSessionId: null, items: [] };
    // A session the panel holds in full, as loadSession leaves it.
    globalThis.fetch = (async () => new Response(JSON.stringify(session))) as typeof fetch;
    await loadSession('s-1');
    applyEvent({ type: 'item', sessionId: 's-1', item: { kind: 'assistant', id: 'a', ts: '', text: 'Hel', streaming: true } });
    applyEvent({ type: 'delta', sessionId: 's-1', itemId: 'a', text: 'lo' });
    applyEvent({ type: 'item', sessionId: 's-1', item: { kind: 'error', id: 'e', ts: '', text: 'x' } });
    expect(getState().full['s-1']!.items.map((i) => (i.kind === 'assistant' ? i.text : i.kind))).toEqual(['Hello', 'error']);
    applyEvent({ type: 'truncate', sessionId: 's-1', itemId: 'e' });
    expect(getState().full['s-1']!.items).toHaveLength(1);
    applyEvent({ type: 'session', session: { ...session, status: 'idle', updatedAt: '2026-01-02' } });
    expect(getState().full['s-1']!.status).toBe('idle');
    expect(getState().sessions[0]!.id).toBe('s-1');
    applyEvent({ type: 'session-removed', sessionId: 's-1' });
    expect(getState().full['s-1']).toBeUndefined();
    expect(getState().sessions).toEqual([]);
  });

  it('falls back to the first model and drops an effort it does not offer', () => {
    const models = [
      { id: 'a', label: 'A', efforts: ['low' as const] },
      { id: 'b', label: 'B', efforts: [] },
    ];
    expect(resolveModelChoice(models, { model: 'b', effort: 'low' })).toEqual({ model: models[1], effort: null });
    expect(resolveModelChoice(models, { model: 'gone', effort: 'low' })).toEqual({ model: models[0], effort: 'low' });
    expect(resolveModelChoice([], undefined)).toBeNull();
  });

  it('asks for a fix of what the check found', () => {
    const text = checkFixPrompt([{ workbook: 'budget', sheet: 'Sheet1', address: 'B4', severity: 'error', message: '=FOO() → #NAME?' }], 3);
    expect(text).toContain('- budget · Sheet1!B4 [error] =FOO() → #NAME?');
    expect(text).toContain('…and 2 more');
  });
});
