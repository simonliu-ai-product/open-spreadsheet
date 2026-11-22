import { describe, expect, it } from 'vitest';
// @ts-expect-error plain JS module
import { createAgentRunner } from './runner.js';
// @ts-expect-error plain JS module
import { createSessionStore } from './store.js';

type Run = (opts: any) => Promise<void>;

const settle = () => new Promise((r) => setTimeout(r, 0));

function setup(run: Run, hooks: { onTurnStart?: (t: any) => void; onTurnEnd?: (t: any) => void } = {}) {
  const store = createSessionStore({ file: null });
  const provider = { id: 'codex', name: 'Fake', sdkPackage: 'fake', cliCommand: 'fake', askInstruction: 'Ask with the fake block.', asksPermission: true, models: async () => [], run };
  const runner = createAgentRunner({ cwd: '/ws', writableRoots: ['/ws/sheets', '/ws/assets'], store, providers: { codex: provider }, promptOptions: () => ({ checkCommand: 'open-spreadsheet check' }), ...hooks });
  return { store, runner };
}

const kinds = (store: any, id: string) => store.get(id).items.map((i: any) => i.kind);

describe('agent runner', () => {
  it('records a turn: the message, streamed text, tools and the result', async () => {
    const seen: any[] = [];
    const { store, runner } = setup(async (opts) => {
      seen.push(opts);
      opts.emit({ type: 'session-id', id: 'thread-1' });
      opts.emit({ type: 'text-delta', id: 't1', text: 'Sum' });
      opts.emit({ type: 'text-delta', id: 't1', text: 'med' });
      opts.emit({ type: 'tool-start', id: 'x', name: 'mcp', title: 'write_cells', detail: 'budget · Sheet1 · B4' });
      opts.emit({ type: 'tool-end', id: 'x', ok: true });
      opts.emit({ type: 'text-end', id: 't1', text: 'Summed B2:B3.' });
      opts.emit({ type: 'result', durationMs: 7, costUsd: null });
    });
    const s = store.create('codex', null);
    await runner.send(s.id, { text: 'add a total', context: { workbook: 'budget', sheet: 'Sheet1', active: 'B4', selection: 'B4' } });
    expect(runner.isRunning(s.id)).toBe(true);
    await settle();
    const after = store.get(s.id);
    expect(after).toMatchObject({ status: 'idle', title: 'add a total', workbook: 'budget', providerSessionId: 'thread-1' });
    expect(kinds(store, s.id)).toEqual(['user', 'assistant', 'tool', 'result']);
    expect(after.items[1]).toMatchObject({ text: 'Summed B2:B3.', streaming: false });
    expect(after.items[2]).toMatchObject({ status: 'done', title: 'write_cells' });
    expect(after.items[3]).toMatchObject({ durationMs: 7, aborted: false });
    expect(seen[0].prompt).toContain('<open-spreadsheet-context>');
    expect(seen[0].systemPromptAppend).toContain('Ask with the fake block.');
    expect(seen[0].writableRoots).toEqual(['/ws/sheets', '/ws/assets']);
    expect(seen[0].resume).toBeNull();
  });

  it('resumes the provider session on the next turn and reports turn start and end', async () => {
    const resumes: (string | null)[] = [];
    const events: string[] = [];
    const { store, runner } = setup(
      async (opts) => {
        resumes.push(opts.resume);
        opts.emit({ type: 'session-id', id: 'thread-9' });
      },
      { onTurnStart: () => events.push('start'), onTurnEnd: () => events.push('end') },
    );
    const s = store.create('codex', null);
    await runner.send(s.id, { text: 'one', model: 'gpt-x', effort: 'high' });
    await settle();
    await settle();
    await runner.send(s.id, { text: 'two' });
    await settle();
    await settle();
    expect(resumes).toEqual([null, 'thread-9']);
    expect(events).toEqual(['start', 'end', 'start', 'end']);
  });

  it('refuses a second turn while one runs, and stops one on request', async () => {
    const { store, runner } = setup(
      (opts) =>
        new Promise<void>((resolve) => {
          opts.emit({ type: 'tool-start', id: 'c', name: 'command', title: 'Run command', detail: 'ls' });
          opts.signal.addEventListener('abort', () => resolve());
        }),
    );
    const s = store.create('codex', null);
    await runner.send(s.id, { text: 'long' });
    await expect(runner.send(s.id, { text: 'again' })).rejects.toThrow('already running');
    expect(runner.stop(s.id)).toBe(true);
    await settle();
    const items = store.get(s.id).items;
    expect(items.find((i: any) => i.kind === 'tool').status).toBe('error');
    expect(items.at(-1)).toMatchObject({ kind: 'result', aborted: true });
  });

  it('turns a provider failure into an error item', async () => {
    const { store, runner } = setup(async () => {
      throw new Error('Codex is not signed in');
    });
    const s = store.create('codex', null);
    await runner.send(s.id, { text: 'go' });
    await settle();
    expect(kinds(store, s.id)).toEqual(['user', 'error', 'result']);
    expect(store.get(s.id).items[1].text).toBe('Codex is not signed in');
  });

  it('answers questions that ended the turn with a new turn', async () => {
    const prompts: string[] = [];
    const { store, runner } = setup(async (opts) => {
      prompts.push(opts.prompt);
      if (prompts.length === 1) opts.emit({ type: 'questions', questions: [{ header: 'Period', question: 'Which months?', options: [{ label: 'Q1' }], multiSelect: false }] });
    });
    const s = store.create('codex', null);
    await runner.send(s.id, { text: 'summarise', context: { workbook: 'sales' } });
    await settle();
    const q = store.get(s.id).items.find((i: any) => i.kind === 'question');
    expect(q.status).toBe('pending');
    expect(runner.respond(s.id, q.id, { answers: { 'Which months?': 'Q1' } })).toBe(true);
    await settle();
    expect(store.get(s.id).items.find((i: any) => i.id === q.id)).toMatchObject({ status: 'answered', answers: { 'Which months?': 'Q1' } });
    expect(prompts[1]).toContain('Which months? Q1');
    // The question stands in for a user message; the context carries over.
    expect(kinds(store, s.id).filter((k: string) => k === 'user')).toHaveLength(1);
    expect(prompts[1]).toContain('workbook "sales"');
  });

  it('holds a permission prompt until the person answers', async () => {
    let allowed: boolean | null = null;
    const { store, runner } = setup(async (opts) => {
      allowed = await opts.requestPermission({ tool: 'command', title: 'Run command', detail: 'rm -rf sheets' });
    });
    const s = store.create('codex', null);
    await runner.send(s.id, { text: 'clean up' });
    await settle();
    const p = store.get(s.id).items.find((i: any) => i.kind === 'permission');
    expect(p.status).toBe('pending');
    runner.respond(s.id, p.id, { allow: false });
    await settle();
    expect(allowed).toBe(false);
    expect(store.get(s.id).items.find((i: any) => i.id === p.id).status).toBe('denied');
  });

  it('edits a message: the first restarts the thread, a later one is sent as an edit', async () => {
    const calls: { prompt: string; resume: string | null }[] = [];
    const { store, runner } = setup(async (opts) => {
      calls.push({ prompt: opts.prompt, resume: opts.resume });
      opts.emit({ type: 'session-id', id: `thread-${calls.length}` });
    });
    const s = store.create('codex', null);
    await runner.send(s.id, { text: 'first' });
    await settle();
    await runner.send(s.id, { text: 'second' });
    await settle();
    const [u1, u2] = store.get(s.id).items.filter((i: any) => i.kind === 'user');
    await runner.send(s.id, { text: 'second, better', edit: u2.id });
    await settle();
    expect(calls[2].prompt).toContain('<original-message>\nsecond\n</original-message>');
    expect(calls[2].resume).toBe('thread-2');
    await runner.send(s.id, { text: 'first, better', edit: u1.id });
    await settle();
    expect(calls[3]).toMatchObject({ prompt: 'first, better', resume: null });
    expect(store.get(s.id).items.filter((i: any) => i.kind === 'user').map((i: any) => i.text)).toEqual(['first, better']);
  });
});
