import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
// @ts-expect-error plain JS module
import { createCodexMapper, createCodexProvider, parseCodexModelsCache } from './codex.js';

const mapped = (events: any[], cwd = '/ws') => {
  const out: any[] = [];
  const mapper = createCodexMapper((e: any) => out.push(e), cwd);
  for (const e of events) mapper.handle(e);
  return { out, mapper };
};

describe('codex event mapping', () => {
  it('maps a thread to session, tool and text events', () => {
    const { out, mapper } = mapped([
      { type: 'thread.started', thread_id: 'th-1' },
      { type: 'turn.started' },
      { type: 'item.started', item: { id: 'm1', type: 'mcp_tool_call', server: 'open-spreadsheet', tool: 'write_cells', arguments: { workbook: 'budget', sheet: 'Sheet1', cells: { B4: '=SUM(B2:B3)', A4: 'Total' } }, status: 'in_progress' } },
      { type: 'item.updated', item: { id: 'm1', type: 'mcp_tool_call', server: 'open-spreadsheet', tool: 'write_cells', arguments: {}, status: 'in_progress' } },
      { type: 'item.completed', item: { id: 'm1', type: 'mcp_tool_call', server: 'open-spreadsheet', tool: 'write_cells', arguments: {}, status: 'completed' } },
      { type: 'item.started', item: { id: 'c1', type: 'command_execution', command: 'node bin.js check', aggregated_output: '', status: 'in_progress' } },
      { type: 'item.completed', item: { id: 'c1', type: 'command_execution', command: 'node bin.js check', aggregated_output: '', exit_code: 1, status: 'failed' } },
      { type: 'item.completed', item: { id: 'f1', type: 'file_change', changes: [{ path: '/ws/sheets/budget/workbook.json', kind: 'update' }], status: 'completed' } },
      { type: 'item.completed', item: { id: 'r1', type: 'reasoning', text: 'thinking' } },
      { type: 'item.completed', item: { id: 'a1', type: 'agent_message', text: 'Added a total.' } },
      { type: 'turn.completed', usage: {} },
    ]);
    expect(out).toEqual([
      { type: 'session-id', id: 'th-1' },
      { type: 'tool-start', id: 'm1', name: 'mcp', title: 'write_cells', detail: 'budget · Sheet1 · B4, A4' },
      { type: 'tool-end', id: 'm1', ok: true },
      { type: 'tool-start', id: 'c1', name: 'command', title: 'Run command', detail: 'node bin.js check' },
      { type: 'tool-end', id: 'c1', ok: false },
      { type: 'tool-start', id: 'f1', name: 'file_change', title: 'Edit', detail: 'sheets/budget/workbook.json' },
      { type: 'tool-end', id: 'f1', ok: true },
      { type: 'text-start', id: 'a1' },
      { type: 'text-end', id: 'a1', text: 'Added a total.' },
      { type: 'result', durationMs: null, costUsd: null },
    ]);
    expect(mapper.resultSeen).toBe(true);
  });

  it('names other MCP servers and turns failures into errors', () => {
    const { out } = mapped([
      { type: 'item.completed', item: { id: 'x', type: 'mcp_tool_call', server: 'github', tool: 'search', arguments: {}, status: 'completed', error: { message: 'no' } } },
      { type: 'item.completed', item: { id: 'e', type: 'error', message: 'model overloaded' } },
      { type: 'turn.failed', error: { message: 'Not signed in' } },
    ]);
    expect(out).toEqual([
      { type: 'tool-start', id: 'x', name: 'mcp', title: 'github: search', detail: null },
      { type: 'tool-end', id: 'x', ok: false },
      { type: 'error', text: 'model overloaded' },
      { type: 'error', text: 'Not signed in' },
      { type: 'result', durationMs: null, costUsd: null },
    ]);
  });

  it('turns a tagged question block into questions', () => {
    const { out } = mapped([{ type: 'item.completed', item: { id: 'a', type: 'agent_message', text: 'First:\n<open-spreadsheet-questions>[{"question":"Which sheet?","options":[{"label":"Sales"},{"label":"Costs"}]}]</open-spreadsheet-questions>' } }]);
    expect(out.map((e) => e.type)).toEqual(['text-start', 'text-end', 'questions']);
    expect(out[1].text).toBe('First:');
    expect(out[2].questions[0].question).toBe('Which sheet?');
  });
});

describe('codex models', () => {
  it("reads the CLI's model cache in its own order", () => {
    expect(
      parseCodexModelsCache({
        models: [
          { slug: 'b', display_name: 'B', priority: 2, supported_reasoning_levels: [{ effort: 'low' }, { effort: 'ultra' }, { effort: 'Not a level!' }] },
          { slug: 'hidden', visibility: 'hide' },
          { slug: 'a', priority: 1, supported_reasoning_levels: ['high'] },
          { nope: true },
        ],
      }),
    ).toEqual([
      { id: 'a', label: 'a', efforts: ['high'] },
      // A level the panel has no word for yet (ultra) still shows; malformed ones do not.
      { id: 'b', label: 'B', efforts: ['low', 'ultra'] },
    ]);
    expect(parseCodexModelsCache(null)).toEqual([]);
  });
});

describe('codex provider', () => {
  it('runs a thread from the workspace SDK with the open-spreadsheet MCP server and the panel preamble', async () => {
    // A stand-in SDK in the workspace's node_modules: no real Codex, login or network.
    const ws = mkdtempSync(path.join(tmpdir(), 'oss-codex-'));
    const dir = path.join(ws, 'node_modules', '@openai', 'codex-sdk');
    mkdirSync(dir, { recursive: true });
    writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ name: '@openai/codex-sdk', version: '9.9.9', type: 'module', exports: { '.': { import: './index.js' } } }));
    writeFileSync(
      path.join(dir, 'index.js'),
      `export class Codex {
  constructor(options) { globalThis.__codexOptions = options; }
  startThread(options) { return this.thread(null, options); }
  resumeThread(id, options) { return this.thread(id, options); }
  thread(id, options) {
    globalThis.__codexThread = { id, options };
    return { id: id ?? 'th-new', async runStreamed(input) {
      globalThis.__codexInput = input;
      return { events: (async function* () {
        yield { type: 'thread.started', thread_id: 'th-new' };
        yield { type: 'item.completed', item: { id: 'a', type: 'agent_message', text: 'ok' } };
        yield { type: 'turn.completed', usage: {} };
      })() };
    } };
  }
}`,
    );
    const g = globalThis as any;
    const provider = createCodexProvider({ mcpServer: { command: '/usr/bin/node', args: ['/core/mcp.js', '--cwd', ws] } });
    const events: any[] = [];
    const base = { cwd: ws, writableRoots: [path.join(ws, 'sheets'), '/elsewhere/shared'], prompt: 'hi', systemPromptAppend: 'PREAMBLE', model: 'gpt-x', effort: 'low', signal: new AbortController().signal, emit: (e: any) => events.push(e), askUser: async () => null, requestPermission: async () => false };
    await provider.run({ ...base, resume: null });
    expect(g.__codexOptions.config).toEqual({
      mcp_servers: { 'open-spreadsheet': { command: '/usr/bin/node', args: ['/core/mcp.js', '--cwd', ws], default_tools_approval_mode: 'approve' } },
    });
    expect(g.__codexThread.options).toMatchObject({ workingDirectory: ws, sandboxMode: 'workspace-write', approvalPolicy: 'never', model: 'gpt-x', modelReasoningEffort: 'low', additionalDirectories: ['/elsewhere/shared'] });
    expect(g.__codexInput).toBe('PREAMBLE\n\nhi');
    expect(events.map((e) => e.type)).toEqual(['session-id', 'text-start', 'text-end', 'result', 'session-id']);
    // A resumed thread already has the preamble.
    await provider.run({ ...base, resume: 'th-new' });
    expect(g.__codexThread.id).toBe('th-new');
    expect(g.__codexInput).toBe('hi');
  });

  it('says how to connect when the SDK is not in the workspace', async () => {
    const provider = createCodexProvider();
    await expect(provider.run({ cwd: mkdtempSync(path.join(tmpdir(), 'oss-codex-none-')), writableRoots: [], prompt: '', systemPromptAppend: '', resume: null, model: null, effort: null, signal: new AbortController().signal, emit() {}, askUser: async () => null, requestPermission: async () => false })).rejects.toThrow('install @openai/codex-sdk');
  });
});
