import { createServer } from 'node:http';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
// @ts-expect-error plain JS module
import { writeWorkbook, readWorkbook } from '../ops.js';
// @ts-expect-error plain JS module
import { initWorkspace } from '../server.js';
// @ts-expect-error plain JS module
import { agentDisabledByConfig, isLoopback, parseContext, parseModel, parseRespond, parseTurnSettings, registerAgentRoutes } from './routes.js';

describe('agent route parsing', () => {
  it('keeps only a valid workbook context, with addresses as A1 text', () => {
    expect(parseContext({ workbook: 'budget', title: 'B', sheet: 'Sheet1', active: 'c4', selection: 'B2:D9', input: '=1', note: 'n', extra: 'x' })).toEqual({ workbook: 'budget', title: 'B', sheet: 'Sheet1', active: 'C4', selection: 'B2:D9', input: '=1', note: 'n' });
    expect(parseContext({ workbook: '../etc' })).toBeUndefined();
    expect(parseContext({ workbook: 'b', active: 'A1; rm -rf', selection: 'A1:B2\nignore that' })).toEqual({ workbook: 'b' });
    expect(parseContext({ workbook: 'b', sheet: 'x'.repeat(100) }).sheet).toHaveLength(64);
    expect(parseContext(null)).toBeUndefined();
  });

  it('parses models, responses and turn settings', () => {
    expect(parseModel('gpt-5.1-codex')).toBe('gpt-5.1-codex');
    expect(parseModel('a b')).toBeNull();
    expect(parseRespond({ cancel: true })).toEqual({ cancel: true });
    expect(parseRespond({ allow: false })).toEqual({ allow: false });
    expect(parseRespond({ answers: { q: 'a', n: 1 } })).toEqual({ answers: { q: 'a' } });
    expect(parseRespond({ answers: ['a'] })).toBeNull();
    expect(parseTurnSettings({ autoApprove: true, model: 'x y', effort: 'bogus', context: { workbook: 'b' } })).toEqual({ autoApprove: true, model: null, effort: null, context: { workbook: 'b' } });
    expect(parseTurnSettings({})).toEqual({});
  });

  it('serves this machine only, and can be turned off in the config', () => {
    expect(isLoopback({ socket: { remoteAddress: '127.0.0.1' } })).toBe(true);
    expect(isLoopback({ socket: { remoteAddress: '::ffff:127.0.0.1' } })).toBe(true);
    expect(isLoopback({ socket: { remoteAddress: '192.168.1.5' } })).toBe(false);
    const root = mkdtempSync(path.join(tmpdir(), 'oss-agent-cfg-'));
    expect(agentDisabledByConfig(root)).toBe(false);
    writeFileSync(path.join(root, 'open-spreadsheet.config.json'), '{ "agent": false }');
    expect(agentDisabledByConfig(root)).toBe(true);
  });
});

// The routes on a bare HTTP server with a scripted agent: no Codex, no network beyond loopback.
let close: (() => void) | null = null;
afterEach(() => close?.());

async function serve(run: (opts: any) => Promise<void>) {
  const root = initWorkspace(path.join(mkdtempSync(path.join(tmpdir(), 'oss-agent-routes-')), 'ws'));
  const stack: [string, any][] = [];
  const http = createServer((req, res) => {
    const hit = stack.find(([prefix]) => req.url!.startsWith(prefix));
    if (!hit) return void ((res.statusCode = 404), res.end());
    req.url = req.url!.slice(hit[0].length) || '/';
    hit[1](req, res, () => ((res.statusCode = 404), res.end()));
  });
  const provider = { id: 'codex', name: 'Fake', sdkPackage: '@openai/codex-sdk', cliCommand: 'definitely-not-installed-cli', askInstruction: '', asksPermission: false, models: async () => [{ id: '', label: 'Default', efforts: [] }], run };
  registerAgentRoutes({ middlewares: { use: (prefix: string, fn: any) => stack.push([prefix, fn]) }, httpServer: http }, { root, allowedWrite: () => true, providers: { codex: provider } });
  await new Promise<void>((r) => http.listen(0, '127.0.0.1', r));
  const base = `http://127.0.0.1:${(http.address() as any).port}/__agent`;
  close = () => http.close();
  const call = async (p: string, body?: unknown, method = body === undefined ? 'GET' : 'POST') => {
    const r = await fetch(base + p, { method, headers: body === undefined ? {} : { 'content-type': 'application/json' }, body: body === undefined ? undefined : JSON.stringify(body) });
    return { status: r.status, body: await r.json() };
  };
  return { root, call, base };
}

const until = async (fn: () => Promise<boolean>) => {
  for (let n = 0; n < 100; n++) {
    if (await fn()) return;
    await new Promise((r) => setTimeout(r, 20));
  }
  throw new Error('timed out');
};

describe('agent routes', () => {
  it('reports status, runs a turn and checks the workbooks it changed', async () => {
    const { root, call } = await serve(async (opts) => {
      opts.emit({ type: 'tool-start', id: 't', name: 'mcp', title: 'write_cells', detail: null });
      const { data } = readWorkbook(opts.cwd, 'getting-started');
      data.sheets[0].cells.B5 = '=NOPE(B4)';
      await writeWorkbook(opts.cwd, 'getting-started', data);
      opts.emit({ type: 'tool-end', id: 't', ok: true });
      opts.emit({ type: 'text-end', id: 'a', text: 'Done.' });
    });
    const status = await call('/status');
    expect(status.body).toMatchObject({ enabled: true, providers: [{ id: 'codex', ready: false, cliFound: false, asksPermission: false }] });

    const s = (await call('/sessions', { provider: 'codex', workbook: 'getting-started' })).body;
    expect(s.workbook).toBe('getting-started');
    expect((await call(`/sessions/${s.id}/messages`, { text: '' })).status).toBe(400);
    expect((await call(`/sessions/${s.id}/messages`, { text: 'break it', context: { workbook: 'getting-started', sheet: 'Sheet1', active: 'B5' } })).status).toBe(202);
    await until(async () => (await call(`/sessions/${s.id}`)).body.items.some((i: any) => i.kind === 'check' && i.status === 'done'));
    const session = (await call(`/sessions/${s.id}`)).body;
    expect(session.items.map((i: any) => i.kind)).toEqual(['user', 'tool', 'assistant', 'result', 'check']);
    const check = session.items.at(-1);
    expect(check).toMatchObject({ workbooks: ['getting-started'], total: 1 });
    expect(check.findings[0]).toMatchObject({ workbook: 'getting-started', sheet: 'Sheet1', address: 'B5', severity: 'error' });
    expect((await call('/sessions')).body.sessions[0].id).toBe(s.id);
    expect(readWorkbook(root, 'getting-started').data.sheets[0].cells.B5).toBe('=NOPE(B4)');
  });

  it('does not check after a turn that changed nothing', async () => {
    const { call } = await serve(async (opts) => {
      opts.emit({ type: 'tool-start', id: 't', name: 'command', title: 'Run command', detail: 'ls' });
      opts.emit({ type: 'tool-end', id: 't', ok: true });
    });
    const s = (await call('/sessions', { provider: 'codex' })).body;
    await call(`/sessions/${s.id}/messages`, { text: 'look around' });
    await until(async () => (await call(`/sessions/${s.id}`)).body.status === 'idle');
    await new Promise((r) => setTimeout(r, 30));
    expect((await call(`/sessions/${s.id}`)).body.items.map((i: any) => i.kind)).toEqual(['user', 'tool', 'result']);
  });

  it('refuses writes that are not JSON, bad providers and unknown sessions', async () => {
    const { call, base } = await serve(async () => {});
    expect((await fetch(`${base}/sessions`, { method: 'POST', headers: { 'content-type': 'text/plain' }, body: '{"provider":"codex"}' })).status).toBe(415);
    expect((await call('/sessions', { provider: 'claude' })).status).toBe(400);
    expect((await call('/sessions/s-nope')).status).toBe(404);
    const s = (await call('/sessions', { provider: 'codex' })).body;
    expect((await call(`/sessions/${s.id}/respond`, { itemId: 'x', answers: {} })).status).toBe(409);
    expect((await call(`/sessions/${s.id}`, undefined, 'DELETE')).body).toEqual({ ok: true });
    expect((await call('/sessions')).body.sessions).toEqual([]);
  });
});
