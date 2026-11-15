import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
// @ts-expect-error plain JS module
import { checkWorkbook, readRange, readWorkbook, writeWorkbook } from './ops.js';
// @ts-expect-error plain JS module
import { initWorkspace } from './server.js';

const fresh = () => initWorkspace(path.join(mkdtempSync(path.join(tmpdir(), 'oss-agent-')), 'ws'));

describe('init', () => {
  it('writes agent files and skills', () => {
    const root = fresh();
    for (const f of ['AGENTS.md', 'CLAUDE.md', '.mcp.json', '.claude/skills/sheet-authoring/SKILL.md', '.agents/skills/current-sheet/SKILL.md', 'sheets/getting-started/workbook.json']) {
      expect(existsSync(path.join(root, f)), f).toBe(true);
    }
    expect(readFileSync(path.join(root, '.gitignore'), 'utf8')).toContain('sheets/.trash');
  });
});

describe('ops', () => {
  it('refuses a write based on a stale version', async () => {
    const root = fresh();
    const { data, hash } = readWorkbook(root, 'getting-started');
    await writeWorkbook(root, 'getting-started', { ...data, title: 'One' }, { expectedHash: hash });
    await expect(writeWorkbook(root, 'getting-started', { ...data, title: 'Two' }, { expectedHash: hash })).rejects.toMatchObject({ code: 'conflict' });
    expect(readWorkbook(root, 'getting-started').data.title).toBe('One');
  });

  it('reads a range with inputs, values and display text', () => {
    const out = readRange(
      { title: 't', sheets: [{ name: 'S', cells: { A1: '0.5', A2: '=A1*2' }, styles: { A1: { fmt: '0%' } } }] },
      'S',
      'A1:A2',
    );
    expect(out.cells).toEqual([
      { address: 'A1', input: '0.5', value: 0.5, display: '50%' },
      { address: 'A2', input: '=A1*2', value: 1 },
    ]);
  });

  it('checks for broken formulas and malformed entries', () => {
    const problems = checkWorkbook({
      title: 't',
      sheets: [{ name: 'S', cells: { A1: '=FOO(1)', A2: '=Nope!A1', A3: '=A3+1', B1: '=1/0', ZZZZ9: 'x' }, merges: ['A1:B2', 'B2:C3'] }],
    });
    const at = (a: string) => problems.find((p: { address: string }) => p.address === a);
    expect(at('A1').message).toContain('#NAME?');
    expect(at('A2').message).toContain('#REF!');
    expect(at('A3').message).toContain('#CIRC!');
    expect(at('B1').severity).toBe('warning');
    expect(at('ZZZZ9').severity).toBe('error');
    expect(problems.some((p: { message: string }) => p.message.includes('overlap'))).toBe(true);
  });
});

describe('MCP over stdio', () => {
  it('lists tools, writes cells and reports problems', async () => {
    const root = fresh();
    const child = spawn(process.execPath, [path.join(__dirname, 'mcp.js'), '--cwd', root]);
    const replies = new Map<number, any>();
    let buf = '';
    child.stdout.on('data', (d) => {
      buf += d;
      for (let i = buf.indexOf('\n'); i >= 0; i = buf.indexOf('\n')) {
        const m = JSON.parse(buf.slice(0, i));
        buf = buf.slice(i + 1);
        replies.set(m.id, m);
      }
    });
    const call = (id: number, method: string, params?: object) => {
      child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
      return new Promise<any>((res) => {
        const t = setInterval(() => replies.has(id) && (clearInterval(t), res(replies.get(id))), 10);
      });
    };
    await call(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 't', version: '1' } });
    const list = await call(2, 'tools/list');
    expect(list.result.tools.map((t: { name: string }) => t.name)).toContain('write_cells');
    const wrote = await call(3, 'tools/call', { name: 'write_cells', arguments: { workbook: 'getting-started', sheet: 'Sheet1', cells: { B3: '200', C1: '=FOO()' } } });
    expect(wrote.result.isError).toBeUndefined();
    expect(wrote.result.structuredContent.problems[0].address).toBe('C1');
    const read = await call(4, 'tools/call', { name: 'read_range', arguments: { workbook: 'getting-started', range: 'B4' } });
    expect(read.result.structuredContent.cells[0].value).toBe(320);
    const renamed = await call(5, 'tools/call', { name: 'rename_sheet', arguments: { workbook: 'getting-started', from: 'Sheet1', to: 'Costs' } });
    expect(renamed.result.isError).toBeUndefined();
    const file = JSON.parse(readFileSync(path.join(root, 'sheets/getting-started/workbook.json'), 'utf8'));
    expect(file.sheets[0].name).toBe('Costs');
    expect(file.sheets[0].styles.A1.b).toBe(true);
    const noted = await call(6, 'tools/call', { name: 'set_note', arguments: { workbook: 'getting-started', sheet: 'Costs', address: 'b4', text: 'Show as NT$' } });
    expect(noted.result.isError).toBeUndefined();
    const notes = await call(7, 'tools/call', { name: 'list_notes', arguments: {} });
    expect(notes.result.structuredContent.notes).toMatchObject([{ sheet: 'Costs', address: 'B4', input: '=SUM(B2:B3)', note: 'Show as NT$' }]);
    await call(8, 'tools/call', { name: 'insert_rows', arguments: { workbook: 'getting-started', sheet: 'Costs', before: 1, count: 1 } });
    await call(9, 'tools/call', { name: 'set_note', arguments: { workbook: 'getting-started', sheet: 'Costs', address: 'B5', text: '' } });
    const after = JSON.parse(readFileSync(path.join(root, 'sheets/getting-started/workbook.json'), 'utf8'));
    expect(after.sheets[0].notes).toBeUndefined();
    const named = await call(10, 'tools/call', { name: 'set_names', arguments: { workbook: 'getting-started', names: { Amounts: 'Costs!$B$3:$B$4' } } });
    expect(named.result.isError).toBeUndefined();
    const total = await call(11, 'tools/call', { name: 'evaluate', arguments: { workbook: 'getting-started', formula: '=SUM(Amounts)' } });
    expect(total.result.structuredContent.value).toBe(320);
    const bad = await call(12, 'tools/call', { name: 'set_names', arguments: { workbook: 'getting-started', names: { A1: 'Costs!$B$3' } } });
    expect(bad.result.isError).toBe(true);
    expect(JSON.parse(readFileSync(path.join(root, 'sheets/getting-started/workbook.json'), 'utf8')).names).toEqual({ Amounts: 'Costs!$B$3:$B$4' });
    const traced = await call(13, 'tools/call', { name: 'trace_cell', arguments: { workbook: 'getting-started', sheet: 'costs', address: 'B3' } });
    expect(traced.result.structuredContent).toMatchObject({ sheet: 'Costs', address: 'B3', precedents: [], dependents: [{ sheet: 'Costs', address: 'B5', input: '=SUM(B3:B4)' }] });
    const from = await call(14, 'tools/call', { name: 'trace_cell', arguments: { workbook: 'getting-started', address: 'B5' } });
    expect(from.result.structuredContent.precedents).toEqual([{ sheet: 'Costs', range: 'B3:B4' }]);
    const copy = await call(15, 'tools/call', { name: 'duplicate_workbook', arguments: { workbook: 'getting-started' } });
    expect(copy.result.structuredContent).toMatchObject({ id: 'getting-started-copy', title: 'Getting started (copy)' });
    const copied = JSON.parse(readFileSync(path.join(root, 'sheets/getting-started-copy/workbook.json'), 'utf8'));
    expect(copied.names).toEqual({ Amounts: 'Costs!$B$3:$B$4' });
    expect((await call(16, 'tools/call', { name: 'duplicate_workbook', arguments: { workbook: 'getting-started', new_id: 'getting-started-copy' } })).result.isError).toBe(true);
    await call(17, 'tools/call', { name: 'rename_workbook', arguments: { workbook: 'getting-started-copy', title: 'Q3 draft' } });
    expect(JSON.parse(readFileSync(path.join(root, 'sheets/getting-started-copy/workbook.json'), 'utf8')).title).toBe('Q3 draft');
    const gone = await call(18, 'tools/call', { name: 'delete_workbook', arguments: { workbook: 'getting-started-copy' } });
    expect(gone.result.structuredContent.trashedTo).toMatch(/^sheets\/\.trash\/getting-started-copy-/);
    expect(existsSync(path.join(root, 'sheets/getting-started-copy'))).toBe(false);
    const listed = await call(19, 'tools/call', { name: 'list_workbooks', arguments: {} });
    expect(JSON.stringify(listed.result.structuredContent)).not.toContain('getting-started-copy');
    child.kill();
  }, 20_000);
});

describe('export', () => {
  it('writes computed values with plain numbers, ISO dates and a header row', async () => {
    // @ts-expect-error plain JS module
    const { exportRows, toCsvText, toJsonRecords } = await import('./ops.js');
    const data = {
      title: 't',
      sheets: [{ name: 'S', cells: { A1: 'Day', B1: 'Amount', A2: '2026-01-05', B2: '1,200', A3: '2026-01-06', B3: '=B2*2', C2: '=1/0' }, styles: { B2: { fmt: '#,##0' } } }],
    };
    const { rows, firstCol } = exportRows(data, 'S', 'A1:B3');
    expect(firstCol).toBe(0);
    expect(toCsvText(rows)).toBe('Day,Amount\n2026-01-05,1200\n2026-01-06,2400\n');
    expect(toJsonRecords(rows)).toEqual([
      { Day: '2026-01-05', Amount: 1200 },
      { Day: '2026-01-06', Amount: 2400 },
    ]);
    expect(exportRows(data, 'S', 'C2').rows).toEqual([[null]]);
  });
});

describe('plugin', () => {
  it('ships the same skills as the CLI installs', async () => {
    const { readdirSync } = await import('node:fs');
    const core = path.join(__dirname, 'skills');
    const plugin = path.join(__dirname, '../../plugins/open-spreadsheet/skills');
    for (const name of readdirSync(core)) {
      expect(readFileSync(path.join(plugin, name, 'SKILL.md'), 'utf8'), name).toBe(readFileSync(path.join(core, name, 'SKILL.md'), 'utf8'));
    }
  });
});

describe('packaging', () => {
  // Node will not strip types inside node_modules, and an installed package only has what `files` lists
  // (open-doc 0.14.0 shipped without sources its viewer imported). Both failures only show after npm install.
  it('ships every file core loads, and imports the engine by its package name', async () => {
    const { readdirSync, statSync } = await import('node:fs');
    const pkg = JSON.parse(readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
    const packed = (rel: string) => pkg.files.some((e: string) => rel === e || rel.startsWith(`${e}/`));
    const sources = ['bin.js', 'mcp.js', 'server.js', 'ops.js', ...readdirSync(path.join(__dirname, 'app')).filter((f) => /\.ts$/.test(f) && !f.endsWith('.test.ts')).map((f) => `app/${f}`)];
    const problems: string[] = [];
    for (const rel of sources) {
      const text = readFileSync(path.join(__dirname, rel), 'utf8');
      for (const m of text.matchAll(/^(?:import|export)\s+(?!type\b)[^;]*?from\s*['"]([^'"]+)['"]/gms)) {
        const spec = m[1]!;
        if (spec.startsWith('.')) {
          const target = path.relative(__dirname, path.resolve(path.dirname(path.join(__dirname, rel)), spec));
          if (target.startsWith('..')) problems.push(`${rel} reaches outside the package: ${spec}`);
          else if (!packed(target) || !statSync(path.join(__dirname, target)).isFile()) problems.push(`${rel} imports unpacked ${spec}`);
        } else if (!spec.startsWith('node:') && !(spec in pkg.dependencies) && !spec.startsWith('vite')) {
          problems.push(`${rel} imports ${spec}, which is not a dependency`);
        }
      }
    }
    expect(problems).toEqual([]);
  });

  it('builds an engine that runs from node_modules without type stripping', async () => {
    const { execFileSync } = await import('node:child_process');
    const dir = path.join(mkdtempSync(path.join(tmpdir(), 'oss-pack-')), 'node_modules', 'engine');
    execFileSync(process.execPath, [path.join(__dirname, '../engine/scripts/build.mjs'), dir]);
    const engine = JSON.parse(readFileSync(path.join(__dirname, '../engine/package.json'), 'utf8'));
    expect(engine.publishConfig.exports['.'].default).toBe('./dist/index.js');
    expect(engine.files).toContain('dist');
    const out = execFileSync(process.execPath, ['--input-type=module', '-e', `import { Workbook } from ${JSON.stringify(path.join(dir, 'index.js'))}; const wb = Workbook.fromJSON({ sheets: [{ name: 'S', cells: { A1: '2', A2: '=A1*21' } }] }); console.log(wb.value('S', 1, 0));`], { encoding: 'utf8' });
    expect(out.trim()).toBe('42');
  });
});
