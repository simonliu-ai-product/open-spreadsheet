import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
// @ts-expect-error plain JS module
import { readWorkbook, writeWorkbook } from '../ops.js';
// @ts-expect-error plain JS module
import { initWorkspace } from '../server.js';
// @ts-expect-error plain JS module
import { CHECK_FINDINGS_KEPT, changedWorkbooks, checkItemResult, checkWorkbooks, shouldCheckAfterTurn, snapshotWorkbooks } from './check-after-turn.js';

const fresh = () => initWorkspace(path.join(mkdtempSync(path.join(tmpdir(), 'oss-agent-check-')), 'ws'));

describe('check after a turn', () => {
  it('finds the workbooks a turn changed or created', async () => {
    const root = fresh();
    const before = snapshotWorkbooks(root);
    expect([...before.keys()]).toEqual(['getting-started']);
    const { data } = readWorkbook(root, 'getting-started');
    await writeWorkbook(root, 'getting-started', { ...data, title: 'Changed' });
    await writeWorkbook(root, 'fresh-one', { title: 'New', sheets: [{ name: 'S', cells: {} }] }, { create: true });
    expect(changedWorkbooks(before, snapshotWorkbooks(root)).sort()).toEqual(['fresh-one', 'getting-started']);
    expect(changedWorkbooks(before, before)).toEqual([]);
  });

  it('checks only a finished turn that used a tool and changed a workbook', () => {
    const items = [
      { kind: 'user', id: 'u1', ts: '' },
      { kind: 'tool', id: 't', ts: '' },
      { kind: 'result', id: 'r', ts: '', aborted: false },
    ];
    expect(shouldCheckAfterTurn({ session: { items }, userItemId: 'u1', changed: ['a'] })).toBe(true);
    expect(shouldCheckAfterTurn({ session: { items }, userItemId: 'u1', changed: [] })).toBe(false);
    expect(shouldCheckAfterTurn({ session: { items: [items[0], items[2]] }, userItemId: 'u1', changed: ['a'] })).toBe(false);
    expect(shouldCheckAfterTurn({ session: { items: [items[0], items[1], { ...items[2], aborted: true }] }, userItemId: 'u1', changed: ['a'] })).toBe(false);
    expect(shouldCheckAfterTurn({ session: { items }, userItemId: 'missing', changed: ['a'] })).toBe(false);
  });

  it("runs the framework's own check and keeps errors first", async () => {
    const root = fresh();
    await writeWorkbook(root, 'broken', { title: 'B', sheets: [{ name: 'S', cells: { A1: '=1/0', A2: '=FOO(1)' } }] }, { create: true });
    const results = checkWorkbooks(root, ['getting-started', 'broken', 'gone']);
    expect(results.map((r: any) => r.problems.length)).toEqual([0, 2, 1]);
    const item = checkItemResult(results);
    expect(item.total).toBe(3);
    expect(item.findings.map((f: any) => [f.workbook, f.address, f.severity])).toEqual([
      ['broken', 'A2', 'error'],
      ['gone', '', 'error'],
      ['broken', 'A1', 'warning'],
    ]);
    const many = checkItemResult([{ id: 'x', problems: Array.from({ length: 30 }, (_, n) => ({ severity: 'warning', sheet: 'S', address: `A${n + 1}`, message: 'm' })) }]);
    expect(many.findings).toHaveLength(CHECK_FINDINGS_KEPT);
    expect(many.total).toBe(30);
  });
});
