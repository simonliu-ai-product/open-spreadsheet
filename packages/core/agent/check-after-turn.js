/**
 * After a turn, the dev server runs `open-spreadsheet check` (ops.js
 * `checkWorkbook`) on every workbook the turn changed and puts the findings
 * in the transcript, so a broken formula never waits for the person to find it.
 *
 * Which workbooks changed: the hash of each workbook.json when the turn
 * started, compared with when it ended (new workbooks count too).
 *
 * @typedef {import('./types.d.ts').AgentCheckFinding} AgentCheckFinding
 */
import { existsSync, readFileSync } from 'node:fs';
import { checkWorkbook, hashOf, listWorkbooks, readWorkbook, workbookFile } from '../ops.js';

/** How many findings a check item keeps; the rest are only counted. */
export const CHECK_FINDINGS_KEPT = 20;

/** @returns {Map<string, string>} workbook id → SHA-256 of its file */
export function snapshotWorkbooks(root) {
  const hashes = new Map();
  for (const w of listWorkbooks(root)) {
    const file = workbookFile(root, w.id);
    try {
      if (existsSync(file)) hashes.set(w.id, hashOf(readFileSync(file, 'utf8')));
    } catch {
      // gone mid-read
    }
  }
  return hashes;
}

/** Workbooks added or changed between two snapshots, in the later one's order. */
export function changedWorkbooks(before, after) {
  return [...after].filter(([id, hash]) => before.get(id) !== hash).map(([id]) => id);
}

/**
 * Whether to check after the turn that just ended: it ran to its end, used a
 * tool (a person editing the grid meanwhile is no reason), and changed a workbook.
 */
export function shouldCheckAfterTurn({ session, userItemId, changed }) {
  if (!changed.length) return false;
  const start = session.items.findIndex((i) => i.id === userItemId);
  if (start === -1) return false;
  const turn = session.items.slice(start + 1);
  const result = [...turn].reverse().find((i) => i.kind === 'result');
  if (result?.aborted) return false;
  return turn.some((i) => i.kind === 'tool');
}

/** Checks workbooks the way `open-spreadsheet check` does. @returns {{ id: string, problems: object[] }[]} */
export function checkWorkbooks(root, ids) {
  return ids.map((id) => {
    try {
      return { id, problems: checkWorkbook(readWorkbook(root, id).data) };
    } catch (e) {
      return { id, problems: [{ severity: 'error', sheet: '', address: '', message: e.message }] };
    }
  });
}

/** Check results as the transcript keeps them: errors first, the first few kept, all counted. */
export function checkItemResult(results) {
  /** @type {AgentCheckFinding[]} */
  const all = results.flatMap((r) =>
    r.problems.map((p) => ({ workbook: r.id, sheet: p.sheet ?? '', address: p.address ?? '', severity: p.severity === 'error' ? 'error' : 'warning', message: p.message })),
  );
  const ordered = all.map((f, n) => ({ f, n })).sort((a, b) => Number(b.f.severity === 'error') - Number(a.f.severity === 'error') || a.n - b.n);
  return { findings: ordered.slice(0, CHECK_FINDINGS_KEPT).map((o) => o.f), total: all.length };
}
