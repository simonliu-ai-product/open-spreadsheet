// Recalculation timings for three shapes real sheets have.
//   node --experimental-strip-types packages/engine/bench/recalc.ts
import { Workbook } from '../src/workbook.ts';

function time<T>(label: string, f: () => T): T {
  const t = performance.now();
  const r = f();
  console.log(`${label.padEnd(52)} ${(performance.now() - t).toFixed(0).padStart(6)} ms`);
  return r;
}

{
  const N = 100_000;
  const wb = new Workbook();
  wb.addSheet('S');
  const cells: Record<string, string> = { A1: '1' };
  for (let i = 2; i <= N; i++) cells[`A${i}`] = `=A${i - 1}+1`;
  time(`chain of ${N.toLocaleString()} formulas: build + first calc`, () => wb.setMany('S', cells));
  time('chain: edit A1, recalculate everything after it', () => wb.set('S', 'A1', '10'));
  console.log(`  A${N} = ${wb.get('S', `A${N}`)}  (expected ${N + 9})`);
}

{
  const N = 10_000;
  const wb = new Workbook();
  wb.addSheet('S');
  const cells: Record<string, string> = {};
  for (let i = 1; i <= N; i++) {
    cells[`A${i}`] = String(i);
    cells[`B${i}`] = `=A${i}*2`;
  }
  cells.C1 = '=SUM(B:B)';
  time(`${N.toLocaleString()} rows + a column total: build + first calc`, () => wb.setMany('S', cells));
  time('rows: edit one input', () => wb.set('S', 'A5000', '0'));
  console.log(`  C1 = ${wb.get('S', 'C1')}  (expected ${N * (N + 1) - 10_000})`);
}

{
  const N = 2_000;
  const wb = new Workbook();
  wb.addSheet('S');
  const cells: Record<string, string> = {};
  for (let i = 1; i <= N; i++) {
    cells[`A${i}`] = '1';
    cells[`B${i}`] = `=SUM($A$1:A${i})`;
  }
  time(`${N.toLocaleString()} running totals: build + first calc`, () => wb.setMany('S', cells));
  time('running totals: edit the first input', () => wb.set('S', 'A1', '5'));
  console.log(`  B${N} = ${wb.get('S', `B${N}`)}  (expected ${N + 4})`);
}
