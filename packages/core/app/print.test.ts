import { describe, expect, it } from 'vitest';
import { printHtml, type PrintInput } from './print.ts';

const base = (over: Partial<PrintInput> = {}): PrintInput => ({
  title: 'Budget',
  sheetName: 'Q1',
  rows: 3,
  cols: 3,
  cell: (r, c) => ({ text: `${String.fromCharCode(65 + c)}${r + 1}`, align: c === 2 ? 'right' : 'left' }),
  colWidth: () => 100,
  rowHidden: () => false,
  colHidden: () => false,
  merges: [],
  headerRows: 0,
  charts: [],
  lang: 'zh-Hant',
  ...over,
});

describe('print', () => {
  it('prints a font size relative to the default 11pt, and ignores one that is not a number', () => {
    const cell = (size: unknown) => () => ({ text: 'x', align: 'left' as const, style: { size } as never });
    expect(printHtml(base({ rows: 1, cols: 1, cell: cell(22) }))).toContain('font-size:2.000em');
    const top = () => ({ text: 'x', align: 'left' as const, style: { valign: 'top' } as never });
    expect(printHtml(base({ rows: 1, cols: 1, cell: top }))).toContain('vertical-align:top');
    const box = (text: string) => () => ({ text, align: 'center' as const, style: { check: true } as never });
    expect(printHtml(base({ rows: 1, cols: 1, cell: box('TRUE') }))).toContain('>☑</td>');
    expect(printHtml(base({ rows: 1, cols: 1, cell: box('FALSE') }))).toContain('>☐</td>');
    expect(printHtml(base({ rows: 1, cols: 1, cell: cell('1em;color:red') }))).toContain('<td style="text-align:left">x</td>');
  });

  it('lays out the used area with alignment and the language', () => {
    const html = printHtml(base());
    expect(html).toContain('<html lang="zh-Hant">');
    expect(html.match(/<tr>/g)).toHaveLength(3);
    expect(html).toContain('<td style="text-align:right">C1</td>');
  });

  it('leaves out hidden rows and columns and repeats frozen rows', () => {
    const html = printHtml(base({ rowHidden: (r) => r === 1, colHidden: (c) => c === 1, headerRows: 1 }));
    expect(html).not.toContain('>A2<');
    expect(html).not.toContain('>B1<');
    expect(html).toMatch(/<thead><tr><td[^>]*>A1<\/td><td[^>]*>C1<\/td><\/tr><\/thead>/);
  });

  it('spans merges, draws styles and borders, and escapes text', () => {
    const html = printHtml(
      base({
        merges: [{ r0: 0, c0: 0, r1: 0, c1: 2 }],
        cell: (r, c) =>
          r === 0 && c === 0
            ? { text: '<Total>', align: 'center', style: { b: true, bg: '#e2efda', border: { bottom: '#000000', thick: true } } }
            : { text: 'x', align: 'left' },
      }),
    );
    expect(html).toContain('<td colspan="3" rowspan="1" style="text-align:center;font-weight:600;background:#e2efda;border-bottom:2px solid #000000">&lt;Total&gt;</td>');
    expect(html.match(/<tr>/g)).toHaveLength(3);
  });
});
