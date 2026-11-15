import { describe, expect, it } from 'vitest';
import { formatValue, impliedFormat } from './format.ts';
import { dateToSerial } from './dates.ts';
import { parseNumericText } from './values.ts';
import { offsetFormula, Workbook } from './workbook.ts';

const f = (v: number | string, code: string) => formatValue(v, code).text;

describe('number formats', () => {
  it('formats digits, grouping and decimals', () => {
    expect(f(1234.567, '#,##0.00')).toBe('1,234.57');
    expect(f(1234.567, '0')).toBe('1235');
    expect(f(0.5, '#,##0')).toBe('1');
    expect(f(0.25, '0.0')).toBe('0.3');
    expect(f(5, '000')).toBe('005');
    expect(f(1234567, '#,##0')).toBe('1,234,567');
    expect(f(-1234.5, '#,##0.00')).toBe('-1,234.50');
    expect(f(0, '#,##0.00')).toBe('0.00');
    expect(f(1.5, '0.##')).toBe('1.5');
    expect(f(2, '0.##')).toBe('2.');
  });

  it('formats percent, currency, scaling and scientific', () => {
    expect(f(0.1234, '0%')).toBe('12%');
    expect(f(0.1234, '0.00%')).toBe('12.34%');
    expect(f(1500, '$#,##0')).toBe('$1,500');
    expect(f(1500, '[$NT$-404]#,##0')).toBe('NT$1,500');
    expect(f(1234567, '#,##0,')).toBe('1,235');
    expect(f(12345, '0.00E+00')).toBe('1.23E+04');
    expect(f(0.000123, '0.00E+00')).toBe('1.23E-04');
    expect(f(5, '0" pcs"')).toBe('5 pcs');
  });

  it('uses the negative, zero and text sections', () => {
    expect(f(-5, '0;(0)')).toBe('(5)');
    expect(f(0, '0;-0;"zero"')).toBe('zero');
    expect(f('abc', '0;-0;0;"[" @ "]"')).toBe('[ abc ]');
    expect(formatValue(-5, '0;[Red]-0').color).toBe('#ff0000');
  });

  it('formats dates and times', () => {
    const d = dateToSerial(2026, 1, 5);
    expect(f(d, 'yyyy-mm-dd')).toBe('2026-01-05');
    expect(f(d, 'm/d/yy')).toBe('1/5/26');
    expect(f(d, 'mmm d, yyyy')).toBe('Jan 5, 2026');
    expect(f(d, 'dddd')).toBe('Monday');
    expect(f(d + 0.75, 'hh:mm')).toBe('18:00');
    expect(f(d + 0.75, 'h:mm AM/PM')).toBe('6:00 PM');
    expect(f(1.5, '[h]:mm')).toBe('36:00');
    expect(f(d + 14.5 / 24, 'yyyy-mm-dd hh:mm:ss')).toBe('2026-01-05 14:30:00');
  });
});

describe('typed input', () => {
  it('reads dates, times and currency as numbers', () => {
    expect(parseNumericText('2026-01-05')).toBe(dateToSerial(2026, 1, 5));
    expect(parseNumericText('2026/1/5')).toBe(dateToSerial(2026, 1, 5));
    expect(parseNumericText('12:00')).toBe(0.5);
    expect(parseNumericText('6:00 pm')).toBe(0.75);
    expect(parseNumericText('$1,200')).toBe(1200);
    expect(parseNumericText('2026-13-05')).toBeNull();
  });

  it('implies a display format', () => {
    expect(impliedFormat('2026-01-05')).toBe('yyyy-mm-dd');
    expect(impliedFormat('12%')).toBe('0%');
    expect(impliedFormat('$1,200')).toBe('$#,##0');
    expect(impliedFormat('1,200.5')).toBe('#,##0.00');
    expect(impliedFormat('hello')).toBeUndefined();
  });

  it('does date arithmetic on typed dates', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setMany('S', { A1: '2026-01-05', A2: '=A1+30', A3: '=TEXT(A2,"yyyy/mm/dd")', A4: '=DAYS("2026-03-01","2026-01-01")' });
    expect(wb.get('S', 'A3')).toBe('2026/02/04');
    expect(wb.get('S', 'A4')).toBe(59);
  });
});

describe('copying formulas', () => {
  it('moves relative references and keeps absolute ones', () => {
    expect(offsetFormula('=A1+B$2+$C3+$D$4', 2, 1)).toBe('=B3+C$2+$C5+$D$4');
    expect(offsetFormula('=SUM(A1:A3)', 1, 0)).toBe('=SUM(A2:A4)');
    expect(offsetFormula('=SUM(A:A)', 0, 2)).toBe('=SUM(C:C)');
    expect(offsetFormula('=A1', -1, 0)).toBe('=#REF!');
    expect(offsetFormula('plain', 3, 3)).toBe('plain');
  });
});

describe('sheet operations', () => {
  const book = () => {
    const wb = new Workbook();
    wb.addSheet('Data');
    wb.addSheet('Report');
    wb.setMany('Data', { A1: '10', A2: '20' });
    wb.setMany('Report', { A1: '=SUM(Data!A1:A2)', A2: "='Data'!A1*2" });
    return wb;
  };

  it('renames a sheet and every formula that names it', () => {
    const wb = book();
    wb.renameSheet('Data', 'Raw data');
    expect(wb.sheetNames()).toEqual(['Raw data', 'Report']);
    expect(wb.input('Report', 'A1')).toBe("=SUM('Raw data'!A1:A2)");
    expect(wb.get('Report', 'A1')).toBe(30);
    wb.set('Raw data', 'A1', '15');
    expect(wb.get('Report', 'A2')).toBe(30);
  });

  it('removes a sheet, leaving #REF! where formulas pointed into it', () => {
    const wb = book();
    wb.removeSheet('Data');
    expect(wb.sheetNames()).toEqual(['Report']);
    expect(wb.input('Report', 'A1')).toBe('=SUM(#REF!)');
    expect(String(wb.get('Report', 'A1'))).toBe('#REF!');
  });

  it('reorders sheets', () => {
    const wb = book();
    wb.reorderSheets(['Report', 'Data']);
    expect(wb.sheetNames()).toEqual(['Report', 'Data']);
    expect(wb.get('Report', 'A1')).toBe(30);
  });
});

describe('apostrophe', () => {
  it('keeps typed text that looks like a number', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setMany('S', { A1: "'00123", A2: '=LEN(A1)' });
    expect(wb.get('S', 'A1')).toBe('00123');
    expect(wb.get('S', 'A2')).toBe(5);
  });
});

describe('range index', () => {
  it('recalculates through narrow, wide, whole-column and cross-sheet ranges', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.setMany('A', { A1: '1', A2: '2', CZ5: '10', B1: '=SUM(A1:A2)', B3: '=SUM(A:A)', B4: '=B1*2' });
    wb.setMany('B', { A1: '=SUM(A!A1:A5)', B2: '=SUM(A!A1:CZ9)' });
    wb.set('A', 'A2', '5');
    expect(wb.get('A', 'B1')).toBe(6);
    expect(wb.get('A', 'B4')).toBe(12);
    expect(wb.get('B', 'A1')).toBe(6);
    // A1 1 + A2 5 + CZ5 10 + B1 6 + B3 6 + B4 12, through a range wider than one bucket per column.
    expect(wb.get('B', 'B2')).toBe(40);
    wb.set('A', 'CZ5', '20');
    expect(wb.get('B', 'B2')).toBe(50);
  });

  it('keeps running totals right after edits, inserts and a rename', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    const cells: Record<string, string> = {};
    for (let r = 1; r <= 50; r++) {
      cells[`A${r}`] = String(r);
      cells[`B${r}`] = `=SUM($A$1:A${r})`;
    }
    wb.setMany('S', cells);
    expect(wb.get('S', 'B50')).toBe(1275);
    wb.set('S', 'A10', '110');
    expect(wb.get('S', 'B9')).toBe(45);
    expect(wb.get('S', 'B50')).toBe(1375);
    wb.insertRows('S', 0, 1);
    wb.set('S', 'A2', '0');
    expect(wb.get('S', 'B51')).toBe(1374);
    wb.renameSheet('S', 'T');
    wb.set('T', 'A51', '0');
    expect(wb.get('T', 'B51')).toBe(1324);
  });
});

describe('defined names', () => {
  const book = () =>
    Workbook.fromJSON({
      names: { Sales: 'Orders!$B$2:$B$4', 稅率: "'Settings'!$A$1" },
      sheets: [
        { name: 'Orders', cells: { B2: '100', B3: '200', B4: '300' } },
        { name: 'Settings', cells: { A1: '0.05' } },
        { name: 'Report', cells: { A1: '=SUM(Sales)', A2: '=A1*稅率', A3: '=SUM(sales)+1', A4: '=Nope' } },
      ],
    });

  it('evaluate case-insensitively, CJK names included, and recalculate through them', () => {
    const wb = book();
    expect(wb.get('Report', 'A1')).toBe(600);
    expect(wb.get('Report', 'A2')).toBe(30);
    expect(wb.get('Report', 'A3')).toBe(601);
    expect(String(wb.get('Report', 'A4'))).toBe('#NAME?');
    wb.set('Orders', 'B3', '1200');
    expect(wb.get('Report', 'A1')).toBe(1600);
  });

  it('follow inserted rows and renamed or removed sheets', () => {
    const wb = book();
    wb.insertRows('Orders', 0, 2);
    expect(wb.namesJSON().Sales).toBe('Orders!$B$4:$B$6');
    expect(wb.get('Report', 'A1')).toBe(600);
    wb.renameSheet('Orders', 'Raw orders');
    expect(wb.namesJSON().Sales).toBe("'Raw orders'!$B$4:$B$6");
    wb.removeSheet('Settings');
    expect(String(wb.get('Report', 'A2'))).toBe('#REF!');
  });

  it('can be defined after the formulas that use them', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setMany('S', { A1: '7', B1: '=Seven*2' });
    expect(String(wb.get('S', 'B1'))).toBe('#NAME?');
    wb.setNames({ Seven: 'S!$A$1' });
    expect(wb.get('S', 'B1')).toBe(14);
    wb.set('S', 'A1', '8');
    expect(wb.get('S', 'B1')).toBe(16);
  });

  it('refuse names that are not names, or not references', () => {
    expect(Workbook.validName('A1')).toBe(false);
    expect(Workbook.validName('TRUE')).toBe(false);
    expect(Workbook.validName('R1C1')).toBe(false);
    expect(Workbook.validName('Q1_sales')).toBe(true);
    expect(() => new Workbook().setNames({ X: '=1+2' })).toThrow(/must refer/);
    expect(() => new Workbook().setNames({ X: 'A1' })).toThrow(/with its sheet/);
  });
});

describe('precedents and dependents', () => {
  const wb = Workbook.fromJSON({
    sheets: [
      { name: 'Data', cells: { A1: '10', A2: '20', B1: '=A1*2' } },
      { name: 'Sum', cells: { A1: '=SUM(Data!A1:A2)+Data!B1', A2: '=Total*2', A3: '=A1+A1' } },
    ],
    names: { Total: 'Sum!$A$1' },
  });

  it('lists what a formula reads, names resolved and repeats dropped', () => {
    expect(wb.precedents('Sum', 0, 0)).toEqual([
      { sheet: 'Data', r0: 0, c0: 1, r1: 0, c1: 1 },
      { sheet: 'Data', r0: 0, c0: 0, r1: 1, c1: 0 },
    ]);
    expect(wb.precedents('Sum', 1, 0)).toEqual([{ sheet: 'Sum', r0: 0, c0: 0, r1: 0, c1: 0 }]);
    expect(wb.precedents('Sum', 2, 0)).toHaveLength(1);
    expect(wb.precedents('Data', 0, 0)).toEqual([]);
  });

  it('lists the formulas that read a cell, through a reference, a range or a name', () => {
    expect(wb.dependents('Data', 0, 0)).toEqual([
      { sheet: 'Data', row: 0, col: 1 },
      { sheet: 'Sum', row: 0, col: 0 },
    ]);
    expect(wb.dependents('Data', 1, 0)).toEqual([{ sheet: 'Sum', row: 0, col: 0 }]);
    expect(wb.dependents('Sum', 0, 0)).toEqual([
      { sheet: 'Sum', row: 1, col: 0 },
      { sheet: 'Sum', row: 2, col: 0 },
    ]);
  });
});
