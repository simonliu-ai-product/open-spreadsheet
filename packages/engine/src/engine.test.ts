import { describe, expect, it } from 'vitest';
import { parseFormula } from './parse.ts';
import { serialize } from './serialize.ts';
import { CellError, type Scalar } from './values.ts';
import { Workbook } from './workbook.ts';

/** Evaluates one formula in an otherwise empty workbook (or one prepared by `setup`). */
function calc(formula: string, setup?: (wb: Workbook) => void): Scalar {
  const wb = new Workbook();
  wb.addSheet('Sheet1');
  wb.now = () => new Date(2026, 10, 15, 12, 0, 0);
  setup?.(wb);
  wb.set('Sheet1', 'Z100', formula);
  return wb.get('Sheet1', 'Z100');
}
const code = (v: Scalar) => (v instanceof CellError ? v.code : v);
const close = (v: Scalar, expected: number, digits = 2) => expect(Number((v as number).toFixed(digits))).toBe(expected);

describe('operators follow Excel precedence and coercion', () => {
  it.each([
    ['=-2^2', 4], // negation binds tighter than ^
    ['=2^3^2', 64], // ^ is left-associative
    ['=1+2*3', 7],
    ['=(1+2)*3', 9],
    ['=10%', 0.1],
    ['=-50%*2', -1],
    ['="3"+1', 4],
    ['="1,000"+1', 1001],
    ['=TRUE+TRUE', 2],
    ['=1&2', '12'],
    ['="a"="A"', true], // text compares case-insensitively
    ['=0.1+0.2=0.3', true], // 15-significant-digit comparison
    ['=1<"a"', true], // numbers sort before text
    ['="Z"<TRUE', true], // text sorts before booleans
  ])('%s → %s', (f, expected) => expect(calc(f)).toBe(expected));

  it.each([
    ['=1/0', '#DIV/0!'],
    ['=0^0', '#NUM!'],
    ['="abc"+1', '#VALUE!'],
    ['=NOSUCHFUNCTION(1)', '#NAME?'],
    ['=#N/A+1', '#N/A'],
    ['=SQRT(-1)', '#NUM!'],
  ])('%s → %s', (f, expected) => expect(code(calc(f))).toBe(expected));
});

describe('math', () => {
  it.each([
    ['=ROUND(2.675,2)', 2.68], // JavaScript's toFixed gives 2.67
    ['=ROUND(-2.5,0)', -3], // half away from zero
    ['=ROUND(1234.567,-2)', 1200],
    ['=ROUNDUP(3.2,0)', 4],
    ['=ROUNDDOWN(-3.7,0)', -3],
    ['=INT(-8.9)', -9],
    ['=TRUNC(-8.9)', -8],
    ['=MOD(-3,2)', 1], // the result takes the divisor's sign
    ['=MOD(3,-2)', -1],
    ['=SUM("3",TRUE,1)', 5], // typed directly, text and TRUE count
    ['=SUMPRODUCT({1,2,3},{4,5,6})', 32],
    ['=PRODUCT(2,3,4)', 24],
    ['=POWER(2,10)', 1024],
  ])('%s → %s', (f, expected) => expect(calc(f)).toBe(expected));

  it('skips text and blanks inside a range', () => {
    const setup = (wb: Workbook) => wb.setMany('Sheet1', { A1: 10, A2: 'text', A4: 20, A5: 'TRUE' });
    expect(calc('=SUM(A1:A5)', setup)).toBe(30);
    expect(calc('=AVERAGE(A1:A5)', setup)).toBe(15);
    expect(calc('=COUNT(A1:A5)', setup)).toBe(2);
    expect(calc('=COUNTA(A1:A5)', setup)).toBe(4);
    expect(calc('=COUNTBLANK(A1:A5)', setup)).toBe(1);
    expect(calc('=MAX(A1:A5)', setup)).toBe(20);
    expect(code(calc('=AVERAGE(A2:A3)', setup))).toBe('#DIV/0!');
  });
});

describe('criteria functions', () => {
  const setup = (wb: Workbook) =>
    wb.setMany('Sheet1', {
      A1: 'apple', B1: 1, C1: 'east',
      A2: 'banana', B2: 3, C2: 'west',
      A3: 'apricot', B3: 5, C3: 'east',
      A4: 'cherry', B4: 7, C4: 'east',
    });
  it.each([
    ['=COUNTIF(B1:B5,">2")', 3],
    ['=COUNTIF(A1:A4,"ap*")', 2],
    ['=COUNTIF(A1:A4,"?pple")', 1],
    ['=COUNTIF(A1:A5,"")', 1],
    ['=COUNTIF(C1:C4,"<>east")', 1],
    ['=SUMIF(C1:C4,"east",B1:B4)', 13],
    ['=SUMIF(B1:B4,">=5")', 12],
    ['=SUMIFS(B1:B4,C1:C4,"east",A1:A4,"a*")', 6],
    ['=COUNTIFS(C1:C4,"east",B1:B4,">1")', 2],
    ['=AVERAGEIF(C1:C4,"east",B1:B4)', 13 / 3],
  ])('%s → %s', (f, expected) => expect(calc(f, setup)).toBe(expected));
});

describe('lookup', () => {
  const setup = (wb: Workbook) =>
    wb.setMany('Sheet1', {
      A1: 1, B1: 'one', A2: 3, B2: 'three', A3: 5, B3: 'five', A4: 7, B4: 'seven',
      D1: 'Q1', E1: 'Q2', F1: 'Q3', D2: 100, E2: 200, F2: 300,
    });
  it.each([
    ['=VLOOKUP(5,A1:B4,2,FALSE)', 'five'],
    ['=VLOOKUP(4,A1:B4,2)', 'three'], // approximate: the last key <= 4
    ['=VLOOKUP(4,A1:B4,2,FALSE)', '#N/A'],
    ['=VLOOKUP(5,A1:B4,3,FALSE)', '#REF!'],
    ['=HLOOKUP("Q2",D1:F2,2,FALSE)', 200],
    ['=MATCH(5,A1:A4,0)', 3],
    ['=MATCH(6,A1:A4,1)', 3],
    ['=INDEX(A1:B4,2,2)', 'three'],
    ['=INDEX(A1:B4,MATCH(7,A1:A4,0),2)', 'seven'],
    ['=XLOOKUP(3,A1:A4,B1:B4)', 'three'],
    ['=XLOOKUP(9,A1:A4,B1:B4,"none")', 'none'],
    ['=XLOOKUP(4,A1:A4,B1:B4,,-1)', 'three'], // next smaller
    ['=XLOOKUP(4,A1:A4,B1:B4,,1)', 'five'], // next larger
    ['=CHOOSE(2,"a","b","c")', 'b'],
    ['=ROWS(A1:B4)*10+COLUMNS(A1:B4)', 42],
    ['=ROW(C7)+COLUMN(C7)', 10],
  ])('%s → %s', (f, expected) => expect(code(calc(f, setup))).toBe(expected));
});

describe('text', () => {
  it.each([
    ['=LEFT("中文字",2)', '中文'],
    ['=LEN("中文")', 2],
    ['=MID("open-spreadsheet",6,6)', 'spread'],
    ['=RIGHT("abc",5)', 'abc'],
    ['=PROPER("hello world")', 'Hello World'],
    ['=TRIM("  a   b ")', 'a b'],
    ['=SUBSTITUTE("a-b-c","-","+",2)', 'a-b+c'],
    ['=REPLACE("abcdef",2,3,"X")', 'aXef'],
    ['=FIND("b","abcb")', 2],
    ['=SEARCH("B*","abc")', 2],
    ['=TEXTJOIN("/",TRUE,"a","","b")', 'a/b'],
    ['=CONCAT("x",1,TRUE)', 'x1TRUE'],
    ['=EXACT("a","A")', false],
    ['=VALUE("12.5")', 12.5],
    ['=TEXT(1234.567,"#,##0.00")', '1,234.57'],
    ['=TEXT(0.256,"0.0%")', '25.6%'],
    ['=TEXT(DATE(2026,11,15),"yyyy-mm-dd")', '2026-11-15'],
    ['=1/3&""', '0.333333333333333'], // 15 significant digits
  ])('%s → %s', (f, expected) => expect(calc(f)).toBe(expected));
});

describe('logic', () => {
  it.each([
    ['=IF(TRUE,1,1/0)', 1], // the branch not taken is never evaluated
    ['=IF(FALSE,1)', false],
    ['=IFERROR(1/0,"x")', 'x'],
    ['=IFNA(NA(),"missing")', 'missing'],
    ['=AND(TRUE,1,2>1)', true],
    ['=OR(FALSE,0)', false],
    ['=IFS(1>2,"a",2>1,"b")', 'b'],
    ['=SWITCH(2,1,"one",2,"two","other")', 'two'],
    ['=ISBLANK(A1)', true],
    ['=ISNUMBER("1")', false],
  ])('%s → %s', (f, expected) => expect(calc(f)).toBe(expected));
});

describe('dates: Excel 1900 system, leap-year bug included', () => {
  it.each([
    ['=DATE(2023,1,1)', 44927],
    ['=DATE(2023,3,15)', 45000],
    ['=DATE(1900,2,28)', 59],
    ['=DATE(1900,3,1)', 61], // serial 60 is the 29 February 1900 that never existed
    ['=DAY(60)*100+MONTH(60)', 2902],
    ['=DATE(2026,14,1)=DATE(2027,2,1)', true], // months overflow into years
    ['=WEEKDAY(DATE(2023,3,15))', 4], // a Wednesday
    ['=WEEKDAY(DATE(2023,3,15),2)', 3],
    ['=EDATE(DATE(2023,1,31),1)', 44985], // clamped to 28 February
    ['=EOMONTH(DATE(2023,1,15),1)', 44985],
    ['=YEAR(45000)', 2023],
    ['=DAYS(DATE(2026,12,25),DATE(2026,11,15))', 40],
    ['=TODAY()', 46341],
  ])('%s → %s', (f, expected) => expect(calc(f)).toBe(expected));
});

describe('finance (Microsoft documentation examples)', () => {
  it('PMT, FV, PV and NPV', () => {
    close(calc('=PMT(0.08/12,10,10000)'), -1037.03);
    close(calc('=FV(0.06/12,10,-200,-500,1)'), 2581.4);
    close(calc('=PV(0.08/12,12*20,500,,0)'), -59777.15);
    close(calc('=NPV(0.1,-10000,3000,4200,6800)'), 1188.44);
  });
});

describe('the workbook recalculates what depends on a change', () => {
  it('follows chains, ranges and other sheets', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.addSheet('My Sheet');
    wb.setMany('Sheet1', { A1: 1, A2: '=A1*2', A3: '=A2+1', B1: '=SUM(A1:A3)' });
    wb.set('My Sheet', 'A1', "=Sheet1!A3*10");
    wb.set('Sheet1', 'C1', "='My Sheet'!A1+1");
    expect(wb.get('Sheet1', 'B1')).toBe(6);
    wb.set('Sheet1', 'A1', '5');
    expect(wb.get('Sheet1', 'A3')).toBe(11);
    expect(wb.get('Sheet1', 'B1')).toBe(26);
    expect(wb.get('My Sheet', 'A1')).toBe(110);
    expect(wb.get('Sheet1', 'C1')).toBe(111);
  });

  it('reports a circular reference instead of looping', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.setMany('Sheet1', { A1: '=B1+1', B1: '=A1+1' });
    expect(code(wb.get('Sheet1', 'A1'))).toBe('#CIRC!');
    expect(code(wb.get('Sheet1', 'B1'))).toBe('#CIRC!');
  });

  it('reads a whole column as its used rows', () => {
    expect(calc('=SUM(A:A)', (wb) => wb.setMany('Sheet1', { A1: 1, A2: 2, A500: 3 }))).toBe(6);
  });
});

describe('inserting and deleting rows rewrites formulas, as Excel does', () => {
  it('moves cells and grows ranges on insert', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.setMany('Sheet1', { A1: 1, A2: 2, A3: 3, A4: '=SUM(A1:A3)', B1: '=$A$3*2' });
    wb.insertRows('Sheet1', 1, 1); // a blank row 2
    expect(wb.input('Sheet1', 'A5')).toBe('=SUM(A1:A4)');
    expect(wb.input('Sheet1', 'B1')).toBe('=$A$4*2');
    wb.set('Sheet1', 'A2', '10');
    expect(wb.get('Sheet1', 'A5')).toBe(16);
  });

  it('shrinks ranges and turns references to deleted cells into #REF!', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.setMany('Sheet1', { A1: 1, A2: 2, A3: 3, A4: '=SUM(A1:A3)', B1: '=A2*2' });
    wb.insertRows('Sheet1', 1, -1); // delete row 2
    expect(wb.input('Sheet1', 'A3')).toBe('=SUM(A1:A2)');
    expect(wb.get('Sheet1', 'A3')).toBe(4);
    expect(wb.input('Sheet1', 'B1')).toBe('=#REF!*2');
    expect(code(wb.get('Sheet1', 'B1'))).toBe('#REF!');
  });
});

describe('parse → serialize round trip', () => {
  it.each([
    '=SUM($A$1:B2,C:C,3:3)',
    "='My Sheet'!A1+Other!$B$2",
    '=IF(A1>=10,"big","small")&"!"',
    '=-(1+2)*3^2',
    '=(1-2)-3',
    '=1-(2-3)',
    '=VLOOKUP(A1,{1,"a";2,"b"},2,FALSE)',
    '=50%',
  ])('%s', (f) => expect(serialize(parseFormula(f))).toBe(f));
});

describe('saving and loading', () => {
  it('round-trips a workbook through JSON, formulas recomputed on load', () => {
    const wb = new Workbook();
    wb.addSheet('預算');
    wb.setMany('預算', { A1: '項目', B1: '金額', A2: '房租', B2: 18000, A3: '餐費', B3: 9000, B4: '=SUM(B2:B3)' });
    const json = JSON.parse(JSON.stringify(wb.toJSON()));
    expect(json.sheets[0].cells).toEqual({ A1: '項目', B1: '金額', A2: '房租', B2: '18000', A3: '餐費', B3: '9000', B4: '=SUM(B2:B3)' });
    const again = Workbook.fromJSON(json);
    expect(again.get('預算', 'B4')).toBe(27000);
    again.setCell('預算', 2, 1, '10000');
    expect(again.value('預算', 3, 1)).toBe(28000);
  });
});

describe('one-value functions given a range work cell by cell, as Excel does', () => {
  const sales = (wb: Workbook) =>
    wb.setMany('Sheet1', {
      A1: '2026-01-05', B1: 100, C1: 'apple',
      A2: '2026-01-20', B2: 200, C2: 'kiwi',
      A3: '2026-02-03', B3: 50, C3: 'banana',
      A4: '2026-03-15', B4: 7, C4: '',
    });
  it.each([
    ['=SUMPRODUCT((MONTH(A1:A4)=1)*B1:B4)', 300],
    ['=SUMPRODUCT((YEAR(A1:A4)=2026)*(DAY(A1:A4)>10)*B1:B4)', 207],
    ['=SUMPRODUCT(LEN(C1:C4))', 15],
    ['=SUMPRODUCT(--ISNUMBER(SEARCH("an",C1:C4)))', 1],
    ['=SUMPRODUCT(--ISBLANK(C1:C4))', 1], // typing nothing leaves C4 empty
    ['=SUM(IF(B1:B4>60,B1:B4,0))', 300],
    ['=SUM(IFERROR(1/(B1:B4-50),0))', 0.0034],
    ['=SUMPRODUCT(ROUND(B1:B4/3,0))', 119],
    ['=SUM(LEN({"a","bb","ccc"}))', 6],
    ['=MONTH(A3:A4)', 2], // in one cell, the first value until spilling lands
  ])('%s', (f, expected) => {
    const v = calc(f, sales);
    if (typeof expected === 'number' && !Number.isInteger(expected)) close(v, expected, 4);
    else expect(v).toBe(expected);
  });
  it('still evaluates only the IF branch it takes', () => {
    expect(calc('=IF(TRUE,1,1/0)')).toBe(1);
    expect(code(calc('=IF(A1:A2>0,1/0,2)', (wb) => wb.setMany('Sheet1', { A1: 1, A2: -1 })))).toBe('#DIV/0!');
    expect(calc('=SUM(IFERROR(IF(A1:A2>0,1/0,2),5))', (wb) => wb.setMany('Sheet1', { A1: 1, A2: -1 }))).toBe(7);
  });
});
