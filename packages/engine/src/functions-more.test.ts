// Expected values are the worked examples in Microsoft's documentation for each function.
import { describe, expect, it } from 'vitest';
import { CellError } from './values.ts';
import { Workbook } from './workbook.ts';

function f(formula: string, cells: Record<string, string> = {}): unknown {
  const wb = new Workbook();
  wb.addSheet('S');
  wb.setMany('S', { ...cells, ZZ1: formula });
  const v = wb.get('S', 'ZZ1');
  return v instanceof CellError ? v.code : v;
}
const near = (formula: string, expected: number, cells?: Record<string, string>, digits = 6) => expect(f(formula, cells) as number).toBeCloseTo(expected, digits);

describe('math', () => {
  it('rounds to multiples', () => {
    expect(f('=CEILING(2.5,1)')).toBe(3);
    expect(f('=CEILING(-2.5,-2)')).toBe(-4);
    expect(f('=CEILING(-2.5,2)')).toBe(-2);
    expect(f('=CEILING(0.234,0.01)')).toBe(0.24);
    expect(f('=FLOOR(3.7,2)')).toBe(2);
    expect(f('=FLOOR(-2.5,-2)')).toBe(-2);
    expect(f('=FLOOR(0.234,0.01)')).toBe(0.23);
    expect(f('=CEILING.MATH(24.3,5)')).toBe(25);
    expect(f('=CEILING.MATH(-5.5,2)')).toBe(-4);
    expect(f('=CEILING.MATH(-5.5,2,-1)')).toBe(-6);
    expect(f('=FLOOR.MATH(-8.1,2)')).toBe(-10);
    expect(f('=FLOOR.MATH(-5.5,2,-1)')).toBe(-4);
    expect(f('=MROUND(10,3)')).toBe(9);
    expect(f('=MROUND(-10,-3)')).toBe(-9);
    expect(f('=MROUND(1.3,0.2)')).toBe(1.4);
    expect(f('=MROUND(5,-2)')).toBe('#NUM!');
    expect([f('=EVEN(1.5)'), f('=EVEN(3)'), f('=EVEN(-1)')]).toEqual([2, 4, -2]);
    expect([f('=ODD(1.5)'), f('=ODD(3)'), f('=ODD(2)'), f('=ODD(-1)'), f('=ODD(-2)')]).toEqual([3, 3, 3, -1, -3]);
  });

  it('does logs, factorials, combinations and divisors', () => {
    expect(f('=LOG(8,2)')).toBe(3);
    near('=LOG10(86)', 1.934498451);
    near('=LN(86)', 4.454347296);
    expect(f('=FACT(5)')).toBe(120);
    expect(f('=COMBIN(8,2)')).toBe(28);
    expect(f('=PERMUT(100,3)')).toBe(970200);
    expect(f('=GCD(24,36)')).toBe(12);
    expect(f('=LCM(24,36)')).toBe(72);
    expect(f('=QUOTIENT(-10,3)')).toBe(-3);
    expect(f('=SUMSQ(3,4)')).toBe(25);
    near('=DEGREES(PI())', 180);
    near('=SIN(PI()/2)', 1);
    expect(f('=ASIN(2)')).toBe('#NUM!');
  });

  it('runs SUBTOTAL by its function code', () => {
    const cells = { A1: '120', A2: '10', A3: '150', A4: '23' };
    expect(f('=SUBTOTAL(9,A1:A4)', cells)).toBe(303);
    expect(f('=SUBTOTAL(1,A1:A4)', cells)).toBe(75.75);
    expect(f('=SUBTOTAL(104,A1:A4)', cells)).toBe(150);
  });
});

describe('statistics', () => {
  const data = { A1: '1345', A2: '1301', A3: '1368', A4: '1322', A5: '1310', A6: '1370', A7: '1318', A8: '1350', A9: '1303', A10: '1299' };
  it('spreads', () => {
    near('=STDEV.P(A1:A10)', 26.05455814, data);
    near('=VAR.S(A1:A10)', 754.2666667, data);
    near('=VAR.P(A1:A10)', 678.84, data);
    near('=STDEV(A1:A10)', 27.46392, data, 5);
  });

  it('modes, percentiles, quartiles and ranks', () => {
    expect(f('=MODE.SNGL(5.6,4,4,3,2,4)')).toBe(4);
    near('=PERCENTILE.INC({1,3,2,4},0.3)', 1.9);
    near('=QUARTILE.INC({1,2,4,7,8,9,10,12},1)', 3.5);
    near('=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.25)', 2.5);
    const r = { A1: '7', A2: '3.5', A3: '3.5', A4: '1', A5: '2' };
    expect(f('=RANK.EQ(A2,A1:A5,1)', r)).toBe(3);
    expect(f('=RANK.EQ(A1,A1:A5,1)', r)).toBe(5);
    expect(f('=RANK.AVG(94,{89,88,92,101,94,97,95})')).toBe(4);
  });

  it('correlates and regresses', () => {
    near('=CORREL({3,2,4,5,6},{9,7,12,15,17})', 0.997054486);
    near('=SLOPE({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', 0.305555556);
    near('=INTERCEPT({2,3,9,1,8},{6,5,11,7,5})', 0.0483871);
    near('=FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40})', 10.607253);
    near('=GEOMEAN(4,5,8,7,11,4,3)', 5.476987);
    near('=HARMEAN(4,5,8,7,11,4,3)', 5.028376);
    near('=AVEDEV(4,5,6,7,5,4,3)', 1.020408163);
    expect(f('=DEVSQ(4,5,8,7,11,4,3)')).toBe(48);
    near('=NORM.DIST(42,40,1.5,TRUE)', 0.9087888);
    near('=NORM.DIST(42,40,1.5,FALSE)', 0.10934005);
  });

  it('averages, maxima and minima under criteria', () => {
    const cells = { A1: '89', A2: '93', A3: '96', A4: '85', A5: '91', A6: '88', B1: '1', B2: '2', B3: '2', B4: '3', B5: '1', B6: '1' };
    expect(f('=MAXIFS(A1:A6,B1:B6,1)', cells)).toBe(91);
    expect(f('=MINIFS(A1:A6,B1:B6,1)', cells)).toBe(88);
    near('=AVERAGEIFS(A1:A6,B1:B6,">1")', 91.3333333, cells);
  });
});

describe('text', () => {
  it('handles codes and cleaning', () => {
    expect(f('=CHAR(65)')).toBe('A');
    expect(f('=CODE("A")')).toBe(65);
    expect(f('=CLEAN(CHAR(9)&"Monthly report"&CHAR(10))')).toBe('Monthly report');
    expect(f('=UNICHAR(9731)')).toBe('☃');
    expect(f('=T(123)')).toBe('');
    expect(f('=N(TRUE)')).toBe(1);
    expect(f('=NUMBERVALUE("2.500,27",",",".")')).toBe(2500.27);
  });

  it('formats money and fixed decimals', () => {
    expect(f('=DOLLAR(1234.567,2)')).toBe('$1,234.57');
    expect(f('=DOLLAR(-1234.567,-2)')).toBe('($1,200)');
    expect(f('=FIXED(1234.567,1)')).toBe('1,234.6');
    expect(f('=FIXED(-1234.567,-1)')).toBe('-1,230');
    expect(f('=FIXED(44.332)')).toBe('44.33');
  });

  it('splits around a delimiter', () => {
    const s = '"Red riding hood\'s, red hood"';
    expect(f(`=TEXTBEFORE(${s},"hood")`)).toBe('Red riding ');
    expect(f(`=TEXTAFTER(${s},"hood")`)).toBe("'s, red hood");
    expect(f(`=TEXTBEFORE(${s},"hood",2)`)).toBe("Red riding hood's, red ");
    expect(f(`=TEXTAFTER(${s},"hood",-1)`)).toBe('');
    expect(f(`=TEXTAFTER(${s},"x")`)).toBe('#N/A');
  });
});

describe('dates', () => {
  it('times', () => {
    expect(f('=TIME(12,0,0)')).toBe(0.5);
    expect(f('=HOUR(0.75)')).toBe(18);
    expect(f('=MINUTE("12:45")')).toBe(45);
    expect(f('=DATEVALUE("2026-01-05")')).toBe(46027);
  });

  it('weeks and working days', () => {
    expect(f('=WEEKNUM(DATE(2012,3,9))')).toBe(10);
    expect(f('=WEEKNUM(DATE(2012,3,9),2)')).toBe(11);
    expect(f('=ISOWEEKNUM(DATE(2012,3,9))')).toBe(10);
    expect(f('=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))')).toBe(110);
    expect(f('=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),DATE(2012,11,22))')).toBe(109);
    expect(f('=WORKDAY(DATE(2008,10,1),151)')).toBe(39933);
    expect(f('=WORKDAY(DATE(2008,10,1),151,A1:A3)', { A1: '2008-11-26', A2: '2008-12-04', A3: '2009-01-21' })).toBe(39938);
  });

  it('differences and year fractions', () => {
    expect(f('=DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Y")')).toBe(2);
    expect(f('=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"D")')).toBe(440);
    expect(f('=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YD")')).toBe(75);
    expect(f('=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"MD")')).toBe(14);
    expect(f('=DAYS360(DATE(2011,1,1),DATE(2011,12,31))')).toBe(360);
    expect(f('=DAYS360(DATE(2011,1,30),DATE(2011,2,1))')).toBe(1);
    near('=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))', 0.58055556);
    near('=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)', 0.57650273);
    near('=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)', 0.57808219);
  });
});

describe('lookup and information', () => {
  it('matches and builds addresses', () => {
    expect(f('=XMATCH("Gra*",{"Apple","Grape","Pear","Banana","Cherry"},2)')).toBe(2);
    expect(f('=XMATCH("Gra?",{"Apple","Grape","Pear","Banana","Cherry"},2)')).toBe('#N/A');
    expect(f('=XMATCH(25,{10,20,30,40},1)')).toBe(3);
    expect(f('=XMATCH(25,{10,20,30,40},-1)')).toBe(2);
    expect(f('=LOOKUP(4.19,A1:A5,B1:B5)', { A1: '4.14', A2: '4.19', A3: '5.17', A4: '5.77', A5: '6.39', B1: 'red', B2: 'orange', B3: 'yellow', B4: 'green', B5: 'blue' })).toBe('orange');
    expect(f('=ADDRESS(2,3)')).toBe('$C$2');
    expect(f('=ADDRESS(2,3,2)')).toBe('C$2');
    expect(f('=ADDRESS(2,3,2,FALSE)')).toBe('R2C[3]');
    expect(f('=ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")')).toBe("'[Book1]Sheet1'!R2C3");
  });

  it('types and errors', () => {
    expect([f('=TYPE(1)'), f('=TYPE("a")'), f('=TYPE(TRUE)'), f('=TYPE(1/0)'), f('=TYPE({1,2})')]).toEqual([1, 2, 4, 16, 64]);
    expect(f('=ERROR.TYPE(1/0)')).toBe(2);
    expect([f('=ISEVEN(-1)'), f('=ISODD(5)'), f('=ISERR(NA())'), f('=ISNONTEXT(1)')]).toEqual([false, true, false, true]);
  });
});

describe('finance', () => {
  it('periods, rates and payment parts', () => {
    near('=NPER(0.12/12,-100,-1000,10000,1)', 59.6738657);
    near('=RATE(4*12,-200,8000)', 0.00770147, undefined, 8);
    near('=IPMT(0.1/12,1,3*12,8000)', -66.6666667);
    near('=PPMT(0.1/12,1,2*12,2000)', -75.62, undefined, 2);
  });

  it('cash flows', () => {
    near('=IRR({-70000,12000,15000,18000,21000,26000})', 0.086630948);
    const flows = { A1: '-10000', A2: '2750', A3: '4250', A4: '3250', A5: '2750', B1: '2008-01-01', B2: '2008-03-01', B3: '2008-10-30', B4: '2009-02-15', B5: '2009-04-01' };
    near('=XNPV(0.09,A1:A5,B1:B5)', 2086.647602, flows, 4);
    near('=XIRR(A1:A5,B1:B5)', 0.373362535, flows);
  });

  it('depreciation and conversions', () => {
    expect(f('=SLN(30000,7500,10)')).toBe(2250);
    near('=SYD(30000,7500,10,1)', 4090.909091);
    expect(f('=DDB(2400,300,120,1,2)')).toBe(40);
    near('=DDB(2400,300,10*365,1)', 1.315068493);
    near('=EFFECT(0.0525,4)', 0.053542667);
    near('=NOMINAL(0.053543,4)', 0.05250032);
  });
});

describe('database functions', () => {
  const db = {
    A1: 'Tree', B1: 'Height', C1: 'Age', D1: 'Yield', E1: 'Profit',
    A2: 'Apple', B2: '18', C2: '20', D2: '14', E2: '105',
    A3: 'Pear', B3: '12', C3: '12', D3: '10', E3: '96',
    A4: 'Cherry', B4: '13', C4: '14', D4: '9', E4: '105',
    A5: 'Apple', B5: '14', C5: '15', D5: '10', E5: '75',
    A6: 'Pear', B6: '9', C6: '8', D6: '8', E6: '76.8',
    A7: 'Apple', B7: '8', C7: '9', D7: '6', E7: '45',
    G1: 'Tree', H1: 'Height', I1: 'Height',
    G2: 'Apple', H2: '>10', I2: '<16',
    G3: 'Pear',
  };
  it('follows Microsoft’s orchard example', () => {
    expect(f('=DCOUNT(A1:E7,"Age",G1:I2)', db)).toBe(1);
    expect(f('=DMAX(A1:E7,"Profit",G1:I3)', db)).toBe(96);
    expect(f('=DMIN(A1:E7,"Profit",G1:I2)', db)).toBe(75);
    expect(f('=DSUM(A1:E7,"Profit",G1:G2)', db)).toBe(225);
    expect(f('=DGET(A1:E7,"Yield",G1:I2)', db)).toBe(10);
    expect(f('=DGET(A1:E7,"Yield",G1:G2)', db)).toBe('#NUM!');
  });
});

describe('HYPERLINK', () => {
  it('shows the friendly name, or the link when there is none (Microsoft examples)', () => {
    const wb = Workbook.fromJSON({
      sheets: [{ name: 'S', cells: { A1: 'https://example.com/report', A2: '=HYPERLINK(A1, "Click for report")', A3: '=HYPERLINK("https://example.com")', A4: '=HYPERLINK(A1, 1/0)' } }],
    });
    expect(wb.value('S', 1, 0)).toBe('Click for report');
    expect(wb.value('S', 2, 0)).toBe('https://example.com');
    expect(String(wb.value('S', 3, 0))).toBe('#DIV/0!');
  });
});
