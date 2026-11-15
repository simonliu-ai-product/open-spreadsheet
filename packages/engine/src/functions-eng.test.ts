// Expected values are the worked examples in Microsoft's documentation for each function.
import { describe, expect, it } from 'vitest';
import { CellError } from './values.ts';
import { Workbook } from './workbook.ts';

function f(formula: string): unknown {
  const wb = new Workbook();
  wb.addSheet('S');
  wb.set('S', 'A1', formula);
  const v = wb.get('S', 'A1');
  return v instanceof CellError ? v.code : v;
}
const near = (formula: string, expected: number, digits = 6) => expect(f(formula) as number).toBeCloseTo(expected, digits);

describe('number bases and bits', () => {
  it('converts between bases with ten-digit two’s complement', () => {
    expect(f('=BIN2DEC("1100100")')).toBe(100);
    expect(f('=BIN2DEC("1111111111")')).toBe(-1);
    expect(f('=DEC2BIN(9,4)')).toBe('1001');
    expect(f('=DEC2BIN(-100)')).toBe('1110011100');
    expect(f('=HEX2DEC("A5")')).toBe(165);
    expect(f('=HEX2DEC("FFFFFFFF5B")')).toBe(-165);
    expect(f('=DEC2HEX(100,4)')).toBe('0064');
    expect(f('=DEC2HEX(-54)')).toBe('FFFFFFFFCA');
    expect(f('=OCT2DEC("54")')).toBe(44);
    expect(f('=DEC2OCT(58,3)')).toBe('072');
    expect(f('=BIN2HEX("11111011",4)')).toBe('00FB');
    expect(f('=DEC2BIN(9,2)')).toBe('#NUM!');
  });

  it('does bitwise arithmetic', () => {
    expect([f('=BITAND(1,5)'), f('=BITOR(23,10)'), f('=BITXOR(5,3)'), f('=BITLSHIFT(4,2)'), f('=BITRSHIFT(13,2)')]).toEqual([1, 31, 6, 16, 3]);
    expect([f('=DELTA(5,4)'), f('=DELTA(5,5)'), f('=GESTEP(5,4)'), f('=GESTEP(-4,-5)')]).toEqual([0, 1, 1, 1]);
  });

  it('writes Roman numerals and other bases', () => {
    expect(f('=ROMAN(499)')).toBe('CDXCIX');
    expect(f('=ARABIC("LVII")')).toBe(57);
    expect(f('=BASE(7,2)')).toBe('111');
    expect(f('=BASE(100,16,4)')).toBe('0064');
    expect(f('=DECIMAL("FF",16)')).toBe(255);
  });
});

describe('complex numbers', () => {
  it('builds and reads them', () => {
    expect(f('=COMPLEX(3,4)')).toBe('3+4i');
    expect(f('=COMPLEX(3,4,"j")')).toBe('3+4j');
    expect(f('=COMPLEX(0,1)')).toBe('i');
    expect(f('=IMREAL("6-9i")')).toBe(6);
    expect(f('=IMAGINARY("3+4i")')).toBe(4);
    expect(f('=IMAGINARY("-i")')).toBe(-1);
    expect(f('=IMABS("5+12i")')).toBe(13);
    expect(f('=IMCONJUGATE("3+4i")')).toBe('3-4i');
    near('=IMARGUMENT("3+4i")', 0.92729522);
  });

  it('does arithmetic', () => {
    expect(f('=IMSUM("3+4i","5-3i")')).toBe('8+i');
    expect(f('=IMSUB("13+4i","5+3i")')).toBe('8+i');
    expect(f('=IMPRODUCT("3+4i","5-3i")')).toBe('27+11i');
    expect(f('=IMDIV("-238+240i","10+24i")')).toBe('5+12i');
    expect(f('=IMSQRT("1+i")')).toBe('1.09868411346781+0.455089860562227i');
  });
});

describe('units', () => {
  it('converts within a kind and refuses across kinds', () => {
    near('=CONVERT(1,"lbm","kg")', 0.45359237, 8);
    expect(f('=CONVERT(68,"F","C")')).toBeCloseTo(20, 10);
    expect(f('=CONVERT(100,"ft","m")')).toBe(30.48);
    expect(f('=CONVERT(6,"tsp","tbs")')).toBe(2);
    expect(f('=CONVERT(2.5,"ft","sec")')).toBe('#N/A');
  });
});

describe('math and statistics', () => {
  it('counts and series', () => {
    expect(f('=MULTINOMIAL(2,3,4)')).toBe(1260);
    expect(f('=COMBINA(4,3)')).toBe(20);
    expect(f('=PERMUTATIONA(3,2)')).toBe(9);
    near('=GAMMALN(4)', 1.791759469);
    near('=GAMMA(2.5)', 1.329340388);
    expect(f('=GAMMA(-1)')).toBe('#NUM!');
    near('=FISHER(0.75)', 0.972955075);
    near('=FISHERINV(0.972955)', 0.75);
  });

  it('shapes of a distribution', () => {
    near('=KURT(3,4,5,2,3,4,5,6,4,7)', -0.151799637);
    near('=SKEW(3,4,5,2,3,4,5,6,4,7)', 0.359543071);
    near('=STDEVA(1345,1301,1368,1322,1310,1370,1318,1350,1303,1299)', 27.46391572);
    near('=ERF(0.745)', 0.707928920);
    near('=ERFC(1)', 0.157299207);
  });
});

describe('distributions', () => {
  it('normal, lognormal, exponential, Poisson, binomial', () => {
    near('=NORM.INV(0.908789,40,1.5)', 42.000002, 5);
    near('=NORM.S.INV(0.908789)', 1.3333347, 5);
    near('=LOGNORM.DIST(4,3.5,1.2,TRUE)', 0.0390836);
    near('=LOGNORM.INV(0.039084,3.5,1.2)', 4.0000252, 4);
    near('=EXPON.DIST(0.2,10,TRUE)', 0.86466472);
    near('=EXPON.DIST(0.2,10,FALSE)', 1.35335283);
    near('=POISSON.DIST(2,5,TRUE)', 0.124652);
    near('=POISSON.DIST(2,5,FALSE)', 0.084224);
    near('=BINOM.DIST(6,10,0.5,FALSE)', 0.2050781);
  });

  it('t, chi-squared, F, gamma and beta', () => {
    near('=T.DIST(60,1,TRUE)', 0.99469533);
    near('=T.DIST(8,3,FALSE)', 0.00073691);
    near('=T.DIST.2T(1.959999998,60)', 0.054645);
    near('=T.DIST.RT(1.959999998,60)', 0.027322);
    near('=T.INV.2T(0.546449,60)', 0.606533);
    near('=T.INV(0.75,2)', 0.8164966);
    near('=CHISQ.DIST(0.5,1,TRUE)', 0.52049988);
    near('=CHISQ.DIST.RT(18.307,10)', 0.0500006);
    near('=CHISQ.INV.RT(0.050001,10)', 18.306973, 4);
    near('=F.DIST(15.2069,6,4,TRUE)', 0.99, 4);
    near('=F.DIST.RT(15.2069,6,4)', 0.01, 4);
    near('=GAMMA.DIST(10.00001131,9,2,TRUE)', 0.068094);
    near('=BETA.DIST(2,8,10,TRUE,1,3)', 0.6854706);
    near('=CONFIDENCE.NORM(0.05,2.5,50)', 0.692952);
    near('=CONFIDENCE.T(0.05,1,50)', 0.284196855);
  });
});

describe('information and web', () => {
  it('ISREF and ENCODEURL', () => {
    expect(f('=ISREF(B2)')).toBe(true);
    expect(f('=ISREF("B2")')).toBe(false);
    expect(f('=ENCODEURL("a b&c")')).toBe('a%20b%26c');
  });
});
