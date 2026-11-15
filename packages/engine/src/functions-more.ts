/**
 * The second batch of Excel functions: more math and trigonometry, statistics,
 * text, dates, lookup, information, finance, the D-functions and the names
 * older workbooks use. Registered into the same table as the first batch.
 */
import { dateToSerial, serialToDate, weekday } from './dates.ts';
import { formatValue } from './format.ts';
import {
  asMatrix,
  type CallContext,
  criterion,
  FUNCTIONS,
  type FunctionImpl,
  findApprox,
  lineOf,
  matchingPositions,
  num,
  numbers,
  pairsFrom,
  pmt,
  roundWith,
  scalarArg,
  sum,
  text,
  wildcardRegex,
} from './functions.ts';
import { columnName } from './address.ts';
import type { Node } from './parse.ts';
import { compare, err, fail, flatten, isError, isMatrix, parseNumericText, roundTo15, type Scalar, toBoolean, type Value } from './values.ts';

const mean = (xs: number[]) => sum(xs) / xs.length;
const need = (xs: number[], n = 1) => (xs.length < n ? fail(err.div0) : xs);
function variance(xs: number[], sample: boolean): number {
  need(xs, sample ? 2 : 1);
  const m = mean(xs);
  return sum(xs.map((x) => (x - m) ** 2)) / (xs.length - (sample ? 1 : 0));
}
/** Numbers in one argument, as a range sees them (text and blanks skipped). */
function rangeNumbers(ctx: CallContext, node: Node | undefined): number[] {
  if (!node) return [];
  return numbers(ctx, [node]);
}
/** Cells of a range flattened, keeping positions (for pairwise statistics). */
function cells(ctx: CallContext, node: Node | undefined): Scalar[] {
  if (!node) return [];
  return [...flatten(asMatrix(ctx.evaluate(node)))];
}
/** x/y pairs where both are numbers, for CORREL, SLOPE and friends. */
function pairs(ctx: CallContext, ys: Node | undefined, xs: Node | undefined): [number, number][] {
  const a = cells(ctx, ys);
  const b = cells(ctx, xs);
  if (a.length !== b.length) fail(err.na);
  const out: [number, number][] = [];
  a.forEach((y, i) => {
    const x = b[i]!;
    if (isError(y)) fail(y);
    if (isError(x)) fail(x);
    if (typeof y === 'number' && typeof x === 'number') out.push([y, x]);
  });
  return out;
}
function regression(ctx: CallContext, ys: Node | undefined, xs: Node | undefined) {
  const p = pairs(ctx, ys, xs);
  need(p.map(([y]) => y), 1);
  const my = mean(p.map(([y]) => y));
  const mx = mean(p.map(([, x]) => x));
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (const [y, x] of p) {
    sxy += (x - mx) * (y - my);
    sxx += (x - mx) ** 2;
    syy += (y - my) ** 2;
  }
  return { n: p.length, mx, my, sxy, sxx, syy };
}
function percentileInc(xs: number[], k: number): number {
  if (!xs.length || k < 0 || k > 1) fail(err.num);
  const s = [...xs].sort((a, b) => a - b);
  const pos = k * (s.length - 1);
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return s[lo]! + (s[hi]! - s[lo]!) * (pos - lo);
}
function percentileExc(xs: number[], k: number): number {
  const s = [...xs].sort((a, b) => a - b);
  const pos = k * (s.length + 1) - 1;
  if (!s.length || pos < 0 || pos > s.length - 1) fail(err.num);
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return s[lo]! + (s[hi]! - s[lo]!) * (pos - lo);
}
function rank(ctx: CallContext, a: Node[], average: boolean): number {
  const x = num(ctx, a[0]);
  const xs = rangeNumbers(ctx, a[1]);
  const asc = num(ctx, a[2], 0) !== 0;
  if (!xs.some((v) => roundTo15(v) === roundTo15(x))) fail(err.na);
  const better = xs.filter((v) => (asc ? v < x : v > x)).length;
  if (!average) return better + 1;
  const ties = xs.filter((v) => roundTo15(v) === roundTo15(x)).length;
  return better + (ties + 1) / 2;
}
function mode(xs: number[]): number {
  const counts = new Map<number, number>();
  let best: number | null = null;
  let bestCount = 1;
  for (const x of xs) {
    const c = (counts.get(x) ?? 0) + 1;
    counts.set(x, c);
    if (c > bestCount) {
      bestCount = c;
      best = x;
    }
  }
  // Ties go to the value that appears first, as in Excel.
  if (best === null) fail(err.na);
  const top = Math.max(...counts.values());
  return xs.find((x) => counts.get(x) === top)!;
}
const gcd2 = (a: number, b: number): number => (b === 0 ? a : gcd2(b, a % b));
function factorial(n: number): number {
  if (n < 0) fail(err.num);
  let f = 1;
  for (let i = 2; i <= Math.trunc(n); i++) f *= i;
  return f;
}
/** Excel's A ifs: averages, maxima and minima over the cells meeting every criterion. */
function ifsValues(ctx: CallContext, a: Node[]): number[] {
  const target = asMatrix(ctx.evaluate(a[0]!));
  const { hits } = matchingPositions(ctx, pairsFrom(a, 1));
  return hits.map(([r, c]) => target[r]?.[c] ?? null).filter((v): v is number => typeof v === 'number');
}

// ---------- dates ----------

const parts = (serial: number) => serialToDate(Math.floor(serial));
const timeOf = (serial: number) => {
  let secs = Math.round((serial - Math.floor(serial)) * 86400);
  if (secs >= 86400) secs = 0;
  return { h: Math.floor(secs / 3600), m: Math.floor((secs % 3600) / 60), s: secs % 60 };
};
function daysInMonth(y: number, m: number): number {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}
function holidaysOf(ctx: CallContext, node: Node | undefined): Set<number> {
  return new Set(rangeNumbers(ctx, node).map(Math.floor));
}
const isWorkday = (serial: number, holidays: Set<number>) => {
  const d = weekday(serial);
  return d !== 1 && d !== 7 && !holidays.has(serial);
};
/** 30/360 day counts: US (NASD) or European. */
function days360(start: number, end: number, european: boolean): number {
  const a = parts(start);
  const b = parts(end);
  let d1 = a.day;
  let d2 = b.day;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    const lastFeb = (p: { year: number; month: number; day: number }) => p.month === 2 && p.day === daysInMonth(p.year, 2);
    if (lastFeb(a) && lastFeb(b)) d2 = 30;
    if (lastFeb(a)) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
    if (d1 === 31) d1 = 30;
  }
  return (b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1);
}

// ---------- finance ----------

function solve(f: (x: number) => number, guess: number): number {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    if (Math.abs(y) < 1e-10) return x;
    const h = 1e-7 * Math.max(1, Math.abs(x));
    const d = (f(x + h) - y) / h;
    if (!Number.isFinite(d) || d === 0) break;
    const next = x - y / d;
    if (!Number.isFinite(next)) break;
    if (Math.abs(next - x) < 1e-12) return next;
    x = next;
  }
  return fail(err.num);
}
function fvOf(r: number, n: number, p: number, pv: number, type: number): number {
  if (r === 0) return -(pv + p * n);
  const f = (1 + r) ** n;
  return -(pv * f + (p * (1 + r * type) * (f - 1)) / r);
}
// ---------- the D-functions ----------

/** The values of `field` in the rows of `database` that `criteria` selects (rows OR, columns AND). */
function dValues(ctx: CallContext, a: Node[]): Scalar[] {
  const db = asMatrix(ctx.evaluate(a[0]!));
  const crit = asMatrix(ctx.evaluate(a[2]!));
  const head = db[0] ?? [];
  const fieldArg = scalarArg(ctx, a[1]);
  let col: number;
  if (typeof fieldArg === 'number') col = Math.trunc(fieldArg) - 1;
  else col = head.findIndex((h) => typeof h === 'string' && typeof fieldArg === 'string' && h.toLowerCase() === fieldArg.toLowerCase());
  if (col < 0 || col >= head.length) fail(err.value);
  const critHead = crit[0] ?? [];
  const critCols = critHead.map((h) => head.findIndex((x) => typeof x === 'string' && typeof h === 'string' && x.toLowerCase() === h.toLowerCase()));
  const rules = crit.slice(1).map((row) => row.map((c, i) => (c === null || c === '' || critCols[i]! < 0 ? null : { col: critCols[i]!, test: criterion(c) })));
  return db.slice(1).filter((row) => rules.length === 0 || rules.some((r) => r.every((t) => !t || t.test(row[t.col] ?? null)))).map((row) => row[col] ?? null);
}
const dNumbers = (ctx: CallContext, a: Node[]) => dValues(ctx, a).filter((v): v is number => typeof v === 'number');

// ---------- the table ----------

const trig = (f: (x: number) => number): FunctionImpl => (a, c) => f(num(c, a[0]));

export const MORE: Record<string, FunctionImpl> = {
  // Math and trigonometry
  CEILING: (a, c) => {
    const n = num(c, a[0]);
    const s = num(c, a[1], n < 0 ? -1 : 1);
    if (s === 0) return 0;
    if (n > 0 && s < 0) fail(err.num);
    return roundTo15(Math.ceil(roundTo15(n / s)) * s);
  },
  'CEILING.MATH': (a, c) => {
    const n = num(c, a[0]);
    const s = Math.abs(num(c, a[1], 1)) || 1;
    const away = num(c, a[2], 0) !== 0 && n < 0;
    return roundTo15((away ? Math.floor(roundTo15(n / s)) : Math.ceil(roundTo15(n / s))) * s);
  },
  FLOOR: (a, c) => {
    const n = num(c, a[0]);
    const s = num(c, a[1], n < 0 ? -1 : 1);
    if (s === 0) return fail(err.div0);
    if (n > 0 && s < 0) fail(err.num);
    return roundTo15(Math.floor(roundTo15(n / s)) * s);
  },
  'FLOOR.MATH': (a, c) => {
    const n = num(c, a[0]);
    const s = Math.abs(num(c, a[1], 1)) || 1;
    const toward = num(c, a[2], 0) !== 0 && n < 0;
    return roundTo15((toward ? Math.ceil(roundTo15(n / s)) : Math.floor(roundTo15(n / s))) * s);
  },
  MROUND: (a, c) => {
    const n = num(c, a[0]);
    const m = num(c, a[1]);
    if (m === 0) return 0;
    if (Math.sign(n) * Math.sign(m) < 0) fail(err.num);
    return roundTo15(roundWith(n / m, 0, (x) => Math.floor(x + 0.5)) * m);
  },
  EVEN: (a, c) => {
    const n = num(c, a[0]);
    const v = Math.ceil(Math.abs(n) / 2) * 2;
    return n < 0 ? -v : v;
  },
  ODD: (a, c) => {
    const n = num(c, a[0]);
    let v = Math.ceil(Math.abs(n));
    if (v % 2 === 0) v += 1;
    return n < 0 ? -v : v;
  },
  EXP: (a, c) => Math.exp(num(c, a[0])),
  LN: (a, c) => {
    const n = num(c, a[0]);
    return n <= 0 ? fail(err.num) : Math.log(n);
  },
  LOG10: (a, c) => {
    const n = num(c, a[0]);
    return n <= 0 ? fail(err.num) : Math.log10(n);
  },
  LOG: (a, c) => {
    const n = num(c, a[0]);
    const b = num(c, a[1], 10);
    if (n <= 0 || b <= 0) fail(err.num);
    if (b === 1) fail(err.div0);
    return Math.log(n) / Math.log(b);
  },
  FACT: (a, c) => factorial(num(c, a[0])),
  FACTDOUBLE: (a, c) => {
    const n = Math.trunc(num(c, a[0]));
    if (n < -1) fail(err.num);
    let f = 1;
    for (let i = n; i > 1; i -= 2) f *= i;
    return f;
  },
  COMBIN: (a, c) => {
    const n = Math.trunc(num(c, a[0]));
    const k = Math.trunc(num(c, a[1]));
    if (n < 0 || k < 0 || k > n) fail(err.num);
    let r = 1;
    for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
    return Math.round(r);
  },
  PERMUT: (a, c) => {
    const n = Math.trunc(num(c, a[0]));
    const k = Math.trunc(num(c, a[1]));
    if (n < 0 || k < 0 || k > n) fail(err.num);
    let r = 1;
    for (let i = 0; i < k; i++) r *= n - i;
    return r;
  },
  GCD: (a, c) => {
    const xs = numbers(c, a).map(Math.trunc);
    if (xs.some((x) => x < 0)) fail(err.num);
    return xs.reduce((g, x) => gcd2(g, x), 0);
  },
  LCM: (a, c) => {
    const xs = numbers(c, a).map(Math.trunc);
    if (xs.some((x) => x < 0)) fail(err.num);
    if (xs.some((x) => x === 0)) return 0;
    return xs.reduce((l, x) => (l * x) / gcd2(l, x), 1);
  },
  QUOTIENT: (a, c) => {
    const d = num(c, a[1]);
    return d === 0 ? fail(err.div0) : Math.trunc(num(c, a[0]) / d);
  },
  SUMSQ: (a, c) => sum(numbers(c, a).map((x) => x * x)),
  SUMX2MY2: (a, c) => sum(pairs(c, a[0], a[1]).map(([x, y]) => x * x - y * y)),
  SUMX2PY2: (a, c) => sum(pairs(c, a[0], a[1]).map(([x, y]) => x * x + y * y)),
  SUMXMY2: (a, c) => sum(pairs(c, a[0], a[1]).map(([x, y]) => (x - y) ** 2)),
  RAND: () => Math.random(),
  RANDBETWEEN: (a, c) => {
    const lo = Math.ceil(num(c, a[0]));
    const hi = Math.floor(num(c, a[1]));
    if (lo > hi) fail(err.num);
    return lo + Math.floor(Math.random() * (hi - lo + 1));
  },
  RADIANS: trig((x) => (x * Math.PI) / 180),
  DEGREES: trig((x) => (x * 180) / Math.PI),
  SIN: trig(Math.sin),
  COS: trig(Math.cos),
  TAN: trig(Math.tan),
  ASIN: (a, c) => {
    const x = num(c, a[0]);
    return Math.abs(x) > 1 ? fail(err.num) : Math.asin(x);
  },
  ACOS: (a, c) => {
    const x = num(c, a[0]);
    return Math.abs(x) > 1 ? fail(err.num) : Math.acos(x);
  },
  ATAN: trig(Math.atan),
  ATAN2: (a, c) => {
    const x = num(c, a[0]);
    const y = num(c, a[1]);
    return x === 0 && y === 0 ? fail(err.div0) : Math.atan2(y, x);
  },
  SINH: trig(Math.sinh),
  COSH: trig(Math.cosh),
  TANH: trig(Math.tanh),
  SQRTPI: (a, c) => {
    const x = num(c, a[0]);
    return x < 0 ? fail(err.num) : Math.sqrt(x * Math.PI);
  },
  SUBTOTAL: (a, c) => {
    const code = Math.trunc(num(c, a[0])) % 100;
    const names = ['', 'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P'];
    const name = names[code];
    if (!name) return fail(err.value);
    return FUNCTIONS[name]!(a.slice(1), c);
  },

  // Statistics
  AVERAGEIFS: (a, c) => {
    const xs = ifsValues(c, a);
    return xs.length ? mean(xs) : fail(err.div0);
  },
  MAXIFS: (a, c) => {
    const xs = ifsValues(c, a);
    return xs.length ? Math.max(...xs) : 0;
  },
  MINIFS: (a, c) => {
    const xs = ifsValues(c, a);
    return xs.length ? Math.min(...xs) : 0;
  },
  AVERAGEA: (a, c) => {
    const xs: number[] = [];
    for (const n of a) for (const v of flatten(asMatrix(c.evaluate(n)))) {
      if (isError(v)) fail(v);
      if (v === null) continue;
      xs.push(typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0);
    }
    return xs.length ? mean(xs) : fail(err.div0);
  },
  'STDEV.P': (a, c) => Math.sqrt(variance(numbers(c, a), false)),
  'VAR.S': (a, c) => variance(numbers(c, a), true),
  'VAR.P': (a, c) => variance(numbers(c, a), false),
  'MODE.SNGL': (a, c) => mode(numbers(c, a)),
  'PERCENTILE.INC': (a, c) => percentileInc(rangeNumbers(c, a[0]), num(c, a[1])),
  'PERCENTILE.EXC': (a, c) => percentileExc(rangeNumbers(c, a[0]), num(c, a[1])),
  'QUARTILE.INC': (a, c) => {
    const q = Math.trunc(num(c, a[1]));
    if (q < 0 || q > 4) fail(err.num);
    return percentileInc(rangeNumbers(c, a[0]), q / 4);
  },
  'QUARTILE.EXC': (a, c) => {
    const q = Math.trunc(num(c, a[1]));
    if (q < 1 || q > 3) fail(err.num);
    return percentileExc(rangeNumbers(c, a[0]), q / 4);
  },
  'RANK.EQ': (a, c) => rank(c, a, false),
  'RANK.AVG': (a, c) => rank(c, a, true),
  'PERCENTRANK.INC': (a, c) => {
    const xs = [...rangeNumbers(c, a[0])].sort((p, q) => p - q);
    const x = num(c, a[1]);
    const digits = num(c, a[2], 3);
    if (!xs.length || x < xs[0]! || x > xs[xs.length - 1]!) fail(err.na);
    let i = xs.findIndex((v) => v >= x);
    let r: number;
    if (xs[i] === x) r = i / (xs.length - 1);
    else {
      i -= 1;
      r = (i + (x - xs[i]!) / (xs[i + 1]! - xs[i]!)) / (xs.length - 1);
    }
    return roundWith(r, digits, Math.floor);
  },
  CORREL: (a, c) => {
    const { sxy, sxx, syy } = regression(c, a[0], a[1]);
    return sxx === 0 || syy === 0 ? fail(err.div0) : sxy / Math.sqrt(sxx * syy);
  },
  RSQ: (a, c) => {
    const { sxy, sxx, syy } = regression(c, a[0], a[1]);
    return sxx === 0 || syy === 0 ? fail(err.div0) : sxy ** 2 / (sxx * syy);
  },
  SLOPE: (a, c) => {
    const { sxy, sxx } = regression(c, a[0], a[1]);
    return sxx === 0 ? fail(err.div0) : sxy / sxx;
  },
  INTERCEPT: (a, c) => {
    const { sxy, sxx, mx, my } = regression(c, a[0], a[1]);
    return sxx === 0 ? fail(err.div0) : my - (sxy / sxx) * mx;
  },
  'FORECAST.LINEAR': (a, c) => {
    const x = num(c, a[0]);
    const { sxy, sxx, mx, my } = regression(c, a[1], a[2]);
    return sxx === 0 ? fail(err.div0) : my + (sxy / sxx) * (x - mx);
  },
  'COVARIANCE.P': (a, c) => {
    const { sxy, n } = regression(c, a[0], a[1]);
    return sxy / n;
  },
  'COVARIANCE.S': (a, c) => {
    const { sxy, n } = regression(c, a[0], a[1]);
    return n < 2 ? fail(err.div0) : sxy / (n - 1);
  },
  GEOMEAN: (a, c) => {
    const xs = need(numbers(c, a));
    if (xs.some((x) => x <= 0)) fail(err.num);
    return Math.exp(mean(xs.map(Math.log)));
  },
  HARMEAN: (a, c) => {
    const xs = need(numbers(c, a));
    if (xs.some((x) => x <= 0)) fail(err.num);
    return xs.length / sum(xs.map((x) => 1 / x));
  },
  AVEDEV: (a, c) => {
    const xs = need(numbers(c, a));
    const m = mean(xs);
    return mean(xs.map((x) => Math.abs(x - m)));
  },
  DEVSQ: (a, c) => {
    const xs = numbers(c, a);
    const m = xs.length ? mean(xs) : 0;
    return sum(xs.map((x) => (x - m) ** 2));
  },
  'NORM.DIST': (a, c) => {
    const x = num(c, a[0]);
    const m = num(c, a[1]);
    const sd = num(c, a[2]);
    if (sd <= 0) fail(err.num);
    const z = (x - m) / sd;
    return toBoolean(scalarArg(c, a[3])) ? 0.5 * (1 + erf(z / Math.SQRT2)) : Math.exp(-0.5 * z * z) / (sd * Math.sqrt(2 * Math.PI));
  },
  'NORM.S.DIST': (a, c) => {
    const z = num(c, a[0]);
    return toBoolean(scalarArg(c, a[1])) ? 0.5 * (1 + erf(z / Math.SQRT2)) : Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
  },
  STANDARDIZE: (a, c) => {
    const sd = num(c, a[2]);
    return sd <= 0 ? fail(err.num) : (num(c, a[0]) - num(c, a[1])) / sd;
  },

  // Text
  CHAR: (a, c) => {
    const n = Math.trunc(num(c, a[0]));
    return n < 1 || n > 255 ? fail(err.value) : String.fromCharCode(n);
  },
  CODE: (a, c) => {
    const s = text(c, a[0]);
    return s ? s.charCodeAt(0) : fail(err.value);
  },
  UNICHAR: (a, c) => {
    const n = Math.trunc(num(c, a[0]));
    return n < 1 || n > 0x10ffff ? fail(err.value) : String.fromCodePoint(n);
  },
  UNICODE: (a, c) => {
    const s = text(c, a[0]);
    return s ? s.codePointAt(0)! : fail(err.value);
  },
  CLEAN: (a, c) => text(c, a[0]).replace(/[\x00-\x1f]/g, ''),
  T: (a, c) => {
    const v = scalarArg(c, a[0]);
    return typeof v === 'string' ? v : '';
  },
  /** The cell shows `friendly`, or the link itself; the grid makes it clickable. */
  HYPERLINK: (a, c) => {
    const link = scalarArg(c, a[0]);
    if (isError(link)) return link;
    if (a.length < 2) return typeof link === 'string' ? link : text(c, a[0]);
    const shown = scalarArg(c, a[1]);
    return shown === null ? 0 : shown;
  },
  N: (a, c) => {
    const v = scalarArg(c, a[0]);
    if (isError(v)) return v;
    return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0;
  },
  NUMBERVALUE: (a, c) => {
    const s = text(c, a[0]);
    const dec = a[1] ? text(c, a[1]) : '.';
    const group = a[2] ? text(c, a[2]) : ',';
    const cleaned = s.replaceAll(group, '').replaceAll(dec, '.').replace(/\s/g, '');
    if (cleaned === '') return 0;
    const n = parseNumericText(cleaned);
    return n === null ? fail(err.value) : n;
  },
  DOLLAR: (a, c) => {
    const d = Math.trunc(num(c, a[1], 2));
    const n = d < 0 ? roundWith(num(c, a[0]), d, (x) => Math.floor(x + 0.5)) : num(c, a[0]);
    const code = `$#,##0${d > 0 ? `.${'0'.repeat(d)}` : ''}`;
    return formatValue(n, `${code};($#,##0${d > 0 ? `.${'0'.repeat(d)}` : ''})`).text;
  },
  FIXED: (a, c) => {
    const d = Math.trunc(num(c, a[1], 2));
    const noCommas = a[2] ? toBoolean(scalarArg(c, a[2])) : false;
    const n = roundWith(num(c, a[0]), d, (x) => Math.floor(x + 0.5));
    return formatValue(n, `${noCommas ? '0' : '#,##0'}${d > 0 ? `.${'0'.repeat(d)}` : ''}`).text;
  },
  TEXTBEFORE: (a, c) => textAround(a, c, true),
  TEXTAFTER: (a, c) => textAround(a, c, false),
  REGEXTEST: (a, c) => {
    try {
      return new RegExp(text(c, a[1]), num(c, a[2], 0) === 1 ? 'i' : '').test(text(c, a[0]));
    } catch {
      return fail(err.value);
    }
  },

  // Date and time
  TIME: (a, c) => {
    const secs = Math.trunc(num(c, a[0])) * 3600 + Math.trunc(num(c, a[1])) * 60 + Math.trunc(num(c, a[2]));
    if (secs < 0) fail(err.num);
    return (secs % 86400) / 86400;
  },
  HOUR: (a, c) => timeOf(num(c, a[0])).h,
  MINUTE: (a, c) => timeOf(num(c, a[0])).m,
  SECOND: (a, c) => timeOf(num(c, a[0])).s,
  DATEVALUE: (a, c) => {
    const n = parseNumericText(text(c, a[0]));
    return n === null ? fail(err.value) : Math.floor(n);
  },
  TIMEVALUE: (a, c) => {
    const n = parseNumericText(text(c, a[0]));
    return n === null ? fail(err.value) : n - Math.floor(n);
  },
  WEEKNUM: (a, c) => {
    const serial = Math.floor(num(c, a[0]));
    const type = num(c, a[1], 1);
    const start = type === 2 || type === 11 ? 2 : type === 1 || type === 17 ? 1 : type === 12 ? 3 : type === 13 ? 4 : type === 14 ? 5 : type === 15 ? 6 : type === 16 ? 7 : 0;
    if (type === 21) return isoWeek(serial);
    if (!start) fail(err.num);
    const jan1 = dateToSerial(parts(serial).year, 1, 1);
    const offset = (weekday(jan1) - start + 7) % 7;
    return Math.floor((serial - jan1 + offset) / 7) + 1;
  },
  ISOWEEKNUM: (a, c) => isoWeek(Math.floor(num(c, a[0]))),
  NETWORKDAYS: (a, c) => {
    let s = Math.floor(num(c, a[0]));
    let e = Math.floor(num(c, a[1]));
    const h = holidaysOf(c, a[2]);
    const sign = e < s ? -1 : 1;
    if (e < s) [s, e] = [e, s];
    let n = 0;
    for (let d = s; d <= e; d++) if (isWorkday(d, h)) n++;
    return n * sign;
  },
  WORKDAY: (a, c) => {
    let d = Math.floor(num(c, a[0]));
    let left = Math.trunc(num(c, a[1]));
    const h = holidaysOf(c, a[2]);
    const step = left < 0 ? -1 : 1;
    while (left !== 0) {
      d += step;
      if (isWorkday(d, h)) left -= step;
    }
    return d;
  },
  DATEDIF: (a, c) => {
    const s = Math.floor(num(c, a[0]));
    const e = Math.floor(num(c, a[1]));
    if (s > e) fail(err.num);
    const unit = text(c, a[2]).toUpperCase();
    const p = parts(s);
    const q = parts(e);
    let months = (q.year - p.year) * 12 + (q.month - p.month);
    if (q.day < p.day) months -= 1;
    switch (unit) {
      case 'D':
        return e - s;
      case 'M':
        return months;
      case 'Y':
        return Math.floor(months / 12);
      case 'YM':
        return months % 12;
      case 'MD': {
        if (q.day >= p.day) return q.day - p.day;
        const prev = q.month === 1 ? { y: q.year - 1, m: 12 } : { y: q.year, m: q.month - 1 };
        return daysInMonth(prev.y, prev.m) - p.day + q.day;
      }
      case 'YD': {
        let anniversary = dateToSerial(q.year, p.month, p.day);
        if (anniversary > e) anniversary = dateToSerial(q.year - 1, p.month, p.day);
        return e - anniversary;
      }
      default:
        return fail(err.num);
    }
  },
  DAYS360: (a, c) => days360(Math.floor(num(c, a[0])), Math.floor(num(c, a[1])), a[2] ? toBoolean(scalarArg(c, a[2])) : false),
  YEARFRAC: (a, c) => {
    let s = Math.floor(num(c, a[0]));
    let e = Math.floor(num(c, a[1]));
    if (s > e) [s, e] = [e, s];
    const basis = Math.trunc(num(c, a[2], 0));
    switch (basis) {
      case 0:
        return days360(s, e, false) / 360;
      case 1: {
        const p = parts(s);
        const q = parts(e);
        if (p.year === q.year) return (e - s) / (isLeap(p.year) ? 366 : 365);
        const span = q.year - p.year + 1;
        const days = dateToSerial(q.year + 1, 1, 1) - dateToSerial(p.year, 1, 1);
        return (e - s) / (days / span);
      }
      case 2:
        return (e - s) / 360;
      case 3:
        return (e - s) / 365;
      case 4:
        return days360(s, e, true) / 360;
      default:
        return fail(err.num);
    }
  },

  // Lookup and reference
  XMATCH: (a, c) => {
    const x = scalarArg(c, a[0]);
    const values = lineOf(asMatrix(c.evaluate(a[1]!)));
    const matchMode = num(c, a[2], 0);
    const reverse = num(c, a[3], 1) < 0;
    const order = reverse ? [...values.keys()].reverse() : [...values.keys()];
    if (matchMode === 2 && typeof x === 'string') {
      const re = wildcardRegex(x);
      const i = order.find((k) => typeof values[k] === 'string' && re.test(values[k] as string));
      return i === undefined ? fail(err.na) : i + 1;
    }
    const exact = order.find((k) => values[k] !== null && !isError(values[k]) && typeof values[k] === typeof x && compare(values[k]!, x) === 0);
    if (exact !== undefined) return exact + 1;
    if (matchMode === -1 || matchMode === 1) {
      let best = -1;
      for (const k of order) {
        const v = values[k]!;
        if (v === null || isError(v) || typeof v !== typeof x) continue;
        const cmp = compare(v, x);
        if (matchMode === -1 && cmp < 0 && (best < 0 || compare(v, values[best]!) > 0)) best = k;
        if (matchMode === 1 && cmp > 0 && (best < 0 || compare(v, values[best]!) < 0)) best = k;
      }
      return best < 0 ? fail(err.na) : best + 1;
    }
    return fail(err.na);
  },
  LOOKUP: (a, c) => {
    const x = scalarArg(c, a[0]);
    const m = asMatrix(c.evaluate(a[1]!));
    const wide = (m[0]?.length ?? 0) > m.length;
    const keys = a[2] ? lineOf(m) : wide ? m[0]! : m.map((r) => r[0] ?? null);
    const results = a[2] ? [...flatten(asMatrix(c.evaluate(a[2])))] : wide ? m[m.length - 1]! : m.map((r) => r[r.length - 1] ?? null);
    const i = findApprox(keys, x);
    return i < 0 ? fail(err.na) : (results[i] ?? fail(err.na));
  },
  ADDRESS: (a, c) => {
    const row = Math.trunc(num(c, a[0]));
    const col = Math.trunc(num(c, a[1]));
    if (row < 1 || col < 1) fail(err.value);
    const abs = Math.trunc(num(c, a[2], 1));
    const a1 = a[3] ? toBoolean(scalarArg(c, a[3])) : true;
    const sheet = a[4] ? text(c, a[4]) : '';
    const prefix = sheet ? (/^[A-Za-z_][\w.]*$/.test(sheet) ? `${sheet}!` : `'${sheet.replaceAll("'", "''")}'!`) : '';
    if (!a1) return `${prefix}R${abs === 1 || abs === 2 ? row : `[${row}]`}C${abs === 1 || abs === 3 ? col : `[${col}]`}`;
    const colAbs = abs === 1 || abs === 3 ? '$' : '';
    const rowAbs = abs === 1 || abs === 2 ? '$' : '';
    return `${prefix}${colAbs}${columnName(col - 1)}${rowAbs}${row}`;
  },
  AREAS: () => 1,

  // Information
  ISEVEN: (a, c) => Math.trunc(num(c, a[0])) % 2 === 0,
  ISODD: (a, c) => Math.abs(Math.trunc(num(c, a[0])) % 2) === 1,
  ISNONTEXT: (a, c) => typeof scalarArg(c, a[0]) !== 'string',
  ISERR: (a, c) => {
    const v = scalarArg(c, a[0]);
    return isError(v) && v.code !== '#N/A';
  },
  TYPE: (a, c) => {
    const v = a[0] ? c.evaluate(a[0]) : null;
    if (isMatrix(v) && a[0]!.kind !== 'ref' && a[0]!.kind !== 'range') return 64;
    const s = isMatrix(v) ? (v[0]?.[0] ?? null) : v;
    return typeof s === 'number' || s === null ? 1 : typeof s === 'string' ? 2 : typeof s === 'boolean' ? 4 : 16;
  },
  'ERROR.TYPE': (a, c) => {
    const v = scalarArg(c, a[0]);
    if (!isError(v)) return fail(err.na);
    const codes = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'];
    const i = codes.indexOf(v.code);
    return i < 0 ? 8 : i + 1;
  },

  // Finance
  NPER: (a, c) => {
    const r = num(c, a[0]);
    const p = num(c, a[1]);
    const pv = num(c, a[2]);
    const fv = num(c, a[3], 0);
    const type = num(c, a[4], 0);
    if (r === 0) return p === 0 ? fail(err.num) : -(pv + fv) / p;
    const k = p * (1 + r * type);
    return Math.log((k - fv * r) / (k + pv * r)) / Math.log(1 + r);
  },
  RATE: (a, c) => {
    const n = num(c, a[0]);
    const p = num(c, a[1]);
    const pv = num(c, a[2]);
    const fv = num(c, a[3], 0);
    const type = num(c, a[4], 0);
    const guess = num(c, a[5], 0.1);
    return solve((r) => (Math.abs(r) < 1e-12 ? pv + p * n + fv : pv * (1 + r) ** n + (p * (1 + r * type) * ((1 + r) ** n - 1)) / r + fv), guess);
  },
  IPMT: (a, c) => interestPart(a, c),
  PPMT: (a, c) => {
    const r = num(c, a[0]);
    const n = num(c, a[2]);
    const pv = num(c, a[3]);
    const fv = num(c, a[4], 0);
    const type = num(c, a[5], 0);
    return pmt(r, n, pv, fv, type) - (interestPart(a, c) as number);
  },
  IRR: (a, c) => {
    const flows = rangeNumbers(c, a[0]);
    if (!flows.some((f) => f > 0) || !flows.some((f) => f < 0)) fail(err.num);
    return solve((r) => flows.reduce((acc, f, i) => acc + f / (1 + r) ** i, 0), num(c, a[1], 0.1));
  },
  XNPV: (a, c) => {
    const r = num(c, a[0]);
    const flows = rangeNumbers(c, a[1]);
    const dates = rangeNumbers(c, a[2]);
    if (flows.length !== dates.length) fail(err.num);
    return flows.reduce((acc, f, i) => acc + f / (1 + r) ** ((dates[i]! - dates[0]!) / 365), 0);
  },
  XIRR: (a, c) => {
    const flows = rangeNumbers(c, a[0]);
    const dates = rangeNumbers(c, a[1]);
    if (flows.length !== dates.length || !flows.some((f) => f > 0) || !flows.some((f) => f < 0)) fail(err.num);
    return solve((r) => flows.reduce((acc, f, i) => acc + f / (1 + r) ** ((dates[i]! - dates[0]!) / 365), 0), num(c, a[2], 0.1));
  },
  SLN: (a, c) => {
    const life = num(c, a[2]);
    return life === 0 ? fail(err.div0) : (num(c, a[0]) - num(c, a[1])) / life;
  },
  SYD: (a, c) => {
    const cost = num(c, a[0]);
    const salvage = num(c, a[1]);
    const life = num(c, a[2]);
    const per = num(c, a[3]);
    if (life <= 0 || per < 1 || per > life) fail(err.num);
    return ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1));
  },
  DDB: (a, c) => {
    const cost = num(c, a[0]);
    const salvage = num(c, a[1]);
    const life = num(c, a[2]);
    const per = num(c, a[3]);
    const factor = num(c, a[4], 2);
    if (life <= 0 || per < 1 || per > life) fail(err.num);
    let book = cost;
    let dep = 0;
    for (let i = 1; i <= per; i++) {
      dep = Math.min(book * (factor / life), Math.max(0, book - salvage));
      book -= dep;
    }
    return dep;
  },
  EFFECT: (a, c) => {
    const r = num(c, a[0]);
    const n = Math.trunc(num(c, a[1]));
    if (r <= 0 || n < 1) fail(err.num);
    return (1 + r / n) ** n - 1;
  },
  NOMINAL: (a, c) => {
    const r = num(c, a[0]);
    const n = Math.trunc(num(c, a[1]));
    if (r <= 0 || n < 1) fail(err.num);
    return n * ((1 + r) ** (1 / n) - 1);
  },

  // Database
  DSUM: (a, c) => sum(dNumbers(c, a)),
  DAVERAGE: (a, c) => {
    const xs = dNumbers(c, a);
    return xs.length ? mean(xs) : fail(err.div0);
  },
  DCOUNT: (a, c) => dNumbers(c, a).length,
  DCOUNTA: (a, c) => dValues(c, a).filter((v) => v !== null && v !== '').length,
  DMAX: (a, c) => {
    const xs = dNumbers(c, a);
    return xs.length ? Math.max(...xs) : 0;
  },
  DMIN: (a, c) => {
    const xs = dNumbers(c, a);
    return xs.length ? Math.min(...xs) : 0;
  },
  DPRODUCT: (a, c) => dNumbers(c, a).reduce((p, x) => p * x, 1),
  DGET: (a, c) => {
    const vs = dValues(c, a);
    if (vs.length === 0) fail(err.value);
    if (vs.length > 1) fail(err.num);
    return vs[0]!;
  },
  DSTDEV: (a, c) => Math.sqrt(variance(dNumbers(c, a), true)),
  DSTDEVP: (a, c) => Math.sqrt(variance(dNumbers(c, a), false)),
  DVAR: (a, c) => variance(dNumbers(c, a), true),
  DVARP: (a, c) => variance(dNumbers(c, a), false),
};

function textAround(a: Node[], c: CallContext, before: boolean): Scalar {
  const s = text(c, a[0]);
  const delim = text(c, a[1]);
  let n = Math.trunc(num(c, a[2], 1));
  const ci = num(c, a[3], 0) === 1;
  if (n === 0 || delim === '') return fail(err.value);
  const hay = ci ? s.toLowerCase() : s;
  const needle = ci ? delim.toLowerCase() : delim;
  const idx: number[] = [];
  for (let i = hay.indexOf(needle); i >= 0; i = hay.indexOf(needle, i + needle.length)) idx.push(i);
  if (n < 0) n = idx.length + n + 1;
  const at = idx[n - 1];
  if (at === undefined) return a[5] ? scalarArg(c, a[5]) : fail(err.na);
  return before ? s.slice(0, at) : s.slice(at + delim.length);
}

function interestPart(a: Node[], c: CallContext): Value {
  const r = num(c, a[0]);
  const per = Math.trunc(num(c, a[1]));
  const n = num(c, a[2]);
  const pv = num(c, a[3]);
  const fv = num(c, a[4], 0);
  const type = num(c, a[5], 0);
  if (per < 1 || per > n) fail(err.num);
  const p = pmt(r, n, pv, fv, type);
  if (type === 1) {
    if (per === 1) return 0;
    return (fvOf(r, per - 2, p, pv, 1) - p) * r;
  }
  return fvOf(r, per - 1, p, pv, 0) * r;
}

const isLeap = (y: number) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;

function isoWeek(serial: number): number {
  // Thursday decides the year a week belongs to.
  const dow = (weekday(serial) + 5) % 7; // Monday 0
  const thursday = serial - dow + 3;
  const jan1 = dateToSerial(parts(thursday).year, 1, 1);
  return Math.floor((thursday - jan1) / 7) + 1;
}

/** The error function, to 1e-7 (Abramowitz and Stegun 7.1.26), for NORM.DIST. */
function erf(x: number): number {
  const t = 1 / (1 + 0.3275911 * Math.abs(x));
  const y = 1 - ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x);
  return x >= 0 ? y : -y;
}

// Names older workbooks use for the functions above.
const ALIASES: Record<string, string> = {
  STDEV: 'STDEV.S',
  STDEVP: 'STDEV.P',
  VAR: 'VAR.S',
  VARP: 'VAR.P',
  MODE: 'MODE.SNGL',
  PERCENTILE: 'PERCENTILE.INC',
  QUARTILE: 'QUARTILE.INC',
  RANK: 'RANK.EQ',
  PERCENTRANK: 'PERCENTRANK.INC',
  FORECAST: 'FORECAST.LINEAR',
  COVAR: 'COVARIANCE.P',
  NORMDIST: 'NORM.DIST',
  NORMSDIST: 'NORM.S.DIST',
  'CEILING.PRECISE': 'CEILING.MATH',
  'FLOOR.PRECISE': 'FLOOR.MATH',
  PEARSON: 'CORREL',
};

Object.assign(FUNCTIONS, MORE);
for (const [alias, target] of Object.entries(ALIASES)) FUNCTIONS[alias] ??= FUNCTIONS[target]!;
