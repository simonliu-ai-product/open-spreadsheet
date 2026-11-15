/**
 * The third batch: engineering (number bases, bits, complex numbers, units),
 * probability distributions, and the rest of math and statistics.
 */
import { FUNCTIONS, type FunctionImpl, num, numbers, scalarArg, sum, text } from './functions.ts';
import { err, fail, flatten, isError, roundTo15, toBoolean } from './values.ts';
import type { Node } from './parse.ts';
import type { CallContext } from './functions.ts';

const bool = (c: CallContext, n: Node | undefined, d = false) => (n === undefined || n.kind === 'missing' ? d : toBoolean(scalarArg(c, n)));

// ---------- number bases ----------

/** Excel's 10-digit two's complement bases: 1023 in binary, 0x7FFFFFFFFF in hex. */
const BASES = { 2: { digits: 10, name: 'BIN' }, 8: { digits: 10, name: 'OCT' }, 16: { digits: 10, name: 'HEX' } } as const;
type Radix = 2 | 8 | 16;
function fromBase(s: string, radix: Radix): number {
  const t = s.trim().toUpperCase();
  if (t.length > 10 || !/^[0-9A-F]*$/.test(t)) fail(err.num);
  if (t === '') return 0;
  const n = Number.parseInt(t, radix);
  if (Number.isNaN(n) || [...t].some((ch) => Number.parseInt(ch, 16) >= radix)) fail(err.num);
  const full = radix ** BASES[radix].digits;
  return t.length === 10 && n >= full / 2 ? n - full : n;
}
function toBase(n: number, radix: Radix, places?: number): string {
  const full = radix ** BASES[radix].digits;
  const v = Math.trunc(n);
  if (v < -full / 2 || v >= full / 2) fail(err.num);
  if (v < 0) return (full + v).toString(radix).toUpperCase();
  const s = v.toString(radix).toUpperCase();
  if (places === undefined) return s;
  if (places < s.length || places > 10) fail(err.num);
  return s.padStart(Math.trunc(places), '0');
}
const placesOf = (c: CallContext, n: Node | undefined) => (n === undefined || n.kind === 'missing' ? undefined : num(c, n));
const conv =
  (from: Radix | 10, to: Radix | 10): FunctionImpl =>
  (a, c) => {
    const value = from === 10 ? num(c, a[0]) : fromBase(text(c, a[0]), from);
    if (to === 10) return value;
    return toBase(value, to, placesOf(c, a[1]));
  };

// ---------- complex numbers ----------

type Cx = { re: number; im: number; unit: 'i' | 'j' };
/** "3+4i", "-2.5j", "i", "7": real and imaginary parts, as Excel's IM functions read them. */
function parseCx(v: unknown): Cx {
  if (typeof v === 'number') return { re: v, im: 0, unit: 'i' };
  const s = String(v ?? '').replace(/\s/g, '');
  if (s === '') return { re: 0, im: 0, unit: 'i' };
  if (!/[ij]$/.test(s)) {
    const n = Number(s);
    return Number.isNaN(n) ? fail(err.num) : { re: n, im: 0, unit: 'i' };
  }
  const unit = s.endsWith('j') ? 'j' : 'i';
  const body = s.slice(0, -1);
  // The sign that starts the imaginary part: the last + or - that is not an exponent's.
  let k = -1;
  for (let i = body.length - 1; i > 0; i--) {
    if ((body[i] === '+' || body[i] === '-') && !/e/i.test(body[i - 1]!)) {
      k = i;
      break;
    }
  }
  const reText = k > 0 ? body.slice(0, k) : '';
  const imText = k > 0 ? body.slice(k) : body;
  const im = imText === '' || imText === '+' ? 1 : imText === '-' ? -1 : Number(imText);
  const re = reText === '' ? 0 : Number(reText);
  if (Number.isNaN(re) || Number.isNaN(im)) return fail(err.num);
  return { re, im, unit };
}
const fmt = (x: number) => {
  const r = roundTo15(x);
  return Object.is(r, -0) ? '0' : String(r);
};
function showCx(z: Cx): string {
  const re = roundTo15(z.re);
  const im = roundTo15(z.im);
  if (im === 0) return fmt(re);
  const imText = im === 1 ? '' : im === -1 ? '-' : fmt(im);
  if (re === 0) return `${imText}${z.unit}`;
  return `${fmt(re)}${im > 0 ? '+' : ''}${imText}${z.unit}`;
}
const cx = (c: CallContext, n: Node | undefined) => parseCx(scalarArg(c, n));
const mul = (a: Cx, b: Cx): Cx => ({ re: a.re * b.re - a.im * b.im, im: a.re * b.im + a.im * b.re, unit: a.unit });
const abs = (z: Cx) => Math.hypot(z.re, z.im);
const arg = (z: Cx) => Math.atan2(z.im, z.re);
const polar = (r: number, t: number, unit: 'i' | 'j'): Cx => ({ re: r * Math.cos(t), im: r * Math.sin(t), unit });
function cxList(c: CallContext, a: Node[]): Cx[] {
  const out: Cx[] = [];
  for (const n of a) {
    for (const v of flatten(c.evaluate(n))) {
      if (isError(v)) fail(v);
      if (v !== null) out.push(parseCx(v));
    }
  }
  return out;
}

// ---------- units for CONVERT ----------

const UNITS: Record<string, [kind: string, factor: number]> = {
  m: ['len', 1], km: ['len', 1000], cm: ['len', 0.01], mm: ['len', 0.001], um: ['len', 1e-6], nm: ['len', 1e-9],
  in: ['len', 0.0254], ft: ['len', 0.3048], yd: ['len', 0.9144], mi: ['len', 1609.344], Nmi: ['len', 1852], ang: ['len', 1e-10],
  g: ['mass', 1], kg: ['mass', 1000], mg: ['mass', 0.001], lbm: ['mass', 453.59237], ozm: ['mass', 28.349523125], ton: ['mass', 907184.74], stone: ['mass', 6350.29318],
  sec: ['time', 1], s: ['time', 1], mn: ['time', 60], min: ['time', 60], hr: ['time', 3600], day: ['time', 86400], d: ['time', 86400], yr: ['time', 31557600],
  l: ['vol', 1], L: ['vol', 1], lt: ['vol', 1], ml: ['vol', 0.001], mL: ['vol', 0.001], gal: ['vol', 3.785411784], qt: ['vol', 0.946352946], pt: ['vol', 0.473176473], cup: ['vol', 0.2365882365], oz: ['vol', 0.0295735295625], tbs: ['vol', 0.01478676478125], tsp: ['vol', 0.00492892159375], m3: ['vol', 1000],
  m2: ['area', 1], km2: ['area', 1e6], ha: ['area', 1e4], ft2: ['area', 0.09290304], in2: ['area', 0.00064516], acre: ['area', 4046.8564224], uk_acre: ['area', 4046.8564224],
  J: ['energy', 1], kJ: ['energy', 1000], cal: ['energy', 4.1868], kcal: ['energy', 4186.8], Wh: ['energy', 3600], kWh: ['energy', 3.6e6], BTU: ['energy', 1055.05585262],
  W: ['power', 1], kW: ['power', 1000], HP: ['power', 745.69987158227], PS: ['power', 735.49875],
  Pa: ['pressure', 1], kPa: ['pressure', 1000], atm: ['pressure', 101325], mmHg: ['pressure', 133.322], psi: ['pressure', 6894.75729316836], bar: ['pressure', 1e5],
  'm/s': ['speed', 1], 'm/h': ['speed', 1 / 3600], mph: ['speed', 0.44704], kn: ['speed', 0.514444444444444], 'km/h': ['speed', 1 / 3.6],
  bit: ['info', 1], byte: ['info', 8], kbyte: ['info', 8000], Mbyte: ['info', 8e6], Gbyte: ['info', 8e9],
};
function convert(x: number, from: string, to: string): number {
  const temp: Record<string, [(v: number) => number, (k: number) => number]> = {
    C: [(v) => v + 273.15, (k) => k - 273.15], cel: [(v) => v + 273.15, (k) => k - 273.15],
    F: [(v) => ((v - 32) * 5) / 9 + 273.15, (k) => ((k - 273.15) * 9) / 5 + 32], fah: [(v) => ((v - 32) * 5) / 9 + 273.15, (k) => ((k - 273.15) * 9) / 5 + 32],
    K: [(v) => v, (k) => k], kel: [(v) => v, (k) => k],
  };
  if (temp[from] && temp[to]) return temp[to][1](temp[from][0](x));
  const a = UNITS[from];
  const b = UNITS[to];
  if (!a || !b || a[0] !== b[0]) return fail(err.na);
  return roundTo15((x * a[1]) / b[1]);
}

// ---------- special functions ----------

export function gammaln(x: number): number {
  const g = [76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  let y = x;
  const tmp = x + 5.5 - (x + 0.5) * Math.log(x + 5.5);
  let ser = 1.000000000190015;
  for (const k of g) ser += k / ++y;
  return -tmp + Math.log((2.5066282746310005 * ser) / x);
}
function gammaFn(x: number): number {
  if (x <= 0 && Number.isInteger(x)) return fail(err.num);
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gammaFn(1 - x));
  return Math.exp(gammaln(x));
}
/** Regularized lower incomplete gamma P(a, x). */
function gammaP(a: number, x: number): number {
  if (x <= 0) return 0;
  if (x < a + 1) {
    let sumv = 1 / a;
    let del = sumv;
    let ap = a;
    for (let n = 0; n < 500; n++) {
      ap += 1;
      del *= x / ap;
      sumv += del;
      if (Math.abs(del) < Math.abs(sumv) * 1e-15) break;
    }
    return sumv * Math.exp(-x + a * Math.log(x) - gammaln(a));
  }
  let b = x + 1 - a;
  let cc = 1 / 1e-300;
  let d = 1 / b;
  let h = d;
  for (let i = 1; i < 500; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < 1e-300) d = 1e-300;
    cc = b + an / cc;
    if (Math.abs(cc) < 1e-300) cc = 1e-300;
    d = 1 / d;
    const del = d * cc;
    h *= del;
    if (Math.abs(del - 1) < 1e-15) break;
  }
  return 1 - Math.exp(-x + a * Math.log(x) - gammaln(a)) * h;
}
function betacf(a: number, b: number, x: number): number {
  const qab = a + b;
  const qap = a + 1;
  const qam = a - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < 1e-300) d = 1e-300;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= 500; m++) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < 1e-300) d = 1e-300;
    c = 1 + aa / c;
    if (Math.abs(c) < 1e-300) c = 1e-300;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < 1e-300) d = 1e-300;
    c = 1 + aa / c;
    if (Math.abs(c) < 1e-300) c = 1e-300;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-15) break;
  }
  return h;
}
/** Regularized incomplete beta I_x(a, b). */
function betaI(x: number, a: number, b: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const bt = Math.exp(gammaln(a + b) - gammaln(a) - gammaln(b) + a * Math.log(x) + b * Math.log(1 - x));
  return x < (a + 1) / (a + b + 2) ? (bt * betacf(a, b, x)) / a : 1 - (bt * betacf(b, a, 1 - x)) / b;
}
function erf(x: number): number {
  // Through the incomplete gamma: erf(x) = P(1/2, x²), accurate to ~1e-15.
  const v = gammaP(0.5, x * x);
  return x < 0 ? -v : v;
}
const normCdf = (z: number) => 0.5 * (1 + erf(z / Math.SQRT2));
/** Inverse standard normal (Acklam), refined by one Newton step. */
function normInv(p: number): number {
  if (p <= 0 || p >= 1) return fail(err.num);
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const cc = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const lo = 0.02425;
  let x: number;
  if (p < lo) {
    const q = Math.sqrt(-2 * Math.log(p));
    x = (((((cc[0]! * q + cc[1]!) * q + cc[2]!) * q + cc[3]!) * q + cc[4]!) * q + cc[5]!) / ((((d[0]! * q + d[1]!) * q + d[2]!) * q + d[3]!) * q + 1);
  } else if (p <= 1 - lo) {
    const q = p - 0.5;
    const r = q * q;
    x = ((((((a[0]! * r + a[1]!) * r + a[2]!) * r + a[3]!) * r + a[4]!) * r + a[5]!) * q) / (((((b[0]! * r + b[1]!) * r + b[2]!) * r + b[3]!) * r + b[4]!) * r + 1);
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((cc[0]! * q + cc[1]!) * q + cc[2]!) * q + cc[3]!) * q + cc[4]!) * q + cc[5]!) / ((((d[0]! * q + d[1]!) * q + d[2]!) * q + d[3]!) * q + 1);
  }
  const e = normCdf(x) - p;
  const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2);
  return x - u / (1 + (x * u) / 2);
}
const tCdf = (x: number, v: number) => {
  const ib = betaI(v / (v + x * x), v / 2, 0.5);
  return x >= 0 ? 1 - ib / 2 : ib / 2;
};
const tPdf = (x: number, v: number) => Math.exp(gammaln((v + 1) / 2) - gammaln(v / 2)) / Math.sqrt(v * Math.PI) * (1 + (x * x) / v) ** (-(v + 1) / 2);
/** x with f(x) = target on [lo, hi] for an increasing f, by bisection. */
function bisect(f: (x: number) => number, target: number, lo: number, hi: number): number {
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    if (f(mid) < target) lo = mid;
    else hi = mid;
    if (hi - lo < 1e-13 * Math.max(1, Math.abs(mid))) break;
  }
  return (lo + hi) / 2;
}
const tInv2 = (p: number, v: number) => bisect((x) => 1 - betaI(v / (v + x * x), v / 2, 0.5), 1 - p, 0, 1e7);
const binom = (n: number, k: number) => Math.exp(gammaln(n + 1) - gammaln(k + 1) - gammaln(n - k + 1));
function positive(...xs: number[]) {
  if (xs.some((x) => !(x > 0))) fail(err.num);
}

// ---------- statistics over mixed values (the A functions) ----------

function aValues(c: CallContext, a: Node[]): number[] {
  const out: number[] = [];
  for (const n of a) {
    const v = c.evaluate(n);
    for (const s of flatten(v)) {
      if (isError(s)) fail(s);
      if (s === null) continue;
      out.push(typeof s === 'number' ? s : typeof s === 'boolean' ? (s ? 1 : 0) : 0);
    }
  }
  return out;
}
function varianceOf(xs: number[], sample: boolean): number {
  if (xs.length < (sample ? 2 : 1)) fail(err.div0);
  const m = sum(xs) / xs.length;
  return sum(xs.map((x) => (x - m) ** 2)) / (xs.length - (sample ? 1 : 0));
}

const ROMAN: [number, string][] = [
  [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
];

export const ENGINEERING: Record<string, FunctionImpl> = {
  BIN2DEC: conv(2, 10), BIN2OCT: conv(2, 8), BIN2HEX: conv(2, 16),
  OCT2DEC: conv(8, 10), OCT2BIN: conv(8, 2), OCT2HEX: conv(8, 16),
  HEX2DEC: conv(16, 10), HEX2BIN: conv(16, 2), HEX2OCT: conv(16, 8),
  DEC2BIN: conv(10, 2), DEC2OCT: conv(10, 8), DEC2HEX: conv(10, 16),
  BITAND: (a, c) => Number(BigInt(Math.trunc(num(c, a[0]))) & BigInt(Math.trunc(num(c, a[1])))),
  BITOR: (a, c) => Number(BigInt(Math.trunc(num(c, a[0]))) | BigInt(Math.trunc(num(c, a[1])))),
  BITXOR: (a, c) => Number(BigInt(Math.trunc(num(c, a[0]))) ^ BigInt(Math.trunc(num(c, a[1])))),
  BITLSHIFT: (a, c) => {
    const s = Math.trunc(num(c, a[1]));
    const n = BigInt(Math.trunc(num(c, a[0])));
    return Number(s >= 0 ? n << BigInt(s) : n >> BigInt(-s));
  },
  BITRSHIFT: (a, c) => {
    const s = Math.trunc(num(c, a[1]));
    const n = BigInt(Math.trunc(num(c, a[0])));
    return Number(s >= 0 ? n >> BigInt(s) : n << BigInt(-s));
  },
  DELTA: (a, c) => (num(c, a[0]) === num(c, a[1], 0) ? 1 : 0),
  GESTEP: (a, c) => (num(c, a[0]) >= num(c, a[1], 0) ? 1 : 0),
  ERF: (a, c) => (a[1] ? erf(num(c, a[1])) - erf(num(c, a[0])) : erf(num(c, a[0]))),
  'ERF.PRECISE': (a, c) => erf(num(c, a[0])),
  ERFC: (a, c) => 1 - erf(num(c, a[0])),
  'ERFC.PRECISE': (a, c) => 1 - erf(num(c, a[0])),
  CONVERT: (a, c) => convert(num(c, a[0]), text(c, a[1]), text(c, a[2])),

  COMPLEX: (a, c) => {
    const unit = a[2] ? text(c, a[2]) : 'i';
    if (unit !== 'i' && unit !== 'j') return fail(err.value);
    return showCx({ re: num(c, a[0]), im: num(c, a[1]), unit });
  },
  IMREAL: (a, c) => cx(c, a[0]).re,
  IMAGINARY: (a, c) => cx(c, a[0]).im,
  IMABS: (a, c) => abs(cx(c, a[0])),
  IMARGUMENT: (a, c) => {
    const z = cx(c, a[0]);
    return z.re === 0 && z.im === 0 ? fail(err.div0) : arg(z);
  },
  IMCONJUGATE: (a, c) => {
    const z = cx(c, a[0]);
    return showCx({ ...z, im: -z.im });
  },
  IMSUM: (a, c) => {
    const zs = cxList(c, a);
    return showCx(zs.reduce((p, z) => ({ re: p.re + z.re, im: p.im + z.im, unit: z.unit }), { re: 0, im: 0, unit: zs[0]?.unit ?? 'i' }));
  },
  IMSUB: (a, c) => {
    const x = cx(c, a[0]);
    const y = cx(c, a[1]);
    return showCx({ re: x.re - y.re, im: x.im - y.im, unit: x.unit });
  },
  IMPRODUCT: (a, c) => {
    const zs = cxList(c, a);
    return showCx(zs.reduce((p, z) => mul(p, z), { re: 1, im: 0, unit: zs[0]?.unit ?? 'i' }));
  },
  IMDIV: (a, c) => {
    const x = cx(c, a[0]);
    const y = cx(c, a[1]);
    const d = y.re * y.re + y.im * y.im;
    if (d === 0) return fail(err.num);
    return showCx({ re: (x.re * y.re + x.im * y.im) / d, im: (x.im * y.re - x.re * y.im) / d, unit: x.unit });
  },
  IMSQRT: (a, c) => {
    const z = cx(c, a[0]);
    return showCx(polar(Math.sqrt(abs(z)), arg(z) / 2, z.unit));
  },
  IMEXP: (a, c) => {
    const z = cx(c, a[0]);
    return showCx(polar(Math.exp(z.re), z.im, z.unit));
  },
  IMLN: (a, c) => {
    const z = cx(c, a[0]);
    if (abs(z) === 0) return fail(err.num);
    return showCx({ re: Math.log(abs(z)), im: arg(z), unit: z.unit });
  },
  IMPOWER: (a, c) => {
    const z = cx(c, a[0]);
    const n = num(c, a[1]);
    return showCx(polar(abs(z) ** n, arg(z) * n, z.unit));
  },

  BASE: (a, c) => {
    const n = Math.trunc(num(c, a[0]));
    const radix = Math.trunc(num(c, a[1]));
    if (n < 0 || radix < 2 || radix > 36) return fail(err.num);
    return n.toString(radix).toUpperCase().padStart(Math.trunc(num(c, a[2], 0)), '0');
  },
  DECIMAL: (a, c) => {
    const radix = Math.trunc(num(c, a[1]));
    const s = text(c, a[0]).trim();
    if (radix < 2 || radix > 36 || ![...s.toLowerCase()].every((ch) => Number.parseInt(ch, 36) < radix)) return fail(err.num);
    return Number.parseInt(s, radix);
  },
  ROMAN: (a, c) => {
    let n = Math.trunc(num(c, a[0]));
    if (n < 0 || n > 3999) return fail(err.value);
    let out = '';
    for (const [v, s] of ROMAN) while (n >= v) (out += s), (n -= v);
    return out;
  },
  ARABIC: (a, c) => {
    const s = text(c, a[0]).trim().toUpperCase();
    const neg = s.startsWith('-');
    const t = neg ? s.slice(1) : s;
    const val: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
    let total = 0;
    for (let i = 0; i < t.length; i++) {
      const v = val[t[i]!];
      if (!v) return fail(err.value);
      const next = val[t[i + 1] ?? ''] ?? 0;
      total += v < next ? -v : v;
    }
    return neg ? -total : total;
  },

  // Math
  MULTINOMIAL: (a, c) => {
    const xs = numbers(c, a).map(Math.trunc);
    if (xs.some((x) => x < 0)) return fail(err.num);
    return Math.round(Math.exp(gammaln(sum(xs) + 1) - sum(xs.map((x) => gammaln(x + 1)))));
  },
  COMBINA: (a, c) => {
    const n = Math.trunc(num(c, a[0]));
    const k = Math.trunc(num(c, a[1]));
    if (n < 0 || k < 0) return fail(err.num);
    return Math.round(binom(n + k - 1, k));
  },
  PERMUTATIONA: (a, c) => Math.trunc(num(c, a[0])) ** Math.trunc(num(c, a[1])),
  SERIESSUM: (a, c) => {
    const x = num(c, a[0]);
    const n = num(c, a[1]);
    const m = num(c, a[2]);
    const coef = numbers(c, a.slice(3));
    return sum(coef.map((k, i) => k * x ** (n + i * m)));
  },
  SEC: (a, c) => 1 / Math.cos(num(c, a[0])),
  CSC: (a, c) => 1 / Math.sin(num(c, a[0])),
  COT: (a, c) => 1 / Math.tan(num(c, a[0])),
  ACOT: (a, c) => Math.PI / 2 - Math.atan(num(c, a[0])),
  ASINH: (a, c) => Math.asinh(num(c, a[0])),
  ACOSH: (a, c) => {
    const x = num(c, a[0]);
    return x < 1 ? fail(err.num) : Math.acosh(x);
  },
  ATANH: (a, c) => {
    const x = num(c, a[0]);
    return Math.abs(x) >= 1 ? fail(err.num) : Math.atanh(x);
  },
  GAMMA: (a, c) => gammaFn(num(c, a[0])),
  GAMMALN: (a, c) => {
    const x = num(c, a[0]);
    return x <= 0 ? fail(err.num) : gammaln(x);
  },
  'GAMMALN.PRECISE': (a, c) => {
    const x = num(c, a[0]);
    return x <= 0 ? fail(err.num) : gammaln(x);
  },

  // Statistics over everything typed (A functions)
  MAXA: (a, c) => {
    const xs = aValues(c, a);
    return xs.length ? Math.max(...xs) : 0;
  },
  MINA: (a, c) => {
    const xs = aValues(c, a);
    return xs.length ? Math.min(...xs) : 0;
  },
  STDEVA: (a, c) => Math.sqrt(varianceOf(aValues(c, a), true)),
  STDEVPA: (a, c) => Math.sqrt(varianceOf(aValues(c, a), false)),
  VARA: (a, c) => varianceOf(aValues(c, a), true),
  VARPA: (a, c) => varianceOf(aValues(c, a), false),
  KURT: (a, c) => {
    const xs = numbers(c, a);
    const n = xs.length;
    if (n < 4) return fail(err.div0);
    const m = sum(xs) / n;
    const s = Math.sqrt(varianceOf(xs, true));
    const k = sum(xs.map((x) => ((x - m) / s) ** 4));
    return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * k - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
  },
  SKEW: (a, c) => {
    const xs = numbers(c, a);
    const n = xs.length;
    if (n < 3) return fail(err.div0);
    const m = sum(xs) / n;
    const s = Math.sqrt(varianceOf(xs, true));
    return (n / ((n - 1) * (n - 2))) * sum(xs.map((x) => ((x - m) / s) ** 3));
  },
  FISHER: (a, c) => {
    const x = num(c, a[0]);
    return Math.abs(x) >= 1 ? fail(err.num) : 0.5 * Math.log((1 + x) / (1 - x));
  },
  FISHERINV: (a, c) => Math.tanh(num(c, a[0])),

  // Distributions
  'NORM.INV': (a, c) => {
    const sd = num(c, a[2]);
    positive(sd);
    return num(c, a[1]) + sd * normInv(num(c, a[0]));
  },
  'NORM.S.INV': (a, c) => normInv(num(c, a[0])),
  'LOGNORM.DIST': (a, c) => {
    const x = num(c, a[0]);
    const m = num(c, a[1]);
    const sd = num(c, a[2]);
    positive(x, sd);
    const z = (Math.log(x) - m) / sd;
    return bool(c, a[3]) ? normCdf(z) : Math.exp(-0.5 * z * z) / (x * sd * Math.sqrt(2 * Math.PI));
  },
  'LOGNORM.INV': (a, c) => {
    const sd = num(c, a[2]);
    positive(sd);
    return Math.exp(num(c, a[1]) + sd * normInv(num(c, a[0])));
  },
  'EXPON.DIST': (a, c) => {
    const x = num(c, a[0]);
    const l = num(c, a[1]);
    if (x < 0 || l <= 0) return fail(err.num);
    return bool(c, a[2]) ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x);
  },
  'POISSON.DIST': (a, c) => {
    const k = Math.trunc(num(c, a[0]));
    const m = num(c, a[1]);
    if (k < 0 || m < 0) return fail(err.num);
    const pmf = (i: number) => Math.exp(-m + i * Math.log(m) - gammaln(i + 1));
    if (!bool(c, a[2])) return pmf(k);
    let s = 0;
    for (let i = 0; i <= k; i++) s += pmf(i);
    return s;
  },
  'BINOM.DIST': (a, c) => {
    const k = Math.trunc(num(c, a[0]));
    const n = Math.trunc(num(c, a[1]));
    const p = num(c, a[2]);
    if (k < 0 || k > n || p < 0 || p > 1) return fail(err.num);
    const pmf = (i: number) => binom(n, i) * p ** i * (1 - p) ** (n - i);
    if (!bool(c, a[3])) return pmf(k);
    let s = 0;
    for (let i = 0; i <= k; i++) s += pmf(i);
    return s;
  },
  'T.DIST': (a, c) => {
    const x = num(c, a[0]);
    const v = num(c, a[1]);
    if (v < 1) return fail(err.num);
    return bool(c, a[2]) ? tCdf(x, v) : tPdf(x, v);
  },
  'T.DIST.2T': (a, c) => {
    const x = num(c, a[0]);
    const v = num(c, a[1]);
    if (x < 0 || v < 1) return fail(err.num);
    return betaI(v / (v + x * x), v / 2, 0.5);
  },
  'T.DIST.RT': (a, c) => {
    const v = num(c, a[1]);
    if (v < 1) return fail(err.num);
    return 1 - tCdf(num(c, a[0]), v);
  },
  'T.INV.2T': (a, c) => {
    const p = num(c, a[0]);
    const v = num(c, a[1]);
    if (p <= 0 || p > 1 || v < 1) return fail(err.num);
    return tInv2(p, v);
  },
  'T.INV': (a, c) => {
    const p = num(c, a[0]);
    const v = num(c, a[1]);
    if (p <= 0 || p >= 1 || v < 1) return fail(err.num);
    const x = tInv2(2 * Math.min(p, 1 - p), v);
    return p < 0.5 ? -x : x;
  },
  'CHISQ.DIST': (a, c) => {
    const x = num(c, a[0]);
    const k = num(c, a[1]);
    if (x < 0 || k < 1) return fail(err.num);
    if (bool(c, a[2])) return gammaP(k / 2, x / 2);
    return Math.exp((k / 2 - 1) * Math.log(x) - x / 2 - (k / 2) * Math.LN2 - gammaln(k / 2));
  },
  'CHISQ.DIST.RT': (a, c) => {
    const x = num(c, a[0]);
    const k = num(c, a[1]);
    if (x < 0 || k < 1) return fail(err.num);
    return 1 - gammaP(k / 2, x / 2);
  },
  'CHISQ.INV.RT': (a, c) => {
    const p = num(c, a[0]);
    const k = num(c, a[1]);
    if (p <= 0 || p > 1 || k < 1) return fail(err.num);
    return bisect((x) => gammaP(k / 2, x / 2), 1 - p, 0, 1e7);
  },
  'F.DIST': (a, c) => {
    const x = num(c, a[0]);
    const d1 = num(c, a[1]);
    const d2 = num(c, a[2]);
    if (x < 0 || d1 < 1 || d2 < 1) return fail(err.num);
    if (bool(c, a[3])) return betaI((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2);
    return Math.exp(0.5 * (d1 * Math.log(d1 * x) + d2 * Math.log(d2) - (d1 + d2) * Math.log(d1 * x + d2)) - Math.log(x) - (gammaln(d1 / 2) + gammaln(d2 / 2) - gammaln((d1 + d2) / 2)));
  },
  'F.DIST.RT': (a, c) => {
    const x = num(c, a[0]);
    const d1 = num(c, a[1]);
    const d2 = num(c, a[2]);
    if (x < 0 || d1 < 1 || d2 < 1) return fail(err.num);
    return 1 - betaI((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2);
  },
  'GAMMA.DIST': (a, c) => {
    const x = num(c, a[0]);
    const al = num(c, a[1]);
    const be = num(c, a[2]);
    if (x < 0) return fail(err.num);
    positive(al, be);
    if (bool(c, a[3])) return gammaP(al, x / be);
    return Math.exp((al - 1) * Math.log(x) - x / be - gammaln(al) - al * Math.log(be));
  },
  'BETA.DIST': (a, c) => {
    const lo = num(c, a[4], 0);
    const hi = num(c, a[5], 1);
    const x = (num(c, a[0]) - lo) / (hi - lo);
    const al = num(c, a[1]);
    const be = num(c, a[2]);
    positive(al, be);
    if (x < 0 || x > 1) return fail(err.num);
    if (bool(c, a[3])) return betaI(x, al, be);
    return Math.exp((al - 1) * Math.log(x) + (be - 1) * Math.log(1 - x) - (gammaln(al) + gammaln(be) - gammaln(al + be))) / (hi - lo);
  },
  'CONFIDENCE.NORM': (a, c) => {
    const alpha = num(c, a[0]);
    const sd = num(c, a[1]);
    const n = Math.trunc(num(c, a[2]));
    if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 1) return fail(err.num);
    return (normInv(1 - alpha / 2) * sd) / Math.sqrt(n);
  },
  'CONFIDENCE.T': (a, c) => {
    const alpha = num(c, a[0]);
    const sd = num(c, a[1]);
    const n = Math.trunc(num(c, a[2]));
    if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 2) return fail(err.div0);
    return (tInv2(alpha, n - 1) * sd) / Math.sqrt(n);
  },

  // Information and web
  ISREF: (a) => a[0]?.kind === 'ref' || a[0]?.kind === 'range',
  ENCODEURL: (a, c) => encodeURIComponent(text(c, a[0])),
};

const ALIASES: Record<string, string> = {
  NORMINV: 'NORM.INV',
  NORMSINV: 'NORM.S.INV',
  LOGNORMDIST: 'LOGNORM.DIST',
  LOGINV: 'LOGNORM.INV',
  EXPONDIST: 'EXPON.DIST',
  POISSON: 'POISSON.DIST',
  BINOMDIST: 'BINOM.DIST',
  CHIDIST: 'CHISQ.DIST.RT',
  CHIINV: 'CHISQ.INV.RT',
  FDIST: 'F.DIST.RT',
  GAMMADIST: 'GAMMA.DIST',
  CONFIDENCE: 'CONFIDENCE.NORM',
  TINV: 'T.INV.2T',
};

Object.assign(FUNCTIONS, ENGINEERING);
for (const [alias, target] of Object.entries(ALIASES)) FUNCTIONS[alias] ??= FUNCTIONS[target]!;
// TDIST(x, df, tails) is the older two-in-one of T.DIST.RT and T.DIST.2T.
FUNCTIONS.TDIST ??= (a, c) => (num(c, a[2]) === 2 ? ENGINEERING['T.DIST.2T']!(a, c) : ENGINEERING['T.DIST.RT']!(a, c));
