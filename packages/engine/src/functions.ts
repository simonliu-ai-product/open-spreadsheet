import { dateToSerial, serialFromJsDate, serialToDate, weekday } from './dates.ts';
import { formatValue } from './format.ts';
import type { Node } from './parse.ts';
import {
  CellError,
  compare,
  err,
  ErrorSignal,
  fail,
  flatten,
  formatGeneral,
  isError,
  isMatrix,
  type Matrix,
  parseNumericText,
  roundTo15,
  type Scalar,
  scalarOf,
  toBoolean,
  toNumber,
  toText,
  type Value,
} from './values.ts';

/** What a function needs from the evaluator. */
export interface CallContext {
  evaluate(node: Node): Value;
  /** The row and column of a reference argument, for ROW() and COLUMN(). */
  position(node: Node): { row: number; col: number; rows: number; cols: number } | null;
  /** The cell being computed, for ROW() and COLUMN() with no argument. */
  here: { row: number; col: number };
  now(): Date;
}

export type FunctionImpl = (args: Node[], ctx: CallContext) => Value;

export const scalarArg = (ctx: CallContext, node: Node | undefined): Scalar =>
  node === undefined || node.kind === 'missing' ? null : scalarOf(ctx.evaluate(node));
export const num = (ctx: CallContext, node: Node | undefined, fallback?: number): number => {
  if ((node === undefined || node.kind === 'missing') && fallback !== undefined) return fallback;
  return toNumber(scalarArg(ctx, node));
};
export const text = (ctx: CallContext, node: Node | undefined) => toText(scalarArg(ctx, node));
export const asMatrix = (v: Value): Matrix => (isMatrix(v) ? v : [[v]]);
export const isRefNode = (n: Node) => n.kind === 'ref' || n.kind === 'range' || n.kind === 'call' || n.kind === 'name';

/**
 * The numbers an aggregate sees. Inside a range, text, booleans and blanks are
 * skipped; typed directly as an argument they are coerced (SUM("3",TRUE) is 4).
 */
export function numbers(ctx: CallContext, args: Node[]): number[] {
  const out: number[] = [];
  for (const a of args) {
    if (a.kind === 'missing') continue;
    const v = ctx.evaluate(a);
    if (isMatrix(v) || isRefNode(a)) {
      for (const s of flatten(v)) {
        if (isError(s)) fail(s);
        if (typeof s === 'number') out.push(s);
      }
    } else {
      out.push(toNumber(v));
    }
  }
  return out;
}

export const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);

/** ROUND-family helper: scale, clean the binary noise to 15 digits, then cut. */
export function roundWith(n: number, digits: number, cut: (x: number) => number): number {
  const f = 10 ** Math.trunc(digits);
  const scaled = roundTo15(Math.abs(n) * f);
  return (Math.sign(n) * cut(scaled)) / f;
}

// ---------- criteria (COUNTIF, SUMIF, ...) ----------

export function wildcardRegex(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i]!;
    if (c === '~' && i + 1 < pattern.length) re += pattern[++i]!.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (c === '*') re += '.*';
    else if (c === '?') re += '.';
    else re += c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'is');
}

/** A predicate for an Excel criterion: 5, ">=10", "<>done", "app*", "". */
export function criterion(c: Scalar): (v: Scalar) => boolean {
  if (typeof c === 'number' || typeof c === 'boolean') return (v) => v !== null && !isError(v) && compare(v, c) === 0;
  if (isError(c)) return (v) => isError(v) && v.code === c.code;
  const s = c ?? '';
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s)!;
  const op = m[1] ?? '=';
  const rest = m[2]!;
  const n = parseNumericText(rest);
  if (rest === '') {
    if (op === '=') return (v) => v === null || v === '';
    if (op === '<>') return (v) => v !== null && v !== '';
  }
  const target: Scalar = n !== null ? n : rest.toUpperCase() === 'TRUE' ? true : rest.toUpperCase() === 'FALSE' ? false : rest;
  if (op === '=' || op === '<>') {
    const eq =
      typeof target === 'string'
        ? (v: Scalar) => typeof v === 'string' && wildcardRegex(target).test(v)
        : (v: Scalar) => v !== null && !isError(v) && typeof v === typeof target && compare(v, target) === 0;
    return op === '=' ? eq : (v) => !eq(v);
  }
  return (v) => {
    if (v === null || isError(v) || typeof v !== typeof target) return false;
    const c2 = compare(v, target);
    return op === '<' ? c2 < 0 : op === '<=' ? c2 <= 0 : op === '>' ? c2 > 0 : c2 >= 0;
  };
}

/** Cells of `ranges[i]` at each position where every criterion holds. */
export function matchingPositions(ctx: CallContext, pairs: [Node, Node][]): { matrices: Matrix[]; hits: [number, number][] } {
  const matrices = pairs.map(([r]) => asMatrix(ctx.evaluate(r)));
  const tests = pairs.map(([, c]) => criterion(scalarArg(ctx, c)));
  const first = matrices[0]!;
  for (const m of matrices) {
    if (m.length !== first.length || (m[0]?.length ?? 0) !== (first[0]?.length ?? 0)) fail(err.value);
  }
  const hits: [number, number][] = [];
  first.forEach((row, r) =>
    row.forEach((_, c) => {
      if (tests.every((t, i) => t(matrices[i]![r]![c] ?? null))) hits.push([r, c]);
    }),
  );
  return { matrices, hits };
}

export const pairsFrom = (args: Node[], offset: number): [Node, Node][] => {
  const out: [Node, Node][] = [];
  for (let i = offset; i + 1 < args.length; i += 2) out.push([args[i]!, args[i + 1]!]);
  return out;
};

// ---------- lookup helpers ----------

export function findExact(values: Scalar[], x: Scalar): number {
  if (typeof x === 'string' && /[*?~]/.test(x)) {
    const re = wildcardRegex(x);
    return values.findIndex((v) => typeof v === 'string' && re.test(v));
  }
  return values.findIndex((v) => v !== null && !isError(v) && typeof v === typeof x && compare(v, x) === 0);
}

/** Approximate match on ascending data: the last position whose value is <= x. */
export function findApprox(values: Scalar[], x: Scalar): number {
  let found = -1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i]!;
    if (v === null || isError(v) || typeof v !== typeof x) continue;
    if (compare(v, x) <= 0) found = i;
    else break;
  }
  return found;
}

export const column = (m: Matrix, c: number) => m.map((row) => row[c] ?? null);
export const lineOf = (m: Matrix) => (m.length === 1 ? m[0]! : column(m, 0));

// ---------- TEXT() ----------

// ---------- finance ----------

export function pmt(rate: number, nper: number, pv: number, fv = 0, type = 0): number {
  if (rate === 0) return -(pv + fv) / nper;
  const f = (1 + rate) ** nper;
  return (-rate * (fv + pv * f)) / ((1 + rate * type) * (f - 1));
}

// ---------- the table ----------

export const FUNCTIONS: Record<string, FunctionImpl> = {
  // Math
  SUM: (a, c) => sum(numbers(c, a)),
  PRODUCT: (a, c) => numbers(c, a).reduce((x, y) => x * y, 1),
  ABS: (a, c) => Math.abs(num(c, a[0])),
  INT: (a, c) => Math.floor(num(c, a[0])),
  TRUNC: (a, c) => roundWith(num(c, a[0]), num(c, a[1], 0), Math.floor),
  ROUND: (a, c) => roundWith(num(c, a[0]), num(c, a[1]), (x) => Math.floor(x + 0.5)),
  ROUNDUP: (a, c) => roundWith(num(c, a[0]), num(c, a[1]), Math.ceil),
  ROUNDDOWN: (a, c) => roundWith(num(c, a[0]), num(c, a[1]), Math.floor),
  MOD: (a, c) => {
    const n = num(c, a[0]);
    const d = num(c, a[1]);
    if (d === 0) return err.div0;
    return roundTo15(n - d * Math.floor(n / d));
  },
  POWER: (a, c) => {
    const r = num(c, a[0]) ** num(c, a[1]);
    return Number.isFinite(r) ? r : err.num;
  },
  SQRT: (a, c) => {
    const n = num(c, a[0]);
    return n < 0 ? err.num : Math.sqrt(n);
  },
  PI: () => Math.PI,
  SIGN: (a, c) => Math.sign(num(c, a[0])),
  SUMPRODUCT: (a, c) => {
    const ms = a.map((n) => asMatrix(c.evaluate(n)));
    const rows = ms[0]!.length;
    const cols = ms[0]![0]?.length ?? 0;
    if (ms.some((m) => m.length !== rows || (m[0]?.length ?? 0) !== cols)) return err.value;
    let total = 0;
    for (let r = 0; r < rows; r++) {
      for (let k = 0; k < cols; k++) {
        total += ms.reduce((p, m) => {
          const v = m[r]![k] ?? null;
          if (isError(v)) fail(v);
          return p * (typeof v === 'number' ? v : 0);
        }, 1);
      }
    }
    return total;
  },
  SUMIF: (a, c) => {
    const { matrices, hits } = matchingPositions(c, [[a[0]!, a[1]!]]);
    const target = a[2] ? asMatrix(c.evaluate(a[2])) : matrices[0]!;
    return sum(hits.map(([r, k]) => target[r]?.[k] ?? null).filter((v): v is number => typeof v === 'number'));
  },
  SUMIFS: (a, c) => {
    const { hits } = matchingPositions(c, pairsFrom(a, 1));
    const target = asMatrix(c.evaluate(a[0]!));
    return sum(hits.map(([r, k]) => target[r]?.[k] ?? null).filter((v): v is number => typeof v === 'number'));
  },

  // Statistics
  COUNT: (a, c) => {
    let n = 0;
    for (const arg of a) {
      const v = c.evaluate(arg);
      for (const s of flatten(v)) {
        if (typeof s === 'number') n++;
        else if (!isMatrix(v) && !isRefNode(arg) && typeof s === 'string' && parseNumericText(s) !== null) n++;
      }
    }
    return n;
  },
  COUNTA: (a, c) => {
    let n = 0;
    for (const arg of a) for (const s of flatten(c.evaluate(arg))) if (s !== null) n++;
    return n;
  },
  COUNTBLANK: (a, c) => {
    let n = 0;
    for (const s of flatten(c.evaluate(a[0]!))) if (s === null || s === '') n++;
    return n;
  },
  COUNTIF: (a, c) => matchingPositions(c, [[a[0]!, a[1]!]]).hits.length,
  COUNTIFS: (a, c) => matchingPositions(c, pairsFrom(a, 0)).hits.length,
  AVERAGE: (a, c) => {
    const xs = numbers(c, a);
    return xs.length ? sum(xs) / xs.length : err.div0;
  },
  AVERAGEIF: (a, c) => {
    const { matrices, hits } = matchingPositions(c, [[a[0]!, a[1]!]]);
    const target = a[2] ? asMatrix(c.evaluate(a[2])) : matrices[0]!;
    const xs = hits.map(([r, k]) => target[r]?.[k] ?? null).filter((v): v is number => typeof v === 'number');
    return xs.length ? sum(xs) / xs.length : err.div0;
  },
  MIN: (a, c) => {
    const xs = numbers(c, a);
    return xs.length ? Math.min(...xs) : 0;
  },
  MAX: (a, c) => {
    const xs = numbers(c, a);
    return xs.length ? Math.max(...xs) : 0;
  },
  MEDIAN: (a, c) => {
    const xs = numbers(c, a).sort((x, y) => x - y);
    if (!xs.length) return err.num;
    const mid = Math.floor(xs.length / 2);
    return xs.length % 2 ? xs[mid]! : (xs[mid - 1]! + xs[mid]!) / 2;
  },
  LARGE: (a, c) => {
    const xs = numbers(c, [a[0]!]).sort((x, y) => y - x);
    const k = Math.ceil(num(c, a[1]));
    return k >= 1 && k <= xs.length ? xs[k - 1]! : err.num;
  },
  SMALL: (a, c) => {
    const xs = numbers(c, [a[0]!]).sort((x, y) => x - y);
    const k = Math.ceil(num(c, a[1]));
    return k >= 1 && k <= xs.length ? xs[k - 1]! : err.num;
  },
  'STDEV.S': (a, c) => {
    const xs = numbers(c, a);
    if (xs.length < 2) return err.div0;
    const mean = sum(xs) / xs.length;
    return Math.sqrt(sum(xs.map((x) => (x - mean) ** 2)) / (xs.length - 1));
  },

  // Logic (IF and friends are lazy: only the branch taken is evaluated)
  IF: (a, c) => {
    const cond = toBoolean(scalarArg(c, a[0]));
    const branch = cond ? a[1] : a[2];
    if (branch === undefined) return cond ? true : false;
    return branch.kind === 'missing' ? 0 : c.evaluate(branch);
  },
  IFS: (a, c) => {
    for (let i = 0; i + 1 < a.length; i += 2) if (toBoolean(scalarArg(c, a[i]))) return c.evaluate(a[i + 1]!);
    return err.na;
  },
  IFERROR: (a, c) => {
    const v = c.evaluate(a[0]!);
    return isError(scalarOf(v)) ? c.evaluate(a[1]!) : v;
  },
  IFNA: (a, c) => {
    const v = scalarOf(c.evaluate(a[0]!));
    return isError(v) && v.code === '#N/A' ? c.evaluate(a[1]!) : v;
  },
  AND: (a, c) => a.every((n) => [...flatten(c.evaluate(n))].filter((s) => s !== null && typeof s !== 'string').every(toBoolean)),
  OR: (a, c) => a.some((n) => [...flatten(c.evaluate(n))].filter((s) => s !== null && typeof s !== 'string').some(toBoolean)),
  XOR: (a, c) => a.flatMap((n) => [...flatten(c.evaluate(n))]).filter((s) => s !== null && typeof s !== 'string').filter(toBoolean).length % 2 === 1,
  NOT: (a, c) => !toBoolean(scalarArg(c, a[0])),
  TRUE: () => true,
  FALSE: () => false,
  SWITCH: (a, c) => {
    const v = scalarArg(c, a[0]);
    for (let i = 1; i + 1 < a.length; i += 2) if (compare(v, scalarArg(c, a[i])) === 0) return c.evaluate(a[i + 1]!);
    return a.length % 2 === 0 ? c.evaluate(a[a.length - 1]!) : err.na;
  },

  // Text
  CONCAT: (a, c) => a.flatMap((n) => [...flatten(c.evaluate(n))].map(toText)).join(''),
  CONCATENATE: (a, c) => a.map((n) => text(c, n)).join(''),
  TEXTJOIN: (a, c) => {
    const sep = text(c, a[0]);
    const skip = toBoolean(scalarArg(c, a[1]));
    return a
      .slice(2)
      .flatMap((n) => [...flatten(c.evaluate(n))].map(toText))
      .filter((s) => !(skip && s === ''))
      .join(sep);
  },
  LEFT: (a, c) => [...text(c, a[0])].slice(0, num(c, a[1], 1)).join(''),
  RIGHT: (a, c) => {
    const chars = [...text(c, a[0])];
    const n = num(c, a[1], 1);
    return n <= 0 ? '' : chars.slice(-n).join('');
  },
  MID: (a, c) => {
    const start = num(c, a[1]);
    const n = num(c, a[2]);
    if (start < 1 || n < 0) return err.value;
    return [...text(c, a[0])].slice(start - 1, start - 1 + n).join('');
  },
  LEN: (a, c) => [...text(c, a[0])].length,
  UPPER: (a, c) => text(c, a[0]).toUpperCase(),
  LOWER: (a, c) => text(c, a[0]).toLowerCase(),
  PROPER: (a, c) => text(c, a[0]).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (_, p, l: string) => p + l.toUpperCase()),
  TRIM: (a, c) => text(c, a[0]).replace(/ +/g, ' ').trim(),
  REPT: (a, c) => text(c, a[0]).repeat(Math.max(0, Math.floor(num(c, a[1])))),
  EXACT: (a, c) => text(c, a[0]) === text(c, a[1]),
  SUBSTITUTE: (a, c) => {
    const s = text(c, a[0]);
    const from = text(c, a[1]);
    const to = text(c, a[2]);
    if (from === '') return s;
    if (a[3] === undefined) return s.split(from).join(to);
    const nth = num(c, a[3]);
    let seen = 0;
    return s.replace(new RegExp(from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), (m) => (++seen === nth ? to : m));
  },
  REPLACE: (a, c) => {
    const chars = [...text(c, a[0])];
    const start = num(c, a[1]);
    const n = num(c, a[2]);
    chars.splice(start - 1, n, text(c, a[3]));
    return chars.join('');
  },
  FIND: (a, c) => {
    const i = text(c, a[1]).indexOf(text(c, a[0]), num(c, a[2], 1) - 1);
    return i < 0 ? err.value : i + 1;
  },
  SEARCH: (a, c) => {
    const hay = text(c, a[1]);
    const from = num(c, a[2], 1) - 1;
    const re = new RegExp(wildcardRegex(text(c, a[0])).source.slice(1, -1), 'is');
    const m = re.exec(hay.slice(from));
    return m ? from + m.index + 1 : err.value;
  },
  VALUE: (a, c) => {
    const v = scalarArg(c, a[0]);
    if (typeof v === 'number') return v;
    const n = parseNumericText(toText(v));
    return n === null ? err.value : n;
  },
  TEXT: (a, c) => {
    const v = scalarArg(c, a[0]);
    const fmt = text(c, a[1]);
    if (typeof v !== 'number') {
      const n = typeof v === 'string' ? parseNumericText(v) : null;
      if (n === null) return formatValue(v, fmt).text;
      return formatValue(n, fmt).text;
    }
    return formatValue(v, fmt).text;
  },

  // Lookup and reference
  VLOOKUP: (a, c) => {
    const x = scalarArg(c, a[0]);
    const table = asMatrix(c.evaluate(a[1]!));
    const col = num(c, a[2]);
    const approx = a[3] === undefined || a[3].kind === 'missing' ? true : toBoolean(scalarArg(c, a[3]));
    if (col < 1 || col > (table[0]?.length ?? 0)) return err.ref;
    const keys = column(table, 0);
    const i = approx ? findApprox(keys, x) : findExact(keys, x);
    return i < 0 ? err.na : (table[i]![col - 1] ?? null);
  },
  HLOOKUP: (a, c) => {
    const x = scalarArg(c, a[0]);
    const table = asMatrix(c.evaluate(a[1]!));
    const row = num(c, a[2]);
    const approx = a[3] === undefined || a[3].kind === 'missing' ? true : toBoolean(scalarArg(c, a[3]));
    if (row < 1 || row > table.length) return err.ref;
    const keys = table[0]!;
    const i = approx ? findApprox(keys, x) : findExact(keys, x);
    return i < 0 ? err.na : (table[row - 1]![i] ?? null);
  },
  MATCH: (a, c) => {
    const x = scalarArg(c, a[0]);
    const line = lineOf(asMatrix(c.evaluate(a[1]!)));
    const type = num(c, a[2], 1);
    let i: number;
    if (type === 0) i = findExact(line, x);
    else if (type > 0) i = findApprox(line, x);
    else {
      i = -1;
      for (let k = 0; k < line.length; k++) {
        const v = line[k]!;
        if (v !== null && !isError(v) && typeof v === typeof x && compare(v, x) >= 0) i = k;
        else break;
      }
    }
    return i < 0 ? err.na : i + 1;
  },
  XLOOKUP: (a, c) => {
    const x = scalarArg(c, a[0]);
    const lookup = lineOf(asMatrix(c.evaluate(a[1]!)));
    const result = asMatrix(c.evaluate(a[2]!));
    const mode = num(c, a[4], 0);
    const reverse = num(c, a[5], 1) === -1;
    const order = [...lookup.keys()];
    if (reverse) order.reverse();
    let hit = -1;
    let best: Scalar = null;
    for (const i of order) {
      const v = lookup[i]!;
      if (v === null || isError(v) || typeof v !== typeof x) continue;
      const cmp = compare(v, x);
      if (cmp === 0) {
        hit = i;
        break;
      }
      if (mode === 2 && typeof x === 'string' && wildcardRegex(x).test(String(v))) {
        hit = i;
        break;
      }
      if ((mode === -1 && cmp < 0 && (best === null || compare(v, best) > 0)) || (mode === 1 && cmp > 0 && (best === null || compare(v, best) < 0))) {
        best = v;
        hit = i;
      }
    }
    if (hit < 0) return a[3] && a[3].kind !== 'missing' ? c.evaluate(a[3]) : err.na;
    return result.length === 1 ? (result[0]![hit] ?? null) : (result[hit]?.[0] ?? null);
  },
  INDEX: (a, c) => {
    const m = asMatrix(c.evaluate(a[0]!));
    let r = num(c, a[1], 0);
    let k = num(c, a[2], 0);
    if (m.length === 1 && (a[2] === undefined || a[2].kind === 'missing')) {
      k = r;
      r = 1;
    }
    if (r < 0 || k < 0 || r > m.length || k > (m[0]?.length ?? 0)) return err.ref;
    if (r === 0) return m.map((row) => [row[k - 1] ?? null]);
    if (k === 0) return [m[r - 1]!];
    return m[r - 1]![k - 1] ?? null;
  },
  CHOOSE: (a, c) => {
    const i = Math.floor(num(c, a[0]));
    return i < 1 || i >= a.length ? err.value : c.evaluate(a[i]!);
  },
  ROWS: (a, c) => asMatrix(c.evaluate(a[0]!)).length,
  COLUMNS: (a, c) => asMatrix(c.evaluate(a[0]!))[0]?.length ?? 0,
  ROW: (a, c) => (a[0] ? (c.position(a[0])?.row ?? fail(err.value)) : c.here.row) + 1,
  COLUMN: (a, c) => (a[0] ? (c.position(a[0])?.col ?? fail(err.value)) : c.here.col) + 1,

  // Date and time
  DATE: (a, c) => {
    const y = Math.floor(num(c, a[0]));
    const year = y < 1900 ? y + 1900 : y;
    const s = dateToSerial(year, Math.floor(num(c, a[1])), Math.floor(num(c, a[2])));
    return s < 0 ? err.num : s;
  },
  YEAR: (a, c) => serialToDate(num(c, a[0])).year,
  MONTH: (a, c) => serialToDate(num(c, a[0])).month,
  DAY: (a, c) => serialToDate(num(c, a[0])).day,
  WEEKDAY: (a, c) => {
    const w = weekday(num(c, a[0]));
    const type = num(c, a[1], 1);
    if (type === 1) return w;
    if (type === 2) return ((w + 5) % 7) + 1;
    if (type === 3) return (w + 5) % 7;
    return err.num;
  },
  TODAY: (_a, c) => Math.floor(serialFromJsDate(c.now())),
  NOW: (_a, c) => serialFromJsDate(c.now()),
  EDATE: (a, c) => {
    const { year, month, day } = serialToDate(num(c, a[0]));
    const months = Math.trunc(num(c, a[1]));
    const last = new Date(Date.UTC(year, month - 1 + months + 1, 0)).getUTCDate();
    return dateToSerial(year, month + months, Math.min(day, last));
  },
  EOMONTH: (a, c) => {
    const { year, month } = serialToDate(num(c, a[0]));
    return dateToSerial(year, month + Math.trunc(num(c, a[1])) + 1, 0);
  },
  DAYS: (a, c) => Math.floor(num(c, a[0])) - Math.floor(num(c, a[1])),

  // Information
  ISBLANK: (a, c) => scalarArg(c, a[0]) === null,
  ISNUMBER: (a, c) => typeof scalarArg(c, a[0]) === 'number',
  ISTEXT: (a, c) => typeof scalarArg(c, a[0]) === 'string',
  ISLOGICAL: (a, c) => typeof scalarArg(c, a[0]) === 'boolean',
  ISERROR: (a, c) => isError(scalarArg(c, a[0])),
  ISNA: (a, c) => {
    const v = scalarArg(c, a[0]);
    return isError(v) && v.code === '#N/A';
  },
  NA: () => err.na,

  // Finance
  PMT: (a, c) => pmt(num(c, a[0]), num(c, a[1]), num(c, a[2]), num(c, a[3], 0), num(c, a[4], 0)),
  FV: (a, c) => {
    const r = num(c, a[0]);
    const n = num(c, a[1]);
    const p = num(c, a[2]);
    const pv = num(c, a[3], 0);
    const type = num(c, a[4], 0);
    if (r === 0) return -(pv + p * n);
    const f = (1 + r) ** n;
    return -(pv * f + (p * (1 + r * type) * (f - 1)) / r);
  },
  PV: (a, c) => {
    const r = num(c, a[0]);
    const n = num(c, a[1]);
    const p = num(c, a[2]);
    const fv = num(c, a[3], 0);
    const type = num(c, a[4], 0);
    if (r === 0) return -(fv + p * n);
    const f = (1 + r) ** n;
    return -(fv + (p * (1 + r * type) * (f - 1)) / r) / f;
  },
  NPV: (a, c) => {
    const r = num(c, a[0]);
    return numbers(c, a.slice(1)).reduce((acc, v, i) => acc + v / (1 + r) ** (i + 1), 0);
  },
};

/** Runs a function, turning a thrown error value into a returned one. */
export function callFunction(name: string, args: Node[], ctx: CallContext): Value {
  const fn = FUNCTIONS[name];
  if (!fn) return err.name;
  try {
    const v = fn(args, ctx);
    if (typeof v === 'number' && !Number.isFinite(v)) return err.num;
    return v;
  } catch (e) {
    if (e instanceof ErrorSignal) return e.error;
    throw e;
  }
}

export { CellError, formatGeneral };
