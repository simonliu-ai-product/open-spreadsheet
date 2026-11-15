/**
 * Excel's value model. A cell holds a number, text, a boolean, an error, or
 * nothing; a range evaluates to a matrix of those. Every coercion rule here
 * follows Excel's, because that is where most compatibility bugs come from.
 */

import { dateToSerial } from './dates.ts';

export const ERROR_CODES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!'] as const;
export type ErrorCode = (typeof ERROR_CODES)[number] | '#CIRC!';

export class CellError {
  readonly code: ErrorCode;
  constructor(code: ErrorCode) {
    this.code = code;
  }
  toString(): string {
    return this.code;
  }
}

export const err = {
  null: new CellError('#NULL!'),
  div0: new CellError('#DIV/0!'),
  value: new CellError('#VALUE!'),
  ref: new CellError('#REF!'),
  name: new CellError('#NAME?'),
  num: new CellError('#NUM!'),
  na: new CellError('#N/A'),
  circ: new CellError('#CIRC!'),
};

/** One cell's value. `null` is a blank cell, distinct from the empty string. */
export type Scalar = number | string | boolean | CellError | null;
/** A range or array: rows of scalars. */
export type Matrix = Scalar[][];
export type Value = Scalar | Matrix;

export const isError = (v: unknown): v is CellError => v instanceof CellError;
export const isMatrix = (v: Value): v is Matrix => Array.isArray(v);

/** Thrown inside a function to return an error value. */
export class ErrorSignal extends Error {
  readonly error: CellError;
  constructor(error: CellError) {
    super(error.code);
    this.error = error;
  }
}
export const fail = (e: CellError): never => {
  throw new ErrorSignal(e);
};

/** The top-left value of a range, as Excel uses when a range meets a scalar operator. */
export function scalarOf(v: Value): Scalar {
  if (!isMatrix(v)) return v;
  return v[0]?.[0] ?? null;
}

const NUMERIC = /^[+-]?(\d+(\.\d*)?|\.\d+)(e[+-]?\d+)?%?$/i;
const DATE = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
const TIME = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(am|pm)?$/i;

/**
 * Text that Excel reads as a number: "12", "1,000", "-3.5e2", "50%", "$1,200",
 * and dates and times ("2026-01-05", "2026/1/5 09:30", "14:30") as serials.
 */
export function parseNumericText(text: string): number | null {
  let t = text.trim().replace(/,(?=\d{3}(\D|$))/g, '');
  const date = DATE.exec(t);
  if (date) {
    const [, y, mo, d, h, mi, se] = date;
    const month = Number(mo);
    const day = Number(d);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    const serial = dateToSerial(Number(y), month, day);
    return h === undefined ? serial : serial + (Number(h) * 3600 + Number(mi) * 60 + Number(se ?? 0)) / 86400;
  }
  const time = TIME.exec(t);
  if (time) {
    let h = Number(time[1]);
    const m = Number(time[2]);
    const sec = Number(time[3] ?? 0);
    const ap = time[4]?.toLowerCase();
    if (ap) {
      if (h < 1 || h > 12) return null;
      h = (h % 12) + (ap === 'pm' ? 12 : 0);
    }
    if (m > 59 || sec >= 60) return null;
    return (h * 3600 + m * 60 + sec) / 86400;
  }
  let negative = false;
  if (/^-?\$/.test(t)) {
    negative = t.startsWith('-');
    t = t.replace(/^-?\$/, '');
  }
  if (t === '' || !NUMERIC.test(t)) return null;
  const pct = t.endsWith('%');
  const n = Number(pct ? t.slice(0, -1) : t) * (negative ? -1 : 1);
  return Number.isFinite(n) ? (pct ? n / 100 : n) : null;
}

/** Coercion for arithmetic: blanks are 0, TRUE is 1, numeric text is its number. */
export function toNumber(v: Scalar): number {
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return fail(v);
  const n = parseNumericText(v);
  return n === null ? fail(err.value) : n;
}

/** Excel's General format for a number: at most 15 significant digits, no trailing zeros. */
export function formatGeneral(n: number): string {
  const p = roundTo15(n);
  const abs = Math.abs(p);
  // Thresholds to verify against Excel with the oracle corpus (PLAN.md §8).
  if (abs !== 0 && (abs >= 1e15 || abs < 1e-9)) {
    const [mantissa, exponent] = p.toExponential().split('e');
    const e = Number(exponent);
    return `${mantissa}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
  }
  return String(p);
}

export function toText(v: Scalar): string {
  if (typeof v === 'string') return v;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return fail(v);
  return formatGeneral(v);
}

export function toBoolean(v: Scalar): boolean {
  if (typeof v === 'boolean') return v;
  if (v === null) return false;
  if (typeof v === 'number') return v !== 0;
  if (isError(v)) return fail(v);
  const u = v.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return fail(err.value);
}

/**
 * Excel compares with 15 significant digits, so 0.1+0.2=0.3 is TRUE and a
 * sum that drifts by one ulp still equals its target.
 */
export function roundTo15(n: number): number {
  if (!Number.isFinite(n) || n === 0) return n;
  return Number.parseFloat(n.toPrecision(15));
}

/** Excel's ordering across types: numbers < text < booleans; text compares case-insensitively. */
export function compare(a: Scalar, b: Scalar): number {
  const rank = (v: Scalar) => (typeof v === 'number' || v === null ? 0 : typeof v === 'string' ? 1 : 2);
  if (isError(a)) return fail(a);
  if (isError(b)) return fail(b);
  // A blank compares as 0 against a number, as "" against text, as FALSE against a boolean.
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (typeof a === 'number' && typeof b === 'number') {
    const x = roundTo15(a);
    const y = roundTo15(b);
    return x === y ? 0 : x < y ? -1 : 1;
  }
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x === y ? 0 : x < y ? -1 : 1;
  }
  return a === b ? 0 : a ? 1 : -1;
}

/** Every scalar in a value, row by row. */
export function* flatten(v: Value): Generator<Scalar> {
  if (isMatrix(v)) {
    for (const row of v) yield* row;
  } else {
    yield v;
  }
}
