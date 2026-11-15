import { columnIndex, MAX_COLS, MAX_ROWS } from './address.ts';
import { ERROR_CODES, type ErrorCode } from './values.ts';

export interface RefPart {
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

export type Node =
  | { kind: 'num'; value: number }
  | { kind: 'str'; value: string }
  | { kind: 'bool'; value: boolean }
  | { kind: 'err'; code: ErrorCode }
  | ({ kind: 'ref'; sheet?: string } & RefPart)
  | { kind: 'range'; sheet?: string; start: RefPart; end: RefPart; whole?: 'cols' | 'rows' }
  | { kind: 'name'; name: string }
  | { kind: 'call'; name: string; args: Node[] }
  | { kind: 'unary'; op: '-' | '+'; arg: Node }
  | { kind: 'percent'; arg: Node }
  /** Parentheses the person typed, kept so a rewritten formula reads as they wrote it. */
  | { kind: 'paren'; arg: Node }
  | { kind: 'binary'; op: BinaryOp; left: Node; right: Node }
  | { kind: 'array'; rows: Node[][] }
  | { kind: 'missing' };

export type BinaryOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=';

export class FormulaSyntaxError extends Error {}

type Token =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'err'; v: ErrorCode }
  | { t: 'ref'; node: Node }
  | { t: 'func'; v: string }
  | { t: 'ident'; v: string }
  | { t: 'op'; v: string }
  | { t: 'end' };

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const IDENT_START = /[A-Za-z_\\À-￿]/;
const IDENT_CHAR = /[A-Za-z0-9_.\\À-￿]/;

function refPart(text: string): RefPart | null {
  const m = CELL.exec(text);
  if (!m) return null;
  const col = columnIndex(m[1]!);
  const row = Number(m[2]) - 1;
  if (row < 0 || row >= MAX_ROWS || col >= MAX_COLS) return null;
  return { row, col, colAbs: text.startsWith('$'), rowAbs: text.slice(1).includes('$') };
}

function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  const peek = (k = 0) => src[i + k] ?? '';

  const readSheetPrefix = (): string | undefined => {
    if (peek() === "'") {
      let j = i + 1;
      let name = '';
      while (j < src.length) {
        if (src[j] === "'" && src[j + 1] === "'") {
          name += "'";
          j += 2;
        } else if (src[j] === "'") break;
        else name += src[j++];
      }
      if (src[j] === "'" && src[j + 1] === '!') {
        i = j + 2;
        return name;
      }
      return undefined;
    }
    const m = /^([A-Za-z_À-￿][\w.À-￿]*)!/.exec(src.slice(i));
    if (m) {
      i += m[0].length;
      return m[1];
    }
    return undefined;
  };

  /** A1, A1:B2, A:C or 1:3 at the cursor, with an optional sheet already read. */
  const readReference = (sheet?: string): Node | null => {
    const rest = src.slice(i);
    let m = /^(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?(?![\w(.])/.exec(rest);
    if (m) {
      const a = refPart(m[1]!);
      const b = m[2] ? refPart(m[2]) : null;
      if (a && (!m[2] || b)) {
        i += m[0].length;
        return b ? { kind: 'range', sheet, start: a, end: b } : { kind: 'ref', sheet, ...a };
      }
    }
    m = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/.exec(rest);
    if (m) {
      i += m[0].length;
      return {
        kind: 'range',
        sheet,
        whole: 'cols',
        start: { row: 0, col: columnIndex(m[2]!), rowAbs: true, colAbs: m[1] === '$' },
        end: { row: MAX_ROWS - 1, col: columnIndex(m[4]!), rowAbs: true, colAbs: m[3] === '$' },
      };
    }
    m = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\d.])/.exec(rest);
    if (m) {
      i += m[0].length;
      return {
        kind: 'range',
        sheet,
        whole: 'rows',
        start: { row: Number(m[2]) - 1, col: 0, rowAbs: m[1] === '$', colAbs: true },
        end: { row: Number(m[4]) - 1, col: MAX_COLS - 1, rowAbs: m[3] === '$', colAbs: true },
      };
    }
    return null;
  };

  while (i < src.length) {
    const c = peek();
    if (c === ' ' || c === '\n' || c === '\t' || c === '\r') {
      i++;
      continue;
    }
    if (c === '"') {
      let s = '';
      i++;
      for (;;) {
        if (i >= src.length) throw new FormulaSyntaxError('Unterminated text');
        if (src[i] === '"' && src[i + 1] === '"') {
          s += '"';
          i += 2;
        } else if (src[i] === '"') {
          i++;
          break;
        } else s += src[i++];
      }
      out.push({ t: 'str', v: s });
      continue;
    }
    if (c === '#') {
      const code = ERROR_CODES.find((e) => src.startsWith(e, i)) ?? (src.startsWith('#N/A', i) ? '#N/A' : null);
      if (!code) throw new FormulaSyntaxError(`Unknown error literal at ${i}`);
      i += code.length;
      out.push({ t: 'err', v: code });
      continue;
    }
    // A row range (1:3) looks like a number; try references first.
    if (/[\d$]/.test(c) || IDENT_START.test(c) || c === "'") {
      const start = i;
      const sheet = readSheetPrefix();
      const ref = readReference(sheet);
      if (ref) {
        out.push({ t: 'ref', node: ref });
        continue;
      }
      if (sheet !== undefined) throw new FormulaSyntaxError(`Bad reference after ${sheet}!`);
      i = start;
    }
    if (/[\d.]/.test(c)) {
      const m = /^(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i));
      if (!m) throw new FormulaSyntaxError(`Bad number at ${i}`);
      i += m[0].length;
      out.push({ t: 'num', v: Number(m[0]) });
      continue;
    }
    if (IDENT_START.test(c)) {
      let j = i;
      while (j < src.length && IDENT_CHAR.test(src[j]!)) j++;
      const word = src.slice(i, j);
      i = j;
      out.push(peek() === '(' ? { t: 'func', v: word.toUpperCase() } : { t: 'ident', v: word });
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      out.push({ t: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%(),;{}'.includes(c)) {
      out.push({ t: 'op', v: c });
      i++;
      continue;
    }
    throw new FormulaSyntaxError(`Unexpected "${c}" at ${i}`);
  }
  out.push({ t: 'end' });
  return out;
}

/** Parses a formula, with or without its leading "=". */
export function parseFormula(formula: string): Node {
  const tokens = tokenize(formula.startsWith('=') ? formula.slice(1) : formula);
  let p = 0;
  const peek = () => tokens[p]!;
  const isOp = (v: string) => {
    const t = peek();
    return t.t === 'op' && t.v === v;
  };
  const expect = (v: string) => {
    if (!isOp(v)) throw new FormulaSyntaxError(`Expected "${v}"`);
    p++;
  };

  const binaryLevel = (ops: string[], next: () => Node) => (): Node => {
    let left = next();
    for (;;) {
      const t = peek();
      if (t.t !== 'op' || !ops.includes(t.v)) return left;
      p++;
      left = { kind: 'binary', op: t.v as BinaryOp, left, right: next() };
    }
  };

  // Excel's precedence, lowest first: comparison, &, + -, * /, ^, negation, %.
  const comparison = (): Node => binaryLevel(['=', '<>', '<', '>', '<=', '>='], concat)();
  const concat = (): Node => binaryLevel(['&'], additive)();
  const additive = (): Node => binaryLevel(['+', '-'], multiplicative)();
  const multiplicative = (): Node => binaryLevel(['*', '/'], power)();
  // ^ is left-associative in Excel: 2^3^2 is 64.
  const power = (): Node => binaryLevel(['^'], unary)();
  // Negation binds tighter than ^: -2^2 is 4.
  const unary = (): Node => {
    if (isOp('-') || isOp('+')) {
      const op = (peek() as { v: '-' | '+' }).v;
      p++;
      return { kind: 'unary', op, arg: unary() };
    }
    return postfix();
  };
  const postfix = (): Node => {
    let node = primary();
    while (isOp('%')) {
      p++;
      node = { kind: 'percent', arg: node };
    }
    return node;
  };

  const args = (): Node[] => {
    const list: Node[] = [];
    if (isOp(')')) {
      p++;
      return list;
    }
    for (;;) {
      list.push(isOp(',') || isOp(')') ? { kind: 'missing' } : comparison());
      if (isOp(',')) {
        p++;
        continue;
      }
      expect(')');
      return list;
    }
  };

  const arrayLiteral = (): Node => {
    const rows: Node[][] = [[]];
    for (;;) {
      rows[rows.length - 1]!.push(unary());
      if (isOp(',')) {
        p++;
      } else if (isOp(';')) {
        p++;
        rows.push([]);
      } else {
        expect('}');
        return { kind: 'array', rows };
      }
    }
  };

  function primary(): Node {
    const t = peek();
    p++;
    switch (t.t) {
      case 'num':
        return { kind: 'num', value: t.v };
      case 'str':
        return { kind: 'str', value: t.v };
      case 'err':
        return { kind: 'err', code: t.v };
      case 'ref':
        return t.node;
      case 'func':
        p++; // the "(" the tokenizer already saw
        return { kind: 'call', name: t.v.replace(/^_XLFN\./, '').replace(/^_XLWS\./, ''), args: args() };
      case 'ident': {
        const u = t.v.toUpperCase();
        if (u === 'TRUE' || u === 'FALSE') return { kind: 'bool', value: u === 'TRUE' };
        return { kind: 'name', name: t.v };
      }
      case 'op':
        if (t.v === '(') {
          const inner = comparison();
          expect(')');
          return { kind: 'paren', arg: inner };
        }
        if (t.v === '{') return arrayLiteral();
        throw new FormulaSyntaxError(`Unexpected "${t.v}"`);
      default:
        throw new FormulaSyntaxError('Unexpected end of formula');
    }
  }

  const node = comparison();
  if (peek().t !== 'end') throw new FormulaSyntaxError('Unexpected text after the formula');
  return node;
}
