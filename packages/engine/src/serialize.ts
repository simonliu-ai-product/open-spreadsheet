import { columnName, formatA1, MAX_COLS, MAX_ROWS } from './address.ts';
import type { Node, RefPart } from './parse.ts';

const PRECEDENCE: Record<string, number> = {
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
  '&': 2,
  '+': 3, '-': 3,
  '*': 4, '/': 4,
  '^': 5,
};

const sheetPrefix = (sheet?: string) =>
  sheet === undefined ? '' : /^[A-Za-z_][\w.]*$/.test(sheet) ? `${sheet}!` : `'${sheet.replaceAll("'", "''")}'!`;

const part = (r: RefPart) => formatA1(r.row, r.col, r.rowAbs, r.colAbs);

/** Formula text for a parse tree, with the leading "=". */
export function serialize(node: Node): string {
  return `=${expr(node, 0)}`;
}

function expr(node: Node, parent: number): string {
  switch (node.kind) {
    case 'num':
      return String(node.value);
    case 'str':
      return `"${node.value.replaceAll('"', '""')}"`;
    case 'bool':
      return node.value ? 'TRUE' : 'FALSE';
    case 'err':
      return node.code;
    case 'ref':
      return node.row < 0 ? '#REF!' : sheetPrefix(node.sheet) + part(node);
    case 'range': {
      if (node.start.row < 0) return '#REF!';
      const pre = sheetPrefix(node.sheet);
      if (node.whole === 'cols') {
        return `${pre}${node.start.colAbs ? '$' : ''}${columnName(node.start.col)}:${node.end.colAbs ? '$' : ''}${columnName(node.end.col)}`;
      }
      if (node.whole === 'rows') {
        return `${pre}${node.start.rowAbs ? '$' : ''}${node.start.row + 1}:${node.end.rowAbs ? '$' : ''}${node.end.row + 1}`;
      }
      return `${pre}${part(node.start)}:${part(node.end)}`;
    }
    case 'name':
      return node.name;
    case 'call':
      return `${node.name}(${node.args.map((a) => expr(a, 0)).join(',')})`;
    case 'missing':
      return '';
    case 'unary':
      return `${node.op}${expr(node.arg, 6)}`;
    case 'percent':
      return `${expr(node.arg, 7)}%`;
    case 'paren':
      return `(${expr(node.arg, 0)})`;
    case 'array':
      return `{${node.rows.map((r) => r.map((c) => expr(c, 0)).join(',')).join(';')}}`;
    case 'binary': {
      const prec = PRECEDENCE[node.op]!;
      // Every level is left-associative, so a right operand at the same level needs parentheses.
      const text = `${expr(node.left, prec)}${node.op}${expr(node.right, prec + 1)}`;
      return prec < parent ? `(${text})` : text;
    }
  }
}

/**
 * Rewrites references after `count` rows (or columns) are inserted before
 * index `at` on `sheet` (negative count deletes). A reference to a deleted
 * cell becomes #REF!, as in Excel; a range shrinks or grows around the edit.
 */
export function shiftReferences(
  node: Node,
  edit: { sheet: string; axis: 'row' | 'col'; at: number; count: number },
  formulaSheet: string,
): Node {
  const onSheet = (sheet?: string) => (sheet ?? formulaSheet) === edit.sheet;
  const key = edit.axis;
  const limit = key === 'row' ? MAX_ROWS : MAX_COLS;

  const moveIndex = (i: number): number | null => {
    if (i < edit.at) return i;
    if (edit.count < 0 && i < edit.at - edit.count) return null; // inside the deleted band
    return i + edit.count;
  };

  const visit = (n: Node): Node => {
    switch (n.kind) {
      case 'ref': {
        if (!onSheet(n.sheet)) return n;
        const moved = moveIndex(n[key]);
        if (moved === null || moved >= limit) return { kind: 'err', code: '#REF!' };
        return { ...n, [key]: moved };
      }
      case 'range': {
        if (!onSheet(n.sheet) || (key === 'row' && n.whole === 'cols') || (key === 'col' && n.whole === 'rows')) return n;
        let lo = n.start[key];
        let hi = n.end[key];
        if (edit.count > 0) {
          if (lo >= edit.at) lo += edit.count;
          if (hi >= edit.at) hi += edit.count;
        } else {
          const removed = -edit.count;
          const d0 = edit.at;
          const d1 = edit.at + removed - 1;
          if (lo >= d0 && hi <= d1) return { kind: 'err', code: '#REF!' }; // the whole range is gone
          if (lo > d1) {
            lo -= removed;
            hi -= removed;
          } else if (hi >= d0) {
            // The deleted band overlaps the range: it shrinks around the gap.
            hi = hi > d1 ? hi - removed : d0 - 1;
            if (lo > d0) lo = d0;
          }
        }
        return { ...n, start: { ...n.start, [key]: lo }, end: { ...n.end, [key]: Math.min(hi, limit - 1) } };
      }
      case 'call':
        return { ...n, args: n.args.map(visit) };
      case 'unary':
      case 'percent':
      case 'paren':
        return { ...n, arg: visit(n.arg) };
      case 'binary':
        return { ...n, left: visit(n.left), right: visit(n.right) };
      case 'array':
        return { ...n, rows: n.rows.map((r) => r.map(visit)) };
      default:
        return n;
    }
  };
  return visit(node);
}

/** Applies `f` to every reference and range in a formula, bottom-up. */
function mapRefs(node: Node, f: (n: Node) => Node): Node {
  switch (node.kind) {
    case 'ref':
    case 'range':
      return f(node);
    case 'call':
      return { ...node, args: node.args.map((a) => mapRefs(a, f)) };
    case 'unary':
    case 'percent':
    case 'paren':
      return { ...node, arg: mapRefs(node.arg, f) };
    case 'binary':
      return { ...node, left: mapRefs(node.left, f), right: mapRefs(node.right, f) };
    case 'array':
      return { ...node, rows: node.rows.map((r) => r.map((c) => mapRefs(c, f))) };
    default:
      return node;
  }
}

/**
 * Moves relative references by (dr, dc), as Excel does when a formula is
 * copied, filled or pasted elsewhere. Absolute parts ($A$1) stay; a
 * reference pushed off the sheet becomes #REF!.
 */
export function offsetReferences(node: Node, dr: number, dc: number): Node {
  const REF: Node = { kind: 'err', code: '#REF!' };
  const move = (p: RefPart): RefPart | null => {
    const row = p.rowAbs ? p.row : p.row + dr;
    const col = p.colAbs ? p.col : p.col + dc;
    if (row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) return null;
    return { ...p, row, col };
  };
  return mapRefs(node, (n) => {
    if (n.kind === 'ref') {
      const m = move(n);
      return m ? { ...n, ...m } : REF;
    }
    if (n.kind !== 'range') return n;
    if (n.whole === 'cols') {
      const s = n.start.colAbs ? n.start.col : n.start.col + dc;
      const e = n.end.colAbs ? n.end.col : n.end.col + dc;
      if (s < 0 || e >= MAX_COLS) return REF;
      return { ...n, start: { ...n.start, col: s }, end: { ...n.end, col: e } };
    }
    if (n.whole === 'rows') {
      const s = n.start.rowAbs ? n.start.row : n.start.row + dr;
      const e = n.end.rowAbs ? n.end.row : n.end.row + dr;
      if (s < 0 || e >= MAX_ROWS) return REF;
      return { ...n, start: { ...n.start, row: s }, end: { ...n.end, row: e } };
    }
    const s = move(n.start);
    const e = move(n.end);
    return s && e ? { ...n, start: s, end: e } : REF;
  });
}

/** Points references at sheet `from` to sheet `to` (to = null: the sheet is gone, so #REF!). */
export function retargetSheet(node: Node, from: string, to: string | null): Node {
  const same = (s?: string) => s !== undefined && s.toLowerCase() === from.toLowerCase();
  return mapRefs(node, (n) => {
    if ((n.kind === 'ref' || n.kind === 'range') && same(n.sheet)) {
      return to === null ? { kind: 'err', code: '#REF!' } : { ...n, sheet: to };
    }
    return n;
  });
}
