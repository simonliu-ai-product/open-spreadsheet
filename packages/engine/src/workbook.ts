import { cellKey, formatA1, parseA1 } from './address.ts';
import { callFunction, type CallContext } from './functions.ts';
import './functions-more.ts';
import './functions-eng.ts';
import { FormulaSyntaxError, type Node, parseFormula } from './parse.ts';
import { offsetReferences, retargetSheet, serialize, shiftReferences } from './serialize.ts';
import {
  CellError,
  compare,
  err,
  ErrorSignal,
  isError,
  isMatrix,
  type Matrix,
  parseNumericText,
  type Scalar,
  scalarOf,
  toNumber,
  toText,
  type Value,
} from './values.ts';

interface Cell {
  /** What the person typed: "=SUM(A1:A3)", "42", "hello". */
  input: string;
  formula?: Node;
  value: Scalar;
}

/** sheet + row + col, the key of a cell across the workbook. */
type CellId = string;
const idOf = (sheet: string, row: number, col: number): CellId => `${sheet}!${row},${col}`;
const parseId = (id: CellId) => {
  const bang = id.lastIndexOf('!');
  const [row, col] = id.slice(bang + 1).split(',').map(Number) as [number, number];
  return { sheet: id.slice(0, bang), row, col };
};

interface RangeDep {
  sheet: string;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** Functions whose result can change with no input changing. */
const VOLATILE = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT']);

/** A literal the way Excel reads typed input: numbers, TRUE/FALSE, errors, else text. */
export function parseLiteral(input: string): Scalar {
  if (input === '') return null;
  // A leading apostrophe keeps the rest as text: '00123 stays "00123".
  if (input.startsWith("'")) return input.slice(1);
  const n = parseNumericText(input);
  if (n !== null) return n;
  const u = input.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return input;
}

/**
 * A formula as it reads after being copied (dr, dc) cells away: relative
 * references move, absolute ones stay. Non-formulas come back unchanged.
 */
export function offsetFormula(input: string, dr: number, dc: number): string {
  if (!input.startsWith('=') || (dr === 0 && dc === 0)) return input;
  try {
    return serialize(offsetReferences(parseFormula(input), dr, dc));
  } catch {
    return input;
  }
}

const contains = (g: RangeDep, sheet: string, row: number, col: number) => g.sheet === sheet && row >= g.r0 && row <= g.r1 && col >= g.c0 && col <= g.c1;
/** Ranges wider than this many columns are kept in one list instead of a bucket per column. */
const WIDE = 64;

/** Dirty cells, also findable by sheet and column, so a range finds its dirty cells without scanning them all. */
class DirtySet implements Iterable<CellId> {
  private readonly ids = new Set<CellId>();
  private readonly byCol = new Map<string, Set<number>>();
  get size(): number {
    return this.ids.size;
  }
  has(id: CellId): boolean {
    return this.ids.has(id);
  }
  add(id: CellId, at: { sheet: string; row: number; col: number } = parseId(id)): void {
    if (this.ids.has(id)) return;
    this.ids.add(id);
    const { sheet, row, col } = at;
    const key = `${sheet}\u0000${col}`;
    let rows = this.byCol.get(key);
    if (!rows) this.byCol.set(key, (rows = new Set()));
    rows.add(row);
  }
  delete(id: CellId, at?: { sheet: string; row: number; col: number }): void {
    if (!this.ids.delete(id)) return;
    const { sheet, row, col } = at ?? parseId(id);
    this.byCol.get(`${sheet}\u0000${col}`)?.delete(row);
  }
  clear(): void {
    this.ids.clear();
    this.byCol.clear();
  }
  rows(sheet: string, col: number): Set<number> | undefined {
    return this.byCol.get(`${sheet}\u0000${col}`);
  }
  [Symbol.iterator](): Iterator<CellId> {
    return this.ids[Symbol.iterator]();
  }
}

/**
 * Formulas that read through ranges, bucketed by the columns they cover, so a
 * changed cell looks only at ranges over its own column instead of every range
 * in the workbook.
 */
class RangeIndex {
  private readonly cols = new Map<string, Map<CellId, RangeDep[]>>();
  private readonly wide = new Map<CellId, RangeDep[]>();
  add(reader: CellId, ranges: RangeDep[]): void {
    for (const g of ranges) {
      if (g.c1 - g.c0 >= WIDE) {
        const list = this.wide.get(reader) ?? [];
        list.push(g);
        this.wide.set(reader, list);
        continue;
      }
      for (let c = g.c0; c <= g.c1; c++) {
        const key = `${g.sheet}\u0000${c}`;
        let bucket = this.cols.get(key);
        if (!bucket) this.cols.set(key, (bucket = new Map()));
        const list = bucket.get(reader) ?? [];
        list.push(g);
        bucket.set(reader, list);
      }
    }
  }
  remove(reader: CellId, ranges: RangeDep[]): void {
    this.wide.delete(reader);
    for (const g of ranges) {
      if (g.c1 - g.c0 >= WIDE) continue;
      for (let c = g.c0; c <= g.c1; c++) this.cols.get(`${g.sheet}\u0000${c}`)?.delete(reader);
    }
  }
  clear(): void {
    this.cols.clear();
    this.wide.clear();
  }
  /** Formulas with a range that holds this cell. */
  *readersOf(sheet: string, row: number, col: number): Generator<CellId> {
    for (const [reader, gs] of this.cols.get(`${sheet}\u0000${col}`) ?? []) if (gs.some((g) => contains(g, sheet, row, col))) yield reader;
    for (const [reader, gs] of this.wide) if (gs.some((g) => contains(g, sheet, row, col))) yield reader;
  }
}

/** Parses name definitions; a definition that is not a reference or range with its sheet is refused. */
function parseNames(defs: Record<string, string>): Map<string, { name: string; node: Node }> {
  const out = new Map<string, { name: string; node: Node }>();
  for (const [name, ref] of Object.entries(defs)) {
    if (!Workbook.validName(name)) throw new Error(`"${name}" is not a valid name`);
    let node: Node;
    try {
      node = parseFormula(ref.startsWith('=') ? ref : `=${ref}`);
    } catch {
      throw new Error(`The name ${name} must refer to a cell or range, e.g. Sheet1!$A$1:$B$10`);
    }
    if ((node.kind !== 'ref' && node.kind !== 'range') || node.sheet === undefined) {
      throw new Error(`The name ${name} must refer to a cell or range with its sheet, e.g. Sheet1!$A$1:$B$10`);
    }
    out.set(name.toUpperCase(), { name, node });
  }
  return out;
}

function mapNames(names: Map<string, { name: string; node: Node }>, f: (n: Node) => Node): Map<string, { name: string; node: Node }> {
  return new Map([...names].map(([k, v]) => [k, { name: v.name, node: f(v.node) }]));
}

export class Workbook {
  private readonly sheets = new Map<string, Map<string, Cell>>();
  /** cell -> the cells and ranges its formula reads */
  private readonly reads = new Map<CellId, { cells: CellId[]; ranges: RangeDep[]; volatile: boolean }>();
  /** cell -> formulas that read it directly */
  private readonly readers = new Map<CellId, Set<CellId>>();
  /** formulas that read through a range, indexed by column */
  private readonly rangeReaders = new RangeIndex();
  private readonly dirty = new DirtySet();
  /** While loading a whole workbook every formula is dirty anyway, so a cell need not mark its readers. */
  private loading = false;
  /** Defined names (upper-cased) -> the reference or range they stand for, always with its sheet. */
  private names = new Map<string, { name: string; node: Node }>();

  // ---------- defined names ----------

  /** A name Excel would accept: a letter, underscore or backslash first; not TRUE/FALSE; not a cell address like A1 or R1C1. */
  static validName(name: string): boolean {
    return (
      /^[A-Za-z_\\\u00C0-\uFFFF][\w.\u00C0-\uFFFF]{0,254}$/.test(name) &&
      !/^(TRUE|FALSE)$/i.test(name) &&
      !/^\$?[A-Za-z]{1,3}\$?\d+$/.test(name) &&
      !/^[Rr]\d*[Cc]\d*$/.test(name)
    );
  }

  /** The names as written: { Sales: "Orders!$F$2:$F$9" }. */
  namesJSON(): Record<string, string> {
    return Object.fromEntries([...this.names.values()].map(({ name, node }) => [name, serialize(node).slice(1)]));
  }

  /**
   * Replaces the defined names, then relinks every formula, so one that used
   * a name before it existed (and showed #NAME?) now reads it.
   */
  setNames(defs: Record<string, string>): void {
    this.names = parseNames(defs);
    this.rebuild((sheet, row, col, cell) => ({ sheet, row, col, input: cell.input }));
  }
  private readonly computing = new Set<CellId>();
  now: () => Date = () => new Date();

  addSheet(name: string): void {
    if (!this.sheets.has(name)) this.sheets.set(name, new Map());
  }

  sheetNames(): string[] {
    return [...this.sheets.keys()];
  }

  private sheet(name: string): Map<string, Cell> {
    const s = this.sheets.get(name);
    if (!s) throw new Error(`No sheet named ${name}`);
    return s;
  }

  /** Sets one cell from what a person typed, then recalculates what depends on it. */
  set(sheet: string, a1: string, input: string): void {
    const at = parseA1(a1);
    if (!at) throw new Error(`Bad address ${a1}`);
    this.setAt(sheet, at.row, at.col, input);
    this.recalculate();
  }

  /** Sets many cells, recalculating once at the end. */
  setMany(sheet: string, cells: Record<string, string | number | boolean>): void {
    for (const [a1, v] of Object.entries(cells)) {
      const at = parseA1(a1);
      if (!at) throw new Error(`Bad address ${a1}`);
      this.setAt(sheet, at.row, at.col, String(v));
    }
    this.recalculate();
  }

  private setAt(sheetName: string, row: number, col: number, input: string): void {
    const sheet = this.sheet(sheetName);
    const id = idOf(sheetName, row, col);
    this.unlink(id);
    if (input === '') {
      sheet.delete(cellKey(row, col));
    } else if (input.startsWith('=') && input.length > 1) {
      let formula: Node;
      try {
        formula = parseFormula(input);
      } catch (e) {
        if (!(e instanceof FormulaSyntaxError)) throw e;
        sheet.set(cellKey(row, col), { input, value: err.name });
        this.markDependentsDirty(id);
        return;
      }
      sheet.set(cellKey(row, col), { input, formula, value: null });
      this.link(id, sheetName, formula);
      this.dirty.add(id);
    } else {
      sheet.set(cellKey(row, col), { input, value: parseLiteral(input) });
    }
    this.markDependentsDirty(id);
  }

  // ---------- row/column access for an editor ----------

  /** Sets one cell by position and recalculates. */
  setCell(sheet: string, row: number, col: number, input: string): void {
    this.setAt(sheet, row, col, input);
    this.recalculate();
  }

  /** Sets many cells by position, recalculating once. */
  setCells(sheet: string, cells: { row: number; col: number; input: string }[]): void {
    for (const c of cells) this.setAt(sheet, c.row, c.col, c.input);
    this.recalculate();
  }

  value(sheet: string, row: number, col: number): Scalar {
    return this.valueAt(sheet, row, col);
  }

  inputAt(sheet: string, row: number, col: number): string {
    return this.sheets.get(sheet)?.get(cellKey(row, col))?.input ?? '';
  }

  /** Rows and columns in use on a sheet. */
  used(sheet: string): { rows: number; cols: number } {
    return this.extent(sheet);
  }

  /** The workbook as typed: every non-empty input by A1 address, sheet by sheet. */
  toJSON(): { sheets: { name: string; cells: Record<string, string> }[]; names?: Record<string, string> } {
    return {
      ...(this.names.size ? { names: this.namesJSON() } : {}),
      sheets: [...this.sheets].map(([name, cells]) => {
        const sorted = [...cells.entries()]
          .map(([key, cell]) => {
            const [row, col] = key.split(',').map(Number) as [number, number];
            return { row, col, input: cell.input };
          })
          .sort((a, b) => a.row - b.row || a.col - b.col);
        return { name, cells: Object.fromEntries(sorted.map((c) => [formatA1(c.row, c.col), c.input])) };
      }),
    };
  }

  static fromJSON(data: { sheets: { name: string; cells: Record<string, string> }[]; names?: Record<string, string> }): Workbook {
    const wb = new Workbook();
    for (const s of data.sheets) wb.addSheet(s.name);
    if (data.names) wb.names = parseNames(data.names);
    wb.loading = true;
    for (const s of data.sheets) {
      for (const [a1, input] of Object.entries(s.cells)) {
        const at = parseA1(a1);
        if (at) wb.setAt(s.name, at.row, at.col, input);
      }
    }
    wb.loading = false;
    wb.recalculate();
    return wb;
  }

  get(sheet: string, a1: string): Scalar {
    const at = parseA1(a1);
    if (!at) throw new Error(`Bad address ${a1}`);
    return this.valueAt(sheet, at.row, at.col);
  }

  /** The text a person would see in the formula bar. */
  input(sheet: string, a1: string): string {
    const at = parseA1(a1)!;
    return this.sheet(sheet).get(cellKey(at.row, at.col))?.input ?? '';
  }

  // ---------- dependency graph ----------

  private link(id: CellId, sheet: string, formula: Node): void {
    const cells: CellId[] = [];
    const ranges: RangeDep[] = [];
    let volatile = false;
    const walk = (n: Node) => {
      switch (n.kind) {
        case 'ref':
          cells.push(idOf(n.sheet ?? sheet, n.row, n.col));
          break;
        case 'range':
          ranges.push({ sheet: n.sheet ?? sheet, r0: n.start.row, c0: n.start.col, r1: n.end.row, c1: n.end.col });
          break;
        case 'name': {
          const def = this.names.get(n.name.toUpperCase());
          if (def) walk(def.node);
          break;
        }
        case 'call':
          if (VOLATILE.has(n.name)) volatile = true;
          n.args.forEach(walk);
          break;
        case 'unary':
        case 'percent':
        case 'paren':
          walk(n.arg);
          break;
        case 'binary':
          walk(n.left);
          walk(n.right);
          break;
        case 'array':
          n.rows.forEach((r) => r.forEach(walk));
          break;
      }
    };
    walk(formula);
    this.reads.set(id, { cells, ranges, volatile });
    if (ranges.length) this.rangeReaders.add(id, ranges);
    for (const c of cells) {
      let set = this.readers.get(c);
      if (!set) this.readers.set(c, (set = new Set()));
      set.add(id);
    }
  }

  private unlink(id: CellId): void {
    const r = this.reads.get(id);
    if (!r) return;
    for (const c of r.cells) this.readers.get(c)?.delete(id);
    this.reads.delete(id);
    if (r.ranges.length) this.rangeReaders.remove(id, r.ranges);
  }

  /**
   * What the formula in a cell reads directly — each cell or range with its
   * sheet, defined names already resolved — for tracing where a value comes from.
   */
  precedents(sheet: string, row: number, col: number): { sheet: string; r0: number; c0: number; r1: number; c1: number }[] {
    const r = this.reads.get(idOf(sheet, row, col));
    if (!r) return [];
    const out = [
      ...r.cells.map((c) => {
        const at = parseId(c);
        return { sheet: at.sheet, r0: at.row, c0: at.col, r1: at.row, c1: at.col };
      }),
      ...r.ranges.map((g) => ({ sheet: g.sheet, r0: g.r0, c0: g.c0, r1: g.r1, c1: g.c1 })),
    ];
    const seen = new Set<string>();
    return out.filter((g) => {
      const key = `${g.sheet}!${g.r0},${g.c0}:${g.r1},${g.c1}`;
      return !seen.has(key) && (seen.add(key), true);
    });
  }

  /** The formulas that read a cell directly, by reference or through a range, in sheet then reading order. */
  dependents(sheet: string, row: number, col: number): { sheet: string; row: number; col: number }[] {
    const ids = new Set<CellId>(this.readers.get(idOf(sheet, row, col)) ?? []);
    for (const reader of this.rangeReaders.readersOf(sheet, row, col)) ids.add(reader);
    const order = [...this.sheets.keys()];
    return [...ids]
      .map(parseId)
      .sort((a, b) => order.indexOf(a.sheet) - order.indexOf(b.sheet) || a.row - b.row || a.col - b.col);
  }

  /** Every formula that reads this cell, directly or through a range, transitively. */
  private markDependentsDirty(start: CellId): void {
    if (this.loading) return;
    // Each queued cell carries its parsed address, so an id is parsed once.
    const queue: [CellId, { sheet: string; row: number; col: number }][] = [[start, parseId(start)]];
    while (queue.length) {
      const [id, { sheet, row, col }] = queue.pop()!;
      const visit = (n: CellId) => {
        if (this.dirty.has(n)) return;
        const at = parseId(n);
        this.dirty.add(n, at);
        queue.push([n, at]);
      };
      const direct = this.readers.get(id);
      if (direct) for (const n of direct) visit(n);
      for (const reader of this.rangeReaders.readersOf(sheet, row, col)) visit(reader);
    }
  }

  /** The dirty cells a formula reads, directly or inside one of its ranges. */
  private dirtyPrecedents(id: CellId): CellId[] {
    const r = this.reads.get(id);
    if (!r) return [];
    const out = r.cells.filter((c) => this.dirty.has(c));
    for (const g of r.ranges) {
      if (g.c1 - g.c0 >= WIDE) {
        for (const d of this.dirty) {
          const p = parseId(d);
          if (contains(g, p.sheet, p.row, p.col)) out.push(d);
        }
        continue;
      }
      for (let c = g.c0; c <= g.c1; c++) {
        const rows = this.dirty.rows(g.sheet, c);
        if (!rows?.size) continue;
        // Walk whichever is shorter: the dirty rows of this column, or the range's rows.
        if (rows.size <= g.r1 - g.r0 + 1) {
          for (const row of rows) if (row >= g.r0 && row <= g.r1) out.push(idOf(g.sheet, row, c));
        } else {
          for (let row = g.r0; row <= g.r1; row++) if (rows.has(row)) out.push(idOf(g.sheet, row, c));
        }
      }
    }
    return out;
  }

  /**
   * Recomputes every dirty formula, each after the cells it reads. The order
   * comes from an explicit-stack depth-first walk, not recursion, so a chain
   * a hundred thousand rows long does not overflow the stack.
   */
  recalculate(): void {
    for (const [id, r] of this.reads) if (r.volatile) this.dirty.add(id);
    const state = new Map<CellId, 1 | 2>(); // 1: on the stack, 2: ordered
    const order: CellId[] = [];
    for (const start of this.dirty) {
      if (state.has(start)) continue;
      const stack: [CellId, CellId[] | null][] = [[start, null]];
      while (stack.length) {
        const top = stack[stack.length - 1]!;
        if (top[1] === null) {
          state.set(top[0], 1);
          top[1] = this.dirtyPrecedents(top[0]);
        }
        const next = top[1].pop();
        if (next !== undefined) {
          // A precedent already on the stack is a cycle; valueAt reports it as #CIRC!.
          if (!state.has(next)) stack.push([next, null]);
          continue;
        }
        state.set(top[0], 2);
        order.push(top[0]);
        stack.pop();
      }
    }
    for (const id of order) {
      const { sheet, row, col } = parseId(id);
      this.valueAt(sheet, row, col);
    }
  }

  // ---------- evaluation ----------

  private valueAt(sheetName: string, row: number, col: number): Scalar {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) return err.ref;
    const cell = sheet.get(cellKey(row, col));
    if (!cell) return null;
    if (!cell.formula) return cell.value;
    const id = idOf(sheetName, row, col);
    if (!this.dirty.has(id)) return cell.value;
    if (this.computing.has(id)) return err.circ;
    this.computing.add(id);
    try {
      const v = this.evaluate(cell.formula, sheetName, row, col);
      // A formula returning a range shows its top-left value until spilling lands (PLAN.md §4).
      cell.value = scalarOf(v);
      if (cell.value === null) cell.value = 0; // a formula pointing at a blank shows 0, as in Excel
    } finally {
      this.computing.delete(id);
      this.dirty.delete(id, { sheet: sheetName, row, col });
    }
    return cell.value;
  }

  /** The used extent of a sheet, so A:A reads the filled rows, not a million. */
  private extent(sheetName: string): { rows: number; cols: number } {
    let rows = 0;
    let cols = 0;
    for (const key of this.sheets.get(sheetName)?.keys() ?? []) {
      const [r, c] = key.split(',').map(Number) as [number, number];
      if (r + 1 > rows) rows = r + 1;
      if (c + 1 > cols) cols = c + 1;
    }
    return { rows, cols };
  }

  private evaluate(node: Node, sheet: string, row: number, col: number): Value {
    const ctx: CallContext = {
      evaluate: (n) => this.evaluate(n, sheet, row, col),
      position: (n) => {
        if (n.kind === 'ref') return { row: n.row, col: n.col, rows: 1, cols: 1 };
        if (n.kind === 'range') return { row: n.start.row, col: n.start.col, rows: n.end.row - n.start.row + 1, cols: n.end.col - n.start.col + 1 };
        return null;
      },
      here: { row, col },
      now: this.now,
    };
    try {
      return this.evalNode(node, sheet, ctx);
    } catch (e) {
      if (e instanceof ErrorSignal) return e.error;
      throw e;
    }
  }

  private evalNode(node: Node, sheet: string, ctx: CallContext): Value {
    switch (node.kind) {
      case 'num':
        return node.value;
      case 'str':
        return node.value;
      case 'bool':
        return node.value;
      case 'err':
        return new CellError(node.code);
      case 'missing':
        return null;
      case 'name': {
        const def = this.names.get(node.name.toUpperCase());
        return def ? this.evalNode(def.node, sheet, ctx) : err.name;
      }
      case 'ref':
        return this.valueAt(node.sheet ?? sheet, node.row, node.col);
      case 'range': {
        const s = node.sheet ?? sheet;
        if (!this.sheets.has(s)) return err.ref;
        const ext = node.whole ? this.extent(s) : null;
        const r1 = ext && node.whole === 'cols' ? Math.min(node.end.row, Math.max(ext.rows - 1, node.start.row)) : node.end.row;
        const c1 = ext && node.whole === 'rows' ? Math.min(node.end.col, Math.max(ext.cols - 1, node.start.col)) : node.end.col;
        const m: Matrix = [];
        for (let r = node.start.row; r <= r1; r++) {
          const line: Scalar[] = [];
          for (let c = node.start.col; c <= c1; c++) line.push(this.valueAt(s, r, c));
          m.push(line);
        }
        return m;
      }
      case 'array':
        return node.rows.map((r) => r.map((c) => scalarOf(this.evalNode(c, sheet, ctx))));
      case 'call':
        return callFunction(node.name, node.args, ctx);
      case 'unary':
        return mapValue(this.evalNode(node.arg, sheet, ctx), (v) => (node.op === '-' ? -toNumber(v) : toNumber(v)));
      case 'percent':
        return mapValue(this.evalNode(node.arg, sheet, ctx), (v) => toNumber(v) / 100);
      case 'paren':
        return this.evalNode(node.arg, sheet, ctx);
      case 'binary': {
        const left = this.evalNode(node.left, sheet, ctx);
        const right = this.evalNode(node.right, sheet, ctx);
        return mapPair(left, right, (a, b) => binary(node.op, a, b));
      }
    }
  }

  // ---------- structure edits ----------

  /** Inserts (count > 0) or deletes (count < 0) rows before `at`, rewriting every formula that points across. */
  insertRows(sheetName: string, at: number, count: number): void {
    this.restructure({ sheet: sheetName, axis: 'row', at, count });
  }

  insertColumns(sheetName: string, at: number, count: number): void {
    this.restructure({ sheet: sheetName, axis: 'col', at, count });
  }

  private restructure(edit: { sheet: string; axis: 'row' | 'col'; at: number; count: number }): void {
    this.names = mapNames(this.names, (node) => shiftReferences(node, edit, ''));
    this.rebuild((name, row, col, cell) => {
      if (name === edit.sheet) {
        const i = edit.axis === 'row' ? row : col;
        if (i >= edit.at) {
          if (edit.count < 0 && i < edit.at - edit.count) return null; // deleted
          if (edit.axis === 'row') row += edit.count;
          else col += edit.count;
        }
      }
      const input = cell.formula ? serialize(shiftReferences(cell.formula, edit, name)) : cell.input;
      return { sheet: name, row, col, input };
    });
  }

  /** Renames a sheet; every formula that names it follows. */
  renameSheet(from: string, to: string): void {
    if (from === to) return;
    if (!this.sheets.has(from)) throw new Error(`No sheet named ${from}`);
    if ([...this.sheets.keys()].some((k) => k !== from && k.toLowerCase() === to.toLowerCase())) throw new Error(`A sheet named ${to} exists`);
    const order = [...this.sheets.keys()].map((k) => (k === from ? to : k));
    this.names = mapNames(this.names, (node) => retargetSheet(node, from, to));
    this.rebuild(
      (name, row, col, cell) => ({
        sheet: name === from ? to : name,
        row,
        col,
        input: cell.formula ? serialize(retargetSheet(cell.formula, from, to)) : cell.input,
      }),
      order,
    );
  }

  /** Removes a sheet; formulas that pointed into it show #REF!, as in Excel. */
  removeSheet(name: string): void {
    if (!this.sheets.has(name)) return;
    const order = [...this.sheets.keys()].filter((k) => k !== name);
    this.names = mapNames(this.names, (node) => retargetSheet(node, name, null));
    this.rebuild(
      (sheet, row, col, cell) =>
        sheet === name ? null : { sheet, row, col, input: cell.formula ? serialize(retargetSheet(cell.formula, name, null)) : cell.input },
      order,
    );
  }

  /** Puts the sheets in this order (names not listed keep their place at the end). */
  reorderSheets(names: string[]): void {
    const rest = [...this.sheets.keys()].filter((k) => !names.includes(k));
    const order = [...names.filter((n) => this.sheets.has(n)), ...rest];
    const entries = new Map(order.map((n) => [n, this.sheets.get(n)!]));
    this.sheets.clear();
    for (const [k, v] of entries) this.sheets.set(k, v);
  }

  /** Rewrites every cell through `map` (null drops it), then rebuilds the graph and recalculates. */
  private rebuild(
    map: (sheet: string, row: number, col: number, cell: Cell) => { sheet: string; row: number; col: number; input: string } | null,
    order?: string[],
  ): void {
    const entries: { sheet: string; row: number; col: number; input: string }[] = [];
    for (const [name, cells] of this.sheets) {
      for (const [key, cell] of cells) {
        const [row, col] = key.split(',').map(Number) as [number, number];
        const e = map(name, row, col, cell);
        if (e) entries.push(e);
      }
    }
    if (order) {
      this.sheets.clear();
      for (const n of order) this.sheets.set(n, new Map());
    }
    for (const cells of this.sheets.values()) cells.clear();
    this.reads.clear();
    this.readers.clear();
    this.rangeReaders.clear();
    this.dirty.clear();
    this.loading = true;
    for (const e of entries) this.setAt(e.sheet, e.row, e.col, e.input);
    this.loading = false;
    this.recalculate();
  }
}

function mapValue(v: Value, f: (s: Scalar) => Scalar): Value {
  const one = (s: Scalar): Scalar => {
    if (isError(s)) return s;
    try {
      return f(s);
    } catch (e) {
      if (e instanceof ErrorSignal) return e.error;
      throw e;
    }
  };
  return isMatrix(v) ? v.map((r) => r.map(one)) : one(v);
}

/** Element-wise over arrays, broadcasting a single value or a single row/column. */
function mapPair(a: Value, b: Value, f: (x: Scalar, y: Scalar) => Scalar): Value {
  const safe = (x: Scalar, y: Scalar): Scalar => {
    try {
      return f(x, y);
    } catch (e) {
      if (e instanceof ErrorSignal) return e.error;
      throw e;
    }
  };
  if (!isMatrix(a) && !isMatrix(b)) return safe(a, b);
  const A = isMatrix(a) ? a : [[a]];
  const B = isMatrix(b) ? b : [[b]];
  const rows = Math.max(A.length, B.length);
  const cols = Math.max(A[0]?.length ?? 0, B[0]?.length ?? 0);
  const pick = (M: Matrix, r: number, c: number): Scalar => {
    const row = M.length === 1 ? M[0]! : M[r];
    if (!row) return err.na;
    const v = row.length === 1 ? row[0] : row[c];
    return v === undefined ? err.na : v;
  };
  const out: Matrix = [];
  for (let r = 0; r < rows; r++) {
    const line: Scalar[] = [];
    for (let c = 0; c < cols; c++) line.push(safe(pick(A, r, c), pick(B, r, c)));
    out.push(line);
  }
  return out;
}

function binary(op: string, a: Scalar, b: Scalar): Scalar {
  if (isError(a)) return a;
  if (isError(b)) return b;
  switch (op) {
    case '&':
      return toText(a) + toText(b);
    case '=':
      return compare(a, b) === 0;
    case '<>':
      return compare(a, b) !== 0;
    case '<':
      return compare(a, b) < 0;
    case '>':
      return compare(a, b) > 0;
    case '<=':
      return compare(a, b) <= 0;
    case '>=':
      return compare(a, b) >= 0;
  }
  const x = toNumber(a);
  const y = toNumber(b);
  let r: number;
  switch (op) {
    case '+':
      r = x + y;
      break;
    case '-':
      r = x - y;
      break;
    case '*':
      r = x * y;
      break;
    case '/':
      if (y === 0) return err.div0;
      r = x / y;
      break;
    case '^':
      if (x === 0 && y === 0) return err.num;
      r = x ** y;
      break;
    default:
      return err.value;
  }
  return Number.isFinite(r) ? r : err.num;
}
