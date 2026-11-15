/**
 * Conditional formatting, data validation and filters: what they compute for
 * a cell, and how their ranges follow inserted and deleted rows and columns.
 * Pure functions over a sheet and a way to read its values; the editor paints.
 */
import { CellError, compare, formatA1, parseA1, type Scalar } from '@open-spreadsheet/engine';
import type { ConditionalRule, HighlightStyle, SheetData } from './types.ts';

export interface Rect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export function parseRange(text: string): Rect | null {
  const [a, b] = text.trim().toUpperCase().replaceAll('$', '').split(':');
  const p = parseA1(a ?? '');
  const q = b ? parseA1(b) : p;
  if (!p || !q) return null;
  return { r0: Math.min(p.row, q.row), c0: Math.min(p.col, q.col), r1: Math.max(p.row, q.row), c1: Math.max(p.col, q.col) };
}
export const rangeText = (r: Rect) => (r.r0 === r.r1 && r.c0 === r.c1 ? formatA1(r.r0, r.c0) : `${formatA1(r.r0, r.c0)}:${formatA1(r.r1, r.c1)}`);
export const inRect = (r: Rect, row: number, col: number) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

/**
 * A range after `count` rows or columns are inserted before `at` (negative:
 * deleted): it grows, shrinks or moves; null when every cell of it is gone.
 */
/** Whether a cell lies in one of the sheet's protected ranges. */
export function isProtectedCell(sheet: SheetData, row: number, col: number): boolean {
  return (sheet.protected ?? []).some((t) => {
    const r = parseRange(t);
    return r !== null && row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;
  });
}

/** Whether any cell of a rectangle is protected. */
export function touchesProtected(sheet: SheetData, rect: { r0: number; c0: number; r1: number; c1: number }): boolean {
  return (sheet.protected ?? []).some((t) => {
    const r = parseRange(t);
    return r !== null && r.r0 <= rect.r1 && rect.r0 <= r.r1 && r.c0 <= rect.c1 && rect.c0 <= r.c1;
  });
}

export function shiftRange(text: string, axis: 'row' | 'col', at: number, count: number): string | null {
  const m = parseRange(text);
  if (!m) return null;
  const lo = axis === 'row' ? m.r0 : m.c0;
  const hi = axis === 'row' ? m.r1 : m.c1;
  let a = lo;
  let b = hi;
  if (count > 0) {
    if (lo >= at) a += count;
    if (hi >= at) b += count;
  } else {
    const d0 = at;
    const d1 = at - count - 1;
    if (lo >= d0 && hi <= d1) return null;
    a = lo > d1 ? lo + count : lo >= d0 ? d0 : lo;
    b = hi > d1 ? hi + count : hi >= d0 ? d0 - 1 : hi;
  }
  return rangeText(axis === 'row' ? { ...m, r0: a, r1: b } : { ...m, c0: a, c1: b });
}

// ---------- conditional formatting ----------

export const DEFAULT_HIGHLIGHT: HighlightStyle = { bg: '#ffc7ce', color: '#9c0006' };

export interface CellLook {
  bg?: string;
  color?: string;
  b?: boolean;
  /** A data bar: how far it reaches (0–1) and its colour. */
  bar?: { pct: number; color: string };
  /** An icon-set icon: the glyph and its colour. */
  icon?: { glyph: string; color: string };
}

/** How many icons a set has: "3Arrows" → 3. */
export const iconCount = (set: string) => Math.min(5, Math.max(3, Number(/^(\d)/.exec(set)?.[1] ?? 3)));

/**
 * The glyph for level `i` (0 = lowest) of an icon set, drawn as text so it needs
 * no image: arrows, lights, symbols, flags, ratings and the rest by family.
 */
export function iconGlyph(set: string, i: number, n: number): { glyph: string; color: string } {
  const red = '#d93025';
  const amber = '#f9ab00';
  const green = '#1e8e3e';
  const grey = '#80868b';
  const traffic = n === 3 ? [red, amber, green] : n === 4 ? ['#202124', red, amber, green] : [red, '#e37400', amber, '#7cb342', green];
  const name = set.replace(/^\d/, '');
  if (/Arrows/.test(name)) {
    const glyphs = n === 3 ? ['↓', '→', '↑'] : n === 4 ? ['↓', '↘', '↗', '↑'] : ['↓', '↘', '→', '↗', '↑'];
    return { glyph: glyphs[i]!, color: /Gray/.test(name) ? grey : traffic[i]! };
  }
  if (/Traffic|Signs/.test(name)) return { glyph: '●', color: traffic[i]! };
  if (/Symbols/.test(name)) return { glyph: ['✖', '!', '✔'][Math.min(2, Math.round((i * 2) / (n - 1)))]!, color: [red, amber, green][Math.min(2, Math.round((i * 2) / (n - 1)))]! };
  if (/Flags/.test(name)) return { glyph: '⚑', color: [red, amber, green][Math.min(2, Math.round((i * 2) / (n - 1)))]! };
  if (/Stars/.test(name)) return { glyph: ['☆', '⯪', '★'][Math.min(2, i)]!, color: amber };
  if (/Rating|Quarters|Boxes|Triangles/.test(name)) return { glyph: ['○', '◔', '◑', '◕', '●'][Math.round((i * 4) / (n - 1))]!, color: '#1a73e8' };
  if (/RedToBlack/.test(name)) return { glyph: '●', color: ['#202124', grey, '#f28b82', red][Math.min(3, i)]! };
  return { glyph: '●', color: traffic[i]! };
}

const hex = (c: string) => {
  const m = /^#?([0-9a-f]{6})$/i.exec(c.trim());
  const n = m ? Number.parseInt(m[1]!, 16) : 0;
  return [n >> 16, (n >> 8) & 255, n & 255] as const;
};
const mix = (a: string, b: string, t: number) => {
  const x = hex(a);
  const y = hex(b);
  return `#${x.map((v, i) => Math.round(v + (y[i]! - v) * t).toString(16).padStart(2, '0')).join('')}`;
};

const asNumber = (v: Scalar) => (typeof v === 'number' ? v : null);
const shown = (v: Scalar) => (v === null ? '' : v instanceof CellError ? v.code : String(v));

/**
 * Prepares a sheet's rules once per paint: the extremes a colour scale or a
 * data bar stretches over, the values that repeat. Returns what one cell gets.
 */
export function conditionalLook(sheet: SheetData, value: (row: number, col: number) => Scalar, used: { rows: number; cols: number }): (row: number, col: number) => CellLook | null {
  const rules = (sheet.conditional ?? [])
    .map((rule) => {
      const r = parseRange(rule.range);
      if (!r) return null;
      // Whole columns or rows reach only as far as the data.
      const rect = { ...r, r1: Math.min(r.r1, Math.max(used.rows - 1, r.r0)), c1: Math.min(r.c1, Math.max(used.cols - 1, r.c0)) };
      let min = Infinity;
      let max = -Infinity;
      const nums: number[] = [];
      const counts = new Map<string, number>();
      if (rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'duplicate' || rule.type === 'top' || rule.type === 'average' || rule.type === 'iconSet') {
        for (let row = rect.r0; row <= rect.r1; row++) {
          for (let col = rect.c0; col <= rect.c1; col++) {
            const v = value(row, col);
            const n = asNumber(v);
            if (n !== null) {
              nums.push(n);
              if (n < min) min = n;
              if (n > max) max = n;
            }
            if (rule.type === 'duplicate' && v !== null && v !== '') {
              const key = shown(v).toLowerCase();
              counts.set(key, (counts.get(key) ?? 0) + 1);
            }
          }
        }
      }
      nums.sort((a, b) => a - b);
      const mid = nums.length ? nums[Math.floor((nums.length - 1) / 2)]! : 0;
      const mean = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
      // The value a Top/Bottom rule cuts at: the rank-th from the top (or bottom).
      let cut = Number.NaN;
      if (rule.type === 'top' && nums.length) {
        const k = Math.max(1, rule.percent ? Math.floor((nums.length * rule.rank) / 100) : Math.floor(rule.rank));
        cut = rule.bottom ? nums[Math.min(k, nums.length) - 1]! : nums[Math.max(0, nums.length - k)]!;
      }
      return { rule, rect, min, max, mid, counts, mean, cut };
    })
    .filter((x) => x !== null);
  if (!rules.length) return () => null;

  return (row, col) => {
    let look: CellLook | null = null;
    for (const { rule, rect, min, max, mid, counts, mean, cut } of rules) {
      if (!inRect(rect, row, col)) continue;
      const v = value(row, col);
      if (v === null || v === '') continue;
      const n = asNumber(v);
      if (rule.type === 'top' && (n === null || Number.isNaN(cut) || (rule.bottom ? n > cut : n < cut))) continue;
      if (rule.type === 'average' && (n === null || (rule.below ? (rule.equal ? n > mean : n >= mean) : rule.equal ? n < mean : n <= mean))) continue;
      if (rule.type === 'iconSet') {
        if (n === null || !Number.isFinite(min)) continue;
        const count = iconCount(rule.set);
        const cuts = rule.thresholds?.length === count - 1 ? rule.thresholds : Array.from({ length: count - 1 }, (_, i) => Math.round(((i + 1) * 100) / count));
        const pct = max === min ? 100 : ((n - min) / (max - min)) * 100;
        let level = cuts.filter((c) => pct >= c).length;
        if (rule.reverse) level = count - 1 - level;
        look ??= {};
        look.icon = iconGlyph(rule.set, level, count);
        continue;
      }
      const hit = test(rule, v, counts);
      if (hit === false) continue;
      look ??= {};
      if (rule.type === 'colorScale') {
        const n = asNumber(v);
        if (n === null || !Number.isFinite(min)) continue;
        const [lo, a, b] = rule.colors.length >= 3 ? [rule.colors[0]!, rule.colors[1]!, rule.colors[2]!] : [rule.colors[0] ?? '#ffffff', null, rule.colors[1] ?? '#63be7b'];
        if (a && b) look.bg = n <= mid ? mix(lo, a, mid === min ? 1 : (n - min) / (mid - min)) : mix(a, b, max === mid ? 1 : (n - mid) / (max - mid));
        else look.bg = mix(lo, b, max === min ? 1 : (n - min) / (max - min));
        look.color = '#1f1f1f';
      } else if (rule.type === 'dataBar') {
        const n = asNumber(v);
        if (n === null || !Number.isFinite(max)) continue;
        const top = Math.max(max, 0);
        const bottom = Math.min(min, 0);
        look.bar = { pct: top === bottom ? 1 : Math.max(0, (n - bottom) / (top - bottom)), color: rule.color ?? '#638ec6' };
      } else {
        const s = rule.style ?? DEFAULT_HIGHLIGHT;
        if (s.bg) look.bg = s.bg;
        if (s.color) look.color = s.color;
        if (s.b) look.b = true;
      }
    }
    return look;
  };
}

function test(rule: ConditionalRule, v: Scalar, counts: Map<string, number>): boolean {
  if (v instanceof CellError) return false;
  switch (rule.type) {
    case 'gt':
    case 'lt':
    case 'ge':
    case 'le':
    case 'eq':
    case 'ne': {
      // A number in the rule compares with numbers only; text compares with text, ignoring case.
      const raw = rule.value;
      const target = typeof raw === 'number' ? raw : raw !== '' && Number.isFinite(Number(raw)) ? Number(raw) : raw;
      if (typeof target !== typeof v) return rule.type === 'ne';
      const c = compare(v, target);
      return rule.type === 'gt' ? c > 0 : rule.type === 'lt' ? c < 0 : rule.type === 'ge' ? c >= 0 : rule.type === 'le' ? c <= 0 : rule.type === 'eq' ? c === 0 : c !== 0;
    }
    case 'between': {
      const lo = Math.min(rule.value, rule.value2);
      const hi = Math.max(rule.value, rule.value2);
      return typeof v === 'number' && v >= lo && v <= hi;
    }
    case 'contains':
      return shown(v).toLowerCase().includes(rule.text.toLowerCase());
    case 'duplicate':
      return (counts.get(shown(v).toLowerCase()) ?? 0) > 1;
    default:
      return true;
  }
}

// ---------- data validation ----------

export function validationAt(sheet: SheetData, row: number, col: number): string[] | null {
  for (const v of sheet.validation ?? []) {
    const r = parseRange(v.range);
    if (r && inRect(r, row, col)) return v.list;
  }
  return null;
}

// ---------- filters ----------

export const BLANK = '\u0000blank';

/** Rows a filter hides: those whose shown text in a filtered column is on that column's hidden list. */
export function filteredRows(sheet: SheetData, text: (row: number, col: number) => string, used: { rows: number }): Set<number> {
  const out = new Set<number>();
  const f = sheet.filter;
  const r = f ? parseRange(f.range) : null;
  if (!f || !r || !f.hidden) return out;
  const cols = Object.entries(f.hidden)
    .map(([letter, values]) => ({ col: parseA1(`${letter}1`)?.col ?? -1, values: new Set(values) }))
    .filter((x) => x.col >= r.c0 && x.col <= r.c1 && x.values.size);
  const last = Math.min(r.r1, Math.max(used.rows - 1, r.r0));
  for (let row = r.r0 + 1; row <= last; row++) {
    if (cols.some(({ col, values }) => values.has(text(row, col) || BLANK))) out.add(row);
  }
  return out;
}

/** The shown texts of a column inside the filter's data rows, in first-seen order. */
export function columnValues(sheet: SheetData, col: number, text: (row: number, col: number) => string, used: { rows: number }): string[] {
  const r = sheet.filter ? parseRange(sheet.filter.range) : null;
  if (!r) return [];
  const seen = new Set<string>();
  const last = Math.min(r.r1, Math.max(used.rows - 1, r.r0));
  for (let row = r.r0 + 1; row <= last; row++) seen.add(text(row, col) || BLANK);
  return [...seen].sort((a, b) => (a === BLANK ? 1 : b === BLANK ? -1 : a.localeCompare(b, undefined, { numeric: true })));
}
