/**
 * The workbook editor: header, toolbar, formula bar, the grid (frozen panes,
 * merges, formats), sheet tabs and status bar. Cells live in the engine; the
 * rest of a sheet (styles, sizes, merges, freeze, tab colour) lives in `data`.
 */
import {
  CellError,
  columnIndex,
  columnName,
  formatA1,
  formatValue,
  impliedFormat,
  offsetFormula,
  parseA1,
  type Scalar,
  serialFromJsDate,
  Workbook,
  parseFormula,
  type Node as FormulaNode,
} from '@open-spreadsheet/engine';
import { BLANK, columnValues, conditionalLook, type CellLook, DEFAULT_HIGHLIGHT, filteredRows, shiftRange, validationAt } from './features.ts';
import { chartData, chartSvg } from './charts.ts';
import { printHtml } from './print.ts';
import { api } from './home.ts';
import { helpFor, suggest } from './help.ts';
import { isZh, LANGS, lang, locale, T } from './i18n.ts';
import { icon } from './icons.ts';
import { exportXlsx, parseCsv, readSpreadsheetFile, toCsv } from './io.ts';
import type { CellStyle, ChartSpec, ChartType, ConditionalRule, SheetData, WorkbookData } from './types.ts';
import { type Command, closeMenus, commandPalette, confirmDialog, currentTheme, downloadBlob, el, iconButton, type MenuEntry, MOD, pickFile, promptDialog, showMenu, toast, toggleTheme } from './ui.ts';

const HEAD_W = 46;
const HEAD_H = 24;
const DEFAULT_COL = 100;
const DEFAULT_ROW = 24;
/** Excel's default font size; it draws at the grid's 13px, and other sizes in proportion. */
const BASE_PT = 11;
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];
const ptToPx = (pt: number) => (pt * 13) / BASE_PT;
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
const REF_COLORS = ['#1a73e8', '#e8710a', '#9334e6', '#188038', '#d93025', '#12b5cb'];
const TEXT_COLORS = ['#000000', '#434343', '#666666', '#999999', '#cccccc', '#ffffff', '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];
const FILL_COLORS = [null, '#f2f2f2', '#d9d9d9', '#fde9e7', '#fce4d6', '#fff2cc', '#e2efda', '#ddebf7', '#e4dfec', '#ffc7ce', '#ffeb9c', '#c6efce', '#bdd7ee', '#ffff00', '#00b050', '#217346'];
const TAB_COLORS = [null, '#e53935', '#fb8c00', '#fdd835', '#43a047', '#00897b', '#1e88e5', '#3949ab', '#8e24aa', '#6d4c41', '#757575'];

interface Pos {
  row: number;
  col: number;
}
interface Rect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}
interface Snapshot {
  json: string;
  si: number;
  active: Pos;
  anchor: Pos;
}
interface Clip {
  text: string;
  cells: { input: string; style?: CellStyle }[][];
  from: Rect;
  sheet: string;
  cut: boolean;
}

const rectOf = (a: Pos, b: Pos): Rect => ({ r0: Math.min(a.row, b.row), c0: Math.min(a.col, b.col), r1: Math.max(a.row, b.row), c1: Math.max(a.col, b.col) });
const inRect = (r: Rect, row: number, col: number) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;
const overlaps = (a: Rect, b: Rect) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
const rangeText = (r: Rect) => (r.r0 === r.r1 && r.c0 === r.c1 ? formatA1(r.r0, r.c0) : `${formatA1(r.r0, r.c0)}:${formatA1(r.r1, r.c1)}`);
function parseRange(text: string): Rect | null {
  const [a, b] = text.trim().toUpperCase().replaceAll('$', '').split(':');
  const p = parseA1(a ?? '');
  const q = b ? parseA1(b) : p;
  return p && q ? rectOf(p, q) : null;
}
const isDateFormat = (f?: string) => Boolean(f && /[yd]|h:|m+:|:s/i.test(f.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '')));

/** Dark or light text for a fill colour, by its relative luminance. */
function inkFor(bg: string): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(bg.trim());
  if (!m) return 'inherit';
  const n = Number.parseInt(m[1]!, 16);
  const lin = (c: number) => {
    const v = c / 255;
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  };
  const l = 0.2126 * lin(n >> 16) + 0.7152 * lin((n >> 8) & 255) + 0.0722 * lin(n & 255);
  return l > 0.36 ? '#1f1f1f' : '#ffffff';
}

/** The currency button's format, by interface language. */
const currencyFormat = () => ({ en: '$#,##0.00', zh: '"NT$"#,##0', 'zh-CN': '"¥"#,##0.00', ja: '"¥"#,##0', ko: '"₩"#,##0' })[lang];
const longDate = () =>
  isZh() || lang === 'ja'
    ? { code: 'yyyy"年"m"月"d"日"', sample: '2026年11月15日' }
    : lang === 'ko'
      ? { code: 'yyyy"년" m"월" d"일"', sample: '2026년 11월 15일' }
      : { code: 'mmmm d, yyyy', sample: 'November 15, 2026' };

let teardown: (() => void) | null = null;

export async function openBook(app: HTMLElement, id: string): Promise<void> {
  teardown?.();
  let data: WorkbookData;
  /** The version on disk this page last read or wrote; saves send it so they never overwrite an agent's change. */
  let etag = '';
  try {
    ({ data, etag } = await api.load(id));
  } catch {
    location.hash = '#/';
    return;
  }
  if (!data.sheets?.length) data.sheets = [{ name: 'Sheet1', cells: {} }];
  let wb = Workbook.fromJSON(data);
  let si = 0;
  try {
    const saved = Number(localStorage.getItem(`open-spreadsheet:sheet:${id}`));
    if (saved > 0 && saved < data.sheets.length) si = saved;
  } catch {
    // first sheet
  }
  let zoom = 1;
  let active: Pos = { row: 0, col: 0 };
  let anchor: Pos = { row: 0, col: 0 };
  /** null, or how the cell is being edited: typed over ('enter') or opened for editing ('edit'). */
  let editing: { mode: 'enter' | 'edit'; source: 'cell' | 'bar' } | null = null;
  /** While a formula is edited, the reference being pointed at with the mouse or arrows. */
  let pointer: { start: number; end: number; anchor: Pos; at: Pos } | null = null;
  let clip: Clip | null = null;
  const undoStack: Snapshot[] = [];
  const redoStack: Snapshot[] = [];
  let extraRows = 0;
  let extraCols = 0;

  const sheet = (): SheetData => data.sheets[si]!;
  const sname = () => sheet().name;
  document.title = `${data.title} — open-spreadsheet`;

  // ---------- DOM ----------
  const root = el('div', 'book');

  // Header
  const head = el('header', 'book-head');
  const back = iconButton('arrowLeft', T.back, () => (location.hash = '#/'));
  const titleBtn = el('button', 'book-title', data.title);
  titleBtn.type = 'button';
  titleBtn.title = T.rename;
  const status = el('span', 'book-status', T.saved);
  const findBtn = iconButton('search', `${T.findReplace} (${MOD}F)`, () => openFind());
  const themeBtn = iconButton(currentTheme() === 'dark' ? 'moon' : 'sun', T.theme, () => {
    toggleTheme();
    themeBtn.replaceChildren(icon(currentTheme() === 'dark' ? 'moon' : 'sun'));
    render();
  });
  const helpBtn = iconButton('keyboard', T.shortcuts, () => showShortcuts());
  const paletteBtn = iconButton('list', `${T.palette.replace('…', '')} (${MOD}K)`, () => void openPalette());
  const notesBtn = iconButton('pencil', T.notes, () => notesMenu());
  const notesBadge = el('span', 'badge');
  notesBtn.append(notesBadge);
  const downloadBtn = el('button', 'btn primary');
  downloadBtn.type = 'button';
  downloadBtn.append(icon('download', 15), el('span', '', T.download), icon('chevronDown', 14));
  const fileMark = el('span', 'brand-mark');
  fileMark.append(icon('sheet', 14));
  head.append(back, fileMark, titleBtn, status, el('div', 'grow'), notesBtn, paletteBtn, findBtn, helpBtn, themeBtn, downloadBtn);

  // Toolbar
  const toolbar = el('div', 'toolbar');
  const tb = {
    undo: iconButton('undo', `${T.undo} (${MOD}Z)`, () => undo()),
    redo: iconButton('redo', `${T.redo} (⇧${MOD}Z)`, () => redo()),
    format: el('button', 'tbtn select'),
    size: el('button', 'tbtn select size'),
    currency: iconButton('dollar', T.fmtCurrency, () => setFormat(currencyFormat())),
    percent: iconButton('percent', T.fmtPercent, () => setFormat('0%')),
    dec: iconButton('decDec', T.decDecimals, () => bumpDecimals(-1)),
    inc: iconButton('incDec', T.incDecimals, () => bumpDecimals(1)),
    bold: iconButton('bold', `${T.bold} (${MOD}B)`, () => toggleStyle('b')),
    italic: iconButton('italic', `${T.italic} (${MOD}I)`, () => toggleStyle('i')),
    underline: iconButton('underline', `${T.underline} (${MOD}U)`, () => toggleStyle('u')),
    strike: iconButton('strike', T.strike, () => toggleStyle('s')),
    color: iconButton('textColor', T.textColor),
    fill: iconButton('fill', T.fillColor),
    border: iconButton('border', T.borders),
    left: iconButton('alignLeft', T.alignLeft, () => setAlign('left')),
    center: iconButton('alignCenter', T.alignCenter, () => setAlign('center')),
    right: iconButton('alignRight', T.alignRight, () => setAlign('right')),
    wrap: iconButton('wrap', T.wrap, () => toggleStyle('wrap')),
    merge: iconButton('merge', T.merge, () => toggleMerge()),
    freeze: iconButton('snowflake', T.freeze),
    sort: iconButton('sortAsc', T.sort),
    sum: iconButton('sigma', T.autosum),
    filter: iconButton('filter', T.filterLabel),
    cond: iconButton('conditional', T.conditional),
    chart: iconButton('chart', T.insertChart),
    painter: iconButton('brush', `${T.formatPainter} (⇧${MOD}C / ⇧${MOD}V)`),
  };
  tb.format.type = 'button';
  tb.format.title = T.format;
  tb.size.type = 'button';
  tb.size.title = `${T.fontSize} (⇧${MOD}> / ⇧${MOD}<)`;
  const colorBar = el('span', 'color-bar');
  tb.color.append(colorBar);
  const fillBar = el('span', 'color-bar');
  tb.fill.append(fillBar);
  const sep = () => el('span', 'tsep');
  toolbar.append(
    tb.undo, tb.redo, tb.painter, sep(),
    tb.format, tb.currency, tb.percent, tb.dec, tb.inc, sep(),
    tb.size, tb.bold, tb.italic, tb.underline, tb.strike, tb.color, tb.fill, tb.border, sep(),
    tb.left, tb.center, tb.right, tb.wrap, tb.merge, sep(),
    tb.freeze, tb.sort, tb.filter, tb.cond, tb.chart, tb.sum,
  );
  // Toolbar clicks keep the keyboard focus on the grid.
  toolbar.addEventListener('mousedown', (e) => {
    if ((e.target as HTMLElement).closest('button')) e.preventDefault();
  });

  // Formula bar
  const fbar = el('div', 'fbar');
  const namebox = el('input', 'namebox');
  namebox.spellcheck = false;
  namebox.title = T.nameboxTitle;
  const fx = el('span', 'fx');
  fx.append(icon('fx', 16));
  const finput = el('input', 'finput');
  finput.spellcheck = false;
  fbar.append(namebox, fx, finput);

  // Grid
  const grid = el('div', 'grid');
  const corner = el('div', 'corner');
  const colheads = el('div', 'colheads');
  const colFrozen = el('div', 'heads-frozen');
  const colScroll = el('div', 'heads-scroll');
  const colInner = el('div', 'heads-inner');
  colScroll.append(colInner);
  colheads.append(colFrozen, colScroll);
  const rowheads = el('div', 'rowheads');
  const rowFrozen = el('div', 'heads-frozen');
  const rowScroll = el('div', 'heads-scroll');
  const rowInner = el('div', 'heads-inner');
  rowScroll.append(rowInner);
  rowheads.append(rowFrozen, rowScroll);
  const paneMain = el('div', 'pane');
  const paneTop = el('div', 'pane frozen-pane');
  const paneLeft = el('div', 'pane frozen-pane');
  const paneCorner = el('div', 'pane frozen-pane');
  const body = el('div', 'body');
  const spacer = el('div', 'spacer');
  body.append(spacer);
  const overlay = el('div', 'overlay');
  const editor = el('input', 'editor idle');
  editor.spellcheck = false;
  editor.autocomplete = 'off';
  editor.setAttribute('aria-label', 'cell editor');
  const fillHandle = el('div', 'fill-handle');
  const dvBtn = el('button', 'dv-btn');
  dvBtn.type = 'button';
  dvBtn.hidden = true;
  dvBtn.append(icon('chevronDown', 12));
  const popup = el('div', 'ac-popup');
  popup.hidden = true;
  const chartsLayer = el('div', 'charts-layer');
  grid.append(paneMain, paneTop, paneLeft, paneCorner, body, overlay, chartsLayer, colheads, rowheads, corner);
  overlay.append(fillHandle, dvBtn, editor);

  // Find bar
  const findbar = el('div', 'findbar');
  findbar.hidden = true;

  // Tabs and status
  const tabsBar = el('div', 'tabs-bar');
  const addTab = iconButton('plus', T.addSheet, () => addSheet());
  const listTabs = iconButton('list', T.allSheets);
  const tabs = el('div', 'tabs');
  const stats = el('div', 'stats');
  const zoomSel = el('select', 'zoom');
  zoomSel.title = T.zoom;
  for (const z of [50, 75, 90, 100, 110, 125, 150, 200]) {
    const o = el('option', '', `${z}%`);
    o.value = String(z / 100);
    if (z === 100) o.selected = true;
    zoomSel.append(o);
  }
  tabsBar.append(addTab, listTabs, tabs, el('div', 'grow'), stats, zoomSel);

  const gridWrap = el('div', 'grid-wrap');
  gridWrap.append(grid, findbar);
  root.append(head, toolbar, fbar, gridWrap, tabsBar);
  app.replaceChildren(root);
  document.body.append(popup);

  // ---------- geometry ----------
  let colX: number[] = [0];
  let rowY: number[] = [0];
  let nRows = 0;
  let nCols = 0;
  let mergeAt = new Map<string, Rect>(); // "r,c" of every covered cell -> its merge
  let merges: Rect[] = [];

  let hiddenColSet = new Set<number>();
  const colW = (c: number) => (hiddenColSet.has(c) ? 0 : Math.round((sheet().cols?.[columnName(c)] ?? DEFAULT_COL) * zoom));
  /** Rows the sheet's filter hides; they get no height. */
  let hiddenRows = new Set<number>();
  /** The largest font size on each row, for rows whose height nobody set: they grow to fit, as in Excel. */
  let rowFont = new Map<number, number>();
  const rowH = (r: number) => {
    if (hiddenRows.has(r)) return 0;
    const set = sheet().rows?.[String(r + 1)];
    const auto = Math.max(DEFAULT_ROW, Math.ceil(ptToPx(rowFont.get(r) ?? BASE_PT) * 1.25 + 8));
    return Math.round((set ?? auto) * zoom);
  };
  const fr = () => Math.min(sheet().freeze?.rows ?? 0, nRows - 1);
  const fc = () => Math.min(sheet().freeze?.cols ?? 0, nCols - 1);

  const layout = () => {
    const used = wb.used(sname());
    nRows = Math.min(MAX_ROWS, Math.max(100, used.rows + 50, active.row + 30, anchor.row + 30) + extraRows);
    nCols = Math.min(MAX_COLS, Math.max(26, used.cols + 6, active.col + 6, anchor.col + 6) + extraCols);
    hiddenRows = filteredRows(sheet(), (r, c) => (wb.inputAt(sname(), r, c) ? show(r, c).text : ''), used);
    for (const n of sheet().hiddenRows ?? []) hiddenRows.add(n - 1);
    hiddenColSet = new Set((sheet().hiddenCols ?? []).map((l) => columnIndex(l)));
    colX = new Array(nCols + 1);
    colX[0] = 0;
    for (let c = 0; c < nCols; c++) colX[c + 1] = colX[c]! + colW(c);
    rowFont = new Map();
    for (const [a, st] of Object.entries(sheet().styles ?? {})) {
      if (!st.size || st.size <= BASE_PT) continue;
      const p = parseA1(a);
      if (p && st.size > (rowFont.get(p.row) ?? 0)) rowFont.set(p.row, st.size);
    }
    rowY = new Array(nRows + 1);
    rowY[0] = 0;
    for (let r = 0; r < nRows; r++) rowY[r + 1] = rowY[r]! + rowH(r);
    merges = (sheet().merges ?? []).map(parseRange).filter((m): m is Rect => m !== null);
    mergeAt = new Map();
    for (const m of merges) for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) mergeAt.set(`${r},${c}`, m);
    grid.style.setProperty('--z', String(zoom));
  };

  /** Index i with arr[i] <= x < arr[i+1]. */
  const search = (arr: number[], x: number, n: number) => {
    let lo = 0;
    let hi = n - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (arr[mid]! <= x) lo = mid;
      else hi = mid - 1;
    }
    return Math.max(0, Math.min(n - 1, lo));
  };
  const frozenW = () => colX[fc()]!;
  const frozenH = () => rowY[fr()]!;
  /** Sheet x of a point in the grid's cell area (after the headers). */
  const sheetX = (x: number) => (x < frozenW() ? x : x + body.scrollLeft);
  const sheetY = (y: number) => (y < frozenH() ? y : y + body.scrollTop);
  const viewX = (c: number) => (c < fc() ? colX[c]! : colX[c]! - body.scrollLeft);
  const viewY = (r: number) => (r < fr() ? rowY[r]! : rowY[r]! - body.scrollTop);

  /** Where a range shows in the cell area, clipped against the frozen panes, or null when hidden. */
  const viewRect = (r: Rect) => {
    let x0 = viewX(r.c0);
    let x1 = viewX(r.c1) + (colX[r.c1 + 1]! - colX[r.c1]!);
    let y0 = viewY(r.r0);
    let y1 = viewY(r.r1) + (rowY[r.r1 + 1]! - rowY[r.r1]!);
    if (r.c0 >= fc()) x0 = Math.max(x0, frozenW());
    if (r.c1 >= fc()) x1 = Math.max(x1, frozenW());
    if (r.r0 >= fr()) y0 = Math.max(y0, frozenH());
    if (r.r1 >= fr()) y1 = Math.max(y1, frozenH());
    if (r.c0 < fc() && r.c1 >= fc()) x1 = Math.max(x1, frozenW());
    if (x1 <= x0 || y1 <= y0) return null;
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  };

  const mergeOf = (row: number, col: number) => mergeAt.get(`${row},${col}`);
  /** A range grown until it holds every merge it touches. */
  const expand = (r: Rect): Rect => {
    const out = { ...r };
    for (let changed = true; changed; ) {
      changed = false;
      for (const m of merges) {
        if (overlaps(out, m) && (m.r0 < out.r0 || m.c0 < out.c0 || m.r1 > out.r1 || m.c1 > out.c1)) {
          out.r0 = Math.min(out.r0, m.r0);
          out.c0 = Math.min(out.c0, m.c0);
          out.r1 = Math.max(out.r1, m.r1);
          out.c1 = Math.max(out.c1, m.c1);
          changed = true;
        }
      }
    }
    return out;
  };
  const selection = (): Rect => expand(rectOf(active, anchor));

  // ---------- values and styles ----------
  const styleAt = (row: number, col: number): CellStyle | undefined => sheet().styles?.[formatA1(row, col)];
  const noteAt = (row: number, col: number): string | undefined => sheet().notes?.[formatA1(row, col)];
  const filterRect = () => (sheet().filter ? parseRange(sheet().filter!.range) : null);
  const isFilterHead = (row: number, col: number) => {
    const r = filterRect();
    return Boolean(r && row === r.r0 && col >= r.c0 && col <= r.c1);
  };
  const filterHidden = (col: number) => sheet().filter?.hidden?.[columnName(col)] ?? [];
  const noteCount = () => data.sheets.reduce((n, sh) => n + Object.keys(sh.notes ?? {}).length, 0);
  async function editNote(row: number, col: number) {
    const a = formatA1(row, col);
    const v = await promptDialog({ title: `${noteAt(row, col) ? T.editNote : T.addNote} · ${a}`, body: T.noteHint, fields: [{ name: 'text', label: T.noteLabel, value: noteAt(row, col) ?? '', multiline: true }], ok: T.save });
    editor.focus();
    if (!v) return;
    setNote(row, col, v.text!.trim());
  }
  function setNote(row: number, col: number, text: string) {
    change(() => {
      const sh = sheet();
      sh.notes ??= {};
      if (text) sh.notes[formatA1(row, col)] = text;
      else delete sh.notes[formatA1(row, col)];
    }, { relayout: false });
    paintNotesBadge();
  }

  /** The format a cell shows with: its own, else what its input implies, else its formula's first reference's. */
  const formatOf = (row: number, col: number, depth = 0): string | undefined => {
    const own = styleAt(row, col)?.fmt;
    if (own) return own;
    const input = wb.inputAt(sname(), row, col);
    if (!input.startsWith('=')) return impliedFormat(input);
    if (depth > 3 || !/^=\s*(\(?\$?[A-Z]{1,3}\$?\d+|SUM\(|MIN\(|MAX\(|AVERAGE\(|ROUND\()/i.test(input)) return undefined;
    const m = /(?<![A-Za-z_!'])\$?([A-Z]{1,3})\$?(\d+)/i.exec(input.slice(1));
    if (!m) return undefined;
    const ref = parseA1(`${m[1]}${m[2]}`);
    if (!ref || (ref.row === row && ref.col === col)) return undefined;
    const f = formatOf(ref.row, ref.col, depth + 1);
    // A date times two is not a date: dates carry over only through + and -.
    if (f && isDateFormat(f)) return /[*/^]/.test(input.replace(/"[^"]*"/g, '')) ? undefined : f;
    return f && /[%$]|NT\$|#,##0/.test(f) ? f : undefined;
  };

  /**
   * Where a cell links to: HYPERLINK's first argument (a string, a reference or
   * text joined with &), or a value that is itself a web or mail address. Only
   * http(s) and mailto, so a sheet can never run javascript: on a click.
   */
  const linkAt = (row: number, col: number): string | null => {
    const input = wb.inputAt(sname(), row, col);
    let link: unknown = null;
    if (/^=\s*HYPERLINK\s*\(/i.test(input)) {
      const textOf = (n: FormulaNode): string | null => {
        if (n.kind === 'str') return n.value;
        if (n.kind === 'ref') {
          const v = wb.value(n.sheet ?? sname(), n.row, n.col);
          return typeof v === 'string' ? v : null;
        }
        if (n.kind === 'binary' && n.op === '&') {
          const a = textOf(n.left);
          const b = textOf(n.right);
          return a === null || b === null ? null : a + b;
        }
        if (n.kind === 'paren') return textOf(n.arg);
        return null;
      };
      try {
        const node = parseFormula(input);
        if (node.kind === 'call' && node.name.toUpperCase() === 'HYPERLINK' && node.args[0]) link = textOf(node.args[0]);
      } catch {
        link = null;
      }
    } else if (!input.startsWith('=')) link = wb.value(sname(), row, col);
    if (typeof link !== 'string' || !/^(https?:\/\/|mailto:)\S+$/i.test(link.trim())) return null;
    try {
      return new URL(link.trim()).href;
    } catch {
      return null;
    }
  };
  const openLink = (url: string) => window.open(url, '_blank', 'noopener,noreferrer');

  /** Excel's Show Formulas (⌃`): formula cells show what was typed, not the result. */
  let showFormulas = false;
  const show = (row: number, col: number) => {
    if (showFormulas) {
      const input = wb.inputAt(sname(), row, col);
      if (input.startsWith('=')) return { v: input as Scalar, text: input, color: undefined as string | undefined };
    }
    const v = wb.value(sname(), row, col);
    const f = formatValue(v, formatOf(row, col));
    return { v, text: f.text, color: f.color };
  };

  // ---------- painting ----------
  function paintPane(pane: HTMLElement, rows: [number, number], cols: [number, number], ox: number, oy: number) {
    const frag = document.createDocumentFragment();
    const [r0, r1] = rows;
    const [c0, c1] = cols;
    // Grid lines.
    for (let r = r0; r < r1; r++) {
      const line = el('div', 'hline');
      line.style.top = `${rowY[r + 1]! - 1 - oy}px`;
      frag.append(line);
    }
    for (let c = c0; c < c1; c++) {
      const line = el('div', 'vline');
      line.style.left = `${colX[c + 1]! - 1 - ox}px`;
      frag.append(line);
    }
    const drawn = new Set<Rect>();
    const name = sname();
    for (let r = r0; r < r1; r++) {
      if (rowY[r + 1] === rowY[r]) continue; // hidden by a filter
      // Start a little to the left, so text spilling in from a hidden column still shows.
      const start = Math.max(cols === undefined ? 0 : c0 - 8, c0 >= fc() ? fc() : 0);
      for (let c = start; c < c1; c++) {
        if (colX[c + 1] === colX[c]) continue; // a hidden column
        const m = mergeOf(r, c);
        if (m) {
          if (drawn.has(m)) continue;
          drawn.add(m);
          frag.append(cellNode(m.r0, m.c0, colX[m.c0]! - ox, rowY[m.r0]! - oy, colX[m.c1 + 1]! - colX[m.c0]!, rowY[m.r1 + 1]! - rowY[m.r0]!, true));
          continue;
        }
        const input = wb.inputAt(name, r, c);
        const style = styleAt(r, c);
        if (!input && !style?.bg && !style?.border && !style?.check && !noteAt(r, c) && !isFilterHead(r, c)) continue;
        let w = colX[c + 1]! - colX[c]!;
        const s = input ? show(r, c) : null;
        const isText = s && (typeof s.v === 'string' || s.v === null) && !style?.wrap && (style?.align ?? 'left') === 'left';
        if (isText && s.text) {
          let e = c + 1;
          while (e < c1 + 8 && e < nCols && !wb.inputAt(name, r, e) && !mergeOf(r, e) && e - c < 30) e++;
          w = colX[e]! - colX[c]!;
          // Only as wide as the text needs, so it does not hide grid lines it does not cover.
          const need = measure(s.text, style) + 10;
          if (need < w) {
            let k = c + 1;
            while (colX[k]! - colX[c]! < need && k < e) k++;
            w = colX[k]! - colX[c]!;
          }
        }
        if (c < c0 && colX[c]! + w <= colX[c0]!) continue;
        frag.append(cellNode(r, c, colX[c]! - ox, rowY[r]! - oy, w, rowY[r + 1]! - rowY[r]!, false, s, w > colX[c + 1]! - colX[c]!));
      }
    }
    // Merges that start above or left of the painted rows still show.
    for (const m of merges) {
      if (drawn.has(m) || m.r1 < r0 || m.r0 >= r1 || m.c1 < c0 || m.c0 >= c1) continue;
      frag.append(cellNode(m.r0, m.c0, colX[m.c0]! - ox, rowY[m.r0]! - oy, colX[m.c1 + 1]! - colX[m.c0]!, rowY[m.r1 + 1]! - rowY[m.r0]!, true));
    }
    pane.replaceChildren(frag);
  }

  const measureCtx = document.createElement('canvas').getContext('2d')!;
  function measure(text: string, style?: CellStyle) {
    measureCtx.font = `${style?.i ? 'italic ' : ''}${style?.b ? '600 ' : ''}${ptToPx(style?.size ?? BASE_PT) * zoom}px -apple-system, BlinkMacSystemFont, "PingFang TC", sans-serif`;
    return measureCtx.measureText(text).width;
  }

  function cellNode(row: number, col: number, x: number, y: number, w: number, h: number, merged: boolean, s?: ReturnType<typeof show> | null, spills = false) {
    const style = styleAt(row, col);
    const shown = s ?? (wb.inputAt(sname(), row, col) ? show(row, col) : null);
    const d = el('div', 'cell');
    d.style.cssText = `left:${x}px;top:${y}px;width:${w}px;height:${h}px`;
    if (style?.check && !showFormulas && !(shown?.v instanceof CellError)) {
      const on = shown?.v === true;
      const box = el('span', `cb${on ? ' on' : ''}`);
      box.setAttribute('role', 'checkbox');
      box.setAttribute('aria-checked', String(on));
      if (on) box.innerHTML = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="m3.5 8.5 3 3 6-7" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
      d.append(box);
      d.classList.add('a-center', 'has-cb');
      if (style.bg) d.style.background = style.bg;
      return d;
    }
    if (shown) {
      const span = el('span', 'cell-text', shown.text);
      d.append(span);
      const v = shown.v;
      let align = style?.align;
      if (!align) align = typeof v === 'number' ? 'right' : typeof v === 'boolean' || v instanceof CellError ? 'center' : 'left';
      d.classList.add(`a-${align}`);
      if (v instanceof CellError) d.classList.add('err');
      else if (linkAt(row, col)) {
        d.classList.add('link');
        d.title = `${linkAt(row, col)}\n${T.openLinkHint}`;
      }
      if (shown.color) d.style.color = shown.color;
    }
    if (style) {
      if (style.b) d.classList.add('b');
      if (style.i) d.classList.add('i');
      if (style.u || style.s) d.style.textDecoration = [style.u ? 'underline' : '', style.s ? 'line-through' : ''].join(' ').trim();
      if (style.color && !shown?.color) d.style.color = style.color;
      if (style.bg) {
        d.style.background = style.bg;
        // A fill is the same in both themes, so its text follows the fill, not the theme.
        if (!style.color && !shown?.color) d.style.color = inkFor(style.bg);
      }
      if (style.wrap) d.classList.add('wrap');
      if (style.valign) d.classList.add(`v-${style.valign}`);
      if (style.size && style.size !== BASE_PT) d.style.fontSize = `calc(${ptToPx(style.size).toFixed(2)}px * var(--z))`;
      if (style.border) {
        const w = style.border.thick ? 2 : 1;
        const b = style.border;
        // Black lines follow the theme's ink, so they stay visible in dark mode.
        const ink = (x?: string) => (x && /^#(1f1f1f|000000)$/i.test(x) ? 'var(--cell-fg)' : x);
        const shadows = [b.top && `inset 0 ${w}px 0 ${ink(b.top)}`, b.bottom && `inset 0 -${w}px 0 ${ink(b.bottom)}`, b.left && `inset ${w}px 0 0 ${ink(b.left)}`, b.right && `inset -${w}px 0 0 ${ink(b.right)}`].filter(Boolean);
        if (shadows.length) d.style.boxShadow = shadows.join(', ');
      }
    }
    const look = shown ? lookAt(row, col) : null;
    if (look) {
      if (look.bg) d.style.background = look.bg;
      if (look.color) d.style.color = look.color;
      if (look.b) d.classList.add('b');
      if (look.bar) {
        const bar = el('span', 'data-bar');
        bar.style.width = `${Math.round(look.bar.pct * 100)}%`;
        bar.style.background = `linear-gradient(90deg, ${look.bar.color}, ${look.bar.color}55)`;
        d.prepend(bar);
      }
    }
    if (isFilterHead(row, col)) {
      const fb = el('span', `filter-btn${filterHidden(col).length ? ' on' : ''}`);
      fb.dataset.col = String(col);
      fb.append(icon(filterHidden(col).length ? 'filter' : 'chevronDown', 12));
      d.append(fb);
      d.classList.add('filter-head');
    }
    if (merged || spills) d.classList.add('solid');
    if (noteAt(row, col)) {
      const mark = el('span', 'note-mark');
      // On a spilling cell the mark belongs to the cell's own corner, not the spill's.
      if (spills) mark.style.left = `${colX[col + 1]! - colX[col]! - 7}px`;
      d.append(mark);
    }
    return d;
  }

  function visibleRange(start: number, size: number, arr: number[], n: number, from: number): [number, number] {
    const a = Math.max(from, search(arr, start, n));
    const b = Math.min(n, search(arr, start + size, n) + 1);
    return [a, Math.max(a, b)];
  }

  function paintHeads() {
    const sel = selection();
    const full = sel.r0 === 0 && sel.r1 >= nRows - 1;
    const fullC = sel.c0 === 0 && sel.c1 >= nCols - 1;
    const cw = grid.clientWidth - HEAD_W;
    const ch = grid.clientHeight - HEAD_H;
    const [mc0, mc1] = visibleRange(frozenW() + body.scrollLeft, cw - frozenW(), colX, nCols, fc());
    const [mr0, mr1] = visibleRange(frozenH() + body.scrollTop, ch - frozenH(), rowY, nRows, fr());
    const colHead = (c: number, ox: number) => {
      if (colX[c + 1] === colX[c]) return document.createComment('hidden');
      const h = el('div', `head${c >= sel.c0 && c <= sel.c1 ? (full ? ' picked' : ' active') : ''}`, columnName(c));
      h.style.cssText = `left:${colX[c]! - ox}px;width:${colX[c + 1]! - colX[c]!}px;height:${HEAD_H}px`;
      h.dataset.col = String(c);
      if (c > 0 && colX[c] === colX[c - 1]) h.classList.add('after-hidden');
      return h;
    };
    const rowHead = (r: number, oy: number) => {
      if (rowY[r + 1] === rowY[r]) return document.createComment('hidden');
      const h = el('div', `head${r >= sel.r0 && r <= sel.r1 ? (fullC ? ' picked' : ' active') : ''}`, String(r + 1));
      h.style.cssText = `top:${rowY[r]! - oy}px;height:${rowY[r + 1]! - rowY[r]!}px;width:${HEAD_W}px`;
      h.dataset.row = String(r);
      if (r > 0 && rowY[r] === rowY[r - 1]) h.classList.add('after-hidden');
      return h;
    };
    colFrozen.style.width = `${frozenW()}px`;
    colScroll.style.left = `${frozenW()}px`;
    rowFrozen.style.height = `${frozenH()}px`;
    rowScroll.style.top = `${frozenH()}px`;
    const cf = document.createDocumentFragment();
    for (let c = 0; c < fc(); c++) cf.append(colHead(c, 0));
    colFrozen.replaceChildren(cf);
    const cs = document.createDocumentFragment();
    for (let c = mc0; c < mc1; c++) cs.append(colHead(c, frozenW() + body.scrollLeft));
    colInner.replaceChildren(cs);
    const rf = document.createDocumentFragment();
    for (let r = 0; r < fr(); r++) rf.append(rowHead(r, 0));
    rowFrozen.replaceChildren(rf);
    const rs = document.createDocumentFragment();
    for (let r = mr0; r < mr1; r++) rs.append(rowHead(r, frozenH() + body.scrollTop));
    rowInner.replaceChildren(rs);
    corner.classList.toggle('picked', full && fullC);
  }

  /** What conditional formatting gives each cell; rebuilt every paint. */
  let lookAt: (row: number, col: number) => CellLook | null = () => null;
  function render() {
    lookAt = conditionalLook(sheet(), (r, c) => wb.value(sname(), r, c), wb.used(sname()));
    const cw = grid.clientWidth - HEAD_W;
    const ch = grid.clientHeight - HEAD_H;
    const fw = frozenW();
    const fh = frozenH();
    spacer.style.width = `${colX[nCols]! - fw}px`;
    spacer.style.height = `${rowY[nRows]! - fh}px`;
    for (const [node, css] of [
      [body, { left: `${HEAD_W + fw}px`, top: `${HEAD_H + fh}px` }],
      [paneMain, { left: `${HEAD_W + fw}px`, top: `${HEAD_H + fh}px`, right: '0', bottom: '0' }],
      [paneTop, { left: `${HEAD_W + fw}px`, top: `${HEAD_H}px`, height: `${fh}px`, right: '0' }],
      [paneLeft, { left: `${HEAD_W}px`, top: `${HEAD_H + fh}px`, width: `${fw}px`, bottom: '0' }],
      [paneCorner, { left: `${HEAD_W}px`, top: `${HEAD_H}px`, width: `${fw}px`, height: `${fh}px` }],
    ] as const) Object.assign(node.style, css);
    paneTop.hidden = fh === 0;
    paneLeft.hidden = fw === 0;
    paneCorner.hidden = fh === 0 || fw === 0;
    grid.classList.toggle('has-freeze-row', fh > 0);
    grid.classList.toggle('has-freeze-col', fw > 0);
    grid.style.setProperty('--fw', `${fw}px`);
    grid.style.setProperty('--fh', `${fh}px`);

    const cols = visibleRange(fw + body.scrollLeft, cw - fw, colX, nCols, fc());
    const rows = visibleRange(fh + body.scrollTop, ch - fh, rowY, nRows, fr());
    const ox = fw + body.scrollLeft;
    const oy = fh + body.scrollTop;
    paintPane(paneMain, rows, cols, ox, oy);
    if (fh) paintPane(paneTop, [0, fr()], cols, ox, 0);
    else paneTop.replaceChildren();
    if (fw) paintPane(paneLeft, rows, [0, fc()], 0, oy);
    else paneLeft.replaceChildren();
    if (fh && fw) paintPane(paneCorner, [0, fr()], [0, fc()], 0, 0);
    else paneCorner.replaceChildren();
    paintHeads();
    paintOverlay();
    paintCharts();
    paintToolbar();
    paintStats();
  }

  // ---------- overlay: selection, references, editor ----------
  let copyBox: HTMLElement | null = null;
  function paintOverlay() {
    for (const n of [...overlay.querySelectorAll('.sel, .active-box, .ref-box, .copy-box, .fill-preview, .trace-box, .trace-arrows')]) n.remove();
    paintTrace();
    const sel = selection();
    const multi = sel.r0 !== sel.r1 || sel.c0 !== sel.c1;
    const am = mergeOf(active.row, active.col) ?? { r0: active.row, c0: active.col, r1: active.row, c1: active.col };
    const vr = viewRect(sel);
    if (vr && multi) {
      const box = el('div', 'sel');
      box.style.cssText = `left:${vr.x}px;top:${vr.y}px;width:${vr.w}px;height:${vr.h}px`;
      overlay.prepend(box);
    }
    const av = viewRect(am);
    if (av) {
      const box = el('div', 'active-box');
      box.style.cssText = `left:${av.x - 1}px;top:${av.y - 1}px;width:${av.w + 1}px;height:${av.h + 1}px`;
      overlay.prepend(box);
    }
    if (vr && !editing) {
      fillHandle.hidden = false;
      fillHandle.style.left = `${vr.x + vr.w - 4}px`;
      fillHandle.style.top = `${vr.y + vr.h - 4}px`;
    } else fillHandle.hidden = true;
    if (clip && clip.sheet === sname()) {
      const cv = viewRect(clip.from);
      if (cv) {
        copyBox = el('div', 'copy-box');
        copyBox.style.cssText = `left:${cv.x - 1}px;top:${cv.y - 1}px;width:${cv.w + 1}px;height:${cv.h + 1}px`;
        overlay.prepend(copyBox);
      }
    }
    if (editing && draft().startsWith('=')) {
      formulaRefs(draft()).forEach((ref, i) => {
        const v = viewRect(expand(ref.rect));
        if (!v) return;
        const box = el('div', 'ref-box');
        const color = REF_COLORS[i % REF_COLORS.length]!;
        box.style.cssText = `left:${v.x - 1}px;top:${v.y - 1}px;width:${v.w + 1}px;height:${v.h + 1}px;border-color:${color};background:${color}14`;
        overlay.prepend(box);
      });
    }
    placeEditor(av);
    const list = validationAt(sheet(), active.row, active.col);
    dvBtn.hidden = !list || Boolean(editing) || !av;
    if (list && av) {
      dvBtn.style.left = `${av.x + av.w - 18}px`;
      dvBtn.style.top = `${av.y + (av.h - 18) / 2}px`;
    }
    reportCurrent();
    const s = selection();
    if (document.activeElement !== namebox) namebox.value = multi ? rangeText(s) : formatA1(active.row, active.col);
    if (!editing && document.activeElement !== finput) finput.value = wb.inputAt(sname(), am.r0, am.c0);
  }

  // ---------- tracing precedents and dependents ----------
  /** The cell whose precedents or dependents are drawn; recomputed on every paint, cleared by an edit or Esc. */
  let trace: { kind: 'pre' | 'dep'; sheet: string; row: number; col: number } | null = null;
  const sheetCase = (n: string) => data.sheets.find((s) => s.name.toLowerCase() === n.toLowerCase())?.name ?? n;
  function traceItems(): { sheet: string; rect: Rect }[] {
    if (!trace) return [];
    if (trace.kind === 'pre') return wb.precedents(trace.sheet, trace.row, trace.col).map((g) => ({ sheet: sheetCase(g.sheet), rect: { r0: g.r0, c0: g.c0, r1: g.r1, c1: g.c1 } }));
    return wb.dependents(trace.sheet, trace.row, trace.col).map((d) => ({ sheet: sheetCase(d.sheet), rect: { r0: d.row, c0: d.col, r1: d.row, c1: d.col } }));
  }
  function paintTrace() {
    if (!trace) return;
    const here = sname();
    const items = traceItems().filter((t) => t.sheet === here);
    const origin = trace.sheet === here ? viewRect({ r0: trace.row, c0: trace.col, r1: trace.row, c1: trace.col }) : null;
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.classList.add('trace-arrows');
    const centre = (v: { x: number; y: number; w: number; h: number }) => [v.x + Math.min(v.w, 60) / 2, v.y + v.h / 2] as const;
    for (const t of items) {
      const v = viewRect(t.rect);
      if (!v) continue;
      const box = el('div', 'trace-box');
      box.style.cssText = `left:${v.x - 1}px;top:${v.y - 1}px;width:${v.w + 1}px;height:${v.h + 1}px`;
      overlay.prepend(box);
      if (!origin) continue;
      // Excel's arrows: from what is read to what reads it, a dot at the tail.
      const [ax, ay] = trace.kind === 'pre' ? centre(v) : centre(origin);
      const [bx, by] = trace.kind === 'pre' ? centre(origin) : centre(v);
      const len = Math.hypot(bx - ax, by - ay);
      if (len < 8) continue;
      const ux = (bx - ax) / len;
      const uy = (by - ay) / len;
      const hx = bx - ux * 7;
      const hy = by - uy * 7;
      svg.insertAdjacentHTML(
        'beforeend',
        `<circle cx="${ax}" cy="${ay}" r="2.5"/><path d="M${ax},${ay}L${bx},${by}M${hx - uy * 4},${hy + ux * 4}L${bx},${by}L${hx + uy * 4},${hy - ux * 4}"/>`,
      );
    }
    if (svg.childNodes.length) overlay.append(svg);
  }
  /** Draws where the active cell's value comes from (precedents) or what uses it (dependents), and lists them to jump to. */
  function traceCells(kind: 'pre' | 'dep') {
    trace = { kind, sheet: sname(), row: active.row, col: active.col };
    const items = traceItems();
    render();
    if (!items.length) {
      trace = null;
      render();
      return toast(kind === 'pre' ? T.noPrecedents : T.noDependents);
    }
    const v = viewRect({ r0: active.row, c0: active.col, r1: active.row, c1: active.col });
    const box = overlay.getBoundingClientRect();
    const at = v ? { x: box.left + v.x + v.w, y: box.top + v.y + v.h } : namebox;
    showMenu(
      items.map((t) => {
        const where = `${t.sheet === sname() ? '' : `${quoteSheet(t.sheet)}!`}${rangeText(t.rect)}`;
        const single = t.rect.r0 === t.rect.r1 && t.rect.c0 === t.rect.c1;
        const input = single ? wb.inputAt(t.sheet, t.rect.r0, t.rect.c0) : '';
        return {
          label: input ? `${where}　${input.length > 40 ? `${input.slice(0, 39)}…` : input}` : where,
          icon: kind === 'pre' ? 'arrowRight' : 'arrowLeft',
          action: () => {
            const i = data.sheets.findIndex((s) => s.name === t.sheet);
            if (i >= 0) switchSheet(i);
            moveTo({ row: t.rect.r0, col: t.rect.c0 });
            moveTo({ row: t.rect.r1, col: t.rect.c1 }, true);
          },
        };
      }),
      at,
      () => editor.focus(),
    );
  }
  function toggleFormulas() {
    showFormulas = !showFormulas;
    grid.classList.toggle('show-formulas', showFormulas);
    render();
    toast(showFormulas ? T.formulasShown : T.formulasHidden);
  }
  /** Goes to the next cell showing an error, through the following sheets and round again, and says what the error means. */
  function nextError() {
    const n = data.sheets.length;
    const live = new Map(wb.toJSON().sheets.map((s) => [s.name, s.cells]));
    const after = (s: number, r: number, c: number) => s === 0 && (r > active.row || (r === active.row && c > active.col));
    for (let k = 0; k <= n; k++) {
      const i = (si + k) % n;
      const sheetName = data.sheets[i]!.name;
      const found = Object.keys(live.get(sheetName) ?? {})
        .map((a) => parseA1(a))
        .filter((p): p is { row: number; col: number } => p !== null && wb.value(sheetName, p.row, p.col) instanceof CellError)
        .sort((a, b) => a.row - b.row || a.col - b.col)
        .find((p) => (k === 0 ? after(0, p.row, p.col) : true));
      if (!found) continue;
      if (i !== si) switchSheet(i);
      moveTo(found);
      const code = (wb.value(sheetName, found.row, found.col) as CellError).code;
      return toast(`${formatA1(found.row, found.col)} ${code}：${T.errorWhy[code as keyof typeof T.errorWhy] ?? ''}`);
    }
    const here = wb.value(sname(), active.row, active.col);
    if (here instanceof CellError) return toast(`${formatA1(active.row, active.col)} ${here.code}：${T.errorWhy[here.code as keyof typeof T.errorWhy] ?? ''}`);
    toast(T.noErrors);
  }
  function clearTrace() {
    trace = null;
    render();
  }

  function placeEditor(av: { x: number; y: number; w: number; h: number } | null) {
    const r = av ?? { x: -9999, y: -9999, w: 100, h: 24 };
    const style = styleAt(active.row, active.col);
    editor.style.left = `${r.x - 1}px`;
    editor.style.top = `${r.y - 1}px`;
    editor.style.height = `${r.h + 1}px`;
    editor.style.minWidth = `${r.w + 1}px`;
    editor.style.width = editing ? `${Math.max(r.w + 1, measure(editor.value, style) + 24)}px` : `${r.w + 1}px`;
    editor.style.fontWeight = style?.b ? '600' : '';
    editor.style.fontSize = style?.size && style.size !== BASE_PT ? `${(ptToPx(style.size) * zoom).toFixed(2)}px` : '';
    editor.style.fontStyle = style?.i ? 'italic' : '';
    editor.style.textAlign = editing ? 'left' : '';
  }

  /** The references in a formula, for colouring them: on this sheet only. */
  function formulaRefs(text: string): { rect: Rect; start: number; end: number }[] {
    const out: { rect: Rect; start: number; end: number }[] = [];
    const re = /("(?:[^"]|"")*")|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)(?![\w(!])/g;
    for (let m = re.exec(text); m; m = re.exec(text)) {
      if (m[1]) continue;
      const before = text[m.index - 1] ?? '';
      if (/[\w.$]/.test(before)) continue;
      const sheetName = m[2]?.slice(0, -1).replace(/^'|'$/g, '').replaceAll("''", "'");
      if (sheetName && sheetName.toLowerCase() !== sname().toLowerCase()) continue;
      const rect = parseRange(m[3]!);
      if (rect) out.push({ rect, start: m.index, end: m.index + m[0].length });
    }
    return out;
  }

  // Tell the dev server what is selected, so agents can resolve "this cell" (node_modules/.open-spreadsheet/current.json).
  let reportTimer: ReturnType<typeof setTimeout> | null = null;
  let lastReport = '';
  function reportCurrent() {
    if (reportTimer) clearTimeout(reportTimer);
    reportTimer = setTimeout(() => {
      const s = selection();
      const v = wb.value(name(), active.row, active.col);
      const payload = {
        workbook: id,
        title: data.title,
        sheet: name(),
        active: formatA1(active.row, active.col),
        selection: rangeText(s),
        input: wb.inputAt(name(), active.row, active.col),
        value: v instanceof CellError ? v.code : v,
        note: noteAt(active.row, active.col),
      };
      const key = JSON.stringify(payload);
      if (key === lastReport) return;
      lastReport = key;
      import.meta.hot?.send('open-spreadsheet:current', payload);
    }, 250);
  }

  // ---------- toolbar state ----------
  function paintToolbar() {
    const s = styleAt(active.row, active.col) ?? {};
    tb.bold.classList.toggle('on', Boolean(s.b));
    tb.size.textContent = String(s.size ?? BASE_PT);
    tb.size.append(icon('chevronDown', 12));
    tb.italic.classList.toggle('on', Boolean(s.i));
    tb.underline.classList.toggle('on', Boolean(s.u));
    tb.strike.classList.toggle('on', Boolean(s.s));
    tb.wrap.classList.toggle('on', Boolean(s.wrap));
    tb.left.classList.toggle('on', s.align === 'left');
    tb.center.classList.toggle('on', s.align === 'center');
    tb.right.classList.toggle('on', s.align === 'right');
    tb.merge.classList.toggle('on', Boolean(mergeOf(active.row, active.col)));
    tb.freeze.classList.toggle('on', Boolean(fr() || fc()));
    tb.filter.classList.toggle('on', Boolean(sheet().filter));
    colorBar.style.background = s.color ?? 'currentColor';
    fillBar.style.background = s.bg ?? 'transparent';
    tb.undo.disabled = undoStack.length === 0;
    tb.redo.disabled = redoStack.length === 0;
    const f = formatOf(active.row, active.col);
    tb.format.replaceChildren(el('span', '', formatLabel(f)), icon('chevronDown', 13));
  }

  const FORMATS = () => [
    { label: T.fmtGeneral, code: '' },
    { label: T.fmtNumber, code: '#,##0.00', sample: '1,234.56' },
    { label: T.fmtTwd, code: '"NT$"#,##0', sample: 'NT$1,235' },
    { label: T.fmtCurrency, code: '$#,##0.00', sample: '$1,234.56' },
    { label: T.fmtPercent, code: '0.00%', sample: '12.34%' },
    { label: T.fmtScientific, code: '0.00E+00', sample: '1.23E+03' },
    { label: T.fmtDate, code: 'yyyy-mm-dd', sample: '2026-11-15' },
    { label: T.fmtLongDate, ...longDate() },
    { label: T.fmtTime, code: 'hh:mm', sample: '14:30' },
    { label: T.fmtDateTime, code: 'yyyy-mm-dd hh:mm', sample: '2026-11-15 14:30' },
    { label: T.fmtText, code: '@', sample: 'abc' },
  ];
  function formatLabel(code?: string) {
    if (!code) return T.fmtGeneral;
    const hit = FORMATS().find((f) => f.code === code);
    if (hit) return hit.label;
    if (isDateFormat(code)) return T.fmtDate;
    if (code.includes('%')) return T.fmtPercent;
    if (/\$|NT/.test(code)) return T.fmtCurrency;
    return T.fmtNumber;
  }

  function paintStats() {
    const s = selection();
    const cells = (s.r1 - s.r0 + 1) * (s.c1 - s.c0 + 1);
    stats.replaceChildren();
    if (cells < 2 || cells > 200_000) return;
    let count = 0;
    let nums = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    const name = sname();
    const used = wb.used(name);
    for (let r = s.r0; r <= Math.min(s.r1, used.rows - 1); r++) {
      for (let c = s.c0; c <= Math.min(s.c1, used.cols - 1); c++) {
        const v = wb.value(name, r, c);
        if (v === null || v === '') continue;
        count++;
        if (typeof v === 'number') {
          nums++;
          sum += v;
          if (v < min) min = v;
          if (v > max) max = v;
        }
      }
    }
    if (!count) return;
    const fmt = (n: number) => formatValue(Number.parseFloat(n.toPrecision(15)), formatOf(active.row, active.col) ?? '#,##0.##').text.replace(/\.$/, '');
    const parts: [string, string][] = nums ? [[T.sum, fmt(sum)], [T.average, fmt(sum / nums)], [T.min, fmt(min)], [T.max, fmt(max)], [T.count, String(count)]] : [[T.count, String(count)]];
    for (const [k, v] of parts) {
      const item = el('span', 'stat');
      item.append(el('span', 'stat-k', k), el('b', '', v));
      stats.append(item);
    }
  }

  // ---------- history and saving ----------
  function serialize(): WorkbookData {
    const cells = new Map(wb.toJSON().sheets.map((s) => [s.name, s.cells]));
    const names = wb.namesJSON();
    const { names: _old, ...rest } = data;
    void _old;
    return {
      ...rest,
      ...(Object.keys(names).length ? { names } : {}),
      sheets: data.sheets.map((s) => {
        const out: SheetData = { name: s.name, cells: cells.get(s.name) ?? {} };
        if (s.cols && Object.keys(s.cols).length) out.cols = s.cols;
        if (s.rows && Object.keys(s.rows).length) out.rows = s.rows;
        if (s.styles && Object.keys(s.styles).length) out.styles = s.styles;
        if (s.merges?.length) out.merges = s.merges;
        if (s.freeze && (s.freeze.rows || s.freeze.cols)) out.freeze = s.freeze;
        if (s.color) out.color = s.color;
        if (s.notes && Object.keys(s.notes).length) out.notes = s.notes;
        if (s.conditional?.length) out.conditional = s.conditional;
        if (s.validation?.length) out.validation = s.validation;
        if (s.filter) out.filter = s.filter;
        if (s.hiddenRows?.length) out.hiddenRows = s.hiddenRows;
        if (s.hiddenCols?.length) out.hiddenCols = s.hiddenCols;
        if (s.charts?.length) out.charts = s.charts;
        return out;
      }),
    };
  }

  const snapshot = (): Snapshot => ({ json: JSON.stringify(serialize()), si, active: { ...active }, anchor: { ...anchor } });
  function checkpoint() {
    trace = null;
    undoStack.push(snapshot());
    if (undoStack.length > 200) undoStack.shift();
    redoStack.length = 0;
  }
  function restore(s: Snapshot) {
    data = JSON.parse(s.json) as WorkbookData;
    wb = Workbook.fromJSON(data);
    si = Math.min(s.si, data.sheets.length - 1);
    active = s.active;
    anchor = s.anchor;
    titleBtn.textContent = data.title;
    layout();
    renderTabs();
    scrollIntoView();
    render();
    paintNotesBadge();
    save();
  }
  function undo() {
    if (editing) cancelEdit();
    const s = undoStack.pop();
    if (!s) return;
    redoStack.push(snapshot());
    restore(s);
  }
  function redo() {
    const s = redoStack.pop();
    if (!s) return;
    undoStack.push(snapshot());
    restore(s);
  }

  let saveTimer: ReturnType<typeof setTimeout> | null = null;
  let lastSaved = JSON.stringify(serialize());
  function save() {
    status.textContent = T.saving;
    status.classList.remove('warn');
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(flush, 400);
  }
  async function flush() {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = null;
    const out = serialize();
    const text = JSON.stringify(out);
    if (text === lastSaved) {
      status.textContent = T.saved;
      return;
    }
    lastSaved = text;
    try {
      const r = await api.put(id, out, etag);
      if (r.status === 412) {
        // Someone else (an agent) wrote the file since this page read it: take theirs, keep ours in undo.
        lastSaved = '';
        await reloadFromDisk(true);
        return;
      }
      if (!r.ok) throw new Error(String(r.status));
      etag = r.headers.get('etag') ?? etag;
      status.textContent = T.saved;
    } catch {
      status.textContent = T.unsaved;
      status.classList.add('warn');
    }
  }
  /** Every change: undo point, the change, then layout, paint and save. */
  function change(fn: () => void, opts: { relayout?: boolean } = {}) {
    checkpoint();
    fn();
    if (opts.relayout !== false) layout();
    render();
    save();
  }

  // ---------- cell edits ----------
  const name = () => sname();
  function setInputs(cells: { row: number; col: number; input: string }[]) {
    wb.setCells(name(), cells);
  }
  function forEachCell(r: Rect, fn: (row: number, col: number) => void) {
    const used = wb.used(name());
    const r1 = Math.min(r.r1, Math.max(used.rows, ...Object.keys(sheet().styles ?? {}).map((a) => parseA1(a)!.row + 1)) - 1);
    const c1 = Math.min(r.c1, Math.max(used.cols, ...Object.keys(sheet().styles ?? {}).map((a) => parseA1(a)!.col + 1)) - 1);
    for (let row = r.r0; row <= r1; row++) for (let col = r.c0; col <= c1; col++) fn(row, col);
  }
  function setStyle(r: Rect, patch: Partial<CellStyle> | null) {
    const sh = sheet();
    sh.styles ??= {};
    for (let row = r.r0; row <= r.r1; row++) {
      for (let col = r.c0; col <= r.c1; col++) {
        const a = formatA1(row, col);
        if (patch === null) {
          delete sh.styles[a];
          continue;
        }
        const next: CellStyle = { ...sh.styles[a], ...patch };
        for (const k of Object.keys(next) as (keyof CellStyle)[]) if (next[k] === undefined || next[k] === false || next[k] === '') delete next[k];
        if (Object.keys(next).length) sh.styles[a] = next;
        else delete sh.styles[a];
      }
    }
  }
  /** Style edits apply to the selection; a whole row or column applies to the used part only. */
  const styleRect = () => {
    const s = selection();
    const used = wb.used(name());
    return { ...s, r1: Math.min(s.r1, Math.max(used.rows + 20, s.r0 + 200)), c1: Math.min(s.c1, Math.max(used.cols + 5, s.c0 + 50)) };
  };
  function toggleStyle(k: 'b' | 'i' | 'u' | 's' | 'wrap') {
    const on = !styleAt(active.row, active.col)?.[k];
    change(() => setStyle(styleRect(), { [k]: on }), { relayout: false });
  }
  // ---------- checkboxes ----------
  const isCheck = (row: number, col: number) => Boolean(styleAt(row, col)?.check);
  /**
   * Flips checkboxes, as Google Sheets does with Space: all on when any in the
   * range is off, else all off. A checkbox whose value comes from a formula is left alone.
   */
  function toggleChecks(r: Rect) {
    const cells: { row: number; col: number }[] = [];
    for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) if (isCheck(row, col) && !wb.inputAt(sname(), row, col).startsWith('=')) cells.push({ row, col });
    if (!cells.length) return false;
    const next = cells.some((c) => wb.value(sname(), c.row, c.col) !== true) ? 'TRUE' : 'FALSE';
    change(() => setInputs(cells.map((c) => ({ ...c, input: next }))), { relayout: false });
    return true;
  }
  function insertCheckboxes() {
    const r = styleRect();
    change(() => {
      setStyle(r, { check: true });
      const blanks: { row: number; col: number; input: string }[] = [];
      for (let row = r.r0; row <= Math.min(r.r1, r.r0 + 9999); row++) for (let col = r.c0; col <= r.c1; col++) if (!wb.inputAt(sname(), row, col)) blanks.push({ row, col, input: 'FALSE' });
      setInputs(blanks);
    });
  }
  function removeCheckboxes() {
    change(() => setStyle(styleRect(), { check: undefined }), { relayout: false });
  }

  function setValign(v: 'top' | 'middle' | 'bottom') {
    const cur = styleAt(active.row, active.col)?.valign;
    change(() => setStyle(styleRect(), { valign: cur === v ? undefined : v }), { relayout: false });
  }
  function setAlign(a: 'left' | 'center' | 'right') {
    const cur = styleAt(active.row, active.col)?.align;
    change(() => setStyle(styleRect(), { align: cur === a ? undefined : a }), { relayout: false });
  }
  function setSize(pt: number) {
    const size = Math.max(6, Math.min(72, Math.round(pt * 2) / 2));
    // Rows grow to fit, so the layout is redone.
    change(() => setStyle(styleRect(), { size: size === BASE_PT ? undefined : size }));
  }
  /** ⇧⌘> / ⇧⌘<: the next size in the list, as Excel's Increase / Decrease Font Size. */
  function stepSize(dir: 1 | -1) {
    const cur = styleAt(active.row, active.col)?.size ?? BASE_PT;
    const next = dir > 0 ? FONT_SIZES.find((n) => n > cur) : [...FONT_SIZES].reverse().find((n) => n < cur);
    if (next) setSize(next);
  }
  function setFormat(code: string) {
    change(() => setStyle(styleRect(), { fmt: code || undefined }), { relayout: false });
  }
  function bumpDecimals(d: number) {
    let f = formatOf(active.row, active.col);
    if (!f) {
      const v = wb.value(name(), active.row, active.col);
      const decimals = typeof v === 'number' ? Math.min(10, (formatValue(v, '').text.split('.')[1] ?? '').length) : 0;
      f = decimals ? `0.${'0'.repeat(decimals)}` : '0';
    }
    const next = f.replace(/(0|#)(\.(0+))?(?=[^0#.]*$|%|E|"|;)/, (m, digit: string, _dot: string, zeros?: string) => {
      const n = Math.max(0, (zeros?.length ?? 0) + d);
      return n ? `${digit}.${'0'.repeat(n)}` : digit;
    });
    setFormat(next === f && d > 0 ? `${f}.0` : next);
  }
  function setColor(kind: 'color' | 'bg', c: string | null) {
    change(() => setStyle(styleRect(), { [kind]: c ?? undefined }), { relayout: false });
  }

  function clearContents(r: Rect) {
    const cells: { row: number; col: number; input: string }[] = [];
    forEachCell(r, (row, col) => {
      if (wb.inputAt(name(), row, col)) cells.push({ row, col, input: '' });
    });
    if (cells.length) change(() => setInputs(cells));
  }

  // ---------- editing ----------
  const draft = () => (editing?.source === 'bar' ? finput.value : editor.value);
  function beginEdit(mode: 'enter' | 'edit', initial?: string, source: 'cell' | 'bar' = 'cell') {
    const m = mergeOf(active.row, active.col);
    if (m) {
      active = { row: m.r0, col: m.c0 };
    }
    editing = { mode, source };
    pointer = null;
    const text = initial ?? wb.inputAt(name(), active.row, active.col);
    editor.value = text;
    finput.value = text;
    editor.classList.remove('idle');
    fillHandle.hidden = true;
    if (source === 'cell') {
      editor.focus();
      editor.setSelectionRange(text.length, text.length);
    }
    paintOverlay();
    updatePopup();
  }
  function commitEdit(move?: [number, number]) {
    if (!editing) return;
    let value = draft();
    const row = active.row;
    const col = active.col;
    endEdit();
    // A cell formatted as text keeps what was typed as text.
    if (styleAt(row, col)?.fmt === '@' && value && !value.startsWith("'") && !value.startsWith('=')) value = `'${value}`;
    // Close the parentheses a formula left open, as Excel does.
    if (value.startsWith('=')) {
      const open = (value.match(/\(/g) ?? []).length - (value.match(/\)/g) ?? []).length;
      if (open > 0 && (value.match(/"/g) ?? []).length % 2 === 0) value += ')'.repeat(open);
    }
    const allowed = validationAt(sheet(), row, col);
    if (allowed && value !== '' && !allowed.some((v) => v.toLowerCase() === value.trim().toLowerCase())) {
      toast(T.notInList(allowed.join('、')));
      render();
      editor.focus();
      return;
    }
    if (value !== wb.inputAt(name(), row, col)) change(() => setInputs([{ row, col, input: value }]));
    else render();
    if (move) moveActive(move[0], move[1]);
    editor.focus();
  }
  function cancelEdit() {
    endEdit();
    render();
    editor.focus();
  }
  function endEdit() {
    editing = null;
    pointer = null;
    editor.value = '';
    editor.classList.add('idle');
    popup.hidden = true;
  }
  /** Whether a reference can go at the caret: right after an operator, a comma or an open parenthesis. */
  function canPoint() {
    if (!editing || editing.mode !== 'enter' && !pointer) return false;
    const input = editing.source === 'bar' ? finput : editor;
    const text = input.value;
    if (!text.startsWith('=')) return false;
    const caret = input.selectionStart ?? text.length;
    if (pointer && caret === pointer.end) return true;
    const before = text.slice(0, caret).trimEnd();
    return /[=(,+\-*/^&<>:;%]$/.test(before);
  }
  function pointAt(a: Pos, b: Pos) {
    const input = editing!.source === 'bar' ? finput : editor;
    const text = input.value;
    const ref = rangeText(expand(rectOf(a, b)));
    if (!pointer) {
      const caret = input.selectionStart ?? text.length;
      pointer = { start: caret, end: caret, anchor: a, at: b };
    }
    const next = text.slice(0, pointer.start) + ref + text.slice(pointer.end);
    pointer = { ...pointer, end: pointer.start + ref.length, anchor: a, at: b };
    input.value = next;
    (input === editor ? finput : editor).value = next;
    input.setSelectionRange(pointer.end, pointer.end);
    paintOverlay();
  }

  // Autocomplete and argument hints.
  let acItems: { name: string; signature: string; description: string; isName?: boolean }[] = [];
  let acIndex = 0;
  function updatePopup() {
    const input = editing?.source === 'bar' ? finput : editor;
    const text = input.value;
    if (!editing || !text.startsWith('=')) {
      popup.hidden = true;
      return;
    }
    const caret = input.selectionStart ?? text.length;
    const before = text.slice(0, caret);
    const m = /(?:^=|[=(,+\-*/^&<>:;\s])([A-Za-z\u00C0-\uFFFF_][\w.\u00C0-\uFFFF]*)$/.exec(before);
    acItems = m && !/^[A-Za-z]{1,3}\d+$/.test(m[1]!) ? suggest(m[1]!) : [];
    if (m && !/^[A-Za-z]{1,3}\d+$/.test(m[1]!)) {
      const p = m[1]!.toLowerCase();
      const named = Object.entries(wb.namesJSON())
        .filter(([n]) => n.toLowerCase().startsWith(p))
        .map(([n, ref]) => ({ name: n, signature: ref, description: T.nameLabel, isName: true }));
      acItems = [...named, ...acItems].slice(0, 8);
    }
    popup.replaceChildren();
    if (acItems.length) {
      acIndex = Math.min(acIndex, acItems.length - 1);
      acItems.forEach((it, i) => {
        const row = el('div', `ac-item${i === acIndex ? ' on' : ''}`);
        row.append(el('b', '', it.name), el('span', '', it.description));
        row.addEventListener('mousedown', (e) => {
          e.preventDefault();
          acIndex = i;
          acceptSuggestion();
        });
        popup.append(row);
      });
      const cur = acItems[acIndex]!;
      popup.append(el('div', 'ac-sig', cur.signature));
    } else {
      // Inside a function's parentheses: show its signature.
      let depth = 0;
      let fn: string | null = null;
      for (let i = before.length - 1; i >= 0; i--) {
        const ch = before[i];
        if (ch === ')') depth++;
        else if (ch === '(') {
          if (depth === 0) {
            fn = /([A-Za-z][A-Za-z0-9.]*)$/.exec(before.slice(0, i))?.[1] ?? null;
            break;
          }
          depth--;
        }
      }
      const h = fn ? helpFor(fn) : null;
      if (!h) {
        popup.hidden = true;
        return;
      }
      const sig = el('div', 'ac-sig solo');
      sig.append(el('span', '', h.signature));
      if (h.description) sig.append(el('span', 'ac-desc', h.description));
      popup.append(sig);
    }
    const r = (editing.source === 'bar' ? finput : editor).getBoundingClientRect();
    popup.style.left = `${Math.min(r.left, innerWidth - 380)}px`;
    popup.style.top = `${r.bottom + 4}px`;
    popup.hidden = false;
  }
  function acceptSuggestion() {
    const item = acItems[acIndex];
    if (!item) return;
    const input = editing?.source === 'bar' ? finput : editor;
    const caret = input.selectionStart ?? input.value.length;
    const before = input.value.slice(0, caret).replace(/[A-Za-z\u00C0-\uFFFF_][\w.\u00C0-\uFFFF]*$/, (item as { isName?: boolean }).isName ? item.name : `${item.name}(`);
    const next = before + input.value.slice(caret);
    input.value = next;
    (input === editor ? finput : editor).value = next;
    input.setSelectionRange(before.length, before.length);
    pointer = null;
    paintOverlay();
    updatePopup();
  }
  /** F4: cycles the reference at the caret through A1, $A$1, A$1, $A1. */
  function toggleAbsolute() {
    const input = editing?.source === 'bar' ? finput : editor;
    const text = input.value;
    const caret = input.selectionStart ?? text.length;
    const re = /\$?[A-Za-z]{1,3}\$?\d+/g;
    for (let m = re.exec(text); m; m = re.exec(text)) {
      if (caret < m.index || caret > m.index + m[0].length) continue;
      const plain = m[0].replaceAll('$', '');
      const [, col, row] = /([A-Za-z]+)(\d+)/.exec(plain)!;
      const states = [`${col}${row}`, `$${col}$${row}`, `${col}$${row}`, `$${col}${row}`];
      const next = states[(states.indexOf(m[0]) + 1) % 4]!;
      const out = text.slice(0, m.index) + next + text.slice(m.index + m[0].length);
      input.value = out;
      (input === editor ? finput : editor).value = out;
      input.setSelectionRange(m.index + next.length, m.index + next.length);
      paintOverlay();
      return;
    }
  }

  // ---------- movement ----------
  function scrollIntoView(p: Pos = active) {
    const m = mergeOf(p.row, p.col) ?? { r0: p.row, c0: p.col, r1: p.row, c1: p.col };
    const cw = body.clientWidth;
    const ch = body.clientHeight;
    if (m.r0 >= fr()) {
      const top = rowY[m.r0]! - frozenH();
      const bottom = rowY[m.r1 + 1]! - frozenH();
      if (top < body.scrollTop) body.scrollTop = top;
      else if (bottom > body.scrollTop + ch) body.scrollTop = bottom - ch;
    }
    if (m.c0 >= fc()) {
      const left = colX[m.c0]! - frozenW();
      const right = colX[m.c1 + 1]! - frozenW();
      if (left < body.scrollLeft) body.scrollLeft = left;
      else if (right > body.scrollLeft + cw) body.scrollLeft = Math.max(left, right - cw);
    }
  }
  /** Moves the active cell there; with `extend`, moves only the selection's far corner. */
  function moveTo(p: Pos, extend = false) {
    const q = { row: Math.max(0, Math.min(MAX_ROWS - 1, p.row)), col: Math.max(0, Math.min(MAX_COLS - 1, p.col)) };
    if (extend) anchor = q;
    else {
      active = q;
      anchor = { ...q };
    }
    if (q.row >= nRows - 10 || q.col >= nCols - 3) layout();
    scrollIntoView(q);
    render();
  }
  function moveActive(dr: number, dc: number, extend = false) {
    if (extend) return moveTo({ row: anchor.row + dr, col: anchor.col + dc }, true);
    // Out of a merge from its edge; into a merge at its anchor.
    const m = mergeOf(active.row, active.col);
    let row = active.row;
    let col = active.col;
    if (m && !extend) {
      if (dr > 0) row = m.r1;
      if (dc > 0) col = m.c1;
      if (dr < 0) row = m.r0;
      if (dc < 0) col = m.c0;
    }
    row += dr;
    col += dc;
    // Rows a filter hides are stepped over.
    while (dr !== 0 && hiddenRows.has(row) && row > 0) row += dr;
    while (dc !== 0 && hiddenColSet.has(col) && col > 0) col += dc;
    const n = mergeOf(row, col);
    if (n && !extend) {
      row = n.r0;
      col = n.c0;
    }
    moveTo({ row, col }, extend);
  }
  /** ⌘ + arrow: to the edge of the data block, as in Excel. */
  function jump(dr: number, dc: number, extend: boolean) {
    const filled = (r: number, c: number) => wb.inputAt(name(), r, c) !== '';
    const used = wb.used(name());
    const from = extend ? anchor : active;
    const maxR = Math.max(used.rows, from.row + 1) - 1;
    const maxC = Math.max(used.cols, from.col + 1) - 1;
    let r = from.row;
    let c = from.col;
    const inside = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr <= maxR && cc <= maxC;
    if (filled(r, c) && inside(r + dr, c + dc) && filled(r + dr, c + dc)) {
      while (inside(r + dr, c + dc) && filled(r + dr, c + dc)) {
        r += dr;
        c += dc;
      }
    } else {
      r += dr;
      c += dc;
      while (inside(r, c) && !filled(r, c)) {
        r += dr;
        c += dc;
      }
      if (!inside(r, c)) {
        r = dr > 0 ? maxR : dr < 0 ? 0 : r - dr;
        c = dc > 0 ? maxC : dc < 0 ? 0 : c - dc;
      }
    }
    moveTo({ row: Math.max(0, r), col: Math.max(0, c) }, extend);
  }
  /** Enter and Tab walk inside a selected range instead of leaving it. */
  function stepInSelection(dr: number, dc: number) {
    const s = selection();
    if (s.r0 === s.r1 && s.c0 === s.c1) return moveActive(dr, dc);
    let { row, col } = active;
    if (dr) {
      row += dr;
      if (row > s.r1) [row, col] = [s.r0, col + 1 > s.c1 ? s.c0 : col + 1];
      if (row < s.r0) [row, col] = [s.r1, col - 1 < s.c0 ? s.c1 : col - 1];
    } else {
      col += dc;
      if (col > s.c1) [row, col] = [row + 1 > s.r1 ? s.r0 : row + 1, s.c0];
      if (col < s.c0) [row, col] = [row - 1 < s.r0 ? s.r1 : row - 1, s.c1];
    }
    active = { row, col };
    scrollIntoView();
    render();
  }

  // ---------- clipboard ----------
  function copySelection(cut: boolean): string {
    const s = selection();
    const cells: Clip['cells'] = [];
    const lines: string[] = [];
    for (let r = s.r0; r <= s.r1; r++) {
      const row: Clip['cells'][number] = [];
      const texts: string[] = [];
      for (let c = s.c0; c <= s.c1; c++) {
        row.push({ input: wb.inputAt(name(), r, c), style: styleAt(r, c) });
        const t = show(r, c).text;
        texts.push(/[\t\n"]/.test(t) ? `"${t.replaceAll('"', '""')}"` : t);
      }
      cells.push(row);
      lines.push(texts.join('\t'));
    }
    const text = lines.join('\n');
    clip = { text, cells, from: s, sheet: name(), cut };
    paintOverlay();
    return text;
  }
  function pasteText(text: string, valuesOnly = false) {
    const s = selection();
    if (clip && clip.text === text.replace(/\r\n/g, '\n').replace(/\n$/, '')) return pasteClip(clip, valuesOnly);
    const rows = parseCsv(text.includes('\t') ? text : text.replace(/,/g, '\u0000')).map((r) => r.map((v) => v.replaceAll('\u0000', ',')));
    if (!text.includes('\t')) {
      // One column: split lines only.
      rows.splice(0, rows.length, ...text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => [l]));
    }
    const cells = rows.flatMap((line, i) => line.map((v, k) => ({ row: s.r0 + i, col: s.c0 + k, input: v })));
    if (cells.length) change(() => setInputs(cells));
    anchor = { row: s.r0, col: s.c0 };
    active = { row: s.r0 + rows.length - 1, col: s.c0 + Math.max(...rows.map((r) => r.length)) - 1 };
    render();
  }
  function pasteClip(c: Clip, valuesOnly: boolean) {
    const s = selection();
    const h = c.cells.length;
    const w = c.cells[0]?.length ?? 0;
    // A selection that is a whole multiple of the copied block gets it tiled, as in Excel.
    const reps = (s.r1 - s.r0 + 1) % h === 0 && (s.c1 - s.c0 + 1) % w === 0 ? [(s.r1 - s.r0 + 1) / h, (s.c1 - s.c0 + 1) / w] : [1, 1];
    change(() => {
      const cells: { row: number; col: number; input: string }[] = [];
      const srcSheet = c.sheet;
      if (c.cut) {
        for (let r = c.from.r0; r <= c.from.r1; r++) {
          for (let k = c.from.c0; k <= c.from.c1; k++) {
            if (srcSheet === name()) cells.push({ row: r, col: k, input: '' });
            else wb.setCell(srcSheet, r, k, '');
          }
        }
        const src = data.sheets.find((x) => x.name === srcSheet);
        if (src?.styles) for (let r = c.from.r0; r <= c.from.r1; r++) for (let k = c.from.c0; k <= c.from.c1; k++) delete src.styles[formatA1(r, k)];
      }
      for (let a = 0; a < reps[0]!; a++) {
        for (let b = 0; b < reps[1]!; b++) {
          c.cells.forEach((line, i) =>
            line.forEach((cell, k) => {
              const row = s.r0 + a * h + i;
              const col = s.c0 + b * w + k;
              let input = cell.input;
              if (valuesOnly) {
                const v = clipValue(c, i, k);
                input = v;
              } else if (!c.cut) input = offsetFormula(input, row - (c.from.r0 + i), col - (c.from.c0 + k));
              const idx = cells.findIndex((x) => x.row === row && x.col === col);
              if (idx >= 0) cells.splice(idx, 1);
              cells.push({ row, col, input });
              if (!valuesOnly) {
                setStyle({ r0: row, c0: col, r1: row, c1: col }, null);
                if (cell.style) setStyle({ r0: row, c0: col, r1: row, c1: col }, cell.style);
              }
            }),
          );
        }
      }
      setInputs(cells);
    });
    anchor = { row: s.r0, col: s.c0 };
    active = { row: s.r0 + h * reps[0]! - 1, col: s.c0 + w * reps[1]! - 1 };
    if (c.cut) clip = null;
    render();
  }
  function clipValue(c: Clip, i: number, k: number): string {
    const src = data.sheets.find((x) => x.name === c.sheet);
    if (!src) return c.cells[i]![k]!.input;
    const v = wb.value(c.sheet, c.from.r0 + i, c.from.c0 + k);
    if (v === null) return '';
    if (v instanceof CellError) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return /^[\d=+\-.$]/.test(v) ? `'${v}` : v;
    return String(v);
  }
  async function pasteFromMenu(valuesOnly: boolean) {
    try {
      const text = await navigator.clipboard.readText();
      pasteText(text, valuesOnly);
    } catch {
      if (clip) pasteClip(clip, valuesOnly);
      else toast(T.pasteBlocked);
    }
  }
  async function copyFromMenu(cut: boolean) {
    const text = copySelection(cut);
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      // The internal clipboard still works inside the page.
    }
  }

  // ---------- fill ----------
  /** Fills `target` from `source` (which it extends): series for numbers and dates, shifted formulas, repeated text. */
  function fill(source: Rect, target: Rect) {
    const vertical = target.c0 === source.c0 && target.c1 === source.c1;
    const cells: { row: number; col: number; input: string }[] = [];
    const n = vertical ? source.r1 - source.r0 + 1 : source.c1 - source.c0 + 1;
    const lines = vertical ? source.c1 - source.c0 + 1 : source.r1 - source.r0 + 1;
    const forward = vertical ? target.r1 > source.r1 : target.c1 > source.c1;
    for (let line = 0; line < lines; line++) {
      const src: { row: number; col: number; input: string }[] = [];
      for (let i = 0; i < n; i++) {
        const row = vertical ? source.r0 + i : source.r0 + line;
        const col = vertical ? source.c0 + line : source.c0 + i;
        src.push({ row, col, input: wb.inputAt(name(), row, col) });
      }
      const nums = src.map((s) => (s.input && !s.input.startsWith('=') ? wb.value(name(), s.row, s.col) : null));
      const allNum = nums.every((v) => typeof v === 'number');
      const step = allNum && n >= 2 ? ((nums[n - 1] as number) - (nums[0] as number)) / (n - 1) : allNum && n === 1 && isDateFormat(formatOf(src[0]!.row, src[0]!.col)) ? 1 : null;
      const textNum = n === 1 ? /^(.*?)(\d+)$/.exec(src[0]!.input) : null;
      const count = vertical ? target.r1 - target.r0 + 1 - n : target.c1 - target.c0 + 1 - n;
      for (let k = 1; k <= count; k++) {
        const idx = forward ? n - 1 + k : -k;
        const row = vertical ? (forward ? source.r1 + k : source.r0 - k) : source.r0 + line;
        const col = vertical ? source.c0 + line : forward ? source.c1 + k : source.c0 - k;
        const from = src[((idx % n) + n) % n]!;
        let input: string;
        if (step !== null) {
          const v = (nums[0] as number) + step * (forward ? n - 1 + k : -k);
          input = isDateFormat(formatOf(src[0]!.row, src[0]!.col)) && !impliedFormat(from.input) ? String(v) : formatLike(from.input, v);
        } else if (textNum && !src[0]!.input.startsWith('=') && textNum[1] && !/^\d/.test(src[0]!.input)) {
          input = `${textNum[1]}${Number(textNum[2]) + (forward ? k : -k)}`;
        } else input = offsetFormula(from.input, row - from.row, col - from.col);
        cells.push({ row, col, input });
        const st = styleAt(from.row, from.col);
        setStyle({ r0: row, c0: col, r1: row, c1: col }, null);
        if (st) setStyle({ r0: row, c0: col, r1: row, c1: col }, st);
      }
    }
    setInputs(cells);
  }
  /** A number typed the way the source was: a date stays a date, a percent a percent. */
  function formatLike(input: string, v: number): string {
    const f = impliedFormat(input);
    if (f && isDateFormat(f)) return formatValue(v, f).text;
    if (f?.includes('%')) return `${Number.parseFloat((v * 100).toPrecision(12))}%`;
    return String(Number.parseFloat(v.toPrecision(15)));
  }

  // ---------- structure ----------
  function shiftMeta(axis: 'row' | 'col', at: number, count: number) {
    const sh = sheet();
    const moveIdx = (i: number): number | null => (i < at ? i : count < 0 && i < at - count ? null : i + count);
    if (sh.styles) {
      const next: Record<string, CellStyle> = {};
      for (const [a, s] of Object.entries(sh.styles)) {
        const p = parseA1(a);
        if (!p) continue;
        const m = moveIdx(axis === 'row' ? p.row : p.col);
        if (m === null) continue;
        next[axis === 'row' ? formatA1(m, p.col) : formatA1(p.row, m)] = s;
      }
      sh.styles = next;
    }
    if (sh.notes) {
      const next: Record<string, string> = {};
      for (const [a, text] of Object.entries(sh.notes)) {
        const p = parseA1(a);
        if (!p) continue;
        const m = moveIdx(axis === 'row' ? p.row : p.col);
        if (m !== null) next[axis === 'row' ? formatA1(m, p.col) : formatA1(p.row, m)] = text;
      }
      sh.notes = next;
    }
    if (axis === 'row' && sh.rows) {
      const next: Record<string, number> = {};
      for (const [k, v] of Object.entries(sh.rows)) {
        const m = moveIdx(Number(k) - 1);
        if (m !== null) next[String(m + 1)] = v;
      }
      sh.rows = next;
    }
    if (axis === 'col' && sh.cols) {
      const next: Record<string, number> = {};
      for (const [k, v] of Object.entries(sh.cols)) {
        const m = moveIdx(columnIndex(k));
        if (m !== null) next[columnName(m)] = v;
      }
      sh.cols = next;
    }
    if (axis === 'row' && sh.hiddenRows) sh.hiddenRows = sh.hiddenRows.flatMap((n) => {
      const m = moveIdx(n - 1);
      return m === null ? [] : [m + 1];
    });
    if (axis === 'col' && sh.hiddenCols) sh.hiddenCols = sh.hiddenCols.flatMap((l) => {
      const m = moveIdx(columnIndex(l));
      return m === null ? [] : [columnName(m)];
    });
    if (sh.charts) sh.charts = sh.charts.flatMap((ch) => {
      const range = shiftRange(ch.range, axis, at, count);
      return range ? [{ ...ch, range }] : [];
    });
    if (sh.conditional) sh.conditional = sh.conditional.flatMap((r) => {
      const range = shiftRange(r.range, axis, at, count);
      return range ? [{ ...r, range }] : [];
    });
    if (sh.validation) sh.validation = sh.validation.flatMap((v) => {
      const range = shiftRange(v.range, axis, at, count);
      return range ? [{ ...v, range }] : [];
    });
    if (sh.filter) {
      const range = shiftRange(sh.filter.range, axis, at, count);
      if (!range) sh.filter = undefined;
      else {
        const hidden: Record<string, string[]> = {};
        for (const [letter, values] of Object.entries(sh.filter.hidden ?? {})) {
          const c = columnIndex(letter);
          const m = axis === 'col' ? (c < at ? c : count < 0 && c < at - count ? null : c + count) : c;
          if (m !== null) hidden[columnName(m)] = values;
        }
        sh.filter = { range, ...(Object.keys(hidden).length ? { hidden } : {}) };
      }
    }
    if (sh.merges) {
      sh.merges = sh.merges
        .map(parseRange)
        .filter((m): m is Rect => m !== null)
        .map((m) => {
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
          const r = axis === 'row' ? { ...m, r0: a, r1: b } : { ...m, c0: a, c1: b };
          return r.r0 === r.r1 && r.c0 === r.c1 ? null : rangeText(r);
        })
        .filter((x): x is string => x !== null);
    }
  }
  function insertRows(at: number, count: number) {
    change(() => {
      wb.insertRows(name(), at, count);
      shiftMeta('row', at, count);
    });
  }
  function insertCols(at: number, count: number) {
    change(() => {
      wb.insertColumns(name(), at, count);
      shiftMeta('col', at, count);
    });
  }
  function deleteRows(r0: number, r1: number) {
    change(() => {
      wb.insertRows(name(), r0, -(r1 - r0 + 1));
      shiftMeta('row', r0, -(r1 - r0 + 1));
    });
    moveTo({ row: r0, col: active.col });
  }
  function deleteCols(c0: number, c1: number) {
    change(() => {
      wb.insertColumns(name(), c0, -(c1 - c0 + 1));
      shiftMeta('col', c0, -(c1 - c0 + 1));
    });
    moveTo({ row: active.row, col: c0 });
  }

  function toggleMerge() {
    const s = selection();
    const existing = merges.filter((m) => overlaps(m, s));
    if (existing.length) {
      change(() => {
        sheet().merges = (sheet().merges ?? []).filter((m) => !existing.some((e) => rangeText(e) === m));
      });
      return;
    }
    if (s.r0 === s.r1 && s.c0 === s.c1) return;
    let lost = false;
    forEachCell(s, (r, c) => {
      if ((r !== s.r0 || c !== s.c0) && wb.inputAt(name(), r, c)) lost = true;
    });
    change(() => {
      const cells: { row: number; col: number; input: string }[] = [];
      forEachCell(s, (r, c) => {
        if ((r !== s.r0 || c !== s.c0) && wb.inputAt(name(), r, c)) cells.push({ row: r, col: c, input: '' });
      });
      setInputs(cells);
      sheet().merges = [...(sheet().merges ?? []), rangeText(s)];
      if (!styleAt(s.r0, s.c0)?.align) setStyle({ r0: s.r0, c0: s.c0, r1: s.r0, c1: s.c0 }, { align: 'center' });
    });
    active = { row: s.r0, col: s.c0 };
    anchor = { ...active };
    if (lost) toast(T.mergeLoses);
    render();
  }

  function setFreeze(rows: number, cols: number) {
    change(() => {
      sheet().freeze = rows || cols ? { rows, cols } : undefined;
    });
    body.scrollTop = 0;
    body.scrollLeft = 0;
    render();
  }

  /** The block of data around a cell, as Excel's current region. */
  function currentRegion(p: Pos): Rect {
    const r = { r0: p.row, c0: p.col, r1: p.row, c1: p.col };
    const filled = (row: number, col: number) => row >= 0 && col >= 0 && wb.inputAt(name(), row, col) !== '';
    for (let grew = true; grew; ) {
      grew = false;
      const rowHas = (row: number) => {
        for (let c = r.c0 - 1; c <= r.c1 + 1; c++) if (filled(row, c)) return true;
        return false;
      };
      const colHas = (col: number) => {
        for (let row = r.r0 - 1; row <= r.r1 + 1; row++) if (filled(row, col)) return true;
        return false;
      };
      if (r.r0 > 0 && rowHas(r.r0 - 1)) (r.r0--, (grew = true));
      if (rowHas(r.r1 + 1)) (r.r1++, (grew = true));
      if (r.c0 > 0 && colHas(r.c0 - 1)) (r.c0--, (grew = true));
      if (colHas(r.c1 + 1)) (r.c1++, (grew = true));
    }
    return r;
  }

  /** The range a sort works on, and whether its first row is a header that stays put. */
  function sortRange(): { s: Rect; header: boolean } {
    let s = selection();
    let header = false;
    if (s.r0 === s.r1) {
      s = currentRegion(active);
      // A first row of text over data is a header and stays put.
      const firstText = Array.from({ length: s.c1 - s.c0 + 1 }, (_, i) => wb.value(name(), s.r0, s.c0 + i)).every((v) => typeof v === 'string' || v === null);
      const keyBelow = wb.value(name(), s.r0 + 1, active.col);
      header = firstText && s.r1 > s.r0 && (typeof keyBelow === 'number' || typeof wb.value(name(), s.r0, active.col) === 'string');
    }
    return { s, header };
  }
  function sortBy(desc: boolean) {
    const { s, header } = sortRange();
    sortRows(s, header, [{ col: active.col, desc }]);
  }
  /** Excel's Custom Sort: up to three columns, each ascending or descending, the header kept or not. */
  async function customSort() {
    const { s, header } = sortRange();
    if (s.r1 <= s.r0 && !header) return;
    const cols = Array.from({ length: s.c1 - s.c0 + 1 }, (_, i) => s.c0 + i);
    const label = (c: number, withHeader: boolean) => {
      const h = withHeader ? wb.value(name(), s.r0, c) : null;
      return typeof h === 'string' && h ? `${columnName(c)}：${h}` : `${T.column} ${columnName(c)}`;
    };
    const options = (none: boolean) => [
      ...(none ? [{ value: '', label: T.sortNone }] : []),
      ...cols.flatMap((c) => [
        { value: `${c}:a`, label: `${label(c, header)}　↑ ${T.sortAscShort}` },
        { value: `${c}:d`, label: `${label(c, header)}　↓ ${T.sortDescShort}` },
      ]),
    ];
    const v = await promptDialog({
      title: T.customSort.replace('…', ''),
      body: `${rangeText(s)}`,
      fields: [
        { name: 'k1', label: T.sortBy, options: options(false), value: `${Math.min(Math.max(active.col, s.c0), s.c1)}:a` },
        { name: 'k2', label: T.thenBy, options: options(true), value: '' },
        { name: 'k3', label: T.thenBy, options: options(true), value: '' },
        { name: 'h', label: T.sortHeader, options: [{ value: '1', label: T.yes }, { value: '0', label: T.no }], value: header ? '1' : '0' },
      ],
      ok: T.sortOk,
    });
    editor.focus();
    if (!v) return;
    const keys = [v.k1, v.k2, v.k3]
      .filter((k): k is string => Boolean(k))
      .map((k) => ({ col: Number(k.split(':')[0]), desc: k.endsWith(':d') }))
      .filter((k, i, all) => all.findIndex((o) => o.col === k.col) === i);
    sortRows(s, v.h === '1', keys);
  }
  function sortRows(s: Rect, header: boolean, keys: { col: number; desc: boolean }[]) {
    const r0 = header ? s.r0 + 1 : s.r0;
    if (s.r1 <= r0 || !keys.length) return;
    const rank = (v: Scalar) => (v === null || v === '' ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
    const rows = Array.from({ length: s.r1 - r0 + 1 }, (_, i) => r0 + i);
    const one = (a: number, b: number, key: number, desc: boolean) => {
      const x = wb.value(name(), a, key);
      const y = wb.value(name(), b, key);
      const rx = rank(x);
      const ry = rank(y);
      // Blanks go last whichever way, as in Excel.
      if (rx === 3 || ry === 3) return rx - ry;
      if (rx !== ry) return (rx - ry) * (desc ? -1 : 1);
      const cmp = typeof x === 'number' ? (x as number) - (y as number) : String(x).localeCompare(String(y), locale(), { numeric: true });
      return desc ? -cmp : cmp;
    };
    // Array.prototype.sort is stable, so rows equal on every key keep their order.
    const sorted = [...rows].sort((a, b) => {
      for (const k of keys) {
        const d = one(a, b, k.col, k.desc);
        if (d) return d;
      }
      return 0;
    });
    change(() => {
      const cells: { row: number; col: number; input: string }[] = [];
      const styles: (CellStyle | undefined)[][] = sorted.map((from) => Array.from({ length: s.c1 - s.c0 + 1 }, (_, k) => styleAt(from, s.c0 + k)));
      sorted.forEach((from, i) => {
        const to = r0 + i;
        for (let c = s.c0; c <= s.c1; c++) cells.push({ row: to, col: c, input: offsetFormula(wb.inputAt(name(), from, c), to - from, 0) });
      });
      setInputs(cells);
      sorted.forEach((_, i) =>
        styles[i]!.forEach((st, k) => {
          const rr = { r0: r0 + i, c0: s.c0 + k, r1: r0 + i, c1: s.c0 + k };
          setStyle(rr, null);
          if (st) setStyle(rr, st);
        }),
      );
    });
  }

  function autosum(fn: string) {
    const s = selection();
    const numeric = (r: number, c: number) => typeof wb.value(name(), r, c) === 'number';
    if (s.r0 === s.r1 && s.c0 === s.c1) {
      // The run of numbers above, else to the left.
      let r = active.row - 1;
      while (r >= 0 && numeric(r, active.col)) r--;
      if (r < active.row - 1) return beginEdit('enter', `=${fn}(${rangeText({ r0: r + 1, c0: active.col, r1: active.row - 1, c1: active.col })})`);
      let c = active.col - 1;
      while (c >= 0 && numeric(active.row, c)) c--;
      if (c < active.col - 1) return beginEdit('enter', `=${fn}(${rangeText({ r0: active.row, c0: c + 1, r1: active.row, c1: active.col - 1 })})`);
      return beginEdit('enter', `=${fn}(`);
    }
    // A range: a total under each column.
    change(() => {
      const cells = [];
      for (let c = s.c0; c <= s.c1; c++) cells.push({ row: s.r1 + 1, col: c, input: `=${fn}(${rangeText({ r0: s.r0, c0: c, r1: s.r1, c1: c })})` });
      setInputs(cells);
    });
  }

  // ---------- sizes ----------
  function setColWidth(cols: number[], px: number | null) {
    const sh = sheet();
    sh.cols ??= {};
    for (const c of cols) {
      if (px === null || Math.round(px) === DEFAULT_COL) delete sh.cols[columnName(c)];
      else sh.cols[columnName(c)] = Math.max(4, Math.round(px));
    }
  }
  function setRowHeight(rows: number[], px: number | null) {
    const sh = sheet();
    sh.rows ??= {};
    for (const r of rows) {
      if (px === null || Math.round(px) === DEFAULT_ROW) delete sh.rows[String(r + 1)];
      else sh.rows[String(r + 1)] = Math.max(4, Math.round(px));
    }
  }
  function fitCol(c: number) {
    const used = wb.used(name());
    let w = 30;
    for (let r = 0; r < Math.min(used.rows, 5000); r++) {
      if (mergeOf(r, c)) continue;
      const t = wb.inputAt(name(), r, c) ? show(r, c).text : '';
      if (t) w = Math.max(w, measure(t, styleAt(r, c)) / zoom + 14);
    }
    return Math.min(600, Math.ceil(w));
  }
  function fitRow(r: number) {
    let h = DEFAULT_ROW;
    const used = wb.used(name());
    for (let c = 0; c < used.cols; c++) {
      const st = styleAt(r, c);
      if (!st?.wrap || !wb.inputAt(name(), r, c)) continue;
      const lines = Math.ceil(measure(show(r, c).text, st) / zoom / Math.max(20, (colX[c + 1]! - colX[c]!) / zoom - 10));
      h = Math.max(h, lines * 17 + 7);
    }
    return h;
  }
  async function askSize(kind: 'row' | 'col', idx: number[]) {
    const current = kind === 'col' ? colW(idx[0]!) / zoom : rowH(idx[0]!) / zoom;
    const v = await promptDialog({
      title: kind === 'col' ? T.colWidth.replace('…', '') : T.rowHeight.replace('…', ''),
      fields: [{ name: 'px', label: T.sizePx, value: String(Math.round(current)), type: 'number', validate: (x) => (Number(x) > 0 && Number(x) <= 2000 ? null : '1–2000') }],
      ok: T.ok,
    });
    editor.focus();
    if (!v) return;
    change(() => (kind === 'col' ? setColWidth(idx, Number(v.px)) : setRowHeight(idx, Number(v.px))));
  }

  // ---------- sheets ----------
  const badName = (n: string) => !n.trim() || /[:\\/?*[\]]/.test(n) || n.length > 31;
  const nameTaken = (n: string, except?: string) => data.sheets.some((s) => s.name.toLowerCase() === n.toLowerCase() && s.name !== except);
  function uniqueName(base: string) {
    let n = 1;
    let cand = base;
    while (nameTaken(cand)) cand = `${base} (${++n})`;
    return cand;
  }
  function switchSheet(i: number) {
    if (i === si || i < 0 || i >= data.sheets.length) return;
    if (editing) commitEdit();
    si = i;
    active = { row: 0, col: 0 };
    anchor = { row: 0, col: 0 };
    extraRows = 0;
    extraCols = 0;
    clip = clip && clip.sheet === name() ? clip : clip;
    try {
      localStorage.setItem(`open-spreadsheet:sheet:${id}`, String(si));
    } catch {
      // not remembered
    }
    layout();
    body.scrollTop = 0;
    body.scrollLeft = 0;
    renderTabs();
    render();
    editor.focus();
  }
  function addSheet() {
    let n = data.sheets.length + 1;
    while (nameTaken(`Sheet${n}`)) n++;
    const nm = `Sheet${n}`;
    change(() => {
      wb.addSheet(nm);
      data.sheets.splice(si + 1, 0, { name: nm, cells: {} });
      wb.reorderSheets(data.sheets.map((s) => s.name));
    });
    switchSheet(si + 1);
  }
  async function renameSheet(i: number) {
    const old = data.sheets[i]!.name;
    const v = await promptDialog({
      title: T.renameSheet,
      fields: [{ name: 'name', label: T.sheetName, value: old, validate: (x) => (badName(x) ? T.sheetNameBad : nameTaken(x.trim(), old) ? T.sheetNameTaken : null) }],
      ok: T.save,
    });
    editor.focus();
    if (!v || v.name!.trim() === old) return;
    const nm = v.name!.trim();
    change(() => {
      wb.renameSheet(old, nm);
      data.sheets[i]!.name = nm;
    });
    renderTabs();
  }
  function duplicateSheet(i: number) {
    const src = data.sheets[i]!;
    const nm = uniqueName(src.name.replace(/ \(\d+\)$/, ''));
    change(() => {
      const copy: SheetData = JSON.parse(JSON.stringify({ ...src, cells: {} }));
      copy.name = nm;
      wb.addSheet(nm);
      const cells = wb.toJSON().sheets.find((s) => s.name === src.name)?.cells ?? {};
      wb.setCells(nm, Object.entries(cells).map(([a, input]) => ({ ...parseA1(a)!, input })));
      data.sheets.splice(i + 1, 0, copy);
      wb.reorderSheets(data.sheets.map((s) => s.name));
    });
    switchSheet(i + 1);
    renderTabs();
  }
  async function deleteSheet(i: number) {
    if (data.sheets.length === 1) return toast(T.lastSheet);
    const nm = data.sheets[i]!.name;
    if (!(await confirmDialog(T.deleteSheetQ(nm), T.deleteSheetBody, T.deleteSheet, true))) return editor.focus();
    change(() => {
      wb.removeSheet(nm);
      data.sheets.splice(i, 1);
      si = Math.min(si, data.sheets.length - 1);
    });
    renderTabs();
    editor.focus();
  }
  function moveSheet(i: number, to: number) {
    if (to < 0 || to >= data.sheets.length || to === i) return;
    const current = name();
    change(() => {
      const [s] = data.sheets.splice(i, 1);
      data.sheets.splice(to, 0, s!);
      wb.reorderSheets(data.sheets.map((x) => x.name));
      si = data.sheets.findIndex((x) => x.name === current);
    });
    renderTabs();
  }
  function sheetMenu(i: number, at: { x: number; y: number } | HTMLElement) {
    showMenu(
      [
        { label: T.renameSheet, icon: 'pencil', action: () => void renameSheet(i) },
        { label: T.duplicateSheet, icon: 'copy', action: () => duplicateSheet(i) },
        { label: T.moveLeft, icon: 'arrowLeft', disabled: i === 0, action: () => moveSheet(i, i - 1) },
        { label: T.moveRight, icon: 'arrowLeft', disabled: i === data.sheets.length - 1, action: () => moveSheet(i, i + 1) },
        { label: T.tabColor, swatches: { colors: TAB_COLORS, current: data.sheets[i]!.color ?? null, pick: (c) => change(() => (data.sheets[i]!.color = c ?? undefined), { relayout: false }) } },
        '-',
        { label: T.addSheet, icon: 'plus', action: () => addSheet() },
        { label: T.deleteSheet, icon: 'trash', danger: true, disabled: data.sheets.length === 1, action: () => void deleteSheet(i) },
      ],
      at,
      () => editor.focus(),
    );
  }
  let dragTab = -1;
  function renderTabs() {
    tabs.replaceChildren(
      ...data.sheets.map((s, i) => {
        const b = el('button', `tab${i === si ? ' active' : ''}`);
        b.type = 'button';
        b.draggable = true;
        if (s.color) b.style.setProperty('--tab', s.color);
        b.classList.toggle('colored', Boolean(s.color));
        b.append(el('span', '', s.name));
        b.addEventListener('mousedown', (e) => {
          if (e.button !== 0) return;
          e.preventDefault(); // keep the keyboard on the grid
          switchSheet(i);
          editor.focus();
        });
        b.addEventListener('dblclick', () => void renameSheet(i));
        b.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          switchSheet(i);
          sheetMenu(i, { x: e.clientX, y: e.clientY - 4 });
        });
        b.addEventListener('dragstart', (e) => {
          dragTab = i;
          e.dataTransfer!.effectAllowed = 'move';
        });
        b.addEventListener('dragover', (e) => {
          e.preventDefault();
          b.classList.add('drop');
        });
        b.addEventListener('dragleave', () => b.classList.remove('drop'));
        b.addEventListener('drop', (e) => {
          e.preventDefault();
          b.classList.remove('drop');
          if (dragTab >= 0) moveSheet(dragTab, i);
          dragTab = -1;
        });
        return b;
      }),
    );
    tabs.querySelector('.tab.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
  listTabs.addEventListener('click', () =>
    showMenu(
      data.sheets.map((s, i) => ({ label: s.name, checked: i === si, action: () => switchSheet(i) })),
      listTabs,
      () => editor.focus(),
    ),
  );

  // ---------- menus ----------
  function cellMenu(at: { x: number; y: number }) {
    const s = selection();
    const rows = s.r1 - s.r0 + 1;
    const cols = s.c1 - s.c0 + 1;
    const wholeRows = s.c0 === 0 && s.c1 >= nCols - 1;
    const wholeCols = s.r0 === 0 && s.r1 >= nRows - 1;
    const entries: MenuEntry[] = [
      { label: T.cut, icon: 'scissors', shortcut: `${MOD}X`, action: () => void copyFromMenu(true) },
      { label: T.copy, icon: 'copy', shortcut: `${MOD}C`, action: () => void copyFromMenu(false) },
      { label: T.paste, icon: 'clipboard', shortcut: `${MOD}V`, action: () => void pasteFromMenu(false) },
      { label: T.pasteValues, shortcut: `⇧${MOD}V`, action: () => void pasteFromMenu(true) },
      { label: T.pasteTranspose, action: () => void pasteTransposed() },
      '-',
    ];
    if (!wholeCols) entries.push({ label: rows > 1 ? T.insertRows(rows) : T.insertRowAbove, icon: 'rowAbove', action: () => insertRows(s.r0, rows) }, { label: T.insertRowBelow, icon: 'rowBelow', action: () => insertRows(s.r1 + 1, 1) });
    if (!wholeRows) entries.push({ label: cols > 1 ? T.insertCols(cols) : T.insertColLeft, icon: 'colLeft', action: () => insertCols(s.c0, cols) }, { label: T.insertColRight, icon: 'colRight', action: () => insertCols(s.c1 + 1, 1) });
    entries.push(
      '-',
      ...(wholeCols ? [] : [{ label: T.deleteRows(rows), icon: 'trash', action: () => deleteRows(s.r0, s.r1) }]),
      ...(wholeRows ? [] : [{ label: T.deleteCols(cols), icon: 'trash', action: () => deleteCols(s.c0, s.c1) }]),
      { label: T.clearContents, icon: 'eraser', shortcut: '⌫', action: () => clearContents(s) },
      { label: T.clearFormat, action: () => change(() => setStyle(styleRect(), null), { relayout: false }) },
      '-',
      { label: T.sortAsc, icon: 'sortAsc', action: () => sortBy(false) },
      { label: T.sortDesc, icon: 'sortDesc', action: () => sortBy(true) },
      { label: T.customSort, icon: 'sortAsc', action: () => void customSort() },
      '-',
      { label: mergeOf(active.row, active.col) ? T.unmerge : T.merge, icon: 'merge', disabled: !mergeOf(active.row, active.col) && rows * cols < 2, action: () => toggleMerge() },
      { label: T.freezeHere(formatA1(active.row, active.col)), icon: 'snowflake', action: () => setFreeze(active.row, active.col) },
      { label: sheet().filter ? T.removeFilter : T.addFilter, icon: 'filter', action: toggleFilter },
      { label: T.conditional, icon: 'conditional', action: () => conditionalMenu(at) },
      {
        label: T.valign,
        icon: 'alignCenter',
        submenu: (['top', 'middle', 'bottom'] as const).map((v) => ({ label: T[`valign_${v}`], checked: styleAt(active.row, active.col)?.valign === v, action: () => setValign(v) })),
      },
      { label: T.validation, icon: 'listCheck', action: () => void editValidation() },
      isCheck(active.row, active.col) ? { label: T.removeCheckbox, icon: 'check', action: removeCheckboxes } : { label: T.insertCheckbox, icon: 'check', action: insertCheckboxes },
      { label: T.defineName, icon: 'pencil', action: () => void defineName() },
      ...(linkAt(active.row, active.col) ? [{ label: T.openLink, icon: 'arrowRight', action: () => openLink(linkAt(active.row, active.col)!) }] : []),
      { label: T.tracePrecedents, icon: 'arrowRight', disabled: !wb.inputAt(sname(), active.row, active.col).startsWith('='), action: () => traceCells('pre') },
      { label: T.traceDependents, icon: 'arrowLeft', action: () => traceCells('dep') },
      ...(trace ? [{ label: T.removeArrows, icon: 'x', action: clearTrace }] : []),
      '-',
      { label: noteAt(active.row, active.col) ? T.editNote : T.addNote, icon: 'pencil', shortcut: '⇧F2', action: () => void editNote(active.row, active.col) },
      ...(noteAt(active.row, active.col) ? [{ label: T.deleteNote, icon: 'trash', action: () => setNote(active.row, active.col, '') }] : []),
    );
    showMenu(entries, at, () => editor.focus());
  }
  function headMenu(kind: 'row' | 'col', at: { x: number; y: number }) {
    const s = selection();
    const idx = kind === 'col' ? Array.from({ length: s.c1 - s.c0 + 1 }, (_, i) => s.c0 + i) : Array.from({ length: s.r1 - s.r0 + 1 }, (_, i) => s.r0 + i);
    const n = idx.length;
    const entries: MenuEntry[] =
      kind === 'col'
        ? [
            { label: T.cut, icon: 'scissors', action: () => void copyFromMenu(true) },
            { label: T.copy, icon: 'copy', action: () => void copyFromMenu(false) },
            { label: T.paste, icon: 'clipboard', action: () => void pasteFromMenu(false) },
            '-',
            { label: n > 1 ? T.insertCols(n) : T.insertColLeft, icon: 'colLeft', action: () => insertCols(s.c0, n) },
            { label: T.insertColRight, icon: 'colRight', action: () => insertCols(s.c1 + 1, 1) },
            { label: T.deleteCols(n), icon: 'trash', action: () => deleteCols(s.c0, s.c1) },
            { label: T.clearContents, icon: 'eraser', action: () => clearContents(s) },
            '-',
            { label: T.sortAsc, icon: 'sortAsc', action: () => sortBy(false) },
            { label: T.sortDesc, icon: 'sortDesc', action: () => sortBy(true) },
            '-',
            { label: T.hideCols(n), icon: 'eyeOff', action: () => hideCols(s.c0, s.c1) },
            ...(hasHiddenColsNear(s.c0, s.c1) ? [{ label: T.unhideCols, icon: 'eye', action: () => unhideCols(s.c0, s.c1) }] : []),
            { label: T.colWidth, action: () => void askSize('col', idx) },
            { label: T.autoFitCol, action: () => change(() => idx.forEach((c) => setColWidth([c], fitCol(c)))) },
            { label: T.freezeHere(columnName(s.c1)), icon: 'snowflake', action: () => setFreeze(fr(), s.c1 + 1) },
          ]
        : [
            { label: T.cut, icon: 'scissors', action: () => void copyFromMenu(true) },
            { label: T.copy, icon: 'copy', action: () => void copyFromMenu(false) },
            { label: T.paste, icon: 'clipboard', action: () => void pasteFromMenu(false) },
            '-',
            { label: n > 1 ? T.insertRows(n) : T.insertRowAbove, icon: 'rowAbove', action: () => insertRows(s.r0, n) },
            { label: T.insertRowBelow, icon: 'rowBelow', action: () => insertRows(s.r1 + 1, 1) },
            { label: T.deleteRows(n), icon: 'trash', action: () => deleteRows(s.r0, s.r1) },
            { label: T.clearContents, icon: 'eraser', action: () => clearContents(s) },
            '-',
            { label: T.hideRows(n), icon: 'eyeOff', action: () => hideRows(s.r0, s.r1) },
            ...(hasHiddenRowsNear(s.r0, s.r1) ? [{ label: T.unhideRows, icon: 'eye', action: () => unhideRows(s.r0, s.r1) }] : []),
            { label: T.rowHeight, action: () => void askSize('row', idx) },
            { label: T.autoFitRow, action: () => change(() => idx.forEach((r) => setRowHeight([r], fitRow(r)))) },
            { label: T.freezeHere(String(s.r1 + 1)), icon: 'snowflake', action: () => setFreeze(s.r1 + 1, fc()) },
          ];
    showMenu(entries, at, () => editor.focus());
  }

  // Toolbar dropdowns
  tb.size.addEventListener('click', () => {
    const cur = styleAt(active.row, active.col)?.size ?? BASE_PT;
    showMenu(
      [
        ...FONT_SIZES.map((n) => ({ label: String(n), checked: n === cur, action: () => setSize(n) })),
        '-' as const,
        {
          label: T.fontSizeOther,
          action: async () => {
            const v = await promptDialog({
              title: T.fontSize,
              fields: [{ name: 'n', label: T.fontSize, value: String(cur), type: 'number', validate: (x) => (Number(x) >= 6 && Number(x) <= 72 ? null : '6–72') }],
              ok: T.ok,
            });
            editor.focus();
            if (v) setSize(Number(v.n));
          },
        },
      ],
      tb.size,
      () => editor.focus(),
    );
  });
  tb.format.addEventListener('click', () => {
    const cur = formatOf(active.row, active.col) ?? '';
    showMenu(
      [
        ...FORMATS().map((f) => ({ label: f.sample ? `${f.label}　${f.sample}` : f.label, checked: f.code === cur, action: () => setFormat(f.code) })),
        '-' as const,
        {
          label: T.fmtCustom,
          action: async () => {
            const v = await promptDialog({ title: T.fmtCustom.replace('…', ''), fields: [{ name: 'code', label: T.customFormat, value: cur, hint: T.customFormatHint, mono: true }], ok: T.ok });
            editor.focus();
            if (v) setFormat(v.code!);
          },
        },
      ],
      tb.format,
      () => editor.focus(),
    );
  });
  tb.color.addEventListener('click', () =>
    showMenu([{ label: T.textColor, swatches: { colors: [null, ...TEXT_COLORS], current: styleAt(active.row, active.col)?.color ?? null, pick: (c) => setColor('color', c) } }], tb.color, () => editor.focus()),
  );
  tb.fill.addEventListener('click', () =>
    showMenu([{ label: T.fillColor, swatches: { colors: FILL_COLORS, current: styleAt(active.row, active.col)?.bg ?? null, pick: (c) => setColor('bg', c) } }], tb.fill, () => editor.focus()),
  );
  tb.freeze.addEventListener('click', () =>
    showMenu(
      [
        { label: T.freezeNone, checked: !fr() && !fc(), action: () => setFreeze(0, 0) },
        { label: T.freezeRow, checked: fr() === 1 && !fc(), action: () => setFreeze(1, 0) },
        { label: T.freezeCol, checked: fc() === 1 && !fr(), action: () => setFreeze(0, 1) },
        { label: T.freezeHere(formatA1(active.row, active.col)), action: () => setFreeze(active.row, active.col) },
      ],
      tb.freeze,
      () => editor.focus(),
    ),
  );
  tb.sort.addEventListener('click', () =>
    showMenu(
      [
        { label: T.sortAsc, icon: 'sortAsc', action: () => sortBy(false) },
        { label: T.sortDesc, icon: 'sortDesc', action: () => sortBy(true) },
        '-',
        { label: T.removeDuplicates, icon: 'eraser', action: removeDuplicates },
        { label: T.textToColumns, icon: 'colRight', action: () => void textToColumns() },
      ],
      tb.sort,
      () => editor.focus(),
    ),
  );
  tb.sum.addEventListener('click', () =>
    showMenu(
      [
        ...['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((f) => ({ label: f, action: () => autosum(f) })),
        '-' as const,
        { label: T.moreFunctions, action: () => beginEdit('enter', '=') },
      ],
      tb.sum,
      () => editor.focus(),
    ),
  );
  downloadBtn.addEventListener('click', () =>
    showMenu(
      [
        { label: T.exportXlsx, icon: 'download', action: () => downloadBlob(exportXlsx(serialize(), wb), `${id}.xlsx`) },
        { label: T.printPdf, icon: 'file', shortcut: `${MOD}P`, action: printSheet },
        {
          label: T.exportCsv,
          icon: 'file',
          action: () => {
            const used = wb.used(name());
            const rows: string[][] = [];
            for (let r = 0; r < used.rows; r++) {
              const line: string[] = [];
              for (let c = 0; c < used.cols; c++) line.push(wb.inputAt(name(), r, c) ? show(r, c).text : '');
              rows.push(line);
            }
            downloadBlob(new Blob([toCsv(rows)], { type: 'text/csv;charset=utf-8' }), `${id}-${name()}.csv`);
          },
        },
        '-',
        {
          label: T.importCsv,
          icon: 'upload',
          action: async () => {
            const file = await pickFile('.csv,.tsv,.txt,.xlsx');
            if (!file) return;
            try {
              const incoming = await readSpreadsheetFile(file);
              const before = data.sheets.length;
              change(() => {
                for (const s of incoming.sheets) {
                  const nm = uniqueName(s.name);
                  wb.addSheet(nm);
                  wb.setCells(nm, Object.entries(s.cells).map(([a, input]) => ({ ...parseA1(a)!, input })));
                  data.sheets.push({ ...s, name: nm, cells: {} });
                }
              });
              toast(T.importedSheets(incoming.sheets.length));
              switchSheet(before);
              renderTabs();
            } catch {
              toast(T.importFailed);
            }
          },
        },
      ],
      downloadBtn,
      () => editor.focus(),
    ),
  );
  titleBtn.addEventListener('click', async () => {
    const v = await promptDialog({ title: T.rename, fields: [{ name: 'title', label: T.title, value: data.title, validate: (x) => (x.trim() ? null : T.title) }], ok: T.save });
    editor.focus();
    if (!v) return;
    change(() => {
      data.title = v.title!.trim();
    }, { relayout: false });
    titleBtn.textContent = data.title;
    document.title = `${data.title} — open-spreadsheet`;
  });

  // ---------- charts ----------
  let selectedChart: string | null = null;
  const chartNodes = new Map<string, { node: HTMLElement; key: string }>();
  function chartCells(spec: ChartSpec): Scalar[][] {
    const r = parseRange(spec.range);
    if (!r) return [];
    const used = wb.used(name());
    const r1 = Math.min(r.r1, Math.max(r.r0, used.rows - 1));
    const c1 = Math.min(r.c1, Math.max(r.c0, used.cols - 1));
    const out: Scalar[][] = [];
    for (let row = r.r0; row <= r1; row++) {
      if (hiddenRows.has(row)) continue; // charts follow filters and hidden rows, as in Excel
      const line: Scalar[] = [];
      for (let col = r.c0; col <= c1; col++) line.push(wb.value(name(), row, col));
      out.push(line);
    }
    return out;
  }
  function paintCharts() {
    const specs = sheet().charts ?? [];
    const seen = new Set<string>();
    for (const spec of specs) {
      seen.add(spec.id);
      let entry = chartNodes.get(spec.id);
      if (!entry) {
        const node = el('div', 'chart');
        node.dataset.id = spec.id;
        node.append(el('div', 'chart-body'), el('div', 'chart-resize'));
        node.addEventListener('mousedown', (e) => chartMouseDown(e, spec.id));
        node.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          e.stopPropagation();
          selectChart(spec.id);
          chartMenu(spec.id, { x: e.clientX, y: e.clientY });
        });
        chartsLayer.append(node);
        entry = { node, key: '' };
        chartNodes.set(spec.id, entry);
      }
      const data = chartData(chartCells(spec));
      const key = JSON.stringify([spec.type, spec.title, spec.w, spec.h, data, currentTheme()]);
      if (key !== entry.key) {
        entry.node.firstElementChild!.innerHTML = chartSvg(spec, data, spec.w, spec.h);
        entry.key = key;
      }
      Object.assign(entry.node.style, {
        left: `${spec.x * zoom - body.scrollLeft}px`,
        top: `${spec.y * zoom - body.scrollTop}px`,
        width: `${spec.w * zoom}px`,
        height: `${spec.h * zoom}px`,
      });
      entry.node.classList.toggle('selected', spec.id === selectedChart);
    }
    for (const [id, entry] of chartNodes) {
      if (!seen.has(id)) {
        entry.node.remove();
        chartNodes.delete(id);
      }
    }
  }
  function selectChart(id: string | null) {
    selectedChart = id;
    for (const [cid, entry] of chartNodes) entry.node.classList.toggle('selected', cid === id);
  }
  const chartSpec = (id: string) => sheet().charts?.find((c) => c.id === id);
  function updateChart(id: string, patch: Partial<ChartSpec>) {
    change(() => {
      const spec = chartSpec(id);
      if (spec) Object.assign(spec, patch);
    }, { relayout: false });
  }
  function deleteChart(id: string) {
    change(() => {
      const sh = sheet();
      sh.charts = (sh.charts ?? []).filter((c) => c.id !== id);
      if (!sh.charts.length) sh.charts = undefined;
    }, { relayout: false });
    selectedChart = null;
    editor.focus();
  }
  function chartMouseDown(e: MouseEvent, id: string) {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    if (editing) commitEdit();
    selectChart(id);
    editor.focus();
    const spec = chartSpec(id);
    if (!spec) return;
    const resizing = (e.target as HTMLElement).classList.contains('chart-resize');
    const start = { x: e.clientX, y: e.clientY, sx: spec.x, sy: spec.y, w: spec.w, h: spec.h };
    const node = chartNodes.get(id)!.node;
    let moved = false;
    const move = (ev: MouseEvent) => {
      const dx = (ev.clientX - start.x) / zoom;
      const dy = (ev.clientY - start.y) / zoom;
      if (Math.abs(dx) + Math.abs(dy) > 2) moved = true;
      if (resizing) {
        node.style.width = `${Math.max(160, start.w + dx) * zoom}px`;
        node.style.height = `${Math.max(120, start.h + dy) * zoom}px`;
      } else {
        node.style.left = `${Math.max(0, start.sx + dx) * zoom - body.scrollLeft}px`;
        node.style.top = `${Math.max(0, start.sy + dy) * zoom - body.scrollTop}px`;
      }
    };
    const up = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      if (!moved) return;
      const dx = (ev.clientX - start.x) / zoom;
      const dy = (ev.clientY - start.y) / zoom;
      if (resizing) updateChart(id, { w: Math.round(Math.max(160, start.w + dx)), h: Math.round(Math.max(120, start.h + dy)) });
      else updateChart(id, { x: Math.round(Math.max(0, start.sx + dx)), y: Math.round(Math.max(0, start.sy + dy)) });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }
  function insertChart(type: ChartType) {
    const s = selection();
    const r = s.r0 === s.r1 && s.c0 === s.c1 ? currentRegion(active) : s;
    const id = `chart-${Date.now().toString(36)}`;
    const spec: ChartSpec = { id, type, range: rangeText(r), x: Math.round(colX[Math.min(r.c1 + 1, nCols)]! / zoom) + 16, y: Math.round(rowY[r.r0]! / zoom), w: 440, h: 270 };
    change(() => {
      const sh = sheet();
      sh.charts = [...(sh.charts ?? []), spec];
    }, { relayout: false });
    selectChart(id);
  }
  function chartMenu(id: string, at: { x: number; y: number } | HTMLElement) {
    const spec = chartSpec(id);
    if (!spec) return;
    showMenu(
      [
        { label: T.chartBar, icon: 'chart', checked: spec.type === 'bar', action: () => updateChart(id, { type: 'bar' }) },
        { label: T.chartLine, icon: 'lineChart', checked: spec.type === 'line', action: () => updateChart(id, { type: 'line' }) },
        { label: T.chartPie, icon: 'pieChart', checked: spec.type === 'pie', action: () => updateChart(id, { type: 'pie' }) },
        '-',
        {
          label: T.chartTitle,
          icon: 'pencil',
          action: async () => {
            const v = await promptDialog({ title: T.chartTitle.replace('…', ''), fields: [{ name: 't', label: T.title, value: spec.title ?? '' }], ok: T.save });
            editor.focus();
            if (v) updateChart(id, { title: v.t!.trim() || undefined });
          },
        },
        {
          label: T.chartRange,
          icon: 'sheet',
          action: async () => {
            const v = await promptDialog({
              title: T.chartRange.replace('…', ''),
              fields: [{ name: 'r', label: T.chartRange.replace('…', ''), value: spec.range, mono: true, validate: (x) => (parseRange(x) ? null : 'A1:B10') }],
              ok: T.save,
            });
            editor.focus();
            if (v) updateChart(id, { range: v.r!.trim().toUpperCase() });
          },
        },
        '-',
        { label: T.deleteChart, icon: 'trash', danger: true, shortcut: '⌫', action: () => deleteChart(id) },
      ],
      at,
      () => editor.focus(),
    );
  }

  // ---------- defined names ----------
  const quoteSheet = (n: string) => (/^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replaceAll("'", "''")}'`);
  /** Selects what a name refers to, switching sheet when it lives on another. */
  function goToName(ref: string) {
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(ref);
    if (!m) return;
    const sheetName = (m[1] ?? m[2] ?? '').replaceAll("''", "'");
    const r = parseRange(m[3]!);
    const i = data.sheets.findIndex((s) => s.name.toLowerCase() === sheetName.toLowerCase());
    if (!r || i < 0) return;
    switchSheet(i);
    moveTo({ row: r.r0, col: r.c0 });
    moveTo({ row: r.r1, col: r.c1 }, true);
  }
  async function defineName() {
    const s = selection();
    const abs = (row: number, col: number) => formatA1(row, col, true, true);
    const ref = `${quoteSheet(name())}!${s.r0 === s.r1 && s.c0 === s.c1 ? abs(s.r0, s.c0) : `${abs(s.r0, s.c0)}:${abs(s.r1, s.c1)}`}`;
    const header = wb.value(name(), s.r0 - 1 >= 0 ? s.r0 - 1 : s.r0, s.c0);
    const suggestion = typeof header === 'string' ? header.replace(/[^\w\u00C0-\uFFFF.]/g, '_') : '';
    const existing = wb.namesJSON();
    const v = await promptDialog({
      title: T.defineName.replace('…', ''),
      fields: [
        {
          name: 'n',
          label: T.nameLabel,
          value: Workbook.validName(suggestion) ? suggestion : '',
          validate: (x) => (!Workbook.validName(x.trim()) ? T.nameBad : Object.keys(existing).some((k) => k.toLowerCase() === x.trim().toLowerCase()) ? T.nameDefined : null),
        },
        { name: 'r', label: T.refersTo, value: ref, mono: true },
      ],
      ok: T.save,
    });
    editor.focus();
    if (!v) return;
    try {
      const next = { ...existing, [v.n!.trim()]: v.r!.trim().replace(/^=/, '') };
      checkpoint();
      wb.setNames(next);
      layout();
      render();
      save();
    } catch (e) {
      undoStack.pop();
      toast((e as Error).message);
    }
  }
  function nameManager(at: HTMLElement | { x: number; y: number }) {
    const names = Object.entries(wb.namesJSON());
    showMenu(
      names.length
        ? [
            ...names.map(([n, ref]) => ({ label: `${n}　${ref}`, icon: 'sheet', action: () => goToName(ref) })),
            '-' as const,
            {
              label: T.deleteName,
              icon: 'trash',
              submenu: names.map(([n]) => ({
                label: n,
                action: () =>
                  change(() => {
                    const rest = wb.namesJSON();
                    delete rest[n];
                    wb.setNames(rest);
                  }),
              })),
            },
          ]
        : [{ label: T.noNames, disabled: true }],
      at,
      () => editor.focus(),
    );
  }

  // ---------- data tools ----------
  /** The block to work on: the selection, or the data around the cell when one cell is selected. */
  const dataBlock = () => {
    const s = selection();
    return s.r0 === s.r1 && s.c0 === s.c1 ? currentRegion(active) : { ...s, r1: Math.min(s.r1, wb.used(name()).rows - 1), c1: Math.min(s.c1, wb.used(name()).cols - 1) };
  };
  const looksLikeHeader = (r: { r0: number; c0: number; r1: number; c1: number }) =>
    r.r1 > r.r0 && Array.from({ length: r.c1 - r.c0 + 1 }, (_, i) => wb.value(name(), r.r0, r.c0 + i)).every((v) => typeof v === 'string' && v !== '');

  /** Excel's Remove Duplicates: rows equal in every column of the block keep their first copy; the rest close up. */
  function removeDuplicates() {
    const r = dataBlock();
    const header = looksLikeHeader(r);
    const start = header ? r.r0 + 1 : r.r0;
    const seen = new Set<string>();
    const keep: number[] = [];
    for (let row = start; row <= r.r1; row++) {
      const key = JSON.stringify(Array.from({ length: r.c1 - r.c0 + 1 }, (_, i) => show(row, r.c0 + i).text.toLowerCase()));
      if (seen.has(key)) continue;
      seen.add(key);
      keep.push(row);
    }
    const removed = r.r1 - start + 1 - keep.length;
    if (!removed) return toast(T.noDuplicates);
    change(() => {
      const cells: { row: number; col: number; input: string }[] = [];
      const styles: { at: { r0: number; c0: number; r1: number; c1: number }; st?: CellStyle }[] = [];
      keep.forEach((from, i) => {
        const to = start + i;
        for (let c = r.c0; c <= r.c1; c++) {
          cells.push({ row: to, col: c, input: offsetFormula(wb.inputAt(name(), from, c), to - from, 0) });
          styles.push({ at: { r0: to, c0: c, r1: to, c1: c }, st: styleAt(from, c) });
        }
      });
      for (let row = start + keep.length; row <= r.r1; row++) {
        for (let c = r.c0; c <= r.c1; c++) {
          cells.push({ row, col: c, input: '' });
          styles.push({ at: { r0: row, c0: c, r1: row, c1: c } });
        }
      }
      setInputs(cells);
      for (const { at, st } of styles) {
        setStyle(at, null);
        if (st) setStyle(at, st);
      }
    });
    toast(T.removedDuplicates(removed, keep.length));
  }

  /** Excel's Text to Columns: split each cell of one column at a delimiter into the columns to its right. */
  async function textToColumns() {
    const s = selection();
    const r = s.r0 === s.r1 ? currentRegion(active) : s;
    const col = active.col;
    const v = await promptDialog({ title: T.textToColumns.replace('…', ''), fields: [{ name: 'd', label: T.delimiter, value: ',', hint: T.delimiterHint, mono: true }], ok: T.ok });
    editor.focus();
    if (!v) return;
    const d = v.d === '\\t' || v.d === 'tab' ? '\t' : v.d === '' ? ' ' : v.d!;
    let split = 0;
    change(() => {
      const cells: { row: number; col: number; input: string }[] = [];
      for (let row = r.r0; row <= r.r1; row++) {
        const input = wb.inputAt(name(), row, col);
        if (!input || input.startsWith('=') || !input.includes(d)) continue;
        const parts = input.split(d).map((p) => p.trim());
        parts.forEach((p, i) => cells.push({ row, col: col + i, input: p }));
        split++;
      }
      setInputs(cells);
    });
    toast(T.splitDone(split));
  }

  /** Paste with rows and columns swapped (the internal clipboard, else the system's text). */
  async function pasteTransposed() {
    let grid2: { input: string; style?: CellStyle }[][] | null = clip ? clip.cells : null;
    let from: { r0: number; c0: number } | null = clip ? clip.from : null;
    if (!grid2) {
      try {
        const text = await navigator.clipboard.readText();
        grid2 = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => l.split('\t').map((input) => ({ input })));
      } catch {
        return toast(T.pasteBlocked);
      }
    }
    const src = grid2;
    const at = { row: selection().r0, col: selection().c0 };
    change(() => {
      const cells: { row: number; col: number; input: string }[] = [];
      src.forEach((line, i) =>
        line.forEach((cell, k) => {
          const row = at.row + k;
          const col = at.col + i;
          const input = from ? offsetFormula(cell.input, row - (from.r0 + i), col - (from.c0 + k)) : cell.input;
          cells.push({ row, col, input });
          setStyle({ r0: row, c0: col, r1: row, c1: col }, null);
          if (cell.style) setStyle({ r0: row, c0: col, r1: row, c1: col }, cell.style);
        }),
      );
      setInputs(cells);
    });
    from = null;
    anchor = { row: at.row + (src[0]?.length ?? 1) - 1, col: at.col + src.length - 1 };
    active = at;
    render();
  }

  // ---------- print ----------
  /** The sheet as a paged table in a hidden frame, handed to the print dialog (which can save a PDF). */
  function printSheet() {
    if (editing) commitEdit();
    const used = wb.used(name());
    const sh = sheet();
    const html = printHtml({
      title: data.title,
      sheetName: name(),
      rows: Math.max(used.rows, ...merges.map((m) => m.r1 + 1)),
      cols: Math.max(used.cols, ...merges.map((m) => m.c1 + 1)),
      cell: (r, c) => {
        const input = wb.inputAt(name(), r, c);
        const style = styleAt(r, c);
        if (!input && !style) return null;
        const s = input ? show(r, c) : null;
        const v = s?.v ?? null;
        const look = input ? lookAt(r, c) : null;
        const align = style?.align ?? (typeof v === 'number' ? 'right' : typeof v === 'boolean' || v instanceof CellError ? 'center' : 'left');
        return { text: s?.text ?? '', align, style, bg: look?.bg, color: s?.color ?? look?.color };
      },
      colWidth: (c) => sh.cols?.[columnName(c)] ?? DEFAULT_COL,
      rowHidden: (r) => hiddenRows.has(r),
      colHidden: (c) => hiddenColSet.has(c),
      merges,
      headerRows: sh.freeze?.rows ?? 0,
      charts: (sh.charts ?? []).map((spec) => chartSvg(spec, chartData(chartCells(spec)), spec.w, spec.h)),
      lang: LANGS.find((l) => l.id === lang)?.html ?? 'en',
    });
    const frame = el('iframe', 'print-frame');
    frame.setAttribute('aria-hidden', 'true');
    document.body.append(frame);
    const doc = frame.contentDocument!;
    doc.open();
    doc.write(html);
    doc.close();
    const done = () => setTimeout(() => frame.remove(), 1000);
    frame.contentWindow!.addEventListener('afterprint', done);
    setTimeout(() => {
      frame.contentWindow!.focus();
      frame.contentWindow!.print();
      done();
    }, 50);
    lastPrintHtml = html;
  }
  /** The last page handed to the print dialog, for tests. */
  let lastPrintHtml = '';
  (window as unknown as { __lastPrint?: () => string }).__lastPrint = () => lastPrintHtml;

  // ---------- format painter ----------
  /** The styles picked up by the format painter, by offset from the copied block's corner. */
  let painter: { cells: (CellStyle | undefined)[][] } | null = null;
  function pickUpFormat() {
    const s = selection();
    painter = { cells: Array.from({ length: s.r1 - s.r0 + 1 }, (_, i) => Array.from({ length: s.c1 - s.c0 + 1 }, (_, k) => styleAt(s.r0 + i, s.c0 + k))) };
    tb.painter.classList.add('on');
    grid.classList.add('painting');
  }
  function applyFormat() {
    if (!painter) return;
    const s = selection();
    const p = painter;
    const h = p.cells.length;
    const w = p.cells[0]?.length ?? 1;
    // One cell picked: every target cell gets it; a block tiles over the target.
    const rows = Math.max(s.r1 - s.r0 + 1, h);
    const cols = Math.max(s.c1 - s.c0 + 1, w);
    change(() => {
      for (let i = 0; i < rows; i++) {
        for (let k = 0; k < cols; k++) {
          const st = p.cells[i % h]![k % w];
          const at = { r0: s.r0 + i, c0: s.c0 + k, r1: s.r0 + i, c1: s.c0 + k };
          setStyle(at, null);
          if (st) setStyle(at, st);
        }
      }
    }, { relayout: false });
    painter = null;
    tb.painter.classList.remove('on');
    grid.classList.remove('painting');
  }
  tb.painter.addEventListener('click', () => {
    if (painter) {
      painter = null;
      tb.painter.classList.remove('on');
      grid.classList.remove('painting');
    } else pickUpFormat();
  });

  // ---------- hiding rows and columns ----------
  function hideRows(r0: number, r1: number) {
    change(() => {
      const set = new Set(sheet().hiddenRows ?? []);
      for (let r = r0; r <= r1; r++) set.add(r + 1);
      sheet().hiddenRows = [...set].sort((a, b) => a - b);
    });
  }
  function hideCols(c0: number, c1: number) {
    change(() => {
      const set = new Set(sheet().hiddenCols ?? []);
      for (let c = c0; c <= c1; c++) set.add(columnName(c));
      sheet().hiddenCols = [...set].sort((a, b) => columnIndex(a) - columnIndex(b));
    });
  }
  /** Unhides what lies in, or right beside, the selection — as Excel does when you select across a hidden band. */
  function unhideRows(r0: number, r1: number) {
    change(() => {
      const next = (sheet().hiddenRows ?? []).filter((n) => n - 1 < r0 - 1 || n - 1 > r1 + 1);
      sheet().hiddenRows = next.length ? next : undefined;
    });
  }
  function unhideCols(c0: number, c1: number) {
    change(() => {
      const next = (sheet().hiddenCols ?? []).filter((l) => columnIndex(l) < c0 - 1 || columnIndex(l) > c1 + 1);
      sheet().hiddenCols = next.length ? next : undefined;
    });
  }
  const hasHiddenRowsNear = (r0: number, r1: number) => (sheet().hiddenRows ?? []).some((n) => n - 1 >= r0 - 1 && n - 1 <= r1 + 1);
  const hasHiddenColsNear = (c0: number, c1: number) => (sheet().hiddenCols ?? []).some((l) => columnIndex(l) >= c0 - 1 && columnIndex(l) <= c1 + 1);

  // ---------- borders ----------
  function setBorders(kind: 'all' | 'outer' | 'thick' | 'top' | 'bottom' | 'left' | 'right' | 'none') {
    const s = styleRect();
    const ink = '#1f1f1f';
    change(() => {
      const sh = sheet();
      sh.styles ??= {};
      for (let r = s.r0; r <= s.r1; r++) {
        for (let c = s.c0; c <= s.c1; c++) {
          const a = formatA1(r, c);
          const cur = { ...sh.styles[a] };
          let b = kind === 'none' ? {} : { ...cur.border };
          const edge = { top: r === s.r0, bottom: r === s.r1, left: c === s.c0, right: c === s.c1 };
          if (kind === 'all') b = { ...b, top: ink, bottom: ink, left: ink, right: ink };
          if (kind === 'outer' || kind === 'thick') {
            for (const side of ['top', 'bottom', 'left', 'right'] as const) if (edge[side]) b[side] = ink;
            if (kind === 'thick' && Object.values(edge).some(Boolean)) b.thick = true;
          }
          if (kind === 'top' && edge.top) b.top = ink;
          if (kind === 'bottom' && edge.bottom) b.bottom = ink;
          if (kind === 'left' && edge.left) b.left = ink;
          if (kind === 'right' && edge.right) b.right = ink;
          if (Object.keys(b).filter((k) => k !== 'thick').length) cur.border = b;
          else delete cur.border;
          if (Object.keys(cur).length) sh.styles[a] = cur;
          else delete sh.styles[a];
        }
      }
    }, { relayout: false });
  }
  const borderItems = (): MenuEntry[] => [
    { label: T.borderAll, icon: 'border', action: () => setBorders('all') },
    { label: T.borderOuter, action: () => setBorders('outer') },
    { label: T.borderThickOuter, action: () => setBorders('thick') },
    { label: T.borderTop, action: () => setBorders('top') },
    { label: T.borderBottom, action: () => setBorders('bottom') },
    { label: T.borderLeft, action: () => setBorders('left') },
    { label: T.borderRight, action: () => setBorders('right') },
    '-',
    { label: T.borderNone, icon: 'x', action: () => setBorders('none') },
  ];

  // ---------- filters, validation, conditional formatting ----------
  const shownText = (r: number, c: number) => (wb.inputAt(name(), r, c) ? show(r, c).text : '');
  function toggleFilter() {
    if (sheet().filter) {
      change(() => {
        sheet().filter = undefined;
      });
      return;
    }
    const s = selection();
    const r = s.r0 === s.r1 ? currentRegion(active) : s;
    change(() => {
      sheet().filter = { range: rangeText(r) };
    });
  }
  function filterMenu(col: number, at: HTMLElement | { x: number; y: number }) {
    const r = filterRect();
    if (!r) return;
    const letter = columnName(col);
    const hidden = new Set(filterHidden(col));
    const values = columnValues(sheet(), col, shownText, wb.used(name())).slice(0, 200);
    const setHidden = (next: Set<string>) =>
      change(() => {
        const f = sheet().filter!;
        const h = { ...(f.hidden ?? {}) };
        if (next.size) h[letter] = [...next];
        else delete h[letter];
        f.hidden = Object.keys(h).length ? h : undefined;
      });
    const sortColumn = (desc: boolean) => {
      active = { row: r.r0 + 1, col };
      anchor = { row: Math.max(r.r0 + 1, Math.min(r.r1, wb.used(name()).rows - 1)), col: col === r.c0 ? r.c1 : r.c0 };
      sortBy(desc);
    };
    const pos = at instanceof HTMLElement ? at.getBoundingClientRect() : null;
    const point = pos ? { x: pos.left, y: pos.bottom + 4 } : (at as { x: number; y: number });
    showMenu(
      [
        { label: T.sortAsc, icon: 'sortAsc', action: () => sortColumn(false) },
        { label: T.sortDesc, icon: 'sortDesc', action: () => sortColumn(true) },
        '-',
        { label: T.showAll, icon: 'filter', disabled: hidden.size === 0, action: () => setHidden(new Set()) },
        ...values.map((v) => ({
          label: v === BLANK ? T.blanks : v,
          checked: !hidden.has(v),
          action: () => {
            const next = new Set(hidden);
            if (next.has(v)) next.delete(v);
            else next.add(v);
            setHidden(next);
            filterMenu(col, point);
          },
        })),
      ],
      point,
      () => editor.focus(),
    );
  }
  async function editValidation() {
    const s = selection();
    const current = validationAt(sheet(), active.row, active.col) ?? [];
    const v = await promptDialog({
      title: `${T.validation.replace('…', '')} · ${rangeText(s)}`,
      fields: [{ name: 'list', label: T.validationList, value: current.join(', '), hint: T.validationHint, placeholder: 'Yes, No, Maybe' }],
      ok: T.save,
    });
    editor.focus();
    if (!v) return;
    const list = v.list!.split(/[,，、]/).map((x) => x.trim()).filter(Boolean);
    const range = rangeText(s);
    change(() => {
      const sh = sheet();
      const rest = (sh.validation ?? []).filter((x) => x.range !== range);
      sh.validation = list.length ? [...rest, { range, list }] : rest;
      if (!sh.validation.length) sh.validation = undefined;
    }, { relayout: false });
  }
  dvBtn.addEventListener('mousedown', (e) => {
    e.preventDefault();
    e.stopPropagation();
    const list = validationAt(sheet(), active.row, active.col) ?? [];
    showMenu(
      list.map((v) => ({ label: v, checked: wb.inputAt(name(), active.row, active.col) === v, action: () => change(() => setInputs([{ row: active.row, col: active.col, input: v }])) })),
      dvBtn,
      () => editor.focus(),
    );
  });
  function addRule(rule: ConditionalRule) {
    change(() => {
      const sh = sheet();
      sh.conditional = [...(sh.conditional ?? []), rule];
    }, { relayout: false });
  }
  async function askRule(type: 'gt' | 'lt' | 'eq' | 'between' | 'contains') {
    const range = rangeText(selection());
    const fields =
      type === 'between'
        ? [
            { name: 'a', label: T.cfValue, validate: (x: string) => (Number.isFinite(Number(x)) && x !== '' ? null : T.cfValue) },
            { name: 'b', label: T.cfValue2, validate: (x: string) => (Number.isFinite(Number(x)) && x !== '' ? null : T.cfValue) },
          ]
        : [{ name: 'a', label: type === 'contains' ? T.cfText : T.cfValue, validate: (x: string) => (x.trim() ? null : T.cfValue) }];
    const title = { gt: T.cfGreater, lt: T.cfLess, eq: T.cfEqual, between: T.cfBetween, contains: T.cfContains }[type].replace('…', '');
    const v = await promptDialog({ title: `${title} · ${range}`, fields, ok: T.ok });
    editor.focus();
    if (!v) return;
    const value = (x: string) => (Number.isFinite(Number(x)) && x.trim() !== '' ? Number(x) : x);
    if (type === 'between') addRule({ range, type, value: Number(v.a), value2: Number(v.b), style: DEFAULT_HIGHLIGHT });
    else if (type === 'contains') addRule({ range, type, text: v.a!, style: DEFAULT_HIGHLIGHT });
    else addRule({ range, type, value: value(v.a!), style: DEFAULT_HIGHLIGHT });
  }
  const conditionalMenu = (at: HTMLElement | { x: number; y: number }) => {
    const range = rangeText(selection());
    const s = selection();
    const touching = (sheet().conditional ?? []).filter((r) => {
      const g = parseRange(r.range);
      return g && overlaps(g, s);
    });
    showMenu(
      [
        { label: T.cfGreater, action: () => void askRule('gt') },
        { label: T.cfLess, action: () => void askRule('lt') },
        { label: T.cfBetween, action: () => void askRule('between') },
        { label: T.cfEqual, action: () => void askRule('eq') },
        { label: T.cfContains, action: () => void askRule('contains') },
        { label: T.cfDuplicates, action: () => addRule({ range, type: 'duplicate', style: DEFAULT_HIGHLIGHT }) },
        '-',
        { label: T.cfScaleRYG, action: () => addRule({ range, type: 'colorScale', colors: ['#f8696b', '#ffeb84', '#63be7b'] }) },
        { label: T.cfScaleWG, action: () => addRule({ range, type: 'colorScale', colors: ['#ffffff', '#63be7b'] }) },
        { label: T.cfDataBar, action: () => addRule({ range, type: 'dataBar', color: '#638ec6' }) },
        '-',
        {
          label: T.cfClearSel,
          disabled: touching.length === 0,
          action: () =>
            change(() => {
              sheet().conditional = (sheet().conditional ?? []).filter((r) => !touching.includes(r));
              if (!sheet().conditional!.length) sheet().conditional = undefined;
            }, { relayout: false }),
        },
        {
          label: T.cfClearAll,
          disabled: !sheet().conditional?.length,
          action: () =>
            change(() => {
              sheet().conditional = undefined;
            }, { relayout: false }),
        },
      ],
      at,
      () => editor.focus(),
    );
  };
  tb.cond.addEventListener('click', () => conditionalMenu(tb.cond));
  tb.chart.addEventListener('click', () =>
    showMenu(
      [
        { label: T.chartBar, icon: 'chart', action: () => insertChart('bar') },
        { label: T.chartLine, icon: 'lineChart', action: () => insertChart('line') },
        { label: T.chartPie, icon: 'pieChart', action: () => insertChart('pie') },
      ],
      tb.chart,
      () => editor.focus(),
    ),
  );
  tb.border.addEventListener('click', () => showMenu(borderItems(), tb.border, () => editor.focus()));
  tb.filter.addEventListener('click', () => {
    if (!sheet().filter) return toggleFilter();
    showMenu(
      [
        { label: T.showAll, icon: 'filter', disabled: !sheet().filter?.hidden, action: () => change(() => (sheet().filter!.hidden = undefined)) },
        { label: T.removeFilter, icon: 'x', action: toggleFilter },
      ],
      tb.filter,
      () => editor.focus(),
    );
  });

  // ---------- find and replace ----------
  const findInput = el('input', 'input');
  findInput.placeholder = T.find;
  findInput.spellcheck = false;
  const findCount = el('span', 'find-count');
  const replaceInput = el('input', 'input');
  replaceInput.placeholder = T.replaceWith;
  replaceInput.spellcheck = false;
  const caseBox = el('label', 'find-case');
  const caseCheck = el('input');
  caseCheck.type = 'checkbox';
  caseBox.append(caseCheck, el('span', '', T.matchCase));
  const findPrev = iconButton('arrowUp', '⇧Enter', () => findStep(-1));
  const findNext = iconButton('arrowDown', 'Enter', () => findStep(1));
  const findClose = iconButton('x', 'Esc', () => closeFind());
  const replaceOne = el('button', 'btn small', T.replace);
  const replaceAllBtn = el('button', 'btn small', T.replaceAll);
  const row1 = el('div', 'find-row');
  row1.append(findInput, findCount, findPrev, findNext, findClose);
  const row2 = el('div', 'find-row');
  row2.append(replaceInput, replaceOne, replaceAllBtn);
  const row3 = el('div', 'find-row');
  row3.append(caseBox);
  findbar.append(row1, row2, row3);
  let hits: Pos[] = [];
  let hitIndex = -1;
  function findAll() {
    const q = findInput.value;
    hits = [];
    if (!q) return;
    const used = wb.used(name());
    const norm = (s: string) => (caseCheck.checked ? s : s.toLowerCase());
    const needle = norm(q);
    for (let r = 0; r < used.rows; r++) {
      for (let c = 0; c < used.cols; c++) {
        const input = wb.inputAt(name(), r, c);
        if (!input) continue;
        if (norm(input).includes(needle) || norm(show(r, c).text).includes(needle)) hits.push({ row: r, col: c });
      }
    }
  }
  function findStep(d: number) {
    findAll();
    if (!hits.length) {
      findCount.textContent = findInput.value ? T.noResults : '';
      return;
    }
    const cur = hits.findIndex((h) => h.row === active.row && h.col === active.col);
    hitIndex = cur === -1 ? (d > 0 ? hits.findIndex((h) => h.row > active.row || (h.row === active.row && h.col > active.col)) : -1) : cur + d;
    if (hitIndex < 0 || hitIndex >= hits.length) hitIndex = d > 0 ? 0 : hits.length - 1;
    if (cur === -1 && hitIndex === -1) hitIndex = 0;
    findCount.textContent = `${hitIndex + 1} / ${hits.length}`;
    moveTo(hits[hitIndex]!);
  }
  function replaceIn(input: string): string {
    const q = findInput.value;
    const flags = caseCheck.checked ? 'g' : 'gi';
    return input.replace(new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags), () => replaceInput.value);
  }
  replaceOne.addEventListener('click', () => {
    findAll();
    const at = hits.find((h) => h.row === active.row && h.col === active.col);
    if (at) {
      const input = wb.inputAt(name(), at.row, at.col);
      const next = replaceIn(input);
      if (next !== input) change(() => setInputs([{ ...at, input: next }]));
    }
    findStep(1);
  });
  replaceAllBtn.addEventListener('click', () => {
    findAll();
    const cells = hits.map((h) => ({ ...h, input: replaceIn(wb.inputAt(name(), h.row, h.col)) })).filter((c) => c.input !== wb.inputAt(name(), c.row, c.col));
    if (cells.length) change(() => setInputs(cells));
    toast(T.replaced(cells.length));
    findAll();
    findCount.textContent = hits.length ? `${hits.length}` : '';
  });
  findInput.addEventListener('input', () => {
    findAll();
    findCount.textContent = findInput.value ? (hits.length ? `${hits.length}` : T.noResults) : '';
  });
  findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      findStep(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') closeFind();
  });
  replaceInput.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeFind();
    if (e.key === 'Enter') replaceOne.click();
  });
  function openFind() {
    findbar.hidden = false;
    findInput.focus();
    findInput.select();
  }
  function closeFind() {
    findbar.hidden = true;
    editor.focus();
  }

  function paintNotesBadge() {
    const n = noteCount();
    notesBadge.textContent = String(n);
    notesBadge.hidden = n === 0;
  }
  function notesMenu() {
    const items: MenuEntry[] = data.sheets.flatMap((sh, i) =>
      Object.entries(sh.notes ?? {}).map(([a, text]) => ({
        label: `${sh.name}!${a}　${text.replace(/\s+/g, ' ').slice(0, 48)}${text.length > 48 ? '…' : ''}`,
        icon: 'pencil',
        action: () => {
          switchSheet(i);
          const p = parseA1(a)!;
          moveTo(p);
        },
      })),
    );
    showMenu(items.length ? items : [{ label: T.noNotes, disabled: true }], notesBtn, () => editor.focus());
  }
  const notePop = el('div', 'note-pop');
  notePop.hidden = true;
  document.body.append(notePop);
  let noteHover = '';
  grid.addEventListener('mousemove', (e) => {
    if (drag || resize) return (notePop.hidden = true), undefined;
    const r = grid.getBoundingClientRect();
    if (e.clientX - r.left < HEAD_W || e.clientY - r.top < HEAD_H) {
      notePop.hidden = true;
      noteHover = '';
      return;
    }
    const at = cellFromEvent(e);
    const m = mergeOf(at.row, at.col);
    const p = m ? { row: m.r0, col: m.c0 } : at;
    const text = noteAt(p.row, p.col);
    const key = text ? `${name()}!${p.row},${p.col}` : '';
    if (key === noteHover) return;
    noteHover = key;
    if (!text) {
      notePop.hidden = true;
      return;
    }
    const v = viewRect(m ?? { r0: p.row, c0: p.col, r1: p.row, c1: p.col });
    notePop.replaceChildren(el('b', '', formatA1(p.row, p.col)), document.createTextNode(text));
    notePop.hidden = false;
    if (v) {
      notePop.style.left = `${Math.min(innerWidth - 310, r.left + HEAD_W + v.x + v.w + 6)}px`;
      notePop.style.top = `${r.top + HEAD_H + v.y}px`;
    }
  });
  grid.addEventListener('mouseleave', () => {
    notePop.hidden = true;
    noteHover = '';
  });

  async function openPalette() {
    if (editing) commitEdit();
    const s = selection();
    const rows = s.r1 - s.r0 + 1;
    const cols = s.c1 - s.c0 + 1;
    const cmd = (group: string, label: string, run: () => void, icon?: string, hint?: string, keywords?: string): Command => ({ group, label, run, icon, hint, keywords });
    const list: Command[] = [
      cmd(T.gEdit, T.undo, undo, 'undo', `${MOD}Z`, 'undo'),
      cmd(T.gEdit, T.redo, redo, 'redo', `⇧${MOD}Z`, 'redo'),
      cmd(T.gEdit, T.copy, () => void copyFromMenu(false), 'copy', `${MOD}C`, 'copy'),
      cmd(T.gEdit, T.cut, () => void copyFromMenu(true), 'scissors', `${MOD}X`, 'cut'),
      cmd(T.gEdit, T.paste, () => void pasteFromMenu(false), 'clipboard', `${MOD}V`, 'paste'),
      cmd(T.gEdit, T.pasteValues, () => void pasteFromMenu(true), undefined, `⇧${MOD}V`, 'paste values'),
      cmd(T.gEdit, T.clearContents, () => clearContents(selection()), 'eraser', '⌫', 'clear delete'),
      cmd(T.gEdit, T.findReplace, openFind, 'search', `${MOD}F`, 'find replace search'),
      cmd(T.gEdit, noteAt(active.row, active.col) ? T.editNote : T.addNote, () => void editNote(active.row, active.col), 'pencil', '⇧F2', 'note comment'),
      cmd(T.gEdit, T.notes, notesMenu, 'pencil', undefined, 'notes comments list'),
      cmd(T.gEdit, T.defineName, () => void defineName(), 'pencil', undefined, 'define name named range'),
      cmd(T.gEdit, T.tracePrecedents, () => traceCells('pre'), 'arrowRight', undefined, 'trace precedents where from source audit'),
      cmd(T.gEdit, T.traceDependents, () => traceCells('dep'), 'arrowLeft', undefined, 'trace dependents used by audit'),
      cmd(T.gEdit, T.removeArrows, clearTrace, 'x', undefined, 'remove arrows trace clear'),
      cmd(T.gEdit, T.customSort, () => void customSort(), 'sortAsc', undefined, 'custom sort multiple columns levels'),
      cmd(T.gView, showFormulas ? T.hideFormulas : T.showFormulas, toggleFormulas, 'fx', '⌃`', 'show formulas audit'),
      cmd(T.gEdit, T.nextError, nextError, 'search', undefined, 'error checking next error audit div ref name'),
      cmd(T.gEdit, T.nameManager, () => nameManager(namebox), 'list', undefined, 'name manager names'),
      ...Object.entries(wb.namesJSON()).map(([n, ref]) => cmd(T.gEdit, `${T.nameLabel}：${n}`, () => goToName(ref), 'sheet', ref, `go name ${n}`)),
      cmd(T.gEdit, T.removeDuplicates, removeDuplicates, 'eraser', undefined, 'remove duplicates unique'),
      cmd(T.gEdit, T.textToColumns, () => void textToColumns(), 'colRight', undefined, 'text to columns split delimiter'),
      cmd(T.gEdit, T.pasteTranspose, () => void pasteTransposed(), 'clipboard', undefined, 'paste transpose'),
      cmd(T.gEdit, T.sortAsc, () => sortBy(false), 'sortAsc', undefined, 'sort ascending'),
      cmd(T.gEdit, sheet().filter ? T.removeFilter : T.addFilter, toggleFilter, 'filter', undefined, 'filter autofilter'),
      cmd(T.gEdit, T.validation, () => void editValidation(), 'listCheck', undefined, 'data validation dropdown list'),
      ...(['all', 'outer', 'thick', 'bottom', 'none'] as const).map((k) =>
        cmd(T.gFormat, `${T.borders}：${{ all: T.borderAll, outer: T.borderOuter, thick: T.borderThickOuter, bottom: T.borderBottom, none: T.borderNone }[k]}`, () => setBorders(k), 'border', undefined, 'border lines'),
      ),
      cmd(T.gRowsCols, T.hideRows(s.r1 - s.r0 + 1), () => hideRows(s.r0, s.r1), 'eyeOff', undefined, 'hide rows'),
      cmd(T.gRowsCols, T.hideCols(s.c1 - s.c0 + 1), () => hideCols(s.c0, s.c1), 'eyeOff', undefined, 'hide columns'),
      cmd(T.gRowsCols, T.unhideRows, () => unhideRows(s.r0, s.r1), 'eye', undefined, 'unhide show rows'),
      cmd(T.gRowsCols, T.unhideCols, () => unhideCols(s.c0, s.c1), 'eye', undefined, 'unhide show columns'),
      ...(['bar', 'line', 'pie'] as const).map((k) => cmd(T.chart, `${T.insertChart}：${{ bar: T.chartBar, line: T.chartLine, pie: T.chartPie }[k]}`, () => insertChart(k), k === 'bar' ? 'chart' : k === 'line' ? 'lineChart' : 'pieChart', undefined, `chart graph ${k}`)),
      cmd(T.gFormat, T.conditional, () => conditionalMenu(tb.cond), 'conditional', undefined, 'conditional formatting color scale data bar highlight'),
      cmd(T.gEdit, T.sortDesc, () => sortBy(true), 'sortDesc', undefined, 'sort descending'),
      ...['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((f) => cmd(T.autosum, `${f}( )`, () => autosum(f), 'sigma', undefined, `autosum ${f.toLowerCase()}`)),
      cmd(T.gFormat, T.bold, () => toggleStyle('b'), 'bold', `${MOD}B`, 'bold'),
      cmd(T.gEdit, T.insertCheckbox, insertCheckboxes, 'check', undefined, 'checkbox tick box todo'),
      cmd(T.gEdit, T.removeCheckbox, removeCheckboxes, 'check', undefined, 'checkbox remove'),
      ...(['top', 'middle', 'bottom'] as const).map((v) => cmd(T.gFormat, `${T.valign}：${T[`valign_${v}`]}`, () => setValign(v), 'alignCenter', undefined, `vertical align ${v}`)),
      cmd(T.gFormat, T.italic, () => toggleStyle('i'), 'italic', `${MOD}I`, 'italic'),
      cmd(T.gFormat, T.underline, () => toggleStyle('u'), 'underline', `${MOD}U`, 'underline'),
      cmd(T.gFormat, T.strike, () => toggleStyle('s'), 'strike', undefined, 'strikethrough'),
      cmd(T.gFormat, T.wrap, () => toggleStyle('wrap'), 'wrap', undefined, 'wrap'),
      cmd(T.gFormat, T.alignLeft, () => setAlign('left'), 'alignLeft', undefined, 'align left'),
      cmd(T.gFormat, T.alignCenter, () => setAlign('center'), 'alignCenter', undefined, 'align center'),
      cmd(T.gFormat, T.alignRight, () => setAlign('right'), 'alignRight', undefined, 'align right'),
      ...FORMATS().map((f) => cmd(T.gFormat, `${T.format}：${f.label}`, () => setFormat(f.code), undefined, f.sample, `format number ${f.code}`)),
      cmd(T.gFormat, T.incDecimals, () => bumpDecimals(1), 'incDec', undefined, 'decimals'),
      cmd(T.gFormat, T.decDecimals, () => bumpDecimals(-1), 'decDec', undefined, 'decimals'),
      cmd(T.gFormat, mergeOf(active.row, active.col) ? T.unmerge : T.merge, toggleMerge, 'merge', undefined, 'merge cells'),
      cmd(T.gFormat, T.clearFormat, () => change(() => setStyle(styleRect(), null), { relayout: false }), 'eraser', undefined, 'clear formatting'),
      cmd(T.gRowsCols, rows > 1 ? T.insertRows(rows) : T.insertRowAbove, () => insertRows(s.r0, rows), 'rowAbove', undefined, 'insert row'),
      cmd(T.gRowsCols, T.insertRowBelow, () => insertRows(s.r1 + 1, 1), 'rowBelow', undefined, 'insert row'),
      cmd(T.gRowsCols, cols > 1 ? T.insertCols(cols) : T.insertColLeft, () => insertCols(s.c0, cols), 'colLeft', undefined, 'insert column'),
      cmd(T.gRowsCols, T.insertColRight, () => insertCols(s.c1 + 1, 1), 'colRight', undefined, 'insert column'),
      cmd(T.gRowsCols, T.deleteRows(rows), () => deleteRows(s.r0, s.r1), 'trash', undefined, 'delete row'),
      cmd(T.gRowsCols, T.deleteCols(cols), () => deleteCols(s.c0, s.c1), 'trash', undefined, 'delete column'),
      cmd(T.gRowsCols, T.autoFitCol, () => change(() => range(s.c0, s.c1).forEach((c) => setColWidth([c], fitCol(c)))), undefined, undefined, 'fit column width'),
      cmd(T.gView, T.freezeNone, () => setFreeze(0, 0), 'snowflake', undefined, 'unfreeze'),
      cmd(T.gView, T.freezeRow, () => setFreeze(1, 0), 'snowflake', undefined, 'freeze top row'),
      cmd(T.gView, T.freezeCol, () => setFreeze(0, 1), 'snowflake', undefined, 'freeze first column'),
      cmd(T.gView, T.freezeHere(formatA1(active.row, active.col)), () => setFreeze(active.row, active.col), 'snowflake', undefined, 'freeze panes'),
      cmd(T.gView, T.toggleTheme, () => themeBtn.click(), 'moon', undefined, 'theme dark light'),
      cmd(T.gView, T.shortcuts, showShortcuts, 'keyboard', undefined, 'keyboard shortcuts help'),
      ...[50, 75, 100, 125, 150, 200].map((z) => cmd(T.gView, `${T.zoom} ${z}%`, () => ((zoomSel.value = String(z / 100)), zoomSel.dispatchEvent(new Event('change'))), undefined, undefined, 'zoom')),
      cmd(T.gSheets, T.addSheet, addSheet, 'plus', undefined, 'add new sheet'),
      cmd(T.gSheets, `${T.renameSheet}：${name()}`, () => void renameSheet(si), 'pencil', undefined, 'rename sheet'),
      cmd(T.gSheets, `${T.duplicateSheet}：${name()}`, () => duplicateSheet(si), 'copy', undefined, 'duplicate sheet'),
      ...data.sheets.map((sh, i) => cmd(T.gSheets, T.goSheet(sh.name), () => switchSheet(i), 'sheet', i === si ? '●' : undefined, 'go sheet')),
      cmd(T.gFile, T.printPdf, printSheet, 'file', `${MOD}P`, 'print pdf'),
      cmd(T.gFile, T.exportXlsx, () => downloadBlob(exportXlsx(serialize(), wb), `${id}.xlsx`), 'download', undefined, 'export download excel xlsx'),
      cmd(T.gFile, T.rename, () => titleBtn.click(), 'pencil', undefined, 'rename title'),
      cmd(T.gBooks, T.allBooks, () => (location.hash = '#/'), 'arrowLeft', undefined, 'home back'),
    ];
    try {
      const res = await api.list();
      for (const w of res.workbooks) if (w.id !== id) list.push(cmd(T.gBooks, T.openBook(w.title), () => (location.hash = `#/w/${encodeURIComponent(w.id)}`), 'sheet', undefined, `open ${w.id}`));
    } catch {
      // the workbook list is optional here
    }
    commandPalette(list, T.palette, () => editor.focus());
  }

  function showShortcuts() {
    const keys = [
      'Enter / Tab', 'F2', 'Esc', 'F4', `${MOD}Z / ⇧${MOD}Z`, `${MOD}C / X / V`, `⇧${MOD}V`, `${MOD}B / I / U`, `${MOD}D / ${MOD}R`,
      `${MOD}Enter`, `${MOD}; / ⇧${MOD};`, `${MOD}↑↓←→`, `${MOD}A`, `${MOD}F`, `${MOD}K`, '⌫ / Delete', '⇧ + ↑↓←→', '⌥↑ / ⌥↓', '⌃`',
    ];
    const rows = keys.map((k, i) => [k, T.shortcutList[i] ?? ''] as [string, string]);
    const back = el('div', 'dialog-back');
    const box = el('div', 'dialog wide');
    box.append(el('h2', '', T.shortcuts));
    const list = el('dl', 'shortcuts');
    for (const [k, v] of rows) list.append(el('dt', '', k), el('dd', '', v));
    const ok = el('button', 'btn primary', T.ok);
    const actions = el('div', 'dialog-actions');
    actions.append(ok);
    box.append(list, actions);
    back.append(box);
    document.body.append(back);
    ok.focus();
    const close = () => {
      back.remove();
      editor.focus();
    };
    ok.addEventListener('click', close);
    back.addEventListener('mousedown', (e) => e.target === back && close());
    back.addEventListener('keydown', (e) => e.key === 'Escape' && close());
  }

  // ---------- keyboard ----------
  editor.addEventListener('keydown', (e) => onKey(e));
  // Keys pressed while a button (a tab, a toolbar button) has focus still reach the grid.
  const onWindowKey = (e: KeyboardEvent) => {
    const t = e.target as HTMLElement;
    if (t === editor || t.closest('input, textarea, select, .dialog, .menu') || !root.isConnected) return;
    if (e.key.length === 1 && !e.metaKey && !e.ctrlKey) {
      editor.focus();
      return; // the character lands in the editor
    }
    onKey(e);
    editor.focus();
  };
  window.addEventListener('keydown', onWindowKey);
  function onKey(e: KeyboardEvent) {
    if (e.isComposing || e.keyCode === 229) return;
    if (selectedChart && !editing) {
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        return deleteChart(selectedChart);
      }
      if (e.key === 'Escape') return selectChart(null);
      if (e.key.length === 1 || e.key.startsWith('Arrow') || e.key === 'Enter' || e.key === 'Tab') selectChart(null);
    }
    const mod = e.metaKey || e.ctrlKey;
    const k = e.key;
    if (editing) {
      if (!popup.hidden && acItems.length) {
        if (k === 'ArrowDown' || k === 'ArrowUp') {
          e.preventDefault();
          acIndex = (acIndex + (k === 'ArrowDown' ? 1 : -1) + acItems.length) % acItems.length;
          updatePopup();
          return;
        }
        if (k === 'Tab' || (k === 'Enter' && !mod)) {
          e.preventDefault();
          acceptSuggestion();
          return;
        }
      }
      const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (arrows[k] && canPoint()) {
        e.preventDefault();
        const [dr, dc] = arrows[k]!;
        const from = pointer?.at ?? active;
        const at = { row: Math.max(0, from.row + dr), col: Math.max(0, from.col + dc) };
        pointAt(e.shiftKey && pointer ? pointer.anchor : at, at);
        return;
      }
      if (k === 'Enter') {
        e.preventDefault();
        if (mod) {
          // Fill the whole selection with this entry.
          const value = draft();
          const s = selection();
          endEdit();
          change(() => {
            const cells = [];
            for (let r = s.r0; r <= s.r1; r++) for (let c = s.c0; c <= s.c1; c++) cells.push({ row: r, col: c, input: offsetFormula(value, r - active.row, c - active.col) });
            setInputs(cells);
          });
          return;
        }
        commitEdit();
        stepInSelection(e.shiftKey ? -1 : 1, 0);
      } else if (k === 'Tab') {
        e.preventDefault();
        commitEdit();
        stepInSelection(0, e.shiftKey ? -1 : 1);
      } else if (k === 'Escape') {
        e.preventDefault();
        cancelEdit();
      } else if (k === 'F4') {
        e.preventDefault();
        toggleAbsolute();
      } else if (editing.mode === 'enter' && arrows[k] && !e.shiftKey) {
        e.preventDefault();
        commitEdit(arrows[k]);
      }
      return;
    }
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) {
      e.preventDefault();
      if (e.altKey && (k === 'ArrowUp' || k === 'ArrowDown')) return switchSheet(si + (k === 'ArrowDown' ? 1 : -1));
      if (mod) return jump(...arrows[k]!, e.shiftKey);
      return moveActive(...arrows[k]!, e.shiftKey);
    }
    if (k === 'Enter') {
      e.preventDefault();
      return stepInSelection(e.shiftKey ? -1 : 1, 0);
    }
    if (k === 'Tab') {
      e.preventDefault();
      return stepInSelection(0, e.shiftKey ? -1 : 1);
    }
    if (k === 'PageDown' || k === 'PageUp') {
      e.preventDefault();
      const rows = Math.max(1, Math.floor(body.clientHeight / rowH(active.row)) - 1);
      return moveActive(k === 'PageDown' ? rows : -rows, 0, e.shiftKey);
    }
    if (k === 'Home') {
      e.preventDefault();
      return moveTo(mod ? { row: 0, col: 0 } : { row: active.row, col: 0 }, e.shiftKey);
    }
    if (k === 'End' && mod) {
      e.preventDefault();
      const u = wb.used(name());
      return moveTo({ row: Math.max(0, u.rows - 1), col: Math.max(0, u.cols - 1) }, e.shiftKey);
    }
    if (k === ' ' && !mod && !e.shiftKey && !editing && toggleChecks(selection())) {
      e.preventDefault();
      return;
    }
    if (k === '`' && e.ctrlKey) {
      e.preventDefault();
      return toggleFormulas();
    }
    if (k === 'F2') {
      e.preventDefault();
      if (e.shiftKey) return void editNote(active.row, active.col);
      return beginEdit('edit');
    }
    if (k === 'Escape') {
      if (painter) {
        painter = null;
        tb.painter.classList.remove('on');
        grid.classList.remove('painting');
      }
      clip = null;
      trace = null;
      closeMenus();
      return render();
    }
    if (k === 'Delete' || k === 'Backspace') {
      e.preventDefault();
      return clearContents(selection());
    }
    if (mod) {
      const key = k.toLowerCase();
      if (key === 'z') {
        e.preventDefault();
        return e.shiftKey ? redo() : undo();
      }
      if (key === 'y') {
        e.preventDefault();
        return redo();
      }
      if (key === 'b' || key === 'i' || key === 'u') {
        e.preventDefault();
        return toggleStyle(key);
      }
      if (e.shiftKey && (key === '>' || key === '.' || key === '<' || key === ',')) {
        e.preventDefault();
        return stepSize(key === '>' || key === '.' ? 1 : -1);
      }
      if (key === 'a') {
        e.preventDefault();
        active = { row: 0, col: 0 };
        anchor = { row: nRows - 1, col: nCols - 1 };
        return render();
      }
      if (key === 'p') {
        e.preventDefault();
        return printSheet();
      }
      if (key === 'k') {
        e.preventDefault();
        return void openPalette();
      }
      if (key === 'f' || key === 'h') {
        e.preventDefault();
        return openFind();
      }
      if (key === 'd' || key === 'r') {
        e.preventDefault();
        const s = selection();
        if (key === 'd' && s.r1 > s.r0) change(() => fill({ ...s, r1: s.r0 }, s));
        if (key === 'r' && s.c1 > s.c0) change(() => fill({ ...s, c1: s.c0 }, s));
        if (key === 'd' && s.r1 === s.r0 && s.r0 > 0) change(() => fill({ ...s, r0: s.r0 - 1, r1: s.r0 - 1 }, { ...s, r0: s.r0 - 1 }));
        if (key === 'r' && s.c1 === s.c0 && s.c0 > 0) change(() => fill({ ...s, c0: s.c0 - 1, c1: s.c0 - 1 }, { ...s, c0: s.c0 - 1 }));
        return;
      }
      if (k === ';' || k === ':' || e.code === 'Semicolon') {
        e.preventDefault();
        const now = new Date();
        const text = e.shiftKey
          ? `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`
          : formatValue(Math.floor(serialFromJsDate(now)), 'yyyy-mm-dd').text;
        return beginEdit('enter', text);
      }
      if (key === 's') {
        e.preventDefault();
        return void flush();
      }
      if (key === 'c' && e.shiftKey) {
        e.preventDefault();
        return pickUpFormat();
      }
      if (key === 'v' && e.shiftKey && painter) {
        e.preventDefault();
        return applyFormat();
      }
      if (key === 'v' && e.shiftKey) {
        // Paste values: handled on the paste event below.
        pasteValuesNext = true;
      }
      return;
    }
    if (k === '=' || k === '+') {
      // Typing starts a fresh entry; let the character land in the input.
      beginEdit('enter', '');
      return;
    }
  }
  let pasteValuesNext = false;
  /**
   * Excel's AutoComplete: typing text that starts an entry already in this
   * column fills in the rest, selected, so typing on replaces it and Enter keeps it.
   */
  function autoComplete(e: Event) {
    const ie = e as InputEvent;
    if (!editing || editing.mode !== 'enter' || ie.isComposing || ie.inputType?.startsWith('delete')) return;
    const typed = editor.value;
    if (!typed || typed.startsWith('=') || /^[\d.,$%+-]/.test(typed) || editor.selectionStart !== typed.length) return;
    const lower = typed.toLowerCase();
    const seen = new Set<string>();
    const used = wb.used(name());
    for (const step of [-1, 1]) {
      // The block of data around the cell, up and down, as Excel looks.
      for (let r = active.row + step; r >= 0 && r < used.rows; r += step) {
        const v = wb.value(name(), r, active.col);
        if (v === null || v === '') break;
        if (typeof v === 'string' && v.toLowerCase().startsWith(lower) && v.length > typed.length) seen.add(v);
      }
    }
    // Only an unambiguous match completes, as in Excel.
    if (seen.size !== 1) return;
    const [match] = [...seen];
    editor.value = typed + match!.slice(typed.length);
    editor.setSelectionRange(typed.length, editor.value.length);
    finput.value = editor.value;
  }

  editor.addEventListener('input', (e) => {
    if (!editing) {
      const v = editor.value;
      beginEdit('enter', v);
      autoComplete(e);
      return;
    }
    autoComplete(e);
    pointer = pointer && editor.selectionStart === pointer.end ? pointer : null;
    finput.value = editor.value;
    paintOverlay();
    updatePopup();
  });
  editor.addEventListener('compositionstart', () => {
    if (!editing) {
      editing = { mode: 'enter', source: 'cell' };
      editor.value = '';
      editor.classList.remove('idle');
      fillHandle.hidden = true;
      paintOverlay();
    }
  });
  editor.addEventListener('click', () => editing && updatePopup());
  editor.addEventListener('keyup', (e) => {
    if (editing && (e.key === 'ArrowLeft' || e.key === 'ArrowRight' || e.key === 'Home' || e.key === 'End')) updatePopup();
  });
  editor.addEventListener('blur', () => {
    setTimeout(() => {
      if (editing?.source === 'cell' && document.activeElement !== editor && document.activeElement !== finput && !document.querySelector('.dialog-back, .menu-layer')) commitEdit();
    }, 0);
  });
  editor.addEventListener('copy', (e) => {
    if (editing) return;
    e.preventDefault();
    e.clipboardData?.setData('text/plain', copySelection(false));
  });
  editor.addEventListener('cut', (e) => {
    if (editing) return;
    e.preventDefault();
    e.clipboardData?.setData('text/plain', copySelection(true));
  });
  editor.addEventListener('paste', (e) => {
    if (editing) return;
    e.preventDefault();
    pasteText(e.clipboardData?.getData('text/plain') ?? '', pasteValuesNext);
    pasteValuesNext = false;
  });

  // Formula bar
  finput.addEventListener('focus', () => {
    if (!editing) {
      editing = { mode: 'edit', source: 'bar' };
      editor.value = finput.value;
      editor.classList.remove('idle');
      fillHandle.hidden = true;
      paintOverlay();
    } else editing.source = 'bar';
  });
  finput.addEventListener('input', () => {
    editor.value = finput.value;
    pointer = pointer && finput.selectionStart === pointer.end ? pointer : null;
    paintOverlay();
    updatePopup();
  });
  finput.addEventListener('keydown', (e) => {
    if (e.isComposing || e.keyCode === 229) return;
    if (!popup.hidden && acItems.length && (e.key === 'Tab' || e.key === 'Enter' || e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
      e.preventDefault();
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        acIndex = (acIndex + (e.key === 'ArrowDown' ? 1 : -1) + acItems.length) % acItems.length;
        updatePopup();
      } else acceptSuggestion();
      return;
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      commitEdit([1, 0]);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commitEdit([0, 1]);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      cancelEdit();
    } else if (e.key === 'F4') {
      e.preventDefault();
      toggleAbsolute();
    }
  });
  finput.addEventListener('blur', () => {
    setTimeout(() => {
      if (editing?.source === 'bar' && document.activeElement !== finput && document.activeElement !== editor && !document.querySelector('.dialog-back, .menu-layer')) commitEdit();
    }, 0);
  });

  // Name box: go to an address or select a range.
  namebox.addEventListener('focus', () => namebox.select());
  namebox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      const def = Object.entries(wb.namesJSON()).find(([n]) => n.toLowerCase() === namebox.value.trim().toLowerCase());
      if (def) {
        editor.focus(); // first, so the name box shows the range it lands on
        goToName(def[1]);
        return;
      }
      const r = parseRange(namebox.value);
      if (r) {
        moveTo({ row: r.r0, col: r.c0 });
        moveTo({ row: r.r1, col: r.c1 }, true);
      }
      editor.focus();
    } else if (e.key === 'Escape') editor.focus();
  });

  // ---------- mouse ----------
  const cellFromEvent = (e: MouseEvent): Pos => {
    const r = grid.getBoundingClientRect();
    const x = Math.max(0, e.clientX - r.left - HEAD_W);
    const y = Math.max(0, e.clientY - r.top - HEAD_H);
    return { row: search(rowY, sheetY(y), nRows), col: search(colX, sheetX(x), nCols) };
  };
  let drag: null | { kind: 'select' | 'point' | 'fill' | 'rows' | 'cols'; start: Pos; source?: Rect; target?: Rect } = null;
  let resize: null | { kind: 'row' | 'col'; index: number; start: number; size: number; guide: HTMLElement } = null;

  grid.addEventListener('mousedown', (e) => {
    const target = e.target as HTMLElement;
    if (target.closest('.chart')) return;
    if (selectedChart) selectChart(null);
    if (target === editor || target.closest('.findbar')) return;
    const r = grid.getBoundingClientRect();
    const bx = e.clientX - r.left;
    const by = e.clientY - r.top;
    // Clicks on the scrollbars belong to the browser.
    if (bx > HEAD_W + frozenW() + body.clientWidth || by > HEAD_H + frozenH() + body.clientHeight) return;

    // Header edges resize; headers select whole rows and columns.
    if (by < HEAD_H || bx < HEAD_W) {
      e.preventDefault();
      if (by < HEAD_H && bx < HEAD_W) {
        if (editing) commitEdit();
        active = { row: 0, col: 0 };
        anchor = { row: nRows - 1, col: nCols - 1 };
        render();
        return editor.focus();
      }
      const isCol = by < HEAD_H;
      const pos = isCol ? sheetX(bx - HEAD_W) : sheetY(by - HEAD_H);
      const arr = isCol ? colX : rowY;
      const n = isCol ? nCols : nRows;
      const i = search(arr, pos, n);
      const nearEnd = arr[i + 1]! - pos <= 4;
      const nearStart = pos - arr[i]! <= 3 && i > 0;
      if (e.button === 0 && (nearEnd || nearStart)) {
        const index = nearEnd ? i : i - 1;
        if (e.detail === 2) {
          // Double-click a border: fit to the contents.
          const s = selection();
          const many = isCol ? (s.r0 === 0 && s.r1 >= nRows - 1 && index >= s.c0 && index <= s.c1 ? range(s.c0, s.c1) : [index]) : s.c0 === 0 && s.c1 >= nCols - 1 && index >= s.r0 && index <= s.r1 ? range(s.r0, s.r1) : [index];
          change(() => (isCol ? many.forEach((c) => setColWidth([c], fitCol(c))) : many.forEach((rr) => setRowHeight([rr], fitRow(rr)))));
          return;
        }
        const guide = el('div', `resize-guide ${isCol ? 'v' : 'h'}`);
        grid.append(guide);
        resize = { kind: isCol ? 'col' : 'row', index, start: isCol ? e.clientX : e.clientY, size: (isCol ? colX[index + 1]! - colX[index]! : rowY[index + 1]! - rowY[index]!) / zoom, guide };
        moveGuide(e);
        return;
      }
      if (editing) commitEdit();
      const s = selection();
      const wholeAlready = isCol ? s.r0 === 0 && s.r1 >= nRows - 1 && i >= s.c0 && i <= s.c1 : s.c0 === 0 && s.c1 >= nCols - 1 && i >= s.r0 && i <= s.r1;
      if (e.button === 2 && wholeAlready) {
        return headMenu(isCol ? 'col' : 'row', { x: e.clientX, y: e.clientY });
      }
      const startAt = e.shiftKey ? (isCol ? active.col : active.row) : i;
      if (isCol) {
        active = { row: 0, col: startAt };
        anchor = { row: nRows - 1, col: i };
      } else {
        active = { row: startAt, col: 0 };
        anchor = { row: i, col: nCols - 1 };
      }
      render();
      editor.focus();
      if (e.button === 2) return headMenu(isCol ? 'col' : 'row', { x: e.clientX, y: e.clientY });
      drag = { kind: isCol ? 'cols' : 'rows', start: { row: startAt, col: startAt } };
      return;
    }

    // The filter button at the right of a header cell (painted under the scroll layer, so found by position).
    {
      const at = cellFromEvent(e);
      if (e.button === 0 && isFilterHead(at.row, at.col)) {
        const right = viewX(at.col) + (colX[at.col + 1]! - colX[at.col]!);
        if (right - (bx - HEAD_W) <= 20) {
          e.preventDefault();
          if (editing) commitEdit();
          filterMenu(at.col, { x: r.left + HEAD_W + right - 20, y: r.top + HEAD_H + viewY(at.row) + (rowY[at.row + 1]! - rowY[at.row]!) + 2 });
          return;
        }
      }
    }
    if (e.button === 0 && !e.shiftKey && !e.metaKey && !e.ctrlKey && !editing) {
      // The box is painted under the scroll layer, so it is found by position: centred in its cell.
      const at = cellFromEvent(e);
      const cell = isCheck(at.row, at.col) ? viewRect({ r0: at.row, c0: at.col, r1: at.row, c1: at.col }) : null;
      const reach = 10 * zoom;
      const onBox = cell && Math.abs(bx - HEAD_W - (cell.x + cell.w / 2)) <= reach && Math.abs(by - HEAD_H - (cell.y + cell.h / 2)) <= reach;
      if (onBox) {
        e.preventDefault();
        active = at;
        anchor = at;
        toggleChecks({ r0: at.row, c0: at.col, r1: at.row, c1: at.col });
        editor.focus();
        return;
      }
    }
    if (e.button === 0 && (e.metaKey || e.ctrlKey) && !editing) {
      const at = cellFromEvent(e);
      const url = linkAt(at.row, at.col);
      if (url) {
        e.preventDefault();
        openLink(url);
        return;
      }
    }
    if (target === fillHandle) {
      e.preventDefault();
      drag = { kind: 'fill', start: cellFromEvent(e), source: selection() };
      return;
    }

    const at = cellFromEvent(e);
    if (editing && e.button === 0 && canPoint()) {
      e.preventDefault();
      pointer = pointer ?? null;
      pointAt(at, at);
      drag = { kind: 'point', start: at };
      return;
    }
    e.preventDefault();
    if (editing) commitEdit();
    if (e.button === 2) {
      // A right-click inside the selection keeps it.
      if (!inRect(selection(), at.row, at.col)) {
        active = at;
        anchor = { ...at };
        render();
      }
      editor.focus();
      return cellMenu({ x: e.clientX, y: e.clientY });
    }
    const m = mergeOf(at.row, at.col);
    if (e.shiftKey) anchor = at; // extend from the active cell to the click
    else {
      active = m ? { row: m.r0, col: m.c0 } : at;
      anchor = { ...active };
    }
    render();
    editor.focus();
    drag = { kind: 'select', start: e.shiftKey ? active : at };
  });
  const range = (a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
  grid.addEventListener('dblclick', (e) => {
    const r = grid.getBoundingClientRect();
    if (e.clientX - r.left < HEAD_W || e.clientY - r.top < HEAD_H) return;
    if ((e.target as HTMLElement) === fillHandle) {
      // Double-click the fill handle: fill down as far as the column to the left goes.
      const s = selection();
      let end = s.r1;
      const probe = s.c0 > 0 ? s.c0 - 1 : s.c1 + 1;
      while (wb.inputAt(name(), end + 1, probe)) end++;
      if (end > s.r1) change(() => fill(s, { ...s, r1: end }));
      return;
    }
    if (!editing) beginEdit('edit');
  });
  grid.addEventListener('contextmenu', (e) => e.preventDefault());

  function moveGuide(e: MouseEvent) {
    if (!resize) return;
    const delta = (resize.kind === 'col' ? e.clientX : e.clientY) - resize.start;
    const size = Math.max(4, resize.size * zoom + delta);
    if (resize.kind === 'col') resize.guide.style.left = `${HEAD_W + viewX(resize.index) + size - 1}px`;
    else resize.guide.style.top = `${HEAD_H + viewY(resize.index) + size - 1}px`;
    resize.guide.dataset.size = String(Math.round(size / zoom));
  }
  const onMove = (e: MouseEvent) => {
    if (resize) return moveGuide(e);
    if (!drag) return;
    const at = cellFromEvent(e);
    autoScroll(e);
    if (drag.kind === 'select') {
      if (anchor.row !== at.row || anchor.col !== at.col) {
        anchor = at;
        render();
      }
    } else if (drag.kind === 'point') {
      pointAt(drag.start, at);
    } else if (drag.kind === 'cols') {
      active = { row: 0, col: drag.start.col };
      anchor = { row: nRows - 1, col: at.col };
      render();
    } else if (drag.kind === 'rows') {
      active = { row: drag.start.row, col: 0 };
      anchor = { row: at.row, col: nCols - 1 };
      render();
    } else if (drag.kind === 'fill' && drag.source) {
      const s = drag.source;
      const down = at.row > s.r1 ? at.row - s.r1 : at.row < s.r0 ? s.r0 - at.row : 0;
      const side = at.col > s.c1 ? at.col - s.c1 : at.col < s.c0 ? s.c0 - at.col : 0;
      let t: Rect = s;
      if (down >= side && down > 0) t = at.row > s.r1 ? { ...s, r1: at.row } : { ...s, r0: at.row };
      else if (side > 0) t = at.col > s.c1 ? { ...s, c1: at.col } : { ...s, c0: at.col };
      drag.target = t;
      overlay.querySelector('.fill-preview')?.remove();
      const v = viewRect(t);
      if (v && t !== s) {
        const box = el('div', 'fill-preview');
        box.style.cssText = `left:${v.x - 1}px;top:${v.y - 1}px;width:${v.w + 1}px;height:${v.h + 1}px`;
        overlay.prepend(box);
      }
    }
  };
  const onUp = () => {
    stopAutoScroll();
    if (resize) {
      const size = Number(resize.guide.dataset.size);
      resize.guide.remove();
      const { kind, index } = resize;
      resize = null;
      if (Number.isFinite(size) && size > 0) {
        const s = selection();
        const many =
          kind === 'col'
            ? s.r0 === 0 && s.r1 >= nRows - 1 && index >= s.c0 && index <= s.c1
              ? range(s.c0, s.c1)
              : [index]
            : s.c0 === 0 && s.c1 >= nCols - 1 && index >= s.r0 && index <= s.r1
              ? range(s.r0, s.r1)
              : [index];
        change(() => (kind === 'col' ? setColWidth(many, size) : setRowHeight(many, size)));
      }
      return;
    }
    if (drag?.kind === 'select' && painter) {
      drag = null;
      applyFormat();
      return;
    }
    if (drag?.kind === 'fill' && drag.source && drag.target && drag.target !== drag.source) {
      const { source, target } = drag;
      drag = null;
      const shrink = target.r1 - target.r0 < source.r1 - source.r0 || target.c1 - target.c0 < source.c1 - source.c0;
      if (!shrink) {
        change(() => fill(source, target));
        active = { row: target.r0, col: target.c0 };
        anchor = { row: target.r1, col: target.c1 };
        render();
      }
      return;
    }
    drag = null;
    if (editing) (editing.source === 'bar' ? finput : editor).focus();
  };
  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseup', onUp);

  let scrollTimer: ReturnType<typeof setInterval> | null = null;
  let lastMouse: MouseEvent | null = null;
  function autoScroll(e: MouseEvent) {
    lastMouse = e;
    if (scrollTimer) return;
    scrollTimer = setInterval(() => {
      if (!lastMouse || !drag) return stopAutoScroll();
      const r = body.getBoundingClientRect();
      const dx = lastMouse.clientX > r.right ? 20 : lastMouse.clientX < r.left && body.scrollLeft > 0 && drag.kind !== 'rows' ? -20 : 0;
      const dy = lastMouse.clientY > r.bottom ? 20 : lastMouse.clientY < r.top && body.scrollTop > 0 && drag.kind !== 'cols' ? -20 : 0;
      if (dx || dy) {
        body.scrollLeft += dx;
        body.scrollTop += dy;
        onMove(lastMouse);
      }
    }, 40);
  }
  function stopAutoScroll() {
    if (scrollTimer) clearInterval(scrollTimer);
    scrollTimer = null;
  }

  // Scrolling the frozen panes and headers scrolls the body.
  grid.addEventListener(
    'wheel',
    (e) => {
      if (body.contains(e.target as Node)) return;
      e.preventDefault();
      body.scrollLeft += e.deltaX;
      body.scrollTop += e.deltaY;
    },
    { passive: false },
  );
  body.addEventListener('scroll', () => {
    // Grow the sheet as it scrolls toward its end.
    if (body.scrollTop + body.clientHeight > spacer.offsetHeight - 400 && nRows < MAX_ROWS) {
      extraRows += 200;
      layout();
    }
    if (body.scrollLeft + body.clientWidth > spacer.offsetWidth - 300 && nCols < MAX_COLS) {
      extraCols += 10;
      layout();
    }
    render();
    if (!popup.hidden) updatePopup();
  });

  zoomSel.addEventListener('change', () => {
    zoom = Number(zoomSel.value);
    layout();
    scrollIntoView();
    render();
    editor.focus();
  });

  const ro = new ResizeObserver(() => render());
  ro.observe(grid);

  // An agent (or a person in another editor) changed the file: load it, unless this page wrote it.
  const onExternal = async (payload: { id: string }) => {
    if (payload.id !== id) return;
    await reloadFromDisk(false);
  };
  async function reloadFromDisk(conflict: boolean) {
    const loaded = await api.load(id).catch(() => null);
    if (!loaded) return;
    const fresh = loaded.data;
    etag = loaded.etag;
    const text = JSON.stringify(fresh);
    if (text === lastSaved && !conflict) return;
    if (editing) cancelEdit();
    checkpoint();
    const current = name();
    data = fresh;
    if (!data.sheets?.length) data.sheets = [{ name: 'Sheet1', cells: {} }];
    wb = Workbook.fromJSON(data);
    lastSaved = text;
    si = Math.max(0, data.sheets.findIndex((s) => s.name === current));
    titleBtn.textContent = data.title;
    layout();
    renderTabs();
    render();
    paintNotesBadge();
    toast(T.reloaded);
    status.textContent = T.saved;
  }
  import.meta.hot?.on('open-spreadsheet:changed', onExternal);

  const onBeforeUnload = () => {
    if (saveTimer) void flush();
  };
  window.addEventListener('beforeunload', onBeforeUnload);

  teardown = () => {
    if (saveTimer) void flush();
    window.removeEventListener('mousemove', onMove);
    window.removeEventListener('mouseup', onUp);
    window.removeEventListener('beforeunload', onBeforeUnload);
    window.removeEventListener('keydown', onWindowKey);
    ro.disconnect();
    popup.remove();
    notePop.remove();
    closeMenus();
    import.meta.hot?.off?.('open-spreadsheet:changed', onExternal);
    teardown = null;
  };
  const onHash = () => {
    teardown?.();
    window.removeEventListener('hashchange', onHash);
  };
  window.addEventListener('hashchange', onHash);

  layout();
  renderTabs();
  render();
  paintNotesBadge();
  editor.focus();
}
