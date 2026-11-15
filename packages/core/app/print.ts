/**
 * A sheet as a printable page: the used area as one table, with the shown
 * text, styles, merges and borders; hidden rows and columns left out; frozen
 * rows repeated at the top of every printed page (a <thead>); charts after.
 * Pure: the editor passes what each cell shows.
 */
import type { CellStyle } from './types.ts';

export interface PrintCell {
  text: string;
  align: 'left' | 'center' | 'right';
  style?: CellStyle;
  /** Colours a conditional format or a [Red] number format adds. */
  bg?: string;
  color?: string;
}

export interface PrintInput {
  title: string;
  sheetName: string;
  rows: number;
  cols: number;
  cell: (row: number, col: number) => PrintCell | null;
  colWidth: (col: number) => number;
  rowHidden: (row: number) => boolean;
  colHidden: (col: number) => boolean;
  /** Merged blocks, as row/column rectangles. */
  merges: { r0: number; c0: number; r1: number; c1: number }[];
  headerRows: number;
  charts: string[];
  lang: string;
}

const esc = (s: string) => s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
const safeColor = (c?: string) => (c && /^#[0-9a-f]{3,8}$/i.test(c) ? c : undefined);

export function printHtml(p: PrintInput): string {
  const cover = new Map<string, { r0: number; c0: number; r1: number; c1: number }>();
  for (const m of p.merges) for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) cover.set(`${r},${c}`, m);
  const cols = Array.from({ length: p.cols }, (_, c) => c).filter((c) => !p.colHidden(c));
  const row = (r: number) => {
    const tds: string[] = [];
    for (const c of cols) {
      const m = cover.get(`${r},${c}`);
      if (m && (m.r0 !== r || m.c0 !== c)) continue; // inside a merge, drawn by its corner
      const span = m
        ? ` colspan="${cols.filter((x) => x >= m.c0 && x <= m.c1).length}" rowspan="${Array.from({ length: m.r1 - m.r0 + 1 }, (_, i) => m.r0 + i).filter((x) => !p.rowHidden(x)).length}"`
        : '';
      const cell = p.cell(r, c);
      const st = cell?.style;
      const css: string[] = [];
      if (cell) css.push(`text-align:${cell.align}`);
      if (st?.b) css.push('font-weight:600');
      if (st?.i) css.push('font-style:italic');
      // Relative to the printed table's own size, as the grid is to its 13px; a number only, never a string from the file.
      if (typeof st?.size === 'number' && Number.isFinite(st.size) && st.size !== 11) css.push(`font-size:${(st.size / 11).toFixed(3)}em`);
      if (st?.u || st?.s) css.push(`text-decoration:${[st.u ? 'underline' : '', st.s ? 'line-through' : ''].join(' ').trim()}`);
      const color = safeColor(cell?.color) ?? safeColor(st?.color);
      if (color) css.push(`color:${color}`);
      const bg = safeColor(cell?.bg) ?? safeColor(st?.bg);
      if (bg) css.push(`background:${bg}`);
      if (st?.wrap) css.push('white-space:pre-wrap');
      if (st?.valign === 'top' || st?.valign === 'middle' || st?.valign === 'bottom') css.push(`vertical-align:${st.valign}`);
      const b = st?.border;
      if (b) {
        const w = b.thick ? 2 : 1;
        for (const side of ['top', 'right', 'bottom', 'left'] as const) {
          const bc = safeColor(b[side]);
          if (bc) css.push(`border-${side}:${w}px solid ${bc}`);
        }
      }
      // A checkbox prints as a box, ticked when the cell is TRUE.
      const text = st?.check ? (cell?.text === 'TRUE' ? '☑' : '☐') : (cell?.text ?? '');
      tds.push(`<td${span}${css.length ? ` style="${css.join(';')}"` : ''}>${esc(text)}</td>`);
    }
    return `<tr>${tds.join('')}</tr>`;
  };
  const head: string[] = [];
  const body: string[] = [];
  for (let r = 0; r < p.rows; r++) {
    if (p.rowHidden(r)) continue;
    (r < p.headerRows ? head : body).push(row(r));
  }
  const colgroup = `<colgroup>${cols.map((c) => `<col style="width:${Math.round(p.colWidth(c))}px">`).join('')}</colgroup>`;
  return `<!doctype html>
<html lang="${esc(p.lang)}"><head><meta charset="utf-8"><title>${esc(p.title)} — ${esc(p.sheetName)}</title>
<style>
@page { margin: 12mm; }
body { margin: 0; font: 10pt -apple-system, BlinkMacSystemFont, "PingFang TC", "Noto Sans TC", "Segoe UI", sans-serif; color: #1f1f1f; }
h1 { font-size: 13pt; margin: 0 0 8pt; }
table { border-collapse: collapse; table-layout: fixed; }
td { border: 0.5pt solid #d9d9d9; padding: 2pt 4pt; height: 14pt; overflow: hidden; white-space: nowrap; font-variant-numeric: tabular-nums; }
thead { display: table-header-group; }
tr { break-inside: avoid; }
.charts { margin-top: 12pt; display: flex; flex-wrap: wrap; gap: 12pt; }
.charts > div { break-inside: avoid; border: 0.5pt solid #d9d9d9; border-radius: 6pt; }
:root { --sheet: #ffffff; }
.c-title { font: 600 13px sans-serif; fill: #1f1f1f; }
.c-tick, .c-muted { font: 11px sans-serif; fill: #5f6368; }
.c-grid { stroke: #e3e3e3; stroke-width: 1; }
.c-axis { stroke: #5f6368; stroke-width: 1; }
</style></head>
<body><h1>${esc(p.title)} · ${esc(p.sheetName)}</h1>
<table>${colgroup}${head.length ? `<thead>${head.join('')}</thead>` : ''}<tbody>${body.join('')}</tbody></table>
${p.charts.length ? `<div class="charts">${p.charts.map((svg) => `<div>${svg}</div>`).join('')}</div>` : ''}
</body></html>`;
}
