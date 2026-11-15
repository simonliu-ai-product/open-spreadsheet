/**
 * Charts drawn as SVG from a range: bar (grouped), line and pie. Hand-written,
 * like open-dashboard's: one y-axis, zero in the domain, a fixed categorical
 * palette, ticks from one formatter, text in ink colours.
 */
import { CellError, formatValue, type Scalar } from '@open-spreadsheet/engine';
import type { ChartSpec } from './types.ts';

export type { ChartSpec, ChartType } from './types.ts';

export interface ChartData {
  categories: string[];
  series: { name: string; values: (number | null)[] }[];
}

export const PALETTE = ['#2a78d6', '#e8710a', '#1e8e3e', '#d93025', '#9334e6', '#12b5cb', '#c5221f', '#7f6000'];

const text = (v: Scalar) => (v === null ? '' : v instanceof CellError ? v.code : typeof v === 'number' ? formatValue(v, '').text : String(v));

/**
 * Reads a block of cells the way Excel's chart wizard does: if the top-left
 * cell is empty or the first column is text, the first column holds the
 * categories; a first row of text holds the series names.
 */
/** Whether a block's first row holds series names and its first column categories. */
export function chartShape(cells: Scalar[][]): { headerRow: boolean; labelCol: boolean } {
  const cols = cells[0]?.length ?? 0;
  const isText = (v: Scalar) => typeof v === 'string' && v !== '';
  const headerRow = Boolean(cells[0]?.slice(1).some(isText)) || (cols === 1 && isText(cells[0]?.[0] ?? null));
  const labelCol = cols > 1 && cells.slice(headerRow ? 1 : 0).some((r) => typeof r[0] !== 'number');
  return { headerRow, labelCol };
}

export function chartData(cells: Scalar[][]): ChartData {
  if (!cells.length || !cells[0]?.length) return { categories: [], series: [] };
  const cols = cells[0]!.length;
  const { headerRow, labelCol } = chartShape(cells);
  const r0 = headerRow ? 1 : 0;
  const c0 = labelCol ? 1 : 0;
  const categories = cells.slice(r0).map((r, i) => (labelCol ? text(r[0] ?? null) : String(i + 1)));
  const series = [];
  for (let c = c0; c < cols; c++) {
    series.push({
      name: headerRow ? text(cells[0]![c] ?? null) : `Series ${c - c0 + 1}`,
      values: cells.slice(r0).map((r) => (typeof r[c] === 'number' ? (r[c] as number) : null)),
    });
  }
  return { categories, series };
}

/** Round axis steps: 1, 2 or 5 times a power of ten. */
export function niceTicks(min: number, max: number, count = 5): number[] {
  if (min === max) max = min + 1;
  const raw = (max - min) / count;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 5, 10].map((m) => m * mag).find((s) => s >= raw) ?? raw;
  // From the step at or below the minimum to the step at or above the maximum, so every mark fits.
  const start = Math.floor(min / step) * step;
  const end = Math.ceil(max / step - 1e-9) * step;
  const out: number[] = [];
  for (let v = start; v <= end + step * 1e-9; v += step) out.push(Number.parseFloat(v.toPrecision(12)));
  return out;
}
function tickFormat(ticks: number[]): (n: number) => string {
  const max = Math.max(...ticks.map(Math.abs));
  if (max >= 1e6) return (n) => (n === 0 ? '0' : `${Number.parseFloat((n / 1e6).toPrecision(3))}M`);
  if (max >= 1e4) return (n) => (n === 0 ? '0' : `${Number.parseFloat((n / 1e3).toPrecision(3))}K`);
  return (n) => formatValue(n, '#,##0.##').text.replace(/\.$/, '');
}
const esc = (s: string) => s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');

export function chartSvg(spec: ChartSpec, data: ChartData, w: number, h: number): string {
  const titleH = spec.title ? 26 : 8;
  const legendH = data.series.length > 1 || spec.type === 'pie' ? 22 : 0;
  const title = spec.title ? `<text x="${w / 2}" y="18" text-anchor="middle" class="c-title">${esc(spec.title)}</text>` : '';
  const empty = !data.series.length || data.series.every((s) => s.values.every((v) => v === null));
  if (empty) return `<svg viewBox="0 0 ${w} ${h}" width="${w}" height="${h}">${title}<text x="${w / 2}" y="${h / 2}" text-anchor="middle" class="c-muted">—</text></svg>`;

  if (spec.type === 'pie') {
    const s = data.series[0]!;
    const items = s.values.map((v, i) => ({ v: Math.max(0, v ?? 0), name: data.categories[i] ?? '' })).filter((x) => x.v > 0);
    const total = items.reduce((a, b) => a + b.v, 0);
    const r = Math.max(10, Math.min(w, h - titleH - legendH) / 2 - 10);
    const cx = w / 2;
    const cy = titleH + (h - titleH - legendH) / 2;
    let angle = -Math.PI / 2;
    const slices = items
      .map((it, i) => {
        const a = (it.v / total) * Math.PI * 2;
        const x0 = cx + r * Math.cos(angle);
        const y0 = cy + r * Math.sin(angle);
        angle += a;
        const x1 = cx + r * Math.cos(angle);
        const y1 = cy + r * Math.sin(angle);
        const path = items.length === 1 ? `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0` : `M${cx},${cy}L${x0},${y0}A${r},${r} 0 ${a > Math.PI ? 1 : 0},1 ${x1},${y1}Z`;
        return `<path d="${path}" fill="${PALETTE[i % PALETTE.length]}" stroke="var(--sheet)" stroke-width="2"><title>${esc(it.name)}: ${esc(formatValue(it.v, '#,##0.##').text.replace(/\.$/, ''))} (${Math.round((it.v / total) * 100)}%)</title></path>`;
      })
      .join('');
    return `<svg viewBox="0 0 ${w} ${h}" width="${w}" height="${h}">${title}${slices}${legend(items.map((it) => it.name), w, h)}</svg>`;
  }

  const padL = 44;
  const padR = 12;
  const top = titleH + 4;
  const bottom = h - legendH - 22;
  const plotW = Math.max(10, w - padL - padR);
  const plotH = Math.max(10, bottom - top);
  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
  const ticks = niceTicks(Math.min(0, ...all), Math.max(0, ...all));
  const lo = ticks[0]!;
  const hi = ticks[ticks.length - 1]!;
  const y = (v: number) => bottom - ((v - lo) / (hi - lo)) * plotH;
  const fmt = tickFormat(ticks);
  const n = data.categories.length;
  const band = plotW / Math.max(1, n);
  const grid = ticks.map((t) => `<line x1="${padL}" x2="${w - padR}" y1="${y(t)}" y2="${y(t)}" class="c-grid"/><text x="${padL - 6}" y="${y(t) + 4}" text-anchor="end" class="c-tick">${esc(fmt(t))}</text>`).join('');
  const every = Math.max(1, Math.ceil((n * 48) / plotW));
  const xLabels = data.categories
    .map((c, i) => (i % every ? '' : `<text x="${padL + band * (i + 0.5)}" y="${bottom + 15}" text-anchor="middle" class="c-tick">${esc(c.length > 10 ? `${c.slice(0, 9)}…` : c)}</text>`))
    .join('');
  let marks = '';
  if (spec.type === 'bar') {
    const k = data.series.length;
    const barW = Math.min(24, (band * 0.72) / k);
    data.series.forEach((s, si) => {
      s.values.forEach((v, i) => {
        if (v === null) return;
        const x = padL + band * i + band / 2 - (barW * k) / 2 + si * barW;
        const y0 = y(Math.max(0, v));
        const hgt = Math.max(1, Math.abs(y(v) - y(0)));
        marks += `<rect x="${x + 1}" y="${y0}" width="${Math.max(1, barW - 2)}" height="${hgt}" rx="2" fill="${PALETTE[si % PALETTE.length]}"><title>${esc(s.name)} · ${esc(data.categories[i] ?? '')}: ${esc(fmt(v))}</title></rect>`;
      });
    });
  } else {
    data.series.forEach((s, si) => {
      const pts = s.values.map((v, i) => (v === null ? null : [padL + band * (i + 0.5), y(v)] as const));
      const d = pts.reduce((acc, p, i) => (p ? `${acc}${acc && pts[i - 1] ? 'L' : 'M'}${p[0]},${p[1]}` : acc), '');
      marks += `<path d="${d}" fill="none" stroke="${PALETTE[si % PALETTE.length]}" stroke-width="2" stroke-linejoin="round"/>`;
      pts.forEach((p, i) => {
        if (p) marks += `<circle cx="${p[0]}" cy="${p[1]}" r="3" fill="${PALETTE[si % PALETTE.length]}"><title>${esc(s.name)} · ${esc(data.categories[i] ?? '')}: ${esc(fmt(s.values[i]!))}</title></circle>`;
      });
    });
  }
  const zero = `<line x1="${padL}" x2="${w - padR}" y1="${y(0)}" y2="${y(0)}" class="c-axis"/>`;
  return `<svg viewBox="0 0 ${w} ${h}" width="${w}" height="${h}">${title}${grid}${zero}${marks}${xLabels}${legendH ? legend(data.series.map((s) => s.name), w, h) : ''}</svg>`;
}

function legend(names: string[], w: number, h: number): string {
  const items = names.slice(0, 8);
  // Wide (CJK) characters take about twice the room of Latin ones at 11px.
  const textWidth = (n: string) => [...n.slice(0, 14)].reduce((w, ch) => w + (/[\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/.test(ch) ? 11.5 : 6.5), 0);
  const widths = items.map((n) => 26 + textWidth(n));
  const total = widths.reduce((a, b) => a + b, 0);
  let x = Math.max(4, (w - total) / 2);
  return items
    .map((name, i) => {
      const out = `<rect x="${x}" y="${h - 15}" width="9" height="9" rx="2" fill="${PALETTE[i % PALETTE.length]}"/><text x="${x + 13}" y="${h - 7}" class="c-tick">${esc(name.length > 14 ? `${name.slice(0, 13)}…` : name)}</text>`;
      x += widths[i]!;
      return out;
    })
    .join('');
}
