/**
 * Charts drawn as SVG from a range: columns (grouped or stacked), horizontal
 * bars, lines, areas, pie and doughnut, scatter, and combo. Hand-written,
 * like open-dashboard's: one y-axis, zero in the domain, a fixed categorical
 * palette, ticks from one formatter, text in ink colours.
 */
import { CellError, formatValue, type Scalar } from '@open-spreadsheet/engine';
import type { ChartSpec, ChartType } from './types.ts';

export const CHART_TYPES: ChartType[] = ['bar', 'barH', 'stacked', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo'];

export type { ChartSpec, ChartType } from './types.ts';

export interface ChartData {
  categories: string[];
  series: { name: string; values: (number | null)[] }[];
}

export const PALETTE = ['#2a78d6', '#e8710a', '#1e8e3e', '#d93025', '#9334e6', '#12b5cb', '#c5221f', '#7f6000'];

/**
 * How a chart file looks: series colours, background, text, grid and type.
 * A few come built in; a workspace's own live in assets/chart-styles.json so
 * every workbook, teammate and agent draws with the same house style.
 */
export interface ChartStyle {
  id: string;
  name: string;
  colors: string[];
  background: string;
  text: string;
  muted: string;
  grid: string;
  /** Grid lines drawn at all. */
  gridLines: boolean;
  font: 'sans' | 'serif' | 'mono' | 'rounded';
  /** Corner radius of bars, 0–8. */
  radius: number;
}

export const CHART_FONTS: Record<ChartStyle['font'], string> = {
  sans: '-apple-system,BlinkMacSystemFont,"PingFang TC","Noto Sans TC","Segoe UI",sans-serif',
  serif: '"Noto Serif TC","Songti TC",Georgia,"Times New Roman",serif',
  mono: '"SF Mono",Menlo,Consolas,"Noto Sans Mono CJK TC",monospace',
  rounded: '"SF Pro Rounded","Hiragino Maru Gothic ProN","Arial Rounded MT Bold",-apple-system,sans-serif',
};

export const BUILTIN_STYLES: ChartStyle[] = [
  { id: 'default', name: 'Default', colors: PALETTE, background: '#ffffff', text: '#1f1f1f', muted: '#5f6368', grid: '#e8eaed', gridLines: true, font: 'sans', radius: 2 },
  { id: 'ocean', name: 'Ocean', colors: ['#0b3c5d', '#1d7ab8', '#4fb0e0', '#8fd0ee', '#328cc1', '#17577e', '#a9d6e5', '#5c7f99'], background: '#ffffff', text: '#0b2233', muted: '#52687a', grid: '#e3edf3', gridLines: true, font: 'sans', radius: 3 },
  { id: 'warm', name: 'Warm', colors: ['#c2410c', '#ea8a2e', '#f2c14e', '#9a3412', '#d9775b', '#7c2d12', '#e7b98a', '#b45309'], background: '#fffaf3', text: '#3b2417', muted: '#8a6a55', grid: '#f1e4d3', gridLines: true, font: 'serif', radius: 0 },
  { id: 'pastel', name: 'Pastel', colors: ['#7aa6e8', '#f4a6a0', '#8fd3b6', '#f6cf7d', '#b9a4e6', '#86d0dc', '#f2b3d6', '#a9c18a'], background: '#ffffff', text: '#2e3440', muted: '#6b7280', grid: '#eef0f3', gridLines: true, font: 'rounded', radius: 6 },
  { id: 'mono', name: 'Mono', colors: ['#111111', '#555555', '#8c8c8c', '#bdbdbd', '#333333', '#707070', '#a6a6a6', '#d4d4d4'], background: '#ffffff', text: '#111111', muted: '#666666', grid: '#ebebeb', gridLines: false, font: 'mono', radius: 0 },
  { id: 'night', name: 'Night', colors: ['#7cc4ff', '#ffb86b', '#7ee2a8', '#ff7a90', '#c4a7ff', '#5ee0e0', '#ffd866', '#a0a8b8'], background: '#15181e', text: '#eef1f6', muted: '#9aa3b2', grid: '#2a2f39', gridLines: true, font: 'sans', radius: 2 },
];
export const DEFAULT_STYLE = BUILTIN_STYLES[0]!;

const HEX = /^#[0-9a-f]{6}$/i;
/** A style read from a file, every field checked; anything missing or wrong falls back to Default's. */
export function cleanStyle(v: unknown): ChartStyle | null {
  if (!v || typeof v !== 'object') return null;
  const o = v as Record<string, unknown>;
  if (typeof o.id !== 'string' || !/^[a-z0-9][a-z0-9-]{0,40}$/.test(o.id)) return null;
  const hex = (x: unknown, d: string) => (typeof x === 'string' && HEX.test(x) ? x.toLowerCase() : d);
  const colors = Array.isArray(o.colors) ? o.colors.filter((c): c is string => typeof c === 'string' && HEX.test(c)).slice(0, 12) : [];
  const d = DEFAULT_STYLE;
  return {
    id: o.id,
    name: typeof o.name === 'string' && o.name.trim() ? o.name.trim().slice(0, 40) : o.id,
    colors: colors.length ? colors.map((c) => c.toLowerCase()) : d.colors,
    background: hex(o.background, d.background),
    text: hex(o.text, d.text),
    muted: hex(o.muted, d.muted),
    grid: hex(o.grid, d.grid),
    gridLines: o.gridLines !== false,
    font: typeof o.font === 'string' && o.font in CHART_FONTS ? (o.font as ChartStyle['font']) : 'sans',
    radius: typeof o.radius === 'number' && Number.isFinite(o.radius) ? Math.max(0, Math.min(8, Math.round(o.radius))) : d.radius,
  };
}

/** The style a chart names, among the workspace's own and the built-in ones; Default when it is gone. */
export function findStyle(id: string | undefined, own: ChartStyle[] = []): ChartStyle {
  return own.find((s) => s.id === id) ?? BUILTIN_STYLES.find((s) => s.id === id) ?? DEFAULT_STYLE;
}

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

export function chartSvg(spec: ChartSpec, data: ChartData, w: number, h: number, style: ChartStyle = DEFAULT_STYLE): string {
  const type = spec.type;
  const round = type === 'pie' || type === 'doughnut';
  const scatter = type === 'scatter';
  // Scatter reads its first numeric column as x: the series are the rest.
  const xs = scatter && data.series.length > 1 ? data.series[0]!.values : null;
  const series = xs ? data.series.slice(1) : data.series;
  const titleH = spec.title ? 26 : 8;
  const legendH = series.length > 1 || round ? 22 : 0;
  const title = spec.title ? `<text x="${w / 2}" y="18" text-anchor="middle" class="c-title">${esc(spec.title)}</text>` : '';
  const svg = (body: string) => `<svg viewBox="0 0 ${w} ${h}" width="${w}" height="${h}">${title}${body}</svg>`;
  const color = (i: number) => style.colors[i % style.colors.length]!;
  const empty = !series.length || series.every((s) => s.values.every((v) => v === null));
  if (empty) return svg(`<text x="${w / 2}" y="${h / 2}" text-anchor="middle" class="c-muted">—</text>`);
  const tip = (name: string, cat: string, v: string) => `<title>${esc(name)}${cat ? ` · ${esc(cat)}` : ''}: ${esc(v)}</title>`;

  if (round) {
    const s = series[0]!;
    const items = s.values.map((v, i) => ({ v: Math.max(0, v ?? 0), name: data.categories[i] ?? '' })).filter((x) => x.v > 0);
    const total = items.reduce((a, b) => a + b.v, 0);
    const r = Math.max(10, Math.min(w, h - titleH - legendH) / 2 - 10);
    const inner = type === 'doughnut' ? r * 0.58 : 0;
    const cx = w / 2;
    const cy = titleH + (h - titleH - legendH) / 2;
    let angle = -Math.PI / 2;
    const at = (rad: number, a: number) => `${cx + rad * Math.cos(a)},${cy + rad * Math.sin(a)}`;
    const slices = items
      .map((it, i) => {
        const a = (it.v / total) * Math.PI * 2;
        const a0 = angle;
        angle += a;
        const large = a > Math.PI ? 1 : 0;
        let path: string;
        if (items.length === 1) {
          path = `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0` + (inner ? `M${cx - inner},${cy}a${inner},${inner} 0 1,1 ${2 * inner},0a${inner},${inner} 0 1,1 ${-2 * inner},0` : '');
        } else if (inner) {
          path = `M${at(r, a0)}A${r},${r} 0 ${large},1 ${at(r, angle)}L${at(inner, angle)}A${inner},${inner} 0 ${large},0 ${at(inner, a0)}Z`;
        } else path = `M${cx},${cy}L${at(r, a0)}A${r},${r} 0 ${large},1 ${at(r, angle)}Z`;
        return `<path d="${path}" fill="${color(i)}" fill-rule="evenodd" stroke="var(--sheet)" stroke-width="2">${tip(it.name, '', `${formatValue(it.v, '#,##0.##').text.replace(/\.$/, '')} (${Math.round((it.v / total) * 100)}%)`)}</path>`;
      })
      .join('');
    const centre = inner ? `<text x="${cx}" y="${cy + 6}" text-anchor="middle" class="c-total">${esc(tickFormat([total])(total))}</text>` : '';
    return svg(`${slices}${centre}${legend(items.map((it) => it.name), w, h, color)}`);
  }

  const horizontal = type === 'barH';
  const combo = type === 'combo' && series.length > 1;
  const left = combo ? series.slice(0, 1) : series;
  const right = combo ? series.slice(1) : [];
  const n = data.categories.length;
  const nums = (list: typeof series) => list.flatMap((s) => s.values.filter((v): v is number => v !== null));
  // Stacked columns reach the sum of each category, above and below zero.
  const stackExtent = () => {
    let lo = 0;
    let hi = 0;
    for (let i = 0; i < n; i++) {
      const pos = series.reduce((a, s) => a + Math.max(0, s.values[i] ?? 0), 0);
      const neg = series.reduce((a, s) => a + Math.min(0, s.values[i] ?? 0), 0);
      hi = Math.max(hi, pos);
      lo = Math.min(lo, neg);
    }
    return [lo, hi];
  };
  const [vLo, vHi] = type === 'stacked' ? stackExtent() : [Math.min(0, ...nums(left)), Math.max(0, ...nums(left))];
  const ticks = scatter ? niceTicks(Math.min(...nums(series)), Math.max(...nums(series))) : niceTicks(vLo!, vHi!);
  const fmt = tickFormat(ticks);
  const lo = ticks[0]!;
  const hi = ticks[ticks.length - 1]!;
  const ticks2 = combo ? niceTicks(Math.min(0, ...nums(right)), Math.max(0, ...nums(right))) : [];
  const fmt2 = combo ? tickFormat(ticks2) : fmt;

  // Horizontal bars: categories down the left, values along the bottom.
  if (horizontal) {
    const labelW = Math.min(140, Math.max(40, ...data.categories.map((c) => textWidth(c) + 10)));
    const padL = labelW + 6;
    const padR = 16;
    const top = titleH + 6;
    const bottom = h - legendH - 22;
    const plotW = Math.max(10, w - padL - padR);
    const band = (bottom - top) / Math.max(1, n);
    const x = (v: number) => padL + ((v - lo) / (hi - lo)) * plotW;
    const k = series.length;
    const barH = Math.min(22, (band * 0.72) / k);
    const grid = ticks.map((t) => `<line x1="${x(t)}" x2="${x(t)}" y1="${top}" y2="${bottom}" class="c-grid"/><text x="${x(t)}" y="${bottom + 15}" text-anchor="middle" class="c-tick">${esc(fmt(t))}</text>`).join('');
    const labels = data.categories.map((c, i) => `<text x="${padL - 8}" y="${top + band * (i + 0.5) + 4}" text-anchor="end" class="c-tick">${esc(c.length > 16 ? `${c.slice(0, 15)}…` : c)}</text>`).join('');
    let marks = '';
    series.forEach((s, si) =>
      s.values.forEach((v, i) => {
        if (v === null) return;
        const yy = top + band * i + band / 2 - (barH * k) / 2 + si * barH;
        marks += `<rect x="${Math.min(x(0), x(v))}" y="${yy + 1}" width="${Math.max(1, Math.abs(x(v) - x(0)))}" height="${Math.max(1, barH - 2)}" rx="${style.radius}" fill="${color(si)}">${tip(s.name, data.categories[i] ?? '', fmt(v))}</rect>`;
      }),
    );
    const axis = `<line x1="${x(0)}" x2="${x(0)}" y1="${top}" y2="${bottom}" class="c-axis"/>`;
    return svg(`${grid}${axis}${marks}${labels}${legendH ? legend(series.map((s) => s.name), w, h, color) : ''}`);
  }

  const padL = 46;
  const padR = combo ? 46 : 14;
  const top = titleH + 4;
  const bottom = h - legendH - 22;
  const plotW = Math.max(10, w - padL - padR);
  const plotH = Math.max(10, bottom - top);
  const y = (v: number) => bottom - ((v - lo) / (hi - lo)) * plotH;
  const lo2 = ticks2[0] ?? 0;
  const hi2 = ticks2[ticks2.length - 1] ?? 1;
  const y2 = (v: number) => bottom - ((v - lo2) / (hi2 - lo2)) * plotH;
  const band = plotW / Math.max(1, n);
  const grid = ticks.map((t) => `<line x1="${padL}" x2="${w - padR}" y1="${y(t)}" y2="${y(t)}" class="c-grid"/><text x="${padL - 6}" y="${y(t) + 4}" text-anchor="end" class="c-tick">${esc(fmt(t))}</text>`).join('');
  const rightAxis = combo ? ticks2.map((t) => `<text x="${w - padR + 6}" y="${y2(t) + 4}" class="c-tick">${esc(fmt2(t))}</text>`).join('') : '';
  let marks = '';
  let xLabels = '';

  if (scatter) {
    // x is the first column when it is numbers, else the row's position.
    const xv = xs ?? data.categories.map((_, i) => i + 1);
    const xn = xv.filter((v): v is number => v !== null);
    const xt = niceTicks(Math.min(...xn), Math.max(...xn));
    const xfmt = tickFormat(xt);
    const xx = (v: number) => padL + ((v - xt[0]!) / (xt[xt.length - 1]! - xt[0]!)) * plotW;
    xLabels = xt.map((t) => `<line x1="${xx(t)}" x2="${xx(t)}" y1="${top}" y2="${bottom}" class="c-grid"/><text x="${xx(t)}" y="${bottom + 15}" text-anchor="middle" class="c-tick">${esc(xfmt(t))}</text>`).join('');
    series.forEach((s, si) =>
      s.values.forEach((v, i) => {
        const xi = xv[i];
        if (v === null || xi === null || xi === undefined) return;
        marks += `<circle cx="${xx(xi)}" cy="${y(v)}" r="4" fill="${color(si)}" fill-opacity="0.8">${tip(s.name, xfmt(xi), fmt(v))}</circle>`;
      }),
    );
    return svg(`${grid}${xLabels}${marks}${legendH ? legend(series.map((s) => s.name), w, h, color) : ''}`);
  }

  const every = Math.max(1, Math.ceil((n * 48) / plotW));
  xLabels = data.categories
    .map((c, i) => (i % every ? '' : `<text x="${padL + band * (i + 0.5)}" y="${bottom + 15}" text-anchor="middle" class="c-tick">${esc(c.length > 10 ? `${c.slice(0, 9)}…` : c)}</text>`))
    .join('');

  const columns = (list: typeof series, offset: number) => {
    const k = list.length;
    const barW = Math.min(28, (band * 0.72) / k);
    list.forEach((s, si) =>
      s.values.forEach((v, i) => {
        if (v === null) return;
        const x = padL + band * i + band / 2 - (barW * k) / 2 + si * barW;
        marks += `<rect x="${x + 1}" y="${y(Math.max(0, v))}" width="${Math.max(1, barW - 2)}" height="${Math.max(1, Math.abs(y(v) - y(0)))}" rx="${style.radius}" fill="${color(si + offset)}">${tip(s.name, data.categories[i] ?? '', fmt(v))}</rect>`;
      }),
    );
  };
  const lines = (list: typeof series, offset: number, yy: (v: number) => number, f: (n: number) => string, fill: boolean) => {
    list.forEach((s, si) => {
      const c = color(si + offset);
      const pts = s.values.map((v, i) => (v === null ? null : ([padL + band * (i + 0.5), yy(v)] as const)));
      const d = pts.reduce((acc, p, i) => (p ? `${acc}${acc && pts[i - 1] ? 'L' : 'M'}${p[0]},${p[1]}` : acc), '');
      if (fill) {
        const run = pts.filter((p): p is readonly [number, number] => p !== null);
        if (run.length) marks += `<path d="M${run[0]![0]},${yy(Math.max(lo, 0))}${run.map((p) => `L${p[0]},${p[1]}`).join('')}L${run[run.length - 1]![0]},${yy(Math.max(lo, 0))}Z" fill="${c}" fill-opacity="0.16"/>`;
      }
      marks += `<path d="${d}" fill="none" stroke="${c}" stroke-width="2" stroke-linejoin="round"/>`;
      pts.forEach((p, i) => {
        if (p) marks += `<circle cx="${p[0]}" cy="${p[1]}" r="3" fill="${c}">${tip(s.name, data.categories[i] ?? '', f(s.values[i]!))}</circle>`;
      });
    });
  };

  if (type === 'bar') columns(series, 0);
  else if (type === 'stacked') {
    const k = 1;
    const barW = Math.min(36, band * 0.62 * k);
    for (let i = 0; i < n; i++) {
      let pos = 0;
      let neg = 0;
      series.forEach((s, si) => {
        const v = s.values[i];
        if (v === null || v === undefined || v === 0) return;
        const from = v > 0 ? pos : neg;
        const to = from + v;
        if (v > 0) pos = to;
        else neg = to;
        const x = padL + band * (i + 0.5) - barW / 2;
        marks += `<rect x="${x}" y="${Math.min(y(from), y(to))}" width="${barW}" height="${Math.max(1, Math.abs(y(to) - y(from)))}" fill="${color(si)}">${tip(s.name, data.categories[i] ?? '', fmt(v))}</rect>`;
      });
    }
  } else if (type === 'combo') {
    columns(left, 0);
    lines(right, 1, combo ? y2 : y, fmt2, false);
  } else lines(series, 0, y, fmt, type === 'area');

  const zero = `<line x1="${padL}" x2="${w - padR}" y1="${y(Math.max(lo, Math.min(0, hi)))}" y2="${y(Math.max(lo, Math.min(0, hi)))}" class="c-axis"/>`;
  return svg(`${grid}${rightAxis}${zero}${marks}${xLabels}${legendH ? legend(series.map((s) => s.name), w, h, color) : ''}`);
}

/** Wide (CJK) characters take about twice the room of Latin ones at 11px. */
function textWidth(n: string): number {
  return [...n.slice(0, 16)].reduce((w, ch) => w + (/[ᄀ-ᇿ⺀-꓏가-힯豈-﫿＀-￯]/.test(ch) ? 11.5 : 6.5), 0);
}

function legend(names: string[], w: number, h: number, color: (i: number) => string): string {
  const items = names.slice(0, 8);
  const widths = items.map((n) => 26 + textWidth(n));
  const total = widths.reduce((a, b) => a + b, 0);
  let x = Math.max(4, (w - total) / 2);
  return items
    .map((name, i) => {
      const out = `<rect x="${x}" y="${h - 15}" width="9" height="9" rx="2" fill="${color(i)}"/><text x="${x + 13}" y="${h - 7}" class="c-tick">${esc(name.length > 14 ? `${name.slice(0, 13)}…` : name)}</text>`;
      x += widths[i]!;
      return out;
    })
    .join('');
}

/**
 * A chart as a file of its own — for the Assets folder, a note or another
 * program: the SVG with its namespace, a white background and the colours the
 * page would give it written in, since no stylesheet comes with it.
 */
export interface ChartSource {
  sheet: string;
  range: string;
  type: ChartSpec['type'];
  title?: string;
  /** A chart style's id; Default when absent. */
  style?: string;
}

/** Where a chart file was drawn from, read back from its <metadata>; null for any other SVG. */
export function chartSourceOf(svg: string): ChartSource | null {
  const m = /<metadata id="open-spreadsheet-chart">([^<]*)<\/metadata>/.exec(svg);
  if (!m) return null;
  try {
    const v = JSON.parse(m[1]!.replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&quot;', '"').replaceAll('&amp;', '&')) as ChartSource;
    return typeof v.sheet === 'string' && typeof v.range === 'string' && CHART_TYPES.includes(v.type) ? v : null;
  } catch {
    return null;
  }
}

export function chartFile(spec: ChartSpec, data: ChartData, w = 720, h = 405, source?: ChartSource, look: ChartStyle = DEFAULT_STYLE): string {
  const css =
    `<style>text{font-family:${CHART_FONTS[look.font]}}` +
    `.c-title{font-weight:600;font-size:13px;fill:${look.text}}.c-tick{font-size:11px;fill:${look.muted}}.c-muted{font-size:14px;fill:${look.muted}}` +
    `.c-grid{stroke:${look.gridLines ? look.grid : 'none'};stroke-width:1}.c-axis{stroke:${look.muted};stroke-opacity:.6;stroke-width:1}.c-total{font-weight:650;font-size:22px;fill:${look.text}}</style>`;
  return chartSvg(spec, data, w, h, look)
    .replace(/^<svg /, '<svg xmlns="http://www.w3.org/2000/svg" ')
    .replace('>', `>${source ? `<metadata id="open-spreadsheet-chart">${esc(JSON.stringify(source))}</metadata>` : ''}${css}<rect width="100%" height="100%" fill="${look.background}"/>`)
    .replaceAll('var(--sheet)', look.background);
}
