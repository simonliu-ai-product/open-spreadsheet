// Assets: images and Markdown notes that sit beside a workbook
// (sheets/<id>/assets/) or are shared by the whole workspace (assets/).
// Charts are drawn here from a range — the editor's renderer — so a person, an
// agent and `apply-notes` all get the same picture. Notes are context: agents
// read them before working on a workbook.
import { existsSync, mkdirSync, readdirSync, readFileSync, renameSync, statSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { Workbook } from '@open-spreadsheet/engine';
import { ID, OpsError, parseRangeText, readWorkbook, workbookFile } from './ops.js';

export const SHARED = 'shared';
/** One file name, no folders: letters (any script), digits, space, _ . ( ) -, and a kind we serve. */
const NAME = /^[\p{L}\p{N}_][\p{L}\p{N}_ .()-]{0,118}\.(svg|png|jpe?g|gif|webp|md)$/iu;
export const IMAGE = /\.(svg|png|jpe?g|gif|webp)$/i;
export const MIME = { svg: 'image/svg+xml', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', md: 'text/markdown; charset=utf-8' };

// The editor's renderer: its source in this repo, the stripped copy once installed.
const charts = () => import(new URL(import.meta.url.includes('/node_modules/') ? './dist/charts.js' : './app/charts.ts', import.meta.url).href);

/** Where a scope's assets live: a workbook's folder, or the workspace's shared one. */
export function assetsDir(root, scope) {
  if (scope === SHARED) return path.join(root, 'assets');
  if (!ID.test(String(scope ?? '')) || !existsSync(workbookFile(root, scope))) throw new OpsError('not_found', `No workbook ${scope} (or use scope "${SHARED}")`);
  return path.join(path.dirname(workbookFile(root, scope)), 'assets');
}

export function assetPath(root, scope, name) {
  if (!NAME.test(String(name ?? '')) || name.includes('..')) {
    throw new OpsError('bad_name', `"${name}" is not an asset name: one file name ending in .md, .svg, .png, .jpg, .gif or .webp`);
  }
  return path.join(assetsDir(root, scope), name);
}

const firstLine = (text) => (text.split('\n').find((l) => l.trim()) ?? '').replace(/^#+\s*/, '').slice(0, 120);

/** A scope's images and notes, newest first; a note carries its first line as a title. */
export function listAssets(root, scope) {
  const dir = assetsDir(root, scope);
  if (!existsSync(dir)) return [];
  return readdirSync(dir)
    .filter((n) => NAME.test(n))
    .map((name) => {
      const st = statSync(path.join(dir, name));
      const kind = IMAGE.test(name) ? 'image' : 'note';
      return { name, kind, size: st.size, updatedAt: st.mtime.toISOString(), ...(kind === 'note' ? { title: firstLine(readFileSync(path.join(dir, name), 'utf8')) } : {}) };
    })
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

/** A note's text, or an image's bytes. */
export function readAsset(root, scope, name) {
  const file = assetPath(root, scope, name);
  if (!existsSync(file)) throw new OpsError('not_found', `No asset ${name} in ${scope}`);
  return IMAGE.test(name) ? readFileSync(file) : readFileSync(file, 'utf8');
}

function writeAtomic(file, content) {
  mkdirSync(path.dirname(file), { recursive: true });
  const tmp = `${file}.${process.pid}.${Date.now().toString(36)}.tmp`;
  writeFileSync(tmp, content);
  renameSync(tmp, file);
}

/** The name itself when free, else name-2, name-3… before the extension. */
function freeName(root, scope, name) {
  const dir = assetsDir(root, scope);
  const dot = name.lastIndexOf('.');
  const stem = name.slice(0, dot);
  const ext = name.slice(dot);
  let candidate = name;
  for (let n = 2; existsSync(path.join(dir, candidate)); n++) candidate = `${stem}-${n}${ext}`;
  return candidate;
}

export function writeNote(root, scope, name, text) {
  if (!/\.md$/i.test(String(name ?? ''))) throw new OpsError('bad_name', 'A note is a .md file');
  const file = assetPath(root, scope, name);
  writeAtomic(file, String(text ?? ''));
  return { scope, name };
}

/** Stores an image (an upload); a taken name gets a number rather than replacing the file. */
export function saveImage(root, scope, name, bytes, opts = {}) {
  if (!IMAGE.test(String(name ?? ''))) throw new OpsError('bad_name', 'An image is a .svg, .png, .jpg, .gif or .webp file');
  assetPath(root, scope, name);
  // A redrawn chart keeps its name; anything else new gets a number instead of replacing a file.
  const free = opts.replace ? name : freeName(root, scope, name);
  writeAtomic(assetPath(root, scope, free), bytes);
  return { scope, name: free };
}

/** Moves an asset to the scope's .trash folder; nothing is erased. */
export function trashAsset(root, scope, name) {
  const file = assetPath(root, scope, name);
  if (!existsSync(file)) throw new OpsError('not_found', `No asset ${name} in ${scope}`);
  const trash = path.join(path.dirname(file), '.trash');
  mkdirSync(trash, { recursive: true });
  const to = path.join(trash, `${Date.now()}-${name}`);
  renameSync(file, to);
  return { scope, name, trashedTo: path.relative(root, to) };
}

/** Renames an image or note in place, keeping its kind; a star follows it. */
export function renameAsset(root, scope, name, to) {
  const from = assetPath(root, scope, name);
  const dest = assetPath(root, scope, to);
  if (!existsSync(from)) throw new OpsError('not_found', `No asset ${name} in ${scope}`);
  const ext = (n) => n.split('.').pop().toLowerCase();
  if (ext(name) !== ext(to)) throw new OpsError('bad_name', `Keep the .${ext(name)} ending.`);
  // Only a change of case may land on the "same" file (macOS folders ignore case).
  if (to !== name && existsSync(dest) && to.toLowerCase() !== name.toLowerCase()) throw new OpsError('exists', `${to} is taken in ${scope}`);
  renameSync(from, dest);
  const favs = readFavorites(root);
  if (favs.some((f) => f.scope === scope && f.name === name)) {
    writeAtomic(favouritesFile(root), `${JSON.stringify(favs.map((f) => (f.scope === scope && f.name === name ? { scope, name: to } : f)), null, 2)}\n`);
  }
  return { scope, from: name, name: to };
}

const slug = (text) =>
  String(text)
    .normalize('NFKC')
    .replace(/[^\p{L}\p{N}]+/gu, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);

/**
 * Draws a chart from a workbook range and saves it as an SVG in the workbook's
 * assets (or the shared ones): the same picture the editor shows.
 */
export async function renderChart(root, workbook, opts) {
  const { data } = readWorkbook(root, workbook);
  const sheet = opts.sheet ? data.sheets.find((s) => s.name.toLowerCase() === String(opts.sheet).toLowerCase()) : data.sheets[0];
  if (!sheet) throw new OpsError('no_sheet', `No sheet "${opts.sheet}". Sheets: ${data.sheets.map((s) => s.name).join(', ')}`);
  const type = opts.type ?? 'bar';
  if (!['bar', 'barH', 'stacked', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo'].includes(type)) throw new OpsError('bad_type', 'type is one of bar, barH, stacked, line, area, pie, doughnut, scatter, combo');
  const r = parseRangeText(opts.range);
  if (!r) throw new OpsError('bad_range', `Not a range: ${opts.range}`);
  if ((r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) > 5000) throw new OpsError('too_big', 'Chart at most 5,000 cells; pass a smaller range.');
  const wb = Workbook.fromJSON(data);
  const cells = [];
  for (let row = r.r0; row <= r.r1; row++) {
    const line = [];
    for (let col = r.c0; col <= r.c1; col++) line.push(wb.value(sheet.name, row, col));
    cells.push(line);
  }
  const { chartData, chartFile, findStyle, BUILTIN_STYLES } = await charts();
  const own = await readChartStyles(root);
  if (opts.style && !own.styles.some((s) => s.id === opts.style) && !BUILTIN_STYLES.some((s) => s.id === opts.style)) {
    throw new OpsError('bad_style', `No chart style "${opts.style}". Styles: ${[...BUILTIN_STYLES, ...own.styles].map((s) => s.id).join(', ')}`);
  }
  const spec = { id: 'asset', type, range: opts.range, title: opts.title ?? '', x: 0, y: 0, w: 720, h: 405 };
  // The source rides in the file, so the editor can redraw it when the numbers change.
  // No style named: it follows the workspace's default, now and when redrawn.
  const source = { sheet: sheet.name, range: String(opts.range).toUpperCase(), type, ...(opts.title ? { title: String(opts.title) } : {}), ...(opts.style ? { style: String(opts.style) } : {}) };
  const svg = chartFile(spec, chartData(cells), Number(opts.width) || 720, Number(opts.height) || 405, source, findStyle(opts.style ?? own.default, own.styles));
  const scope = opts.scope ?? workbook;
  const wanted = opts.name ? (/\.svg$/i.test(opts.name) ? opts.name : `${opts.name}.svg`) : `${slug(opts.title || `${sheet.name}-${opts.range}`) || 'chart'}.svg`;
  assetPath(root, scope, wanted);
  const name = opts.replace ? wanted : freeName(root, scope, wanted);
  writeAtomic(assetPath(root, scope, name), svg);
  return { scope, name, sheet: sheet.name, range: opts.range, type, bytes: Buffer.byteLength(svg) };
}

// ---------- chart styles ----------
// The workspace's own looks for chart files, beside the built-in ones. Shared
// (assets/chart-styles.json, in git) so teammates and agents draw alike.
const stylesFile = (root) => path.join(root, 'assets', 'chart-styles.json');

/** { default, styles }: the style new charts follow, and the workspace's own styles. */
export async function readChartStyles(root) {
  const { cleanStyle, BUILTIN_STYLES } = await charts();
  let raw = {};
  try {
    raw = JSON.parse(readFileSync(stylesFile(root), 'utf8'));
  } catch {
    // none yet
  }
  const seen = new Set(BUILTIN_STYLES.map((s) => s.id));
  const styles = [];
  for (const s of Array.isArray(raw?.styles) ? raw.styles : []) {
    const c = cleanStyle(s);
    if (c && !seen.has(c.id)) {
      seen.add(c.id);
      styles.push(c);
    }
  }
  const def = typeof raw?.default === 'string' && seen.has(raw.default) ? raw.default : 'default';
  return { default: def, styles };
}

/** Replaces the workspace's styles and default; bad entries are refused, not dropped. */
export async function writeChartStyles(root, next) {
  const { cleanStyle, BUILTIN_STYLES } = await charts();
  const builtin = new Set(BUILTIN_STYLES.map((s) => s.id));
  const ids = new Set();
  const styles = (Array.isArray(next?.styles) ? next.styles : []).map((s) => {
    const c = cleanStyle(s);
    if (!c) throw new OpsError('bad_style', 'A style needs an id of lowercase letters, numbers and dashes.');
    if (builtin.has(c.id) || ids.has(c.id)) throw new OpsError('bad_style', `The style id "${c.id}" is taken.`);
    ids.add(c.id);
    return c;
  });
  const def = typeof next?.default === 'string' ? next.default : 'default';
  if (!builtin.has(def) && !ids.has(def)) throw new OpsError('bad_style', `No chart style "${def}" to make the default.`);
  const out = { default: def, styles };
  writeAtomic(stylesFile(root), `${JSON.stringify(out, null, 2)}
`);
  return out;
}

// ---------- favourite notes ----------
// A person's pick, not part of the workbook: kept in .open-spreadsheet/ (git-ignored).
const favouritesFile = (root) => path.join(root, '.open-spreadsheet', 'favorites.json');

/** The starred notes, as { scope, name }, in the order they were starred. */
export function readFavorites(root) {
  try {
    const list = JSON.parse(readFileSync(favouritesFile(root), 'utf8'));
    return Array.isArray(list) ? list.filter((f) => f && typeof f.scope === 'string' && typeof f.name === 'string') : [];
  } catch {
    return [];
  }
}

export function setFavorite(root, scope, name, on) {
  assetPath(root, scope, name);
  const rest = readFavorites(root).filter((f) => !(f.scope === scope && f.name === name));
  const next = on ? [...rest, { scope, name }] : rest;
  writeAtomic(favouritesFile(root), `${JSON.stringify(next, null, 2)}\n`);
  return next;
}
