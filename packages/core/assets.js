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
export function saveImage(root, scope, name, bytes) {
  if (!IMAGE.test(String(name ?? ''))) throw new OpsError('bad_name', 'An image is a .svg, .png, .jpg, .gif or .webp file');
  assetPath(root, scope, name);
  const free = freeName(root, scope, name);
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
  if (!['bar', 'line', 'pie'].includes(type)) throw new OpsError('bad_type', 'type is bar, line or pie');
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
  const { chartData, chartFile } = await charts();
  const spec = { id: 'asset', type, range: opts.range, title: opts.title ?? '', x: 0, y: 0, w: 720, h: 405 };
  const svg = chartFile(spec, chartData(cells), Number(opts.width) || 720, Number(opts.height) || 405);
  const scope = opts.scope ?? workbook;
  const wanted = opts.name ? (/\.svg$/i.test(opts.name) ? opts.name : `${opts.name}.svg`) : `${slug(opts.title || `${sheet.name}-${opts.range}`) || 'chart'}.svg`;
  assetPath(root, scope, wanted);
  const name = opts.replace ? wanted : freeName(root, scope, wanted);
  writeAtomic(assetPath(root, scope, name), svg);
  return { scope, name, sheet: sheet.name, range: opts.range, type, bytes: Buffer.byteLength(svg) };
}
