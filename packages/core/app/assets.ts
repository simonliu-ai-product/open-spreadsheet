// The Assets and Notes views beside a workbook's sheet: chart images and
// pictures (sheets/<id>/assets/*.svg|png|…) and Markdown notes
// (sheets/<id>/assets/*.md), each with the workspace's shared ones (assets/).
// Both are a library: a sidebar (Home, Recent, Favorites, All, search) and,
// on the right, the overview, one item, or — for Assets — the chart styles.
// A chart image carries where it came from, so it is redrawn from the current
// numbers (and in its style) whenever it is shown, its file kept in step for agents.
import { BUILTIN_STYLES, CHART_TYPES, type ChartSource, type ChartStyle, chartFile, chartSourceOf, findStyle } from './charts.ts';
import { locale, T } from './i18n.ts';
import { icon } from './icons.ts';
import { type ImageResolver, mdToHtml } from './markdown.ts';
import { createNoteEditor } from './note-editor.ts';
import type { ChartType } from './types.ts';
import { confirmDialog, downloadBlob, el, iconButton, pickFile, promptDialog, showMenu, toast } from './ui.ts';

export const SHARED = 'shared';

export interface AssetEntry {
  name: string;
  kind: 'image' | 'note';
  size: number;
  updatedAt: string;
  title?: string;
}

const api = (path: string, params: Record<string, string>) => `/__api/assets${path}?${new URLSearchParams(params).toString()}`;

async function failure(res: Response): Promise<Error> {
  try {
    return new Error(((await res.json()) as { error?: string }).error ?? res.statusText);
  } catch {
    return new Error(res.statusText);
  }
}

export async function listAssets(scope: string): Promise<AssetEntry[]> {
  const res = await fetch(api('', { scope }));
  if (!res.ok) return [];
  return ((await res.json()) as { assets: AssetEntry[] }).assets;
}

/** Stores an image; a taken name comes back numbered unless `replace`. */
export async function uploadAsset(scope: string, name: string, body: Blob | string, type = 'image/svg+xml', replace = false): Promise<string> {
  const res = await fetch(api('/upload', { scope, name, ...(replace ? { replace: '1' } : {}) }), { method: 'POST', headers: { 'content-type': type }, body });
  if (!res.ok) throw await failure(res);
  return ((await res.json()) as { name: string }).name;
}

async function readText(scope: string, name: string): Promise<string> {
  const res = await fetch(api('/file', { scope, name, t: String(Date.now()) }));
  return res.ok ? res.text() : '';
}

async function writeNote(scope: string, name: string, text: string): Promise<void> {
  const res = await fetch('/__api/assets/note', { method: 'PUT', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ scope, name, text }) });
  if (!res.ok) throw await failure(res);
}

async function removeAsset(scope: string, name: string): Promise<void> {
  const res = await fetch(api('', { scope, name }), { method: 'DELETE' });
  if (!res.ok) throw await failure(res);
}

/** A file name from a title: letters of any script kept, the rest dashes. */
export const fileStem = (title: string, fallback: string) =>
  title
    .normalize('NFKC')
    .replace(/[^\p{L}\p{N}]+/gu, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60) || fallback;

const stem = (name: string) => name.replace(/\.[a-z]+$/i, '');

export const chartLabel = (t: ChartType) =>
  ({ bar: T.chartBar, barH: T.chartBarH, stacked: T.chartStacked, line: T.chartLine, area: T.chartArea, pie: T.chartPie, doughnut: T.chartDoughnut, scatter: T.chartScatter, combo: T.chartCombo })[t];

/** An image as a PNG at twice its size, for slides and documents. */
async function toPng(src: string): Promise<Blob> {
  const img = new Image();
  img.src = src;
  await img.decode();
  const scale = 2;
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round((img.naturalWidth || 720) * scale));
  canvas.height = Math.max(1, Math.round((img.naturalHeight || 405) * scale));
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  return new Promise((resolve, reject) => canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('PNG'))), 'image/png'));
}
const svgUrl = (svg: string) => `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;

function deleteButton(scope: string, name: string, after: () => void): HTMLButtonElement {
  return iconButton('trash', T.delete, async () => {
    if (!(await confirmDialog(T.deleteAssetTitle(stem(name)), '', T.delete, true))) return;
    try {
      await removeAsset(scope, name);
      after();
    } catch (e) {
      toast((e as Error).message);
    }
  });
}

/** Download as PNG (2×) or the original file; an SVG offers both. */
function downloadButton(scope: string, name: string, src: () => string): HTMLButtonElement {
  const dl = iconButton('download', T.download, () => {
    const original = () => {
      const link = el('a');
      link.href = api('/file', { scope, name, download: '1' });
      link.click();
    };
    const png = async () => {
      try {
        downloadBlob(await toPng(src()), `${stem(name)}.png`);
      } catch (e) {
        toast((e as Error).message);
      }
    };
    if (!/\.svg$/i.test(name)) return original();
    showMenu(
      [
        { label: 'PNG', icon: 'download', action: () => void png() },
        { label: 'SVG', icon: 'download', action: original },
      ],
      dl,
    );
  });
  return dl;
}

/** Large preview over the page; click or Esc closes it. */
function lightbox(src: string, caption: string) {
  const back = el('div', 'lightbox');
  const fig = el('figure', '');
  const img = el('img');
  img.src = src;
  img.alt = caption;
  fig.append(img, el('figcaption', '', caption));
  back.append(fig);
  const close = () => {
    back.remove();
    window.removeEventListener('keydown', onKey, true);
  };
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      close();
    }
  };
  back.addEventListener('click', close);
  window.addEventListener('keydown', onKey, true);
  document.body.append(back);
}

// ---------- the library both views share ----------

interface Entry {
  scope: string;
  name: string;
  title: string;
  updatedAt: string;
}
const keyOf = (n: { scope: string; name: string }) => `${n.scope}/${n.name}`;

async function loadFavourites(): Promise<Set<string>> {
  try {
    const res = await fetch('/__api/favorites');
    if (!res.ok) return new Set();
    return new Set(((await res.json()) as { favorites: { scope: string; name: string }[] }).favorites.map(keyOf));
  } catch {
    return new Set();
  }
}
/** Stars or unstars; the new set, or null when it failed (and said so). */
async function putFavourite(scope: string, name: string, on: boolean): Promise<Set<string> | null> {
  const res = await fetch('/__api/favorites', { method: 'PUT', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ scope, name, on }) });
  if (!res.ok) {
    toast((await failure(res)).message);
    return null;
  }
  return new Set(((await res.json()) as { favorites: { scope: string; name: string }[] }).favorites.map(keyOf));
}

function starButton(on: boolean, toggle: () => void): HTMLButtonElement {
  const b = iconButton(on ? 'starFill' : 'star', on ? T.unfavorite : T.favorite, toggle);
  b.classList.add('lib-star');
  b.classList.toggle('on', on);
  return b;
}

const ago = (iso: string) => {
  const s = (Date.now() - Date.parse(iso)) / 1000;
  if (s < 60) return T.justNow;
  if (s < 3600) return T.minutesAgo(Math.floor(s / 60));
  if (s < 86400) return T.hoursAgo(Math.floor(s / 3600));
  return new Date(iso).toLocaleDateString(locale(), { month: 'short', day: 'numeric' });
};

/** The piece of `text` around the first `q`, with `q` marked; null when it is not there. */
function snippet(text: string, q: string): HTMLElement | null {
  const at = text.toLowerCase().indexOf(q);
  if (at < 0) return null;
  const from = Math.max(0, at - 18);
  const piece = text.slice(from, at + q.length + 40).replace(/[#*`>\n]+/g, ' ');
  const i = piece.toLowerCase().indexOf(q);
  const out = el('span', 'lib-item-meta');
  out.append(`${from ? '…' : ''}${piece.slice(0, Math.max(0, i))}`, el('mark', '', piece.slice(i, i + q.length)), piece.slice(i + q.length));
  return out;
}

type Nav = 'home' | 'recent' | 'fav' | 'all' | 'styles';

/** Head, search and the Home / Recent / Favorites / All nav, with the list under them. */
function librarySide(o: { title: string; actions: HTMLElement[]; searchLabel: string; store: string; extra?: readonly (readonly [Nav, string, string])[]; onNav: (n: Nav) => void; onSearch: () => void }) {
  const side = el('aside', 'lib-side');
  const head = el('div', 'lib-head');
  head.append(el('h2', '', o.title), el('div', 'grow'), ...o.actions);
  const box = el('label', 'lib-search');
  const search = el('input');
  search.type = 'search';
  search.placeholder = o.searchLabel;
  search.setAttribute('aria-label', o.searchLabel);
  box.append(icon('search', 14), search);
  const navEl = el('nav', 'lib-nav');
  const known: Nav[] = ['home', 'recent', 'fav', 'all', ...(o.extra ?? []).map(([k]) => k)];
  let nav: Nav = 'home';
  try {
    const v = localStorage.getItem(o.store) as Nav | null;
    if (v && known.includes(v)) nav = v;
  } catch {
    // remembered for this page only
  }
  const tab = ([key, ic, label]: readonly [Nav, string, string]) => {
    const b = el('button', 'lib-tab', '');
    b.type = 'button';
    b.dataset.nav = key;
    b.append(icon(ic, 15), el('span', '', label));
    b.addEventListener('click', () => select(key));
    return b;
  };
  const main: (readonly [Nav, string, string])[] = [
    ['home', 'home', T.home],
    ['recent', 'clock', T.notesRecent],
    ['fav', 'star', T.notesFavorites],
    ['all', 'list', T.notesAll],
  ];
  const tabs = [...main, ...(o.extra ?? [])].map(tab);
  navEl.append(...tabs.slice(0, 4));
  if (o.extra?.length) navEl.append(el('div', 'lib-divider'), ...tabs.slice(4));
  const list = el('div', 'lib-list');
  side.append(head, box, navEl, list);
  const paint = () => {
    for (const t of tabs) t.setAttribute('aria-selected', String(t.dataset.nav === nav));
  };
  function select(n: Nav, quiet = false) {
    nav = n;
    try {
      localStorage.setItem(o.store, n);
    } catch {
      // remembered for this page only
    }
    paint();
    if (!quiet) o.onNav(n);
  }
  let timer = 0;
  search.addEventListener('input', () => {
    clearTimeout(timer);
    timer = window.setTimeout(o.onSearch, 120);
  });
  search.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && search.value) {
      e.stopPropagation();
      search.value = '';
      o.onSearch();
    }
  });
  paint();
  return {
    el: side,
    list,
    query: () => search.value.trim().toLowerCase(),
    get nav() {
      return nav;
    },
    select,
  };
}

function libItem(n: Entry, fav: boolean, meta: string | HTMLElement | null, onClick: () => void, thumb?: string): HTMLButtonElement {
  const b = el('button', 'lib-item', '');
  b.type = 'button';
  b.dataset.key = keyOf(n);
  b.dataset.name = n.name;
  if (thumb) {
    const img = el('img', 'lib-thumb');
    img.src = thumb;
    img.alt = '';
    img.loading = 'lazy';
    b.append(img);
  }
  const text = el('span', 'lib-item-text');
  const top = el('span', 'lib-item-top');
  top.append(el('span', 'lib-item-title', n.title));
  if (fav) top.append(icon('starFill', 11));
  text.append(top);
  if (meta) text.append(typeof meta === 'string' ? el('span', 'lib-item-meta', meta) : meta);
  b.append(text);
  b.addEventListener('click', onClick);
  return b;
}

function primaryButton(label: string, ic: string, onClick: () => void, primary = true): HTMLButtonElement {
  const b = el('button', primary ? 'btn primary' : 'btn', '');
  b.type = 'button';
  b.append(icon(ic, 14), el('span', '', label));
  b.addEventListener('click', onClick);
  return b;
}

const where = (scope: string) => (scope === SHARED ? T.assetsShared : T.assetsThisBook);
const newestFirst = <E extends Entry>(list: E[]) => [...list].sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));

// ---------- chart styles ----------

interface Styles {
  default: string;
  styles: ChartStyle[];
}
async function loadStyles(): Promise<Styles> {
  try {
    const res = await fetch('/__api/chart-styles');
    if (res.ok) return (await res.json()) as Styles;
  } catch {
    // the built-in ones still work
  }
  return { default: 'default', styles: [] };
}
async function saveStyles(next: Styles): Promise<Styles | null> {
  const res = await fetch('/__api/chart-styles', { method: 'PUT', headers: { 'content-type': 'application/json' }, body: JSON.stringify(next) });
  if (!res.ok) {
    toast((await failure(res)).message);
    return null;
  }
  return (await res.json()) as Styles;
}
const builtin = (id: string) => BUILTIN_STYLES.some((s) => s.id === id);
export const styleName = (s: ChartStyle) => (builtin(s.id) ? T.builtinStyle(s.id) : s.name);

const SAMPLE = {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'A', values: [42, 58, 51, 74] },
    { name: 'B', values: [30, 36, 44, 49] },
    { name: 'C', values: [18, 25, 21, 33] },
  ],
};
/** A small chart in a style, for its swatch and preview. */
const sampleChart = (look: ChartStyle, type: ChartType = 'bar', title = '') =>
  chartFile({ id: 'sample', type, range: 'A1', ...(title ? { title } : {}), x: 0, y: 0, w: 720, h: 405 }, type === 'doughnut' ? { categories: ['A', 'B', 'C', 'D'], series: [{ name: 'S', values: [44, 26, 18, 12] }] } : SAMPLE, 720, 405, undefined, look);

// ---------- Assets ----------

/** What the editor lends the Assets view to make and redraw charts. */
export interface ChartHost {
  sheets(): string[];
  /** The sheet and range a new chart starts from: the selection. */
  selection(): { sheet: string; range: string };
  /** The chart drawn from the numbers now, in a style; '' when its sheet or range is gone. */
  draw(source: ChartSource, look: ChartStyle): string;
}

const RANGE = /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/;

interface ImageEntry extends Entry {
  /** Where a chart image came from; null for any other picture. */
  source: ChartSource | null;
  text: string;
}

export function createAssetsView(opts: { bookId: string; charts: ChartHost }) {
  const page = el('section', 'lib-view asset-view');
  page.dataset.view = 'assets';
  const side = librarySide({
    title: T.viewAssets,
    actions: [iconButton('upload', T.uploadImage, () => void pickAndUpload()), iconButton('plus', T.newChart, () => void chartDialog())],
    searchLabel: T.searchAssets,
    store: 'open-spreadsheet:assets-nav',
    extra: [['styles', 'palette', T.chartStyles]],
    onNav: (n) => {
      if (n === 'home') current = null;
      void paint();
    },
    onSearch: () => void paint(),
  });
  const main = el('div', 'lib-main asset-main');
  page.append(side.el, main);

  let images: ImageEntry[] = [];
  let favourites = new Set<string>();
  let styles: Styles = { default: 'default', styles: [] };
  let current: { scope: string; name: string } | null = null;
  /** The style open in the editor; while it is, refreshes leave the right side alone. */
  let editing: string | null = null;
  let fresh = '';
  const allStyles = () => [...BUILTIN_STYLES, ...styles.styles];
  const lookFor = (source: ChartSource) => findStyle(source.style ?? styles.default, styles.styles);
  const fileUrl = (e: Entry) => api('/file', { scope: e.scope, name: e.name, v: e.updatedAt });

  /** What an image shows now: a chart redrawn from the numbers (its file kept in step), else the file. */
  function liveSrc(e: ImageEntry): string {
    if (!e.source) return fileUrl(e);
    const drawn = opts.charts.draw(e.source, lookFor(e.source)) || e.text;
    if (drawn !== e.text) {
      e.text = drawn;
      void uploadAsset(e.scope, e.name, drawn, 'image/svg+xml', true).catch(() => {});
    }
    return svgUrl(drawn);
  }
  const metaOf = (e: ImageEntry) => (e.source ? `${chartLabel(e.source.type)} · ${e.source.sheet}!${e.source.range}` : where(e.scope));

  async function upload(files: File[]) {
    let last = '';
    for (const file of files) {
      if (!/^image\/(png|jpeg|gif|webp|svg\+xml)$/.test(file.type)) continue;
      try {
        const dot = file.name.lastIndexOf('.');
        const ext = dot > 0 ? file.name.slice(dot).toLowerCase() : '.png';
        last = await uploadAsset(opts.bookId, `${fileStem(dot > 0 ? file.name.slice(0, dot) : file.name, 'image')}${ext}`, file, file.type);
      } catch (e) {
        toast((e as Error).message);
      }
    }
    if (last) await refresh({ scope: opts.bookId, name: last });
  }
  async function pickAndUpload() {
    const file = await pickFile('image/png,image/jpeg,image/gif,image/webp,image/svg+xml');
    if (file) await upload([file]);
  }
  page.addEventListener('dragover', (e) => {
    if (!e.dataTransfer?.types.includes('Files')) return;
    e.preventDefault();
    page.classList.add('dropping');
  });
  page.addEventListener('dragleave', (e) => {
    if (e.target === page) page.classList.remove('dropping');
  });
  page.addEventListener('drop', (e) => {
    e.preventDefault();
    page.classList.remove('dropping');
    void upload([...(e.dataTransfer?.files ?? [])]);
  });

  /** New chart (no `edit`) or a change to one; returns the saved name. */
  async function chartDialog(edit?: { name: string; source: ChartSource }): Promise<string | null> {
    if (!styles.styles.length && styles.default === 'default') styles = await loadStyles();
    const start = edit?.source ?? { ...opts.charts.selection(), type: 'bar' as const };
    const v = await promptDialog({
      title: edit ? T.editChart : T.newChart,
      fields: [
        { name: 'title', label: T.title, value: start.title ?? '' },
        { name: 'type', label: T.chartType, options: CHART_TYPES.map((t) => ({ value: t, label: chartLabel(t) })), value: start.type },
        { name: 'sheet', label: T.sheetLabel, options: opts.charts.sheets().map((s) => ({ value: s, label: s })), value: start.sheet },
        { name: 'range', label: T.rangeLabel, value: start.range, mono: true, validate: (x) => (RANGE.test(x.trim()) ? null : T.rangeLabel) },
        {
          name: 'style',
          label: T.styleLabel,
          options: [{ value: '', label: T.followDefault(styleName(findStyle(styles.default, styles.styles))) }, ...allStyles().map((s) => ({ value: s.id, label: styleName(s) }))],
          value: start.style ?? '',
        },
      ],
      ok: T.save,
    });
    if (!v) return null;
    const source: ChartSource = {
      sheet: v.sheet!,
      range: v.range!.trim().toUpperCase().replaceAll('$', ''),
      type: v.type as ChartSource['type'],
      ...(v.title!.trim() ? { title: v.title!.trim() } : {}),
      ...(v.style ? { style: v.style } : {}),
    };
    const svg = opts.charts.draw(source, lookFor(source));
    if (!svg) {
      toast(T.rangeLabel);
      return null;
    }
    try {
      const name = edit ? await uploadAsset(opts.bookId, edit.name, svg, 'image/svg+xml', true) : await uploadAsset(opts.bookId, `${fileStem(source.title ?? '', `${source.sheet}-${source.range}`)}.svg`, svg);
      fresh = `${opts.bookId}/${name}`;
      if (!page.hidden) await refresh({ scope: opts.bookId, name });
      return name;
    } catch (e) {
      toast((e as Error).message);
      return null;
    }
  }

  async function toggleFavourite(e: Entry) {
    const next = await putFavourite(e.scope, e.name, !favourites.has(keyOf(e)));
    if (!next) return;
    favourites = next;
    await paint();
  }

  // ----- the right side -----

  function card(e: ImageEntry): HTMLElement {
    const c = el('figure', 'asset-card');
    c.dataset.name = e.name;
    if (keyOf(e) === fresh) c.classList.add('fresh');
    const src = liveSrc(e);
    const frame = el('button', 'asset-frame');
    frame.type = 'button';
    frame.title = e.title;
    const img = el('img');
    img.src = src;
    img.alt = e.title;
    img.loading = 'lazy';
    frame.append(img);
    frame.addEventListener('click', () => open(e));
    const tools = el('div', 'asset-tools');
    tools.append(
      starButton(favourites.has(keyOf(e)), () => void toggleFavourite(e)),
      downloadButton(e.scope, e.name, () => src),
      deleteButton(e.scope, e.name, () => void refresh()),
    );
    const cap = el('figcaption', '');
    const t = el('span', 'asset-title', e.title);
    if (favourites.has(keyOf(e))) t.prepend(icon('starFill', 12));
    cap.append(t, el('span', 'asset-meta', metaOf(e)));
    c.append(frame, tools, cap);
    return c;
  }

  function header(title: string, ...actions: HTMLElement[]) {
    const bar = el('div', 'lib-page-head');
    bar.append(el('h1', '', title), el('div', 'grow'), ...actions);
    return bar;
  }

  function showHome() {
    editing = null;
    const body = el('div', 'lib-page');
    body.append(header(T.viewAssets, primaryButton(T.uploadImage, 'upload', () => void pickAndUpload(), false), primaryButton(T.newChart, 'chart', () => void chartDialog())));
    if (!images.length) {
      const empty = el('div', 'lib-empty-state');
      empty.append(icon('chart', 28), el('h3', '', T.noAssetsYet), primaryButton(T.newChart, 'chart', () => void chartDialog()));
      body.append(empty);
      main.replaceChildren(body);
      return;
    }
    const section = (label: string, list: ImageEntry[], scope?: string) => {
      if (!list.length) return;
      const s = el('section', 'asset-section');
      if (scope) s.dataset.scope = scope;
      const grid = el('div', 'asset-grid');
      grid.append(...list.map(card));
      s.append(el('h3', '', label), grid);
      body.append(s);
    };
    section(T.notesFavorites, images.filter((e) => favourites.has(keyOf(e))));
    section(T.assetsThisBook, images.filter((e) => e.scope === opts.bookId), opts.bookId);
    section(T.assetsShared, images.filter((e) => e.scope === SHARED), SHARED);
    main.replaceChildren(body);
    main.querySelector('.asset-card.fresh')?.scrollIntoView({ block: 'nearest' });
  }

  function showDetail(e: ImageEntry) {
    editing = null;
    let src = liveSrc(e);
    const body = el('div', 'lib-page asset-detail');
    // The title on top, edited in place: Enter or leaving it renames the file (and retitles a chart).
    const head = el('div', 'lib-detail-head asset-detail-head');
    const title = el('input', 'asset-title-input');
    title.value = e.title;
    title.spellcheck = false;
    title.setAttribute('aria-label', T.title);
    title.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter') {
        ev.preventDefault();
        title.blur();
      } else if (ev.key === 'Escape') {
        ev.stopPropagation();
        title.value = e.title;
        title.blur();
      }
    });
    title.addEventListener('blur', () => void retitle(e, title));
    head.append(title, el('div', 'grow'));
    if (e.source && e.scope === opts.bookId) {
      const s = e.source;
      // Its style, changed in place: redrawn and saved at once.
      const pick = el('select', 'input select asset-style-pick');
      pick.title = T.styleLabel;
      pick.append(new Option(T.followDefault(styleName(findStyle(styles.default, styles.styles))), ''), ...allStyles().map((st) => new Option(styleName(st), st.id)));
      pick.value = s.style ?? '';
      pick.addEventListener('change', async () => {
        const next: ChartSource = { ...s, ...(pick.value ? { style: pick.value } : {}) };
        if (!pick.value) delete next.style;
        const svg = opts.charts.draw(next, lookFor(next));
        if (!svg) return;
        try {
          await uploadAsset(e.scope, e.name, svg, 'image/svg+xml', true);
          e.source = next;
          e.text = svg;
          src = svgUrl(svg);
          img.src = src;
        } catch (err) {
          toast((err as Error).message);
        }
      });
      head.append(pick, iconButton('pencil', T.editChart, () => void chartDialog({ name: e.name, source: e.source ?? s })));
    }
    head.append(
      starButton(favourites.has(keyOf(e)), () => void toggleFavourite(e)),
      downloadButton(e.scope, e.name, () => src),
      deleteButton(e.scope, e.name, () => {
        current = null;
        void refresh();
      }),
    );
    const frame = el('button', 'asset-stage');
    frame.type = 'button';
    const img = el('img');
    img.src = src;
    img.alt = e.title;
    frame.append(img);
    frame.addEventListener('click', () => lightbox(src, e.title));
    const meta = el('p', 'asset-meta asset-detail-meta', [where(e.scope), ...(e.source ? [`${chartLabel(e.source.type)} · ${e.source.sheet}!${e.source.range}`] : []), ago(e.updatedAt)].join(' · '));
    body.append(head, meta, frame);
    main.replaceChildren(body);
  }

  async function retitle(e: ImageEntry, input: HTMLInputElement) {
    const t = input.value.trim();
    if (!t || t === e.title) {
      input.value = e.title;
      return;
    }
    const dot = e.name.lastIndexOf('.');
    const to = `${fileStem(t, 'image')}${e.name.slice(dot)}`;
    await inflight.catch(() => {});
    try {
      let name = e.name;
      if (to !== e.name) {
        const res = await fetch(api('/rename', { scope: e.scope, name: e.name, to }), { method: 'POST' });
        if (!res.ok) throw await failure(res);
        name = to;
      }
      if (e.source) {
        // The picture's own title follows.
        const next: ChartSource = { ...e.source, title: t };
        const svg = opts.charts.draw(next, lookFor(next));
        if (svg) await uploadAsset(e.scope, name, svg, 'image/svg+xml', true);
      }
      current = { scope: e.scope, name };
      await refresh();
    } catch (err) {
      input.value = e.title;
      toast((err as Error).message);
    }
  }

  // ----- styles -----

  function styleCard(s: ChartStyle): HTMLElement {
    const c = el('button', 'style-card');
    c.type = 'button';
    c.dataset.style = s.id;
    const img = el('img');
    img.src = svgUrl(sampleChart(s));
    img.alt = '';
    const swatches = el('span', 'style-swatches');
    for (const col of s.colors.slice(0, 6)) {
      const dot = el('i');
      dot.style.background = col;
      swatches.append(dot);
    }
    const name = el('span', 'style-name', styleName(s));
    const tags = el('span', 'style-tags');
    if (styles.default === s.id) tags.append(el('span', 'style-tag on', T.defaultBadge));
    if (builtin(s.id)) tags.append(el('span', 'style-tag', T.builtIn));
    const foot = el('span', 'style-foot');
    foot.append(name, tags, swatches);
    c.append(img, foot);
    c.addEventListener('click', () => {
      current = null;
      showStyle(s.id);
      paintList();
    });
    return c;
  }

  function showStyles() {
    editing = null;
    const body = el('div', 'lib-page');
    body.append(header(T.chartStyles, primaryButton(T.newStyle, 'plus', () => void copyStyle(findStyle(styles.default, styles.styles), true))));
    const grid = el('div', 'style-grid');
    grid.append(...allStyles().map(styleCard));
    body.append(grid);
    main.replaceChildren(body);
  }

  async function copyStyle(from: ChartStyle, blank = false) {
    const taken = new Set(allStyles().map((s) => s.id));
    let n = 1;
    const base = blank ? 'style' : `${from.id}`;
    while (taken.has(`${base}-${n}`)) n++;
    const copy: ChartStyle = { ...from, colors: [...from.colors], id: `${base}-${n}`, name: blank ? `${T.newStyle} ${n}` : `${styleName(from)} ${n + 1}` };
    const saved = await saveStyles({ ...styles, styles: [...styles.styles, copy] });
    if (!saved) return;
    styles = saved;
    showStyle(copy.id);
    paintList();
  }

  /** One style: its preview, and — for the workspace's own — every setting, saved as it changes. */
  function showStyle(id: string) {
    const s = allStyles().find((x) => x.id === id);
    if (!s) return showStyles();
    editing = id;
    const own = !builtin(id);
    const draft: ChartStyle = { ...s, colors: [...s.colors] };
    const body = el('div', 'lib-page style-page');
    const back = iconButton('arrowLeft', T.chartStyles, () => {
      showStyles();
      paintList();
    });
    const actions: HTMLElement[] = [];
    if (styles.default !== id) {
      actions.push(
        primaryButton(
          T.makeDefault,
          'check',
          async () => {
            const saved = await saveStyles({ ...styles, default: id });
            if (saved) {
              styles = saved;
              showStyle(id);
              paintList();
            }
          },
          false,
        ),
      );
    } else actions.push(el('span', 'style-tag on', T.defaultBadge));
    if (own) {
      actions.push(
        iconButton('trash', T.delete, async () => {
          if (!(await confirmDialog(T.deleteAssetTitle(draft.name), '', T.delete, true))) return;
          const saved = await saveStyles({ default: styles.default === id ? 'default' : styles.default, styles: styles.styles.filter((x) => x.id !== id) });
          if (saved) {
            styles = saved;
            showStyles();
            paintList();
          }
        }),
      );
    } else actions.push(primaryButton(T.duplicateEdit, 'copy', () => void copyStyle(s)));
    const head = el('div', 'lib-page-head');
    const title = el('h1', '', styleName(s));
    head.append(back, title, el('div', 'grow'), ...actions);

    const previews = el('div', 'style-previews');
    const p1 = el('img');
    const p2 = el('img');
    previews.append(p1, p2);
    const draw = () => {
      p1.src = svgUrl(sampleChart(draft, 'bar', styleName(draft)));
      p2.src = svgUrl(sampleChart(draft, 'doughnut'));
    };
    draw();

    let timer = 0;
    const changed = () => {
      draw();
      clearTimeout(timer);
      timer = window.setTimeout(async () => {
        const saved = await saveStyles({ ...styles, styles: styles.styles.map((x) => (x.id === id ? draft : x)) });
        if (saved) {
          styles = saved;
          paintList();
        }
      }, 450);
    };

    const form = el('div', 'style-form');
    const row = (label: string, control: HTMLElement) => {
      const r = el('label', 'style-row');
      r.append(el('span', '', label), control);
      form.append(r);
    };
    const colour = (value: string, set: (v: string) => void) => {
      const i = el('input', 'style-color');
      i.type = 'color';
      i.value = value;
      i.disabled = !own;
      i.addEventListener('input', () => {
        set(i.value);
        changed();
      });
      return i;
    };
    if (own) {
      const name = el('input', 'input');
      name.value = draft.name;
      name.addEventListener('input', () => {
        draft.name = name.value.trim() || draft.id;
        title.textContent = draft.name;
        changed();
      });
      row(T.title, name);
    }
    const series = el('div', 'style-colors');
    const paintSeries = () => {
      series.replaceChildren(
        ...draft.colors.map((c, i) =>
          colour(c, (v) => {
            draft.colors[i] = v;
          }),
        ),
      );
      if (!own) return;
      if (draft.colors.length < 12)
        series.append(
          iconButton('plus', T.addColor, () => {
            draft.colors.push(draft.colors[draft.colors.length % 8] ?? '#888888');
            paintSeries();
            changed();
          }),
        );
      if (draft.colors.length > 1)
        series.append(
          iconButton('x', T.removeColor, () => {
            draft.colors.pop();
            paintSeries();
            changed();
          }),
        );
    };
    paintSeries();
    row(T.styleColors, series);
    const pair = el('div', 'style-pair');
    for (const [label, key] of [
      [T.styleBackground, 'background'],
      [T.styleText, 'text'],
      [T.styleMuted, 'muted'],
      [T.styleGrid, 'grid'],
    ] as const) {
      const cell = el('label', 'style-swatch');
      cell.append(
        colour(draft[key], (v) => {
          draft[key] = v;
        }),
        el('span', '', label),
      );
      pair.append(cell);
    }
    row(T.styleColorsMore, pair);
    const font = el('select', 'input select');
    for (const [v, label] of [
      ['sans', T.fontSans],
      ['serif', T.fontSerif],
      ['mono', T.fontMono],
      ['rounded', T.fontRounded],
    ] as const)
      font.append(new Option(label, v));
    font.value = draft.font;
    font.disabled = !own;
    font.addEventListener('change', () => {
      draft.font = font.value as ChartStyle['font'];
      changed();
    });
    row(T.styleFont, font);
    const radius = el('input', 'style-range');
    radius.type = 'range';
    radius.min = '0';
    radius.max = '8';
    radius.value = String(draft.radius);
    radius.disabled = !own;
    radius.addEventListener('input', () => {
      draft.radius = Number(radius.value);
      changed();
    });
    row(T.styleCorners, radius);
    const grid = el('input');
    grid.type = 'checkbox';
    grid.checked = draft.gridLines;
    grid.disabled = !own;
    grid.addEventListener('change', () => {
      draft.gridLines = grid.checked;
      changed();
    });
    row(T.styleGridLines, grid);
    if (!own) form.classList.add('locked');

    const split = el('div', 'style-split');
    split.append(previews, form);
    body.append(head, split);
    main.replaceChildren(body);
  }

  // ----- the sidebar -----

  function visible(): ImageEntry[] {
    if (side.nav === 'recent') return newestFirst(images);
    if (side.nav === 'fav') return images.filter((e) => favourites.has(keyOf(e)));
    return images;
  }

  function paintList() {
    const q = side.query();
    const mark = () => {
      for (const n of side.list.querySelectorAll<HTMLElement>('.lib-item')) n.classList.toggle('on', Boolean(current) && n.dataset.key === keyOf(current!));
      for (const n of side.list.querySelectorAll<HTMLElement>('.lib-item[data-style]')) n.classList.toggle('on', n.dataset.style === editing);
    };
    const item = (e: ImageEntry, meta: string | HTMLElement | null = e.source ? metaOf(e) : ago(e.updatedAt)) => libItem(e, favourites.has(keyOf(e)), meta, () => open(e), e.text ? svgUrl(e.text) : fileUrl(e));
    if (q) {
      const hits = images
        .map((e) => {
          const hay = [e.title, e.name, e.source ? `${chartLabel(e.source.type)} ${e.source.sheet}!${e.source.range} ${e.source.title ?? ''}` : ''].join(' ');
          const s = snippet(hay, q);
          return s ? { e, s } : null;
        })
        .filter((x) => x !== null);
      side.list.replaceChildren(...(hits.length ? hits.map((h) => item(h.e, h.s)) : [el('p', 'lib-empty', T.noMatches)]));
      return mark();
    }
    if (side.nav === 'home') return side.list.replaceChildren();
    if (side.nav === 'styles') {
      side.list.replaceChildren(
        ...allStyles().map((s) => {
          const b = el('button', 'lib-item', '');
          b.type = 'button';
          b.dataset.style = s.id;
          const dots = el('span', 'style-swatches');
          for (const col of s.colors.slice(0, 4)) {
            const d = el('i');
            d.style.background = col;
            dots.append(d);
          }
          const text = el('span', 'lib-item-text');
          const top = el('span', 'lib-item-top');
          top.append(el('span', 'lib-item-title', styleName(s)));
          text.append(top, el('span', 'lib-item-meta', styles.default === s.id ? T.defaultBadge : builtin(s.id) ? T.builtIn : ''));
          b.append(dots, text);
          b.addEventListener('click', () => {
            showStyle(s.id);
            paintList();
          });
          return b;
        }),
      );
      return mark();
    }
    const list = visible();
    if (!list.length) {
      side.list.replaceChildren(el('p', 'lib-empty', side.nav === 'fav' ? T.noFavoriteAssets : T.noAssetsYet));
      return;
    }
    if (side.nav === 'all') {
      const groups: HTMLElement[] = [];
      for (const scope of [opts.bookId, SHARED]) {
        const mine = list.filter((e) => e.scope === scope);
        if (!mine.length) continue;
        const g = el('section', 'lib-group');
        g.dataset.scope = scope;
        g.append(el('h3', '', where(scope)), ...mine.map((e) => item(e)));
        groups.push(g);
      }
      side.list.replaceChildren(...groups);
    } else side.list.replaceChildren(...list.map((e) => item(e, `${where(e.scope)} · ${ago(e.updatedAt)}`)));
    mark();
  }

  function open(e: ImageEntry) {
    if (side.nav === 'home' || side.nav === 'styles') side.select('all', true);
    current = { scope: e.scope, name: e.name };
    showDetail(e);
    paintList();
  }

  /** Lists, then fills the right side for where the sidebar is — except a style being edited. */
  async function paint() {
    paintList();
    if (side.nav === 'styles') {
      if (editing && allStyles().some((s) => s.id === editing)) return;
      return showStyles();
    }
    const open = current && images.find((e) => keyOf(e) === keyOf(current!));
    if (open) return showDetail(open);
    current = null;
    if (side.nav === 'home') return showHome();
    const first = visible()[0];
    if (first) {
      current = { scope: first.scope, name: first.name };
      showDetail(first);
      paintList();
    } else showHome();
  }

  const texts = new Map<string, { at: string; text: string }>();
  /** Each refresh bumps this; one that finishes after a newer one started is dropped (its list may be stale). */
  let loading = 0;
  /** The refresh under way; a rename waits for it, so it never reads a name that is gone. */
  let inflight: Promise<void> = Promise.resolve();
  function refresh(prefer?: { scope: string; name: string }): Promise<void> {
    inflight = load(prefer);
    return inflight;
  }
  async function load(prefer?: { scope: string; name: string }) {
    const ticket = ++loading;
    const [mine, shared, favs, st] = await Promise.all([listAssets(opts.bookId), listAssets(SHARED), loadFavourites(), loadStyles()]);
    if (ticket !== loading) return;
    favourites = favs;
    styles = st;
    const entries = await Promise.all(
      [...mine.map((a) => ({ a, scope: opts.bookId })), ...shared.map((a) => ({ a, scope: SHARED }))]
        .filter(({ a }) => a.kind === 'image')
        .map(async ({ a, scope }): Promise<ImageEntry> => {
          const e = { scope, name: a.name, title: stem(a.name), updatedAt: a.updatedAt };
          if (!/\.svg$/i.test(a.name)) return { ...e, source: null, text: '' };
          const k = keyOf(e);
          let cached = texts.get(k);
          if (cached?.at !== a.updatedAt) {
            cached = { at: a.updatedAt, text: await readText(scope, a.name) };
            texts.set(k, cached);
          }
          const source = chartSourceOf(cached.text);
          // A chart reads by the title it shows, not its file name.
          return { ...e, title: source?.title || e.title, source, text: cached.text };
        }),
    );
    if (ticket !== loading) return;
    images = entries;
    if (prefer) {
      current = prefer;
      if (side.nav === 'home' || side.nav === 'styles') side.select('all', true);
    }
    // A style being edited keeps its page; the list still follows.
    if (side.nav === 'styles' && editing) return paintList();
    await paint();
  }

  return { el: page, refresh: () => refresh(), newChart: () => chartDialog() };
}

// ---------- Notes: Markdown, edited in place ----------

/**
 * Images in a note are files beside it, by path relative to the note — so the
 * Markdown reads the same on GitHub or in an editor: `chart.svg` (same folder),
 * `../../../assets/x.svg` (shared, from a workbook's note), `../sheets/<id>/assets/x.svg`
 * (a workbook's, from a shared note). Nothing else is shown.
 */
function noteImages(noteScope: string, absolute = false): ImageResolver {
  return (src) => {
    let s = src.trim();
    try {
      s = decodeURIComponent(s);
    } catch {
      return null;
    }
    s = s.replace(/^\.\//, '');
    let scope = noteScope;
    let name = s;
    const toShared = /^\.\.\/\.\.\/\.\.\/assets\/([^/]+)$/.exec(s);
    const toBook = /^\.\.\/sheets\/([a-z0-9][a-z0-9-]*)\/assets\/([^/]+)$/.exec(s);
    if (toShared && noteScope !== SHARED) [scope, name] = [SHARED, toShared[1]!];
    else if (toBook && noteScope === SHARED) [scope, name] = [toBook[1]!, toBook[2]!];
    else if (s.includes('/')) return null;
    if (!/\.(svg|png|jpe?g|gif|webp)$/i.test(name) || name.startsWith('.')) return null;
    return `${absolute ? location.origin : ''}${api('/file', { scope, name })}`;
  };
}
/** The path a note at `noteScope` uses for an image; spaces and brackets escaped so the Markdown holds. */
function imageRef(noteScope: string, scope: string, name: string): string {
  const n = name.replace(/[ ()]/g, (c) => ({ ' ': '%20', '(': '%28', ')': '%29' })[c]!);
  if (scope === noteScope) return n;
  return scope === SHARED ? `../../../assets/${n}` : `../sheets/${scope}/assets/${n}`;
}

/** A note as a printed page — the print dialog saves it as a PDF. */
function printNote(markdown: string, title: string, scope: string) {
  const esc = (t: string) => t.replaceAll('&', '&amp;').replaceAll('<', '&lt;');
  const html = `<!doctype html><html lang="${document.documentElement.lang || 'en'}"><head><meta charset="utf-8"><title>${esc(title)}</title><style>
@page { size: A4; margin: 20mm 18mm; }
body { margin: 0; color: #1f1f1f; font: 11.5pt/1.7 -apple-system, BlinkMacSystemFont, "PingFang TC", "Noto Sans TC", "Segoe UI", sans-serif; word-break: break-word; }
h1 { font-size: 22pt; line-height: 1.3; margin: 0 0 0.6em; letter-spacing: -0.01em; }
h2 { font-size: 15pt; margin: 1.4em 0 0.5em; } h3 { font-size: 12.5pt; margin: 1.2em 0 0.4em; }
p, ul, ol, blockquote, pre { margin: 0 0 0.8em; } ul, ol { padding-left: 1.4em; }
h1, h2, h3 { break-after: avoid; } pre, blockquote, li { break-inside: avoid; }
blockquote { padding-left: 1em; border-left: 3px solid #1f1f1f; color: #5f6368; }
code { font: 0.9em "SF Mono", Menlo, monospace; background: #f1f3f4; padding: 0.1em 0.3em; border-radius: 3px; }
pre { background: #f1f3f4; padding: 10px 12px; border-radius: 6px; white-space: pre-wrap; } pre code { background: none; padding: 0; }
hr { border: 0; text-align: center; margin: 1.4em 0; } hr::after { content: '···'; letter-spacing: 0.6em; color: #9aa0a6; }
a { color: #1a73e8; }
img { display: block; max-width: 100%; margin: 0.4em 0 0.9em; break-inside: avoid; }
</style></head><body>${mdToHtml(markdown, noteImages(scope, true))}</body></html>`;
  (window as unknown as { __lastNotePrint?: string }).__lastNotePrint = html;
  const frame = el('iframe', 'print-frame');
  frame.setAttribute('aria-hidden', 'true');
  document.body.append(frame);
  const doc = frame.contentDocument!;
  doc.open();
  doc.write(html);
  doc.close();
  const done = () => setTimeout(() => frame.remove(), 1000);
  frame.contentWindow!.addEventListener('afterprint', done);
  // Pictures first: the page is printed once they have loaded (or failed).
  const pictures = [...doc.images].map((img) =>
    img.complete
      ? Promise.resolve()
      : new Promise<void>((ok) => {
          img.addEventListener('load', () => ok(), { once: true });
          img.addEventListener('error', () => ok(), { once: true });
        }),
  );
  void Promise.all(pictures).then(() =>
    setTimeout(() => {
      frame.contentWindow!.focus();
      frame.contentWindow!.print();
      done();
    }, 50),
  );
}

export function createNotesView(opts: { bookId: string }) {
  const page = el('section', 'lib-view note-view');
  page.dataset.view = 'notes';
  const side = librarySide({
    title: T.viewNotes,
    actions: [iconButton('plus', T.newNote, () => void newNote())],
    searchLabel: T.searchNotes,
    store: 'open-spreadsheet:notes-nav',
    onNav: (n) => void navigate(n),
    onSearch: () => void renderList(),
  });
  const main = el('div', 'lib-main note-main');
  page.append(side.el, main);

  let current: { scope: string; name: string } | null = null;
  let dirty = false;
  let timer = 0;
  let saving: Promise<void> = Promise.resolve();
  /** Each open bumps this; a load that finishes after a newer one started is dropped. */
  let opening = 0;
  // Written as it reads; the </> toggle shows the Markdown source.
  const textarea = el('textarea', 'note-text');
  textarea.spellcheck = false;
  const editor = createNoteEditor({ onChange: () => changed(), placeholder: T.notePlaceholder, image: () => noteImages(current?.scope ?? opts.bookId) });
  let source = false;
  const state = el('span', 'note-state');
  const sourceBtn = iconButton('code', T.markdownSource, () => {
    if (source) editor.setMarkdown(textarea.value);
    else textarea.value = editor.getMarkdown();
    source = !source;
    sourceBtn.classList.toggle('on', source);
    (source ? editor.el : textarea).replaceWith(source ? textarea : editor.el);
    (source ? textarea : editor.el).focus();
  });

  // Insert an asset: a picture from this workbook or the shared ones, by its path from the note.
  const insertBtn = iconButton('image', T.insertAsset, async () => {
    const [mine, shared] = await Promise.all([listAssets(opts.bookId), listAssets(SHARED)]);
    const pics = [...mine.map((a) => ({ a, scope: opts.bookId })), ...shared.map((a) => ({ a, scope: SHARED }))].filter(({ a }) => a.kind === 'image');
    if (!current) return;
    if (!pics.length) return toast(T.noAssetsYet);
    const here = current.scope;
    showMenu(
      pics.map(({ a, scope }) => ({
        label: scope === here ? stem(a.name) : `${stem(a.name)} · ${where(scope)}`,
        icon: 'image',
        action: () => {
          const md = `![${stem(a.name)}](${imageRef(here, scope, a.name)})`;
          if (source) {
            const at = textarea.selectionEnd;
            const before = textarea.value.slice(0, at);
            textarea.value = `${before}${before && !before.endsWith('\n') ? '\n\n' : ''}${md}\n${textarea.value.slice(at)}`;
            changed();
          } else editor.insertBlock(md);
        },
      })),
      insertBtn,
    );
  });

  async function save() {
    if (!current || !dirty) return;
    const { scope, name } = current;
    const text = source ? textarea.value : editor.getMarkdown();
    dirty = false;
    state.textContent = T.saving;
    saving = writeNote(scope, name, text).then(
      () => {
        if (!dirty) state.textContent = T.saved;
      },
      (e: Error) => {
        dirty = true;
        state.textContent = e.message;
      },
    );
    await saving;
  }
  function changed() {
    dirty = true;
    state.textContent = T.unsaved;
    clearTimeout(timer);
    timer = window.setTimeout(() => void save(), 700);
  }
  textarea.addEventListener('input', changed);
  for (const target of [textarea, editor.el]) {
    target.addEventListener('keydown', (e) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        clearTimeout(timer);
        void save();
      }
    });
  }

  let notes: Entry[] = [];
  let favourites = new Set<string>();
  /** Note bodies for search and the overview, fetched once per change. */
  const bodies = new Map<string, { at: string; text: string }>();
  async function loadBodies() {
    await Promise.all(
      notes.map(async (n) => {
        const b = bodies.get(keyOf(n));
        if (b?.at === n.updatedAt) return;
        bodies.set(keyOf(n), { at: n.updatedAt, text: await readText(n.scope, n.name) });
      }),
    );
  }

  function mark() {
    for (const n of side.list.querySelectorAll<HTMLElement>('.lib-item')) n.classList.toggle('on', Boolean(current) && n.dataset.key === keyOf(current!));
  }

  async function open(scope: string, name: string) {
    const ticket = ++opening;
    clearTimeout(timer);
    await save();
    const text = await readText(scope, name);
    if (ticket !== opening) return;
    if (side.nav === 'home') {
      side.select('all', true);
      void renderList();
    }
    current = { scope, name };
    textarea.value = text;
    editor.setMarkdown(text);
    dirty = false;
    state.textContent = T.saved;
    const head = el('div', 'lib-detail-head note-head');
    // Download as the Markdown file, or as a PDF through the print dialog.
    const dl = iconButton('download', T.download, () =>
      showMenu(
        [
          {
            label: 'Markdown (.md)',
            icon: 'file',
            action: () => {
              const link = el('a');
              link.href = api('/file', { scope, name, download: '1' });
              link.click();
            },
          },
          { label: 'PDF', icon: 'download', action: () => printNote(source ? textarea.value : editor.getMarkdown(), stem(name), scope) },
        ],
        dl,
      ),
    );
    // The note's own first heading is its title; the bar only says where it lives.
    head.append(
      el('span', 'lib-scope', where(scope)),
      el('div', 'grow'),
      state,
      starButton(favourites.has(keyOf({ scope, name })), () => void toggleFavourite(scope, name)),
      insertBtn,
      sourceBtn,
      dl,
      deleteButton(scope, name, () => {
        current = null;
        void refresh();
      }),
    );
    main.replaceChildren(head, source ? textarea : editor.el);
    mark();
    if (source) textarea.focus();
    else editor.focus();
  }

  /** The overview: favourites, then every note, newest first, each with its opening lines. */
  async function showHome() {
    const ticket = ++opening;
    clearTimeout(timer);
    await save();
    current = null;
    await loadBodies();
    if (ticket !== opening) return;
    const body = el('div', 'lib-page');
    const head = el('div', 'lib-page-head');
    head.append(el('h1', '', T.viewNotes), el('div', 'grow'));
    if (notes.length) head.append(primaryButton(T.newNote, 'plus', () => void newNote()));
    body.append(head);
    if (!notes.length) {
      const empty = el('div', 'lib-empty-state');
      empty.append(icon('file', 28), el('h3', '', T.noNotesYet), primaryButton(T.newNote, 'plus', () => void newNote()));
      body.append(empty);
      main.replaceChildren(body);
      mark();
      return;
    }
    const noteCard = (n: Entry) => {
      const c = el('button', 'note-card');
      c.type = 'button';
      c.dataset.key = keyOf(n);
      const text = bodies.get(keyOf(n))?.text ?? '';
      const lines = text
        .split('\n')
        .filter((l) => l.trim())
        .slice(1)
        .map((l) => l.replace(/!\[[^\]]*\]\([^)]*\)/g, '').replace(/^\s*(#+|[-*+]|\d+\.|>)\s+/, '').replace(/[*`_]+/g, ''))
        .join('  ')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 160);
      const top = el('span', 'note-card-title', n.title);
      if (favourites.has(keyOf(n))) top.prepend(icon('starFill', 12));
      c.append(top, el('span', 'note-card-text', lines), el('span', 'note-card-meta', `${where(n.scope)} · ${ago(n.updatedAt)}`));
      c.addEventListener('click', () => void open(n.scope, n.name));
      return c;
    };
    const section = (label: string, list: Entry[]) => {
      if (!list.length) return;
      const s = el('section', 'note-section');
      const grid = el('div', 'note-grid');
      grid.append(...list.map(noteCard));
      s.append(el('h3', '', label), grid);
      body.append(s);
    };
    section(T.notesFavorites, notes.filter((n) => favourites.has(keyOf(n))));
    section(T.notesRecent, newestFirst(notes));
    main.replaceChildren(body);
    mark();
  }

  function showEmpty() {
    current = null;
    const empty = el('div', 'lib-empty-state');
    empty.append(icon('file', 28), el('h3', '', T.noNotesYet), primaryButton(T.newNote, 'plus', () => void newNote()));
    main.replaceChildren(empty);
  }

  /** What the sidebar lists for Recent, Favorites or All. */
  function visible(): Entry[] {
    if (side.nav === 'recent') return newestFirst(notes);
    if (side.nav === 'fav') return notes.filter((n) => favourites.has(keyOf(n)));
    return notes;
  }

  let searching = 0;
  async function renderList() {
    const q = side.query();
    const item = (n: Entry, meta: string | HTMLElement | null = null) => libItem(n, favourites.has(keyOf(n)), meta, () => void open(n.scope, n.name));
    if (q) {
      const ticket = ++searching;
      await loadBodies();
      if (ticket !== searching) return;
      const hits = notes
        .map((n) => {
          const inTitle = n.title.toLowerCase().includes(q);
          const s = snippet(bodies.get(keyOf(n))?.text ?? '', q);
          if (!inTitle && !s) return null;
          return { n, s: s ?? where(n.scope), rank: inTitle ? 0 : 1 };
        })
        .filter((x) => x !== null)
        .sort((a, b) => a.rank - b.rank || b.n.updatedAt.localeCompare(a.n.updatedAt));
      side.list.replaceChildren(...(hits.length ? hits.map((h) => item(h.n, h.s)) : [el('p', 'lib-empty', T.noMatches)]));
      return mark();
    }
    if (side.nav === 'home') return side.list.replaceChildren();
    const list = visible();
    if (!list.length) {
      side.list.replaceChildren(el('p', 'lib-empty', side.nav === 'fav' ? T.noFavorites : T.noNotesYet));
      return;
    }
    if (side.nav === 'all') {
      const groups: HTMLElement[] = [];
      for (const scope of [opts.bookId, SHARED]) {
        const mine = list.filter((n) => n.scope === scope);
        if (!mine.length) continue;
        const g = el('section', 'lib-group');
        g.dataset.scope = scope;
        g.append(el('h3', '', where(scope)), ...mine.map((n) => item(n)));
        groups.push(g);
      }
      side.list.replaceChildren(...groups);
    } else side.list.replaceChildren(...list.map((n) => item(n, `${where(n.scope)} · ${ago(n.updatedAt)}`)));
    mark();
  }

  async function navigate(n: Nav) {
    await renderList();
    if (n === 'home') return showHome();
    if (current && notes.some((x) => keyOf(x) === keyOf(current!))) return;
    const first = visible()[0];
    if (first) await open(first.scope, first.name);
    else showEmpty();
  }

  async function toggleFavourite(scope: string, name: string) {
    const on = !favourites.has(keyOf({ scope, name }));
    const next = await putFavourite(scope, name, on);
    if (!next) return;
    favourites = next;
    const star = main.querySelector<HTMLButtonElement>('.lib-star');
    if (star) {
      star.replaceChildren(icon(on ? 'starFill' : 'star'));
      star.classList.toggle('on', on);
      star.title = on ? T.unfavorite : T.favorite;
    }
    await renderList();
  }

  /** Lists the notes; opens `prefer`, else keeps the open one, else Home or the first one listed. */
  async function refresh(prefer?: { scope: string; name: string }) {
    const [mine, shared, favs] = await Promise.all([listAssets(opts.bookId), listAssets(SHARED), loadFavourites()]);
    const toEntry = (scope: string) => (a: AssetEntry) => ({ scope, name: a.name, title: a.title || stem(a.name), updatedAt: a.updatedAt });
    notes = [...mine.filter((a) => a.kind === 'note').map(toEntry(opts.bookId)), ...shared.filter((a) => a.kind === 'note').map(toEntry(SHARED))];
    favourites = favs;
    const has = (t: { scope: string; name: string } | null | undefined) => Boolean(t && notes.some((x) => x.scope === t.scope && x.name === t.name));
    if (prefer && has(prefer)) {
      if (side.nav === 'home') side.select('all', true);
      await renderList();
      return open(prefer.scope, prefer.name);
    }
    await renderList();
    if (has(current)) return mark();
    if (side.nav === 'home') return showHome();
    const first = visible()[0] ?? notes[0];
    if (first) await open(first.scope, first.name);
    else showEmpty();
  }

  async function newNote() {
    const v = await promptDialog({
      title: T.newNote,
      fields: [
        { name: 'title', label: T.title, value: '', validate: (x) => (x.trim() ? null : T.title) },
        { name: 'scope', label: T.saveIn, options: [{ value: opts.bookId, label: T.assetsThisBook }, { value: SHARED, label: T.assetsShared }], value: opts.bookId },
      ],
      ok: T.create,
    });
    if (!v) return;
    const t = v.title!.trim();
    const name = `${fileStem(t, 'note')}.md`;
    try {
      await save();
      await writeNote(v.scope!, name, `# ${t}\n\n`);
      await refresh({ scope: v.scope!, name });
    } catch (e) {
      toast((e as Error).message);
    }
  }

  return {
    el: page,
    refresh: () => refresh(),
    async flush() {
      clearTimeout(timer);
      await save();
      await saving;
    },
  };
}
