/** The workspace's home: a sidebar and a gallery of workbooks or templates, like open-doc and open-slide. */
import { formatValue, impliedFormat, Workbook } from '@open-spreadsheet/engine';
import { isZh, LANGS, lang, locale, setLang, T } from './i18n.ts';
import { icon } from './icons.ts';
import { exportXlsx, readSpreadsheetFile } from './io.ts';
import { TEMPLATES } from './templates.ts';
import type { WorkbookData, WorkbookSummary } from './types.ts';
import { commandPalette, confirmDialog, currentTheme, downloadBlob, el, MOD, pickFile, promptDialog, showMenu, toast, toggleTheme } from './ui.ts';

const ID = /^[a-z0-9][a-z0-9._-]{0,63}$/;

interface ListResponse {
  name: string;
  version: string;
  workbooks: WorkbookSummary[];
}

export const api = {
  list: () => fetch('/__api/workbooks').then((r) => r.json() as Promise<ListResponse>),
  get: (id: string) => fetch(`/__api/workbooks/${id}`).then((r) => (r.ok ? (r.json() as Promise<WorkbookData>) : Promise.reject(new Error(String(r.status))))),
  /** The workbook and its version (a SHA-256), for saving without overwriting someone else's change. */
  load: async (id: string) => {
    const r = await fetch(`/__api/workbooks/${id}`);
    if (!r.ok) throw new Error(String(r.status));
    return { data: (await r.json()) as WorkbookData, etag: r.headers.get('etag') ?? '' };
  },
  put: (id: string, data: WorkbookData, etag?: string) =>
    fetch(`/__api/workbooks/${id}`, {
      method: 'PUT',
      headers: { 'content-type': 'application/json', ...(etag ? { 'if-match': etag } : {}) },
      body: JSON.stringify(data),
    }),
  create: async (body: { id: string; title?: string; data?: WorkbookData; from?: string }) => {
    const r = await fetch('/__api/workbooks', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });
    if (!r.ok) throw new Error(((await r.json()) as { error: string }).error);
  },
  remove: (id: string) => fetch(`/__api/workbooks/${id}`, { method: 'DELETE' }),
};

/** A file name from a title: ASCII words joined by dashes, else a dated name. */
export function slugify(title: string, taken: Set<string>): string {
  let base = title
    .normalize('NFKD')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 48);
  if (!base) {
    const d = new Date();
    base = `sheet-${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
  }
  let id = base;
  for (let n = 2; taken.has(id); n++) id = `${base}-${n}`;
  return id;
}

/** Asks for a title and file name, then creates the workbook and opens it. */
export async function createFlow(taken: Set<string>, opts: { title?: string; data?: WorkbookData; from?: string; heading?: string } = {}): Promise<void> {
  const title = opts.title ?? '';
  const values = await promptDialog({
    title: opts.heading ?? T.newSheet,
    ok: T.create,
    fields: [
      {
        name: 'title',
        label: T.title,
        value: title,
        placeholder: T.titlePlaceholder,
        derive: (v, f) => {
          if (!f.id!.dataset.touched) f.id!.value = slugify(v, taken);
        },
        validate: (v) => (v.trim() ? null : T.title),
      },
      {
        name: 'id',
        label: T.fileName,
        value: slugify(title, taken),
        hint: T.fileNameHint,
        mono: true,
        validate: (v) => (!ID.test(v) ? T.fileNameHint : taken.has(v) ? T.nameTaken : null),
      },
    ],
  });
  if (!values) return;
  try {
    await api.create({ id: values.id!, title: values.title!.trim(), data: opts.data ? { ...opts.data, title: values.title!.trim() } : undefined, from: opts.from });
    location.hash = `#/w/${encodeURIComponent(values.id!)}`;
  } catch (e) {
    toast(String((e as Error).message));
  }
}

function thumbnail(data: WorkbookData): HTMLElement {
  const thumb = el('div', 'thumb');
  const sheet = data.sheets[0];
  if (!sheet) return thumb;
  const wb = Workbook.fromJSON({ sheets: [sheet] });
  const grid = el('div', 'thumb-grid');
  for (let r = 0; r < 9; r++) {
    for (let c = 0; c < 6; c++) {
      const a = `${String.fromCharCode(65 + c)}${r + 1}`;
      const style = sheet.styles?.[a];
      const input = sheet.cells[a] ?? '';
      const f = formatValue(wb.value(sheet.name, r, c), style?.fmt ?? impliedFormat(input));
      const span = el('span', typeof wb.value(sheet.name, r, c) === 'number' ? 'n' : '', f.text);
      if (style?.bg) span.style.background = style.bg;
      if (style?.color || f.color) span.style.color = (f.color ?? style?.color)!;
      if (style?.b) span.style.fontWeight = '600';
      grid.append(span);
    }
  }
  thumb.append(grid);
  return thumb;
}

export async function renderHome(app: HTMLElement, view: 'books' | 'templates'): Promise<void> {
  const res = await api.list();
  document.title = `${res.name} — open-spreadsheet`;
  const taken = new Set(res.workbooks.map((w) => w.id));
  const shell = el('div', 'shell');

  // ----- sidebar -----
  const side = el('aside', 'sidebar');
  const brand = el('div', 'brand');
  const mark = el('span', 'brand-mark');
  mark.append(icon('sheet', 14));
  brand.append(mark, el('span', '', 'open-spreadsheet'));
  const search = el('button', 'side-search');
  search.type = 'button';
  search.append(icon('search', 15), el('span', '', T.search), el('kbd', '', `${MOD}K`));
  const nav = el('nav', 'side-nav');
  const navItem = (name: string, label: string, count: number, active: boolean, href: string) => {
    const a = el('a', `side-item${active ? ' active' : ''}`);
    a.href = href;
    a.append(icon(name, 16), el('span', 'side-label', label), el('span', 'side-count', String(count).padStart(2, '0')));
    return a;
  };
  nav.append(navItem('sheet', T.spreadsheets, res.workbooks.length, view === 'books', '#/'), navItem('template', T.templates, TEMPLATES.length, view === 'templates', '#/templates'));
  const section = el('div', 'side-section');
  section.append(el('span', '', T.workspace));
  const ws = el('div', 'side-workspace');
  ws.append(icon('file', 15), el('span', '', res.name));
  const foot = el('div', 'side-foot');
  foot.append(el('span', 'side-version', `v${res.version}`));
  const langBtn = el('button', 'ibtn');
  langBtn.type = 'button';
  langBtn.title = T.language;
  langBtn.append(icon('languages', 15));
  langBtn.addEventListener('click', () =>
    showMenu(
      LANGS.map((l) => ({
        label: l.label,
        checked: l.id === lang,
        action: () => {
          setLang(l.id);
          void renderHome(app, view);
        },
      })),
      langBtn,
    ),
  );
  const themeBtn = el('button', 'ibtn');
  themeBtn.type = 'button';
  themeBtn.title = T.theme;
  themeBtn.append(icon(currentTheme() === 'dark' ? 'moon' : 'sun', 15));
  themeBtn.addEventListener('click', () => {
    toggleTheme();
    themeBtn.replaceChildren(icon(currentTheme() === 'dark' ? 'moon' : 'sun', 15));
  });
  foot.append(el('span', 'grow'), langBtn, themeBtn);
  side.append(brand, search, nav, section, ws, el('div', 'grow'), foot);

  // ----- main panel -----
  const panel = el('main', 'panel');
  const head = el('header', 'panel-head');
  const h = el('h1', 'panel-title');
  h.append(icon(view === 'books' ? 'sheet' : 'template', 20), el('span', '', view === 'books' ? T.spreadsheets : T.templates));
  h.append(el('span', 'panel-count', String(view === 'books' ? res.workbooks.length : TEMPLATES.length).padStart(2, '0')));
  head.append(h, el('div', 'grow'));
  const body = el('div', 'panel-body');

  if (view === 'books') {
    let sort: 'newest' | 'name' = 'newest';
    try {
      if (localStorage.getItem('open-spreadsheet:sort') === 'name') sort = 'name';
    } catch {
      // default order
    }
    const sortBtn = el('button', 'btn');
    sortBtn.type = 'button';
    const paintSort = () => sortBtn.replaceChildren(icon('clock', 15), el('span', '', sort === 'newest' ? T.newest : T.byName), icon('chevronDown', 14));
    paintSort();
    const filterBox = el('label', 'filter');
    const filter = el('input');
    filter.placeholder = T.filter;
    filter.spellcheck = false;
    filterBox.append(icon('search', 15), filter);
    const importBtn = el('button', 'btn');
    importBtn.type = 'button';
    importBtn.title = T.importHint;
    importBtn.classList.add('import-btn');
    importBtn.append(icon('upload', 15), el('span', '', T.import));
    const newBtn = el('button', 'btn primary');
    newBtn.type = 'button';
    newBtn.append(icon('plus', 15), el('span', '', T.newSheet));
    head.append(sortBtn, filterBox, importBtn, newBtn);

    const cards = el('div', 'cards');
    const datas = new Map<string, WorkbookData>();
    const paint = () => {
      const q = filter.value.trim().toLowerCase();
      const list = res.workbooks
        .filter((w) => !q || w.title.toLowerCase().includes(q) || w.id.includes(q))
        .sort((a, b) => (sort === 'name' ? a.title.localeCompare(b.title) : b.updatedAt.localeCompare(a.updatedAt)));
      cards.replaceChildren();
      if (res.workbooks.length === 0) {
        const empty = el('div', 'empty');
        empty.append(icon('sheet', 28), el('p', '', T.empty));
        const row = el('div', 'empty-actions');
        const b1 = el('button', 'btn primary');
        b1.append(icon('plus', 15), el('span', '', T.newSheet));
        b1.onclick = () => void createFlow(taken);
        const b2 = el('a', 'btn', T.templates);
        b2.href = '#/templates';
        row.append(b1, b2);
        empty.append(row);
        cards.append(empty);
        return;
      }
      if (list.length === 0) cards.append(el('p', 'muted', T.noMatch));
      for (const w of list) {
        const card = el('a', 'card');
        card.href = `#/w/${encodeURIComponent(w.id)}`;
        const thumbSlot = el('div', 'thumb');
        const meta = el('div', 'card-meta');
        const date = new Date(w.updatedAt).toLocaleDateString(locale(), { year: 'numeric', month: 'short', day: 'numeric' });
        meta.append(el('div', 'card-title', w.title), el('div', 'card-sub', `${T.sheetsCount(w.sheets)} · ${date}`));
        const more = el('button', 'ibtn card-more');
        more.type = 'button';
        more.title = '…';
        more.append(icon('more', 16));
        const menu = (at: { x: number; y: number } | HTMLElement) =>
          showMenu(
            [
              { label: T.open, icon: 'sheet', action: () => (location.hash = `#/w/${encodeURIComponent(w.id)}`) },
              {
                label: T.rename,
                icon: 'pencil',
                action: async () => {
                  const v = await promptDialog({ title: T.rename, fields: [{ name: 'title', label: T.title, value: w.title, validate: (x) => (x.trim() ? null : T.title) }], ok: T.save });
                  if (!v) return;
                  const data = datas.get(w.id) ?? (await api.get(w.id));
                  await api.put(w.id, { ...data, title: v.title!.trim() });
                  void renderHome(app, view);
                },
              },
              { label: T.duplicate, icon: 'copy', action: () => void createFlow(taken, { title: `${w.title} ${T.copySuffix}`, from: w.id, heading: T.duplicate }) },
              {
                label: T.exportXlsx,
                icon: 'download',
                action: async () => {
                  const data = datas.get(w.id) ?? (await api.get(w.id));
                  downloadBlob(exportXlsx(data, Workbook.fromJSON(data)), `${w.id}.xlsx`);
                },
              },
              '-',
              {
                label: T.delete,
                icon: 'trash',
                danger: true,
                action: async () => {
                  if (!(await confirmDialog(T.deleteBook(w.title), T.deleteBookBody, T.delete, true))) return;
                  await api.remove(w.id);
                  void renderHome(app, view);
                },
              },
            ],
            at,
          );
        more.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          menu(more);
        });
        card.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          menu({ x: e.clientX, y: e.clientY });
        });
        meta.append(more);
        card.append(thumbSlot, meta);
        cards.append(card);
        const fill = (d: WorkbookData) => thumbSlot.replaceWith(thumbnail(d));
        const cached = datas.get(w.id);
        if (cached) fill(cached);
        else
          void api.get(w.id).then((d) => {
            datas.set(w.id, d);
            if (thumbSlot.isConnected) fill(d);
          });
      }
    };
    paint();
    filter.addEventListener('input', paint);
    sortBtn.addEventListener('click', () =>
      showMenu(
        [
          { label: T.newest, checked: sort === 'newest', action: () => ((sort = 'newest'), save(), paintSort(), paint()) },
          { label: T.byName, checked: sort === 'name', action: () => ((sort = 'name'), save(), paintSort(), paint()) },
        ],
        sortBtn,
      ),
    );
    const save = () => {
      try {
        localStorage.setItem('open-spreadsheet:sort', sort);
      } catch {
        // not remembered
      }
    };
    newBtn.addEventListener('click', () => void createFlow(taken));
    const doImport = async (file: File | null) => {
      if (!file) return;
      try {
        const data = await readSpreadsheetFile(file);
        await createFlow(taken, { title: data.title, data, heading: T.import });
      } catch {
        toast(T.importFailed);
      }
    };
    importBtn.addEventListener('click', async () => doImport(await pickFile('.csv,.tsv,.txt,.xlsx')));
    panel.addEventListener('dragover', (e) => {
      e.preventDefault();
      panel.classList.add('dropping');
    });
    panel.addEventListener('dragleave', () => panel.classList.remove('dropping'));
    panel.addEventListener('drop', (e) => {
      e.preventDefault();
      panel.classList.remove('dropping');
      void doImport(e.dataTransfer?.files[0] ?? null);
    });
    body.append(cards);
  } else {
    const cards = el('div', 'cards');
    const zh = isZh();
    const blank = el('button', 'card template');
    blank.type = 'button';
    const bt = el('div', 'thumb blank');
    bt.append(icon('plus', 28));
    const bm = el('div', 'card-meta');
    bm.append(el('div', 'card-title', T.blank), el('div', 'card-sub', T.blankSub));
    blank.append(bt, bm);
    blank.addEventListener('click', () => void createFlow(taken));
    cards.append(blank);
    for (const t of TEMPLATES) {
      const data = t.build(zh);
      const card = el('button', 'card template');
      card.type = 'button';
      const meta = el('div', 'card-meta');
      meta.append(el('div', 'card-title', zh ? t.title.zh : t.title.en), el('div', 'card-sub', zh ? t.description.zh : t.description.en));
      card.append(thumbnail(data), meta);
      card.addEventListener('click', () => void createFlow(taken, { title: data.title, data, heading: T.useTemplate }));
      cards.append(card);
    }
    body.append(cards);
  }

  panel.append(head, body);
  shell.append(side, panel);
  app.replaceChildren(shell);

  const palette = () =>
    commandPalette(
      [
        ...res.workbooks.map((w) => ({ group: T.gBooks, label: w.title, icon: 'sheet', keywords: w.id, run: () => (location.hash = `#/w/${encodeURIComponent(w.id)}`) })),
        { group: T.gFile, label: T.newSheet, icon: 'plus', keywords: 'new create', run: () => void createFlow(taken) },
        { group: T.gFile, label: T.import, icon: 'upload', keywords: 'import csv xlsx excel', run: () => document.querySelector<HTMLButtonElement>('.import-btn')?.click() },
        ...TEMPLATES.map((t) => ({ group: T.templates, label: isZh() ? t.title.zh : t.title.en, icon: 'template', keywords: `template ${t.id}`, run: () => void createFlow(taken, { title: t.build(isZh()).title, data: t.build(isZh()), heading: T.useTemplate }) })),
        { group: T.gView, label: T.toggleTheme, icon: 'moon', keywords: 'theme dark light', run: () => themeBtn.click() },
        { group: T.gView, label: T.toggleLang, icon: 'languages', keywords: 'language', run: () => langBtn.click() },
      ],
      T.palette,
    );
  search.addEventListener('click', palette);
  const onKey = (e: KeyboardEvent) => {
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
      e.preventDefault();
      palette();
    }
  };
  window.addEventListener('keydown', onKey);
  const stop = () => {
    window.removeEventListener('keydown', onKey);
    window.removeEventListener('hashchange', stop);
  };
  window.addEventListener('hashchange', stop);
}
