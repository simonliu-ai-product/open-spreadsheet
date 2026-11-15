// The Assets panel: a workbook's images and Markdown notes (sheets/<id>/assets/)
// and the workspace's shared ones (assets/). Charts are made here as images
// rather than placed on the grid; notes are what people keep about the numbers,
// and agents read them as context.
import { T } from './i18n.ts';
import { icon } from './icons.ts';
import { confirmDialog, el, iconButton, pickFile, promptDialog, toast } from './ui.ts';

export const SHARED = 'shared';

export interface AssetEntry {
  name: string;
  kind: 'image' | 'note';
  size: number;
  updatedAt: string;
  title?: string;
}

const api = (path: string, params: Record<string, string>) => `/__api/assets${path}?${new URLSearchParams(params).toString()}`;

export async function listAssets(scope: string): Promise<AssetEntry[]> {
  const res = await fetch(api('', { scope }));
  if (!res.ok) return [];
  return ((await res.json()) as { assets: AssetEntry[] }).assets;
}

/** Stores an image; a taken name comes back numbered. */
export async function uploadAsset(scope: string, name: string, body: Blob | string, type = 'image/svg+xml'): Promise<string> {
  const res = await fetch(api('/upload', { scope, name }), { method: 'POST', headers: { 'content-type': type }, body });
  const out = (await res.json()) as { name?: string; error?: string };
  if (!res.ok) throw new Error(out.error ?? res.statusText);
  return out.name!;
}

async function readNote(scope: string, name: string): Promise<string> {
  const res = await fetch(api('/file', { scope, name }));
  return res.ok ? res.text() : '';
}

async function writeNote(scope: string, name: string, text: string): Promise<void> {
  const res = await fetch('/__api/assets/note', { method: 'PUT', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ scope, name, text }) });
  if (!res.ok) throw new Error(((await res.json()) as { error?: string }).error ?? res.statusText);
}

async function removeAsset(scope: string, name: string): Promise<void> {
  const res = await fetch(api('', { scope, name }), { method: 'DELETE' });
  if (!res.ok) throw new Error(((await res.json()) as { error?: string }).error ?? res.statusText);
}

/** A file name for a note from its title: letters of any script kept, the rest dashes. */
export const noteName = (title: string) =>
  `${
    title
      .normalize('NFKC')
      .replace(/[^\p{L}\p{N}]+/gu, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 60) || 'note'
  }.md`;

export interface AssetsPanel {
  el: HTMLElement;
  open(): void;
  close(): void;
  toggle(): void;
  isOpen(): boolean;
  refresh(): Promise<void>;
  /** Called with a chart's file name once it is saved, so it can be shown first. */
  showSaved(scope: string, name: string): Promise<void>;
}

export function createAssetsPanel(opts: { bookId: string; onClose?: () => void; newChart: (at: HTMLElement) => void }): AssetsPanel {
  const panel = el('aside', 'assets-panel');
  panel.hidden = true;
  panel.setAttribute('aria-label', T.assets);
  const head = el('div', 'assets-head');
  head.append(el('h2', '', T.assets), el('div', 'grow'), iconButton('x', T.close, () => close()));
  const actions = el('div', 'assets-actions');
  const chartBtn = el('button', 'btn', '');
  chartBtn.type = 'button';
  chartBtn.append(icon('chart', 14), el('span', '', T.newChartImage));
  chartBtn.addEventListener('click', () => opts.newChart(chartBtn));
  const noteBtn = el('button', 'btn', '');
  noteBtn.type = 'button';
  noteBtn.append(icon('pencil', 14), el('span', '', T.newNote));
  noteBtn.addEventListener('click', () => void newNote());
  const uploadBtn = el('button', 'btn', '');
  uploadBtn.type = 'button';
  uploadBtn.append(icon('upload', 14), el('span', '', T.uploadImage));
  uploadBtn.addEventListener('click', () => void upload());
  actions.append(chartBtn, noteBtn, uploadBtn);
  const list = el('div', 'assets-list');
  panel.append(head, el('p', 'assets-hint', T.assetsHint), actions, list);

  let highlight = '';

  function card(scope: string, a: AssetEntry): HTMLElement {
    const c = el('div', `asset-card ${a.kind}`);
    c.dataset.name = a.name;
    if (`${scope}/${a.name}` === highlight) c.classList.add('fresh');
    const src = api('/file', { scope, name: a.name, v: a.updatedAt });
    if (a.kind === 'image') {
      const img = el('img');
      img.src = src;
      img.alt = a.name;
      img.loading = 'lazy';
      const open = el('a', 'asset-preview');
      open.href = src;
      open.target = '_blank';
      open.rel = 'noopener';
      open.append(img);
      c.append(open);
    } else {
      const body = el('button', 'asset-note');
      body.type = 'button';
      body.append(icon('file', 14), el('span', 'asset-note-title', a.title || a.name));
      body.addEventListener('click', () => void editNote(scope, a.name));
      c.append(body);
    }
    const foot = el('div', 'asset-foot');
    foot.append(el('span', 'asset-name', a.name));
    const dl = el('a', 'icon-link');
    dl.href = api('/file', { scope, name: a.name, download: '1' });
    dl.title = T.download;
    dl.append(icon('download', 13));
    const del = iconButton('trash', T.delete, async () => {
      if (!(await confirmDialog(T.deleteAssetTitle(a.name), T.deleteAssetBody, T.delete, true))) return;
      try {
        await removeAsset(scope, a.name);
        await refresh();
      } catch (e) {
        toast((e as Error).message);
      }
    });
    foot.append(dl, del);
    c.append(foot);
    return c;
  }

  async function refresh() {
    const [mine, shared] = await Promise.all([listAssets(opts.bookId), listAssets(SHARED)]);
    const section = (title: string, scope: string, items: AssetEntry[]) => {
      const s = el('section', 'assets-section');
      s.dataset.scope = scope;
      s.append(el('h3', '', title));
      if (!items.length) s.append(el('p', 'assets-empty', T.noAssets));
      const notes = items.filter((a) => a.kind === 'note');
      const images = items.filter((a) => a.kind === 'image');
      for (const a of [...notes, ...images]) s.append(card(scope, a));
      return s;
    };
    list.replaceChildren(section(T.assetsThisBook, opts.bookId, mine), section(T.assetsShared, SHARED, shared));
  }

  async function newNote() {
    const v = await promptDialog({
      title: T.newNote,
      body: T.noteContextHint,
      fields: [
        { name: 'title', label: T.title, value: '', validate: (x) => (x.trim() ? null : T.title) },
        { name: 'text', label: T.noteLabel, value: '', multiline: true },
        { name: 'scope', label: T.saveIn, options: [{ value: opts.bookId, label: T.assetsThisBook }, { value: SHARED, label: T.assetsShared }], value: opts.bookId },
      ],
      ok: T.save,
    });
    if (!v) return;
    const title = v.title!.trim();
    const text = v.text!.trim() ? `# ${title}\n\n${v.text!.trim()}\n` : `# ${title}\n`;
    try {
      await writeNote(v.scope!, noteName(title), text);
      highlight = `${v.scope}/${noteName(title)}`;
      await refresh();
    } catch (e) {
      toast((e as Error).message);
    }
  }

  async function editNote(scope: string, name: string) {
    const text = await readNote(scope, name);
    const v = await promptDialog({ title: name, body: T.noteContextHint, fields: [{ name: 'text', label: T.noteLabel, value: text, multiline: true }], ok: T.save });
    if (!v) return;
    try {
      await writeNote(scope, name, v.text!);
      await refresh();
    } catch (e) {
      toast((e as Error).message);
    }
  }

  async function upload() {
    const file = await pickFile('image/png,image/jpeg,image/gif,image/webp,image/svg+xml');
    if (!file) return;
    try {
      const name = await uploadAsset(opts.bookId, file.name.replace(/[^\p{L}\p{N}_ .()-]+/gu, '-'), file, file.type || 'application/octet-stream');
      highlight = `${opts.bookId}/${name}`;
      await refresh();
    } catch (e) {
      toast((e as Error).message);
    }
  }

  function open() {
    panel.hidden = false;
    void refresh();
  }
  function close() {
    panel.hidden = true;
    opts.onClose?.();
  }
  return {
    el: panel,
    open,
    close,
    toggle: () => (panel.hidden ? open() : close()),
    isOpen: () => !panel.hidden,
    refresh,
    async showSaved(scope, name) {
      highlight = `${scope}/${name}`;
      panel.hidden = false;
      await refresh();
      panel.querySelector(`.asset-card[data-name="${CSS.escape(name)}"]`)?.scrollIntoView({ block: 'nearest' });
    },
  };
}
