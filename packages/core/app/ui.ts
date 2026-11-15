/** Small DOM toolkit: elements, menus, popovers, dialogs, toasts, theme. No framework. */
import { T } from './i18n.ts';
import { icon } from './icons.ts';

export function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls = '', text = ''): HTMLElementTagNameMap[K] {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text) n.textContent = text;
  return n;
}

/** A toolbar or header button: an icon, an optional label, a tooltip. */
export function iconButton(name: string, title: string, onClick?: (e: MouseEvent) => void, label?: string): HTMLButtonElement {
  const b = el('button', `ibtn${label ? ' with-label' : ''}`);
  b.type = 'button';
  b.title = title;
  b.setAttribute('aria-label', title);
  b.append(icon(name));
  if (label) b.append(el('span', '', label));
  if (onClick) b.addEventListener('click', onClick);
  return b;
}

// ---------- menus ----------

export interface MenuItem {
  label: string;
  icon?: string;
  shortcut?: string;
  action?: () => void;
  disabled?: boolean;
  danger?: boolean;
  checked?: boolean;
  /** A row of colour swatches instead of a plain item. */
  swatches?: { colors: (string | null)[]; current?: string | null; pick: (c: string | null) => void };
  submenu?: MenuEntry[];
}
export type MenuEntry = MenuItem | '-';

let openMenu: { root: HTMLElement; close: () => void } | null = null;

export function closeMenus(): void {
  openMenu?.close();
}

/** Opens a menu at a point (a context menu) or under an element (a dropdown). */
export function showMenu(entries: MenuEntry[], at: { x: number; y: number } | HTMLElement, onClose?: () => void): void {
  closeMenus();
  const layer = el('div', 'menu-layer');
  const stack: HTMLElement[] = [];

  const build = (items: MenuEntry[], x: number, y: number, depth: number) => {
    while (stack.length > depth) stack.pop()!.remove();
    const menu = el('div', 'menu');
    menu.setAttribute('role', 'menu');
    for (const item of items) {
      if (item === '-') {
        menu.append(el('div', 'menu-sep'));
        continue;
      }
      if (item.swatches) {
        const row = el('div', 'menu-swatches');
        row.append(el('span', 'menu-swatch-label', item.label));
        const grid = el('div', 'swatch-grid');
        for (const c of item.swatches.colors) {
          const s = el('button', `swatch${c === null ? ' none' : ''}${c === (item.swatches.current ?? null) ? ' current' : ''}`);
          s.type = 'button';
          s.title = c ?? T.noColor;
          if (c) s.style.background = c;
          s.addEventListener('click', () => {
            close();
            item.swatches!.pick(c);
          });
          grid.append(s);
        }
        row.append(grid);
        menu.append(row);
        continue;
      }
      const b = el('button', `menu-item${item.danger ? ' danger' : ''}`);
      b.type = 'button';
      b.setAttribute('role', 'menuitem');
      b.disabled = Boolean(item.disabled);
      const ic = el('span', 'menu-icon');
      if (item.checked) ic.append(icon('check', 14));
      else if (item.icon) ic.append(icon(item.icon, 15));
      b.append(ic, el('span', 'menu-label', item.label));
      if (item.shortcut) b.append(el('span', 'menu-shortcut', item.shortcut));
      if (item.submenu) {
        b.append(el('span', 'menu-shortcut', '›'));
        b.addEventListener('mouseenter', () => {
          const r = b.getBoundingClientRect();
          build(item.submenu!, r.right - 4, r.top - 5, depth + 1);
        });
      } else {
        b.addEventListener('mouseenter', () => {
          while (stack.length > depth + 1) stack.pop()!.remove();
        });
        b.addEventListener('click', () => {
          close();
          item.action?.();
        });
      }
      menu.append(b);
    }
    layer.append(menu);
    stack.push(menu);
    // Keep it on screen.
    const r = menu.getBoundingClientRect();
    const left = Math.max(6, Math.min(x, innerWidth - r.width - 6));
    const top = y + r.height > innerHeight - 6 ? Math.max(6, innerHeight - r.height - 6) : y;
    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;
    menu.querySelector<HTMLButtonElement>('.menu-item:not(:disabled)')?.focus({ preventScroll: true });
    return menu;
  };

  document.body.append(layer);
  let x: number;
  let y: number;
  if (at instanceof HTMLElement) {
    const r = at.getBoundingClientRect();
    x = r.left;
    y = r.bottom + 4;
  } else {
    x = at.x;
    y = at.y;
  }
  build(entries, x, y, 0);

  const onKey = (e: KeyboardEvent) => {
    const menu = stack[stack.length - 1]!;
    const items = [...menu.querySelectorAll<HTMLButtonElement>('.menu-item:not(:disabled)')];
    const i = items.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      if (stack.length > 1) stack.pop()!.remove();
      else close();
    } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      e.stopPropagation();
      const n = items.length;
      items[(i + (e.key === 'ArrowDown' ? 1 : -1) + n) % n]?.focus();
    } else if (e.key === 'ArrowRight' && document.activeElement instanceof HTMLElement) {
      e.preventDefault();
      document.activeElement.dispatchEvent(new MouseEvent('mouseenter'));
      stack[stack.length - 1]?.querySelector<HTMLButtonElement>('.menu-item')?.focus();
    } else if (e.key === 'ArrowLeft' && stack.length > 1) {
      e.preventDefault();
      stack.pop()!.remove();
    }
  };
  const onDown = (e: MouseEvent) => {
    if (!stack.some((m) => m.contains(e.target as Node))) close();
  };
  const close = () => {
    if (!layer.isConnected) return;
    layer.remove();
    window.removeEventListener('keydown', onKey, true);
    window.removeEventListener('mousedown', onDown, true);
    window.removeEventListener('blur', close);
    openMenu = null;
    onClose?.();
  };
  window.addEventListener('keydown', onKey, true);
  window.addEventListener('mousedown', onDown, true);
  window.addEventListener('blur', close);
  openMenu = { root: layer, close };
}

// ---------- dialogs ----------

export interface Field {
  name: string;
  label: string;
  value?: string;
  hint?: string;
  placeholder?: string;
  mono?: boolean;
  type?: 'text' | 'number';
  /** A textarea instead of a one-line input. */
  multiline?: boolean;
  /** A dropdown of these choices instead of a text input. */
  options?: { value: string; label: string }[];
  /** An error message, or null when the value is fine. */
  validate?: (v: string, all: Record<string, string>) => string | null;
  /** Called on input to update other fields (e.g. a file name that follows the title). */
  derive?: (v: string, fields: Record<string, HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => void;
}

/** A modal form. Resolves with the values, or null when cancelled. */
export function promptDialog(opts: { title: string; body?: string; fields: Field[]; ok?: string; danger?: boolean }): Promise<Record<string, string> | null> {
  return new Promise((resolve) => {
    const back = el('div', 'dialog-back');
    const box = el('form', 'dialog');
    box.setAttribute('role', 'dialog');
    box.setAttribute('aria-modal', 'true');
    box.append(el('h2', '', opts.title));
    if (opts.body) box.append(el('p', 'dialog-body', opts.body));
    const inputs: Record<string, HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement> = {};
    const errors: Record<string, HTMLElement> = {};
    for (const f of opts.fields) {
      const label = el('label', 'field');
      label.append(el('span', 'field-label', f.label));
      const input = f.options ? el('select', 'input select') : f.multiline ? el('textarea', 'input textarea') : el('input', `input${f.mono ? ' mono' : ''}`);
      if (input instanceof HTMLSelectElement) for (const o of f.options ?? []) input.append(new Option(o.label, o.value));
      input.value = f.value ?? '';
      if (input instanceof HTMLSelectElement) {
        // a choice needs no typing aids
      } else if (input instanceof HTMLInputElement) {
        input.type = f.type ?? 'text';
        input.autocomplete = 'off';
      } else {
        input.rows = 4;
        // ⌘Enter submits from a textarea; Enter adds a line.
        input.addEventListener('keydown', (e) => {
          if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
            e.preventDefault();
            box.requestSubmit();
          }
        });
      }
      input.spellcheck = false;
      if (f.placeholder && !(input instanceof HTMLSelectElement)) input.placeholder = f.placeholder;
      inputs[f.name] = input;
      label.append(input);
      if (f.hint) label.append(el('span', 'field-hint', f.hint));
      const e = el('span', 'field-error');
      e.hidden = true;
      errors[f.name] = e;
      label.append(e);
      box.append(label);
    }
    for (const f of opts.fields) if (f.derive) inputs[f.name]!.addEventListener(f.options ? 'change' : 'input', () => f.derive!(inputs[f.name]!.value, inputs));
    const actions = el('div', 'dialog-actions');
    const cancel = el('button', 'btn', T.cancel);
    cancel.type = 'button';
    const ok = el('button', `btn ${opts.danger ? 'danger' : 'primary'}`, opts.ok ?? T.ok);
    ok.type = 'submit';
    actions.append(cancel, ok);
    box.append(actions);
    back.append(box);
    document.body.append(back);
    const first = Object.values(inputs)[0];
    if (first) {
      first.focus();
      if (!(first instanceof HTMLSelectElement)) first.select();
    } else ok.focus();

    const done = (v: Record<string, string> | null) => {
      back.remove();
      window.removeEventListener('keydown', onKey, true);
      resolve(v);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        done(null);
      }
    };
    window.addEventListener('keydown', onKey, true);
    cancel.addEventListener('click', () => done(null));
    back.addEventListener('mousedown', (e) => e.target === back && done(null));
    box.addEventListener('submit', (e) => {
      e.preventDefault();
      const values = Object.fromEntries(Object.entries(inputs).map(([k, i]) => [k, i.value]));
      let bad = false;
      for (const f of opts.fields) {
        const msg = f.validate?.(values[f.name]!, values) ?? null;
        errors[f.name]!.textContent = msg ?? '';
        errors[f.name]!.hidden = !msg;
        if (msg && !bad) {
          inputs[f.name]!.focus();
          bad = true;
        }
      }
      if (!bad) done(values);
    });
  });
}

export async function confirmDialog(title: string, body: string, ok: string, danger = false): Promise<boolean> {
  return (await promptDialog({ title, body, fields: [], ok, danger })) !== null;
}

// ---------- toast ----------

let toastTimer: ReturnType<typeof setTimeout> | null = null;
export function toast(message: string): void {
  let t = document.querySelector<HTMLElement>('.toast');
  if (!t) {
    t = el('div', 'toast');
    t.setAttribute('role', 'status');
    document.body.append(t);
  }
  t.textContent = message;
  t.classList.add('show');
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t!.classList.remove('show'), 2400);
}

// ---------- theme ----------

export type Theme = 'light' | 'dark';
const media = matchMedia('(prefers-color-scheme: dark)');

export function currentTheme(): Theme {
  try {
    const saved = localStorage.getItem('open-spreadsheet:theme');
    if (saved === 'light' || saved === 'dark') return saved;
  } catch {
    // Fall back to the system setting.
  }
  return media.matches ? 'dark' : 'light';
}

export function applyTheme(t: Theme = currentTheme()): void {
  document.documentElement.classList.toggle('dark', t === 'dark');
  document.documentElement.style.colorScheme = t;
}

export function toggleTheme(): Theme {
  const next: Theme = currentTheme() === 'dark' ? 'light' : 'dark';
  try {
    localStorage.setItem('open-spreadsheet:theme', next);
  } catch {
    // Applies now even if not remembered.
  }
  applyTheme(next);
  return next;
}
media.addEventListener('change', () => applyTheme());

/** Saves a blob as a file. */
export function downloadBlob(blob: Blob, name: string): void {
  const a = el('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

/** Asks for a file with the browser's picker. */
export function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = el('input');
    input.type = 'file';
    input.accept = accept;
    input.addEventListener('change', () => resolve(input.files?.[0] ?? null));
    input.click();
  });
}

export const isMac = /Mac|iPhone|iPad/.test(navigator.platform);
export const MOD = isMac ? '⌘' : 'Ctrl+';

// ---------- command palette ----------

export interface Command {
  label: string;
  group: string;
  icon?: string;
  hint?: string;
  /** Extra words to match, e.g. the English name in a Chinese UI. */
  keywords?: string;
  run: () => void;
}

/** ⌘K: type to filter every action, ↑↓ to choose, Enter to run. */
export function commandPalette(commands: Command[], placeholder: string, onClose?: () => void): void {
  closeMenus();
  document.querySelector('.palette-back')?.remove();
  const back = el('div', 'palette-back');
  const box = el('div', 'palette');
  box.setAttribute('role', 'dialog');
  const field = el('div', 'palette-field');
  const input = el('input');
  input.placeholder = placeholder;
  input.spellcheck = false;
  field.append(icon('search', 16), input, el('kbd', '', 'esc'));
  const list = el('div', 'palette-list');
  list.setAttribute('role', 'listbox');
  box.append(field, list);
  back.append(box);
  document.body.append(back);
  input.focus();

  let shown: Command[] = [];
  let index = 0;
  const paint = () => {
    const q = input.value.trim().toLowerCase();
    const words = q.split(/\s+/).filter(Boolean);
    shown = commands.filter((c) => {
      const hay = `${c.label} ${c.group} ${c.keywords ?? ''}`.toLowerCase();
      return words.every((w) => hay.includes(w));
    });
    index = Math.min(index, Math.max(0, shown.length - 1));
    list.replaceChildren();
    let group = '';
    shown.slice(0, 80).forEach((c, i) => {
      if (c.group !== group) {
        group = c.group;
        list.append(el('div', 'palette-group', group));
      }
      const row = el('div', `palette-item${i === index ? ' on' : ''}`);
      row.setAttribute('role', 'option');
      const ic = el('span', 'menu-icon');
      if (c.icon) ic.append(icon(c.icon, 15));
      row.append(ic, el('span', 'menu-label', c.label));
      if (c.hint) row.append(el('span', 'menu-shortcut', c.hint));
      row.addEventListener('mousemove', () => {
        if (index !== i) {
          index = i;
          paint();
        }
      });
      row.addEventListener('mousedown', (e) => {
        e.preventDefault();
        index = i;
        run();
      });
      list.append(row);
    });
    if (!shown.length) list.append(el('div', 'palette-empty', '—'));
    list.querySelector('.palette-item.on')?.scrollIntoView({ block: 'nearest' });
  };
  const close = () => {
    back.remove();
    onClose?.();
  };
  const run = () => {
    const c = shown[index];
    close();
    c?.run();
  };
  input.addEventListener('input', () => {
    index = 0;
    paint();
  });
  input.addEventListener('keydown', (e) => {
    if (e.isComposing) return;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      const n = Math.min(shown.length, 80);
      if (n) index = (index + (e.key === 'ArrowDown' ? 1 : -1) + n) % n;
      paint();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      run();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  });
  back.addEventListener('mousedown', (e) => e.target === back && close());
  paint();
}
