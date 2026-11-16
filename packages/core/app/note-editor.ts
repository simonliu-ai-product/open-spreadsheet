// A note edited as it reads, Medium-style: Markdown typed at the start of a
// line becomes the block (#, ##, ###, -, 1., >, ```, ---), and **bold**,
// *italic* and `code` turn into marks as they close. What is saved is Markdown.
import { htmlToMd, type ImageResolver, mdToHtml } from './markdown.ts';
import { el } from './ui.ts';

export interface NoteEditor {
  el: HTMLElement;
  setMarkdown(md: string): void;
  getMarkdown(): string;
  focus(): void;
  /** Puts a line of Markdown (an image) after the block the caret was last in. */
  insertBlock(md: string): void;
}

function caretAtEnd(node: Node) {
  const range = document.createRange();
  range.selectNodeContents(node);
  range.collapse(false);
  const sel = getSelection();
  sel?.removeAllRanges();
  sel?.addRange(range);
}

export function createNoteEditor(opts: { onChange: () => void; placeholder: string; image?: () => ImageResolver }): NoteEditor {
  const doc = el('div', 'note-doc');
  doc.contentEditable = 'true';
  doc.spellcheck = false;
  doc.dataset.placeholder = opts.placeholder;

  /** The top-level block the caret is in. */
  const blockAt = (): HTMLElement | null => {
    let n = getSelection()?.anchorNode ?? null;
    while (n && n.parentNode !== doc) n = n.parentNode;
    return n instanceof HTMLElement ? n : null;
  };
  const replaceBlock = (old: HTMLElement, tag: string, keep: Node[]): HTMLElement => {
    const next = el(tag as 'p');
    next.append(...keep);
    if (!next.textContent) next.append(el('br'));
    old.replaceWith(next);
    return next;
  };
  const emptyEnd = () => {
    const sel = getSelection();
    const block = blockAt();
    if (!sel || !block || !sel.isCollapsed) return false;
    const r = document.createRange();
    r.selectNodeContents(block);
    r.setStart(sel.anchorNode!, sel.anchorOffset);
    return !r.toString();
  };
  /** The block the caret was last in, so a toolbar click can insert after it. */
  let lastBlock: HTMLElement | null = null;
  const remember = () => {
    const b = blockAt();
    if (b) lastBlock = b;
  };
  doc.addEventListener('keyup', remember);
  doc.addEventListener('mouseup', remember);
  const isEmpty = (n: HTMLElement) => !(n.textContent ?? '').replace(/​/g, '').trim();

  /** Block shortcuts: a marker and a space at the start of a paragraph. */
  function blockRules(): boolean {
    const block = blockAt();
    if (!block || !/^(P|DIV)$/.test(block.tagName)) return false;
    const text = (block.textContent ?? '').replaceAll(' ', ' ');
    const m = /^(#{1,3}|[-*]|\d+[.)]|>) (.*)$/s.exec(text);
    if (!m) return false;
    const marker = m[1]!;
    // Drop the marker from the block's first text, keep the rest (and any marks) as they are.
    const first = block.firstChild;
    if (first?.nodeType === Node.TEXT_NODE) first.textContent = (first.textContent ?? '').replace(/^\S+[  ]/, '');
    const keep = [...block.childNodes].filter((n) => !(n instanceof HTMLBRElement) || block.childNodes.length > 1);
    let target: HTMLElement;
    if (marker.startsWith('#')) target = replaceBlock(block, `h${marker.length}`, keep);
    else if (marker === '>') {
      const q = el('blockquote');
      const p = el('p');
      p.append(...keep);
      if (!p.textContent) p.append(el('br'));
      q.append(p);
      block.replaceWith(q);
      target = p;
    } else {
      const list = el(/\d/.test(marker) ? 'ol' : 'ul');
      const li = el('li');
      li.append(...keep);
      if (!li.textContent) li.append(el('br'));
      list.append(li);
      block.replaceWith(list);
      target = li;
    }
    caretAtEnd(target);
    return true;
  }

  /** Inline marks: **bold**, *italic*, `code` turn into elements as they close. */
  function inlineRules() {
    const sel = getSelection();
    const node = sel?.anchorNode;
    if (!sel || !node || node.nodeType !== Node.TEXT_NODE || !sel.isCollapsed) return;
    if ((node.parentElement?.closest('code, pre') ?? null) !== null) return;
    const before = (node.textContent ?? '').slice(0, sel.anchorOffset);
    const rules: [RegExp, string][] = [
      [/\*\*([^*]+)\*\*$/, 'strong'],
      [/`([^`]+)`$/, 'code'],
      [/(?:^|[^*])\*([^*\s][^*]*)\*$/, 'em'],
    ];
    for (const [re, tag] of rules) {
      const m = re.exec(before);
      if (!m) continue;
      const inner = m[1]!;
      const start = before.length - (tag === 'em' ? inner.length + 2 : m[0].length);
      const text = node as Text;
      const after = text.splitText(start);
      after.textContent = (after.textContent ?? '').slice(m[0].length - (tag === 'em' && m[0].length > inner.length + 2 ? 1 : 0));
      const mark = el(tag as 'strong', '', inner);
      const tail = document.createTextNode('​');
      after.before(mark, tail);
      const r = document.createRange();
      r.setStart(tail, 1);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      return;
    }
  }

  doc.addEventListener('input', () => {
    if (!blockRules()) inlineRules();
    opts.onChange();
  });

  doc.addEventListener('keydown', (e) => {
    const block = blockAt();
    if (e.key === 'Enter' && !e.shiftKey && block) {
      const text = (block.textContent ?? '').trim();
      if (/^(P|DIV)$/.test(block.tagName) && (text === '---' || text === '***')) {
        e.preventDefault();
        const hr = el('hr');
        const p = el('p');
        p.append(el('br'));
        block.replaceWith(hr);
        hr.after(p);
        caretAtEnd(p);
        opts.onChange();
        return;
      }
      if (/^(P|DIV)$/.test(block.tagName) && text === '```') {
        e.preventDefault();
        const pre = el('pre');
        const code = el('code');
        code.append(document.createTextNode('\n'));
        pre.append(code);
        block.replaceWith(pre);
        caretAtEnd(code);
        opts.onChange();
        return;
      }
      if (block.tagName === 'PRE') {
        e.preventDefault();
        document.execCommand('insertText', false, '\n');
        return;
      }
      // A heading or quote ends on Enter: the next line is a paragraph.
      if (/^H[1-3]$|^BLOCKQUOTE$/.test(block.tagName) && emptyEnd()) {
        e.preventDefault();
        const p = el('p');
        p.append(el('br'));
        block.after(p);
        caretAtEnd(p);
        opts.onChange();
        return;
      }
      // Enter on an empty list item leaves the list.
      const li = (getSelection()?.anchorNode as Node | null)?.parentElement?.closest('li') ?? (getSelection()?.anchorNode instanceof HTMLLIElement ? (getSelection()!.anchorNode as HTMLLIElement) : null);
      if (li && isEmpty(li) && (block.tagName === 'UL' || block.tagName === 'OL')) {
        e.preventDefault();
        const p = el('p');
        p.append(el('br'));
        li.remove();
        block.after(p);
        if (!block.querySelector('li')) block.remove();
        caretAtEnd(p);
        opts.onChange();
      }
    }
    // Backspace at the start of a heading or quote turns it back into a paragraph.
    if (e.key === 'Backspace' && block && /^H[1-3]$|^BLOCKQUOTE$/.test(block.tagName)) {
      const sel = getSelection();
      if (sel?.isCollapsed) {
        const r = document.createRange();
        r.selectNodeContents(block);
        r.setEnd(sel.anchorNode!, sel.anchorOffset);
        if (!r.toString()) {
          e.preventDefault();
          const inner = block.tagName === 'BLOCKQUOTE' ? (block.firstElementChild ?? block) : block;
          const p = replaceBlock(block, 'p', [...inner.childNodes]);
          const start = document.createRange();
          start.setStart(p, 0);
          start.collapse(true);
          sel.removeAllRanges();
          sel.addRange(start);
          opts.onChange();
        }
      }
    }
  });

  // Plain text only: no styles from a web page, and a pasted Markdown note reads as Markdown.
  doc.addEventListener('paste', (e) => {
    e.preventDefault();
    const text = e.clipboardData?.getData('text/plain') ?? '';
    if (!text) return;
    if (/\n/.test(text) && /^(#{1,3}\s|[-*]\s|\d+[.)]\s|>\s|```)/m.test(text)) {
      document.execCommand('insertHTML', false, mdToHtml(text, opts.image?.()));
    } else document.execCommand('insertText', false, text);
  });
  doc.addEventListener('focus', () => document.execCommand('defaultParagraphSeparator', false, 'p'));

  return {
    el: doc,
    setMarkdown(md) {
      doc.innerHTML = mdToHtml(md, opts.image?.());
      lastBlock = null;
    },
    insertBlock(md) {
      const holder = el('div');
      holder.innerHTML = mdToHtml(md, opts.image?.());
      const block = holder.firstElementChild as HTMLElement | null;
      if (!block) return;
      const after = lastBlock && lastBlock.parentNode === doc ? lastBlock : null;
      // An empty paragraph is replaced rather than left above the picture.
      if (after && isEmpty(after) && !after.querySelector('img')) after.replaceWith(block);
      else if (after) after.after(block);
      else doc.append(block);
      const next = el('p');
      next.append(el('br'));
      if (!block.nextElementSibling) block.after(next);
      doc.focus();
      caretAtEnd(block.nextElementSibling ?? block);
      lastBlock = block;
      opts.onChange();
    },
    getMarkdown: () => htmlToMd(doc).replaceAll('​', ''),
    focus() {
      doc.focus();
      caretAtEnd(doc.lastElementChild ?? doc);
    },
  };
}
