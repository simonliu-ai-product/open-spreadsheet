// Markdown for notes, both ways: a small renderer (the subset people write in
// notes — headings, paragraphs, lists, quotes, code, rules, bold, italic, code,
// links, images) and a serializer back from the edited DOM. Text is always
// escaped and links are http(s) or mailto only: a note may come from an agent.
// An image shows only when the caller resolves its source (to an asset beside
// the note); anything else stays as its alt text.

/** Where an image's Markdown source points, as a URL to show; null to leave it out. */
export type ImageResolver = (src: string) => string | null;

const esc = (s: string) => s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');

const safeHref = (url: string) => (/^(https?:|mailto:)/i.test(url.trim()) ? url.trim() : '');

/** Inline Markdown in one line of text: `code`, **bold**, *italic*, [text](url), ![alt](image). */
export function inline(text: string, image?: ImageResolver): string {
  const parts: string[] = [];
  // Code spans first, so nothing inside them is read as Markdown.
  const re = /`([^`\n]+)`/g;
  let last = 0;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    parts.push(rest(text.slice(last, m.index)), `<code>${esc(m[1]!)}</code>`);
    last = m.index + m[0].length;
  }
  parts.push(rest(text.slice(last)));
  return parts.join('');

  function rest(t: string): string {
    // Images first, held out as tokens so nothing in their alt text or source is read as Markdown.
    const held: string[] = [];
    return esc(t)
      .replace(/!\[([^\]\n]*)\]\(([^)\s]+)\)/g, (_, alt: string, src: string) => {
        const raw = src.replaceAll('&amp;', '&');
        const url = image?.(raw);
        held.push(url ? `<img class="md-img" src="${esc(url)}" alt="${alt}" data-src="${esc(raw)}">` : alt);
        return `\u0000${held.length - 1}\u0000`;
      })
      .replace(/\[([^\]\n]+)\]\(([^)\s]+)\)/g, (_, label: string, url: string) => {
        const href = safeHref(url.replaceAll('&amp;', '&'));
        return href ? `<a href="${esc(href)}" target="_blank" rel="noopener noreferrer">${label}</a>` : label;
      })
      .replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>')
      .replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, '$1<em>$2</em>')
      .replace(/\u0000(\d+)\u0000/g, (_, k: string) => held[Number(k)] ?? '');
  }
}

const inlineWith = (t: string, image?: ImageResolver) => inline(t, image);

/** A note's Markdown as HTML blocks for the editor. */
export function mdToHtml(md: string, image?: ImageResolver): string {
  const inline = (t: string) => inlineWith(t, image);
  const lines = md.replaceAll('\r\n', '\n').split('\n');
  const out: string[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i]!;
    if (/^```/.test(line)) {
      const body: string[] = [];
      i++;
      while (i < lines.length && !/^```/.test(lines[i]!)) body.push(lines[i++]!);
      i++;
      out.push(`<pre><code>${esc(body.join('\n'))}</code></pre>`);
      continue;
    }
    const h = /^(#{1,3})\s+(.*)$/.exec(line);
    if (h) {
      out.push(`<h${h[1]!.length}>${inline(h[2]!)}</h${h[1]!.length}>`);
      i++;
      continue;
    }
    if (/^\s*(-{3,}|\*{3,})\s*$/.test(line)) {
      out.push('<hr>');
      i++;
      continue;
    }
    if (/^\s*[-*]\s+/.test(line) || /^\s*\d+[.)]\s+/.test(line)) {
      const ordered = /^\s*\d+[.)]\s+/.test(line);
      const items: string[] = [];
      while (i < lines.length && (ordered ? /^\s*\d+[.)]\s+/ : /^\s*[-*]\s+/).test(lines[i]!)) {
        items.push(`<li>${inline(lines[i]!.replace(ordered ? /^\s*\d+[.)]\s+/ : /^\s*[-*]\s+/, ''))}</li>`);
        i++;
      }
      out.push(ordered ? `<ol>${items.join('')}</ol>` : `<ul>${items.join('')}</ul>`);
      continue;
    }
    if (/^>\s?/.test(line)) {
      const body: string[] = [];
      while (i < lines.length && /^>\s?/.test(lines[i]!)) body.push(inline(lines[i++]!.replace(/^>\s?/, '')));
      out.push(`<blockquote><p>${body.join('<br>')}</p></blockquote>`);
      continue;
    }
    if (!line.trim()) {
      i++;
      continue;
    }
    const para: string[] = [];
    while (i < lines.length && lines[i]!.trim() && !/^(#{1,3}\s|```|>|\s*[-*]\s+|\s*\d+[.)]\s+|\s*(-{3,}|\*{3,})\s*$)/.test(lines[i]!)) para.push(inline(lines[i++]!));
    out.push(`<p>${para.join('<br>')}</p>`);
  }
  return out.join('') || '<p><br></p>';
}

/** Inline content of an element back to Markdown. */
function inlineMd(node: Node): string {
  let s = '';
  for (const child of node.childNodes) {
    if (child.nodeType === Node.TEXT_NODE) {
      s += (child.textContent ?? '').replaceAll(' ', ' ');
      continue;
    }
    if (!(child instanceof HTMLElement)) continue;
    const tag = child.tagName;
    const inner = inlineMd(child);
    if (tag === 'BR') s += '\n';
    else if (tag === 'STRONG' || tag === 'B') s += inner.trim() ? `**${inner}**` : inner;
    else if (tag === 'EM' || tag === 'I') s += inner.trim() ? `*${inner}*` : inner;
    else if (tag === 'CODE') s += `\`${child.textContent ?? ''}\``;
    else if (tag === 'IMG') s += child.dataset.src ? `![${child.getAttribute('alt') ?? ''}](${child.dataset.src})` : '';
    else if (tag === 'A') s += safeHref(child.getAttribute('href') ?? '') ? `[${inner}](${child.getAttribute('href')})` : inner;
    else s += inner;
  }
  return s;
}

/** The edited document back to Markdown, one block per paragraph. */
export function htmlToMd(root: HTMLElement): string {
  const blocks: string[] = [];
  const text = (el: Element) => inlineMd(el).replace(/\n+$/, '');
  for (const child of root.childNodes) {
    if (child.nodeType === Node.TEXT_NODE) {
      if ((child.textContent ?? '').trim()) blocks.push(child.textContent!.trim());
      continue;
    }
    if (!(child instanceof HTMLElement)) continue;
    const tag = child.tagName;
    if (/^H[1-6]$/.test(tag)) blocks.push(`${'#'.repeat(Math.min(3, Number(tag[1])))} ${text(child).replaceAll('\n', ' ')}`);
    else if (tag === 'UL' || tag === 'OL') {
      const items = [...child.children].filter((li) => li.tagName === 'LI');
      blocks.push(items.map((li, k) => `${tag === 'OL' ? `${k + 1}.` : '-'} ${text(li).replaceAll('\n', ' ')}`).join('\n'));
    } else if (tag === 'BLOCKQUOTE') blocks.push(text(child).split('\n').map((l) => `> ${l}`).join('\n'));
    else if (tag === 'PRE') blocks.push(`\`\`\`\n${(child.textContent ?? '').replace(/\n$/, '')}\n\`\`\``);
    else if (tag === 'HR') blocks.push('---');
    else {
      const t = text(child);
      if (t.trim()) blocks.push(t.trim());
    }
  }
  return `${blocks.join('\n\n')}\n`;
}
