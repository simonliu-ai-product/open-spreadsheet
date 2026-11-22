/**
 * The agent panel's transcript: messages, replies (Markdown), tool rows
 * grouped into steps, questions, permission prompts, errors, the dev server's
 * check and each turn's result. Items are drawn once and redrawn only when
 * they change, so a half-answered question keeps its picks while text streams.
 */
import { type AgentItem, type AgentQuestion, type AgentSession, checkFixPrompt, type RespondPayload } from './agent.ts';
import { T } from './i18n.ts';
import { icon } from './icons.ts';
import { mdToHtml } from './markdown.ts';
import { el } from './ui.ts';

type Tool = Extract<AgentItem, { kind: 'tool' }>;
type Block = { key: string; item: AgentItem } | { key: string; tools: Tool[] };

/** Consecutive tool calls read as one "N steps" row. */
export function toBlocks(items: AgentItem[]): Block[] {
  const blocks: Block[] = [];
  for (const item of items) {
    const last = blocks.at(-1);
    if (item.kind === 'tool' && last && 'tools' in last) last.tools.push(item);
    else if (item.kind === 'tool') blocks.push({ key: item.id, tools: [item] });
    else blocks.push({ key: item.id, item });
  }
  return blocks;
}

/** One reply spans the assistant texts between a message and the turn's result; its copy button sits on the last and copies them all. */
function replyTexts(items: AgentItem[], running: boolean): Map<string, string> {
  const out = new Map<string, string>();
  let parts: { id: string; text: string }[] = [];
  const flush = () => {
    const text = parts
      .map((p) => p.text.trim())
      .filter(Boolean)
      .join('\n\n');
    const last = parts.at(-1);
    if (last && text) out.set(last.id, text);
    parts = [];
  };
  for (const item of items) {
    if (item.kind === 'assistant') parts.push({ id: item.id, text: item.text });
    else if (item.kind === 'user' || item.kind === 'result') flush();
  }
  if (!running) flush();
  return out;
}

const btn = (cls: string, label: string, onClick: () => void) => {
  const b = el('button', cls, label);
  b.type = 'button';
  b.addEventListener('click', onClick);
  return b;
};

/** Runs a click's task once at a time, with the buttons disabled meanwhile. */
const busy = (buttons: HTMLButtonElement[], task: () => Promise<void>) => {
  for (const b of buttons) b.disabled = true;
  void task().finally(() => {
    for (const b of buttons) if (b.isConnected) b.disabled = false;
  });
};

function copyButton(text: string): HTMLButtonElement {
  const b = el('button', 'agent-act');
  b.type = 'button';
  b.title = T.agent.copy;
  b.setAttribute('aria-label', T.agent.copy);
  b.append(icon('copy', 12));
  b.addEventListener('click', () => {
    navigator.clipboard.writeText(text).then(
      () => {
        b.replaceChildren(icon('check', 12));
        b.title = T.agent.copied;
        setTimeout(() => {
          b.replaceChildren(icon('copy', 12));
          b.title = T.agent.copy;
        }, 1500);
      },
      () => {},
    );
  });
  return b;
}

export interface TranscriptHandlers {
  onRespond: (itemId: string, payload: RespondPayload) => Promise<void>;
  onEdit: (itemId: string, text: string) => Promise<void>;
  onFix: (text: string) => Promise<void>;
}

export function createTranscript(h: TranscriptHandlers) {
  const root = el('div', 'agent-transcript');
  const list = el('ol', 'agent-items');
  const working = el('li', 'agent-working');
  working.append(el('span', 'agent-shimmer', T.agent.working));
  let stick = true;
  root.addEventListener('scroll', () => {
    stick = root.scrollHeight - root.scrollTop - root.clientHeight < 48;
  });

  let sessionId: string | null = null;
  /** Drawn blocks by key, with what they were drawn from. */
  const drawn = new Map<string, { sig: unknown[]; node: HTMLElement }>();
  const openGroups = new Set<string>();
  const editing = new Set<string>();

  const same = (a: unknown[], b: unknown[]) => a.length === b.length && a.every((v, i) => v === b[i]);

  function render(session: AgentSession) {
    if (session.id !== sessionId) {
      clear();
      sessionId = session.id;
      stick = true;
    }
    if (!root.contains(list)) root.replaceChildren(list);
    const running = session.status === 'running';
    const replies = replyTexts(session.items, running);
    const lastCheck = [...session.items].reverse().find((i) => i.kind === 'check')?.id;
    const wanted: HTMLElement[] = [];
    const keys = new Set<string>();
    for (const block of toBlocks(session.items)) {
      keys.add(block.key);
      const sig: unknown[] =
        'tools' in block
          ? [...block.tools, openGroups.has(block.key)]
          : [block.item, block.item.kind === 'user' ? running : null, replies.get(block.item.id) ?? null, block.item.id === lastCheck && !running, editing.has(block.item.id)];
      let entry = drawn.get(block.key);
      if (!entry || !same(entry.sig, sig)) {
        const node = el('li');
        node.dataset.kind = 'tools' in block ? 'tools' : block.item.kind;
        node.append('tools' in block ? toolGroup(block.key, block.tools) : itemNode(block.item, running, replies.get(block.item.id), block.item.id === lastCheck && !running));
        entry = { sig, node };
        drawn.set(block.key, entry);
      }
      wanted.push(entry.node);
    }
    for (const k of [...drawn.keys()]) if (!keys.has(k)) drawn.delete(k);
    const last = session.items.at(-1);
    working.firstElementChild!.textContent = T.agent.working;
    if (running && !(last?.kind === 'assistant' && last.streaming)) wanted.push(working);
    // Move only what changed, so focus inside a question form survives.
    wanted.forEach((node, i) => {
      if (list.children[i] !== node) list.insertBefore(node, list.children[i] ?? null);
    });
    while (list.children.length > wanted.length) list.lastElementChild!.remove();
    if (stick) root.scrollTop = root.scrollHeight;
  }

  function clear() {
    drawn.clear();
    openGroups.clear();
    editing.clear();
    list.replaceChildren();
    sessionId = null;
  }

  const redraw = (key: string) => {
    drawn.delete(key);
  };

  function itemNode(item: AgentItem, running: boolean, reply: string | undefined, canFix: boolean): HTMLElement {
    switch (item.kind) {
      case 'user':
        return editing.has(item.id) ? userEditor(item) : userItem(item, running);
      case 'assistant': {
        if (!item.text && item.streaming) return el('span', 'agent-shimmer', T.agent.working);
        const wrap = el('div', 'agent-reply');
        const md = el('div', 'agent-md');
        md.innerHTML = mdToHtml(item.text);
        wrap.append(md);
        if (reply) {
          const acts = el('div', 'agent-acts');
          acts.append(copyButton(reply));
          wrap.append(acts);
        }
        return wrap;
      }
      case 'question':
        return questionItem(item);
      case 'permission':
        return permissionItem(item);
      case 'error': {
        const box = el('div', 'agent-error');
        const eyebrow = el('span', 'agent-eyebrow');
        eyebrow.append(icon('alert', 12), el('span', '', T.agent.error));
        box.append(eyebrow, el('pre', '', item.text));
        return box;
      }
      case 'check':
        return checkItem(item, canFix);
      case 'result': {
        const parts = [item.aborted ? T.agent.stopped : T.agent.done];
        if (item.durationMs !== null) parts.push(T.agent.seconds(Math.max(1, Math.round(item.durationMs / 1000))));
        if (item.costUsd) parts.push(`$${item.costUsd.toFixed(2)}`);
        return el('p', 'agent-result', parts.join(' · '));
      }
    }
    return el('span');
  }

  function userItem(item: Extract<AgentItem, { kind: 'user' }>, running: boolean): HTMLElement {
    const wrap = el('div', 'agent-user');
    wrap.append(el('div', 'agent-bubble', item.text));
    const acts = el('div', 'agent-acts');
    acts.append(copyButton(item.text));
    if (!running) {
      const edit = el('button', 'agent-act');
      edit.type = 'button';
      edit.title = T.agent.edit;
      edit.setAttribute('aria-label', T.agent.edit);
      edit.append(icon('pencil', 12));
      edit.addEventListener('click', () => {
        editing.add(item.id);
        redraw(item.id);
        rerender();
      });
      acts.append(edit);
    }
    wrap.append(acts);
    return wrap;
  }

  function userEditor(item: Extract<AgentItem, { kind: 'user' }>): HTMLElement {
    const box = el('div', 'agent-editor');
    const ta = el('textarea', 'agent-input');
    ta.value = item.text;
    ta.setAttribute('aria-label', T.agent.edit);
    const cancel = btn('btn small', T.cancel, () => {
      editing.delete(item.id);
      redraw(item.id);
      rerender();
    });
    const sendB = btn('btn small primary', T.agent.send, () => submit());
    const submit = () => {
      const text = ta.value.trim();
      if (!text || text === item.text.trim()) return;
      busy([cancel, sendB], async () => {
        await h.onEdit(item.id, text);
        editing.delete(item.id);
      });
    };
    ta.addEventListener('keydown', (e) => {
      if (e.isComposing) return;
      if (e.key === 'Escape') {
        e.preventDefault();
        cancel.click();
      } else if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        submit();
      }
    });
    const row = el('div', 'agent-row-end');
    row.append(cancel, sendB);
    box.append(ta, el('p', 'agent-hint', T.agent.editHint), row);
    requestAnimationFrame(() => {
      ta.focus();
      ta.setSelectionRange(ta.value.length, ta.value.length);
      ta.style.height = `${Math.min(200, ta.scrollHeight)}px`;
    });
    return box;
  }

  function toolRow(t: Tool): HTMLElement {
    const row = el('div', `agent-tool ${t.status}`);
    row.append(el('span', 'agent-dot'), el('span', `agent-tool-title${t.status === 'running' ? ' agent-shimmer' : ''}`, t.title));
    if (t.detail) {
      const d = el('span', 'agent-tool-detail', t.detail);
      d.title = t.detail;
      row.append(d);
    }
    return row;
  }

  function toolGroup(key: string, tools: Tool[]): HTMLElement {
    if (tools.length === 1) return toolRow(tools[0]!);
    const open = openGroups.has(key);
    const wrap = el('div', 'agent-tools');
    const head = el('button', 'agent-steps');
    head.type = 'button';
    head.setAttribute('aria-expanded', String(open));
    const chev = icon('chevronRight', 12);
    head.append(chev, el('span', '', T.agent.steps(tools.length)));
    if (tools.some((t) => t.status === 'error')) head.append(el('span', 'agent-dot error'));
    head.addEventListener('click', () => {
      if (open) openGroups.delete(key);
      else openGroups.add(key);
      rerender();
    });
    wrap.append(head);
    const live = [...tools].reverse().find((t) => t.status === 'running');
    if (open || live) {
      const inner = el('div', 'agent-tools-inner');
      for (const t of open ? tools : [live!]) inner.append(toolRow(t));
      wrap.append(inner);
    }
    return wrap;
  }

  function permissionItem(item: Extract<AgentItem, { kind: 'permission' }>): HTMLElement {
    if (item.status !== 'pending') {
      const p = el('p', `agent-tool ${item.status === 'allowed' ? 'done' : 'error'}`);
      p.append(el('span', 'agent-dot'), el('span', '', item.status === 'allowed' ? T.agent.allowed : T.agent.denied), el('span', 'agent-tool-detail', item.title));
      return p;
    }
    const box = el('div', 'agent-card-box');
    box.append(el('p', 'agent-muted', T.agent.permissionTitle), el('p', 'agent-strong', item.title));
    if (item.detail) box.append(el('pre', 'agent-detail', item.detail));
    const deny = btn('btn small', T.agent.deny, () => busy([deny, allow], () => h.onRespond(item.id, { allow: false })));
    const allow = btn('btn small primary', T.agent.allow, () => busy([deny, allow], () => h.onRespond(item.id, { allow: true })));
    const row = el('div', 'agent-row-end');
    row.append(deny, allow);
    box.append(row);
    return box;
  }

  function questionItem(item: Extract<AgentItem, { kind: 'question' }>): HTMLElement {
    if (item.status !== 'pending') {
      const wrap = el('div', 'agent-answered');
      for (const q of item.questions) {
        const p = el('p', '');
        p.append(el('span', 'agent-muted', q.header || q.question), el('span', 'agent-muted', ' · '), el('span', '', item.status === 'answered' ? item.answers?.[q.question] || '—' : T.agent.questionSkipped));
        wrap.append(p);
      }
      return wrap;
    }
    const picked = new Map<string, string[]>();
    const other = new Map<string, HTMLInputElement>();
    const answerFor = (q: AgentQuestion) => {
      const custom = other.get(q.question)?.value.trim();
      const parts = [...(picked.get(q.question) ?? []), ...(custom ? [custom] : [])];
      return parts.length ? parts.join(', ') : null;
    };
    const box = el('div', 'agent-card-box agent-questions');
    const skip = btn('btn small', T.agent.questionSkip, () => busy([skip, go], () => h.onRespond(item.id, { cancel: true })));
    const go = btn('btn small primary', T.agent.questionSubmit, () => submit());
    const complete = () => item.questions.every((q) => answerFor(q) !== null);
    const paintGo = () => (go.disabled = !complete());
    const submit = () => {
      if (!complete()) return;
      const answers: Record<string, string> = {};
      for (const q of item.questions) answers[q.question] = answerFor(q) ?? '';
      busy([skip, go], () => h.onRespond(item.id, { answers }));
    };
    for (const q of item.questions) {
      const fs = el('fieldset', 'agent-question');
      fs.append(el('legend', '', q.question));
      for (const o of q.options) {
        const b = el('button', 'agent-option');
        b.type = 'button';
        b.setAttribute('aria-pressed', 'false');
        const mark = el('span', `agent-mark${q.multiSelect ? ' multi' : ''}`);
        const text = el('span', 'agent-option-text');
        text.append(el('span', '', o.label));
        if (o.description) text.append(el('span', 'agent-muted', o.description));
        b.append(mark, text);
        b.addEventListener('click', () => {
          const cur = picked.get(q.question) ?? [];
          const on = cur.includes(o.label);
          const next = q.multiSelect ? (on ? cur.filter((l) => l !== o.label) : [...cur, o.label]) : on ? [] : [o.label];
          picked.set(q.question, next);
          for (const sib of fs.querySelectorAll<HTMLButtonElement>('.agent-option')) sib.setAttribute('aria-pressed', String(next.includes(sib.querySelector('.agent-option-text span')!.textContent!)));
          paintGo();
        });
        fs.append(b);
      }
      const input = el('input', 'agent-other');
      input.placeholder = T.agent.questionOther;
      input.addEventListener('input', paintGo);
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.isComposing) {
          e.preventDefault();
          submit();
        }
      });
      other.set(q.question, input);
      fs.append(input);
      box.append(fs);
    }
    paintGo();
    const row = el('div', 'agent-row-end');
    row.append(skip, go);
    box.append(row);
    return box;
  }

  function checkItem(item: Extract<AgentItem, { kind: 'check' }>, canFix: boolean): HTMLElement {
    if (item.status === 'running') {
      const p = el('p', 'agent-check-line');
      p.append(el('span', 'agent-shimmer', T.agent.checkRunning));
      return p;
    }
    if (item.status === 'failed') {
      const p = el('p', 'agent-check-line', T.agent.checkFailed);
      if (item.error && item.error !== 'interrupted') p.append(el('span', 'agent-detail', item.error));
      return p;
    }
    if (item.total === 0) {
      const p = el('p', 'agent-check-line');
      p.append(icon('check', 12), el('span', '', T.agent.checkClean(item.workbooks.length)));
      return p;
    }
    const box = el('div', 'agent-card-box agent-check');
    box.append(el('p', 'agent-muted', T.agent.checkFound(item.total)));
    const ul = el('ul', 'agent-findings');
    const shown = item.findings.slice(0, 5);
    for (const f of shown) {
      const li = el('li', f.severity);
      li.append(el('span', 'agent-dot'));
      const text = el('span', '');
      const where = [item.workbooks.length > 1 ? f.workbook : '', [f.sheet, f.address].filter(Boolean).join('!')].filter(Boolean).join(' · ');
      if (where) text.append(el('span', 'agent-where', where));
      text.append(document.createTextNode(f.message));
      li.append(text);
      ul.append(li);
    }
    box.append(ul);
    if (item.total > shown.length) box.append(el('p', 'agent-muted', T.agent.checkMore(item.total - shown.length)));
    if (canFix) {
      const fix = btn('btn small primary', T.agent.checkFix, () => busy([fix], () => h.onFix(checkFixPrompt(item.findings, item.total))));
      const row = el('div', 'agent-row-end');
      row.append(fix);
      box.append(row);
    }
    return box;
  }

  let last: AgentSession | null = null;
  const rerender = () => last && render(last);
  return {
    el: root,
    render(session: AgentSession) {
      last = session;
      render(session);
    },
    clear() {
      last = null;
      clear();
    },
  };
}
