/**
 * The agent panel: a right-hand column beside the workbook (or the home page)
 * where the person drives Codex. Header with the model menu, history and a new
 * chat; a connect view that finds the CLI and installs the SDK; the transcript
 * (agent-transcript.ts); and the composer. Toggled with the header button or ⌘J.
 */
import {
  agentAvailable,
  type AgentProviderStatus,
  type AgentSessionSummary,
  createSession,
  currentProvider,
  deleteSession,
  getState,
  installProvider,
  loadSession,
  loadSessions,
  readyProviders,
  refreshStatus,
  resolveModelChoice,
  respond,
  sendMessage,
  setActiveSession,
  setAutoApprove,
  setModelChoice,
  setOpen,
  stopSession,
  subscribe,
  toggleOpen,
} from './agent.ts';
import { createTranscript } from './agent-transcript.ts';
import { locale, T } from './i18n.ts';
import { icon } from './icons.ts';
import { el, iconButton, type MenuEntry, MOD, showMenu, toast } from './ui.ts';

const CLI_INSTALL: Record<string, string> = { codex: 'npm install -g @openai/codex' };

/** The header button that shows and hides the panel; null when the panel is not offered. */
export function agentToggleButton(): HTMLButtonElement | null {
  if (!agentAvailable()) return null;
  const b = iconButton('sparkles', `${T.agent.toggle} (${MOD}J)`, () => toggleOpen());
  b.classList.add('agent-toggle');
  const paint = () => {
    b.classList.toggle('on', getState().open);
    b.setAttribute('aria-pressed', String(getState().open));
  };
  paint();
  let seen = false;
  const off = subscribe(() => {
    if (b.isConnected) seen = true;
    else if (seen) return off();
    paint();
  });
  return b;
}

const RELATIVE: [Intl.RelativeTimeFormatUnit, number][] = [
  ['year', 365 * 86400],
  ['month', 30 * 86400],
  ['week', 7 * 86400],
  ['day', 86400],
  ['hour', 3600],
  ['minute', 60],
];
function relativeTime(iso: string): string {
  const seconds = (Date.parse(iso) - Date.now()) / 1000;
  const rtf = new Intl.RelativeTimeFormat(locale(), { numeric: 'auto', style: 'short' });
  for (const [unit, size] of RELATIVE) if (Math.abs(seconds) >= size) return rtf.format(Math.round(seconds / size), unit);
  return rtf.format(0, 'second');
}

/** Puts the panel at the end of `layout` (a row: the page, then the panel). */
export function mountAgentPanel(layout: HTMLElement): void {
  if (!agentAvailable()) return;
  const panel = el('aside', 'agent-panel');
  panel.setAttribute('aria-label', T.agent.title);
  layout.append(panel);

  let view: 'chat' | 'history' = 'chat';
  let manage = false;
  let sending = false;
  let checking = false;

  // ---------- persistent parts: the header, the transcript and the composer ----------
  const progress = el('span', 'agent-progress');
  const head = el('header', 'agent-head');
  const body = el('div', 'agent-body');

  const textarea = el('textarea', 'agent-input');
  textarea.rows = 1;
  const chips = el('div', 'agent-chips');
  const sendBtn = iconButton('send', T.agent.send, () => void send(textarea.value));
  sendBtn.classList.add('agent-send');
  const stopBtn = iconButton('stop', T.agent.stop, () => {
    const id = getState().activeId;
    if (id) stopSession(id).catch((e) => toast(String(e.message ?? e)));
  });
  stopBtn.classList.add('agent-stop');
  const composerBox = el('div', 'agent-composer-box');
  const composerRow = el('div', 'agent-composer-row');
  composerRow.append(chips, sendBtn, stopBtn);
  composerBox.append(textarea, composerRow);
  const composer = el('div', 'agent-composer');
  composer.append(composerBox);

  const grow = () => {
    textarea.style.height = 'auto';
    textarea.style.height = `${Math.min(160, textarea.scrollHeight)}px`;
  };
  textarea.addEventListener('input', () => {
    grow();
    paintComposer();
  });
  textarea.addEventListener('keydown', (e) => {
    if (e.isComposing || e.keyCode === 229) return;
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      void send(textarea.value);
    }
  });

  const transcript = createTranscript({
    onRespond: (itemId, payload) => {
      const id = getState().activeId;
      if (!id) return Promise.resolve();
      return respond(id, itemId, payload).catch((e) => toast(String(e.message ?? e)));
    },
    onEdit: (itemId, text) => {
      const id = getState().activeId;
      if (!id) return Promise.resolve();
      return sendMessage(id, text, itemId).catch((e) => toast(`${T.agent.sendFailed}: ${e.message ?? e}`));
    },
    onFix: (text) => send(text),
  });

  panel.append(progress, head, body);

  // ---------- sending ----------
  async function ensureSession(): Promise<string> {
    const s = getState();
    const provider = currentProvider();
    if (!provider) throw new Error('no agent connected');
    const current = s.activeId ? s.sessions.find((x) => x.id === s.activeId) : null;
    if (current && current.status !== 'running' && current.provider === provider.id) return current.id;
    return (await createSession(provider.id, s.context.workbook ?? null)).id;
  }

  async function send(text: string): Promise<void> {
    const trimmed = text.trim();
    const s = getState();
    const session = s.activeId ? s.full[s.activeId] : null;
    // A second turn would edit the same workbook alongside the running one.
    if (!trimmed || sending || session?.status === 'running') return;
    sending = true;
    paintComposer();
    try {
      const id = await ensureSession();
      if (text === textarea.value) {
        textarea.value = '';
        grow();
      }
      await sendMessage(id, trimmed);
    } catch (e) {
      toast(`${T.agent.sendFailed}: ${(e as Error).message ?? e}`);
    } finally {
      sending = false;
      paintComposer();
    }
  }

  // ---------- painting ----------
  function paintComposer() {
    const s = getState();
    const session = s.activeId ? s.full[s.activeId] : null;
    const running = session?.status === 'running';
    const provider = currentProvider();
    textarea.placeholder = provider ? (s.context.workbook ? T.agent.placeholder(provider.name) : T.agent.placeholderHome(provider.name)) : '';
    sendBtn.hidden = running;
    stopBtn.hidden = !running;
    sendBtn.disabled = sending || !textarea.value.trim();
    chips.replaceChildren();
    const c = s.context;
    if (c.workbook) {
      const chip = el('span', 'agent-chip');
      const where = c.sheet ? `${c.sheet}${c.selection ? `!${c.selection}` : ''}` : '';
      chip.textContent = [c.title || c.workbook, where].filter(Boolean).join(' · ');
      chip.title = `sheets/${c.workbook}/workbook.json`;
      chips.append(chip);
    }
  }

  function paintHead(connected: boolean, showConnect: boolean) {
    head.replaceChildren();
    const left = el('div', 'agent-head-left');
    if (showConnect && connected) {
      left.append(textButton(T.agent.back, () => ((manage = false), paint())));
    } else if (view === 'history') {
      const b = textButton(T.agent.history, () => ((view = 'chat'), paint()));
      b.prepend(icon('arrowLeft', 14));
      left.append(b);
    } else {
      left.append(modelButton());
    }
    const beta = el('span', 'agent-beta', T.agent.beta);
    left.append(beta);
    const right = el('div', 'agent-head-right');
    if (connected && !showConnect) {
      const hist = iconButton('history', T.agent.history, () => {
        view = view === 'history' ? 'chat' : 'history';
        paint();
      });
      hist.classList.toggle('on', view === 'history');
      hist.setAttribute('aria-pressed', String(view === 'history'));
      const fresh = iconButton('plus', T.agent.newChat, () => {
        setActiveSession(null);
        view = 'chat';
        textarea.value = '';
        grow();
        paint();
        textarea.focus();
      });
      right.append(hist, fresh);
    }
    right.append(iconButton('x', T.agent.close, () => setOpen(false)));
    head.append(left, right);
  }

  function modelButton(): HTMLElement {
    const s = getState();
    const provider = currentProvider();
    const b = el('button', 'agent-model');
    b.type = 'button';
    const choice = provider ? resolveModelChoice(provider.models, s.modelChoices[provider.id]) : null;
    const label = choice ? (choice.model.id ? choice.model.label : T.agent.modelDefault) : T.agent.title;
    b.append(icon('sparkles', 14), el('span', 'agent-model-label', label));
    if (choice?.effort) b.append(el('span', 'agent-model-effort', T.agent.efforts[choice.effort]));
    b.append(icon('chevronDown', 12));
    b.addEventListener('click', () => {
      const p = currentProvider();
      const entries: MenuEntry[] = [];
      if (p) {
        const ch = resolveModelChoice(p.models, getState().modelChoices[p.id]);
        for (const m of p.models) {
          entries.push({ label: m.id ? m.label : T.agent.modelDefault, checked: m.id === ch?.model.id, action: () => setModelChoice(p.id, { model: m.id, effort: getState().modelChoices[p.id]?.effort ?? null }) });
        }
        if (ch && ch.model.efforts.length) {
          entries.push('-', {
            label: `${T.agent.reasoning} · ${ch.effort ? T.agent.efforts[ch.effort] : T.agent.effortAuto}`,
            submenu: [null, ...ch.model.efforts].map((e) => ({ label: e ? T.agent.efforts[e] : T.agent.effortAuto, checked: e === ch.effort, action: () => setModelChoice(p.id, { model: ch.model.id, effort: e }) })),
          });
        }
        if (readyProviders().some((x) => x.asksPermission)) entries.push({ label: T.agent.autoApprove, checked: getState().autoApprove, action: () => setAutoApprove(!getState().autoApprove) });
        entries.push('-');
      }
      entries.push({ label: T.agent.manage, icon: 'settings', action: () => ((manage = true), paint()) });
      showMenu(entries, b);
    });
    return b;
  }

  function textButton(label: string, onClick: () => void): HTMLButtonElement {
    const b = el('button', 'agent-text-btn', label);
    b.type = 'button';
    b.addEventListener('click', onClick);
    return b;
  }

  function connectView(): HTMLElement {
    const s = getState();
    const wrap = el('div', 'agent-connect');
    const providers = s.status?.providers ?? [];
    const intro = el('div', 'agent-intro');
    intro.append(el('span', 'agent-eyebrow', T.agent.title), el('h2', '', T.agent.connectTitle), el('p', '', T.agent.connectHint));
    wrap.append(intro);
    for (const p of providers) wrap.append(providerCard(p));
    if (providers.some((p) => !p.ready)) {
      const again = el('button', 'btn small');
      again.type = 'button';
      again.disabled = checking;
      again.append(icon('refresh', 13), el('span', '', T.agent.checkAgain));
      again.addEventListener('click', async () => {
        checking = true;
        paint();
        await refreshStatus();
        checking = false;
        paint();
      });
      wrap.append(again);
    }
    return wrap;
  }

  function providerCard(p: AgentProviderStatus): HTMLElement {
    const s = getState();
    const card = el('div', 'agent-card');
    card.dataset.provider = p.id;
    const top = el('div', 'agent-card-top');
    const name = el('div', 'agent-card-name');
    name.append(el('h3', '', p.name));
    const cli = el('p', `agent-card-cli${p.cliFound ? '' : ' warn'}`);
    cli.append(icon('terminal', 12), el('span', '', p.cliFound ? T.agent.cliFound(p.cliVersion) : T.agent.cliMissing));
    name.append(cli);
    top.append(name);
    const error = s.connectErrors[p.id];
    const missingCli = !p.cliFound && !p.ready;
    if (p.ready) {
      const ok = el('span', 'agent-connected');
      ok.append(icon('check', 12), el('span', '', T.agent.connected));
      top.append(ok);
    } else if (s.installing === p.id) {
      top.append(el('span', 'agent-connecting agent-shimmer', T.agent.connecting));
    } else if (!missingCli) {
      const b = el('button', `btn small${error ? '' : ' primary'}`, error ? T.agent.retry : T.agent.connect);
      b.type = 'button';
      b.disabled = s.installing !== null;
      b.addEventListener('click', async () => {
        try {
          await installProvider(p.id);
          toast(T.agent.connectedToast(p.name));
          manage = false;
          view = 'chat';
          setActiveSession(null);
        } catch {
          toast(T.agent.connectFailed(p.sdkPackage));
        }
      });
      top.append(b);
    }
    card.append(top);
    if (missingCli) {
      card.append(el('p', 'agent-card-hint', T.agent.cliHint(p.name)));
      const code = el('code', 'agent-code', CLI_INSTALL[p.id] ?? '');
      card.append(code);
    } else if (error && s.installing !== p.id) {
      card.append(el('p', 'agent-card-error', error));
    } else {
      card.append(el('p', 'agent-card-hint', p.ready ? `${p.sdkPackage} v${p.sdkVersion}` : `${T.agent.installHint(p.sdkPackage, s.status?.packageManager || 'npm')} ${T.agent.cliHint(p.name)}`));
    }
    return card;
  }

  function historyView(): HTMLElement {
    const s = getState();
    const wrap = el('div', 'agent-history');
    if (!s.sessions.length) {
      const empty = el('div', 'agent-history-empty');
      empty.append(icon('history', 20), el('p', 'strong', T.agent.noSessions), el('p', '', T.agent.historyEmptyHint));
      wrap.append(empty);
      return wrap;
    }
    const here = s.context.workbook ?? null;
    const groups: { label: string | null; items: AgentSessionSummary[] }[] = here
      ? [
          { label: T.agent.historyThisBook, items: s.sessions.filter((x) => x.workbook === here) },
          { label: T.agent.historyOtherBooks, items: s.sessions.filter((x) => x.workbook !== here) },
        ].filter((g) => g.items.length)
      : [{ label: null, items: s.sessions }];
    for (const g of groups) {
      const sec = el('section', 'agent-history-group');
      if (g.label) sec.append(el('h3', 'agent-eyebrow', g.label));
      const ul = el('ul');
      for (const x of g.items) {
        const li = el('li', 'agent-history-row');
        const b = el('button', `agent-history-item${x.id === s.activeId ? ' active' : ''}`);
        b.type = 'button';
        const meta = el('span', 'agent-history-meta');
        meta.append(x.status === 'running' ? el('span', 'agent-shimmer', T.agent.sessionRunning) : el('span', '', relativeTime(x.updatedAt)));
        if (x.workbook && x.workbook !== here) meta.append(el('span', '', '·'), el('span', 'agent-history-book', x.workbook));
        b.append(el('span', 'agent-history-title', x.title), meta);
        b.addEventListener('click', () => {
          setActiveSession(x.id);
          view = 'chat';
          if (x.workbook && x.workbook !== here) location.hash = `#/w/${encodeURIComponent(x.workbook)}`;
          paint();
        });
        const del = iconButton('trash', T.agent.deleteSession, () => deleteSession(x.id).catch((e) => toast(String(e.message ?? e))));
        del.classList.add('agent-history-delete');
        li.append(b, del);
        ul.append(li);
      }
      sec.append(ul);
      wrap.append(sec);
    }
    return wrap;
  }

  function emptyState(): HTMLElement {
    const s = getState();
    const wrap = el('div', 'agent-empty');
    const inner = el('div', 'agent-intro');
    inner.append(el('span', 'agent-eyebrow', T.agent.title), el('h2', '', T.agent.emptyTitle), el('p', '', T.agent.emptyHint));
    const actions = el('div', 'agent-quick');
    const quick = (iconName: string, label: string, prompt: string) => {
      const b = el('button', 'agent-quick-btn');
      b.type = 'button';
      b.append(icon(iconName, 14), el('span', '', label));
      b.addEventListener('click', () => void send(prompt));
      actions.append(b);
    };
    if (s.context.workbook) {
      quick('sheet', T.agent.quickExplain, T.agent.quickExplainPrompt);
      quick('search', T.agent.quickCheck, T.agent.quickCheckPrompt);
      quick('pencil', T.agent.quickNotes, T.agent.quickNotesPrompt);
    } else {
      quick('plus', T.agent.quickCreate, T.agent.quickCreatePrompt);
    }
    inner.append(actions);
    wrap.append(inner);
    return wrap;
  }

  let lastBody: string | null = null;
  function paint() {
    const s = getState();
    panel.hidden = !s.open;
    if (!s.open) return;
    const session = s.activeId ? (s.full[s.activeId] ?? null) : null;
    if (s.activeId && !session) void loadSession(s.activeId);
    const connected = currentProvider() !== null;
    const loading = s.status === null && !s.statusError;
    const showConnect = !loading && s.status?.enabled !== false && (!connected || manage);
    progress.hidden = session?.status !== 'running';
    paintHead(connected, showConnect);

    // Which body: the transcript and composer stay the same nodes between paints.
    const mode = loading ? 'loading' : s.status?.enabled === false || s.statusError ? 'off' : showConnect ? 'connect' : view === 'history' ? 'history' : 'chat';
    if (mode === 'chat') {
      if (lastBody !== 'chat') body.replaceChildren(transcript.el, composer);
      const hasItems = Boolean(session && session.items.length);
      if (hasItems) transcript.render(session!);
      else {
        transcript.clear();
        transcript.el.replaceChildren(emptyState());
      }
      paintComposer();
    } else if (mode === 'loading') {
      body.replaceChildren(el('p', 'agent-note', T.agent.loading));
    } else if (mode === 'off') {
      body.replaceChildren(el('p', 'agent-note', s.status?.reason === 'config' ? T.agent.offConfig : T.agent.offRemote));
    } else {
      body.replaceChildren(mode === 'connect' ? connectView() : historyView());
    }
    lastBody = mode;
    if (focusComposer && textarea.isConnected) {
      focusComposer = false;
      textarea.focus();
    }
  }

  let frame = 0;
  const schedule = () => {
    if (frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      paint();
    });
  };
  let wasOpen = getState().open;
  /** Set when the panel opens: the composer takes the focus once it is drawn. */
  let focusComposer = false;
  subscribe(() => {
    const open = getState().open;
    panel.hidden = !open;
    schedule();
    if (open && !wasOpen) focusComposer = true;
    if (!open && wasOpen) document.querySelector<HTMLInputElement>('.book .editor')?.focus();
    wasOpen = open;
  });

  // ⌘J anywhere (also from the composer) shows or hides the panel.
  window.addEventListener(
    'keydown',
    (e) => {
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'j') {
        e.preventDefault();
        e.stopPropagation();
        toggleOpen();
      }
    },
    { capture: true },
  );

  // The interface language changed (home's language menu): draw everything again in it.
  new MutationObserver(() => {
    for (const [b, label] of [
      [sendBtn, T.agent.send],
      [stopBtn, T.agent.stop],
    ] as const) {
      b.title = label;
      b.setAttribute('aria-label', label);
    }
    panel.setAttribute('aria-label', T.agent.title);
    transcript.clear();
    lastBody = null;
    paint();
  }).observe(document.documentElement, { attributes: true, attributeFilter: ['lang'] });

  if (getState().open) {
    void refreshStatus();
    void loadSessions();
  }
  paint();
}
