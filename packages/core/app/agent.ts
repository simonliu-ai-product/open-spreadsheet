/**
 * The agent panel's client: its state, the /__agent API, and the event
 * stream that keeps transcripts live. No DOM here (agent-panel.ts draws it).
 */
import type {
  AgentCheckFinding,
  AgentEffort,
  AgentItem,
  AgentModelOption,
  AgentProviderId,
  AgentProviderStatus,
  AgentQuestion,
  AgentSession,
  AgentSessionSummary,
  AgentStatus,
  AgentStreamEvent,
  AgentTurnContext,
  RespondPayload,
} from '../agent/types.js';
import { hostHas } from './host.ts';

export type { AgentCheckFinding, AgentItem, AgentProviderStatus, AgentQuestion, AgentSession, AgentSessionSummary, AgentTurnContext, AgentEffort, AgentModelOption, RespondPayload };

type ModelChoice = { model: string; effort: AgentEffort | null };

export interface AgentState {
  status: AgentStatus | null;
  statusError: boolean;
  sessions: AgentSessionSummary[];
  sessionsLoaded: boolean;
  full: Record<string, AgentSession>;
  activeId: string | null;
  provider: AgentProviderId;
  modelChoices: Partial<Record<AgentProviderId, ModelChoice>>;
  autoApprove: boolean;
  open: boolean;
  installing: AgentProviderId | null;
  connectErrors: Partial<Record<AgentProviderId, string>>;
  /** What the person has open: the workbook, sheet and selection (as in current.json). */
  context: AgentTurnContext;
}

const KEY = {
  active: 'open-spreadsheet:agent-session',
  provider: 'open-spreadsheet:agent-provider',
  model: 'open-spreadsheet:agent-model',
  autoApprove: 'open-spreadsheet:agent-auto-approve',
  open: 'open-spreadsheet:agent-open',
};

const read = (key: string): string | null => {
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
};
const write = (key: string, value: string | null) => {
  try {
    if (value === null) localStorage.removeItem(key);
    else localStorage.setItem(key, value);
  } catch {
    // not remembered
  }
};

/**
 * The panel runs on the dev server only, and steps aside for a host that has
 * its own agent (so one window never shows two).
 */
export const agentAvailable = (): boolean => Boolean(import.meta.hot) && !hostHas('agent');

function readModelChoices(): AgentState['modelChoices'] {
  try {
    const raw = JSON.parse(read(KEY.model) ?? '{}') as Record<string, { model?: unknown; effort?: unknown }>;
    const out: AgentState['modelChoices'] = {};
    for (const [p, v] of Object.entries(raw)) {
      if (p === 'codex' && typeof v?.model === 'string') out[p] = { model: v.model, effort: typeof v.effort === 'string' ? (v.effort as AgentEffort) : null };
    }
    return out;
  } catch {
    return {};
  }
}

let state: AgentState = {
  status: null,
  statusError: false,
  sessions: [],
  sessionsLoaded: false,
  full: {},
  activeId: read(KEY.active),
  provider: 'codex',
  modelChoices: readModelChoices(),
  autoApprove: read(KEY.autoApprove) === '1',
  open: read(KEY.open) === '1',
  installing: null,
  connectErrors: {},
  context: {},
};

const listeners = new Set<() => void>();
export const getState = (): AgentState => state;
export function subscribe(fn: () => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}
function setState(patch: Partial<AgentState> | ((s: AgentState) => Partial<AgentState>)) {
  state = { ...state, ...(typeof patch === 'function' ? patch(state) : patch) };
  for (const l of listeners) l();
}

// ---------- the event stream ----------

let stream: EventSource | null = null;
let reconnected = false;

function openStream() {
  if (stream) return;
  stream = new EventSource('/__agent/events');
  stream.onmessage = (m) => applyEvent(JSON.parse(m.data) as AgentStreamEvent);
  stream.onopen = () => {
    // Anything streamed while disconnected is gone: reload what is on screen.
    if (!reconnected) {
      reconnected = true;
      return;
    }
    void refreshStatus();
    void loadSessions();
    for (const id of Object.keys(state.full)) void loadSession(id);
  };
}

const bySince = (a: AgentSessionSummary, b: AgentSessionSummary) => (a.updatedAt < b.updatedAt ? 1 : a.updatedAt > b.updatedAt ? -1 : 0);

function withItems(id: string, fn: (items: AgentItem[]) => AgentItem[]): Partial<AgentState> {
  const current = state.full[id];
  return current ? { full: { ...state.full, [id]: { ...current, items: fn(current.items) } } } : {};
}

export function applyEvent(event: AgentStreamEvent): void {
  switch (event.type) {
    case 'session':
      return setState((s) => {
        const current = s.full[event.session.id];
        return {
          sessions: [event.session, ...s.sessions.filter((x) => x.id !== event.session.id)].sort(bySince),
          full: current ? { ...s.full, [event.session.id]: { ...current, ...event.session } } : s.full,
        };
      });
    case 'session-removed':
      return setState((s) => {
        const { [event.sessionId]: _gone, ...full } = s.full;
        return { sessions: s.sessions.filter((x) => x.id !== event.sessionId), full, activeId: s.activeId === event.sessionId ? null : s.activeId };
      });
    case 'item':
      return setState(() =>
        withItems(event.sessionId, (items) => {
          const i = items.findIndex((x) => x.id === event.item.id);
          return i === -1 ? [...items, event.item] : items.map((x, n) => (n === i ? event.item : x));
        }),
      );
    case 'truncate':
      return setState(() =>
        withItems(event.sessionId, (items) => {
          const i = items.findIndex((x) => x.id === event.itemId);
          return i === -1 ? items : items.slice(0, i);
        }),
      );
    case 'delta':
      return setState(() => withItems(event.sessionId, (items) => items.map((x) => (x.id === event.itemId && x.kind === 'assistant' ? { ...x, text: x.text + event.text } : x))));
  }
}

// ---------- the API ----------

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(path, init);
  if (!res.ok) {
    const body = (await res.json().catch(() => ({}))) as { error?: string };
    throw new Error(body.error ?? `${init?.method ?? 'GET'} ${path} → ${res.status}`);
  }
  return (await res.json()) as T;
}
const post = (body?: unknown): RequestInit => (body === undefined ? { method: 'POST' } : { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });

export async function refreshStatus(): Promise<void> {
  try {
    const status = await request<AgentStatus>('/__agent/status');
    setState({ status, statusError: false });
    if (status.enabled) openStream();
  } catch {
    setState({ statusError: true });
  }
}

export async function loadSessions(): Promise<void> {
  try {
    const { sessions } = await request<{ sessions: AgentSessionSummary[] }>('/__agent/sessions');
    setState((s) => ({ sessions, sessionsLoaded: true, activeId: s.activeId && !sessions.some((x) => x.id === s.activeId) ? null : s.activeId }));
  } catch {
    setState({ sessionsLoaded: true });
  }
}

export async function loadSession(id: string): Promise<void> {
  try {
    const session = await request<AgentSession>(`/__agent/sessions/${encodeURIComponent(id)}`);
    setState((s) => ({ full: { ...s.full, [id]: session } }));
  } catch {
    setState((s) => {
      const { [id]: _gone, ...full } = s.full;
      return { full, activeId: s.activeId === id ? null : s.activeId };
    });
  }
}

export function setActiveSession(id: string | null): void {
  write(KEY.active, id);
  setState({ activeId: id });
  if (id && !state.full[id]) void loadSession(id);
}

export async function createSession(provider: AgentProviderId, workbook: string | null): Promise<AgentSession> {
  const session = await request<AgentSession>('/__agent/sessions', post({ provider, workbook }));
  setState((s) => ({ full: { ...s.full, [session.id]: session }, sessions: [session, ...s.sessions.filter((x) => x.id !== session.id)].sort(bySince) }));
  setActiveSession(session.id);
  return session;
}

export async function deleteSession(id: string): Promise<void> {
  await request(`/__agent/sessions/${encodeURIComponent(id)}`, { method: 'DELETE' });
  applyEvent({ type: 'session-removed', sessionId: id });
}

export function readyProviders(): AgentProviderStatus[] {
  return state.status?.enabled ? state.status.providers.filter((p) => p.ready) : [];
}

/** The agent turns run on: the one picked last, else the first ready one. */
export function currentProvider(): AgentProviderStatus | null {
  const ready = readyProviders();
  return ready.find((p) => p.id === state.provider) ?? ready[0] ?? null;
}

export function resolveModelChoice(models: AgentModelOption[], choice: ModelChoice | undefined): { model: AgentModelOption; effort: AgentEffort | null } | null {
  const model = models.find((m) => m.id === choice?.model) ?? models[0];
  if (!model) return null;
  return { model, effort: choice?.effort && model.efforts.includes(choice.effort) ? choice.effort : null };
}

function turnSettings(id: string) {
  const provider = state.status?.providers.find((p) => p.id === (state.full[id]?.provider ?? state.provider));
  const choice = provider ? resolveModelChoice(provider.models, state.modelChoices[provider.id]) : null;
  return { autoApprove: state.autoApprove, model: choice?.model.id || null, effort: choice?.effort ?? null };
}

export async function sendMessage(id: string, text: string, edit?: string): Promise<void> {
  await request(`/__agent/sessions/${encodeURIComponent(id)}/messages`, post({ text, context: state.context, ...turnSettings(id), edit }));
}

export async function stopSession(id: string): Promise<void> {
  await request(`/__agent/sessions/${encodeURIComponent(id)}/stop`, post());
}

/** Answers that end a turn start the next one, so they carry the same settings and context a message would. */
export async function respond(id: string, itemId: string, payload: RespondPayload): Promise<void> {
  await request(`/__agent/sessions/${encodeURIComponent(id)}/respond`, post({ itemId, ...payload, context: state.context, ...turnSettings(id) }));
}

export async function installProvider(provider: AgentProviderId): Promise<void> {
  setState((s) => ({ installing: provider, connectErrors: { ...s.connectErrors, [provider]: undefined } }));
  try {
    await request('/__agent/install', post({ provider }));
    await refreshStatus();
  } catch (err) {
    setState((s) => ({ connectErrors: { ...s.connectErrors, [provider]: err instanceof Error ? err.message : String(err) } }));
    throw err;
  } finally {
    setState({ installing: null });
  }
}

export function setModelChoice(provider: AgentProviderId, choice: ModelChoice): void {
  const modelChoices = { ...state.modelChoices, [provider]: choice };
  write(KEY.model, JSON.stringify(modelChoices));
  write(KEY.provider, provider);
  setState({ modelChoices, provider });
}

export function setAutoApprove(on: boolean): void {
  write(KEY.autoApprove, on ? '1' : null);
  setState({ autoApprove: on });
}

export function setOpen(open: boolean): void {
  write(KEY.open, open ? '1' : null);
  setState({ open });
  if (open && !state.status) void refreshStatus();
  if (open && !state.sessionsLoaded) void loadSessions();
}
export const toggleOpen = (): void => setOpen(!state.open);

/**
 * The editor reports what is selected (the same payload it writes to
 * current.json). Arriving at another workbook starts a new conversation; the
 * previous one stays in History.
 */
export function setAgentContext(context: AgentTurnContext): void {
  const moved = (context.workbook ?? null) !== (state.context.workbook ?? null);
  const active = state.activeId ? state.sessions.find((s) => s.id === state.activeId) : null;
  setState({ context, ...(moved && active?.workbook && active.workbook !== context.workbook ? { activeId: null } : {}) });
}

/** What the transcript's fix button sends after the dev server's check. */
export function checkFixPrompt(findings: AgentCheckFinding[], total: number): string {
  const lines = findings.map((f) => `- ${f.workbook}${f.sheet || f.address ? ` · ${[f.sheet, f.address].filter(Boolean).join('!')}` : ''} [${f.severity}] ${f.message}`);
  const more = total > findings.length ? `\n- …and ${total - findings.length} more` : '';
  return `\`open-spreadsheet check\` found these problems. Fix them, keeping everything else as it is:\n${lines.join('\n')}${more}`;
}
