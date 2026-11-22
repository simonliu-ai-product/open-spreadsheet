/**
 * Runs the panel's turns: one at a time per conversation, turning a
 * provider's events into transcript items, holding the questions and
 * permission prompts the agent is waiting on, and stopping on request.
 *
 * @typedef {import('./types.d.ts').AgentTurnContext} AgentTurnContext
 * @typedef {import('./types.d.ts').AgentEffort} AgentEffort
 * @typedef {import('./types.d.ts').AgentQuestion} AgentQuestion
 * @typedef {import('./types.d.ts').RespondPayload} RespondPayload
 * @typedef {import('./provider.js').AgentProvider} AgentProvider
 * @typedef {import('./provider.js').ProviderEvent} ProviderEvent
 *
 * @typedef {object} SendOptions
 * @property {string} text
 * @property {AgentTurnContext} [context]
 * @property {boolean} [autoApprove]
 * @property {string | null} [model]
 * @property {AgentEffort | null} [effort]
 * @property {string} [answering]  the question item this turn answers; it stands in for the user message
 * @property {string} [edit]       a user message to replace: the transcript is cut back to it
 *
 * @typedef {Partial<Pick<SendOptions, 'context' | 'autoApprove' | 'model' | 'effort'>>} TurnSettings
 * @typedef {{ sessionId: string, userItemId: string }} Turn
 */
import { answersPrompt, buildTurnPrompt, editedPrompt, systemPromptAppend } from './prompt.js';
import { newId, titleFromPrompt } from './store.js';

/**
 * @param {object} opts
 * @param {string} opts.cwd
 * @param {string[]} opts.writableRoots
 * @param {ReturnType<typeof import('./store.js').createSessionStore>} opts.store
 * @param {Record<string, AgentProvider>} opts.providers
 * @param {() => import('./prompt.js').PromptOptions} opts.promptOptions
 * @param {(turn: Turn) => void} [opts.onTurnStart]
 * @param {(turn: Turn) => void | Promise<void>} [opts.onTurnEnd]
 */
export function createAgentRunner(opts) {
  const { cwd, writableRoots, store, providers, promptOptions, onTurnStart, onTurnEnd } = opts;
  /** @type {Map<string, AbortController>} */
  const running = new Map();
  /** Questions and permission prompts a running turn waits on, by item id. */
  const pending = new Map();
  /** @type {Map<string, SendOptions>} */
  const lastOptions = new Map();
  const now = () => new Date().toISOString();

  const isRunning = (sessionId) => running.has(sessionId);

  const cancelPrompt = (prompt) => prompt.resolve(prompt.kind === 'question' ? null : false);

  /** Starts a turn; resolves once it is under way, not when it ends. @param {SendOptions} options */
  async function send(sessionId, options) {
    const session = store.get(sessionId);
    if (!session) throw new Error(`unknown agent session: ${sessionId}`);
    if (running.has(sessionId)) throw new Error('session is already running');
    const text = options.text.trim();
    if (!text) throw new Error('empty prompt');
    const provider = providers[session.provider];
    if (!provider) throw new Error(`unknown agent: ${session.provider}`);

    let promptText = text;
    if (options.edit) {
      const index = session.items.findIndex((i) => i.id === options.edit);
      const target = session.items[index];
      if (target?.kind !== 'user') throw new Error('unknown message');
      // The first message (with nothing trimmed before it) starts over; a later
      // one rides on top of the old conversation, which the agent sets aside.
      const first = !session.trimmed && !session.items.slice(0, index).some((i) => i.kind === 'user');
      if (first) store.patch(sessionId, { providerSessionId: null });
      else promptText = editedPrompt(target.text, text);
      store.truncateFrom(sessionId, target.id);
    }

    const controller = new AbortController();
    running.set(sessionId, controller);
    lastOptions.set(sessionId, { ...options, edit: undefined });
    if (session.items.length === 0) store.patch(sessionId, { title: titleFromPrompt(text) });
    if (options.context?.workbook && !session.workbook) store.patch(sessionId, { workbook: options.context.workbook });
    for (const item of session.items) {
      if (item.kind === 'question' && item.status === 'pending' && item.id !== options.answering) store.updateItem(sessionId, item.id, 'question', { status: 'cancelled' });
    }
    const userItemId = options.answering ?? newId('u');
    if (!options.answering) store.appendItem(sessionId, { kind: 'user', id: userItemId, ts: now(), text });
    store.patch(sessionId, { status: 'running' });
    onTurnStart?.({ sessionId, userItemId });

    const openText = new Set();
    const openTools = new Set();
    const startedAt = Date.now();
    let resultSeen = false;

    const appendQuestion = (questions) => store.appendItem(sessionId, { kind: 'question', id: newId('q'), ts: now(), questions, answers: null, status: 'pending' });

    /** @param {ProviderEvent} event */
    const emit = (event) => {
      if (!store.get(sessionId)) return;
      switch (event.type) {
        case 'session-id':
          store.patch(sessionId, { providerSessionId: event.id });
          return;
        case 'text-start':
          openText.add(event.id);
          store.appendItem(sessionId, { kind: 'assistant', id: event.id, ts: now(), text: '', streaming: true });
          return;
        case 'text-delta':
          if (!openText.has(event.id)) emit({ type: 'text-start', id: event.id });
          store.appendDelta(sessionId, event.id, event.text);
          return;
        case 'text-end':
          if (!openText.has(event.id)) {
            if (!event.text) return;
            emit({ type: 'text-start', id: event.id });
          }
          openText.delete(event.id);
          store.updateItem(sessionId, event.id, 'assistant', event.text !== null ? { streaming: false, text: event.text } : { streaming: false });
          return;
        case 'tool-start':
          openTools.add(event.id);
          store.appendItem(sessionId, { kind: 'tool', id: event.id, ts: now(), name: event.name, title: event.title, detail: event.detail, status: 'running' });
          return;
        case 'tool-end':
          openTools.delete(event.id);
          store.updateItem(sessionId, event.id, 'tool', { status: event.ok ? 'done' : 'error' });
          return;
        case 'error':
          store.appendItem(sessionId, { kind: 'error', id: newId('e'), ts: now(), text: event.text });
          return;
        case 'questions':
          appendQuestion(event.questions);
          return;
        case 'result':
          resultSeen = true;
          store.appendItem(sessionId, { kind: 'result', id: newId('r'), ts: now(), durationMs: event.durationMs ?? Date.now() - startedAt, costUsd: event.costUsd, aborted: false });
          return;
      }
    };

    const askUser = (questions) =>
      new Promise((resolve) => {
        if (controller.signal.aborted) return resolve(null);
        const item = appendQuestion(questions);
        pending.set(item.id, {
          kind: 'question',
          resolve: (answers) => {
            pending.delete(item.id);
            store.updateItem(sessionId, item.id, 'question', { answers, status: answers ? 'answered' : 'cancelled' });
            resolve(answers);
          },
        });
      });

    const requestPermission = (request) =>
      new Promise((resolve) => {
        if (controller.signal.aborted) return resolve(false);
        const item = store.appendItem(sessionId, { kind: 'permission', id: newId('p'), ts: now(), tool: request.tool, title: request.title, detail: request.detail, status: options.autoApprove ? 'allowed' : 'pending' });
        if (options.autoApprove) return resolve(true);
        pending.set(item.id, {
          kind: 'permission',
          resolve: (allow) => {
            pending.delete(item.id);
            store.updateItem(sessionId, item.id, 'permission', { status: allow ? 'allowed' : 'denied' });
            resolve(allow);
          },
        });
      });

    const finish = () => {
      running.delete(sessionId);
      if (!store.get(sessionId)) return;
      for (const id of openText) store.updateItem(sessionId, id, 'assistant', { streaming: false });
      for (const id of openTools) store.updateItem(sessionId, id, 'tool', { status: 'error' });
      if (!resultSeen) store.appendItem(sessionId, { kind: 'result', id: newId('r'), ts: now(), durationMs: Date.now() - startedAt, costUsd: null, aborted: controller.signal.aborted });
      store.patch(sessionId, { status: 'idle' });
      if (onTurnEnd) {
        void Promise.resolve()
          .then(() => onTurnEnd({ sessionId, userItemId }))
          .catch((err) => console.error('[open-spreadsheet] agent turn-end handler failed:', err));
      }
    };

    void provider
      .run({
        cwd,
        writableRoots,
        prompt: buildTurnPrompt(promptText, options.context),
        systemPromptAppend: systemPromptAppend(provider.askInstruction, promptOptions()),
        resume: store.get(sessionId)?.providerSessionId ?? null,
        model: options.model ?? null,
        effort: options.effort ?? null,
        signal: controller.signal,
        emit,
        askUser,
        requestPermission,
      })
      .catch((err) => {
        if (!controller.signal.aborted) emit({ type: 'error', text: err instanceof Error ? err.message : String(err) });
      })
      .finally(finish);
  }

  function stop(sessionId) {
    const controller = running.get(sessionId);
    if (!controller) return false;
    for (const item of store.get(sessionId)?.items ?? []) {
      const prompt = pending.get(item.id);
      if (prompt) cancelPrompt(prompt);
    }
    controller.abort();
    return true;
  }

  /** Answers a question or permission prompt; a question that ended the turn starts the next one. @param {RespondPayload} payload @param {TurnSettings} next */
  function respond(sessionId, itemId, payload, next = {}) {
    const session = store.get(sessionId);
    if (!session?.items.some((i) => i.id === itemId)) return false;
    const prompt = pending.get(itemId);
    if (!prompt) return answerInNextTurn(sessionId, itemId, payload, next);
    if ('cancel' in payload) {
      cancelPrompt(prompt);
      return true;
    }
    if (prompt.kind === 'question' && 'answers' in payload) {
      prompt.resolve(payload.answers);
      return true;
    }
    if (prompt.kind === 'permission' && 'allow' in payload) {
      prompt.resolve(payload.allow);
      return true;
    }
    return false;
  }

  function answerInNextTurn(sessionId, itemId, payload, next) {
    const item = store.get(sessionId)?.items.find((i) => i.id === itemId);
    if (item?.kind !== 'question' || item.status !== 'pending' || running.has(sessionId)) return false;
    if (!('answers' in payload) && !('cancel' in payload)) return false;
    const answers = 'answers' in payload ? payload.answers : null;
    store.updateItem(sessionId, itemId, 'question', { answers, status: answers ? 'answered' : 'cancelled' });
    // The panel sends its current settings; the last turn's only fill the gaps.
    const previous = { ...lastOptions.get(sessionId), ...next };
    void send(sessionId, { ...previous, text: answersPrompt(item.questions, answers), answering: itemId }).catch((err) => console.error('[open-spreadsheet] agent turn failed:', err));
    return true;
  }

  function stopAll() {
    for (const id of Array.from(running.keys())) stop(id);
  }

  return { send, stop, respond, isRunning, stopAll };
}
