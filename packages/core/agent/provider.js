/**
 * What every agent provider implements, and helpers they share.
 *
 * @typedef {import('./types.d.ts').AgentQuestion} AgentQuestion
 * @typedef {import('./types.d.ts').AgentEffort} AgentEffort
 *
 * @typedef {(
 *   | { type: 'session-id', id: string }
 *   | { type: 'text-start', id: string }
 *   | { type: 'text-delta', id: string, text: string }
 *   | { type: 'text-end', id: string, text: string | null }
 *   | { type: 'tool-start', id: string, name: string, title: string, detail: string | null }
 *   | { type: 'tool-end', id: string, ok: boolean }
 *   | { type: 'error', text: string }
 *   | { type: 'result', durationMs: number | null, costUsd: number | null }
 *   | { type: 'questions', questions: AgentQuestion[] }
 * )} ProviderEvent
 * A `questions` event carries questions the agent wrote into its reply instead
 * of asking through a tool: the turn ends, and the answers arrive as the next prompt.
 *
 * @typedef {{ tool: string, title: string, detail: string | null }} PermissionRequest
 *
 * @typedef {object} ProviderRunOptions
 * @property {string} cwd                 the workspace
 * @property {string[]} writableRoots     folders the agent may write to: sheets/ and the shared assets/
 * @property {string} prompt
 * @property {string} systemPromptAppend  the panel's preamble
 * @property {string | null} resume       the provider's own session id, to continue it
 * @property {string | null} model
 * @property {AgentEffort | null} effort
 * @property {AbortSignal} signal
 * @property {(event: ProviderEvent) => void} emit
 * @property {(questions: AgentQuestion[]) => Promise<Record<string, string> | null>} askUser
 * @property {(request: PermissionRequest) => Promise<boolean>} requestPermission
 *
 * @typedef {object} AgentProvider
 * @property {import('./types.d.ts').AgentProviderId} id
 * @property {string} name
 * @property {string} sdkPackage          resolved from the workspace, never bundled
 * @property {string} cliCommand
 * @property {string} askInstruction      how this agent asks the person something
 * @property {boolean} asksPermission     whether it ever stops for requestPermission
 * @property {() => Promise<import('./types.d.ts').AgentModelOption[]>} models
 * @property {(opts: ProviderRunOptions) => Promise<void>} run
 * @property {() => void} [dispose]       stops anything that outlives a turn
 */
import path from 'node:path';

export class ProviderUnavailableError extends Error {
  /** @param {AgentProvider} provider */
  constructor(provider) {
    super(`${provider.name} is not connected: install ${provider.sdkPackage} in this workspace to use it from the agent panel.`);
    this.name = 'ProviderUnavailableError';
  }
}

const DETAIL_MAX = 400;

/** One line of a tool's input for its transcript row, cut to a readable length. */
export function clipDetail(value) {
  if (value === undefined || value === null) return null;
  let text;
  try {
    text = typeof value === 'string' ? value : JSON.stringify(value);
  } catch {
    text = String(value);
  }
  if (!text) return null;
  const oneLine = text.replace(/\s+/g, ' ').trim();
  return oneLine.length > DETAIL_MAX ? `${oneLine.slice(0, DETAIL_MAX - 1)}…` : oneLine;
}

export function isInsideRoots(cwd, file, roots) {
  if (typeof file !== 'string' || !file) return false;
  const resolved = path.resolve(cwd, file);
  return roots.some((root) => {
    const rel = path.relative(root, resolved);
    return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
  });
}

/** A path as the transcript shows it: relative to the workspace when it is inside. */
export function relativeToCwd(cwd, file) {
  if (typeof file !== 'string' || !file) return null;
  const rel = path.relative(cwd, path.resolve(cwd, file));
  if (rel === '') return '.';
  return rel.startsWith('..') || path.isAbsolute(rel) ? file : rel.split(path.sep).join('/');
}

/** Questions in the shape AskUserQuestion-style tools send, dropping anything malformed. */
export function parseQuestions(input) {
  const raw = Array.isArray(input?.questions) ? input.questions : [];
  /** @type {AgentQuestion[]} */
  const questions = [];
  for (const q of raw) {
    if (!q || typeof q !== 'object' || typeof q.question !== 'string') continue;
    const options = Array.isArray(q.options)
      ? q.options
          .filter((o) => o && typeof o === 'object' && typeof o.label === 'string')
          .map((o) => ({ label: o.label, ...(typeof o.description === 'string' ? { description: o.description } : {}) }))
      : [];
    questions.push({ header: typeof q.header === 'string' ? q.header : '', question: q.question, options, multiSelect: q.multiSelect === true });
  }
  return questions;
}
