/**
 * The agent panel's vocabulary, shared by the dev server and the editor
 * (types.d.ts describes the shapes).
 */

// Anthropic does not allow a third-party product to run on a person's claude.ai
// login, so Claude Code is not offered here; Antigravity waits for official
// support for third-party harnesses. A new agent is one more id here and one
// more file in providers/.
export const AGENT_PROVIDER_IDS = ['codex'];

export const AGENT_EFFORTS = ['minimal', 'low', 'medium', 'high', 'xhigh', 'max'];

/** @returns {value is import('./types.d.ts').AgentProviderId} */
export const isAgentProviderId = (value) => typeof value === 'string' && AGENT_PROVIDER_IDS.includes(value);

// The levels the panel has words for. A model may offer others (a new level such
// as 'ultra'): those pass through as given and show under their own name.
/** @returns {value is import('./types.d.ts').AgentEffort} */
export const isAgentEffort = (value) => typeof value === 'string' && /^[a-z][a-z0-9_-]{0,31}$/.test(value);

/** A session without its transcript and provider ids, as lists and events carry it. */
export function sessionSummary(session) {
  const { providerSessionId: _p, trimmed: _t, items: _i, ...summary } = session;
  return summary;
}
