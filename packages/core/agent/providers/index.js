/** The agents the panel offers, by id (see AGENT_PROVIDER_IDS in ../types.js for why only these). */
import { createCodexProvider } from './codex.js';

/** @param {{ mcpServer?: { command: string, args: string[] } | null }} [opts] */
export function createProviders(opts = {}) {
  return { codex: createCodexProvider(opts) };
}
