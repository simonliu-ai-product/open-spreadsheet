// Shapes the agent panel's server (plain JS beside server.js) and the editor
// (app/agent*.ts) share. The values live in types.js.

export type AgentProviderId = 'codex';
/** The levels the panel has words for; a model may offer others ('ultra'), which pass through as given. */
export type KnownAgentEffort = 'minimal' | 'low' | 'medium' | 'high' | 'xhigh' | 'max';
export type AgentEffort = string;

export declare const AGENT_PROVIDER_IDS: readonly AgentProviderId[];
export declare const AGENT_EFFORTS: readonly KnownAgentEffort[];
export declare function isAgentProviderId(value: unknown): value is AgentProviderId;
export declare function isAgentEffort(value: unknown): value is AgentEffort;
export declare function sessionSummary(session: AgentSession): AgentSessionSummary;

export type AgentModelOption = {
  /** Empty when the option stands for the CLI's own default model. */
  id: string;
  label: string;
  efforts: AgentEffort[];
};

export type AgentProviderStatus = {
  id: AgentProviderId;
  name: string;
  sdkPackage: string;
  sdkVersion: string | null;
  cliCommand: string;
  cliFound: boolean;
  cliVersion: string | null;
  /** The SDK resolves from the workspace, so a turn can run. */
  ready: boolean;
  models: AgentModelOption[];
  /** Whether the agent ever stops to ask before running a command. */
  asksPermission: boolean;
};

export type AgentStatus = {
  enabled: boolean;
  /** Why the panel is off: another machine is asking, or the workspace config turned it off. */
  reason?: 'remote' | 'config';
  packageManager: string;
  providers: AgentProviderStatus[];
};

export type AgentQuestion = {
  header: string;
  question: string;
  options: { label: string; description?: string }[];
  multiSelect: boolean;
};

export type AgentCheckFinding = {
  workbook: string;
  sheet: string;
  address: string;
  severity: 'error' | 'warning';
  message: string;
};

export type AgentItem =
  | { kind: 'user'; id: string; ts: string; text: string }
  | { kind: 'assistant'; id: string; ts: string; text: string; streaming: boolean }
  | { kind: 'tool'; id: string; ts: string; name: string; title: string; detail: string | null; status: 'running' | 'done' | 'error' }
  | { kind: 'question'; id: string; ts: string; questions: AgentQuestion[]; answers: Record<string, string> | null; status: 'pending' | 'answered' | 'cancelled' }
  | { kind: 'permission'; id: string; ts: string; tool: string; title: string; detail: string | null; status: 'pending' | 'allowed' | 'denied' }
  | { kind: 'error'; id: string; ts: string; text: string }
  // `open-spreadsheet check`, run by the dev server on the workbooks a turn changed.
  | {
      kind: 'check';
      id: string;
      ts: string;
      workbooks: string[];
      status: 'running' | 'done' | 'failed';
      /** The first few findings, errors first; `total` counts them all. */
      findings: AgentCheckFinding[];
      total: number;
      error: string | null;
    }
  | { kind: 'result'; id: string; ts: string; durationMs: number | null; costUsd: number | null; aborted: boolean };

export type AgentSessionSummary = {
  id: string;
  provider: AgentProviderId;
  title: string;
  createdAt: string;
  updatedAt: string;
  status: 'idle' | 'running';
  /** The workbook the conversation started on. */
  workbook: string | null;
};

export type AgentSession = AgentSessionSummary & {
  providerSessionId: string | null;
  /** Set once the oldest items were dropped to stay under the item cap. */
  trimmed?: boolean;
  items: AgentItem[];
};

/** What the person has open, as the editor reports it to current.json. */
export type AgentTurnContext = {
  workbook?: string;
  title?: string;
  sheet?: string;
  active?: string;
  selection?: string;
  input?: string;
  note?: string;
};

export type RespondPayload = { answers: Record<string, string> } | { allow: boolean } | { cancel: true };

export type AgentStreamEvent =
  | { type: 'session'; session: AgentSessionSummary }
  | { type: 'session-removed'; sessionId: string }
  | { type: 'item'; sessionId: string; item: AgentItem }
  | { type: 'delta'; sessionId: string; itemId: string; text: string }
  | { type: 'truncate'; sessionId: string; itemId: string };
