/**
 * OpenAI Codex through @openai/codex-sdk, resolved from the workspace (never a
 * dependency of core). The SDK drives the Codex CLI the person signed in to;
 * the panel adds the open-spreadsheet MCP server to the thread, so the agent
 * changes workbooks through the same hash-checked tools as any other client.
 *
 * @typedef {import('../provider.js').AgentProvider} AgentProvider
 * @typedef {import('../provider.js').ProviderEvent} ProviderEvent
 * @typedef {import('../types.d.ts').AgentModelOption} AgentModelOption
 */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { findSpawnableCli, importWorkspacePackage } from '../detect.js';
import { extractQuestionBlock, TEXT_QUESTIONS_INSTRUCTION } from '../prompt.js';
import { clipDetail, ProviderUnavailableError, relativeToCwd } from '../provider.js';
import { isAgentEffort } from '../types.js';

const CODEX_SDK_PACKAGE = '@openai/codex-sdk';

/** @type {AgentModelOption} */
const CODEX_DEFAULT_MODEL = { id: '', label: 'Default', efforts: ['low', 'medium', 'high'] };

/**
 * The Codex CLI caches the model list its own /model picker shows; reading it
 * keeps the panel in step with what the account can use, with no network call.
 * @returns {AgentModelOption[]}
 */
export function parseCodexModelsCache(raw) {
  const list = raw?.models;
  if (!Array.isArray(list)) return [];
  const models = [];
  for (const m of list) {
    if (!m || typeof m !== 'object' || typeof m.slug !== 'string' || !m.slug) continue;
    if (m.visibility !== undefined && m.visibility !== 'list') continue;
    const levels = Array.isArray(m.supported_reasoning_levels) ? m.supported_reasoning_levels : [];
    const efforts = levels.map((l) => (l && typeof l === 'object' ? l.effort : l)).filter(isAgentEffort);
    models.push({
      option: { id: m.slug, label: typeof m.display_name === 'string' && m.display_name ? m.display_name : m.slug, efforts },
      priority: typeof m.priority === 'number' ? m.priority : Number.MAX_SAFE_INTEGER,
    });
  }
  return models.sort((a, b) => a.priority - b.priority).map((m) => m.option);
}

async function readCodexModels() {
  const home = process.env.CODEX_HOME || path.join(os.homedir(), '.codex');
  try {
    const models = parseCodexModelsCache(JSON.parse(await fs.readFile(path.join(home, 'models_cache.json'), 'utf8')));
    return models.length > 0 ? models : [CODEX_DEFAULT_MODEL];
  } catch {
    return [CODEX_DEFAULT_MODEL];
  }
}

/** A tool row's detail for an open-spreadsheet MCP call: where it acts (workbook · sheet · range or cells). */
function mcpDetail(args) {
  if (!args || typeof args !== 'object') return null;
  const where = [args.workbook ?? args.id, args.sheet, args.range ?? args.address ?? (args.cells && typeof args.cells === 'object' ? Object.keys(args.cells).join(', ') : undefined)];
  const text = where.filter((v) => typeof v === 'string' && v).join(' · ');
  return clipDetail(text || args.formula || null);
}

/**
 * Codex thread events as provider events. The SDK reports finished items (no
 * token deltas), so each agent message lands as one block.
 * @param {(event: ProviderEvent) => void} emit
 */
export function createCodexMapper(emit, cwd = process.cwd()) {
  const started = new Set();
  let resultSeen = false;

  const describe = (item) => {
    switch (item.type) {
      case 'command_execution':
        return { name: 'command', title: 'Run command', detail: clipDetail(item.command) };
      case 'file_change':
        return { name: 'file_change', title: 'Edit', detail: clipDetail((item.changes ?? []).map((c) => relativeToCwd(cwd, c.path)).join(', ')) };
      case 'mcp_tool_call':
        return item.server === 'open-spreadsheet'
          ? { name: 'mcp', title: item.tool, detail: mcpDetail(item.arguments) }
          : { name: 'mcp', title: `${item.server}: ${item.tool}`, detail: null };
      case 'web_search':
        return { name: 'web_search', title: 'Web search', detail: clipDetail(item.query) };
      default:
        return { name: item.type, title: item.type, detail: null };
    }
  };

  const handleItem = (phase, item) => {
    if (!item || typeof item !== 'object') return;
    switch (item.type) {
      case 'agent_message': {
        if (phase !== 'completed' || !item.text?.trim()) return;
        const asked = extractQuestionBlock(item.text);
        const text = asked ? asked.text : item.text;
        if (text) {
          emit({ type: 'text-start', id: item.id });
          emit({ type: 'text-end', id: item.id, text });
        }
        if (asked) emit({ type: 'questions', questions: asked.questions });
        return;
      }
      case 'error':
        if (phase === 'completed') emit({ type: 'error', text: item.message });
        return;
      case 'command_execution':
      case 'file_change':
      case 'mcp_tool_call':
      case 'web_search': {
        if (!started.has(item.id)) {
          started.add(item.id);
          emit({ type: 'tool-start', id: item.id, ...describe(item) });
        }
        if (phase === 'completed') emit({ type: 'tool-end', id: item.id, ok: item.status !== 'failed' && !item.error });
        return;
      }
      default:
        return;
    }
  };

  return {
    get resultSeen() {
      return resultSeen;
    },
    handle(event) {
      if (!event || typeof event !== 'object') return;
      switch (event.type) {
        case 'thread.started':
          emit({ type: 'session-id', id: event.thread_id });
          return;
        case 'item.started':
        case 'item.updated':
        case 'item.completed':
          handleItem(event.type.slice(5), event.item);
          return;
        case 'turn.failed':
          resultSeen = true;
          emit({ type: 'error', text: event.error?.message ?? 'The turn failed.' });
          emit({ type: 'result', durationMs: null, costUsd: null });
          return;
        case 'error':
          emit({ type: 'error', text: event.message });
          return;
        case 'turn.completed':
          resultSeen = true;
          emit({ type: 'result', durationMs: null, costUsd: null });
          return;
        default:
          return;
      }
    },
  };
}

/**
 * @param {{ mcpServer?: { command: string, args: string[] } | null }} [opts]
 *   the open-spreadsheet MCP server to add to every thread
 * @returns {AgentProvider}
 */
export function createCodexProvider(opts = {}) {
  return {
    id: 'codex',
    name: 'Codex',
    sdkPackage: CODEX_SDK_PACKAGE,
    cliCommand: 'codex',
    // `codex exec`, which the SDK drives, rejects request_user_input, so
    // questions travel as a tagged block in the final reply instead.
    askInstruction: TEXT_QUESTIONS_INSTRUCTION,
    // Runs with approvals off inside Codex's workspace-write sandbox.
    asksPermission: false,
    models: readCodexModels,
    async run(run) {
      const sdk = await importWorkspacePackage(run.cwd, CODEX_SDK_PACKAGE);
      if (!sdk || typeof sdk.Codex !== 'function') throw new ProviderUnavailableError(this);

      const localCli = findSpawnableCli('codex');
      const codex = new sdk.Codex({
        ...(localCli ? { codexPathOverride: localCli } : {}),
        // The framework's own tools are approved up front: with approvals off, Codex
        // would otherwise refuse every MCP call and could not change a workbook at all.
        ...(opts.mcpServer
          ? {
              config: {
                mcp_servers: {
                  'open-spreadsheet': {
                    command: opts.mcpServer.command,
                    args: opts.mcpServer.args,
                    default_tools_approval_mode: 'approve',
                  },
                },
              },
            }
          : {}),
      });
      // The sandbox can write the working directory (the workspace); the prompt
      // keeps edits to sheets/ and assets/. Roots outside it are added explicitly.
      const outside = run.writableRoots.filter((r) => path.relative(run.cwd, r).startsWith('..'));
      const threadOptions = {
        workingDirectory: run.cwd,
        skipGitRepoCheck: true,
        sandboxMode: 'workspace-write',
        approvalPolicy: 'never',
        ...(outside.length ? { additionalDirectories: outside } : {}),
        ...(run.model ? { model: run.model } : {}),
        ...(run.effort ? { modelReasoningEffort: run.effort } : {}),
      };
      const thread = run.resume ? codex.resumeThread(run.resume, threadOptions) : codex.startThread(threadOptions);

      // Codex has no system-prompt hook, so the panel's preamble rides on the first turn of a thread only.
      const input = run.resume ? run.prompt : `${run.systemPromptAppend}\n\n${run.prompt}`;
      const mapper = createCodexMapper(run.emit, run.cwd);
      const { events } = await thread.runStreamed(input, { signal: run.signal });
      try {
        for await (const event of events) mapper.handle(event);
      } catch (err) {
        if (run.signal.aborted || mapper.resultSeen) return;
        throw err;
      }
      if (thread.id) run.emit({ type: 'session-id', id: thread.id });
    },
  };
}
