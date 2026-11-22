/**
 * What the panel tells the agent: a preamble about this workspace and how to
 * change workbooks safely, and each turn's context (the workbook, sheet and
 * selection on screen).
 *
 * @typedef {import('./types.d.ts').AgentTurnContext} AgentTurnContext
 * @typedef {import('./types.d.ts').AgentQuestion} AgentQuestion
 *
 * @typedef {object} PromptOptions
 * @property {string} checkCommand  how to run `open-spreadsheet check` from the workspace
 * @property {boolean} [mcp]        whether the open-spreadsheet MCP tools are wired into the agent
 */
import { parseQuestions } from './provider.js';

/** @param {string} askInstruction @param {PromptOptions} opts */
export function systemPromptAppend(askInstruction, opts) {
  return [
    'You are running inside the open-spreadsheet dev server, driven from its in-app agent panel in the browser rather than a terminal.',
    'The workspace is an open-spreadsheet workspace: each workbook is sheets/<id>/workbook.json, holding every cell exactly as a person would type it, with assets (chart images, Markdown notes) in sheets/<id>/assets/ and shared notes in assets/. The editor reloads a workbook the moment its file changes, so the person watches your edits land.',
    'Read AGENTS.md, and use the project skills in .agents/skills: sheet-authoring before you create or change a workbook, current-sheet whenever the person says "this cell", "the selection", "this sheet" or "here", create-sheet for a new workbook, apply-notes to act on the notes left on cells. Read the workbook\'s notes and the shared ones before you change it.',
    opts.mcp === false
      ? 'Change workbooks the way sheet-authoring describes.'
      : 'Change workbooks through the open-spreadsheet MCP tools (read_range, write_cells, set_style, set_layout, insert_rows, insert_columns, add_sheet, rename_sheet, set_names, create_workbook, duplicate_workbook, trace_cell, check and the rest): they recalculate, keep formulas, references, styles and merges consistent when rows, columns or sheets change, refuse to touch protected ranges, and refuse a write based on a stale read when you pass the expected_hash from read_range. Prefer them to editing workbook.json by hand.',
    'If you edit workbook.json directly, keep it valid JSON with every input a string and the sheet order kept, and never insert, delete or move rows, columns or sheets by hand (references would break; insert_rows and insert_columns, with a negative count to delete, do it right). Never change protected ranges by any route: ask the person to unprotect them.',
    'Write only under sheets/ and assets/. Ask before changing anything else in the workspace, such as open-spreadsheet.config.json or package.json.',
    `Put totals and derived numbers in formulas, not typed results. After a change, run the check tool (or \`${opts.checkCommand}\`) and fix what it reports. When your turn ends the panel also checks every workbook you changed and shows the person what it finds.`,
    'Reply in the language of the person\'s latest message; in Chinese, keep their script and regional wording (Traditional for Traditional, Simplified for Simplified).',
    'Keep replies short and concrete: the panel is narrow. Prefer doing the work over describing it, and finish with a one-line summary of what changed.',
    askInstruction,
  ].join(' ');
}

// Titles, sheet names, cell inputs and notes come from the workbook, so they
// sit in their own tags with `<` escaped and cannot close the context block.
const escapeTag = (s) => String(s).replace(/</g, '&lt;');

/** @param {AgentTurnContext | undefined} context */
function contextLines(context) {
  if (!context?.workbook) return [];
  const file = `sheets/${context.workbook}/workbook.json`;
  const lines = [`The person is viewing workbook "${context.workbook}": ${file}.`];
  if (context.title) lines.push(`<workbook-title>${escapeTag(context.title)}</workbook-title>`);
  if (context.sheet) {
    const where = context.active
      ? `the active cell is ${context.active}${context.selection && context.selection !== context.active ? ` and the selection is ${context.selection}` : ''}`
      : null;
    lines.push(`The open sheet is in <sheet-name>; ${where ?? 'nothing is selected'}. "This cell", "the selection" and "this sheet" mean them; node_modules/.open-spreadsheet/current.json has the latest (the current-sheet skill), since the person may move while you work.`);
    lines.push(`<sheet-name>${escapeTag(context.sheet)}</sheet-name>`);
  }
  if (context.input) lines.push(`The active cell holds <cell-input>${escapeTag(context.input)}</cell-input>.`);
  if (context.note) lines.push(`It has a note: <cell-note>${escapeTag(context.note)}</cell-note>`);
  return lines;
}

/** The person's message with what they have on screen. */
export function buildTurnPrompt(text, context) {
  const lines = contextLines(context);
  if (lines.length === 0) return text;
  return `${text}\n\n<open-spreadsheet-context>\n${lines.join('\n')}\n</open-spreadsheet-context>`;
}

const QUESTIONS_RE = /<open-spreadsheet-questions>([\s\S]*?)<\/open-spreadsheet-questions>/i;

export const TEXT_QUESTIONS_INSTRUCTION =
  'You cannot ask the person anything interactively here, and nothing you print in the middle of a turn can be answered. When you need answers before continuing, stop working and end your turn with a final message holding only an <open-spreadsheet-questions> block: a JSON array of at most 4 questions shaped like [{"header":"Period","question":"Which months should the summary cover?","multiSelect":false,"options":[{"label":"This quarter (Recommended)","description":"January to March"}]}], each with 2 to 4 options, written in the person\'s language. The panel shows them as a form and sends the answers back as your next message. Write nothing after the block and do not continue the task in the same turn.';

/**
 * Splits a reply written under TEXT_QUESTIONS_INSTRUCTION into its prose and
 * questions; null when it holds no usable block.
 * @returns {{ text: string, questions: AgentQuestion[] } | null}
 */
export function extractQuestionBlock(text) {
  const match = QUESTIONS_RE.exec(text);
  if (!match) return null;
  const body = match[1]
    .trim()
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/, '');
  let raw;
  try {
    raw = JSON.parse(body);
  } catch {
    return null;
  }
  const questions = parseQuestions({ questions: Array.isArray(raw) ? raw : raw?.questions });
  if (questions.length === 0) return null;
  return { text: text.replace(QUESTIONS_RE, '').trim(), questions };
}

/** @param {AgentQuestion[]} questions */
export function answersPrompt(questions, answers) {
  if (!answers) return 'I skipped your questions. Continue the task with your best judgement and sensible defaults.';
  const lines = questions.map((q) => `- ${q.question} ${answers[q.question] || '(no answer)'}`);
  return `My answers:\n${lines.join('\n')}\n\nContinue the task with these answers.`;
}

/** For agents whose history cannot be rewound: the edited message goes on top, and the agent sets the rest aside. */
export function editedPrompt(original, text) {
  return `I went back and edited an earlier message of mine. Treat everything after it in this conversation, including your replies, as discarded, and respond to the edited version instead. The workbooks keep whatever changes were already made, so read them again before changing them.\n\n<original-message>\n${escapeTag(original)}\n</original-message>\n\n${text}`;
}
