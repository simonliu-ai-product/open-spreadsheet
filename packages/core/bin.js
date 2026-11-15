#!/usr/bin/env node
// open-spreadsheet CLI.
//   open-spreadsheet dev [--port <n>] [--host] [--open]   the grid editor for this workspace
//   open-spreadsheet init <dir>                           a new workspace with a starter workbook, skills and AGENTS.md
//   open-spreadsheet check [--json]                       broken formulas and malformed workbooks; exits 1 on errors
//   open-spreadsheet sync-skills                          refresh .claude/skills and .agents/skills
//   open-spreadsheet export <id> [--sheet] [--range] [--format csv|json] [--out] [--root]
//                                                         computed values, for open-dashboard and other tools
//   open-spreadsheet-mcp                                  the MCP server over stdio (a separate bin)
import process from 'node:process';
import { mkdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { checkWorkspace, exportRows, isWorkspace, OpsError, readWorkbook, toCsvText, toJsonRecords } from './ops.js';
import { initWorkspace, startDev, syncSkills, VERSION } from './server.js';

const [command, ...rest] = process.argv.slice(2);
const flag = (name) => {
  const i = rest.indexOf(name);
  return i === -1 ? undefined : (rest[i + 1] ?? true);
};
const root = path.resolve(typeof flag('--root') === 'string' ? flag('--root') : process.cwd());
const needWorkspace = () => {
  if (!isWorkspace(root)) {
    process.stderr.write(`No open-spreadsheet.config.json in ${root}. Run open-spreadsheet init <dir> first.\n`);
    process.exit(1);
  }
};

if (command === 'dev') {
  const port = flag('--port');
  const host = flag('--host');
  await startDev({ root, port: port ? Number(port) : 5373, open: rest.includes('--open'), host: host === true ? true : host });
} else if (command === 'init') {
  const dir = rest.find((a) => !a.startsWith('-'));
  if (!dir) {
    process.stderr.write('Usage: open-spreadsheet init <dir>\n');
    process.exit(1);
  }
  const created = initWorkspace(dir);
  process.stdout.write(`Initialized workspace at ${created}\n  Starter workbook: sheets/getting-started/workbook.json\n  Agent files: AGENTS.md, CLAUDE.md, .mcp.json, .claude/skills, .agents/skills\n`);
} else if (command === 'check') {
  needWorkspace();
  const results = checkWorkspace(root);
  const errors = results.reduce((n, r) => n + r.problems.filter((p) => p.severity === 'error').length, 0);
  if (rest.includes('--json')) process.stdout.write(`${JSON.stringify({ ok: errors === 0, workbooks: results }, null, 2)}\n`);
  else {
    for (const r of results) {
      if (!r.problems.length) {
        process.stdout.write(`✓ ${r.id}\n`);
        continue;
      }
      process.stdout.write(`${r.problems.some((p) => p.severity === 'error') ? '✗' : '!'} ${r.id}\n`);
      for (const p of r.problems) process.stdout.write(`    ${p.severity === 'error' ? 'error  ' : 'warning'} ${[p.sheet, p.address].filter(Boolean).join('!')}  ${p.message}\n`);
    }
    process.stdout.write(`\n${results.length} workbook${results.length === 1 ? '' : 's'}, ${errors} error${errors === 1 ? '' : 's'}\n`);
  }
  process.exit(errors ? 1 : 0);
} else if (command === 'export') {
  needWorkspace();
  const id = rest.find((a, i) => !a.startsWith('-') && !rest[i - 1]?.startsWith('--'));
  if (!id) {
    process.stderr.write('Usage: open-spreadsheet export <workbook> [--sheet <name>] [--range A1:F20] [--format csv|json] [--out <file>] [--root <workspace>]\n');
    process.exit(1);
  }
  try {
    const { data } = readWorkbook(root, id);
    const range = typeof flag('--range') === 'string' ? flag('--range') : undefined;
    const { rows, firstCol } = exportRows(data, typeof flag('--sheet') === 'string' ? flag('--sheet') : undefined, range);
    const out = typeof flag('--out') === 'string' ? flag('--out') : undefined;
    const format = flag('--format') === 'json' || out?.endsWith('.json') ? 'json' : 'csv';
    const text = format === 'json' ? `${JSON.stringify(toJsonRecords(rows, firstCol), null, 2)}\n` : toCsvText(rows);
    if (out) {
      mkdirSync(path.dirname(path.resolve(out)), { recursive: true });
      writeFileSync(path.resolve(out), text);
      process.stderr.write(`Wrote ${rows.length - 1} rows to ${out}\n`);
    } else process.stdout.write(text);
  } catch (e) {
    process.stderr.write(`${e instanceof OpsError ? e.message : String(e)}\n`);
    process.exit(1);
  }
} else if (command === 'sync-skills') {
  needWorkspace();
  const names = syncSkills(root);
  process.stdout.write(`Synced skills: ${names.join(', ')}\n`);
} else if (command === '--version' || command === '-v') {
  process.stdout.write(`${VERSION}\n`);
} else {
  process.stdout.write(
    `open-spreadsheet ${VERSION} — Excel-like spreadsheets for agents and people\n\n  open-spreadsheet dev [--port <n>] [--host] [--open]\n  open-spreadsheet init <dir>\n  open-spreadsheet check [--json]\n  open-spreadsheet sync-skills\n  open-spreadsheet export <workbook> [--sheet] [--range] [--format csv|json] [--out] [--root]\n  open-spreadsheet-mcp            MCP server over stdio\n`,
  );
  process.exit(command ? 1 : 0);
}
