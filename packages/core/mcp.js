#!/usr/bin/env node
/**
 * `open-spreadsheet-mcp` — an MCP server over stdio for Claude Code,
 * Antigravity, Gemini CLI and any client that launches its own servers.
 *
 * The workspace is the first of: `--cwd <dir>`, OPEN_SPREADSHEET_CWD, the
 * launch directory or one of its parents. Every write recalculates, can carry
 * the hash it was based on (refused if the file changed meanwhile), and
 * reports the formula problems it leaves behind.
 *
 * Speaks JSON-RPC 2.0, one message per line, with no dependencies.
 */
import { assetsDir, listAssets, readAsset, renderChart, trashAsset, writeNote } from './assets.js';
import { existsSync, cpSync } from 'node:fs';
import path from 'node:path';
import { createInterface } from 'node:readline';
import { columnIndex, columnName, formatA1, parseA1, Workbook } from '@open-spreadsheet/engine';
import { checkWorkbook, checkWorkspace, evaluate, ID, isWorkspace, listWorkbooks, OpsError, readCurrent, nextCopyId, readRange, readWorkbook, trashWorkbook, traceCell, writeWorkbook } from './ops.js';
import { VERSION } from './server.js';

const argCwd = (() => {
  const i = process.argv.indexOf('--cwd');
  return i >= 0 ? process.argv[i + 1] : undefined;
})();

function findRoot() {
  const explicit = argCwd ?? process.env.OPEN_SPREADSHEET_CWD;
  if (explicit) return existsSync(explicit) && isWorkspace(path.resolve(explicit)) ? path.resolve(explicit) : null;
  for (let dir = process.cwd(); ; dir = path.dirname(dir)) {
    if (isWorkspace(dir)) return dir;
    if (path.dirname(dir) === dir) return null;
  }
}

const str = { type: 'string' };
const TOOLS = [
  {
    name: 'list_workbooks',
    description: 'List the workbooks in this open-spreadsheet workspace: id, title, sheet names, last change.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'read_range',
    description:
      'Read cells from a sheet: for each non-empty cell its address, input (what was typed; formulas start with =), computed value, and how it displays. Omit range to read the used area. Returns the file hash to pass to writes as expected_hash.',
    inputSchema: { type: 'object', properties: { workbook: str, sheet: { ...str, description: 'Sheet name; the first sheet when omitted' }, range: { ...str, description: 'e.g. A1:D20' } }, required: ['workbook'] },
  },
  {
    name: 'trace_cell',
    description:
      "Where a cell's value comes from and what uses it, one level each way (Excel's Trace Precedents / Dependents): the cells and ranges its formula reads, with defined names resolved, and the formulas that read it. Use before changing a cell other formulas depend on.",
    inputSchema: { type: 'object', properties: { workbook: str, sheet: { ...str, description: 'Sheet name; the first sheet when omitted' }, address: { ...str, description: 'e.g. B4' } }, required: ['workbook', 'address'] },
  },
  {
    name: 'write_cells',
    description:
      'Set cells to what a person would type: "120", "2026-01-05", "12%", "=SUM(B2:B9)". An empty string or null clears a cell. Returns the new values of the written cells and any formula problems left in the workbook.',
    inputSchema: {
      type: 'object',
      properties: {
        workbook: str,
        sheet: str,
        cells: { type: 'object', description: 'A1 address → input string (or null to clear)', additionalProperties: { type: ['string', 'null'] } },
        expected_hash: { ...str, description: 'Hash from read_range; the write is refused if the file changed since' },
      },
      required: ['workbook', 'sheet', 'cells'],
    },
  },
  {
    name: 'set_style',
    description:
      'Style a range. style keys: b, i, u, s (booleans), align ("left"|"center"|"right"), wrap, size (font size in points; 11 is the default, e.g. 14 for a title), valign ("top"|"middle"|"bottom"), check (true: a checkbox; the cell holds "TRUE"/"FALSE"), color, bg ("#rrggbb"), border ({ top, right, bottom, left: "#rrggbb", thick: true }), fmt (Excel number format like "#,##0", "0.0%", "\\"NT$\\"#,##0", "yyyy-mm-dd"). Pass null to clear a key; style: null clears all formatting.',
    inputSchema: {
      type: 'object',
      properties: { workbook: str, sheet: str, range: str, style: { type: ['object', 'null'] }, expected_hash: str },
      required: ['workbook', 'sheet', 'range', 'style'],
    },
  },
  {
    name: 'set_layout',
    description:
      'Change a sheet\'s layout: column widths ({"A": 160}), row heights ({"1": 32}), frozen panes ({"rows":1,"cols":0}), merges to add or remove ("A1:D1"), tab colour, the autofilter, dropdown lists (validation) and conditional formats.',
    inputSchema: {
      type: 'object',
      properties: {
        workbook: str,
        sheet: str,
        cols: { type: 'object', additionalProperties: { type: ['number', 'null'] } },
        rows: { type: 'object', additionalProperties: { type: ['number', 'null'] } },
        freeze: { type: 'object', properties: { rows: { type: 'number' }, cols: { type: 'number' } } },
        merge: { type: 'array', items: str },
        unmerge: { type: 'array', items: str },
        color: { type: ['string', 'null'] },
        filter: { type: ['string', 'null'], description: 'Autofilter range, header row first (e.g. "A1:F200"); null removes it' },
        charts: {
          type: 'array',
          items: { type: 'object' },
          description: 'Only [] is accepted: it removes charts an older workbook still has on its grid. New charts are images in Assets — use render_chart.',
        },
        hidden_rows: { type: 'array', items: { type: 'number' }, description: 'Replaces the rows hidden by hand (1-based numbers); [] shows them all' },
        hidden_cols: { type: 'array', items: { type: 'string' }, description: 'Replaces the columns hidden by hand (letters); [] shows them all' },
        validation: { type: 'array', description: 'Replaces the dropdown lists: [{ "range": "C2:C100", "list": ["Open", "Done"] }]', items: { type: 'object' } },
        conditional: {
          type: 'array',
          description:
            'Replaces the conditional formats: { range, type: gt|lt|ge|le|eq|ne (value) | between (value, value2) | contains (text) | duplicate | colorScale (colors: 2 or 3 hex) | dataBar (color), style?: { bg, color, b } }',
          items: { type: 'object' },
        },
        expected_hash: str,
      },
      required: ['workbook', 'sheet'],
    },
  },
  {
    name: 'create_workbook',
    description: 'Create a workbook at sheets/<id>/workbook.json. sheets is optional: [{ name, cells, styles, cols, freeze }]; a blank Sheet1 otherwise.',
    inputSchema: { type: 'object', properties: { id: { ...str, description: 'lowercase letters, numbers and dashes' }, title: str, sheets: { type: 'array', items: { type: 'object' } } }, required: ['id', 'title'] },
  },
  {
    name: 'list_assets',
    description:
      "A workbook's images and Markdown notes (sheets/<id>/assets/), or the workspace's shared ones (scope \"shared\"). Notes are context people keep about the numbers — read them with read_asset before working on a workbook.",
    inputSchema: { type: 'object', properties: { scope: { ...str, description: 'A workbook id, or "shared"' } }, required: ['scope'] },
  },
  {
    name: 'read_asset',
    description: 'The text of a Markdown note in a scope. Images are listed by list_assets; their content is not returned.',
    inputSchema: { type: 'object', properties: { scope: str, name: { ...str, description: 'e.g. notes.md' } }, required: ['scope', 'name'] },
  },
  {
    name: 'write_note',
    description: 'Create or replace a Markdown note (a .md file) in a workbook\'s assets or the shared ones: explanations, sources, decisions, a summary of what you found.',
    inputSchema: { type: 'object', properties: { scope: str, name: { ...str, description: 'ends in .md' }, text: str }, required: ['scope', 'name', 'text'] },
  },
  {
    name: 'render_chart',
    description:
      'Draw a chart from a range and save it as an SVG image in the workbook\'s assets (or shared): first column the categories, first row the series names when they are text. Charts are not placed on the grid; they live in Assets. Use it when a note or a person asks for a chart.',
    inputSchema: {
      type: 'object',
      properties: {
        workbook: str,
        sheet: str,
        range: { ...str, description: 'e.g. A1:C13' },
        type: { type: 'string', enum: ['bar', 'line', 'pie'] },
        title: str,
        name: { ...str, description: 'file name; from the title when omitted' },
        scope: { ...str, description: 'where to save: the workbook (default) or "shared"' },
        replace: { type: 'boolean', description: 'overwrite an image of that name instead of numbering a new one' },
      },
      required: ['workbook', 'range'],
    },
  },
  {
    name: 'delete_asset',
    description: 'Move an image or note to its folder\'s .trash (nothing is erased). Only when the person asked.',
    inputSchema: { type: 'object', properties: { scope: str, name: str }, required: ['scope', 'name'] },
    annotations: { destructiveHint: true },
  },
  {
    name: 'rename_workbook',
    description: "Change a workbook's title (what the list shows). Its id, and so its address, stay.",
    inputSchema: { type: 'object', properties: { workbook: str, title: str, expected_hash: str }, required: ['workbook', 'title'] },
  },
  {
    name: 'duplicate_workbook',
    description: 'Copy a whole workbook — every sheet, style, note, name and chart — to new_id, or <id>-copy when omitted, titled "… (copy)". Use it to try a variation without touching the original.',
    inputSchema: { type: 'object', properties: { workbook: str, new_id: { ...str, description: 'lowercase letters, numbers and dashes' } }, required: ['workbook'] },
  },
  {
    name: 'delete_workbook',
    description: 'Move a workbook to sheets/.trash/ (nothing is erased; moving the folder back restores it). Only when the person asked for the workbook to go.',
    inputSchema: { type: 'object', properties: { workbook: str }, required: ['workbook'] },
    annotations: { destructiveHint: true },
  },
  {
    name: 'add_sheet',
    description: 'Add an empty sheet (after the given sheet, or at the end).',
    inputSchema: { type: 'object', properties: { workbook: str, name: str, after: str }, required: ['workbook', 'name'] },
  },
  {
    name: 'rename_sheet',
    description: 'Rename a sheet; formulas that refer to it are rewritten.',
    inputSchema: { type: 'object', properties: { workbook: str, from: str, to: str }, required: ['workbook', 'from', 'to'] },
  },
  {
    name: 'delete_sheet',
    description: 'Delete a sheet; formulas that pointed into it become #REF!.',
    inputSchema: { type: 'object', properties: { workbook: str, sheet: str }, required: ['workbook', 'sheet'] },
  },
  {
    name: 'insert_rows',
    description: 'Insert (count > 0) or delete (count < 0) rows before a 1-based row number; formulas, styles and merges follow.',
    inputSchema: { type: 'object', properties: { workbook: str, sheet: str, before: { type: 'number' }, count: { type: 'number' } }, required: ['workbook', 'sheet', 'before', 'count'] },
  },
  {
    name: 'insert_columns',
    description: 'Insert (count > 0) or delete (count < 0) columns before a column letter; formulas, styles and merges follow.',
    inputSchema: { type: 'object', properties: { workbook: str, sheet: str, before: str, count: { type: 'number' } }, required: ['workbook', 'sheet', 'before', 'count'] },
  },
  {
    name: 'list_notes',
    description:
      'Every note on every sheet (or one workbook): sheet, address, the cell\'s input and value, and the note. Notes are comments people leave on cells — often requests for you ("make this a percentage", "check this total"); see the apply-notes skill.',
    inputSchema: { type: 'object', properties: { workbook: str } },
  },
  {
    name: 'set_note',
    description: 'Add, change or remove (empty text) the note on a cell. Remove a note once you have done what it asked.',
    inputSchema: { type: 'object', properties: { workbook: str, sheet: str, address: str, text: str, expected_hash: str }, required: ['workbook', 'sheet', 'address', 'text'] },
  },
  {
    name: 'set_names',
    description:
      'Define, change or remove (null) workbook names, e.g. { "Sales": "Orders!$F$2:$F$200", "TaxRate": "Settings!$B$1" }. Formulas can then say =SUM(Sales). A name refers to a cell or range with its sheet.',
    inputSchema: {
      type: 'object',
      properties: { workbook: str, names: { type: 'object', additionalProperties: { type: ['string', 'null'] } }, expected_hash: str },
      required: ['workbook', 'names'],
    },
  },
  {
    name: 'evaluate',
    description: 'Compute a formula against a workbook without changing it, e.g. "=SUMIF(Orders!B:B,\\"May\\",Orders!C:C)".',
    inputSchema: { type: 'object', properties: { workbook: str, sheet: str, formula: str }, required: ['workbook', 'formula'] },
  },
  {
    name: 'check',
    description: 'Find problems: formulas showing #NAME?, #REF!, #CIRC!, #VALUE!, malformed addresses, overlapping merges. One workbook, or all.',
    inputSchema: { type: 'object', properties: { workbook: str } },
  },
  {
    name: 'get_current',
    description: 'What the person has open in the editor: workbook, sheet, active cell, selected range, and that cell\'s input and value.',
    inputSchema: { type: 'object', properties: {} },
  },
];

// ---------- tools ----------

/** Whether a cell lies in a protected range of the sheet. */
const isProtected = (sheet, row, col) =>
  (sheet.protected ?? []).some((t) => {
    const r = parseRange(t);
    return row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;
  });
const refuseProtected = (sheet, cells) => {
  const hit = cells.find((c) => isProtected(sheet, c.row, c.col));
  if (hit)
    throw new OpsError(
      'protected',
      `${formatA1(hit.row, hit.col)} on ${sheet.name} is in a protected range (${sheet.protected.join(', ')}). Ask the person to unprotect it in the editor; agents cannot change protected cells or remove protection.`,
    );
};

/** Refuses a style change on a rectangle that overlaps a protected range. */
const refuseProtectedRect = (sheet, r) => {
  const hit = (sheet.protected ?? []).find((t) => {
    const p = parseRange(t);
    return p.r0 <= r.r1 && r.r0 <= p.r1 && p.c0 <= r.c1 && r.c0 <= p.c1;
  });
  if (hit)
    throw new OpsError(
      'protected',
      `${hit} on ${sheet.name} is protected. Ask the person to unprotect it in the editor; agents cannot change protected cells or remove protection.`,
    );
};

const sheetOf = (data, name) => {
  const s = data.sheets.find((x) => x.name === name);
  if (!s) throw new OpsError('no_sheet', `No sheet "${name}". Sheets: ${data.sheets.map((x) => x.name).join(', ')}`);
  return s;
};

/** Loads a workbook into the engine, lets `fn` change it, and writes it back with the cells the engine holds. */
async function mutate(root, id, expectedHash, fn) {
  const { data, hash } = readWorkbook(root, id);
  if (expectedHash && expectedHash !== hash) throw new OpsError('conflict', `sheets/${id}/workbook.json changed since it was read. Read it again and redo the change.`);
  const wb = Workbook.fromJSON(data);
  const extra = (await fn(data, wb)) ?? {};
  const cells = new Map(wb.toJSON().sheets.map((s) => [s.name, s.cells]));
  const out = { ...data, sheets: data.sheets.map((s) => ({ ...s, cells: cells.get(s.name) ?? {} })) };
  const written = await writeWorkbook(root, id, out, { expectedHash: hash });
  const problems = checkWorkbook(out).filter((p) => p.severity === 'error');
  return { hash: written.hash, ...extra, problems: problems.slice(0, 30), ...(problems.length > 30 ? { more_problems: problems.length - 30 } : {}) };
}

function parseRange(text) {
  const [a, b] = String(text).toUpperCase().replaceAll('$', '').split(':');
  const p = parseA1(a ?? '');
  const q = b ? parseA1(b) : p;
  if (!p || !q) throw new OpsError('bad_range', `Not a range: ${text}`);
  return { r0: Math.min(p.row, q.row), c0: Math.min(p.col, q.col), r1: Math.max(p.row, q.row), c1: Math.max(p.col, q.col) };
}

function shiftMeta(sheet, axis, at, count) {
  const move = (i) => (i < at ? i : count < 0 && i < at - count ? null : i + count);
  if (sheet.styles) {
    const next = {};
    for (const [a, s] of Object.entries(sheet.styles)) {
      const p = parseA1(a);
      if (!p) continue;
      const m = move(axis === 'row' ? p.row : p.col);
      if (m !== null) next[axis === 'row' ? formatA1(m, p.col) : formatA1(p.row, m)] = s;
    }
    sheet.styles = next;
  }
  if (sheet.notes) {
    const next = {};
    for (const [a, t] of Object.entries(sheet.notes)) {
      const p = parseA1(a);
      if (!p) continue;
      const m = move(axis === 'row' ? p.row : p.col);
      if (m !== null) next[axis === 'row' ? formatA1(m, p.col) : formatA1(p.row, m)] = t;
    }
    sheet.notes = next;
  }
  if (axis === 'col' && sheet.cols) sheet.cols = Object.fromEntries(Object.entries(sheet.cols).flatMap(([k, v]) => (move(columnIndex(k)) === null ? [] : [[columnName(move(columnIndex(k))), v]])));
  if (axis === 'row' && sheet.rows) sheet.rows = Object.fromEntries(Object.entries(sheet.rows).flatMap(([k, v]) => (move(Number(k) - 1) === null ? [] : [[String(move(Number(k) - 1) + 1), v]])));
  if (sheet.protected) {
    sheet.protected = sheet.protected.flatMap((t) => {
      const r = parseRange(t);
      const lo = axis === 'row' ? r.r0 : r.c0;
      const hi = axis === 'row' ? r.r1 : r.c1;
      const a = move(lo);
      const b = move(hi);
      if (a === null && b === null) return [];
      const n = axis === 'row' ? { ...r, r0: a ?? at, r1: b ?? at - 1 } : { ...r, c0: a ?? at, c1: b ?? at - 1 };
      if (n.r1 < n.r0 || n.c1 < n.c0) return [];
      return [n.r0 === n.r1 && n.c0 === n.c1 ? formatA1(n.r0, n.c0) : `${formatA1(n.r0, n.c0)}:${formatA1(n.r1, n.c1)}`];
    });
    if (!sheet.protected.length) delete sheet.protected;
  }
  if (sheet.merges) {
    sheet.merges = sheet.merges.flatMap((m) => {
      const r = parseRange(m);
      const lo = axis === 'row' ? r.r0 : r.c0;
      const hi = axis === 'row' ? r.r1 : r.c1;
      const a = move(lo);
      const b = move(hi);
      if (a === null && b === null) return [];
      const n = axis === 'row' ? { ...r, r0: a ?? at, r1: b ?? at - 1 } : { ...r, c0: a ?? at, c1: b ?? at - 1 };
      if (n.r1 < n.r0 || n.c1 < n.c0 || (n.r0 === n.r1 && n.c0 === n.c1)) return [];
      return [`${formatA1(n.r0, n.c0)}:${formatA1(n.r1, n.c1)}`];
    });
  }
}

async function callTool(root, name, a) {
  switch (name) {
    case 'list_workbooks':
      return { workspace: root, workbooks: listWorkbooks(root) };
    case 'read_range': {
      const { data, hash } = readWorkbook(root, a.workbook);
      return { workbook: a.workbook, hash, sheets: data.sheets.map((s) => s.name), ...readRange(data, a.sheet, a.range) };
    }
    case 'trace_cell':
      return traceCell(readWorkbook(root, a.workbook).data, a.sheet, a.address);
    case 'write_cells':
      return mutate(root, a.workbook, a.expected_hash, (data, wb) => {
        const target = sheetOf(data, a.sheet);
        const cells = Object.entries(a.cells ?? {}).map(([addr, input]) => {
          const p = parseA1(addr);
          if (!p) throw new OpsError('bad_address', `Not an A1 address: ${addr}`);
          if (input !== null && typeof input !== 'string') throw new OpsError('bad_input', `${addr}: inputs are strings, e.g. "${input}"`);
          return { ...p, input: input ?? '' };
        });
        refuseProtected(target, cells);
        wb.setCells(a.sheet, cells);
        const values = Object.fromEntries(cells.map((c) => [formatA1(c.row, c.col), wb.value(a.sheet, c.row, c.col)?.toString?.() ?? null]));
        return { written: cells.length, values };
      });
    case 'set_style':
      return mutate(root, a.workbook, a.expected_hash, (data) => {
        const sheet = sheetOf(data, a.sheet);
        const r = parseRange(a.range);
        refuseProtectedRect(sheet, r);
        sheet.styles ??= {};
        for (let row = r.r0; row <= r.r1; row++) {
          for (let col = r.c0; col <= r.c1; col++) {
            const addr = formatA1(row, col);
            if (a.style === null) {
              delete sheet.styles[addr];
              continue;
            }
            const next = { ...sheet.styles[addr], ...a.style };
            for (const k of Object.keys(next)) if (next[k] === null || next[k] === false || next[k] === '') delete next[k];
            if (Object.keys(next).length) sheet.styles[addr] = next;
            else delete sheet.styles[addr];
          }
        }
        return { styled: (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) };
      });
    case 'set_layout':
      return mutate(root, a.workbook, a.expected_hash, (data) => {
        const sheet = sheetOf(data, a.sheet);
        for (const [k, v] of Object.entries(a.cols ?? {})) {
          sheet.cols ??= {};
          if (v === null) delete sheet.cols[k.toUpperCase()];
          else sheet.cols[k.toUpperCase()] = Math.max(4, Math.round(v));
        }
        for (const [k, v] of Object.entries(a.rows ?? {})) {
          sheet.rows ??= {};
          if (v === null) delete sheet.rows[k];
          else sheet.rows[k] = Math.max(4, Math.round(v));
        }
        if (a.freeze) sheet.freeze = a.freeze.rows || a.freeze.cols ? { rows: a.freeze.rows ?? 0, cols: a.freeze.cols ?? 0 } : undefined;
        if (a.unmerge?.length) sheet.merges = (sheet.merges ?? []).filter((m) => !a.unmerge.includes(m));
        for (const m of a.merge ?? []) {
          parseRange(m);
          sheet.merges = [...(sheet.merges ?? []).filter((x) => x !== m), m.toUpperCase()];
        }
        if (a.color !== undefined) sheet.color = a.color ?? undefined;
        if (a.filter !== undefined) {
          if (a.filter === null) delete sheet.filter;
          else {
            parseRange(a.filter);
            sheet.filter = { range: a.filter.toUpperCase() };
          }
        }
        if (a.charts !== undefined) {
          if (a.charts.length) throw new OpsError('charts_in_assets', 'Charts are not placed on the grid any more. Use render_chart to save one as an image in the workbook\'s Assets.');
          sheet.charts = undefined;
        }
        if (a.hidden_rows !== undefined) sheet.hiddenRows = a.hidden_rows.length ? [...new Set(a.hidden_rows.map(Math.trunc))].sort((x, y) => x - y) : undefined;
        if (a.hidden_cols !== undefined) sheet.hiddenCols = a.hidden_cols.length ? [...new Set(a.hidden_cols.map((l) => String(l).toUpperCase()))] : undefined;
        if (a.validation !== undefined) {
          for (const v of a.validation) parseRange(v.range);
          sheet.validation = a.validation.length ? a.validation : undefined;
        }
        if (a.conditional !== undefined) {
          for (const r of a.conditional) parseRange(r.range);
          sheet.conditional = a.conditional.length ? a.conditional : undefined;
        }
      });
    case 'create_workbook': {
      if (!ID.test(a.id ?? '')) throw new OpsError('bad_id', 'Use lowercase letters, numbers and dashes for the id.');
      const sheets = Array.isArray(a.sheets) && a.sheets.length ? a.sheets.map((s, i) => ({ name: s.name ?? `Sheet${i + 1}`, cells: {}, ...s })) : [{ name: 'Sheet1', cells: {} }];
      const data = { title: a.title, sheets };
      const { hash } = await writeWorkbook(root, a.id, data, { create: true });
      return { id: a.id, hash, url: `/#/w/${a.id}`, problems: checkWorkbook(data) };
    }
    case 'list_assets':
      return { scope: a.scope, assets: listAssets(root, a.scope) };
    case 'read_asset': {
      if (!/\.md$/i.test(String(a.name ?? ''))) throw new OpsError('bad_name', 'read_asset reads notes (.md); images are listed by list_assets');
      return { scope: a.scope, name: a.name, text: readAsset(root, a.scope, a.name) };
    }
    case 'write_note':
      return writeNote(root, a.scope, a.name, a.text);
    case 'render_chart':
      return renderChart(root, a.workbook, a);
    case 'delete_asset':
      return trashAsset(root, a.scope, a.name);
    case 'rename_workbook': {
      const title = String(a.title ?? '').trim();
      if (!title || title.length > 200 || /[\n\r]/.test(title)) throw new OpsError('bad_title', 'A title is one line of 1 to 200 characters.');
      return mutate(root, a.workbook, a.expected_hash, (data) => {
        data.title = title;
        return { title };
      });
    }
    case 'duplicate_workbook': {
      const { data } = readWorkbook(root, a.workbook);
      const taken = new Set(listWorkbooks(root).map((w) => w.id));
      if (a.new_id !== undefined && (!ID.test(a.new_id) || taken.has(a.new_id))) {
        throw new OpsError(taken.has(a.new_id) ? 'exists' : 'bad_id', taken.has(a.new_id) ? `A workbook ${a.new_id} exists.` : 'Use lowercase letters, numbers and dashes for the id.');
      }
      const id = a.new_id ?? nextCopyId(a.workbook, taken);
      const title = /\(copy( \d+)?\)$/.test(data.title) ? data.title : `${data.title} (copy)`;
      const { hash } = await writeWorkbook(root, id, { ...data, title }, { create: true });
      const from = assetsDir(root, a.workbook);
      if (existsSync(from)) cpSync(from, assetsDir(root, id), { recursive: true, filter: (f) => !f.includes(`${path.sep}.trash`) });
      return { id, title, hash, url: `/#/w/${id}` };
    }
    case 'delete_workbook':
      return trashWorkbook(root, a.workbook);
    case 'add_sheet':
      return mutate(root, a.workbook, a.expected_hash, (data, wb) => {
        if (data.sheets.some((s) => s.name.toLowerCase() === String(a.name).toLowerCase())) throw new OpsError('exists', `A sheet named ${a.name} exists.`);
        const at = a.after ? data.sheets.findIndex((s) => s.name === a.after) + 1 : data.sheets.length;
        data.sheets.splice(at || data.sheets.length, 0, { name: a.name, cells: {} });
        wb.addSheet(a.name);
      });
    case 'rename_sheet':
      return mutate(root, a.workbook, a.expected_hash, (data, wb) => {
        const sheet = sheetOf(data, a.from);
        wb.renameSheet(a.from, a.to);
        sheet.name = a.to;
      });
    case 'delete_sheet':
      return mutate(root, a.workbook, a.expected_hash, (data, wb) => {
        const doomed = sheetOf(data, a.sheet);
        if (data.sheets.length === 1) throw new OpsError('last_sheet', 'A workbook needs at least one sheet.');
        if (doomed.protected?.length)
          throw new OpsError('protected', `${a.sheet} has protected ranges (${doomed.protected.join(', ')}). Ask the person to unprotect them in the editor before it can be deleted.`);
        wb.removeSheet(a.sheet);
        data.sheets = data.sheets.filter((s) => s.name !== a.sheet);
      });
    case 'insert_rows':
      return mutate(root, a.workbook, a.expected_hash, (data, wb) => {
        const sheet = sheetOf(data, a.sheet);
        wb.insertRows(a.sheet, a.before - 1, a.count);
        shiftMeta(sheet, 'row', a.before - 1, a.count);
      });
    case 'insert_columns':
      return mutate(root, a.workbook, a.expected_hash, (data, wb) => {
        const sheet = sheetOf(data, a.sheet);
        const at = columnIndex(String(a.before).toUpperCase());
        wb.insertColumns(a.sheet, at, a.count);
        shiftMeta(sheet, 'col', at, a.count);
      });
    case 'list_notes': {
      const ids = a.workbook ? [a.workbook] : listWorkbooks(root).map((w) => w.id);
      const notes = [];
      for (const id of ids) {
        const { data, hash } = readWorkbook(root, id);
        const wb = Workbook.fromJSON(data);
        for (const s of data.sheets) {
          for (const [address, note] of Object.entries(s.notes ?? {})) {
            const p = parseA1(address);
            if (!p) continue;
            const v = wb.value(s.name, p.row, p.col);
            notes.push({ workbook: id, hash, sheet: s.name, address, input: wb.inputAt(s.name, p.row, p.col), value: v?.toString?.() ?? null, note });
          }
        }
      }
      return { count: notes.length, notes };
    }
    case 'set_note':
      return mutate(root, a.workbook, a.expected_hash, (data) => {
        const sheet = sheetOf(data, a.sheet);
        const p = parseA1(String(a.address).toUpperCase());
        if (!p) throw new OpsError('bad_address', `Not an A1 address: ${a.address}`);
        const key = formatA1(p.row, p.col);
        sheet.notes ??= {};
        if (a.text) sheet.notes[key] = a.text;
        else delete sheet.notes[key];
        if (!Object.keys(sheet.notes).length) delete sheet.notes;
      });
    case 'set_names':
      return mutate(root, a.workbook, a.expected_hash, (data, wb) => {
        const next = { ...wb.namesJSON() };
        for (const [n, ref] of Object.entries(a.names ?? {})) {
          const key = Object.keys(next).find((k) => k.toLowerCase() === n.toLowerCase()) ?? n;
          if (ref === null || ref === '') delete next[key];
          else next[key] = String(ref).replace(/^=/, '');
        }
        try {
          wb.setNames(next);
        } catch (e) {
          throw new OpsError('bad_name', e.message);
        }
        if (Object.keys(next).length) data.names = next;
        else delete data.names;
        return { names: next };
      });
    case 'evaluate': {
      const { data } = readWorkbook(root, a.workbook);
      const v = evaluate(data, a.sheet, a.formula);
      return { formula: a.formula, value: v };
    }
    case 'check':
      if (a.workbook) return { workbook: a.workbook, problems: checkWorkbook(readWorkbook(root, a.workbook).data) };
      return { workbooks: checkWorkspace(root) };
    case 'get_current':
      return readCurrent(root) ?? { open: false, note: 'The editor is not open, or nothing is selected yet. Ask which workbook and cells.' };
    default:
      throw new OpsError('unknown_tool', `No tool ${name}`);
  }
}

// ---------- JSON-RPC over stdio ----------

const send = (msg) => process.stdout.write(`${JSON.stringify({ jsonrpc: '2.0', ...msg })}\n`);

async function handle(msg) {
  const { id, method, params } = msg;
  if (method === 'initialize') {
    return send({
      id,
      result: {
        protocolVersion: params?.protocolVersion ?? '2025-06-18',
        capabilities: { tools: { listChanged: false } },
        serverInfo: { name: 'open-spreadsheet', title: 'open-spreadsheet', version: VERSION },
        instructions:
          'Spreadsheets as JSON in sheets/<id>/workbook.json. Read with read_range, change with write_cells / set_style / set_layout, verify with check. Cells are typed strings; formulas start with =.',
      },
    });
  }
  if (method === 'ping') return send({ id, result: {} });
  if (method === 'tools/list') return send({ id, result: { tools: TOOLS } });
  if (method === 'tools/call') {
    const root = findRoot();
    if (!root) {
      return send({
        id,
        result: { isError: true, content: [{ type: 'text', text: 'No open-spreadsheet workspace here: run from a folder with open-spreadsheet.config.json, or pass --cwd <dir>.' }] },
      });
    }
    try {
      const out = await callTool(root, params?.name, params?.arguments ?? {});
      return send({ id, result: { content: [{ type: 'text', text: JSON.stringify(out, null, 2) }], structuredContent: out } });
    } catch (e) {
      const text = e instanceof OpsError ? `${e.code}: ${e.message}` : String(e?.message ?? e);
      return send({ id, result: { isError: true, content: [{ type: 'text', text }] } });
    }
  }
  if (id !== undefined && method !== undefined) send({ id, error: { code: -32601, message: `Method not found: ${method}` } });
}

const rl = createInterface({ input: process.stdin });
rl.on('line', (line) => {
  if (!line.trim()) return;
  let msg;
  try {
    msg = JSON.parse(line);
  } catch {
    return send({ id: null, error: { code: -32700, message: 'Parse error' } });
  }
  void handle(msg).catch((e) => msg.id !== undefined && send({ id: msg.id, error: { code: -32603, message: String(e?.message ?? e) } }));
});
