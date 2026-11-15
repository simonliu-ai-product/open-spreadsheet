/**
 * Workspace operations shared by the dev server, the CLI (check) and the MCP
 * server: list, read, write with optimistic concurrency, and check.
 *
 * Writes to one file go through a queue, and a write can carry the SHA-256 of
 * the version it was based on; if the file changed since, the write is
 * refused, so an agent and the viewer never silently overwrite each other.
 */
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, renameSync, statSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { CellError, columnName, formatA1, formatValue, impliedFormat, parseA1, Workbook } from '@open-spreadsheet/engine';

export const ID = /^[a-z0-9][a-z0-9._-]{0,63}$/;
export const workbookFile = (root, id) => path.join(root, 'sheets', id, 'workbook.json');
export const hashOf = (text) => createHash('sha256').update(text).digest('hex');
export const isWorkspace = (dir) => existsSync(path.join(dir, 'open-spreadsheet.config.json'));

export class OpsError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

export function listWorkbooks(root) {
  const dir = path.join(root, 'sheets');
  if (!existsSync(dir)) return [];
  return readdirSync(dir)
    .filter((id) => ID.test(id) && existsSync(workbookFile(root, id)))
    .map((id) => {
      let title = id;
      let sheets = 1;
      let names = [];
      try {
        const data = JSON.parse(readFileSync(workbookFile(root, id), 'utf8'));
        title = data.title || id;
        sheets = Array.isArray(data.sheets) ? data.sheets.length : 1;
        names = Array.isArray(data.sheets) ? data.sheets.map((s) => s.name) : [];
      } catch {
        // A file mid-write: show its id.
      }
      return { id, title, sheets, sheetNames: names, updatedAt: statSync(workbookFile(root, id)).mtime.toISOString() };
    })
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export function readWorkbook(root, id) {
  if (!ID.test(id ?? '')) throw new OpsError('bad_id', `Not a workbook id: ${id}`);
  const file = workbookFile(root, id);
  if (!existsSync(file)) throw new OpsError('not_found', `No workbook ${id}`);
  const text = readFileSync(file, 'utf8');
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new OpsError('bad_json', `sheets/${id}/workbook.json is not valid JSON: ${e.message}`);
  }
  return { data, hash: hashOf(text) };
}

const queues = new Map();
/** Runs writes to one file one after another. */
function queued(file, fn) {
  const prev = queues.get(file) ?? Promise.resolve();
  const next = prev.then(fn, fn);
  queues.set(
    file,
    next.catch(() => {}),
  );
  return next;
}

export const serializeWorkbook = (data) => `${JSON.stringify(data, null, 2)}\n`;

/**
 * Writes a workbook. With `expectedHash`, refuses (code "conflict") unless the
 * file on disk is still that version. Returns the new hash.
 */
export function writeWorkbook(root, id, data, { expectedHash, create = false } = {}) {
  if (!ID.test(id ?? '')) throw new OpsError('bad_id', `Not a workbook id: ${id}`);
  if (!data || !Array.isArray(data.sheets) || data.sheets.length === 0) throw new OpsError('bad_data', 'A workbook needs at least one sheet.');
  const file = workbookFile(root, id);
  return queued(file, () => {
    const exists = existsSync(file);
    if (create && exists) throw new OpsError('exists', `Workbook ${id} already exists.`);
    if (!create && !exists) throw new OpsError('not_found', `No workbook ${id}`);
    if (expectedHash && exists) {
      const current = hashOf(readFileSync(file, 'utf8'));
      if (current !== expectedHash) throw new OpsError('conflict', `sheets/${id}/workbook.json changed since it was read. Read it again and redo the change.`);
    }
    const text = serializeWorkbook(data);
    mkdirSync(path.dirname(file), { recursive: true });
    // Write then rename, so a reader never sees half a file.
    const tmp = `${file}.${process.pid}.tmp`;
    writeFileSync(tmp, text);
    renameSync(tmp, file);
    return { hash: hashOf(text), text };
  });
}

export function trashWorkbook(root, id) {
  const file = workbookFile(root, id);
  if (!existsSync(file)) throw new OpsError('not_found', `No workbook ${id}`);
  const trash = path.join(root, 'sheets', '.trash');
  mkdirSync(trash, { recursive: true });
  const to = path.join(trash, `${id}-${Date.now()}`);
  renameSync(path.dirname(file), to);
  return { id, trashedTo: path.relative(root, to) };
}

/** `budget` → `budget-copy`, then `budget-copy-2`…; a copy of a copy counts on from the original (as open-doc). */
export function nextCopyId(id, taken) {
  const stem = id.replace(/-copy(-\d+)?$/, '');
  let candidate = `${stem}-copy`;
  for (let n = 2; taken.has(candidate); n++) candidate = `${stem}-copy-${n}`;
  return candidate;
}

// ---------- reading cells for agents ----------

function parseRange(text) {
  const [a, b] = String(text).toUpperCase().replaceAll('$', '').split(':');
  const p = parseA1(a ?? '');
  const q = b ? parseA1(b) : p;
  if (!p || !q) return null;
  return { r0: Math.min(p.row, q.row), c0: Math.min(p.col, q.col), r1: Math.max(p.row, q.row), c1: Math.max(p.col, q.col) };
}

const shown = (v, fmt) => formatValue(v, fmt).text;
const plain = (v) => (v instanceof CellError ? v.code : v);

/** A sheet's cells in a range, with what was typed, what it computes to, and how it shows. */
/**
 * Where a cell's value comes from and what uses it: the cells and ranges its
 * formula reads (defined names resolved) and the formulas that read it.
 * One level each way, as Excel's Trace Precedents / Dependents.
 */
export function traceCell(data, sheetName, address) {
  const sheet = sheetName ? data.sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase()) : data.sheets[0];
  if (!sheet) throw new OpsError('no_sheet', `No sheet "${sheetName}". Sheets: ${data.sheets.map((s) => s.name).join(', ')}`);
  const p = parseA1(String(address ?? ''));
  if (!p) throw new OpsError('bad_address', `Not an A1 address: ${address}`);
  const wb = Workbook.fromJSON(data);
  const named = (n) => data.sheets.find((s) => s.name.toLowerCase() === n.toLowerCase())?.name ?? n;
  const cell = (s, row, col) => ({ sheet: s, address: formatA1(row, col), input: wb.inputAt(s, row, col), value: plain(wb.value(s, row, col)) });
  return {
    ...cell(sheet.name, p.row, p.col),
    precedents: wb.precedents(sheet.name, p.row, p.col).map((g) => {
      const s = named(g.sheet);
      return g.r0 === g.r1 && g.c0 === g.c1 ? cell(s, g.r0, g.c0) : { sheet: s, range: `${formatA1(g.r0, g.c0)}:${formatA1(g.r1, g.c1)}` };
    }),
    dependents: wb.dependents(sheet.name, p.row, p.col).map((d) => cell(named(d.sheet), d.row, d.col)),
  };
}

export function readRange(data, sheetName, range) {
  const sheet = sheetName ? data.sheets.find((s) => s.name === sheetName) : data.sheets[0];
  if (!sheet) throw new OpsError('no_sheet', `No sheet "${sheetName}". Sheets: ${data.sheets.map((s) => s.name).join(', ')}`);
  const wb = Workbook.fromJSON(data);
  const used = wb.used(sheet.name);
  const r = range ? parseRange(range) : { r0: 0, c0: 0, r1: Math.max(0, used.rows - 1), c1: Math.max(0, used.cols - 1) };
  if (!r) throw new OpsError('bad_range', `Not a range: ${range}`);
  if ((r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) > 20000) throw new OpsError('too_big', 'Read at most 20,000 cells at a time; pass a smaller range.');
  const cells = [];
  for (let row = r.r0; row <= r.r1; row++) {
    for (let col = r.c0; col <= r.c1; col++) {
      const input = wb.inputAt(sheet.name, row, col);
      if (!input) continue;
      const a1 = formatA1(row, col);
      const v = wb.value(sheet.name, row, col);
      const fmt = sheet.styles?.[a1]?.fmt ?? impliedFormat(input);
      const cell = { address: a1, input, value: plain(v) };
      const text = shown(v, fmt);
      if (text !== String(cell.value ?? '')) cell.display = text;
      if (sheet.notes?.[a1]) cell.note = sheet.notes[a1];
      cells.push(cell);
    }
  }
  const notes = Object.entries(sheet.notes ?? {})
    .filter(([a]) => {
      const p = parseA1(a);
      return p && p.row >= r.r0 && p.row <= r.r1 && p.col >= r.c0 && p.col <= r.c1 && !wb.inputAt(sheet.name, p.row, p.col);
    })
    .map(([address, note]) => ({ address, input: '', value: null, note }));
  cells.push(...notes);
  return { sheet: sheet.name, range: `${formatA1(r.r0, r.c0)}:${formatA1(r.r1, r.c1)}`, used: used.rows ? `A1:${formatA1(used.rows - 1, Math.max(0, used.cols - 1))}` : null, cells };
}

/** Computes a formula against a workbook without writing anything. */
export function evaluate(data, sheetName, formula) {
  const sheet = sheetName ?? data.sheets[0]?.name;
  const copy = JSON.parse(JSON.stringify(data));
  const target = copy.sheets.find((s) => s.name === sheet);
  if (!target) throw new OpsError('no_sheet', `No sheet "${sheet}"`);
  const wb = Workbook.fromJSON(copy);
  const used = wb.used(sheet);
  const scratch = formatA1(used.rows + 5, used.cols + 5);
  wb.set(sheet, scratch, formula.startsWith('=') ? formula : `=${formula}`);
  return plain(wb.get(sheet, scratch));
}

// ---------- export for other tools (open-dashboard's csv and json sources) ----------

const isDateFmt = (f) => Boolean(f && /[yd]|h:|m+:|:s/i.test(f.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '')));

/**
 * A sheet's computed values as rows: the first row of the range is the header.
 * Numbers stay plain (no grouping) so a CSV reader sees numbers; cells shown as
 * dates come out as ISO dates; errors come out empty.
 */
export function exportRows(data, sheetName, range) {
  const sheet = sheetName ? data.sheets.find((s) => s.name === sheetName) : data.sheets[0];
  if (!sheet) throw new OpsError('no_sheet', `No sheet "${sheetName}". Sheets: ${data.sheets.map((s) => s.name).join(', ')}`);
  const wb = Workbook.fromJSON(data);
  const used = wb.used(sheet.name);
  const r = range ? parseRange(range) : { r0: 0, c0: 0, r1: used.rows - 1, c1: used.cols - 1 };
  if (!r) throw new OpsError('bad_range', `Not a range: ${range}`);
  const rows = [];
  for (let row = r.r0; row <= r.r1; row++) {
    const line = [];
    for (let col = r.c0; col <= r.c1; col++) {
      const v = wb.value(sheet.name, row, col);
      const a1 = formatA1(row, col);
      const input = wb.inputAt(sheet.name, row, col);
      const fmt = sheet.styles?.[a1]?.fmt ?? (input.startsWith('=') ? undefined : impliedFormat(input));
      if (v === null || v instanceof CellError) line.push(null);
      else if (typeof v === 'number' && isDateFmt(fmt)) line.push(formatValue(v, Number.isInteger(v) ? 'yyyy-mm-dd' : 'yyyy-mm-dd hh:mm:ss').text);
      else line.push(v);
    }
    rows.push(line);
  }
  return { sheet: sheet.name, rows, firstCol: r.c0 };
}

export function toCsvText(rows) {
  const q = (v) => {
    const s = v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
    return /[",\n\r]/.test(s) ? `"${s.replaceAll('"', '""')}"` : s;
  };
  return `${rows.map((r) => r.map(q).join(',')).join('\n')}\n`;
}

/** Rows as objects keyed by the header row; blank header cells become column letters. */
export function toJsonRecords(rows, firstCol = 0) {
  const [head = [], ...body] = rows;
  const keys = head.map((h, i) => (h === null || h === '' ? columnName(firstCol + i) : String(h)));
  return body.filter((r) => r.some((v) => v !== null && v !== '')).map((r) => Object.fromEntries(keys.map((k, i) => [k, r[i] ?? null])));
}

// ---------- check ----------

/**
 * Problems in a workbook an agent should fix: formulas that fail to parse or
 * name unknown functions (#NAME?), broken references (#REF!), cycles
 * (#CIRC!), other errors, and malformed styles, merges and sizes.
 */
export function checkWorkbook(data) {
  const problems = [];
  const add = (severity, sheet, address, message) => problems.push({ severity, sheet, address, message });
  if (!Array.isArray(data.sheets) || !data.sheets.length) {
    add('error', '', '', 'The workbook has no sheets.');
    return problems;
  }
  const names = new Set();
  for (const s of data.sheets) {
    if (!s.name || /[:\\/?*[\]]/.test(s.name)) add('error', s.name ?? '', '', 'Sheet names cannot be empty or contain : \\ / ? * [ ]');
    if (names.has(String(s.name).toLowerCase())) add('error', s.name, '', 'Two sheets share this name.');
    names.add(String(s.name).toLowerCase());
    for (const a of Object.keys(s.cells ?? {})) if (!parseA1(a)) add('error', s.name, a, 'Not an A1 address.');
    for (const a of Object.keys(s.styles ?? {})) if (!parseA1(a)) add('warning', s.name, a, 'Style on something that is not an A1 address.');
    for (const rule of s.conditional ?? []) {
      if (!parseRange(rule.range ?? '')) add('error', s.name, rule.range ?? '', 'A conditional format needs a range like B2:B20.');
      if (!['gt', 'lt', 'ge', 'le', 'eq', 'ne', 'between', 'contains', 'duplicate', 'colorScale', 'dataBar', 'top', 'average', 'iconSet'].includes(rule.type)) add('error', s.name, rule.range ?? '', `Unknown conditional format type "${rule.type}".`);
      if (rule.type === 'top' && !(Number(rule.rank) >= 1)) add('error', s.name, rule.range ?? '', 'A top/bottom rule needs a rank of 1 or more.');
      if (rule.type === 'iconSet' && !/^[3-5][A-Za-z0-9]+$/.test(rule.set ?? '')) add('error', s.name, rule.range ?? '', 'An icon set is named like Excel\'s: "3Arrows", "3TrafficLights1", "5Rating".');
    }
    for (const v of s.validation ?? []) {
      if (!parseRange(v.range ?? '')) add('error', s.name, v.range ?? '', 'A validation rule needs a range.');
      if (!Array.isArray(v.list) || !v.list.length) add('error', s.name, v.range ?? '', 'A validation rule needs a non-empty list of allowed values.');
    }
    for (const ch of s.charts ?? []) {
      if (!['bar', 'line', 'pie'].includes(ch.type)) add('error', s.name, ch.id ?? '', `Unknown chart type "${ch.type}" (bar, line or pie).`);
      if (!parseRange(ch.range ?? '')) add('error', s.name, ch.id ?? '', 'A chart needs a data range like A1:C13.');
      if (![ch.x, ch.y, ch.w, ch.h].every((n) => typeof n === 'number' && n >= 0)) add('error', s.name, ch.id ?? '', 'A chart needs x, y, w and h in pixels.');
    }
    if (s.filter && !parseRange(s.filter.range ?? '')) add('error', s.name, s.filter.range ?? '', 'The filter needs a range whose first row is the header.');
    for (const [a, t] of Object.entries(s.notes ?? {})) {
      if (!parseA1(a)) add('warning', s.name, a, 'Note on something that is not an A1 address.');
      else if (typeof t !== 'string') add('error', s.name, a, 'A note is a string.');
    }
    for (const t of s.protected ?? []) if (!parseRange(t)) add('error', s.name, t, 'Not a protected range like A1:D1 or B2.');
    const merges = (s.merges ?? []).map((m) => ({ m, r: parseRange(m) }));
    for (const { m, r } of merges) if (!r || !m.includes(':')) add('error', s.name, m, 'Not a merge range like A1:C1.');
    for (let i = 0; i < merges.length; i++)
      for (let j = i + 1; j < merges.length; j++) {
        const a = merges[i].r;
        const b = merges[j].r;
        if (a && b && a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1) add('error', s.name, `${merges[i].m} ${merges[j].m}`, 'Merged ranges overlap.');
      }
  }
  let wb;
  try {
    wb = Workbook.fromJSON(data);
  } catch (e) {
    add('error', '', '', `The engine could not load the workbook: ${e.message}`);
    return problems;
  }
  for (const s of data.sheets) {
    for (const [a, input] of Object.entries(s.cells ?? {})) {
      if (typeof input !== 'string') {
        add('error', s.name, a, `Cell inputs are strings; found ${typeof input}. Write "${input}" in quotes.`);
        continue;
      }
      if (!input.startsWith('=')) continue;
      const v = wb.get(s.name, a);
      if (!(v instanceof CellError)) continue;
      const why =
        v.code === '#NAME?'
          ? 'The formula does not parse or uses a function open-spreadsheet does not have.'
          : v.code === '#REF!'
            ? 'A reference points to a deleted cell or a sheet that does not exist.'
            : v.code === '#CIRC!'
              ? 'A circular reference: the formula depends on itself.'
              : `The formula evaluates to ${v.code}.`;
      add(v.code === '#N/A' || v.code === '#DIV/0!' ? 'warning' : 'error', s.name, a, `${input} → ${v.code}. ${why}`);
    }
  }
  return problems;
}

export function checkWorkspace(root) {
  return listWorkbooks(root).map((w) => {
    try {
      return { id: w.id, problems: checkWorkbook(readWorkbook(root, w.id).data) };
    } catch (e) {
      return { id: w.id, problems: [{ severity: 'error', sheet: '', address: '', message: e.message }] };
    }
  });
}

// ---------- current selection ----------

export const currentFile = (root) => path.join(root, 'node_modules', '.open-spreadsheet', 'current.json');

export function readCurrent(root) {
  try {
    return JSON.parse(readFileSync(currentFile(root), 'utf8'));
  } catch {
    return null;
  }
}
