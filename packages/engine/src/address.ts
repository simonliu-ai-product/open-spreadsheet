/** A1-style cell addresses. Rows and columns are zero-based internally. */

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export function columnName(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function columnIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export interface CellAddress {
  row: number;
  col: number;
}

export function parseA1(a1: string): CellAddress | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(a1);
  if (!m) return null;
  const col = columnIndex(m[1]!);
  const row = Number(m[2]) - 1;
  if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) return null;
  return { row, col };
}

export function formatA1(row: number, col: number, rowAbs = false, colAbs = false): string {
  return `${colAbs ? '$' : ''}${columnName(col)}${rowAbs ? '$' : ''}${row + 1}`;
}

export const cellKey = (row: number, col: number) => `${row},${col}`;
