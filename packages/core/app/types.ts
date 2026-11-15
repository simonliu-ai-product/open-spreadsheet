/** How a cell looks. Kept apart from what it holds, so cells stay plain typed input. */
export interface CellStyle {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean;
  align?: 'left' | 'center' | 'right';
  /** Vertical alignment; unset draws in the middle, as the grid always has. */
  valign?: 'top' | 'middle' | 'bottom';
  /** An Excel number format code, e.g. "#,##0.00" or "yyyy-mm-dd". */
  fmt?: string;
  color?: string;
  bg?: string;
  wrap?: boolean;
  /** A checkbox: the cell holds TRUE or FALSE, drawn as a box a click toggles. */
  check?: boolean;
  /** Font size in points, as in Excel; 11 (the default) draws at 13px. */
  size?: number;
  /** Lines on the cell's edges: a CSS colour per side, drawn 1px (or 2px with `thick`). */
  border?: { top?: string; right?: string; bottom?: string; left?: string; thick?: boolean };
}

export interface SheetData {
  name: string;
  /** What was typed, by A1 address: "120", "=SUM(B2:B3)". */
  cells: Record<string, string>;
  /** Column widths in px by column letter. */
  cols?: Record<string, number>;
  /** Row heights in px by 1-based row number. */
  rows?: Record<string, number>;
  styles?: Record<string, CellStyle>;
  /** Merged ranges like "A1:C1". */
  merges?: string[];
  freeze?: { rows: number; cols: number };
  /** Tab colour. */
  color?: string;
  /** Rows (1-based numbers) and columns (letters) hidden by hand, apart from any filter. */
  hiddenRows?: number[];
  hiddenCols?: string[];
  /** Charts floating over the grid, drawn from a range of this sheet. */
  charts?: ChartSpec[];
  /** Notes on cells, by A1 address: comments for people, or requests for an agent (see the apply-notes skill). */
  notes?: Record<string, string>;
  /** Conditional formatting rules, applied in order; later rules win. */
  conditional?: ConditionalRule[];
  /** Cells that only accept values from a list, shown as a dropdown. */
  validation?: Validation[];
  /** An autofilter: its range (header row first) and, per column letter, the shown texts it hides. */
  filter?: { range: string; hidden?: Record<string, string[]> };
}

/** How a conditional rule highlights a cell; Excel's light red fill with dark red text by default. */
export type HighlightStyle = Pick<CellStyle, 'b' | 'color' | 'bg'>;

export type ConditionalRule =
  | { range: string; type: 'gt' | 'lt' | 'ge' | 'le' | 'eq' | 'ne'; value: number | string; style?: HighlightStyle }
  | { range: string; type: 'between'; value: number; value2: number; style?: HighlightStyle }
  | { range: string; type: 'contains'; text: string; style?: HighlightStyle }
  | { range: string; type: 'duplicate'; style?: HighlightStyle }
  /** Two or three colours, from the smallest number to the largest. */
  | { range: string; type: 'colorScale'; colors: string[] }
  | { range: string; type: 'dataBar'; color?: string };

export interface Validation {
  range: string;
  list: string[];
}

export interface WorkbookData {
  title: string;
  sheets: SheetData[];
  /** Defined names: { "Sales": "Orders!$F$2:$F$200" }, usable in formulas as =SUM(Sales). */
  names?: Record<string, string>;
}

export interface WorkbookSummary {
  id: string;
  title: string;
  sheets: number;
  updatedAt: string;
}

export type ChartType = 'bar' | 'line' | 'pie';

export interface ChartSpec {
  id: string;
  type: ChartType;
  /** The data: first column the categories, first row the series names (when they are text). */
  range: string;
  title?: string;
  /** Position and size in px at 100% zoom, from the sheet's top-left corner. */
  x: number;
  y: number;
  w: number;
  h: number;
}
