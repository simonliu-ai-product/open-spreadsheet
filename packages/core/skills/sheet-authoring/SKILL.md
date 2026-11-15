---
name: sheet-authoring
description: Technical reference for writing or editing open-spreadsheet workbooks — the workbook.json format, how cells, formulas, number formats, styles, merges, column widths and frozen panes are stored, which Excel functions work, and how to verify a change. Consult this whenever you are about to create or modify any file under `sheets/<id>/`, add formulas, format numbers, or lay out a table.
---

# Writing open-spreadsheet workbooks

Each workbook is `sheets/<id>/workbook.json`. The grid editor at `pnpm dev` reloads it the moment it changes, so the person sees your edit live.

## The format

```json
{
  "title": "Q1 sales",
  "sheets": [
    {
      "name": "Orders",
      "cells": { "A1": "Date", "B1": "Rep", "C1": "Amount", "A2": "2026-01-05", "B2": "May", "C2": "32000", "C10": "=SUM(C2:C9)" },
      "styles": { "A1": { "b": true, "bg": "#e8f0e9" }, "C2": { "fmt": "#,##0" }, "C10": { "b": true, "fmt": "#,##0" } },
      "cols": { "A": 110 },
      "freeze": { "rows": 1, "cols": 0 }
    }
  ]
}
```

| Key | What it holds |
| --- | --- |
| `cells` | A1 address → the text a person would type. **Always a string.** |
| `styles` | A1 address → `{ b, i, u, s, align, valign, wrap, check, size, color, bg, fmt, border }`; `size` is points (11 is the default; rows without a set height grow to fit); `border` is `{ top, right, bottom, left }` colours, plus `thick: true` |
| `cols` | Column letter → width in px (default 100) |
| `rows` | 1-based row number → height in px (default 24) |
| `merges` | Ranges like `"A1:D1"`; only the top-left cell keeps a value |
| `freeze` | `{ rows, cols }` frozen at the top and left |
| `hiddenRows` / `hiddenCols` | Rows (1-based numbers) and columns (letters) hidden by hand |
| `names` (top level, beside `sheets`) | Defined names: `{ "Sales": "Orders!$F$2:$F$200" }`, used in formulas as `=SUM(Sales)`; each points at a cell or range with its sheet |
| `protected` | Ranges a person locked, `["A1:D1", "F2"]`. The editor and MCP refuse to change them; never edit or remove them yourself — ask the person |
| `charts` | Charts over the grid: `[{ "id": "sales", "type": "bar", "range": "A1:C13", "title": "Sales by month", "x": 520, "y": 0, "w": 440, "h": 270 }]` — `type` is `bar`, `line` or `pie`; the range's first column holds the categories and its first row the series names; `x`/`y`/`w`/`h` are px at 100% zoom from the sheet's top-left |
| `color` | The sheet tab's colour |
| `notes` | A1 address → a note left on the cell (see the `apply-notes` skill) |
| `conditional` | Conditional formats, later ones win: `{ "range": "D2:D50", "type": "gt", "value": 10000, "style": { "bg": "#ffc7ce", "color": "#9c0006" } }`; types `gt lt ge le eq ne` (value), `between` (value, value2), `contains` (text), `duplicate`, `colorScale` (`colors`: 2 or 3 hex, low → high), `dataBar` (`color`), `top` (`rank`, optional `percent`, `bottom`), `average` (optional `below`, `equal`), `iconSet` (`set` as Excel names it — `3Arrows`, `3TrafficLights1`, `4Rating`, `5Arrows`…; optional `thresholds` in percent, `reverse`). All of them go to and from .xlsx |
| `validation` | Dropdown lists: `[{ "range": "C2:C100", "list": ["Open", "In progress", "Done"] }]`; other input is refused in the editor |
| `filter` | An autofilter: `{ "range": "A1:F200" }` (header row first); `hidden` maps a column letter to the shown values it hides |

## Cells

- Numbers: `"1200"`, `"-3.5"`, `"1,200"`, `"12%"`, `"$1,200"`. Dates: `"2026-01-05"` (shown as a date), times: `"14:30"`.
- Text: anything else. To keep `"00123"` as text, write `"'00123"`.
- Formulas start with `=` and follow Excel exactly: `"=SUM(B2:B9)"`, `"=IF(C2>0,\"up\",\"down\")"`, `"=VLOOKUP(A2,Prices!A:B,2,FALSE)"`, `"='Raw data'!B3*2"`.
- Write totals, ratios and lookups as formulas, never as computed numbers.

## Number formats (`fmt`)

| Want | `fmt` |
| --- | --- |
| Thousands | `#,##0` / `#,##0.00` |
| NT dollars | `"NT$"#,##0` |
| Percent | `0%` / `0.0%` |
| Negative in red | `#,##0;[Red]-#,##0` |
| Date | `yyyy-mm-dd`, `yyyy"年"m"月"d"日"` |
| Time | `hh:mm` |

## Functions

Math and trigonometry (SUM, SUMIF(S), SUMPRODUCT, SUBTOTAL, ROUND/ROUNDUP/ROUNDDOWN/MROUND, CEILING, FLOOR, INT, MOD, ABS, POWER, SQRT, EXP, LN, LOG, FACT, COMBIN, GCD, LCM, SIN/COS/TAN…), statistics (AVERAGE(IF/IFS), COUNT, COUNTA, COUNTIF(S), COUNTBLANK, MAX/MIN(IFS), MEDIAN, MODE, LARGE, SMALL, STDEV(.S/.P), VAR, PERCENTILE, QUARTILE, RANK, CORREL, SLOPE, INTERCEPT, FORECAST, GEOMEAN, NORM.DIST), logic (IF, IFS, IFERROR, IFNA, AND, OR, NOT, XOR, SWITCH), lookup (VLOOKUP, HLOOKUP, XLOOKUP, XMATCH, LOOKUP, INDEX, MATCH, CHOOSE, ADDRESS, ROW, COLUMN), text (LEFT, RIGHT, MID, LEN, FIND, SEARCH, SUBSTITUTE, REPLACE, UPPER, LOWER, PROPER, TRIM, CLEAN, CONCAT, TEXTJOIN, TEXTBEFORE, TEXTAFTER, TEXT, FIXED, DOLLAR, VALUE, NUMBERVALUE, REPT, CHAR, CODE), dates (DATE, TIME, TODAY, NOW, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, WEEKNUM, ISOWEEKNUM, DAYS, DAYS360, DATEDIF, YEARFRAC, EDATE, EOMONTH, NETWORKDAYS, WORKDAY, DATEVALUE), finance (PMT, IPMT, PPMT, PV, FV, NPV, XNPV, IRR, XIRR, RATE, NPER, SLN, SYD, DDB, EFFECT, NOMINAL), database (DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, DGET), information (ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISERR, ISNA, ISEVEN, ISODD, TYPE, N, T), engineering (BIN2DEC, DEC2HEX and the other base conversions, BITAND/OR/XOR, COMPLEX and the IM functions, CONVERT), distributions (NORM.DIST/INV, T.DIST/INV, CHISQ.DIST, F.DIST, BINOM.DIST, POISSON.DIST, EXPON.DIST, GAMMA.DIST, BETA.DIST, CONFIDENCE). Anything else shows `#NAME?` — rewrite it with these.

## Layout habits

- Row 1 is a header: bold, a light `bg`, and `"freeze": { "rows": 1, "cols": 0 }`.
- Give label columns room (`"cols": { "A": 140 }`).
- Money gets `#,##0` or `"NT$"#,##0`; shares get `0.0%`.

## Verify

1. Run `open-spreadsheet check` (or the MCP `check` tool). It lists formulas that show `#NAME?`, `#REF!`, `#CIRC!`, `#VALUE!` and malformed entries.
2. Read the result back with the MCP `read_range` tool, or open the workbook in the viewer.
