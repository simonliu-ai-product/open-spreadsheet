/**
 * CSV and Excel (.xlsx) in and out, with no library: a .xlsx is a zip of XML
 * files. Writing uses stored (uncompressed) entries, which every reader
 * accepts; reading inflates with the browser's DecompressionStream.
 */
import { CellError, columnIndex, columnName, formatA1, formatValue, impliedFormat, offsetFormula, parseA1, type Scalar, Workbook } from '@open-spreadsheet/engine';
import { chartShape } from './charts.ts';
import { BLANK, columnValues, filteredRows, parseRange } from './features.ts';
import type { CellStyle, ChartSpec, ChartType, ConditionalRule, HighlightStyle, SheetData, WorkbookData } from './types.ts';

// ---------- CSV ----------

export function parseCsv(text: string): string[][] {
  const src = text.replace(/^﻿/, '');
  const firstLine = src.slice(0, src.indexOf('\n') === -1 ? undefined : src.indexOf('\n'));
  const count = (ch: string) => firstLine.split(ch).length - 1;
  const delim = count('\t') > count(',') && count('\t') >= count(';') ? '\t' : count(';') > count(',') ? ';' : ',';
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i]!;
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          field += '"';
          i++;
        } else quoted = false;
      } else field += ch;
    } else if (ch === '"' && field === '') {
      quoted = true;
    } else if (ch === delim) {
      row.push(field);
      field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && src[i + 1] === '\n') i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

export function toCsv(rows: string[][]): string {
  const q = (v: string) => (/[",\n\r]/.test(v) ? `"${v.replaceAll('"', '""')}"` : v);
  // A byte-order mark so Excel opens UTF-8 (Chinese text) correctly.
  return `﻿${rows.map((r) => r.map(q).join(',')).join('\r\n')}\r\n`;
}

export function sheetFromRows(name: string, rows: string[][]): SheetData {
  const cells: Record<string, string> = {};
  rows.forEach((r, ri) =>
    r.forEach((v, ci) => {
      if (v !== '') cells[formatA1(ri, ci)] = v;
    }),
  );
  return { name, cells };
}

// ---------- zip ----------

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

export function zip(files: { name: string; data: string | Uint8Array }[]): Blob {
  const enc = new TextEncoder();
  const parts: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const f of files) {
    const name = enc.encode(f.name);
    const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
    const crc = crc32(data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0x0800, true); // UTF-8 names
    local.setUint16(8, 0, true); // stored
    local.setUint32(14, crc, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, data.length, true);
    local.setUint16(26, name.length, true);
    parts.push(new Uint8Array(local.buffer), name, data);
    const cd = new DataView(new ArrayBuffer(46));
    cd.setUint32(0, 0x02014b50, true);
    cd.setUint16(4, 20, true);
    cd.setUint16(6, 20, true);
    cd.setUint16(8, 0x0800, true);
    cd.setUint32(16, crc, true);
    cd.setUint32(20, data.length, true);
    cd.setUint32(24, data.length, true);
    cd.setUint16(28, name.length, true);
    cd.setUint32(42, offset, true);
    central.push(new Uint8Array(cd.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const cdSize = central.reduce((n, p) => n + p.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.length, true);
  end.setUint16(10, files.length, true);
  end.setUint32(12, cdSize, true);
  end.setUint32(16, offset, true);
  return new Blob([...parts, ...central, new Uint8Array(end.buffer)] as BlobPart[], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

export async function unzip(buf: ArrayBuffer): Promise<Map<string, string>> {
  const view = new DataView(buf);
  let eocd = -1;
  for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('not a zip file');
  const count = view.getUint16(eocd + 10, true);
  let p = view.getUint32(eocd + 16, true);
  const dec = new TextDecoder();
  const out = new Map<string, string>();
  for (let i = 0; i < count; i++) {
    const method = view.getUint16(p + 10, true);
    const size = view.getUint32(p + 20, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    const local = view.getUint32(p + 42, true);
    const name = dec.decode(new Uint8Array(buf, p + 46, nameLen));
    p += 46 + nameLen + extraLen + commentLen;
    if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    const raw = new Uint8Array(buf, start, size);
    let data: Uint8Array;
    if (method === 0) data = raw;
    else if (method === 8) {
      const stream = new Blob([raw as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      data = new Uint8Array(await new Response(stream).arrayBuffer());
    } else continue;
    out.set(name, dec.decode(data));
  }
  return out;
}

// ---------- xlsx out ----------

const xml = (s: string) => s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
// Functions Excel stores with the _xlfn. prefix (added after Excel 2007).
const XLFN = /\b(XLOOKUP|XMATCH|IFS|SWITCH|CONCAT|TEXTJOIN|XOR|DAYS|STDEV\.S|STDEV\.P|VAR\.S|VAR\.P|IFNA|MAXIFS|MINIFS|FILTER|SORT|UNIQUE|SEQUENCE|LET|LAMBDA|TEXTSPLIT|TEXTBEFORE|TEXTAFTER|RANK\.EQ|PERCENTILE\.INC|QUARTILE\.INC|MODE\.SNGL|NORM\.DIST|ISOWEEKNUM|CEILING\.MATH|FLOOR\.MATH)\(/g;

function withXlfn(formula: string): string {
  // Only outside string literals.
  return formula.replace(/("[^"]*")|([^"]+)/g, (m, str, code) => (str ? m : code.replace(XLFN, '_xlfn.$1(')));
}

const argb = (hex: string) => `FF${hex.replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6)}`;

export function exportXlsx(data: WorkbookData, wb: Workbook): Blob {
  // Styles: each distinct (font, fill, numFmt, alignment) is one cellXfs entry.
  const numFmts: string[] = [];
  const fonts: string[] = ['<font><sz val="11"/><name val="Calibri"/></font>'];
  const fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const xfIndex = new Map<string, number>();
  const BUILTIN: Record<string, number> = { General: 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '@': 49 };
  const styleId = (s: CellStyle | undefined): number => {
    if (!s || Object.keys(s).length === 0) return 0;
    const key = JSON.stringify(s);
    const hit = xfIndex.get(key);
    if (hit !== undefined) return hit;
    let numFmtId = 0;
    if (s.fmt) {
      if (BUILTIN[s.fmt] !== undefined) numFmtId = BUILTIN[s.fmt]!;
      else {
        let i = numFmts.indexOf(s.fmt);
        if (i === -1) i = numFmts.push(s.fmt) - 1;
        numFmtId = 164 + i;
      }
    }
    let fontId = 0;
    if (s.b || s.i || s.u || s.s || s.color || s.size) {
      fonts.push(
        `<font>${s.b ? '<b/>' : ''}${s.i ? '<i/>' : ''}${s.s ? '<strike/>' : ''}${s.u ? '<u/>' : ''}<sz val="${s.size ?? 11}"/>${s.color ? `<color rgb="${argb(s.color)}"/>` : ''}<name val="Calibri"/></font>`,
      );
      fontId = fonts.length - 1;
    }
    let fillId = 0;
    if (s.bg) {
      fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${argb(s.bg)}"/><bgColor indexed="64"/></patternFill></fill>`);
      fillId = fills.length - 1;
    }
    let borderId = 0;
    if (s.border) {
      const kind = s.border.thick ? 'medium' : 'thin';
      const side = (name: string, c?: string) => (c ? `<${name} style="${kind}"><color rgb="${argb(c)}"/></${name}>` : `<${name}/>`);
      borders.push(`<border>${side('left', s.border.left)}${side('right', s.border.right)}${side('top', s.border.top)}${side('bottom', s.border.bottom)}<diagonal/></border>`);
      borderId = borders.length - 1;
    }
    const vertical = s.valign === 'middle' ? 'center' : s.valign;
    const align = s.align || s.wrap || vertical ? `<alignment${s.align ? ` horizontal="${s.align}"` : ''}${vertical ? ` vertical="${vertical}"` : ''}${s.wrap ? ' wrapText="1"' : ''}/>` : '';
    xfs.push(
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''}${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${align ? ' applyAlignment="1">' + align + '</xf>' : '/>'}`,
    );
    xfIndex.set(key, xfs.length - 1);
    return xfs.length - 1;
  };

  // Differential formats: what a conditional rule paints over a cell.
  const dxfs: string[] = [];
  const dxfId = (s: HighlightStyle | undefined) => {
    const st = s ?? { bg: '#ffc7ce', color: '#9c0006' };
    dxfs.push(
      `<dxf>${st.b || st.color ? `<font>${st.b ? '<b/>' : ''}${st.color ? `<color rgb="${argb(st.color)}"/>` : ''}</font>` : ''}${st.bg ? `<fill><patternFill><bgColor rgb="${argb(st.bg)}"/></patternFill></fill>` : ''}</dxf>`,
    );
    return dxfs.length - 1;
  };
  let priority = 0;
  const OPS: Record<string, string> = { gt: 'greaterThan', lt: 'lessThan', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual', eq: 'equal', ne: 'notEqual' };
  const literal = (v: number | string) => (typeof v === 'number' || (v !== '' && Number.isFinite(Number(v))) ? String(v) : `"${String(v).replaceAll('"', '""')}"`);
  const ruleXml = (r: ConditionalRule): string => {
    priority++;
    const top = parseRange(r.range);
    const first = top ? formatA1(top.r0, top.c0) : 'A1';
    switch (r.type) {
      case 'between':
        return `<cfRule type="cellIs" dxfId="${dxfId(r.style)}" priority="${priority}" operator="between"><formula>${r.value}</formula><formula>${r.value2}</formula></cfRule>`;
      case 'contains':
        return `<cfRule type="containsText" dxfId="${dxfId(r.style)}" priority="${priority}" operator="containsText" text="${xml(r.text)}"><formula>${xml(`NOT(ISERROR(SEARCH(${literal(r.text)},${first})))`)}</formula></cfRule>`;
      case 'duplicate':
        return `<cfRule type="duplicateValues" dxfId="${dxfId(r.style)}" priority="${priority}"/>`;
      case 'colorScale': {
        const cs = r.colors.length >= 3 ? r.colors.slice(0, 3) : r.colors.slice(0, 2);
        const cfvo = cs.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>';
        return `<cfRule type="colorScale" priority="${priority}"><colorScale>${cfvo}${cs.map((c) => `<color rgb="${argb(c)}"/>`).join('')}</colorScale></cfRule>`;
      }
      case 'dataBar':
        return `<cfRule type="dataBar" priority="${priority}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${argb(r.color ?? '#638ec6')}"/></dataBar></cfRule>`;
      default:
        return `<cfRule type="cellIs" dxfId="${dxfId(r.style)}" priority="${priority}" operator="${OPS[r.type]}"><formula>${xml(literal(r.value))}</formula></cfRule>`;
    }
  };

  // Charts: one drawing per sheet that has any, one chart part per chart.
  const extraFiles: { name: string; data: string }[] = [];
  const extraTypes: string[] = [];
  let chartCount = 0;
  const sheetRefPrefix = (name: string) => (/^[A-Za-z_][\w.]*$/.test(name) ? name : `'${name.replaceAll("'", "''")}'`);

  function drawingFor(sheet: SheetData, sheetIndex: number): string {
    const colPx = (c: number) => sheet.cols?.[columnName(c)] ?? 100;
    const rowPx = (r: number) => sheet.rows?.[String(r + 1)] ?? 24;
    const cellAt = (px: number, size: (i: number) => number) => {
      let i = 0;
      let left = px;
      while (left >= size(i) && i < 16384) left -= size(i++);
      return { i, off: Math.round(left * 9525) };
    };
    const anchors: string[] = [];
    const rels: string[] = [];
    (sheet.charts ?? []).forEach((ch, k) => {
      chartCount++;
      const n = chartCount;
      const from = { c: cellAt(ch.x, colPx), r: cellAt(ch.y, rowPx) };
      const to = { c: cellAt(ch.x + ch.w, colPx), r: cellAt(ch.y + ch.h, rowPx) };
      anchors.push(
        `<xdr:twoCellAnchor><xdr:from><xdr:col>${from.c.i}</xdr:col><xdr:colOff>${from.c.off}</xdr:colOff><xdr:row>${from.r.i}</xdr:row><xdr:rowOff>${from.r.off}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${to.c.i}</xdr:col><xdr:colOff>${to.c.off}</xdr:colOff><xdr:row>${to.r.i}</xdr:row><xdr:rowOff>${to.r.off}</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${k + 2}" name="${xml(ch.title ?? `Chart ${k + 1}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${k + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`,
      );
      rels.push(`<Relationship Id="rId${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${n}.xml"/>`);
      extraFiles.push({ name: `xl/charts/chart${n}.xml`, data: chartXml(sheet, ch) });
      extraTypes.push(`<Override PartName="/xl/charts/chart${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
    });
    const d = sheetIndex + 1;
    extraFiles.push({
      name: `xl/drawings/drawing${d}.xml`,
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors.join('')}</xdr:wsDr>`,
    });
    extraFiles.push({
      name: `xl/drawings/_rels/drawing${d}.xml.rels`,
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
    });
    extraFiles.push({
      name: `xl/worksheets/_rels/sheet${d}.xml.rels`,
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${d}.xml"/></Relationships>`,
    });
    extraTypes.push(`<Override PartName="/xl/drawings/drawing${d}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
    return '<drawing r:id="rIdDrawing1"/>';
  }

  /** A chart part whose series point at the sheet's cells, so Excel keeps it live. */
  function chartXml(sheet: SheetData, ch: ChartSpec): string {
    const r = parseRange(ch.range) ?? { r0: 0, c0: 0, r1: 0, c1: 0 };
    const cells: Scalar[][] = [];
    for (let row = r.r0; row <= r.r1; row++) {
      const line: Scalar[] = [];
      for (let col = r.c0; col <= r.c1; col++) line.push(wb.value(sheet.name, row, col));
      cells.push(line);
    }
    const { headerRow, labelCol } = chartShape(cells);
    const pre = `${sheetRefPrefix(sheet.name)}!`;
    const abs = (row: number, col: number) => `$${columnName(col)}$${row + 1}`;
    const d0 = r.r0 + (headerRow ? 1 : 0);
    const cat = labelCol ? `<c:cat><c:strRef><c:f>${xml(`${pre}${abs(d0, r.c0)}:${abs(r.r1, r.c0)}`)}</c:f></c:strRef></c:cat>` : '';
    const series: string[] = [];
    for (let col = r.c0 + (labelCol ? 1 : 0), i = 0; col <= r.c1; col++, i++) {
      if (ch.type === 'pie' && i > 0) break;
      const tx = headerRow ? `<c:tx><c:strRef><c:f>${xml(`${pre}${abs(r.r0, col)}`)}</c:f></c:strRef></c:tx>` : '';
      const marker = ch.type === 'line' ? '<c:marker><c:symbol val="circle"/></c:marker>' : '';
      const smooth = ch.type === 'line' ? '<c:smooth val="0"/>' : '';
      series.push(`<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${marker}${cat}<c:val><c:numRef><c:f>${xml(`${pre}${abs(d0, col)}:${abs(r.r1, col)}`)}</c:f></c:numRef></c:val>${smooth}</c:ser>`);
    }
    const axes =
      '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="10"/></c:valAx>';
    const plot =
      ch.type === 'bar'
        ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series.join('')}<c:axId val="10"/><c:axId val="20"/></c:barChart>${axes}`
        : ch.type === 'line'
          ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series.join('')}<c:marker val="1"/><c:axId val="10"/><c:axId val="20"/></c:lineChart>${axes}`
          : `<c:pieChart><c:varyColors val="1"/>${series.join('')}<c:firstSliceAng val="0"/></c:pieChart>`;
    const title = ch.title
      ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${xml(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
      : '<c:autoTitleDeleted val="1"/>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
  }

  const sheetsXml = data.sheets.map((sheet, sheetIndex) => {
    const drawing = sheet.charts?.length ? drawingFor(sheet, sheetIndex) : '';
    const shown = (r: number, c: number) => {
      const input = sheet.cells[formatA1(r, c)] ?? '';
      if (!input) return '';
      return formatValue(wb.value(sheet.name, r, c), sheet.styles?.[formatA1(r, c)]?.fmt ?? (input.startsWith('=') ? undefined : impliedFormat(input))).text;
    };
    const used = wb.used(sheet.name);
    const hiddenRows = filteredRows(sheet, shown, used);
    for (const n of sheet.hiddenRows ?? []) hiddenRows.add(n - 1);
    const rows = new Map<number, string[]>();
    const addresses = new Set([...Object.keys(sheet.cells), ...Object.keys(sheet.styles ?? {})]);
    const sorted = [...addresses]
      .map((a) => ({ a, at: parseA1(a)! }))
      .filter((c) => c.at)
      .sort((x, y) => x.at.row - y.at.row || x.at.col - y.at.col);
    for (const { a, at } of sorted) {
      const input = sheet.cells[a] ?? '';
      // A typed date or percentage carries its look into Excel as a number format.
      const implied = input && !input.startsWith('=') ? impliedFormat(input) : undefined;
      const own = sheet.styles?.[a];
      const s = styleId(implied && !own?.fmt ? { ...own, fmt: implied } : own);
      const sAttr = s ? ` s="${s}"` : '';
      const value = wb.value(sheet.name, at.row, at.col);
      let c: string;
      if (input === '') c = `<c r="${a}"${sAttr}/>`;
      else if (input.startsWith('=')) {
        const f = `<f>${xml(withXlfn(input.slice(1)))}</f>`;
        if (typeof value === 'number') c = `<c r="${a}"${sAttr}>${f}<v>${value}</v></c>`;
        else if (typeof value === 'boolean') c = `<c r="${a}"${sAttr} t="b">${f}<v>${value ? 1 : 0}</v></c>`;
        else if (value instanceof CellError) c = `<c r="${a}"${sAttr} t="e">${f}<v>${xml(value.code)}</v></c>`;
        else c = `<c r="${a}"${sAttr} t="str">${f}<v>${xml(String(value ?? ''))}</v></c>`;
      } else if (typeof value === 'number') c = `<c r="${a}"${sAttr}><v>${value}</v></c>`;
      else if (typeof value === 'boolean') c = `<c r="${a}"${sAttr} t="b"><v>${value ? 1 : 0}</v></c>`;
      else c = `<c r="${a}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${xml(String(value ?? ''))}</t></is></c>`;
      const list = rows.get(at.row) ?? [];
      list.push(c);
      rows.set(at.row, list);
    }
    for (const r of Object.keys(sheet.rows ?? {})) if (!rows.has(Number(r) - 1)) rows.set(Number(r) - 1, []);
    for (const r of hiddenRows) if (!rows.has(r)) rows.set(r, []);
    const rowXml = [...rows.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([r, cs]) => {
        const h = sheet.rows?.[String(r + 1)];
        return `<row r="${r + 1}"${h ? ` ht="${(h * 0.75).toFixed(2)}" customHeight="1"` : ''}${hiddenRows.has(r) ? ' hidden="1"' : ''}>${cs.join('')}</row>`;
      })
      .join('');
    const hiddenCols = new Set((sheet.hiddenCols ?? []).map((l) => columnIndex(l)));
    const colIdx = new Set([...Object.keys(sheet.cols ?? {}).map((l) => columnIndex(l)), ...hiddenCols]);
    const cols = [...colIdx]
      .sort((a, b) => a - b)
      .map((c) => {
        const px = sheet.cols?.[columnName(c)] ?? 100;
        return `<col min="${c + 1}" max="${c + 1}" width="${Math.max(0, (px - 5) / 7).toFixed(2)}" customWidth="1"${hiddenCols.has(c) ? ' hidden="1"' : ''}/>`;
      })
      .join('');
    const f = sheet.freeze;
    const pane =
      f && (f.rows || f.cols)
        ? `<pane${f.cols ? ` xSplit="${f.cols}"` : ''}${f.rows ? ` ySplit="${f.rows}"` : ''} topLeftCell="${formatA1(f.rows, f.cols)}" activePane="${f.rows && f.cols ? 'bottomRight' : f.rows ? 'bottomLeft' : 'topRight'}" state="frozen"/>`
        : '';
    let autoFilter = '';
    const fr = sheet.filter ? parseRange(sheet.filter.range) : null;
    if (sheet.filter && fr) {
      const columns = Object.entries(sheet.filter.hidden ?? {})
        .map(([letter, hidden]) => {
          const col = columnIndex(letter);
          if (col < fr.c0 || col > fr.c1 || !hidden.length) return '';
          const visible = columnValues(sheet, col, shown, used).filter((v) => !hidden.includes(v));
          const blank = visible.includes(BLANK);
          return `<filterColumn colId="${col - fr.c0}"><filters${blank ? ' blank="1"' : ''}>${visible
            .filter((v) => v !== BLANK)
            .map((v) => `<filter val="${xml(v)}"/>`)
            .join('')}</filters></filterColumn>`;
        })
        .join('');
      autoFilter = `<autoFilter ref="${sheet.filter.range}">${columns}</autoFilter>`;
    }
    const conditional = (sheet.conditional ?? []).map((r) => `<conditionalFormatting sqref="${r.range}">${ruleXml(r)}</conditionalFormatting>`).join('');
    const validations = sheet.validation?.length
      ? `<dataValidations count="${sheet.validation.length}">${sheet.validation
          .map((v) => `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${v.range}"><formula1>${xml(`"${v.list.join(',')}"`)}</formula1></dataValidation>`)
          .join('')}</dataValidations>`
      : '';
    const merges = sheet.merges?.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${sheet.color ? `<sheetPr><tabColor rgb="${argb(sheet.color)}"/></sheetPr>` : ''}<sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${rowXml}</sheetData>${autoFilter}${merges}${conditional}${validations}${drawing}</worksheet>`;
  });

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${xml(f)}"/>`).join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>${dxfs.length ? `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>` : ''}</styleSheet>`;

  const n = data.sheets.length;
  return zip([
    {
      name: '[Content_Types].xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${data.sheets
        .map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
        .join('')}${extraTypes.join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    },
    {
      name: '_rels/.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    },
    {
      name: 'xl/workbook.xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${data.sheets
        .map((s, i) => `<sheet name="${xml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
        .join('')}</sheets>${
        data.names && Object.keys(data.names).length
          ? `<definedNames>${Object.entries(data.names)
              .map(([n, ref]) => `<definedName name="${xml(n)}">${xml(ref)}</definedName>`)
              .join('')}</definedNames>`
          : ''
      }<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${data.sheets
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
        .join('')}<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    },
    { name: 'xl/styles.xml', data: styles },
    ...sheetsXml.map((x, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: x })),
    ...extraFiles,
  ]);
}

// ---------- xlsx in ----------

// Excel's built-in number formats that a file refers to by id only.
const BUILTIN_FMT: Record<number, string> = {
  1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'yyyy/m/d', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy/m/d h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
};

const tags = (node: Document | Element, name: string) => [...node.getElementsByTagNameNS('*', name)];
const first = (node: Document | Element, name: string) => node.getElementsByTagNameNS('*', name)[0] ?? null;
const rgbOf = (e: Element | null) => {
  const v = e?.getAttribute('rgb');
  return v && v.length >= 6 ? `#${v.slice(-6).toLowerCase()}` : undefined;
};

export async function importXlsx(buf: ArrayBuffer): Promise<WorkbookData> {
  const files = await unzip(buf);
  const parse = (name: string) => {
    const text = files.get(name);
    return text ? new DOMParser().parseFromString(text, 'application/xml') : null;
  };
  const book = parse('xl/workbook.xml');
  if (!book) throw new Error('no workbook.xml');
  const rels = parse('xl/_rels/workbook.xml.rels');
  const target = new Map(tags(rels!, 'Relationship').map((r) => [r.getAttribute('Id'), r.getAttribute('Target') ?? '']));

  const strings = tags(parse('xl/sharedStrings.xml') ?? new Document(), 'si').map((si) =>
    tags(si, 't')
      .filter((t) => t.parentElement?.localName !== 'rPh')
      .map((t) => t.textContent ?? '')
      .join(''),
  );

  // Styles: cellXfs index -> our CellStyle.
  const st = parse('xl/styles.xml');
  const custom = new Map(tags(st ?? new Document(), 'numFmt').map((n) => [Number(n.getAttribute('numFmtId')), n.getAttribute('formatCode') ?? '']));
  const fonts = st ? tags(first(st, 'fonts') ?? st, 'font') : [];
  // The workbook's own default size: a cell in it is the grid's normal size.
  const baseSize = Number((fonts[0] && first(fonts[0], 'sz')?.getAttribute('val')) ?? 11) || 11;
  const fills = st ? tags(first(st, 'fills') ?? st, 'fill') : [];
  const borderList = st ? tags(first(st, 'borders') ?? st, 'border') : [];
  const xfStyles: (CellStyle | undefined)[] = (st && first(st, 'cellXfs') ? tags(first(st, 'cellXfs')!, 'xf') : []).map((xf) => {
    const s: CellStyle = {};
    const id = Number(xf.getAttribute('numFmtId') ?? 0);
    const fmt = custom.get(id) ?? BUILTIN_FMT[id];
    if (fmt && fmt !== 'General') s.fmt = fmt;
    const font = fonts[Number(xf.getAttribute('fontId') ?? 0)];
    if (font && Number(xf.getAttribute('fontId')) > 0) {
      if (first(font, 'b') && first(font, 'b')!.getAttribute('val') !== '0') s.b = true;
      if (first(font, 'i') && first(font, 'i')!.getAttribute('val') !== '0') s.i = true;
      if (first(font, 'u')) s.u = true;
      if (first(font, 'strike')) s.s = true;
      const c = rgbOf(first(font, 'color'));
      if (c && c !== '#000000') s.color = c;
      const sz = Number(first(font, 'sz')?.getAttribute('val'));
      if (sz && sz !== baseSize) s.size = sz;
    }
    const fill = fills[Number(xf.getAttribute('fillId') ?? 0)];
    const pf = fill ? first(fill, 'patternFill') : null;
    if (pf?.getAttribute('patternType') === 'solid') {
      const c = rgbOf(first(pf, 'fgColor'));
      if (c) s.bg = c;
    }
    const bd = borderList[Number(xf.getAttribute('borderId') ?? 0)];
    if (bd && Number(xf.getAttribute('borderId')) > 0) {
      const border: NonNullable<CellStyle['border']> = {};
      for (const side of ['top', 'right', 'bottom', 'left'] as const) {
        const e = first(bd, side);
        const kind = e?.getAttribute('style');
        if (!kind || kind === 'none') continue;
        border[side] = rgbOf(first(e!, 'color')) ?? '#000000';
        if (kind === 'medium' || kind === 'thick') border.thick = true;
      }
      if (Object.keys(border).some((k) => k !== 'thick')) s.border = border;
    }
    const al = first(xf, 'alignment');
    const h = al?.getAttribute('horizontal');
    if (h === 'left' || h === 'center' || h === 'right') s.align = h;
    if (al?.getAttribute('wrapText') === '1') s.wrap = true;
    const v = al?.getAttribute('vertical');
    if (v === 'top' || v === 'bottom') s.valign = v;
    else if (v === 'center') s.valign = 'middle';
    return Object.keys(s).length ? s : undefined;
  });

  const dxfStyles: HighlightStyle[] = (st && first(st, 'dxfs') ? tags(first(st, 'dxfs')!, 'dxf') : []).map((d) => {
    const out: HighlightStyle = {};
    const font = first(d, 'font');
    if (font && first(font, 'b')) out.b = true;
    const color = rgbOf(font ? first(font, 'color') : null);
    if (color) out.color = color;
    const fill = first(d, 'fill');
    const bg = fill ? (rgbOf(first(fill, 'bgColor')) ?? rgbOf(first(fill, 'fgColor'))) : undefined;
    if (bg) out.bg = bg;
    return out;
  });

  const sheets: SheetData[] = [];
  for (const sh of tags(book, 'sheet')) {
    const name = sh.getAttribute('name') ?? `Sheet${sheets.length + 1}`;
    const rid = sh.getAttribute('r:id') ?? sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    let path = target.get(rid) ?? '';
    path = path.startsWith('/') ? path.slice(1) : `xl/${path}`;
    const doc = parse(path);
    if (!doc) continue;
    const sheet: SheetData = { name, cells: {} };
    const styles: Record<string, CellStyle> = {};
    const shared = new Map<string, { formula: string; row: number; col: number }>();
    for (const c of tags(doc, 'c')) {
      const ref = c.getAttribute('r');
      const at = ref ? parseA1(ref) : null;
      if (!ref || !at) continue;
      const style = xfStyles[Number(c.getAttribute('s') ?? 0)];
      if (style) styles[ref] = style;
      const t = c.getAttribute('t');
      const f = first(c, 'f');
      const v = first(c, 'v')?.textContent ?? null;
      let input = '';
      if (f) {
        let text = f.textContent ?? '';
        const si = f.getAttribute('si');
        if (f.getAttribute('t') === 'shared' && si !== null) {
          if (text) shared.set(si, { formula: `=${text}`, row: at.row, col: at.col });
          else {
            const master = shared.get(si);
            if (master) text = offsetFormula(master.formula, at.row - master.row, at.col - master.col).slice(1);
          }
        }
        input = text ? `=${text.replace(/_xl(fn|ws)\./g, '')}` : '';
      } else if (t === 's') input = strings[Number(v)] ?? '';
      else if (t === 'inlineStr') input = tags(c, 't').map((x) => x.textContent ?? '').join('');
      else if (t === 'b') input = v === '1' ? 'TRUE' : 'FALSE';
      else if (t === 'e' || t === 'str') input = v ?? '';
      else if (v !== null) input = v;
      // Text that would read as a number keeps a leading apostrophe, as Excel shows it.
      if ((t === 's' || t === 'inlineStr' || t === 'str') && !f && /^[\d.+\-$=]/.test(input) && Number.isFinite(Number(input.replace(/[$,%]/g, '')))) input = `'${input}`;
      if (input !== '') sheet.cells[ref] = input;
    }
    if (Object.keys(styles).length) sheet.styles = styles;
    const cols: Record<string, number> = {};
    for (const col of tags(doc, 'col')) {
      const min = Number(col.getAttribute('min'));
      const max = Math.min(Number(col.getAttribute('max')), min + 100);
      const w = Number(col.getAttribute('width'));
      if (col.getAttribute('hidden') === '1') {
        for (let i = min; i <= max; i++) (sheet.hiddenCols ??= []).push(columnName(i - 1));
      }
      if (!col.getAttribute('customWidth') || !w) continue;
      for (let i = min; i <= max; i++) cols[columnName(i - 1)] = Math.round(w * 7 + 5);
    }
    if (Object.keys(cols).length) sheet.cols = cols;
    const rows: Record<string, number> = {};
    for (const row of tags(doc, 'row')) {
      const ht = row.getAttribute('ht');
      if (ht && row.getAttribute('customHeight') === '1') rows[row.getAttribute('r') ?? ''] = Math.round(Number(ht) / 0.75);
    }
    if (Object.keys(rows).length) sheet.rows = rows;
    // Rows hidden by hand; rows an autofilter hides come back through the filter.
    const af = first(doc, 'autoFilter')?.getAttribute('ref');
    const afRows = af ? parseRange(af) : null;
    const hiddenRows = tags(doc, 'row')
      .filter((row) => row.getAttribute('hidden') === '1')
      .map((row) => Number(row.getAttribute('r')))
      .filter((n) => n > 0 && !(afRows && n - 1 > afRows.r0 && n - 1 <= afRows.r1));
    if (hiddenRows.length) sheet.hiddenRows = hiddenRows;
    const merges = tags(doc, 'mergeCell').map((m) => m.getAttribute('ref') ?? '').filter((r) => r.includes(':'));
    if (merges.length) sheet.merges = merges;
    const pane = first(doc, 'pane');
    if (pane?.getAttribute('state') === 'frozen') {
      sheet.freeze = { rows: Number(pane.getAttribute('ySplit') ?? 0), cols: Number(pane.getAttribute('xSplit') ?? 0) };
    }
    const tab = rgbOf(first(doc, 'tabColor'));
    if (tab) sheet.color = tab;
    readRules(doc, sheet, dxfStyles);
    const charts = readCharts(doc, path, sheet, parse);
    if (charts.length) sheet.charts = charts;
    sheets.push(sheet);
  }
  if (!sheets.length) throw new Error('no sheets');
  // Workbook-level names that point at a cell or range; Excel's own (_xlnm.*) and sheet-scoped ones are left out.
  const names: Record<string, string> = {};
  for (const dn of tags(book, 'definedName')) {
    const n = dn.getAttribute('name') ?? '';
    const ref = (dn.textContent ?? '').trim();
    if (!n || n.startsWith('_xlnm.') || dn.getAttribute('localSheetId') !== null || dn.getAttribute('hidden') === '1') continue;
    if (/^(?:'(?:[^']|'')+'|[^!'",()+\-*/&^=<>:;{} ]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i.test(ref)) names[n] = ref;
  }
  return { title: 'Imported', sheets, ...(Object.keys(names).length ? { names } : {}) };
}

/** Reads a dropped or picked file into sheets. */
export async function readSpreadsheetFile(file: File): Promise<WorkbookData> {
  const base = file.name.replace(/\.[^.]+$/, '');
  if (/\.xlsx$/i.test(file.name)) {
    const data = await importXlsx(await file.arrayBuffer());
    return { ...data, title: base };
  }
  const rows = parseCsv(await file.text());
  return { title: base, sheets: [sheetFromRows(base.slice(0, 31) || 'Sheet1', rows)] };
}

export const displayFor = (v: Scalar) => (v === null ? '' : v instanceof CellError ? v.code : String(v));

const OP_TYPES: Record<string, 'gt' | 'lt' | 'ge' | 'le' | 'eq' | 'ne'> = {
  greaterThan: 'gt',
  lessThan: 'lt',
  greaterThanOrEqual: 'ge',
  lessThanOrEqual: 'le',
  equal: 'eq',
  notEqual: 'ne',
};
const unquote = (f: string) => (/^".*"$/.test(f) ? f.slice(1, -1).replaceAll('""', '"') : f);

/** Conditional formats, dropdown lists and the autofilter of one worksheet. */
function readRules(doc: Document, sheet: SheetData, dxfs: HighlightStyle[]): void {
  const conditional: ConditionalRule[] = [];
  for (const cf of tags(doc, 'conditionalFormatting')) {
    const ranges = (cf.getAttribute('sqref') ?? '').split(/\s+/).filter((r) => parseRange(r));
    for (const rule of tags(cf, 'cfRule')) {
      const type = rule.getAttribute('type');
      const style = dxfs[Number(rule.getAttribute('dxfId') ?? -1)];
      const formulas = tags(rule, 'formula').map((f) => f.textContent ?? '');
      for (const range of ranges) {
        if (type === 'cellIs') {
          const op = rule.getAttribute('operator') ?? '';
          const num = (f: string | undefined) => (f !== undefined && Number.isFinite(Number(f)) ? Number(f) : unquote(f ?? ''));
          if (op === 'between' && formulas.length >= 2) conditional.push({ range, type: 'between', value: Number(formulas[0]), value2: Number(formulas[1]), style });
          else if (OP_TYPES[op]) conditional.push({ range, type: OP_TYPES[op]!, value: num(formulas[0]), style });
        } else if (type === 'containsText' && rule.getAttribute('text')) {
          conditional.push({ range, type: 'contains', text: rule.getAttribute('text')!, style });
        } else if (type === 'duplicateValues') {
          conditional.push({ range, type: 'duplicate', style });
        } else if (type === 'colorScale') {
          const colors = tags(rule, 'color').map((c) => rgbOf(c)).filter((c): c is string => Boolean(c));
          if (colors.length >= 2) conditional.push({ range, type: 'colorScale', colors });
        } else if (type === 'dataBar') {
          conditional.push({ range, type: 'dataBar', color: rgbOf(first(rule, 'color')) ?? '#638ec6' });
        }
      }
    }
  }
  if (conditional.length) sheet.conditional = conditional;

  const validation = tags(doc, 'dataValidation')
    .filter((v) => v.getAttribute('type') === 'list')
    .flatMap((v) => {
      const f = first(v, 'formula1')?.textContent ?? '';
      // Only literal lists; a list drawn from a range would need that range's values.
      if (!/^".*"$/.test(f)) return [];
      const list = unquote(f).split(',').map((x) => x.trim()).filter(Boolean);
      return (v.getAttribute('sqref') ?? '').split(/\s+/).filter((r) => parseRange(r)).map((range) => ({ range, list }));
    });
  if (validation.length) sheet.validation = validation;

  const af = first(doc, 'autoFilter');
  const ref = af?.getAttribute('ref');
  const fr = ref ? parseRange(ref) : null;
  if (af && ref && fr) {
    sheet.filter = { range: ref };
    const hidden: Record<string, string[]> = {};
    for (const fc of tags(af, 'filterColumn')) {
      const col = fr.c0 + Number(fc.getAttribute('colId') ?? 0);
      const filters = first(fc, 'filters');
      if (!filters) continue;
      const visible = new Set(tags(filters, 'filter').map((x) => x.getAttribute('val') ?? ''));
      if (filters.getAttribute('blank') === '1') visible.add(BLANK);
      const text = (r: number, c: number) => (sheet.cells[formatA1(r, c)] ?? '').replace(/^'/, '');
      const all = columnValues(sheet, col, text, { rows: fr.r1 + 1 });
      const off = all.filter((v) => !visible.has(v));
      if (off.length) hidden[columnName(col)] = off;
    }
    if (Object.keys(hidden).length) sheet.filter.hidden = hidden;
  }
}

/** A relative target in a relationships file, resolved against the part that owns it. */
function resolvePart(owner: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = owner.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}
const relsOf = (part: string) => {
  const i = part.lastIndexOf('/');
  return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`;
};
const RID = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Charts drawn on a worksheet: type, title, the range their series read, and where they sit. */
function readCharts(doc: Document, sheetPath: string, sheet: SheetData, parse: (name: string) => Document | null): ChartSpec[] {
  const drawingEl = first(doc, 'drawing');
  const rid = drawingEl?.getAttribute('r:id') ?? drawingEl?.getAttributeNS(RID, 'id');
  if (!rid) return [];
  const sheetRels = parse(relsOf(sheetPath));
  const target = sheetRels ? tags(sheetRels, 'Relationship').find((r) => r.getAttribute('Id') === rid)?.getAttribute('Target') : null;
  if (!target) return [];
  const drawingPath = resolvePart(sheetPath, target);
  const drawing = parse(drawingPath);
  const drawingRels = parse(relsOf(drawingPath));
  if (!drawing || !drawingRels) return [];
  const colPx = (c: number) => sheet.cols?.[columnName(c)] ?? 100;
  const rowPx = (r: number) => sheet.rows?.[String(r + 1)] ?? 24;
  const pos = (el: Element | null, size: (i: number) => number, idx: string, off: string) => {
    const i = Number(el ? first(el, idx)?.textContent : 0) || 0;
    let px = 0;
    for (let k = 0; k < i; k++) px += size(k);
    return px + (Number(el ? first(el, off)?.textContent : 0) || 0) / 9525;
  };
  const out: ChartSpec[] = [];
  for (const anchor of [...tags(drawing, 'twoCellAnchor'), ...tags(drawing, 'oneCellAnchor')]) {
    const chartEl = first(anchor, 'chart');
    const crid = chartEl?.getAttribute('r:id') ?? chartEl?.getAttributeNS(RID, 'id');
    const ctarget = crid ? tags(drawingRels, 'Relationship').find((r) => r.getAttribute('Id') === crid)?.getAttribute('Target') : null;
    const chart = ctarget ? parse(resolvePart(drawingPath, ctarget)) : null;
    if (!chart) continue;
    const type: ChartType | null = first(chart, 'barChart') || first(chart, 'bar3DChart') ? 'bar' : first(chart, 'lineChart') || first(chart, 'line3DChart') ? 'line' : first(chart, 'pieChart') || first(chart, 'pie3DChart') || first(chart, 'doughnutChart') ? 'pie' : null;
    if (!type) continue;
    // The range: the smallest block holding every reference the series make.
    const refs = tags(chart, 'f')
      .map((f) => (f.textContent ?? '').replace(/^.*!/, '').replaceAll('$', ''))
      .map((r) => parseRange(r))
      .filter((r): r is NonNullable<typeof r> => r !== null);
    if (!refs.length) continue;
    const box = refs.reduce((a, b) => ({ r0: Math.min(a.r0, b.r0), c0: Math.min(a.c0, b.c0), r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1) }));
    const titleEl = first(chart, 'title');
    const title = titleEl ? tags(titleEl, 't').map((t) => t.textContent ?? '').join('') : '';
    const from = first(anchor, 'from');
    const to = first(anchor, 'to');
    const x = Math.round(pos(from, colPx, 'col', 'colOff'));
    const y = Math.round(pos(from, rowPx, 'row', 'rowOff'));
    const ext = first(anchor, 'ext');
    const w = to ? Math.round(pos(to, colPx, 'col', 'colOff')) - x : Math.round(Number(ext?.getAttribute('cx') ?? 0) / 9525) || 440;
    const h = to ? Math.round(pos(to, rowPx, 'row', 'rowOff')) - y : Math.round(Number(ext?.getAttribute('cy') ?? 0) / 9525) || 270;
    out.push({
      id: `chart-${out.length + 1}`,
      type,
      range: `${formatA1(box.r0, box.c0)}:${formatA1(box.r1, box.c1)}`,
      ...(title ? { title } : {}),
      x,
      y,
      w: Math.max(160, w),
      h: Math.max(120, h),
    });
  }
  return out;
}
