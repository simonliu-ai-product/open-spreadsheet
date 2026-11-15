/**
 * Excel number formats: "#,##0.00", "0%", "yyyy-mm-dd", "[Red]-#,##0;…",
 * "$#,##0_);($#,##0)". Used by TEXT() and by the grid to show a cell.
 *
 * Supported: up to four sections (positive;negative;zero;text), colours,
 * conditions are ignored, quoted and escaped literals, digit placeholders
 * 0 # ?, thousands separators and scaling commas, percent, scientific E+00,
 * dates and times (y m d h s, AM/PM, elapsed [h]), and @ for text.
 */
import { serialToDate } from './dates.ts';
import { CellError, formatGeneral, type Scalar } from './values.ts';

export interface Formatted {
  text: string;
  /** A colour the format names, like [Red]. */
  color?: string;
}

const COLORS: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#008000',
  magenta: '#ff00ff',
  red: '#ff0000',
  white: '#ffffff',
  yellow: '#ffff00',
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Splits on `;` outside quotes and brackets. */
function sections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quote = false;
  let bracket = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!;
    if (ch === '\\' && !quote) {
      cur += ch + (code[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === '"') quote = !quote;
    else if (!quote && ch === '[') bracket = true;
    else if (!quote && ch === ']') bracket = false;
    if (ch === ';' && !quote && !bracket) {
      out.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

type Tok = { t: 'lit'; v: string } | { t: 'code'; v: string };

/** Literals ("text", \x, _x padding, *x fill) apart from format codes; brackets read out. */
function lex(section: string): { toks: Tok[]; color?: string; elapsed: boolean } {
  const toks: Tok[] = [];
  let color: string | undefined;
  let elapsed = false;
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]!;
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      toks.push({ t: 'lit', v: section.slice(i + 1, end === -1 ? undefined : end) });
      i = end === -1 ? section.length : end;
    } else if (ch === '\\') {
      toks.push({ t: 'lit', v: section[i + 1] ?? '' });
      i++;
    } else if (ch === '_') {
      toks.push({ t: 'lit', v: ' ' });
      i++;
    } else if (ch === '*') {
      i++; // fill to the cell width: not meaningful as text
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      const inner = section.slice(i + 1, end === -1 ? undefined : end);
      i = end === -1 ? section.length : end;
      const lower = inner.toLowerCase();
      if (COLORS[lower]) color = COLORS[lower];
      else if (/^color\d+$/.test(lower)) color = undefined;
      else if (/^(h+|m+|s+)$/i.test(inner)) {
        elapsed = true;
        toks.push({ t: 'code', v: `[${lower}]` });
      } else if (inner.startsWith('$')) {
        // [$NT$-404]: the currency symbol before the dash.
        toks.push({ t: 'lit', v: inner.slice(1).split('-')[0]! });
      }
      // Conditions like [>100] are ignored.
    } else if (/[$\-+/():!^&'~{}<>= ]/.test(ch)) {
      toks.push({ t: 'lit', v: ch });
    } else {
      toks.push({ t: 'code', v: ch });
    }
  }
  return { toks, color, elapsed };
}

const isDateCode = (toks: Tok[]) =>
  toks.some((t) => t.t === 'code' && (/^[ydhsDHSY]$/.test(t.v) || t.v.startsWith('['))) ||
  toks.some((t) => t.t === 'code' && /^[mM]$/.test(t.v)) && !toks.some((t) => t.t === 'code' && /[0#?]/.test(t.v));

/** Formats a value with an Excel format code. */
export function formatValue(v: Scalar, code: string | undefined): Formatted {
  if (v === null) return { text: '' };
  if (v instanceof CellError) return { text: v.code };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE' };
  const fmt = code ?? '';
  if (fmt === '' || /^general$/i.test(fmt)) return { text: typeof v === 'number' ? formatGeneral(v) : v };
  const parts = sections(fmt);
  if (typeof v === 'string') {
    const textSection = parts[3] ?? (parts.length === 1 && parts[0]!.includes('@') ? parts[0] : undefined);
    if (textSection === undefined) return { text: v };
    const { toks, color } = lex(textSection);
    return { text: toks.map((t) => (t.t === 'lit' ? t.v : t.v === '@' ? v : t.v)).join(''), color };
  }
  let section = parts[0]!;
  let n = v;
  let negativeSection = false;
  if (parts.length >= 3 && n === 0) section = parts[2]!;
  else if (parts.length >= 2 && n < 0) {
    section = parts[1]!;
    n = -n;
    negativeSection = true;
  }
  const { toks, color, elapsed } = lex(section);
  if (toks.length === 0) return { text: '', color };
  if (/^general$/i.test(toks.map((t) => t.v).join(''))) return { text: formatGeneral(v), color };
  if (isDateCode(toks)) return { text: formatDate(n, toks, elapsed), color };
  const text = formatNumber(n, toks);
  return { text: !negativeSection && n < 0 && !text.startsWith('-') ? `-${text}` : text, color };
}

function formatDate(serial: number, toks: Tok[], elapsed: boolean): string {
  if (serial < 0 && !elapsed) return '#'.repeat(10);
  const ampm = toks.some((t) => t.t === 'code' && /^[aA]$/.test(t.v));
  const { year, month, day } = serialToDate(serial);
  let secs = Math.round((serial - Math.floor(serial)) * 86400);
  if (secs >= 86400) secs = 86399;
  const hours = Math.floor(secs / 3600);
  const minutes = Math.floor((secs % 3600) / 60);
  const seconds = secs % 60;
  const dow = ((Math.floor(serial) % 7) + 6) % 7; // 0 = Sunday, with serial 1 a Sunday

  // Group runs of the same letter: "yyyy", "mm", "AM/PM".
  const runs: Tok[] = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i]!;
    if (t.t === 'code' && /^[aA]$/.test(t.v)) {
      const rest = toks.slice(i, i + 5).map((x) => x.v).join('');
      if (/^am\/pm$/i.test(rest)) {
        runs.push({ t: 'code', v: 'AM/PM' });
        i += 4;
        continue;
      }
      if (/^a\/p$/i.test(toks.slice(i, i + 3).map((x) => x.v).join(''))) {
        runs.push({ t: 'code', v: 'A/P' });
        i += 2;
        continue;
      }
    }
    const last = runs[runs.length - 1];
    if (t.t === 'code' && last?.t === 'code' && last.v[0]!.toLowerCase() === t.v.toLowerCase() && /^[ymdhs]$/i.test(t.v)) last.v += t.v;
    else runs.push({ ...t });
  }
  const isCode = (t: Tok | undefined, re: RegExp) => t?.t === 'code' && re.test(t.v);
  let out = '';
  for (let i = 0; i < runs.length; i++) {
    const t = runs[i]!;
    if (t.t === 'lit') {
      out += t.v;
      continue;
    }
    const lower = t.v.toLowerCase();
    // "m" means minutes right after an hour or right before a second.
    const prev = runs.slice(0, i).reverse().find((x) => x.t === 'code');
    const next = runs.slice(i + 1).find((x) => x.t === 'code');
    const minute = lower[0] === 'm' && lower.length <= 2 && (isCode(prev, /^h+$|^\[h+\]$/i) || isCode(next, /^s+$/i));
    if (lower === '[h]' || lower === '[hh]') out += String(Math.floor(serial * 24)).padStart(lower.length - 2, '0');
    else if (lower === '[m]' || lower === '[mm]') out += String(Math.floor(serial * 1440));
    else if (lower === '[s]' || lower === '[ss]') out += String(Math.floor(serial * 86400));
    else if (minute) out += lower.length === 2 ? String(minutes).padStart(2, '0') : String(minutes);
    else if (lower.startsWith('y')) out += lower.length <= 2 ? String(year).slice(-2) : String(year);
    else if (lower === 'm') out += String(month);
    else if (lower === 'mm') out += String(month).padStart(2, '0');
    else if (lower === 'mmm') out += MONTHS[month - 1]!.slice(0, 3);
    else if (lower === 'mmmmm') out += MONTHS[month - 1]![0];
    else if (lower.startsWith('mmmm')) out += MONTHS[month - 1];
    else if (lower === 'd') out += String(day);
    else if (lower === 'dd') out += String(day).padStart(2, '0');
    else if (lower === 'ddd') out += DAYS[dow]!.slice(0, 3);
    else if (lower.startsWith('dddd')) out += DAYS[dow];
    else if (lower.startsWith('h')) {
      const h = ampm ? hours % 12 || 12 : hours;
      out += lower.length >= 2 ? String(h).padStart(2, '0') : String(h);
    } else if (lower.startsWith('s')) out += lower.length >= 2 ? String(seconds).padStart(2, '0') : String(seconds);
    else if (t.v === 'AM/PM') out += hours < 12 ? 'AM' : 'PM';
    else if (t.v === 'A/P') out += hours < 12 ? 'A' : 'P';
    else if (lower === 'e') out += String(year);
    else if (lower === '.') out += '.';
    else out += t.v;
  }
  return out;
}

function formatNumber(n: number, toks: Tok[]): string {
  const codes = toks.map((t) => (t.t === 'code' ? t.v : ''));
  const isPh = (c: string) => c === '0' || c === '#' || c === '?';
  // Scientific: everything after E+/E- is the exponent.
  const eAt = toks.findIndex((t, i) => t.t === 'code' && /^[eE]$/.test(t.v) && (toks[i + 1]?.v === '+' || toks[i + 1]?.v === '-'));
  const percents = codes.filter((c) => c === '%').length + toks.filter((t) => t.t === 'lit' && t.v === '%').length;
  let value = n * 100 ** percents;

  const body = eAt === -1 ? toks : toks.slice(0, eAt);
  const dot = body.findIndex((t) => t.t === 'code' && t.v === '.');
  const intToks = dot === -1 ? body : body.slice(0, dot);
  const fracToks = dot === -1 ? [] : body.slice(dot + 1);
  const intPh = intToks.filter((t) => t.t === 'code' && isPh(t.v));
  const fracPh = fracToks.filter((t) => t.t === 'code' && isPh(t.v));

  // Commas after the last integer placeholder scale by 1000 each; between placeholders they group.
  const lastPh = intToks.map((t) => t.t === 'code' && isPh(t.v)).lastIndexOf(true);
  let grouped = false;
  intToks.forEach((t, i) => {
    if (t.t === 'code' && t.v === ',') {
      if (i < lastPh) grouped = true;
      else if (dot === -1 || i > lastPh) value /= 1000;
    }
  });

  let exponent = 0;
  if (eAt !== -1 && value !== 0) {
    const intDigits = Math.max(1, intPh.length);
    exponent = Math.floor(Math.log10(Math.abs(value))) - (intDigits - 1);
    value /= 10 ** exponent;
  }

  const decimals = fracPh.length;
  const neg = value < 0;
  const fixed = roundHalfUp(Math.abs(value), decimals).toFixed(decimals);
  let [intDigits, fracDigits = ''] = fixed.split('.') as [string, string?];
  if (intDigits === '0' && !intPh.some((t) => t.v === '0')) intDigits = '';

  // Integer part: fill placeholders right to left; spare digits go to the leftmost one.
  const intOut: string[] = [];
  let di = intDigits.length - 1;
  const phIdx = intToks.map((t, i) => (t.t === 'code' && isPh(t.v) ? i : -1)).filter((i) => i >= 0);
  const firstPh = phIdx[0] ?? -1;
  for (let i = intToks.length - 1; i >= 0; i--) {
    const t = intToks[i]!;
    if (t.t === 'lit') {
      intOut.unshift(t.v);
    } else if (isPh(t.v)) {
      if (i === firstPh) {
        const rest = di >= 0 ? intDigits.slice(0, di + 1) : t.v === '0' ? '0' : t.v === '?' ? ' ' : '';
        intOut.unshift(rest);
        di = -1;
      } else if (di >= 0) {
        intOut.unshift(intDigits[di]!);
        di--;
      } else {
        intOut.unshift(t.v === '0' ? '0' : t.v === '?' ? ' ' : '');
      }
    } else if (t.v === '%') {
      intOut.unshift('%');
    }
  }
  let intText = intOut.join('');
  if (grouped) {
    // Group the run of digits only, leaving literals where they are.
    intText = intText.replace(/\d+/, (d) => d.replace(/\B(?=(\d{3})+(?!\d))/g, ','));
  }

  let fracText = '';
  let fi = 0;
  for (const t of fracToks) {
    if (t.t === 'lit') fracText += t.v;
    else if (isPh(t.v)) {
      const d = fracDigits[fi++] ?? '0';
      fracText += d;
    } else if (t.v === '%') fracText += '%';
  }
  // Optional decimals (#) drop trailing zeros.
  if (fracPh.length) {
    let trim = 0;
    for (let i = fracPh.length - 1; i >= 0 && fracPh[i]!.v !== '0' && fracDigits[i] === '0'; i--) trim++;
    if (trim) {
      fracDigits = fracDigits.slice(0, fracDigits.length - trim);
      fracText = fracText.replace(new RegExp(`0{${trim}}(?=[^0-9]*$)`), '');
    }
  }
  // Excel keeps the point even when every optional decimal drops: 2 in "0.##" is "2.".
  let out = intText + (dot !== -1 ? `.${fracText}` : fracText);

  if (eAt !== -1) {
    const sign = toks[eAt + 1]!.v;
    const expDigits = toks.slice(eAt + 2).filter((t) => t.t === 'code' && isPh(t.v)).length || 1;
    const e = Math.abs(exponent);
    out += `E${exponent < 0 ? '-' : sign === '+' ? '+' : ''}${String(e).padStart(expDigits, '0')}`;
    out += toks
      .slice(eAt + 2)
      .filter((t) => t.t === 'lit')
      .map((t) => t.v)
      .join('');
  }
  return neg && /[1-9]/.test(out) ? `-${out}` : out;
}

function roundHalfUp(x: number, decimals: number): number {
  const f = 10 ** decimals;
  return Math.round(Number.parseFloat((x * f).toPrecision(15))) / f;
}

/**
 * The format Excel gives a cell from what was typed: "2026-01-05" is a date,
 * "12%" a percentage, "$1,200" currency, "1,200" a grouped number.
 */
export function impliedFormat(input: string): string | undefined {
  const t = input.trim();
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(t)) return 'yyyy-mm-dd';
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}[ T]\d{1,2}:\d{2}(:\d{2})?$/.test(t)) return 'yyyy-mm-dd hh:mm';
  if (/^\d{1,2}:\d{2}:\d{2}$/.test(t)) return 'h:mm:ss';
  if (/^\d{1,2}:\d{2}$/.test(t)) return 'h:mm';
  if (/^[+-]?[\d,]*\.?\d+%$/.test(t)) return t.includes('.') ? '0.00%' : '0%';
  if (/^-?\$[\d,]+(\.\d+)?$/.test(t)) return t.includes('.') ? '$#,##0.00' : '$#,##0';
  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return t.includes('.') ? '#,##0.00' : '#,##0';
  return undefined;
}
