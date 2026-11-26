// Cell value coercion, shared by the ExcelJS backend and the openpyxl fallback.
//
// Deliberately free of any heavy import: when DSH's profile resolver kills the
// ExcelJS module graph (see deps.js), this file still loads, so the fallback can
// reuse exactly the same encoding rules instead of re-implementing them in Python.

const ARGB = (hex) => String(hex || '').replace(/^#/, '').toUpperCase().slice(0, 6);

const DATE_PREFIX = /^date:(\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?)$/;
const NUMBER_PREFIX = /^(?:num|percent):(.*)$/;

/** Beautified first row: white bold on Excel-blue, centered and wrapped. */
export const HEADER_STYLE = { bold: true, color: 'FFFFFF', fill: '4472C4', align: 'center', valign: 'middle', wrap: true };

export function argb(hex) {
  return ARGB(hex);
}

/** `YYYY-MM-DD` from a Date's local calendar fields (Excel dates carry no timezone). */
export function isoLocalDay(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/** `date:`-prefixed ISO string -> Date (null when unparsable). */
export function parsePrefixedDate(s) {
  const m = DATE_PREFIX.exec(s);
  if (!m) return null;
  const iso = m[1].includes('T') ? m[1] : `${m[1].replace(' ', 'T')}T00:00:00`;
  const dt = new Date(iso);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

/** `num:` / `percent:` prefixed text -> number (null when unparsable). */
export function parsePrefixedNumber(s) {
  const m = NUMBER_PREFIX.exec(s);
  if (!m) return null;
  const n = Number(m[1].replaceAll(/[%,\s]/g, ''));
  return Number.isNaN(n) ? null : n;
}

function coerceString(s) {
  if (s.startsWith('=')) return { formula: s.slice(1) };
  const date = parsePrefixedDate(s);
  if (date) return date;
  const num = parsePrefixedNumber(s);
  return num ?? s;
}

/**
 * ExcelJS-flavoured cell value for a tool argument.
 * 一律返回单元格取值(字符串可能被解成公式对象 / 日期 / 数字),类型随入参。
 * @returns {*} ExcelJS 能直接写入的标量或 {formula} 对象
 */
export function coerceValue(raw) {
  if (raw === null || raw === undefined) return null;
  if (raw instanceof Date) return raw;
  if (typeof raw === 'number' || typeof raw === 'boolean') return raw;
  if (typeof raw === 'object') {
    if ('value' in raw) return coerceValue(raw.value);
    return String(raw);
  }
  return coerceString(typeof raw === 'string' ? raw : String(raw));
}

/**
 * Tagged cell value for the Python fallback: {k, v} with k in s/n/b/d/f.
 * A tag keeps the coercion in one place instead of re-deriving `=`/`date:`/`num:`
 * prefixes in the child process.
 */
export function cellTag(raw) {
  if (raw === null || raw === undefined) return { k: 'n', v: null };
  if (typeof raw === 'number') return { k: 'n', v: raw };
  if (typeof raw === 'boolean') return { k: 'b', v: raw };
  if (raw instanceof Date) return { k: 'd', v: isoLocalDay(raw) };
  const s = typeof raw === 'string' ? raw : String(raw);
  if (s.startsWith('=')) return { k: 'f', v: s.slice(1) };
  const date = parsePrefixedDate(s);
  if (date) return { k: 'd', v: isoLocalDay(date) };
  const num = parsePrefixedNumber(s);
  if (num !== null) return { k: 'n', v: num };
  return { k: 's', v: s };
}

/** A tool argument may be `{value, style}`; both backends read it the same way. */
export function splitCell(raw) {
  return {
    value: raw !== null && typeof raw === 'object' && 'value' in raw ? raw.value : raw,
    style: raw !== null && typeof raw === 'object' && 'style' in raw ? raw.style : null,
  };
}
