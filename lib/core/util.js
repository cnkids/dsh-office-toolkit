// Shared helpers: errors, caps, path utils, truncation.
export const CAPS = {
  MAX_WORD_INPUT_BYTES: 40 * 1024 * 1024,   // 40 MB
  MAX_EXCEL_INPUT_BYTES: 60 * 1024 * 1024,  // 60 MB
  MAX_TEXT_BYTES: 256 * 1024,               // textutil/html inputs
  DEFAULT_MAX_CHARS: 90000,                 // default content budget returned to the model
  DEFAULT_MAX_ROWS: 400,                    // xlsx default row window
  DEFAULT_MAX_COLS: 60,
};

export class OfficeError extends Error {
  constructor(message, code = 'OFFICE_ERROR') {
    super(message);
    this.name = 'OfficeError';
    this.code = code;
  }
}

function isAsciiAlnum(value) {
  if (value.length === 0) return false;
  for (let i = 0; i < value.length; i += 1) {
    const cp = value.codePointAt(i);
    if (!isAsciiLetter(cp) && !isAsciiDigit(cp)) return false;
  }
  return true;
}

/** Lower-case extension without the dot; '' when the tail is not `[a-z0-9]+`. */
export function extOf(path) {
  const s = String(path || '').toLowerCase();
  const dot = s.lastIndexOf('.');
  if (dot === -1) return '';
  const ext = s.slice(dot + 1);
  return isAsciiAlnum(ext) ? ext : '';
}

export function fmtBytes(n) {
  if (n == null) return '';
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

export function truncate(str, max) {
  if (str.length <= max) return str;
  return str.slice(0, max) + `\n\n…（内容已按 ${max} 字符截断，可用范围/行数参数读取其余部分）`;
}

const CODE_UPPER_A = 65;
const CODE_UPPER_Z = 90;
const CODE_LOWER_A = 97;
const CODE_LOWER_Z = 122;
const CODE_DIGIT_0 = 48;
const CODE_DIGIT_9 = 57;

export function isAsciiLetter(cp) {
  return (cp >= CODE_UPPER_A && cp <= CODE_UPPER_Z) || (cp >= CODE_LOWER_A && cp <= CODE_LOWER_Z);
}

export function isAsciiDigit(cp) {
  return cp >= CODE_DIGIT_0 && cp <= CODE_DIGIT_9;
}

export function isAllDigits(value) {
  if (typeof value !== 'string' || value.length === 0) return false;
  for (let i = 0; i < value.length; i += 1) {
    if (!isAsciiDigit(value.codePointAt(i))) return false;
  }
  return true;
}

export function isPlainObject(v) {
  return v !== null && typeof v === 'object' && !Array.isArray(v);
}

// Counts: paragraphs / table cells / words (CJK-aware rough count).
export function roughStats(text) {
  const words = (text.match(/[\u4e00-\u9fff]|[A-Za-z0-9]+/g) || []).length;
  const paragraphs = text.split(/\n{2,}/).filter((p) => p.trim()).length;
  return { words, paragraphs };
}

export function isWithinBytes(buf, cap) {
  return Buffer.byteLength(buf) <= cap;
}

export async function readPath(path) {
  const { readFile } = await import('node:fs/promises');
  return readFile(path);
}

export async function writePath(path, data) {
  const { writeFile, mkdir } = await import('node:fs/promises');
  const { dirname } = await import('node:path');
  const dir = dirname(path);
  if (dir && dir !== '.') await mkdir(dir, { recursive: true });
  await writeFile(path, data);
  return path;
}
