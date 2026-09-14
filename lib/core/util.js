// Shared helpers: errors, caps, path utils, truncation.
export const CAPS = {
  MAX_WORD_INPUT_BYTES: 40 * 1024 * 1024,   // 40 MB
  MAX_EXCEL_INPUT_BYTES: 60 * 1024 * 1024,  // 60 MB
  MAX_TEXT_BYTES: 256 * 1024,               // textutil/html inputs
  DEFAULT_MAX_CHARS: 90000,                 // default content budget returned to the model
  DEFAULT_MAX_ROWS: 400,                    // xlsx default row window
  DEFAULT_MAX_COLS: 60,
  // office_query 要算全表,不能被 office_read 的窗口限制住;但仍然要有硬顶,
  // 免得一份超大表把内存吃光(超了会明确告知只统计了前 N 行)
  MAX_QUERY_ROWS: 200000,
  MAX_QUERY_COLS: 200,
  // 解压炸弹防护:docx/xlsx/odt 都是 zip,压缩包体积有上限,解压后体积也必须有限。
  // 绝对上限放到 1 GiB(避免误伤体积大但正常的表格),压缩比才是主要探测手段 ——
  // 正常 OOXML 文本压缩比通常 10~30:1,而压缩炸弹动辄上千比一。
  MAX_UNCOMPRESSED_BYTES: 1024 * 1024 * 1024, // 1 GiB
  MAX_COMPRESSION_RATIO: 150,
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

/**
 * Reject a zip whose declared uncompressed size (or ratio) exceeds the caps.
 * `.docx` / `.xlsx` / `.odt` are zips, so a small file can expand enormously —
 * the compressed-size cap alone is not enough.
 * @param buf zip bytes
 * @param label what is being opened, for the error message
 */
/** OOXML extensions whose bytes must be a zip (`PK` magic). */
export const ZIP_OFFICE_EXTS = new Set(['docx', 'xlsx', 'odt']);

/**
 * Refuse to leave a text file behind under an Office extension. Any OOXML
 * output must start with the zip magic; anything else means the content was
 * written as text (a fake ".xlsx" full of TSV is the classic failure).
 */
export function assertOfficeBinary(buf, ext) {
  const bytes = Buffer.isBuffer(buf) ? buf : Buffer.from(buf || []);
  const isZip = bytes.length > 3
    && bytes[0] === 0x50 && bytes[1] === 0x4b
    && (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07);
  if (!isZip) {
    throw new OfficeError(
      `生成 .${ext} 失败:输出不是有效的 Office 二进制(PK 头缺失),疑似被写成了文本 —— 已中止,避免留下假的 .${ext} 文件`,
      'BAD_OUTPUT_FORMAT'
    );
  }
  return bytes;
}

export async function assertZipBudget(buf, label) {
  const { default: PizZip } = await import('pizzip');
  let zip;
  try {
    zip = new PizZip(buf);
  } catch {
    return; // 不是合法 zip:交给各自的解析器给出更准确的报错
  }
  let total = 0;
  for (const name of Object.keys(zip.files)) {
    const entry = zip.files[name];
    const size = entry?._data?.uncompressedSize;
    if (typeof size === 'number' && size > 0) total += size;
    if (total > CAPS.MAX_UNCOMPRESSED_BYTES) {
      throw new OfficeError(
        `${label} 解压后体积超过 ${fmtBytes(CAPS.MAX_UNCOMPRESSED_BYTES)} 上限(疑似压缩炸弹),已拒绝`,
        'ZIP_BOMB_SUSPECTED'
      );
    }
  }
  const packed = Buffer.byteLength(buf);
  if (packed > 0 && total / packed > CAPS.MAX_COMPRESSION_RATIO) {
    throw new OfficeError(
      `${label} 压缩比异常(${Math.round(total / packed)}:1),已拒绝`,
      'ZIP_BOMB_SUSPECTED'
    );
  }
}

export async function readPath(path) {
  const { readFile } = await import('node:fs/promises');
  return readFile(path);
}

export async function writePath(path, data, options) {
  const { writeFile, mkdir } = await import('node:fs/promises');
  const { dirname } = await import('node:path');
  const dir = dirname(path);
  if (dir && dir !== '.') await mkdir(dir, { recursive: true });
  await writeFile(path, data, options);
  return path;
}
