// Legacy spreadsheets: .xls/.xlsb/.ods/.csv read & conversion via SheetJS (pure JS,
// cross-platform). Legacy Word formats live in converters.js / legacy-read.js.
import { OfficeError, isWithinBytes, CAPS } from './util.js';

export { fmtBytes } from './util.js';

// ---------------- legacy spreadsheets (.xls/.ods/.csv read; conversions) ----------------
let xlsxMod = null;
async function sheetjs() {
  if (!xlsxMod) xlsxMod = await import('@e965/xlsx');
  return xlsxMod;
}

const SHEETJS_READABLE = ['xls', 'xlsb', 'xlsx', 'ods', 'csv', 'tsv', 'txt'];
const SHEETJS_WRITABLE = ['xlsx', 'csv', 'txt', 'html', 'ods'];

export async function readLegacySpreadsheet(buf, ext, opts = {}) {
  if (!isWithinBytes(buf, CAPS.MAX_EXCEL_INPUT_BYTES)) throw new OfficeError('表格文件超过 60 MB 上限', 'OFFICE_TOO_LARGE');
  const XLSX = await sheetjs();
  const wb = parseLegacyWorkbook(XLSX, buf, ext);
  const maxRows = Math.max(1, Math.min(opts.maxRows ?? CAPS.DEFAULT_MAX_ROWS, 5000));
  const maxCols = Math.max(1, Math.min(opts.maxCols ?? CAPS.DEFAULT_MAX_COLS, 200));
  const names = wb.SheetNames || [];
  const parts = [];
  const meta = [];
  for (let i = 0; i < names.length; i++) {
    if (!sheetWanted(opts.sheets, names[i], i)) continue;
    const sheet = renderLegacySheet(XLSX, wb, names[i], { maxRows, maxCols });
    parts.push(sheet.header, '```tsv', sheet.body, '```');
    meta.push(sheet.meta);
  }
  if (!parts.length) parts.push('(未匹配到工作表)');
  return { content: parts.join('\n'), meta: { kind: ext, sheets: names, reported: meta } };
}

// 文本类表格(CSV/TSV/TXT)没有自带编码信息:SheetJS 默认按 latin1 解码,中文会变乱码,
// 所以显式按 UTF-8 读成字符串再交给它;二进制格式(含 zip 的 xlsx/xlsb/ods)保持原样。
const TEXT_TABLE_EXTS = new Set(['csv', 'tsv', 'txt']);

/** Read a sheet buffer with the encoding appropriate for `ext`. */
export function readSheetJs(XLSX, buf, ext) {
  if (!TEXT_TABLE_EXTS.has(ext)) return XLSX.read(buf, { type: 'buffer', cellDates: true });
  const text = Buffer.from(buf).toString('utf8').replace(/^\uFEFF/, '');
  return XLSX.read(text, { type: 'string', cellDates: true });
}

function parseLegacyWorkbook(XLSX, buf, ext) {
  try {
    return readSheetJs(XLSX, buf, ext);
  } catch (err) {
    throw new OfficeError(`无法解析 .${ext} 文件: ${err?.message || err}`, 'BAD_LEGACY_SPREADSHEET');
  }
}

/** Whether the caller asked for this sheet (by name, 0-based index or number). */
function sheetWanted(want, name, index) {
  if (!want) return true;
  const list = Array.isArray(want) ? want : [want];
  return list.some((s) => s === name || s === index || s === String(index));
}

/** Tabs/newlines are made visible so the TSV block stays parseable. */
function escapeCell(v) {
  return String(v).replaceAll('\t', '␉').replaceAll(/\r?\n/g, '⏎');
}

/** Drop trailing tabs (a `/\t+$/g` scan backtracks super-linearly). */
function trimTrailingTabs(s) {
  let end = s.length;
  while (end > 0 && s[end - 1] === '\t') end -= 1;
  return s.slice(0, end);
}

function renderLegacySheet(XLSX, wb, name, { maxRows, maxCols }) {
  const ws = wb.Sheets[name];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const rows = range.e.r - range.s.r + 1;
  const cols = range.e.c - range.s.c + 1;
  const rowsArr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
  const lines = [];
  const limit = Math.min(rowsArr.length, maxRows);
  const cellCount = Math.min(cols, maxCols);
  for (let r = 0; r < limit; r++) {
    const rowArr = rowsArr[r] || [];
    const row = [];
    for (let c = 0; c < cellCount; c++) row.push(rowArr[c] ?? '');
    lines.push(trimTrailingTabs(row.map(escapeCell).join('\t')));
  }
  const shown = lines.length;
  return {
    header: `### 工作表「${name}」(${rows} 行 × ${cols} 列${shown < rows ? ', 仅显示前 ' + shown + ' 行' : ''})`,
    body: lines.join('\n') || '(空)',
    meta: { name, rows, cols, shown },
  };
}
