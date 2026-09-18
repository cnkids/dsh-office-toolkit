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
    const sheet = renderLegacySheet(XLSX, wb, names[i], { maxRows, maxCols, formulaMode: opts.formulaMode });
    parts.push(sheet.header, '```tsv', sheet.body, '```');
    meta.push(sheet.meta);
  }
  if (!parts.length) parts.push('(未匹配到工作表)');
  return { content: parts.join('\n'), meta: { kind: ext, sheets: names, reported: meta } };
}

// 文本类表格(CSV/TSV/TXT)没有自带编码信息:SheetJS 默认按 latin1 解码,中文会变乱码,
// 所以显式按 UTF-8 读成字符串再交给它;二进制格式(含 zip 的 xlsx/xlsb/ods)保持原样。
export const TEXT_TABLE_EXTS = new Set(['csv', 'tsv', 'txt']);

/** Read a sheet buffer with the encoding appropriate for `ext`. */
export function readSheetJs(XLSX, buf, ext) {
  if (!TEXT_TABLE_EXTS.has(ext)) return XLSX.read(buf, { type: 'buffer', cellDates: true });
  const text = Buffer.from(buf).toString('utf8').replace(/^\uFEFF/, '');
  return XLSX.read(text, { type: 'string', cellDates: true });
}

/**
 * CSV/TSV/TXT 没有公式概念,`=` 开头就是普通文本。SheetJS 读入时会把它标成公式
 * (`cell.f`),写出去就成了能触发 DDE/HYPERLINK 的活公式 —— 外部文件不可信,必须中和。
 */
export function neutralizeTextFormulas(wb) {
  for (const name of wb.SheetNames || []) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    for (const addr of Object.keys(ws)) {
      if (addr.startsWith('!')) continue;
      const cell = ws[addr];
      if (cell && typeof cell === 'object' && cell.f !== undefined) {
        delete cell.f;
        cell.t = 's';
      }
    }
  }
}

// ---------------- office_query 的数据源:原始值矩阵 ----------------
/** Date -> YYYY-MM-DD(用本地日历字段:Excel 日期没有时区,走 UTC 会让东八区差一天)。 */
function isoDay(d) {
  if (Number.isNaN(d.getTime())) return '';
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${d.getFullYear()}-${m}-${day}`;
}

/** SheetJS cell -> raw primitive: numbers stay numbers, dates become YYYY-MM-DD. */
function rawCell(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return isoDay(v);
  if (typeof v === 'object') return String(v.w ?? v.v ?? '').trim();
  if (typeof v === 'string') return v.trim();
  return typeof v === 'number' && !Number.isFinite(v) ? '' : v;
}

/** Whole-sheet raw matrix (no row window). */
export async function readMatrixSheets(buf, ext, opts = {}) {
  if (!isWithinBytes(buf, CAPS.MAX_EXCEL_INPUT_BYTES)) throw new OfficeError('表格文件超过 60 MB 上限', 'OFFICE_TOO_LARGE');
  const XLSX = await sheetjs();
  const wb = parseLegacyWorkbook(XLSX, buf, ext);
  const names = wb.SheetNames || [];
  const maxRows = opts.maxRows ?? CAPS.MAX_QUERY_ROWS;
  const maxCols = Math.min(opts.maxCols ?? CAPS.MAX_QUERY_COLS, CAPS.MAX_QUERY_COLS);
  const sheets = [];
  for (const name of names) {
    if (wb.Sheets[name]) sheets.push(readSheetMatrix(XLSX, wb.Sheets[name], name, maxRows, maxCols));
  }
  return { sheets, sheetNames: names };
}

function readSheetMatrix(XLSX, ws, name, maxRows, maxCols) {
  // raw:true 保留数字/日期类型;blankrows:false 跳过空行(与 ExcelJS 侧行为一致)
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false, dateNF: 'yyyy-mm-dd' });
  const matrix = [];
  let truncated = false;
  let totalCols = 0;
  for (const rowArr of aoa) {
    if (matrix.length >= maxRows) { truncated = true; break; }
    totalCols = Math.max(totalCols, rowArr.length);
    matrix.push(rowArr.slice(0, maxCols).map(rawCell));
  }
  return { name, matrix, truncated, totalRows: aoa.length, totalCols };
}

/** 老格式(.xls/.ods)的公式读取:SheetJS 把公式放在格子的 .f 上。 */
function legacyFormulaText(XLSX, ws, range, rowIndex, colIndex, shown, mode) {
  const cell = ws[XLSX.utils.encode_cell({ r: range.s.r + rowIndex, c: range.s.c + colIndex })];
  if (!cell || typeof cell.f !== 'string' || !cell.f) return shown;
  const formula = `=${cell.f}`;
  return mode === 'formula' || shown === '' ? formula : `${formula} → ${shown}`;
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

function renderLegacySheet(XLSX, wb, name, { maxRows, maxCols, formulaMode = 'value' }) {
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
    for (let c = 0; c < cellCount; c++) {
      const shown = rowArr[c] ?? '';
      row.push(formulaMode === 'value' ? shown : legacyFormulaText(XLSX, ws, range, r, c, shown, formulaMode));
    }
    lines.push(trimTrailingTabs(row.map(escapeCell).join('\t')));
  }
  const shown = lines.length;
  return {
    header: `### 工作表「${name}」(${rows} 行 × ${cols} 列${shown < rows ? ', 仅显示前 ' + shown + ' 行' : ''})`,
    body: lines.join('\n') || '(空)',
    meta: { name, rows, cols, shown },
  };
}
