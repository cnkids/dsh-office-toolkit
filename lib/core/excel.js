// .xlsx core on ExcelJS: read windows, build workbooks, apply edit ops.
import ExcelJS from '@wekanteam/exceljs';
import { OfficeError, isWithinBytes, CAPS, isPlainObject, fmtBytes, readPath, isAsciiDigit, isAsciiLetter, assertZipBudget } from './util.js';
import { injectChart } from './charts.js';
import { openOoxml } from './ooxml.js';

const ARGB = (hex) => String(hex || '').replace(/^#/, '').toUpperCase().slice(0, 6);

/** ExcelJS cell value -> display string. Objects may be errors/links/formulas. */
function cellDisplay(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return dateDisplay(value);
  if (typeof value !== 'object') return primitiveDisplay(value);
  return objectDisplay(value);
}

function primitiveDisplay(value) {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

function dateDisplay(value) {
  if (Number.isNaN(value.getTime())) return '';
  // 用本地日历字段而非 toISOString():Excel 日期没有时区,走 UTC 会让东八区回读差一天
  const y = value.getFullYear();
  const m = String(value.getMonth() + 1).padStart(2, '0');
  const d = String(value.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function objectDisplay(value) {
  if (value.error !== undefined) return `#${value.error}`;
  if (value.result !== undefined && value.result !== null) return cellDisplay(value.result);
  if (Array.isArray(value.richText)) return value.richText.map((r) => r.text || '').join('');
  if (value.hyperlink !== undefined) return String(value.text ?? value.hyperlink);
  if (value.formula !== undefined) return String(value.formula);
  return String(value.text ?? value);
}

function rowToCells(row, colCount) {
  const out = [];
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    out.push(cellDisplay(cell.value));
  }
  while (out.length && out.at(-1) === '') out.pop();
  return out;
}

async function loadWorkbook(buf, label) {
  await assertZipBudget(buf, 'xlsx');
  const { buffer } = await openOoxml(buf, 'xlsx'); // 修正非标准条目名与引用
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buffer);
  } catch (err) {
    throw new OfficeError(label + (err?.message || err), 'BAD_XLSX');
  }
  return wb;
}

export async function readWorkbook(buf, opts = {}) {
  if (!isWithinBytes(buf, CAPS.MAX_EXCEL_INPUT_BYTES)) {
    throw new OfficeError('表格文件超过 60 MB 上限', 'OFFICE_TOO_LARGE');
  }
  const wb = await loadWorkbook(buf, '不是有效的 .xlsx 文件(ExcelJS 无法解析): ');

  const maxRows = Math.max(1, Math.min(opts.maxRows ?? CAPS.DEFAULT_MAX_ROWS, 5000));
  const maxCols = Math.max(1, Math.min(opts.maxCols ?? CAPS.DEFAULT_MAX_COLS, 200));
  const sheets = [];
  const wantSheets = normalizeSheetSelector(opts.sheets);

  wb.eachSheet((ws, index) => {
    if (wantSheets && !wantSheets.some((s) => s === index - 0 || s === ws.name || s === String(index - 1))) return;
    sheets.push(readSheetWindow(ws, { ...opts, maxRows, maxCols }));
  });

  const parts = [];
  for (const s of sheets) {
    parts.push(`### 工作表「${s.name}」(${s.rows} 行 × ${s.cols} 列${s.shown < s.rows ? ', 仅显示前 ' + s.shown + ' 行' : ''})`);
    if (s.truncated) parts.push(`> 行数超过窗口(${maxRows} 行)，可用 sheets/range/maxRows 参数分段读取`);
    parts.push('```tsv', s.tsv || '(空)', '```');
  }

  const meta = {
    kind: 'xlsx',
    sheets: wb.worksheets.length,
    sheetNames: wb.worksheets.map((ws) => ws.name),
    reported: sheets.map((s) => ({ name: s.name, rows: s.rows, cols: s.cols, shown: s.shown })),
  };
  return { content: parts.join('\n'), meta, tsvBySheet: Object.fromEntries(sheets.map((s) => [s.name, s.tsv])) };
}

function readSheetWindow(ws, opts) {
  const maxCols = opts.maxCols;
  const maxRows = opts.maxRows;
  const colCount = Math.min(ws.columnCount || 0, maxCols + 60);
  const headers = [];
  const lines = [];
  let shown = 0;
  let truncated = false;
  const startRow = opts.range ? rangeStartRow(opts.range) : 1;
  const endRowLimit = opts.range ? rangeEndRow(opts.range) : startRow + maxRows - 1;

  ws.eachRow({ includeEmpty: true }, (row, rn) => {
    if (rn < startRow) return;
    if (rn > endRowLimit || shown >= maxRows) { truncated = true; return; }
    const cells = rowToCells(row, colCount);
    lines.push(cells.map(escapeTsv).join('\t'));
    if (rn === 1) headers.push(...cells);
    shown += 1;
  });

  return {
    name: ws.name,
    rows: ws.rowCount,
    cols: ws.columnCount || 0,
    shown,
    truncated,
    header: headers.slice(0, 20).join(' | '),
    tsv: lines.join('\n'),
  };
}

function escapeTsv(v) {
  return String(v).replaceAll('\t', '␉').replaceAll(/\r?\n/g, '⏎');
}

function normalizeSheetSelector(sel) {
  if (!sel) return null;
  const list = Array.isArray(sel) ? sel : [sel];
  const out = [];
  for (const s of list) {
    if (typeof s === 'number') out.push(s);
    else out.push(String(s).trim());
  }
  return out.length ? out : null;
}

/** End index of the digit run starting at `from` (-1 when none). */
function digitRunEnd(s, from) {
  let i = from;
  while (i < s.length && isAsciiDigit(s.codePointAt(i))) i += 1;
  return i === from ? -1 : i;
}

function firstDigitIndex(s) {
  for (let i = 0; i < s.length; i += 1) {
    if (isAsciiDigit(s.codePointAt(i))) return i;
  }
  return -1;
}

/** Start index of the digits in the first `:LETTERS DIGITS` group, else -1. */
function endRowDigitsIndex(s) {
  for (let i = 0; i < s.length; i += 1) {
    if (s[i] !== ':') continue;
    let j = i + 1;
    const lettersStart = j;
    while (j < s.length && isAsciiLetter(s.codePointAt(j))) j += 1;
    if (j === lettersStart) continue;
    const digitsStart = j;
    if (digitRunEnd(s, j) !== -1) return digitsStart;
  }
  return -1;
}

function rangeStartRow(range) {
  const s = String(range || '');
  const at = firstDigitIndex(s);
  return at === -1 ? 1 : Math.max(1, Number(s.slice(at, digitRunEnd(s, at))));
}

function rangeEndRow(range) {
  const s = String(range || '');
  const at = endRowDigitsIndex(s);
  return at === -1 ? Infinity : Number(s.slice(at, digitRunEnd(s, at)));
}

// ---------- value coercion (shared by write & edit) ----------
const DATE_PREFIX = /^date:(\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?)$/;
const NUMBER_PREFIX = /^(?:num|percent):(.*)$/;

/**
 * Turn a tool argument into an ExcelJS cell value.
 * Strings may encode formulas (`=`), dates (`date:`) or forced numbers (`num:`).
 * @returns {*} ExcelJS cell value (its type follows the input).
 */
function coerceValue(raw) {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'number' || typeof raw === 'boolean' || raw instanceof Date) return raw;
  if (typeof raw === 'string') return coerceString(raw);
  if (isPlainObject(raw) && 'value' in raw) return coerceValue(raw.value);
  return String(raw);
}

/** @returns {*} ExcelJS cell value encoded by a string argument. */
function coerceString(s) {
  if (s.startsWith('=')) return { formula: s.slice(1) };
  const date = parsePrefixedDate(s);
  if (date) return date;
  const num = NUMBER_PREFIX.exec(s);
  if (num) {
    const n = Number(num[1].replaceAll(/[%,\s]/g, ''));
    return Number.isNaN(n) ? s : n;
  }
  return s;
}

function parsePrefixedDate(s) {
  const m = DATE_PREFIX.exec(s);
  if (!m) return null;
  const iso = m[1].includes('T') ? m[1] : `${m[1].replace(' ', 'T')}T00:00:00`;
  const dt = new Date(iso);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function applyCellStyle(cell, style) {
  if (!isPlainObject(style)) return;
  const font = {};
  if (style.bold) font.bold = true;
  if (style.italic) font.italic = true;
  if (style.fontSize) font.size = Number(style.fontSize);
  if (style.color) font.color = { argb: 'FF' + ARGB(style.color) };
  if (style.name) font.name = String(style.name);
  if (Object.keys(font).length) cell.font = { ...cell.font, ...font };
  if (style.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + ARGB(style.fill) } };
  const align = {};
  if (style.align) align.horizontal = String(style.align);
  if (style.valign) align.vertical = String(style.valign);
  if (style.wrap) align.wrapText = true;
  if (Object.keys(align).length) cell.alignment = { ...cell.alignment, ...align };
  if (style.numFmt) cell.numFmt = String(style.numFmt);
  if (style.border) {
    const b = { style: 'thin', color: { argb: 'FF' + ARGB(style.borderColor || '000000') } };
    cell.border = { top: b, bottom: b, left: b, right: b };
  }
}

const HEADER_STYLE = { bold: true, color: 'FFFFFF', fill: '4472C4', align: 'center', valign: 'middle', wrap: true };

function buildSheet(ws, spec) {
  const { rows = [], header = false, headerStyle, columnWidths, rowHeights } = spec;
  if (!Array.isArray(rows)) throw new OfficeError(`工作表「${ws.name}」的 rows 必须是二维数组`);
  rows.forEach((r, ri) => {
    if (!Array.isArray(r)) throw new OfficeError(`工作表「${ws.name}」第 ${ri + 1} 行不是数组`);
    r.forEach((v, ci) => {
      const cell = ws.getCell(ri + 1, ci + 1);
      const raw = isPlainObject(v) && 'value' in v ? v.value : v;
      const style = isPlainObject(v) && 'style' in v ? v.style : null;
      if (raw === null || raw === undefined) return;
      const val = coerceValue(raw);
      if (val !== null) cell.value = val;
      if (style) applyCellStyle(cell, style);
    });
    if (header && ri === 0) {
      r.forEach((_v, ci) => applyCellStyle(ws.getCell(1, ci + 1), headerStyle === false ? null : (headerStyle || HEADER_STYLE)));
    }
  });
  if (Array.isArray(columnWidths)) {
    columnWidths.forEach((w, i) => {
      const num = Number(w);
      if (num > 0 && num <= 255) ws.getColumn(i + 1).width = num;
    });
  }
  if (Array.isArray(rowHeights)) {
    rowHeights.forEach((h, i) => {
      const num = Number(h);
      if (num > 0) ws.getRow(i + 1).height = num;
    });
  }
  // sensible default widths so CJK text is visible
  if (!Array.isArray(columnWidths)) applyDefaultWidths(ws);
}

function applyDefaultWidths(ws) {
  ws.columns.forEach((col, i) => {
    if (col.width != null) return;
    const longest = (col.values || []).reduce((m, v) => Math.max(m, String(v ?? '').length + 2), 8);
    ws.getColumn(i + 1).width = Math.max(10, Math.min(50, longest));
  });
}

function getWorksheet(wb, sel) {
  if (typeof sel === 'number') {
    // 序号与 readWorkbook、office_read 的 sheets 参数保持一致:从 1 开始
    const ws = wb.worksheets[sel - 1];
    if (ws) return ws;
    throw new OfficeError(`工作表序号 ${sel} 不存在（共 ${wb.worksheets.length} 个，序号从 1 开始）`);
  }
  const name = String(sel);
  const ws = wb.getWorksheet(name);
  if (!ws) throw new OfficeError(`工作表「${name}」不存在（现有: ${wb.worksheets.map((s) => s.name).join('、')}）`);
  return ws;
}

/** Build a fresh workbook buffer from a spec: { sheets: [{name, rows, header, columnWidths, rowHeights}] } */
export async function buildWorkbook(spec) {
  if (!spec || !Array.isArray(spec.sheets) || !spec.sheets.length) throw new OfficeError('至少需要一个工作表(sheets)');
  const wb = new ExcelJS.Workbook();
  wb.creator = 'dsh-office-toolkit';
  wb.created = new Date();
  spec.sheets.forEach((s, i) => {
    if (!isPlainObject(s)) throw new OfficeError(`sheets[${i}] 必须是对象`);
    const ws = wb.addWorksheet(String(s.name || `Sheet${i + 1}`).slice(0, 31));
    buildSheet(ws, s);
  });
  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

/** Apply edit ops to an existing workbook buffer. */
export async function editWorkbook(buf, ops) {
  if (!Array.isArray(ops) || !ops.length) throw new OfficeError('ops 至少需要一个操作');
  if (!isWithinBytes(buf, CAPS.MAX_EXCEL_INPUT_BYTES)) throw new OfficeError('文件超过 60 MB 上限', 'OFFICE_TOO_LARGE');
  const wb = await loadWorkbook(buf, '不是有效的 .xlsx 文件: ');

  const changes = [];
  const chartOps = [];
  for (const op of ops) {
    if (!isPlainObject(op)) throw new OfficeError('每个 op 必须是对象');
    const type = String(op.op || op.type || '');
    if (type === 'add_chart') { chartOps.push(op); continue; }
    changes.push(await runOp(wb, op, type));
  }
  let out = Buffer.from(await wb.xlsx.writeBuffer());
  for (const op of chartOps) out = await runChartOp(out, op, changes);
  return { buf: out, changes };
}

/** Run one op, honouring `optional: true` for non-fatal failures. */
async function runOp(wb, op, type) {
  try {
    await applyOp(wb, op, type);
    return type;
  } catch (err) {
    if (!op.optional) throw err;
    return `${type}(跳过: ${err.message})`;
  }
}

/** Charts need the finished OOXML package, so they run after writeBuffer(). */
async function runChartOp(out, op, changes) {
  const spec = chartSpec(op);
  try {
    const next = await injectChart(out, spec);
    changes.push(`add_chart(${spec.chartType})`);
    return next;
  } catch (err) {
    if (!op.optional) throw err;
    changes.push(`add_chart(跳过: ${err.message})`);
    return out;
  }
}

function chartSpec(op) {
  const ranges = Array.isArray(op.ranges) ? op.ranges : null;
  const series = ranges
    ? ranges.slice(1).map((r, i) => (op.seriesLabels?.[i] ? { range: r, label: op.seriesLabels[i] } : r))
    : op.series || [];
  return {
    sheet: op.sheet ?? 1,
    chartType: op.chartType || op.type || 'bar',
    title: op.title,
    categories: op.categories || (ranges ? ranges[0] : op.range),
    series,
    anchor: op.anchor || 'A1',
    widthPx: op.width,
    heightPx: op.height,
  };
}

function parseRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(String(ref));
  if (!m) throw new OfficeError(`无效单元格引用: ${ref}（如 A1）`);
  return { col: m[1], row: Number(m[2]) };
}

function cellAt(ws, ref) {
  const { col, row } = parseRef(ref);
  return ws.getCell(row, colLettersToIndex(col));
}

function colLettersToIndex(letters) {
  let n = 0;
  for (const ch of String(letters).toUpperCase()) n = n * 26 + (ch.codePointAt(0) - 64);
  return n;
}
// ---------------------------------------------------------------------------
// edit ops
// ---------------------------------------------------------------------------
function opSetCells(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  const m = /^([A-Z]+)(\d+)$/.exec(String(op.start || 'A1'));
  if (!m) throw new OfficeError(`set_cells 的 start 需为单元格引用（如 B2）: ${op.start}`);
  const c0 = colLettersToIndex(m[1]);
  const r0 = Number(m[2]);
  const values = op.values;
  if (!Array.isArray(values)) throw new OfficeError('set_cells 需要 values 二维数组');
  values.forEach((r, ri) => {
    if (!Array.isArray(r)) throw new OfficeError(`values[${ri}] 不是数组`);
    r.forEach((v, ci) => {
      if (v === null || v === undefined) return;
      const cell = ws.getCell(r0 + ri, c0 + ci);
      cell.value = coerceValue(isPlainObject(v) && 'value' in v ? v.value : v);
      if (isPlainObject(v) && v.style) applyCellStyle(cell, v.style);
    });
  });
  if (!op.style) return;
  for (let ri = 0; ri < values.length; ri++) {
    for (let ci = 0; ci < values[ri].length; ci++) applyCellStyle(ws.getCell(r0 + ri, c0 + ci), op.style);
  }
}

function opSetValue(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  const cell = cellAt(ws, String(op.ref || op.cell || 'A1'));
  cell.value = coerceValue(op.value);
  if (op.style) applyCellStyle(cell, op.style);
}

function opSetFormula(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  const cell = cellAt(ws, String(op.ref || op.cell || 'A1'));
  cell.value = { formula: String(op.formula).replace(/^=/, '') };
  if (op.style) applyCellStyle(cell, op.style);
}

function opStyleRange(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  const range = String(op.range || op.ref || '');
  if (!range) throw new OfficeError('style_range 需要 range（如 A1:C10）');
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
  if (!m) { applyCellStyle(cellAt(ws, range), op.style); return; }
  const [c1, r1, c2, r2] = [colLettersToIndex(m[1]), Number(m[2]), colLettersToIndex(m[3]), Number(m[4])];
  for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
    for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) applyCellStyle(ws.getCell(r, c), op.style);
  }
}

function opMerge(wb, op, type) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  const range = String(op.range || '');
  if (!range) throw new OfficeError(`${type} 需要 range（如 A1:C3）`);
  if (type === 'merge') ws.mergeCells(range);
  else ws.unMergeCells(range);
}

function opInsertRows(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  const at = Number(op.at ?? op.index ?? 1);
  ws.spliceRows(at, 0, ...Array.from({ length: Number(op.count ?? 1) }, () => []));
}

function opDeleteRows(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  ws.spliceRows(Number(op.at ?? op.index ?? 1), Number(op.count ?? 1));
}

function opInsertCols(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  const at = Number(op.at ?? op.index ?? 1);
  ws.spliceColumns(at, 0, ...Array.from({ length: Number(op.count ?? 1) }, () => []));
}

function opDeleteCols(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  ws.spliceColumns(Number(op.at ?? op.index ?? 1), Number(op.count ?? 1));
}

function opAddSheet(wb, op) {
  const name = String(op.name || `Sheet${wb.worksheets.length + 1}`).slice(0, 31);
  if (wb.getWorksheet(name)) throw new OfficeError(`工作表「${name}」已存在`);
  buildSheet(wb.addWorksheet(name), { rows: op.rows || [], header: op.header });
}

function opRenameSheet(wb, op) {
  getWorksheet(wb, op.sheet ?? 1).name = String(op.name).slice(0, 31);
}

function opDeleteSheet(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  if (wb.worksheets.length <= 1) throw new OfficeError('不能删除最后一个工作表');
  wb.removeWorksheet(ws.id);
}

function opSetColWidth(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  const col = typeof op.col === 'number' ? op.col : colLettersToIndex(String(op.col));
  ws.getColumn(col).width = Number(op.width);
}

function opSetRowHeight(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  ws.getRow(Number(op.row)).height = Number(op.height);
}

function opFreeze(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  ws.views = [{
    state: 'frozen',
    xSplit: Number(op.cols ?? 0) || undefined,
    ySplit: Number(op.rows ?? 0) || undefined,
    topLeftCell: 'A1',
  }];
}

function opAutoFilter(wb, op) {
  getWorksheet(wb, op.sheet ?? 1).autoFilter = String(op.range || '');
}

async function opAddImage(wb, op) {
  const ws = getWorksheet(wb, op.sheet ?? 1);
  const file = await readPath(String(op.path));
  const id = wb.addImage({
    buffer: file,
    extension: String(op.path).toLowerCase().endsWith('.png') ? 'png' : 'jpeg',
  });
  ws.addImage(id, imageAnchor(op));
}

function imageAnchor(op) {
  if (op.range) return { range: String(op.range) };
  if (!op.cell) throw new OfficeError('add_image 需要 range 或 cell');
  const { col, row } = parseRef(String(op.cell));
  const anchor = { tl: { col: colLettersToIndex(col) - 1, row: row - 1 } };
  if (op.width) anchor.ext = { width: Number(op.width), height: Number(op.height ?? op.width) };
  return anchor;
}

/** `op` name -> handler. Kept as data so no long switch is needed. */
const OP_HANDLERS = {
  set_cells: opSetCells,
  set_value: opSetValue,
  set_formula: opSetFormula,
  style_range: opStyleRange,
  merge: opMerge,
  unmerge: opMerge,
  insert_rows: opInsertRows,
  delete_rows: opDeleteRows,
  insert_cols: opInsertCols,
  delete_cols: opDeleteCols,
  add_sheet: opAddSheet,
  rename_sheet: opRenameSheet,
  delete_sheet: opDeleteSheet,
  set_col_width: opSetColWidth,
  set_row_height: opSetRowHeight,
  freeze: opFreeze,
  auto_filter: opAutoFilter,
  add_image: opAddImage,
};

async function applyOp(wb, op, type) {
  const handler = OP_HANDLERS[type];
  if (!handler) {
    throw new OfficeError(`未知操作类型: ${type}（支持 ${Object.keys(OP_HANDLERS).join('/')}）`);
  }
  return handler(wb, op, type);
}

export async function workbookSummary(buf) {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buf);
  } catch {
    return null;
  }
  return {
    sheets: wb.worksheets.map((ws) => ({ name: ws.name, rows: ws.rowCount, cols: ws.columnCount })),
    bytes: fmtBytes(buf.length),
  };
}
