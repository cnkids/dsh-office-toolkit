// .docx 表格:结构扫描、样式(框线/底纹/列宽/对齐)、增删行列与合并单元格。
//
// 只对表格所在的 XML 片段做手术,不动其它部件;所有扫描单向前进,不用正则,不存在回溯。
import { OfficeError } from './util.js';
import { attrValue, readTagAt } from './markup.js';
import { directChildren, elementInner, elementText, openAndInner, restyleRuns, upsertProps } from './docx-xml.js';
import { autoColumnPercents } from './docx-style.js';

const TBL = 'w:tbl';
const TR = 'w:tr';
const TC = 'w:tc';
const TBL_OPEN = `<${TBL}`;
const TR_OPEN = `<${TR}`;
const TBL_CLOSE = `</${TBL}>`;
const TR_CLOSE = `</${TR}>`;

/** w:tblPr 子元素的 schema 顺序。 */
const TBLPR_ORDER = ['w:tblStyle', 'w:tblpPr', 'w:tblOverlap', 'w:bidiVisual', 'w:tblStyleRowBandSize', 'w:tblStyleColBandSize', 'w:tblW', 'w:jc', 'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders', 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption', 'w:tblDescription', 'w:tblPrChange'];
/** w:trPr 子元素的 schema 顺序。 */
const TRPR_ORDER = ['w:cnfStyle', 'w:divId', 'w:gridBefore', 'w:gridAfter', 'w:wBefore', 'w:wAfter', 'w:cantSplit', 'w:trHeight', 'w:tblHeader', 'w:tblCellSpacing', 'w:jc', 'w:hidden', 'w:ins', 'w:del', 'w:trPrChange'];
/** w:tcPr 子元素的 schema 顺序。 */
const TCPR_ORDER = ['w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge', 'w:tcBorders', 'w:shd', 'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFitText', 'w:vAlign', 'w:hideMark', 'w:headers', 'w:cellIns', 'w:cellDel', 'w:cellMerge', 'w:tcPrChange'];
const OUTSIDE = ['top', 'left', 'bottom', 'right'];
const INSIDE = ['insideH', 'insideV'];
const DEFAULT_TEXT_WIDTH = 9071; // A4 210mm 减去默认 25mm 页边距 ≈ 160mm

// ---------------------------------------------------------------------------
// 扫描
// ---------------------------------------------------------------------------
function childRaw(inner, name) {
  const child = directChildren(inner).find((c) => c.name === name);
  return child ? inner.slice(child.start, child.end) : '';
}

function childAttr(inner, name, attr) {
  const child = directChildren(inner).find((c) => c.name === name);
  if (!child) return undefined;
  return attrValue(inner.slice(child.start, child.end), attr);
}

function cellSpan(cellXml) {
  const props = childRaw(cellXml, 'w:tcPr');
  const span = props ? childAttr(props, 'w:gridSpan', 'w:val') : undefined;
  const parsed = Number.parseInt(span ?? '1', 10);
  return Number.isFinite(parsed) && parsed > 1 ? parsed : 1;
}

function rowRaws(tableInner) {
  return directChildren(tableInner).filter((c) => c.name === TR).map((c) => tableInner.slice(c.start, c.end));
}

function cellRaws(rowXml) {
  const parts = openAndInner(rowXml, TR);
  if (!parts) return [];
  return directChildren(parts.inner).filter((c) => c.name === TC).map((c) => parts.inner.slice(c.start, c.end));
}

function gridColumns(tableInner, rows) {
  const gridInner = elementInner(childRaw(tableInner, 'w:tblGrid'), 'w:tblGrid');
  const declared = directChildren(gridInner).filter((c) => c.name === 'w:gridCol').length;
  const fromRows = rows.reduce((max, row) => Math.max(max, cellRaws(row).reduce((sum, cell) => sum + cellSpan(cell), 0)), 0);
  return Math.max(declared, fromRows, 1);
}

/**
 * 按文档顺序列出表格(嵌套表格也算,顺序按开标签位置)。
 * @returns {Array<{index: number, start: number, end: number, rows: number, cols: number}>}
 */
export function scanTables(xml) {
  const out = [];
  let i = 0;
  while (i < xml.length) {
    const at = xml.indexOf(TBL_OPEN, i);
    if (at === -1) break;
    const tag = readTagAt(xml, at);
    if (tag?.name !== TBL || tag.closing || tag.selfClosing) {
      i = at + 1;
      continue;
    }
    const closeAt = xml.indexOf(TBL_CLOSE, tag.end);
    const end = closeAt === -1 ? xml.length : closeAt + TBL_CLOSE.length;
    const inner = xml.slice(tag.end, end - TBL_CLOSE.length);
    const rows = rowRaws(inner);
    out.push({ index: out.length + 1, start: at, end, rows: rows.length, cols: gridColumns(inner, rows) });
    i = tag.end; // 继续往里扫,嵌套表格也能定位
  }
  return out;
}

/** 每列内容的最大字符数(跨行取最大;跨列合并的单元格不参与)。 */
function columnLengths(rows, cols) {
  const lengths = new Array(cols).fill(0);
  for (const row of rows) {
    let at = 0;
    for (const cell of cellRaws(row)) {
      const span = cellSpan(cell);
      if (span === 1 && at < cols) lengths[at] = Math.max(lengths[at], elementText(cell).trim().length);
      at += span;
    }
  }
  return lengths;
}

// ---------------------------------------------------------------------------
// 样式:框线 / 表头底纹 / 列宽 / 对齐
// ---------------------------------------------------------------------------
function borderTag(side, sz) {
  return `<w:${side} w:val="single" w:sz="${sz}" w:space="0" w:color="auto"/>`;
}

function borderOff(side) {
  return `<w:${side} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`;
}

/** 框线:`all` 全部 / `outline` 仅外框 / `none` 无 / `three-line` 三线表(上下粗线,中间线加在表头行)。 */
function bordersTag(mode) {
  if (mode === 'all') return `<w:tblBorders>${[...OUTSIDE, ...INSIDE].map((s) => borderTag(s, 4)).join('')}</w:tblBorders>`;
  if (mode === 'outline') return `<w:tblBorders>${OUTSIDE.map((s) => borderTag(s, 4)).join('')}${INSIDE.map(borderOff).join('')}</w:tblBorders>`;
  if (mode === 'three-line') return `<w:tblBorders>${['top', 'bottom'].map((s) => borderTag(s, 12)).join('')}${['left', 'right', ...INSIDE].map(borderOff).join('')}</w:tblBorders>`;
  return `<w:tblBorders>${[...OUTSIDE, ...INSIDE].map(borderOff).join('')}</w:tblBorders>`;
}

function lastTag(xml, name) {
  const at = xml.lastIndexOf(`<${name}`);
  if (at === -1) return undefined;
  const tag = readTagAt(xml, at);
  return tag ? xml.slice(at, tag.end) : undefined;
}

/** 表格可用正文宽度(twip):页面宽度减去左右页边距。 */
export function textWidthTwips(xml) {
  const width = Number(attrValue(lastTag(xml, 'w:pgSz') ?? '', 'w:w'));
  const left = Number(attrValue(lastTag(xml, 'w:pgMar') ?? '', 'w:left'));
  const right = Number(attrValue(lastTag(xml, 'w:pgMar') ?? '', 'w:right'));
  const usable = width - (Number.isFinite(left) ? left : 0) - (Number.isFinite(right) ? right : 0);
  return Number.isFinite(usable) && usable > 2000 ? usable : DEFAULT_TEXT_WIDTH;
}

/** 列宽百分比:manual 用给定比例,auto 按内容长度分配(谁内容长谁宽)。 */
function resolvePercents(rows, cols, norm) {
  if (!norm.columnMode) return null;
  if (norm.columnMode === 'manual') {
    const given = norm.columnPercents ?? [];
    return Array.from({ length: cols }, (_, i) => given[i] ?? Math.max(1, Math.round(100 / cols)));
  }
  return autoColumnPercents(columnLengths(rows, cols));
}

function gridTag(percents, textWidth) {
  const body = percents.map((p) => `<w:gridCol w:w="${Math.max(1, Math.round((p / 100) * textWidth))}"/>`).join('');
  return `<w:tblGrid>${body}</w:tblGrid>`;
}

/** 单元格内边距 → <w:tblCellMar>(schema 顺序 top/left/bottom/right)。 */
function cellMarTag(margins) {
  const body = ['top', 'left', 'bottom', 'right']
    .map((side) => `<w:${side} w:w="${margins[side] ?? 0}" w:type="dxa"/>`)
    .join('');
  return `<w:tblCellMar>${body}</w:tblCellMar>`;
}

function cellAdditions(cellXml, ctx) {
  const additions = {};
  const props = childRaw(cellXml, 'w:tcPr');
  if (ctx.percents) {
    const total = ctx.percents.slice(ctx.gridAt, ctx.gridAt + cellSpan(cellXml)).reduce((a, b) => a + b, 0);
    additions['w:tcW'] = `<w:tcW w:type="pct" w:w="${Math.round(Math.max(1, total) * 50)}"/>`;
  }
  if (!ctx.isHeader) return additions;
  if (ctx.norm.headerShading === null) {
    if (props && childRaw(props, 'w:shd')) additions['w:shd'] = '';
  } else if (ctx.norm.headerShading) {
    additions['w:shd'] = `<w:shd w:val="clear" w:color="auto" w:fill="${ctx.norm.headerShading}"/>`;
  }
  if (ctx.norm.borders === 'three-line') {
    additions['w:tcBorders'] = `<w:tcBorders>${borderTag('bottom', 6)}</w:tcBorders>`;
  }
  return additions;
}

/** 单元格垂直对齐(与是否表头无关,所有单元格都套)。 */
function cellAlignAdditions(cellXml, norm) {
  return norm.cellVerticalAlign ? { 'w:vAlign': `<w:vAlign w:val="${norm.cellVerticalAlign}"/>` } : {};
}

/** 行级属性:表头跨页重复、行高、禁止跨页断行。 */
function rowPrAdditions(isHeader, norm) {
  const additions = {};
  if (isHeader && norm.repeatHeader) additions['w:tblHeader'] = '<w:tblHeader/>';
  if (norm.cantSplit) additions['w:cantSplit'] = '<w:cantSplit/>';
  if (norm.rowHeightTwips) additions['w:trHeight'] = `<w:trHeight w:val="${norm.rowHeightTwips}" w:hRule="atLeast"/>`;
  return additions;
}

function styleCell(cellXml, ctx) {
  const additions = { ...cellAdditions(cellXml, ctx), ...cellAlignAdditions(cellXml, ctx.norm) };
  const bold = ctx.isHeader && ctx.norm.headerBold === true;
  if (!Object.keys(additions).length && !bold) return cellXml;
  const styled = Object.keys(additions).length ? upsertProps(cellXml, TC, 'w:tcPr', additions, TCPR_ORDER) : cellXml;
  return bold ? restyleRuns(styled, { 'w:b': '<w:b/>' }) : styled;
}

function styleRow(rowXml, ctx) {
  const parts = openAndInner(rowXml, TR);
  if (!parts) return rowXml;
  const rebuilt = directChildren(parts.inner).map((child) => {
    const raw = parts.inner.slice(child.start, child.end);
    if (child.name !== TC) return raw;
    const styled = styleCell(raw, ctx);
    ctx.gridAt += cellSpan(raw);
    return styled;
  });
  const row = `${parts.openTag}${rebuilt.join('')}${TR_CLOSE}`;
  const additions = rowPrAdditions(ctx.isHeader, ctx.norm);
  return Object.keys(additions).length ? upsertProps(row, TR, 'w:trPr', additions, TRPR_ORDER) : row;
}

/**
 * 给一个表格套样式。
 * @param {string} tableXml 表格 XML 片段
 * @param {object} norm normalizeTableSpec 的结果
 * @param {number} textWidth 可用正文宽度(twip)
 * @returns {string} 新的表格 XML
 */
export function applyTableStyle(tableXml, norm, textWidth) {
  const parts = openAndInner(tableXml, TBL);
  if (!parts) return tableXml;
  const rows = rowRaws(parts.inner);
  const percents = resolvePercents(rows, gridColumns(parts.inner, rows), norm);

  const prAdditions = {};
  if (norm.borders) prAdditions['w:tblBorders'] = bordersTag(norm.borders);
  if (norm.align) prAdditions['w:jc'] = `<w:jc w:val="${norm.align}"/>`;
  if (norm.cellMargins) prAdditions['w:tblCellMar'] = cellMarTag(norm.cellMargins);
  if (percents) {
    prAdditions['w:tblW'] = '<w:tblW w:type="pct" w:w="5000"/>';
    prAdditions['w:tblLayout'] = '<w:tblLayout w:type="fixed"/>';
  }

  let headerDone = false;
  const children = directChildren(parts.inner).map((child) => {
    const raw = parts.inner.slice(child.start, child.end);
    if (child.name === TR) {
      const isHeader = !headerDone;
      headerDone = true;
      return styleRow(raw, { norm, percents, isHeader, gridAt: 0 });
    }
    if (child.name === 'w:tblGrid' && percents) return gridTag(percents, textWidth);
    return raw;
  });

  const body = children.join('');
  const styled = Object.keys(prAdditions).length
    ? elementInner(upsertProps(`<${TBL}>${body}${TBL_CLOSE}`, TBL, 'w:tblPr', prAdditions, TBLPR_ORDER), TBL)
    : body;
  return `${parts.openTag}${styled}${TBL_CLOSE}`;
}

// ---------------------------------------------------------------------------
// 结构:增删行列 / 合并单元格
// ---------------------------------------------------------------------------
function tableParts(tableXml) {
  const parts = openAndInner(tableXml, TBL);
  if (!parts) return null;
  return {
    openTag: parts.openTag,
    children: directChildren(parts.inner).map((c) => ({ name: c.name, raw: parts.inner.slice(c.start, c.end) })),
  };
}

function joinParts(parts) {
  return `${parts.openTag}${parts.children.map((c) => c.raw).join('')}${TBL_CLOSE}`;
}

function rowSlots(children) {
  return children.map((c, i) => (c.name === TR ? i : -1)).filter((i) => i >= 0);
}

function emptyCell() {
  return `<w:tc><w:tcPr><w:tcW w:type="auto" w:w="0"/></w:tcPr><w:p/></w:tc>`;
}

function emptyRow(cols) {
  return `<w:tr>${Array.from({ length: cols }, emptyCell).join('')}${TR_CLOSE}`;
}

function requireCount(value, what) {
  if (value === undefined || value === null || value === '') return 1;
  const n = Number(value);
  if (!Number.isInteger(n) || n < 1 || n > 100) throw new OfficeError(`${what} 需为 1–100 的整数`, 'INVALID_ARGS');
  return n;
}

/** `at` 是 1 起的插入位置(插在它之前);省略或超出则追加。 */
function insertSlot(value, total, what) {
  if (value === undefined || value === null || value === '') return total;
  const n = Number(value);
  if (!Number.isInteger(n) || n < 1 || n > total + 1) {
    throw new OfficeError(`${what} 需为 1–${total + 1} 的整数`, 'INVALID_ARGS');
  }
  return n - 1;
}

function requireAt(value, total, what) {
  if (total === 0) {
    throw new OfficeError(`${what}: 这个表格里已经没有行/列了（删空后的残留）。要整张删掉请用 delete_table`, 'INVALID_ARGS');
  }
  const n = Number(value);
  if (!Number.isInteger(n) || n < 1 || n > total) {
    throw new OfficeError(`${what} 需为 1–${total} 的整数（本表共 ${total} 行/列）`, 'INVALID_ARGS');
  }
  return n - 1;
}

/** 第 position 个数据行之前对应的 children 下标;position 越界表示追加到末尾。 */
function rowInsertIndex(children, slots, position) {
  if (position < slots.length) return slots[position];
  if (!slots.length) return children.length;
  return slots[slots.length - 1] + 1;
}

export function insertTableRows(tableXml, at, count) {
  const parts = tableParts(tableXml);
  if (!parts) return tableXml;
  const slots = rowSlots(parts.children);
  const cols = gridColumns(elementInner(tableXml, TBL), slots.map((i) => parts.children[i].raw));
  const position = insertSlot(at, slots.length, 'insert_table_row 的 at');
  const insertAt = rowInsertIndex(parts.children, slots, position);
  const added = Array.from({ length: requireCount(count, 'insert_table_row 的 count') }, () => ({ name: TR, raw: emptyRow(cols) }));
  parts.children.splice(insertAt, 0, ...added);
  return joinParts(parts);
}

export function deleteTableRows(tableXml, at, count) {
  const parts = tableParts(tableXml);
  if (!parts) return tableXml;
  const slots = rowSlots(parts.children);
  const from = requireAt(at, slots.length, 'delete_table_row 的 at');
  const victims = new Set(slots.slice(from, from + requireCount(count, 'delete_table_row 的 count')));
  parts.children = parts.children.filter((_, i) => !victims.has(i));
  return joinParts(parts);
}

function insertGridCols(gridXml, position, count) {
  const parts = openAndInner(gridXml, 'w:tblGrid');
  if (!parts) return gridXml;
  const cols = directChildren(parts.inner).map((c) => ({ name: c.name, raw: parts.inner.slice(c.start, c.end) }));
  cols.splice(Math.min(position, cols.length), 0, ...Array.from({ length: count }, () => ({ name: 'w:gridCol', raw: '<w:gridCol w:w="1440"/>' })));
  return `${parts.openTag}${cols.map((c) => c.raw).join('')}</w:tblGrid>`;
}

function insertRowCells(rowXml, position, count) {
  const parts = openAndInner(rowXml, TR);
  if (!parts) return rowXml;
  const cells = directChildren(parts.inner).map((c) => ({ name: c.name, raw: parts.inner.slice(c.start, c.end) }));
  cells.splice(Math.min(position, cells.length), 0, ...Array.from({ length: count }, () => ({ name: TC, raw: emptyCell() })));
  return `${parts.openTag}${cells.map((c) => c.raw).join('')}${TR_CLOSE}`;
}

export function insertTableColumns(tableXml, at, count) {
  const parts = tableParts(tableXml);
  if (!parts) return tableXml;
  const firstRow = parts.children.find((c) => c.name === TR);
  const cols = firstRow ? cellRaws(firstRow.raw).length : 1;
  const position = insertSlot(at, cols, 'insert_table_column 的 at');
  const many = requireCount(count, 'insert_table_column 的 count');
  parts.children = parts.children.map((child) => {
    if (child.name === 'w:tblGrid') return { name: child.name, raw: insertGridCols(child.raw, position, many) };
    if (child.name !== TR) return child;
    return { name: child.name, raw: insertRowCells(child.raw, position, many) };
  });
  return joinParts(parts);
}

export function deleteTableColumns(tableXml, at, count) {
  const parts = tableParts(tableXml);
  if (!parts) return tableXml;
  const firstRow = parts.children.find((c) => c.name === TR);
  const cols = firstRow ? cellRaws(firstRow.raw).length : 0;
  const from = requireAt(at, cols, 'delete_table_column 的 at');
  const victims = new Set(Array.from({ length: requireCount(count, 'delete_table_column 的 count') }, (_, i) => from + i));
  parts.children = parts.children.map((child) => {
    if (child.name === 'w:tblGrid') return { name: child.name, raw: keepGridCols(child.raw, victims) };
    if (child.name !== TR) return child;
    return { name: child.name, raw: keepRowCells(child.raw, victims) };
  });
  return joinParts(parts);
}

function keepGridCols(gridXml, victims) {
  const parts = openAndInner(gridXml, 'w:tblGrid');
  if (!parts) return gridXml;
  const kept = directChildren(parts.inner)
    .map((c, i) => ({ name: c.name, raw: parts.inner.slice(c.start, c.end), keep: !victims.has(i) }))
    .filter((c) => c.keep);
  return `${parts.openTag}${kept.map((c) => c.raw).join('')}</w:tblGrid>`;
}

function keepRowCells(rowXml, victims) {
  const parts = openAndInner(rowXml, TR);
  if (!parts) return rowXml;
  const kept = directChildren(parts.inner)
    .map((c, i) => ({ name: c.name, raw: parts.inner.slice(c.start, c.end), keep: !victims.has(i) }))
    .filter((c) => c.keep);
  if (!kept.length) kept.push({ name: TC, raw: emptyCell() });
  return `${parts.openTag}${kept.map((c) => c.raw).join('')}${TR_CLOSE}`;
}

function parseCellRef(raw) {
  const text = String(raw ?? '').trim().toUpperCase();
  let i = 0;
  let col = 0;
  while (i < text.length && text[i] >= 'A' && text[i] <= 'Z') {
    col = col * 26 + (text.codePointAt(i) - 64);
    i += 1;
  }
  const row = Number.parseInt(text.slice(i), 10);
  if (col < 1 || !Number.isInteger(row) || row < 1) throw new OfficeError(`单元格引用 "${raw}" 不合法`, 'INVALID_ARGS');
  return { col, row };
}

/** 解析 `A1:B2`(行列互换也接受)。 */
export function parseCellRange(ref) {
  const text = String(ref ?? '').trim();
  const at = text.indexOf(':');
  if (at <= 0) throw new OfficeError('range 需形如 "A1:B2"', 'INVALID_ARGS');
  const from = parseCellRef(text.slice(0, at));
  const to = parseCellRef(text.slice(at + 1));
  return {
    r1: Math.min(from.row, to.row),
    r2: Math.max(from.row, to.row),
    c1: Math.min(from.col, to.col),
    c2: Math.max(from.col, to.col),
  };
}

/** 取消合并:把 gridSpan/vMerge 去掉,并把占位补回成独立空单元格。 */
export function unmergeTableCells(tableXml, range) {
  const parts = tableParts(tableXml);
  if (!parts) return tableXml;
  const { r1, c1, r2, c2 } = range;
  let seen = 0;
  let touched = 0;
  parts.children = parts.children.map((child) => {
    if (child.name !== TR) return child;
    seen += 1;
    if (seen < r1 || seen > r2) return child;
    const result = unmergeRow(child.raw, { col: c1, width: c2 - c1 + 1 });
    touched += result.count;
    return { name: TR, raw: result.xml };
  });
  if (!touched) throw new OfficeError(`unmerge_table_cells: ${r1 === r2 && c1 === c2 ? '该单元格' : '该区域'}没有合并单元格`, 'TABLE_NOT_MERGED');
  return joinParts(parts);
}

function unmergeRow(rowXml, ctx) {
  const parts = openAndInner(rowXml, TR);
  if (!parts) return { xml: rowXml, count: 0 };
  const cells = directChildren(parts.inner).map((c) => parts.inner.slice(c.start, c.end));
  let gridAt = 0;
  let count = 0;
  const out = [];
  for (const cell of cells) {
    const span = cellSpan(cell);
    const inRange = gridAt >= ctx.col - 1 && gridAt < ctx.col - 1 + ctx.width;
    if (!inRange) {
      out.push(cell);
      gridAt += span;
      continue;
    }
    const merged = span > 1 || /<w:vMerge/.test(cell);
    if (!merged) {
      out.push(cell);
      gridAt += span;
      continue;
    }
    count += 1;
    out.push(upsertProps(cell, TC, 'w:tcPr', { 'w:gridSpan': '', 'w:vMerge': '' }, TCPR_ORDER));
    for (let k = 1; k < span; k += 1) out.push(emptyCell());
    gridAt += span;
  }
  if (!count) return { xml: rowXml, count: 0 };
  return { xml: `${parts.openTag}${out.join('')}${TR_CLOSE}`, count };
}

/** 合并矩形区域(含跨行跨列);保留左上角单元格的内容。 */
export function mergeTableCells(tableXml, range) {
  const parts = tableParts(tableXml);
  if (!parts) return tableXml;
  const rows = rowSlots(parts.children).length;
  const { r1, c1, r2, c2 } = range;
  if (r2 > rows) throw new OfficeError(`merge_table_cells: 第 ${r2} 行超出范围：本表只有 ${rows} 行`, 'INVALID_ARGS');
  let seen = 0;
  parts.children = parts.children.map((child) => {
    if (child.name !== TR) return child;
    seen += 1;
    if (seen < r1 || seen > r2) return child;
    return { name: TR, raw: mergeRow(child.raw, seen, { c1, c2, width: c2 - c1 + 1, top: seen === r1, tall: r2 > r1 }) };
  });
  return joinParts(parts);
}

function mergeRow(rowXml, rowIndex, ctx) {
  const parts = openAndInner(rowXml, TR);
  if (!parts) return rowXml;
  const cells = directChildren(parts.inner).map((c) => parts.inner.slice(c.start, c.end));
  if (ctx.c2 > cells.length) {
    throw new OfficeError(`merge_table_cells: 第 ${rowIndex} 行只有 ${cells.length} 列，取不到第 ${ctx.c2} 列`, 'INVALID_ARGS');
  }
  const kept = [];
  for (let i = 0; i < cells.length; i++) {
    if (i < ctx.c1 - 1 || i > ctx.c2 - 1) {
      kept.push(cells[i]);
      continue;
    }
    if (i > ctx.c1 - 1) continue; // 被合并进去的单元格
    const additions = {};
    if (ctx.width > 1) additions['w:gridSpan'] = `<w:gridSpan w:val="${ctx.width}"/>`;
    if (ctx.tall) additions['w:vMerge'] = ctx.top ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>';
    kept.push(upsertProps(cells[i], TC, 'w:tcPr', additions, TCPR_ORDER));
  }
  return `${parts.openTag}${kept.map((c) => c).join('')}${TR_CLOSE}`;
}
