// 字体 / 行距 / 缩进等排版的单位换算与规格解析。
//
// 两条入口共用同一套规范形式,免得"写新文档"和"改已有文档"对同一个参数解释不一致:
//   - docx-writer.js(HTML→docx)把规范形式映射成 docx@9 的选项
//   - docx-edit.js(就地改 .docx)把规范形式映射成 OOXML 片段
// 规范形式里单位已经统一:字号用半磅(w:sz),长度用 twip(1pt = 20twip)。
import { OfficeError, isPlainObject } from './util.js';

const PT_TO_TWIPS = 20;
const HALF_POINTS_PER_PT = 2;
const LINE_AUTO_PER_UNIT = 240; // w:lineRule="auto" 时 w:line 的单位是 1/240 行

/** 对齐:对外统一 both/center/left/right,对内就是 OOXML 的 w:jc 值。 */
const ALIGNMENTS = new Set(['both', 'center', 'left', 'right']);
const ALIGN_ALIASES = { justify: 'both', justified: 'both', start: 'left', end: 'right' };

function num(value) {
  const n = Number.parseFloat(String(value ?? '').trim());
  return Number.isFinite(n) ? n : null;
}

/** 磅 → 半磅(docx 的 w:sz 单位)。 */
export function ptToHalfPoints(pt) {
  return Math.max(1, Math.round(pt * HALF_POINTS_PER_PT));
}

/** CSS 长度单位(只认这几种,其余当无单位处理)。 */
function unitOf(raw) {
  for (const unit of ['em', 'ch', 'pt', 'px']) {
    if (raw.endsWith(unit)) return unit;
  }
  return '';
}

/** 长度 → twip。支持 pt / px(按 96dpi)/ em / ch(按字号)/ 无单位(当 pt)。 */
export function lengthToTwips(value, emHalfPoints) {
  const raw = String(value ?? '').trim().toLowerCase();
  if (!raw) return null;
  const unit = unitOf(raw);
  const n = num(unit ? raw.slice(0, -unit.length) : raw);
  if (n === null) return null;
  if (unit === 'em' || unit === 'ch') {
    const emPt = (emHalfPoints ?? ptToHalfPoints(10.5)) / HALF_POINTS_PER_PT;
    return Math.round(n * emPt * PT_TO_TWIPS);
  }
  if (unit === 'px') return Math.round(n * 0.75 * PT_TO_TWIPS);
  return Math.round(n * PT_TO_TWIPS);
}

/** `style="a:b;c:d"` → `{a:'b', c:'d'}`(键小写)。 */
export function declarationsFrom(rawStyle) {
  const out = Object.create(null);
  for (const decl of String(rawStyle ?? '').split(';')) {
    const at = decl.indexOf(':');
    if (at > 0) out[decl.slice(0, at).trim().toLowerCase()] = decl.slice(at + 1).trim();
  }
  return out;
}

/** `#abc` / `abc` / `#AABBCC` → `aabbcc`,否则 undefined。 */
export function hexColor(value) {
  const raw = String(value ?? '').trim().replace('#', '').toLowerCase();
  const expand = (s) => s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
  if (raw.length === 3 && /^[0-9a-f]{3}$/.test(raw)) return expand(raw);
  return /^[0-9a-f]{6}$/.test(raw) ? raw : undefined;
}

/**
 * CSS 字体族 → `{ font, fontAscii }`。
 * 约定:`font-family: 中文字体, 西文字体` —— 第一个当作东亚字体(必须写进 w:eastAsia,
 * 否则 Word 会用默认中文字体渲染汉字),第二个(若有)当作西文字体。
 */
function fontFamilies(value) {
  const list = String(value ?? '').split(',').map((f) => f.trim().replace(/^['"]/, '').replace(/['"]$/, '')).filter(Boolean);
  return { font: list[0], fontAscii: list[1] };
}

function isBoldWeight(value) {
  const raw = String(value ?? '').trim().toLowerCase();
  if (raw === 'bold' || raw === 'bolder') return true;
  if (raw === 'normal' || raw === 'lighter') return false;
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) ? n >= 600 : undefined;
}

/** 内联 style 里的字符级属性。 */
export function runStyleFromCss(decls) {
  const out = {};
  const { font, fontAscii } = fontFamilies(decls['font-family']);
  if (font) out.font = font;
  if (fontAscii) out.fontAscii = fontAscii;
  const size = lengthToTwips(decls['font-size'], null);
  if (size) out.sizeHalfPt = Math.max(1, Math.round((size / PT_TO_TWIPS) * HALF_POINTS_PER_PT));
  const bold = isBoldWeight(decls['font-weight']);
  if (bold !== undefined) out.bold = bold;
  if (String(decls['font-style'] ?? '').toLowerCase() === 'italic') out.italics = true;
  const decoration = String(decls['text-decoration'] ?? '').toLowerCase();
  if (decoration.includes('underline')) out.underline = true;
  if (decoration.includes('line-through')) out.strike = true;
  const color = hexColor(decls.color);
  if (color) out.color = color;
  return out;
}

/** 取第一个已定义的值(用于 margin 简写与 margin-top/bottom 的优先级)。 */
function firstDefined(primary, fallback) {
  if (primary === undefined) return fallback;
  return primary;
}

/** 段前/段后:支持 `margin` 简写(1/2/3/4 值的 CSS 语义)。 */
function marginsFrom(decls, emHalfPoints) {
  const out = {};
  const shorthand = String(decls.margin ?? '').trim().split(/\s+/).filter(Boolean);
  const shortBottom = shorthand.length > 2 ? shorthand[2] : shorthand[0];
  const before = lengthToTwips(firstDefined(decls['margin-top'], shorthand[0]), emHalfPoints);
  const bottom = firstDefined(decls['margin-bottom'], shortBottom);
  const after = lengthToTwips(bottom, emHalfPoints);
  if (before !== null) out.beforeTwips = before;
  if (after !== null) out.afterTwips = after;
  return out;
}

/**
 * 无单位数字(纯倍数)才返回数字,带单位返回 null。
 * 注意 `\d*\.?\d+` 是重叠量词(`\.?` 可匹配空串),超长数字串上会退化成 O(n²);
 * 这里两个分支首字符互斥,恒为线性。
 */
const UNITLESS_NUMBER = /^(?:\d+(?:\.\d+)?|\.\d+)$/;

function unitlessMultiple(raw) {
  return UNITLESS_NUMBER.test(raw) ? Number.parseFloat(raw) : null;
}

/** CSS 行高:无单位数字是倍数(auto),带单位是固定值(exact)。 */
function lineFrom(decls, emHalfPoints) {
  const raw = String(decls['line-height'] ?? '').trim().toLowerCase();
  if (!raw || raw === 'normal') return {};
  const multiple = unitlessMultiple(raw);
  if (multiple !== null) {
    return multiple > 0 ? { line: Math.round(multiple * LINE_AUTO_PER_UNIT), lineRule: 'auto' } : {};
  }
  const twips = lengthToTwips(raw, emHalfPoints);
  return twips && twips > 0 ? { line: twips, lineRule: 'exact' } : {};
}

/** 内联 style 里的段落级属性。emHalfPoints 用于把 em/ch 解析成 twip。 */
export function paragraphStyleFromCss(decls, emHalfPoints) {
  const out = { ...marginsFrom(decls, emHalfPoints), ...lineFrom(decls, emHalfPoints) };
  const align = alignmentOf(decls['text-align']);
  if (align) out.align = align;
  const indent = lengthToTwips(decls['text-indent'], emHalfPoints);
  if (indent !== null) out.firstLineTwips = indent;
  return out;
}

/** `text-align` 的 CSS 值 → both/center/left/right。 */
export function alignmentOf(value) {
  const raw = String(value ?? '').trim().toLowerCase();
  if (!raw) return undefined;
  if (ALIGNMENTS.has(raw)) return raw;
  return ALIGN_ALIASES[raw];
}

// ---------------------------------------------------------------------------
// 工具参数里的 style 规格:校验 + 换算成规范形式
// ---------------------------------------------------------------------------
function requireNumber(value, what, { min = 0, max = Infinity } = {}) {
  const n = num(value);
  if (n === null || n < min || n > max) {
    const range = max === Infinity ? `${min} 以上` : `${min}–${max}`;
    const got = JSON.stringify(value);
    throw new OfficeError(`${what} 需为 ${range} 的数字，收到 ${got}`, 'INVALID_ARGS');
  }
  return n;
}

function requireString(value, what) {
  const s = String(value ?? '').trim();
  if (!s) throw new OfficeError(`${what} 不能为空`, 'INVALID_ARGS');
  if (s.length > 64) throw new OfficeError(`${what} 过长（${s.length} 字符）`, 'INVALID_ARGS');
  return s;
}

/**
 * `office_write_docx` / `office_edit_docx` 的 style 参数 → 规范形式。
 * 全部可选;字号用磅(三号=16、四号=14、小四=12),行距用磅(固定值)或倍数。
 * @returns {object} 已换算成半磅 / twip 的规范形式
 */
/** 字符级字段:字体 / 字号 / 加粗 / 字色。 */
function assignRunFields(out, spec) {
  if (spec.font !== undefined) out.font = requireString(spec.font, 'style.font');
  if (spec.fontAscii !== undefined) out.fontAscii = requireString(spec.fontAscii, 'style.fontAscii');
  if (spec.sizePt !== undefined) out.sizeHalfPt = ptToHalfPoints(requireNumber(spec.sizePt, 'style.sizePt', { min: 1, max: 200 }));
  if (spec.bold !== undefined) out.bold = Boolean(spec.bold);
  if (spec.color === undefined) return;
  const color = hexColor(spec.color);
  if (!color) throw new OfficeError(`style.color 需为 RRGGBB，收到 ${JSON.stringify(spec.color)}`, 'INVALID_ARGS');
  out.color = color;
}

/** 缩进字段:字符数或磅值。 */
function assignIndent(out, spec) {
  if (spec.firstLineIndentChars !== undefined) {
    out.firstLineTwips = indentCharsToTwips(spec.firstLineIndentChars, out.sizeHalfPt);
  }
  if (spec.firstLineIndentPt !== undefined) {
    const pt = requireNumber(spec.firstLineIndentPt, 'style.firstLineIndentPt');
    out.firstLineTwips = lengthToTwips(`${pt}pt`, out.sizeHalfPt);
  }
}

function assignAlign(out, spec) {
  if (spec.align === undefined) return;
  const align = alignmentOf(spec.align);
  if (!align) throw new OfficeError('style.align 可为 both/center/left/right', 'INVALID_ARGS');
  out.align = align;
}

export function normalizeStyleSpec(spec, { allowHeadings = true, allowTable = true } = {}) {
  if (spec === undefined || spec === null) return {};
  if (!isPlainObject(spec)) throw new OfficeError('style 需为对象，如 {"font":"仿宋_GB2312","sizePt":16,"lineSpacingPt":28.8}', 'INVALID_ARGS');
  const out = {};
  if (spec.headings !== undefined) {
    if (!allowHeadings) throw new OfficeError('style.headings 只能用于 office_write_docx（整篇默认样式）；改单段/全篇段落请直接给具体属性', 'INVALID_ARGS');
    out.headings = normalizeStyleSpec(spec.headings, { allowHeadings: false });
  }
  if (spec.table !== undefined) {
    if (!allowTable) throw new OfficeError('style.table 只能用于 office_write_docx（写新文档）；改已有文档的表格请用 set_table 操作', 'INVALID_ARGS');
    out.table = normalizeTableSpec(spec.table);
  }
  assignRunFields(out, spec);
  Object.assign(out, lineSpec(spec), spacingSpec(spec));
  assignIndent(out, spec);
  assignAlign(out, spec);
  return out;
}

/** 行距:固定值磅 lineSpacingPt 与倍数 lineSpacingMultiple 二选一。 */
function lineSpec(spec) {
  const hasPt = spec.lineSpacingPt !== undefined && spec.lineSpacingPt !== null;
  const hasMultiple = spec.lineSpacingMultiple !== undefined && spec.lineSpacingMultiple !== null;
  if (hasPt && hasMultiple) throw new OfficeError('lineSpacingPt 与 lineSpacingMultiple 只能给一个', 'INVALID_ARGS');
  if (hasPt) {
    const pt = requireNumber(spec.lineSpacingPt, 'style.lineSpacingPt', { min: 1, max: 500 });
    return { line: Math.round(pt * PT_TO_TWIPS), lineRule: 'exact' };
  }
  if (hasMultiple) {
    const multiple = requireNumber(spec.lineSpacingMultiple, 'style.lineSpacingMultiple', { min: 0.5, max: 10 });
    return { line: Math.round(multiple * LINE_AUTO_PER_UNIT), lineRule: 'auto' };
  }
  return {};
}

function spacingSpec(spec) {
  const out = {};
  if (spec.spacingBeforePt !== undefined) out.beforeTwips = Math.round(requireNumber(spec.spacingBeforePt, 'style.spacingBeforePt', { max: 500 }) * PT_TO_TWIPS);
  if (spec.spacingAfterPt !== undefined) out.afterTwips = Math.round(requireNumber(spec.spacingAfterPt, 'style.spacingAfterPt', { max: 500 }) * PT_TO_TWIPS);
  return out;
}

/** 首行缩进字符数 → twip(需要字号;没给字号时按小四 12pt)。 */
export function indentCharsToTwips(chars, sizeHalfPt) {
  const n = requireNumber(chars, 'style.firstLineIndentChars', { max: 20 });
  const halfPt = sizeHalfPt ?? ptToHalfPoints(12);
  return Math.round(n * (halfPt / HALF_POINTS_PER_PT) * PT_TO_TWIPS);
}

// ---------------------------------------------------------------------------
// 表格规格
// ---------------------------------------------------------------------------
export const TABLE_BORDERS = ['all', 'none', 'outline', 'three-line'];
/** 单元格垂直对齐。 */
export const CELL_ALIGNS = ['top', 'center', 'bottom'];

/**
 * 按内容长度分配列宽（谁内容长谁宽）。
 * 至少给每列 4%、最多 70%，再归一化到 100%，避免超长的一列把其他列挤没。
 * @param {number[]} lengths 每列内容的最大长度（字符数）
 * @returns {number[]} 百分比整数数组，和为 100
 */
export function autoColumnPercents(lengths) {
  const weights = (lengths.length ? lengths : [1]).map((n) => Math.max(1, Math.min(Number(n) || 0, 500)));
  const total = weights.reduce((a, b) => a + b, 0);
  const raw = weights.map((w) => (w / total) * 100);
  const clamped = raw.map((p) => Math.min(70, Math.max(4, p)));
  const sum = clamped.reduce((a, b) => a + b, 0);
  return clamped.map((p) => Math.max(1, Math.round((p / sum) * 100)));
}

function requireBoolean(value, what) {
  if (typeof value !== 'boolean') throw new OfficeError(`${what} 需为 true / false`, 'INVALID_ARGS');
  return value;
}

function columnPercentsOf(value) {
  if (!Array.isArray(value) || !value.length || value.length > 64) {
    throw new OfficeError('style.table.columnWidths 需为 1–64 个百分比的数组，如 [30, 40, 30]', 'INVALID_ARGS');
  }
  const numbers = value.map((v) => requireNumber(v, 'style.table.columnWidths 的每一项', { min: 0.1, max: 100 }));
  const sum = numbers.reduce((a, b) => a + b, 0);
  if (sum <= 0) throw new OfficeError('style.table.columnWidths 之和需大于 0', 'INVALID_ARGS');
  // 用户写 [3,4,3] 或 [30,40,30] 都按比例理解
  return numbers.map((n) => Math.max(1, Math.round((n / sum) * 100)));
}

/**
 * 表格规格 → 规范形式（写新文档与改已有文档共用）。
 * @returns {object} 已换算好的表格样式
 */
function assignBorders(out, spec) {
  if (spec.borders === undefined) return;
  const borders = String(spec.borders).trim();
  if (!TABLE_BORDERS.includes(borders)) {
    throw new OfficeError(`style.table.borders 可为 ${TABLE_BORDERS.join(' / ')}`, 'INVALID_ARGS');
  }
  out.borders = borders;
}

function assignHeaderShading(out, spec) {
  if (spec.headerShading === undefined) return;
  const raw = spec.headerShading;
  if (raw === null || raw === 'none' || raw === '') {
    out.headerShading = null; // 明确表示"去掉底纹"
    return;
  }
  const fill = hexColor(raw);
  if (!fill) throw new OfficeError(`style.table.headerShading 需为 RRGGBB 或 "none"，收到 ${JSON.stringify(raw)}`, 'INVALID_ARGS');
  out.headerShading = fill;
}

function assignTableAlign(out, spec) {
  if (spec.align === undefined) return;
  const align = alignmentOf(spec.align);
  if (!align) throw new OfficeError('style.table.align 可为 left/center/right', 'INVALID_ARGS');
  out.align = align;
}

/** Word 默认的单元格内边距(twip):左右 108,上下 0。 */
const DEFAULT_CELL_MARGINS = { top: 0, bottom: 0, left: 108, right: 108 };

/** 单元格内边距:没给的边用 Word 默认值补齐,保证写出的 tblCellMar 四条边齐全。 */
function assignCellMargins(out, spec) {
  if (spec.cellMargins === undefined) return;
  const raw = spec.cellMargins;
  if (!isPlainObject(raw)) {
    throw new OfficeError('style.table.cellMargins 需为对象，如 {"left":108,"right":108,"top":40,"bottom":40}', 'INVALID_ARGS');
  }
  const margins = { ...DEFAULT_CELL_MARGINS };
  let given = 0;
  for (const side of ['top', 'bottom', 'left', 'right']) {
    if (raw[side] === undefined) continue;
    margins[side] = Math.round(requireNumber(raw[side], `style.table.cellMargins.${side}`, { max: 5000 }));
    given += 1;
  }
  if (!given) throw new OfficeError('style.table.cellMargins 至少要给 top / bottom / left / right 之一', 'INVALID_ARGS');
  out.cellMargins = margins;
}

function assignCellAlign(out, spec) {
  if (spec.cellVerticalAlign === undefined) return;
  const align = String(spec.cellVerticalAlign).trim().toLowerCase();
  if (!CELL_ALIGNS.includes(align)) {
    throw new OfficeError(`style.table.cellVerticalAlign 可为 ${CELL_ALIGNS.join(' / ')}`, 'INVALID_ARGS');
  }
  out.cellVerticalAlign = align;
}

function assignRowHeight(out, spec) {
  if (spec.rowHeightPt === undefined) return;
  const pt = requireNumber(spec.rowHeightPt, 'style.table.rowHeightPt', { max: 2000 });
  out.rowHeightTwips = Math.round(pt * PT_TO_TWIPS);
}

function assignColumnMode(out, spec) {
  if (spec.columnWidthMode !== undefined) {
    const mode = String(spec.columnWidthMode).trim();
    if (mode !== 'auto' && mode !== 'manual') {
      throw new OfficeError('style.table.columnWidthMode 可为 auto（按内容自动分配）或 manual（配合 columnWidths）', 'INVALID_ARGS');
    }
    out.columnMode = mode;
  }
  if (spec.columnWidths !== undefined) {
    out.columnPercents = columnPercentsOf(spec.columnWidths);
    out.columnMode = out.columnMode ?? 'manual';
  }
  if (out.columnMode === 'manual' && !out.columnPercents) {
    throw new OfficeError('style.table.columnWidthMode 为 manual 时需要同时给 columnWidths', 'INVALID_ARGS');
  }
}

/**
 * 表格规格 → 规范形式（写新文档与改已有文档共用）。
 * @returns {object} 已换算好的表格样式
 */
export function normalizeTableSpec(spec) {
  if (spec === undefined || spec === null) return {};
  if (!isPlainObject(spec)) throw new OfficeError('style.table 需为对象，如 {"borders":"three-line","columnWidthMode":"auto"}', 'INVALID_ARGS');
  const out = {};
  assignBorders(out, spec);
  assignHeaderShading(out, spec);
  if (spec.headerBold !== undefined) out.headerBold = requireBoolean(spec.headerBold, 'style.table.headerBold');
  assignTableAlign(out, spec);
  assignCellMargins(out, spec);
  assignCellAlign(out, spec);
  assignRowHeight(out, spec);
  if (spec.repeatHeader !== undefined) out.repeatHeader = requireBoolean(spec.repeatHeader, 'style.table.repeatHeader');
  if (spec.cantSplit !== undefined) out.cantSplit = requireBoolean(spec.cantSplit, 'style.table.cantSplit');
  assignColumnMode(out, spec);
  return out;
}

/** 是否为"空规格"(一个字段都没有)。 */
export function isEmptyStyle(norm) {
  return !norm || Object.keys(norm).length === 0;
}

export { PT_TO_TWIPS, HALF_POINTS_PER_PT };
