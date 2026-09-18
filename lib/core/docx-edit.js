// .docx 局部修改:只对主文档部件做段落级手术,其余部件(图片/页眉页脚/样式/批注/关系)原样保留。
//
// 与 office_write_docx 的区别:那个是整篇重建(原格式全丢),这里是就地改 —— 改一段只动那一段的
// XML。所有扫描都是单向前进(见 markup.js 的说明),不用正则,不存在回溯。
import { OfficeError, isWithinBytes, isPlainObject, CAPS, assertZipBudget } from './util.js';
import { openOoxml, partOf } from './ooxml.js';
import { attrIn, readTagAt } from './markup.js';
import { RPR_ORDER, attrText, directChildren, rebuildChildren, restyleRuns, upsertProps } from './docx-xml.js';
import { normalizeStyleSpec, normalizeTableSpec, isEmptyStyle } from './docx-style.js';
import {
  applyTableStyle, deleteTableColumns, deleteTableRows, insertTableColumns, insertTableRows,
  mergeTableCells, parseCellRange, scanTables, textWidthTwips, unmergeTableCells,
} from './docx-table.js';
import { ensureNumbering, requirePreset } from './docx-numbering.js';

const MAIN_PART = 'word/document.xml';
const MAX_OPS = 200;
const MAX_ENTITY_LEN = 12;
const PARAGRAPH = 'w:p';
const TEXT = 'w:t';
const TABLE = 'w:tbl';

const NAMED_ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

// ---------------------------------------------------------------------------
// XML 文本编解码:匹配与替换都在"显示文本"上做,写回时只转义必须转义的字符
// ---------------------------------------------------------------------------
function decodeEntity(body) {
  if (Object.hasOwn(NAMED_ENTITIES, body)) return NAMED_ENTITIES[body];
  if (!body.startsWith('#')) return undefined;
  const digits = body.slice(1);
  const hex = digits[0] === 'x' || digits[0] === 'X';
  const body2 = hex ? digits.slice(1) : digits;
  if (!body2) return undefined;
  const code = Number.parseInt(body2, hex ? 16 : 10);
  return Number.isInteger(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : undefined;
}

/** XML 文本 → 显示文本(支持命名实体与十进制/十六进制数字实体)。 */
export function decodeXml(s) {
  if (!s.includes('&')) return s;
  let out = '';
  let i = 0;
  while (i < s.length) {
    const at = s.indexOf('&', i);
    if (at === -1) break;
    out += s.slice(i, at);
    const semi = s.indexOf(';', at + 1);
    const decoded = semi === -1 || semi - at > MAX_ENTITY_LEN ? undefined : decodeEntity(s.slice(at + 1, semi));
    if (decoded === undefined) {
      out += '&';
      i = at + 1;
    } else {
      out += decoded;
      i = semi + 1;
    }
  }
  return out + s.slice(i);
}

/** 显示文本 → XML 文本(只转义 & < >,避免把已有实体二次转义)。 */
export function encodeXml(s) {
  return String(s).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

// ---------------------------------------------------------------------------
// 段落扫描
// ---------------------------------------------------------------------------
/** 记录一个 <w:t> 文本段(元素边界 + 解码后的文本),返回继续扫描的位置。 */
function readTextSegment(xml, at, tag, para) {
  if (tag.selfClosing) {
    para.segs.push({ start: at, end: tag.end, openEnd: tag.end, attrs: tag.attrs, text: '' });
    return tag.end;
  }
  const closeAt = xml.indexOf('</w:t>', tag.end);
  const contentEnd = closeAt === -1 ? xml.length : closeAt;
  para.segs.push({
    start: at,
    openEnd: tag.end,
    end: closeAt === -1 ? xml.length : closeAt + '</w:t>'.length,
    attrs: tag.attrs,
    text: decodeXml(xml.slice(tag.end, contentEnd)),
  });
  return closeAt === -1 ? xml.length : closeAt + '</w:t>'.length;
}

/**
 * 按文档顺序列出主文档里的每个 <w:p>。
 * 段落不嵌套(仅文本框里会出现),这里用栈处理:内层段落的文本归内层。
 * @returns {Array<{index: number, start: number, openEnd: number, end: number, text: string, segs: Array}>}
 */
export function scanParagraphs(xml) {
  const out = [];
  const open = [];
  const state = { tables: 0 };
  let i = 0;
  while (i < xml.length) {
    const at = xml.indexOf('<', i);
    if (at === -1) break;
    i = scanParagraphTag(xml, at, state, open, out);
  }
  return out;
}

/** 处理 `<` 处的一个标签,返回继续扫描的位置。 */
function scanParagraphTag(xml, at, state, open, out) {
  const tag = readTagAt(xml, at);
  if (!tag) return at + 1;
  if (tag.name === TABLE) {
    if (tag.closing) state.tables = Math.max(0, state.tables - 1);
    else if (!tag.selfClosing) state.tables += 1;
    return tag.end;
  }
  if (tag.name === PARAGRAPH) return handleParagraphTag(xml, at, tag, open, out, state.tables);
  if (tag.name === TEXT && open.length) return readTextSegment(xml, at, tag, open.at(-1));
  return tag.end;
}

/** 样式名 → 标题级别:Heading1 / 标题 1 / 1 都算。 */
function styleLevel(value) {
  const raw = String(value ?? '').trim().toLowerCase().replaceAll(' ', '');
  if (!raw) return 0;
  let rest = raw;
  if (rest.startsWith('heading')) rest = rest.slice(7);
  else if (rest.startsWith('标题')) rest = rest.slice(2);
  if (rest.length !== 1 || rest < '1' || rest > '9') return 0;
  return Number(rest);
}

/** 段落是不是标题:看 pStyle(Heading1/标题 1/1),再看 outlineLvl。 */
function headingLevelOf(slice) {
  const at = slice.indexOf('<w:pStyle');
  if (at >= 0) {
    const tag = readTagAt(slice, at);
    if (tag?.name === 'w:pStyle') {
      const level = styleLevel(attrIn(tag.attrs, 'w:val'));
      if (level) return level;
    }
  }
  const outline = slice.indexOf('<w:outlineLvl');
  if (outline >= 0) {
    const tag = readTagAt(slice, outline);
    if (tag?.name === 'w:outlineLvl') {
      const n = Number.parseInt(attrIn(tag.attrs, 'w:val') ?? '', 10);
      if (Number.isInteger(n) && n >= 0 && n <= 8) return n + 1;
    }
  }
  return 0;
}

function handleParagraphTag(xml, at, tag, open, out, tables) {
  if (tag.closing) {
    const top = open.pop();
    if (top) out.push(finishParagraph(xml, top, tag.end, out.length + 1));
  } else if (!tag.selfClosing) {
    open.push({ start: at, openEnd: tag.end, segs: [], inTable: tables > 0 });
  }
  return tag.end;
}

function finishParagraph(xml, top, end, index) {
  return {
    index,
    start: top.start,
    openEnd: top.openEnd,
    end,
    inTable: top.inTable,
    headingLevel: headingLevelOf(xml.slice(top.start, end)),
    text: top.segs.map((s) => s.text).join(''),
    segs: top.segs,
  };
}

// ---------------------------------------------------------------------------
// 段落内部的取用:段落属性 / 首个 run 的字符格式 / 新建元素
// ---------------------------------------------------------------------------
/** 从 `from` 起找指定名字的成对标签,返回其完整原文(找不到或自闭合时按需返回)。 */
function pairedAt(slice, from, name) {
  let i = from;
  while (i < slice.length) {
    const at = slice.indexOf(`<${name}`, i);
    if (at === -1) return { start: -1, end: -1, raw: '' };
    const tag = readTagAt(slice, at);
    if (tag && tag.name === name && !tag.closing) {
      if (tag.selfClosing) return { start: at, end: tag.end, raw: slice.slice(at, tag.end) };
      const closeAt = slice.indexOf(`</${name}>`, tag.end);
      if (closeAt === -1) return { start: -1, end: -1, raw: '' };
      return { start: at, end: closeAt + name.length + 3, raw: slice.slice(at, closeAt + name.length + 3) };
    }
    i = at + 1;
  }
  return { start: -1, end: -1, raw: '' };
}

/** 段落开头(必须在最前)的 <w:pPr>…</w:pPr> 原文。 */
function paragraphProps(slice, from) {
  return pairedAt(slice, from, 'w:pPr');
}

/**
 * 新 run 的字符格式:沿用段落里第一个 run 的 <w:rPr>。
 * 从 pPr 之后开始找,免得拿到段落标记(w:pPr 内部)的 rPr。
 */
function firstRunProps(slice, from) {
  return pairedAt(slice, from, 'w:rPr').raw;
}

/**
 * 构造 <w:t>。首尾空白必须配 xml:space="preserve",但原标签上可能已经有这个属性 ——
 * 重复写会得到 "Attribute xml:space redefined",mammoth/Word 都会拒绝整个文档。
 */
function textElement(text, attrs = '') {
  const keep = attrs || '';
  const add = text !== text.trim() && attrIn(keep, 'xml:space') === undefined;
  return `<w:t${add ? ' xml:space="preserve"' : ''}${keep}>${encodeXml(text)}</w:t>`;
}

function runXml(rPr, text) {
  return `<w:r>${rPr}${textElement(text)}</w:r>`;
}

/** 段落开头标签原文,原样保留段落自身的属性(rsid 等)。 */
function paragraphOpenTag(slice, openEndRel) {
  return slice.slice(0, openEndRel);
}

// ---------------------------------------------------------------------------
// 文档对象:每次结构性改动后重新扫描,段落序号始终以当前状态为准
// ---------------------------------------------------------------------------
function openDocument(xml, zip) {
  return { xml, paragraphs: scanParagraphs(xml), zip };
}

function splice(doc, start, end, inner) {
  doc.xml = doc.xml.slice(0, start) + inner + doc.xml.slice(end);
  doc.paragraphs = scanParagraphs(doc.xml);
}

/**
 * 一次性重建:把若干互不重叠的 [start, end) 区间替换成 next,最后只扫一遍段落。
 * 逐个 splice 每次都要复制整份 XML,段落一多就是 O(n²);这里是 O(n)。
 */
function applyEdits(doc, edits) {
  if (!edits.length) return;
  const sorted = [...edits].sort((a, b) => a.start - b.start);
  let out = '';
  let cursor = 0;
  for (const edit of sorted) {
    out += doc.xml.slice(cursor, edit.start) + edit.next;
    cursor = edit.end;
  }
  doc.xml = out + doc.xml.slice(cursor);
  doc.paragraphs = scanParagraphs(doc.xml);
}

function clip(text, max = 32) {
  const s = String(text);
  return s.length > max ? `${s.slice(0, max)}…` : s;
}

function requireText(value, what) {
  if (typeof value !== 'string' || !value.length) throw new OfficeError(`${what} 不能为空`, 'INVALID_ARGS');
  if (value.length > CAPS.MAX_TEXT_BYTES) {
    throw new OfficeError(`${what} 过长（${value.length} 字符，上限 ${CAPS.MAX_TEXT_BYTES}）`, 'INVALID_ARGS');
  }
  return value;
}

/** 用段落序号(≥1)或整段原文定位一个段落。 */
function locate(doc, spec, opName) {
  if (spec.paragraph !== undefined && spec.paragraph !== null && spec.paragraph !== '') {
    const n = Number(spec.paragraph);
    if (!Number.isInteger(n) || n < 1) {
      throw new OfficeError(`${opName}: paragraph 需为 ≥ 1 的整数（正文段落序号，从 1 开始）`, 'INVALID_ARGS');
    }
    const hit = doc.paragraphs[n - 1];
    if (!hit) {
      throw new OfficeError(`${opName}: 段落序号 ${n} 超出范围，本文档共 ${doc.paragraphs.length} 段。可用 office_read 查看实际内容`, 'PARAGRAPH_NOT_FOUND');
    }
    return hit;
  }
  if (typeof spec.match !== 'string' || !spec.match.length) {
    throw new OfficeError(`${opName}: 需要用 paragraph（段落序号）或 match（整段原文）之一来定位段落`, 'INVALID_ARGS');
  }
  const want = spec.match.trim();
  const hits = doc.paragraphs.filter((p) => p.text.trim() === want);
  if (!hits.length) {
    throw new OfficeError(`${opName}: 没有哪一段的正文等于「${clip(want)}」。可用 office_read 看实际内容，或改用 paragraph 序号`, 'PARAGRAPH_NOT_FOUND');
  }
  if (hits.length > 1) {
    throw new OfficeError(
      `${opName}: 有 ${hits.length} 段正文都等于「${clip(want)}」（第 ${hits.map((h) => h.index).join('、')} 段），请改用 paragraph 序号或写更长的 match`,
      'AMBIGUOUS_MATCH'
    );
  }
  return hits[0];
}

// ---------------------------------------------------------------------------
// replace_text:跨 run 的查找替换(Word 常把一句话拆进多个 <w:t>)
// ---------------------------------------------------------------------------
/** 找出段落里所有待替换区间(坐标是段落的显示文本)。 */
function findEdits(text, find, replace, max) {
  const edits = [];
  let from = 0;
  while (edits.length < max) {
    const at = text.indexOf(find, from);
    if (at === -1) break;
    edits.push({ at, to: at + find.length, replace });
    from = at + find.length;
  }
  return edits;
}

/** 段落各段在显示文本里的起点。 */
function segmentOffsets(segs) {
  const offsets = [];
  let at = 0;
  for (const seg of segs) {
    offsets.push(at);
    at += seg.text.length;
  }
  return offsets;
}

/** 该段落里需要重写的 <w:t> 区间(不直接改 XML,交给 applyEdits 一次重建)。 */
function segmentEdits(para, edits) {
  const offsets = segmentOffsets(para.segs);
  const out = [];
  for (let i = 0; i < para.segs.length; i++) {
    const seg = para.segs[i];
    const next = rebuildSegmentText(seg.text, offsets[i], edits);
    if (next !== seg.text) out.push({ start: seg.start, end: seg.end, next: textElement(next, seg.attrs) });
  }
  return out;
}

/** 单段的新文本:按落在它内部(或跨进来)的替换区间拼接。 */
function rebuildSegmentText(text, offset, edits) {
  let out = '';
  let cursor = offset;
  for (const edit of edits) {
    const from = Math.max(edit.at, offset);
    const to = Math.min(edit.to, offset + text.length);
    if (from >= to) continue;
    out += text.slice(cursor - offset, from - offset);
    if (edit.at >= offset) out += edit.replace;
    cursor = to;
  }
  return out + text.slice(cursor - offset);
}

function replaceText(doc, spec) {
  const find = requireText(spec.find, 'replace_text 的 find');
  const replace = spec.replace === undefined || spec.replace === null ? '' : String(spec.replace);
  const max = spec.limit === undefined || spec.limit === null ? Infinity : Number(spec.limit);
  if (max !== Infinity && (!Number.isInteger(max) || max < 1)) {
    throw new OfficeError('replace_text 的 limit 需为 ≥ 1 的整数', 'INVALID_ARGS');
  }

  let left = max;
  let hits = 0;
  const rangeEdits = [];
  for (const para of doc.paragraphs) {
    if (left <= 0) break;
    const edits = findEdits(para.text, find, replace, left);
    if (!edits.length) continue;
    rangeEdits.push(...segmentEdits(para, edits));
    hits += edits.length;
    left -= edits.length;
  }
  applyEdits(doc, rangeEdits);
  return hits ? `替换文字 ${hits} 处` : `没有找到「${clip(find)}」（未改动）`;
}

// ---------------------------------------------------------------------------
// 整段改写 / 插入 / 删除
// ---------------------------------------------------------------------------
function setParagraph(doc, spec) {
  const text = requireText(spec.text, 'set_paragraph 的 text');
  const target = locate(doc, spec, 'set_paragraph');
  const slice = doc.xml.slice(target.start, target.end);
  const openEndRel = target.openEnd - target.start;
  const props = paragraphProps(slice, openEndRel);
  const rPr = firstRunProps(slice, props.end);
  splice(doc, target.start, target.end, `${paragraphOpenTag(slice, openEndRel)}${props.raw}${runXml(rPr, text)}</w:p>`);
  return `改写第 ${target.index} 段`;
}

function headingLevel(raw) {
  if (raw === undefined || raw === null || raw === '') return 0;
  const n = Number(raw);
  if (!Number.isInteger(n) || n < 1 || n > 9) throw new OfficeError('insert_paragraph 的 heading 需为 1–9 的整数（对应 Word 内置样式 Heading1–Heading9）', 'INVALID_ARGS');
  return n;
}

function buildParagraph(text, level) {
  const pPr = level ? `<w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>` : '';
  return `<w:p>${pPr}${runXml('', text)}</w:p>`;
}

/** 正文追加位置:必须落在 <w:sectPr>(页面设置)之前,否则文档打不开。 */
function bodyAppendAt(xml) {
  const close = xml.lastIndexOf('</w:body>');
  if (close === -1) throw new OfficeError('这个 .docx 的正文结构异常（找不到 </w:body>），已放弃修改', 'BAD_DOCX');
  const sectAt = xml.lastIndexOf('<w:sectPr', close);
  if (sectAt === -1) return close;
  return xml.slice(sectAt, close).trimEnd().endsWith('</w:sectPr>') ? sectAt : close;
}

function bodyStartAt(xml) {
  const at = xml.indexOf('<w:body');
  if (at === -1) return -1;
  const tag = readTagAt(xml, at);
  return tag ? tag.end : -1;
}

function insertPlace(doc, spec) {
  const position = spec.position === undefined || spec.position === null ? 'end' : String(spec.position);
  if (position === 'end') return { at: bodyAppendAt(doc.xml), where: '文末' };
  if (position === 'start') {
    const at = bodyStartAt(doc.xml);
    if (at === -1) throw new OfficeError('这个 .docx 的正文结构异常（找不到 <w:body>），已放弃修改', 'BAD_DOCX');
    return { at, where: '文首' };
  }
  if (position !== 'before' && position !== 'after') {
    throw new OfficeError('insert_paragraph 的 position 可为 end/start/before/after', 'INVALID_ARGS');
  }
  const target = locate(doc, spec, 'insert_paragraph');
  return { at: position === 'before' ? target.start : target.end, where: `第 ${target.index} 段${position === 'before' ? '前' : '后'}` };
}

function insertParagraph(doc, spec) {
  const text = requireText(spec.text, 'insert_paragraph 的 text');
  const para = buildParagraph(text, headingLevel(spec.heading));
  const { at, where } = insertPlace(doc, spec);
  splice(doc, at, at, para);
  return `在${where}插入段落`;
}

function deleteParagraph(doc, spec) {
  const target = locate(doc, spec, 'delete_paragraph');
  splice(doc, target.start, target.end, '');
  return `删除第 ${target.index} 段`;
}

// ---------------------------------------------------------------------------
// set_style:改字体 / 行距 / 缩进等排版(直接写在段落与 run 上,不依赖样式表)
// ---------------------------------------------------------------------------
/** w:pPr 子元素的 schema 顺序(只关心会写进去的几个 + 常见锚点)。 */
/** 样式的字符级部分 → OOXML 标签。 */
function runStyleTags(norm) {
  const tags = {};
  if (norm.font || norm.fontAscii) {
    const west = attrText(norm.fontAscii || norm.font);
    const east = attrText(norm.font || west);
    tags['w:rFonts'] = `<w:rFonts w:ascii="${west}" w:hAnsi="${west}" w:eastAsia="${east}" w:cs="${west}"/>`;
  }
  if (norm.sizeHalfPt) {
    tags['w:sz'] = `<w:sz w:val="${norm.sizeHalfPt}"/>`;
    tags['w:szCs'] = `<w:szCs w:val="${norm.sizeHalfPt}"/>`;
  }
  if (norm.bold !== undefined) tags['w:b'] = norm.bold ? '<w:b/>' : '<w:b w:val="0"/>';
  if (norm.color) tags['w:color'] = `<w:color w:val="${norm.color}"/>`;
  return tags;
}

/** 样式的段落级部分 → OOXML 标签(间距合成一个 w:spacing)。 */
function paragraphStyleTags(norm) {
  const tags = {};
  const spacing = [];
  if (norm.beforeTwips !== undefined) spacing.push(`w:before="${norm.beforeTwips}"`);
  if (norm.afterTwips !== undefined) spacing.push(`w:after="${norm.afterTwips}"`);
  if (norm.line !== undefined) spacing.push(`w:line="${norm.line}"`, `w:lineRule="${norm.lineRule}"`);
  if (spacing.length) tags['w:spacing'] = `<w:spacing ${spacing.join(' ')}/>`;
  if (norm.firstLineTwips !== undefined) tags['w:ind'] = `<w:ind w:firstLine="${norm.firstLineTwips}"/>`;
  if (norm.align) tags['w:jc'] = `<w:jc w:val="${norm.align}"/>`;
  return tags;
}

const PPR_ORDER = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'];

/** 重排一个段落的样式:段落属性 + 段落标记 + 各 run。 */
function restyleParagraph(paragraphXml, norm) {
  const runTags = runStyleTags(norm);
  const additions = { ...paragraphStyleTags(norm) };
  if (Object.keys(runTags).length) additions['w:rPr'] = `<w:rPr>${rebuildChildren('', runTags, RPR_ORDER)}</w:rPr>`;
  const styled = Object.keys(additions).length ? upsertProps(paragraphXml, PARAGRAPH, 'w:pPr', additions, PPR_ORDER) : paragraphXml;
  return restyleRuns(styled, runTags);
}

function positiveInt(value, what) {
  const n = Number(value);
  if (!Number.isInteger(n) || n < 1) throw new OfficeError(`${what} 需为 ≥ 1 的整数（段落序号，从 1 开始）`, 'INVALID_ARGS');
  return n;
}

/** set_style 的角色:all 全部 / body 正文(非标题) / headings 标题 / table 表格内段落。 */
const ROLE_SCOPES = new Set(['all', 'body', 'headings', 'table']);

function matchesRole(para, scope) {
  if (scope === 'body') return para.headingLevel === 0;
  if (scope === 'headings') return para.headingLevel > 0;
  if (scope === 'table') return para.inTable;
  return true;
}

/** 段落构成说明,报错时用得上。 */
function describePool(paragraphs) {
  const headings = paragraphs.filter((p) => p.headingLevel > 0).length;
  const inTable = paragraphs.filter((p) => p.inTable).length;
  return `共 ${paragraphs.length} 段（标题 ${headings}、正文 ${paragraphs.length - headings}、表格内 ${inTable}）`;
}

function rangePool(doc, spec, opName = 'set_style') {
  const from = spec.from === undefined ? 1 : positiveInt(spec.from, `${opName} 的 from`);
  const to = spec.to === undefined ? doc.paragraphs.length : positiveInt(spec.to, `${opName} 的 to`);
  if (from > to) throw new OfficeError(`${opName} 的 from(${from}) 不能大于 to(${to})`, 'INVALID_ARGS');
  const pool = doc.paragraphs.slice(from - 1, to);
  if (!pool.length) throw new OfficeError(`${opName} 的区间 ${from}–${to} 超出范围（${describePool(doc.paragraphs)}）`, 'PARAGRAPH_NOT_FOUND');
  return pool;
}

/** set_style 的作用范围:单段 / 区间 + 角色。 */
function styleTargets(doc, spec, opName = 'set_style') {
  const scope = spec.scope === undefined || spec.scope === null || spec.scope === '' ? '' : String(spec.scope);
  if (scope && !ROLE_SCOPES.has(scope)) {
    throw new OfficeError(`${opName} 的 scope 可为 all（全部段落）/ body（正文，不含标题）/ headings（标题）/ table（表格内段落）`, 'INVALID_ARGS');
  }
  const hasRange = spec.from !== undefined || spec.to !== undefined;
  if (!scope && !hasRange) return [locate(doc, spec, opName)];
  const pool = rangePool(doc, spec, opName);
  const picked = scope && scope !== 'all' ? pool.filter((p) => matchesRole(p, scope)) : pool;
  if (!picked.length) {
    const where = hasRange ? '指定区间内' : '本文档中';
    throw new OfficeError(`${opName}: ${where}没有 scope "${scope}" 匹配的段落（${describePool(doc.paragraphs)}）`, 'PARAGRAPH_NOT_FOUND');
  }
  return picked;
}

function setStyle(doc, spec) {
  const norm = normalizeStyleSpec(spec, { allowHeadings: false, allowTable: false });
  if (isEmptyStyle(norm)) {
    throw new OfficeError('set_style 至少要给一个样式属性（font / sizePt / lineSpacingPt / lineSpacingMultiple / firstLineIndentChars / align / spacingBeforePt / spacingAfterPt / bold / color）', 'INVALID_ARGS');
  }
  const targets = styleTargets(doc, spec);
  const edits = targets.map((target) => ({
    start: target.start,
    end: target.end,
    next: restyleParagraph(doc.xml.slice(target.start, target.end), norm),
  }));
  applyEdits(doc, edits);
  return `设置 ${targets.length} 段样式`;
}

// ---------------------------------------------------------------------------
// set_numbering:标题多级自动编号 + 正文列表挂到当前标题层级之下
// ---------------------------------------------------------------------------
/** 段落是否已经有列表编号(w:numPr)。 */
function hasNumPr(slice) {
  const at = slice.indexOf('<w:pPr');
  if (at === -1) return false;
  const tag = readTagAt(slice, at);
  if (tag?.name !== 'w:pPr' || tag.selfClosing) return false;
  const closeAt = slice.indexOf('</w:pPr>', tag.end);
  const inner = closeAt === -1 ? '' : slice.slice(tag.end, closeAt);
  return directChildren(inner).some((c) => c.name === 'w:numPr');
}

/** set_numbering 的目标段落:默认全篇;给了 scope/from-to/paragraph/match 就按那套选。 */
function numberingTargets(doc, spec) {
  const located = spec.scope !== undefined || spec.from !== undefined || spec.to !== undefined
    || spec.paragraph !== undefined || spec.match !== undefined;
  return located ? styleTargets(doc, spec, 'set_numbering') : [...doc.paragraphs];
}

/** 同一套 (preset, linkToHeading) 只往 numbering.xml 里加一次。 */
function requireBoolean(value, what) {
  if (typeof value !== 'boolean') throw new OfficeError(`${what} 需为 true / false`, 'INVALID_ARGS');
  return value;
}

function numberingFor(doc, spec) {
  const preset = requirePreset(spec.style);
  if (spec.linkToHeading !== undefined) requireBoolean(spec.linkToHeading, 'set_numbering 的 linkToHeading');
  const linkToHeading = spec.linkToHeading === undefined ? true : spec.linkToHeading;
  const exclude = spec.exclude === undefined || spec.exclude === null ? [] : spec.exclude;
  if (!Array.isArray(exclude)) throw new OfficeError('set_numbering 的 exclude 需为数组，如 ["Heading1"]', 'INVALID_ARGS');
  const startFrom = spec.startFrom;
  const key = [preset, linkToHeading, JSON.stringify(exclude), JSON.stringify(startFrom ?? null)].join('|');
  doc.numberings = doc.numberings ?? new Map();
  if (!doc.numberings.has(key)) {
    const setup = ensureNumbering(doc.zip, { documentXml: doc.xml, preset, linkToHeading, exclude, startFrom });
    doc.numberings.set(key, { ...setup, preset, linkToHeading });
  }
  return doc.numberings.get(key);
}

/** 给一个段落写上 numPr(返回新的段落 XML,不直接改文档)。 */
function numberedParagraph(slice, ilvl, numId) {
  const numPr = `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
  return upsertProp(slice, numPr);
}

function upsertProp(slice, numPr) {
  return upsertProps(slice, PARAGRAPH, 'w:pPr', { 'w:numPr': numPr }, PPR_ORDER);
}

/**
 * 收集要写 numPr 的段落:标题(linkToHeading=false 时)用自身层级,
 * 正文只处理「本来就是列表项」的段落,层级取当前标题层级的下一级。
 */
/** 标题段落:linkToHeading=false 时按自身链条层级显式编号(exclude 掉的不编号)。 */
function collectHeadingEdit(para, setup, state) {
  if (setup.linkToHeading || !state.chainLevel) return;
  state.edits.push({ para, ilvl: state.chainLevel - 1 });
  state.explicitHeadings += 1;
}

/** 正文段落:只处理「本来就是列表项」的,层级取当前标题在链上的下一级。 */
function collectListEdit(doc, para, state) {
  if (!hasNumPr(doc.xml.slice(para.start, para.end))) return;
  if (!state.chainLevel) {
    // 上方还没有参与编号的标题时给不出正确编号,跳过并如实报出
    state.skipped += 1;
    return;
  }
  state.edits.push({ para, ilvl: Math.min(8, state.chainLevel) });
}

function numberingEdits(doc, targets, setup) {
  const state = { chainLevel: 0, edits: [], explicitHeadings: 0, skipped: 0 };
  for (const para of doc.paragraphs) {
    if (para.headingLevel > 0) state.chainLevel = setup.specs.headingToChainLevel.get(para.headingLevel) ?? 0;
    if (!targets.has(para)) continue;
    if (para.headingLevel > 0) collectHeadingEdit(para, setup, state);
    else collectListEdit(doc, para, state);
  }
  return state;
}

function setNumbering(doc, spec) {
  const setup = numberingFor(doc, spec);
  const targets = new Set(numberingTargets(doc, spec));
  const { edits, explicitHeadings, skipped } = numberingEdits(doc, targets, setup);
  applyEdits(doc, edits.map((edit) => ({
    start: edit.para.start,
    end: edit.para.end,
    next: numberedParagraph(doc.xml.slice(edit.para.start, edit.para.end), edit.ilvl, setup.numId),
  })));
  const lists = edits.length - explicitHeadings;
  if (!lists && !explicitHeadings && !setup.linkToHeading) {
    throw new OfficeError('set_numbering: 没有可编号的段落（正文里没有列表项，标题也没有落进选择范围）', 'PARAGRAPH_NOT_FOUND');
  }
  const head = setup.linkToHeading ? '标题按样式链接自动编号' : `标题 ${explicitHeadings} 段显式编号`;
  const reuse = setup.reused ? '（更新已有编号定义）' : '';
  const skip = skipped ? `；跳过 ${skipped} 段（其上方还没有参与编号的标题）` : '';
  return `多级编号 ${setup.preset}${reuse}：正文列表 ${lists} 段挂到标题层级之下、${head}${skip}`;
}

// ---------------------------------------------------------------------------
// 表格:样式 / 增删行列 / 合并单元格
// ---------------------------------------------------------------------------
/** set_table 与表格结构操作选表:`table` 序号(从 1 起)或 `scope: "all"`;只有一个表时可省略。 */
/** 只留 0 行的空表格(删行删空后的残留,Word 不渲染但 XML 里还在)。 */
function keepEmptyOnly(targets, opName) {
  const empties = targets.filter((t) => t.rows === 0);
  if (!empties.length) throw new OfficeError(`${opName}: 没有 0 行的空表格可删（本次选中 ${targets.length} 个表格，都有行）`, 'TABLE_NOT_FOUND');
  return empties;
}

function tableTargets(doc, spec, opName) {
  const all = scanTables(doc.xml);
  if (!all.length) throw new OfficeError(`${opName}: 这个文档里没有表格`, 'TABLE_NOT_FOUND');
  if (spec.onlyEmpty === true) return keepEmptyOnly(all, opName);
  const scope = spec.scope === undefined || spec.scope === null || spec.scope === '' ? '' : String(spec.scope);
  if (scope) {
    if (scope !== 'all') throw new OfficeError(`${opName}: scope 只支持 "all"（所有表格）；单个表格请用 table 序号`, 'INVALID_ARGS');
    return all;
  }
  if (spec.table === undefined || spec.table === null || spec.table === '') {
    if (all.length === 1) return all;
    throw new OfficeError(`${opName}: 需要 table（表格序号，从 1 开始）或 scope:"all"；本文档共 ${all.length} 个表格`, 'INVALID_ARGS');
  }
  const n = Number(spec.table);
  if (!Number.isInteger(n) || n < 1) throw new OfficeError(`${opName}: table 需为 ≥ 1 的整数`, 'INVALID_ARGS');
  const hit = all[n - 1];
  if (!hit) throw new OfficeError(`${opName}: 表格序号 ${n} 超出范围，本文档共 ${all.length} 个表格`, 'TABLE_NOT_FOUND');
  return [hit];
}

/** 对选中的每个表格套一个变换(从后往前改,前面的偏移保持有效)。 */
function applyToTables(doc, spec, opName, transform, describe) {
  const targets = tableTargets(doc, spec, opName);
  const width = textWidthTwips(doc.xml);
  const edits = targets.map((target) => ({
    start: target.start,
    end: target.end,
    next: transform(doc.xml.slice(target.start, target.end), width),
  }));
  applyEdits(doc, edits);
  // 删表后要兜底(正文非空且不以表格结尾),它可能再改一次 XML
  describe.afterApply?.(doc);
  return describe(targets.length, targets.map((t) => t.index));
}

/** 正文内容结束位置:sectPr 之前(若它是最后一个元素),否则 body 闭合之前。 */
function bodyContentEnd(xml) {
  const close = xml.lastIndexOf('</w:body>');
  if (close === -1) return -1;
  const sectAt = xml.lastIndexOf('<w:sectPr', close);
  if (sectAt === -1) return close;
  return xml.slice(sectAt, close).trimEnd().endsWith('</w:sectPr>') ? sectAt : close;
}

/**
 * 删表后兜底:正文不能为空、也不能以表格结尾(这两种情况 Word 打开时会提示修复)。
 * 需要时补一个空段落。
 */
function ensureBodyEndsWithParagraph(doc) {
  const cut = bodyContentEnd(doc.xml);
  if (cut === -1) return;
  const openAt = doc.xml.indexOf('<w:body');
  const openEnd = openAt === -1 ? 0 : (readTagAt(doc.xml, openAt)?.end ?? 0);
  const body = doc.xml.slice(openEnd, cut);
  if (/<w:p[\s/>]/.test(body) && !body.trimEnd().endsWith('</w:tbl>')) return;
  splice(doc, cut, cut, '<w:p/>');
}

/** 删除整张表格(按 table 序号、scope:"all"，或 onlyEmpty:true 只清 0 行残留)。 */
function deleteTable(doc, spec) {
  const describe = (n, indexes) => `删除表格 ${indexes.join('、')}（共 ${n} 个）`;
  describe.afterApply = ensureBodyEndsWithParagraph;
  return applyToTables(doc, spec, 'delete_table', () => '', describe);
}

function setTable(doc, spec) {
  const norm = normalizeTableSpec(spec);
  if (isEmptyStyle(norm)) {
    throw new OfficeError('set_table 至少要给一个属性（borders / headerShading / headerBold / columnWidthMode / columnWidths / align）', 'INVALID_ARGS');
  }
  return applyToTables(doc, spec, 'set_table', (slice, width) => applyTableStyle(slice, norm, width), (n) => `设置 ${n} 个表格的样式`);
}

const countOf = (spec) => {
  const n = spec.count === undefined || spec.count === null || spec.count === '' ? 1 : Number(spec.count);
  if (!Number.isInteger(n) || n < 1 || n > 100) throw new OfficeError('count 需为 1–100 的整数', 'INVALID_ARGS');
  return n;
};

function insertTableRow(doc, spec) {
  return applyToTables(doc, spec, 'insert_table_row', (slice) => insertTableRows(slice, spec.at, spec.count), (n) => `插入 ${countOf(spec)} 行（${n} 个表格）`);
}

function deleteTableRow(doc, spec) {
  requireText(spec.at === undefined || spec.at === null ? '' : String(spec.at), 'delete_table_row 的 at');
  return applyToTables(doc, spec, 'delete_table_row', (slice) => deleteTableRows(slice, spec.at, spec.count), (n) => `删除第 ${spec.at} 行起的 ${countOf(spec)} 行（${n} 个表格）`);
}

function insertTableColumn(doc, spec) {
  return applyToTables(doc, spec, 'insert_table_column', (slice) => insertTableColumns(slice, spec.at, spec.count), (n) => `插入 ${countOf(spec)} 列（${n} 个表格）`);
}

function deleteTableColumn(doc, spec) {
  requireText(spec.at === undefined || spec.at === null ? '' : String(spec.at), 'delete_table_column 的 at');
  return applyToTables(doc, spec, 'delete_table_column', (slice) => deleteTableColumns(slice, spec.at, spec.count), (n) => `删除第 ${spec.at} 列起的 ${countOf(spec)} 列（${n} 个表格）`);
}

function mergeTablesCells(doc, spec) {
  const range = parseCellRange(spec.range);
  return applyToTables(doc, spec, 'merge_table_cells', (slice) => mergeTableCells(slice, range), (n) => `合并 ${spec.range}（${n} 个表格）`);
}

function unmergeTablesCells(doc, spec) {
  const range = parseCellRange(spec.range);
  return applyToTables(doc, spec, 'unmerge_table_cells', (slice) => unmergeTableCells(slice, range), (n) => `取消合并 ${spec.range}（${n} 个表格）`);
}

// ---------------------------------------------------------------------------
// 入口
// ---------------------------------------------------------------------------
const OPS = {
  replace_text: replaceText,
  set_paragraph: setParagraph,
  set_style: setStyle,
  insert_paragraph: insertParagraph,
  delete_paragraph: deleteParagraph,
  set_table: setTable,
  insert_table_row: insertTableRow,
  delete_table_row: deleteTableRow,
  insert_table_column: insertTableColumn,
  delete_table_column: deleteTableColumn,
  merge_table_cells: mergeTablesCells,
  unmerge_table_cells: unmergeTablesCells,
  delete_table: deleteTable,
  set_numbering: setNumbering,
};

function applyOne(doc, spec, i) {
  if (!isPlainObject(spec)) throw new OfficeError(`ops[${i}] 需为对象，如 {"op":"replace_text","find":"…","replace":"…"}`, 'INVALID_ARGS');
  const run = OPS[spec.op];
  if (!run) {
    throw new OfficeError(`ops[${i}] 不支持的操作「${spec.op}」。可用：${Object.keys(OPS).join('、')}`, 'INVALID_ARGS');
  }
  return run(doc, spec);
}

/**
 * 就地修改 .docx:只重写主文档部件,其余 zip 条目原样保留。
 * 操作按顺序执行,`paragraph` 序号以每一步执行后的文档状态为准。
 * @param {Buffer} buf 原 .docx 字节
 * @param {Array<object>} ops 操作数组
 * @returns {Promise<{buf: Buffer, changes: string[], paragraphsBefore: number, paragraphsAfter: number}>}
 */
export async function editDocx(buf, ops) {
  if (!isWithinBytes(buf, CAPS.MAX_WORD_INPUT_BYTES)) throw new OfficeError('docx 文件超过 40 MB 上限', 'OFFICE_TOO_LARGE');
  if (!Array.isArray(ops) || !ops.length) throw new OfficeError('需要 ops 参数：非空的操作数组', 'INVALID_ARGS');
  if (ops.length > MAX_OPS) throw new OfficeError(`一次最多 ${MAX_OPS} 个操作（收到 ${ops.length} 个）`, 'INVALID_ARGS');
  await assertZipBudget(buf, 'docx');
  const { zip } = await openOoxml(buf, 'docx');
  const main = partOf(zip, MAIN_PART);
  if (!main) throw new OfficeError('这个 .docx 里找不到主文档部件（word/document.xml），无法修改', 'BAD_DOCX');

  const doc = openDocument(main.asText(), zip);
  const before = doc.paragraphs.length;
  const changes = ops.map((spec, i) => applyOne(doc, spec, i));
  zip.file(main.name, doc.xml);
  return {
    buf: zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }),
    changes,
    paragraphsBefore: before,
    paragraphsAfter: doc.paragraphs.length,
  };
}
