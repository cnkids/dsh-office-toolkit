// .docx formatting extraction for 行文规则比对: fonts, sizes, line spacing,
// indents, alignment, page setup — with style inheritance resolved.
//
// mammoth is deliberately lossy (it only wants the text), so this reads
// `word/document.xml` + `word/styles.xml` directly through the linear HTML/XML
// scanner in markup.js. No regexes, no extra dependencies.
import { attrOf, parseHtml } from './md.js';
import { attrIn } from './markup.js';

const TWIPS_PER_PT = 20;
const TWIPS_PER_MM = 1440 / 25.4;
const HALF_POINTS_PER_PT = 2;
const LINES_PER_UNIT = 240; // w:line="240" == single spacing
const CHARS_PER_UNIT = 100; // w:firstLineChars="200" == 2 chars
const MAX_STYLE_DEPTH = 12;
const MAX_REPORTED_PARAGRAPHS = 80;
const MAX_OUTLIERS = 20;
const MARGIN_LABELS = { top: '上', bottom: '下', left: '左', right: '右' };
const TABLE_HEAD = '| # | 段落文字 | 样式 | 对齐 | 行距 | 缩进 | 中文字体 | 西文字体 | 字号 | 加粗 |';
const TABLE_RULE = '| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |';

const ALIGNMENTS = new Map([
  ['left', '左对齐'],
  ['start', '左对齐'],
  ['right', '右对齐'],
  ['end', '右对齐'],
  ['center', '居中'],
  ['both', '两端对齐'],
  ['distribute', '分散对齐'],
  ['justify', '两端对齐'],
]);

/** First child element with the given (lowercased) tag name. */
function childOf(node, tag) {
  return (node?.children || []).find((c) => c.tag === tag);
}

/** Descendant element with the given tag name, breadth-first. */
function findTag(node, tag) {
  if (!node) return undefined;
  const queue = [...(node.children || [])];
  while (queue.length) {
    const current = queue.shift();
    if (current.tag === tag) return current;
    if (current.children) queue.push(...current.children);
  }
  return undefined;
}

/** Drop trailing whitespace without a backtracking regex. */
function trimEnd(text) {
  let end = text.length;
  while (end > 0) {
    const ch = text[end - 1];
    if (ch !== ' ' && ch !== '\t' && ch !== '\n' && ch !== '\r') break;
    end -= 1;
  }
  return text.slice(0, end);
}

/** `w:val` of a toggle property: absent means on, "0"/"false"/"off" means off. */
function isOn(node) {
  if (!node) return false;
  const value = String(attrOf(node, 'w:val') ?? '').toLowerCase();
  return value === '' || value === 'true' || value === '1' || value === 'on';
}

/**
 * Toggle written as an attribute (`w:default="1"`): absent means off.
 * Uses attrIn (not attrOf) on purpose — attrOf maps a missing attribute and an
 * empty one to the same '' , which here would make every style look like the
 * document default.
 */
function isAttrOn(node, name) {
  if (!node) return false;
  const raw = attrIn(node.attrs, name);
  if (raw === undefined) return false;
  const value = String(raw).toLowerCase();
  return value === '' || value === 'true' || value === '1' || value === 'on';
}

/** Integer attribute, or undefined when absent/not numeric. */
function intAttr(node, name) {
  if (!node) return undefined;
  const raw = attrOf(node, name);
  if (raw === '' || raw === undefined) return undefined;
  const value = Number.parseInt(raw, 10);
  return Number.isFinite(value) ? value : undefined;
}

function ptFromTwips(twips) {
  return twips === undefined ? undefined : Math.round((twips / TWIPS_PER_PT) * 100) / 100;
}

function ptFromHalfPoints(half) {
  return half === undefined ? undefined : Math.round((half / HALF_POINTS_PER_PT) * 100) / 100;
}

/** `w:spacing` -> line spacing / paragraph spacing. */
function readSpacingProps(spacing) {
  const props = {};
  if (!spacing) return props;
  const rule = String(attrOf(spacing, 'w:lineRule') || 'auto').toLowerCase();
  const line = intAttr(spacing, 'w:line');
  if (line !== undefined) {
    props.lineSpacing = rule === 'auto'
      ? { rule: 'auto', lines: Math.round((line / LINES_PER_UNIT) * 100) / 100 }
      : { rule, pt: ptFromTwips(line) };
  }
  const beforeLines = intAttr(spacing, 'w:beforeLines');
  const afterLines = intAttr(spacing, 'w:afterLines');
  const before = intAttr(spacing, 'w:before');
  const after = intAttr(spacing, 'w:after');
  if (beforeLines !== undefined) props.spaceBeforeLines = beforeLines / CHARS_PER_UNIT;
  else if (before !== undefined) props.spaceBeforePt = ptFromTwips(before);
  if (afterLines !== undefined) props.spaceAfterLines = afterLines / CHARS_PER_UNIT;
  else if (after !== undefined) props.spaceAfterPt = ptFromTwips(after);
  return props;
}

/** `w:ind` -> first-line / hanging / left indent. */
function readIndentProps(ind) {
  const props = {};
  if (!ind) return props;
  const firstChars = intAttr(ind, 'w:firstlinechars');
  const firstTwips = intAttr(ind, 'w:firstline');
  const hangingChars = intAttr(ind, 'w:hangingchars');
  const hangingTwips = intAttr(ind, 'w:hanging');
  const leftChars = intAttr(ind, 'w:leftchars');
  const leftTwips = intAttr(ind, 'w:left');
  const rightChars = intAttr(ind, 'w:rightchars');
  const rightTwips = intAttr(ind, 'w:right');
  if (firstChars !== undefined) props.firstLineChars = firstChars / CHARS_PER_UNIT;
  else if (firstTwips !== undefined) props.firstLinePt = ptFromTwips(firstTwips);
  if (hangingChars !== undefined) props.hangingChars = hangingChars / CHARS_PER_UNIT;
  else if (hangingTwips !== undefined) props.hangingPt = ptFromTwips(hangingTwips);
  if (leftChars !== undefined) props.leftChars = leftChars / CHARS_PER_UNIT;
  else if (leftTwips !== undefined) props.leftPt = ptFromTwips(leftTwips);
  if (rightChars !== undefined) props.rightChars = rightChars / CHARS_PER_UNIT;
  else if (rightTwips !== undefined) props.rightPt = ptFromTwips(rightTwips);
  return props;
}

/** Paragraph properties carried out of a `w:pPr` element. */
function readParagraphProps(pPr) {
  const props = {};
  if (!pPr) return props;
  const style = childOf(pPr, 'w:pstyle');
  if (style) props.styleId = attrOf(style, 'w:val');
  const align = childOf(pPr, 'w:jc');
  if (align) props.alignment = ALIGNMENTS.get(String(attrOf(align, 'w:val')).toLowerCase()) || attrOf(align, 'w:val');
  const outline = childOf(pPr, 'w:outlinelvl');
  if (outline) props.outlineLevel = intAttr(outline, 'w:val');
  return {
    ...props,
    ...readSpacingProps(childOf(pPr, 'w:spacing')),
    ...readIndentProps(childOf(pPr, 'w:ind')),
  };
}

/**
 * Font faces from `word/theme/theme1.xml`. Word templates reference fonts by
 * theme (`w:eastAsiaTheme="minorEastAsia"`) instead of naming them, so without
 * this a document looks like it has no font at all.
 */
function readTheme(themeXml) {
  if (!themeXml) return {};
  const scheme = findTag(parseHtml(themeXml), 'a:fontscheme');
  const pick = (kind) => {
    const node = findTag(scheme, `a:${kind}font`);
    if (!node) return {};
    const latin = childOf(node, 'a:latin');
    const ea = childOf(node, 'a:ea');
    let hans;
    for (const font of node.children || []) {
      if (font.tag === 'a:font' && String(attrOf(font, 'script')).toLowerCase() === 'hans') hans = attrOf(font, 'typeface');
    }
    return {
      latin: latin ? attrOf(latin, 'typeface') : undefined,
      eastAsia: hans || (ea ? attrOf(ea, 'typeface') : undefined),
    };
  };
  return { major: pick('major'), minor: pick('minor') };
}

/** Resolve `w:asciiTheme` / `w:eastAsiaTheme` against the theme scheme. */
function themeFont(props, theme) {
  const style = String(props.asciiTheme || props.eastAsiaTheme || '').toLowerCase();
  if (!style) return props;
  const scheme = theme[style.startsWith('major') ? 'major' : 'minor'] || {};
  const resolved = { ...props };
  if (props.asciiTheme && !props.fontAscii && scheme.latin) resolved.fontAscii = scheme.latin;
  if (props.eastAsiaTheme && !props.fontEastAsia && scheme.eastAsia) resolved.fontEastAsia = scheme.eastAsia;
  delete resolved.asciiTheme;
  delete resolved.eastAsiaTheme;
  return resolved;
}

/** `w:rFonts` -> explicit faces plus any theme references. */
function readFontProps(fonts) {
  const props = {};
  if (!fonts) return props;
  const ascii = attrOf(fonts, 'w:ascii');
  const eastAsia = attrOf(fonts, 'w:eastAsia');
  const hAnsi = attrOf(fonts, 'w:hAnsi');
  const asciiTheme = attrOf(fonts, 'w:asciiTheme');
  const eastAsiaTheme = attrOf(fonts, 'w:eastAsiaTheme');
  if (ascii) props.fontAscii = ascii;
  if (eastAsia) props.fontEastAsia = eastAsia;
  if (hAnsi) props.fontHAnsi = hAnsi;
  if (asciiTheme) props.asciiTheme = asciiTheme;
  if (eastAsiaTheme) props.eastAsiaTheme = eastAsiaTheme;
  return props;
}

/** `w:b` / `w:i` / `w:u` toggles. */
function readToggleProps(rPr) {
  const props = {};
  const bold = childOf(rPr, 'w:b');
  if (bold) props.bold = isOn(bold);
  const italic = childOf(rPr, 'w:i');
  if (italic) props.italic = isOn(italic);
  const underline = childOf(rPr, 'w:u');
  if (underline) props.underline = isOn(underline);
  return props;
}

/** Run (character) properties carried out of a `w:rPr` element. */
function readRunProps(rPr, theme = {}) {
  if (!rPr) return {};
  const props = readFontProps(childOf(rPr, 'w:rfonts'));
  const style = childOf(rPr, 'w:rstyle');
  if (style) props.runStyleId = attrOf(style, 'w:val');
  const size = intAttr(childOf(rPr, 'w:sz'), 'w:val');
  if (size !== undefined) props.sizePt = ptFromHalfPoints(size);
  const color = childOf(rPr, 'w:color');
  if (color && String(attrOf(color, 'w:val')).toLowerCase() !== 'auto') props.color = attrOf(color, 'w:val');
  Object.assign(props, readToggleProps(rPr));
  return themeFont(props, theme);
}

function readStyle(child, id, theme) {
  const nameEl = childOf(child, 'w:name');
  const basedOn = childOf(child, 'w:basedon');
  return {
    id,
    type: attrOf(child, 'w:type') || 'paragraph',
    isDefault: isAttrOn(child, 'w:default'),
    name: nameEl ? attrOf(nameEl, 'w:val') : id,
    basedOn: basedOn ? attrOf(basedOn, 'w:val') : undefined,
    paragraph: readParagraphProps(childOf(child, 'w:ppr')),
    run: readRunProps(childOf(child, 'w:rpr'), theme),
  };
}

/** Style table plus document defaults from `word/styles.xml`. */
function readStyles(stylesXml, themeXml) {
  const defaults = { paragraph: {}, run: {} };
  const styles = new Map();
  const theme = readTheme(themeXml);
  if (!stylesXml) return { defaults, styles, theme };
  const tree = parseHtml(stylesXml);

  const docDefaults = findTag(tree, 'w:docdefaults');
  if (docDefaults) {
    const runDefault = findTag(docDefaults, 'w:rprdefault');
    const pprDefault = findTag(docDefaults, 'w:pprdefault');
    Object.assign(defaults.run, readRunProps(runDefault && childOf(runDefault, 'w:rpr'), theme));
    Object.assign(defaults.paragraph, readParagraphProps(pprDefault && childOf(pprDefault, 'w:ppr')));
  }

  const walk = (node) => {
    for (const child of node.children || []) {
      if (child.tag === 'w:style') {
        const id = attrOf(child, 'w:styleid');
        if (id) styles.set(id, readStyle(child, id, theme));
        continue;
      }
      walk(child);
    }
  };
  walk(tree);
  // Word 会把 w:default="1" 的段落样式隐式套到没有 w:pStyle 的段落上
  let defaultParagraphId;
  for (const style of styles.values()) {
    if (style.isDefault && style.type === 'paragraph') defaultParagraphId = style.id;
  }
  return { defaults, styles, theme, defaultParagraphId };
}

/** Merge a style chain (base first) onto the given defaults. */
function resolveStyle(styleId, kind, table, base) {
  const chain = [];
  let current = styleId;
  let depth = 0;
  while (current && depth < MAX_STYLE_DEPTH) {
    const style = table.styles.get(current);
    if (!style) break;
    chain.unshift(style);
    current = style.basedOn;
    depth += 1;
  }
  const merged = { ...base };
  let name;
  for (const style of chain) {
    Object.assign(merged, style[kind]);
    name = style.name || name;
  }
  return { props: merged, name };
}

/** Text of a paragraph, including runs inside hyperlinks. */
function paragraphText(p) {
  let text = '';
  const walk = (node) => {
    for (const child of node.children || []) {
      if (child.tag === 'w:t') text += (child.children || []).map((c) => c.text || '').join('');
      else if (child.tag === 'w:tab') text += '\t';
      else if (child.tag === 'w:br' || child.tag === 'w:cr') text += '\n';
      else if (child.tag !== 'w:instrtext') walk(child); // 域代码(目录等)不算正文
    }
  };
  walk(p);
  return text;
}

/** Runs with text, so the dominant character format can be computed. */
function paragraphRuns(p) {
  const runs = [];
  const walk = (node, inherited) => {
    for (const child of node.children || []) {
      if (child.tag === 'w:r') {
        const text = paragraphText(child);
        if (text) runs.push({ ...inherited, ...readRunProps(childOf(child, 'w:rpr')), text });
        continue;
      }
      walk(child, inherited);
    }
  };
  walk(p, {});
  return runs;
}

/** Effective character format of a single run. */
function runFormatOf(run, merged, table) {
  const styled = run.runStyleId ? resolveStyle(run.runStyleId, 'run', table, merged).props : merged;
  const format = { ...styled, ...run };
  delete format.runStyleId;
  delete format.text;
  return format;
}

/** Deterministic grouping key (key order must not split identical formats). */
function formatKey(format) {
  const keys = Object.keys(format).sort((a, b) => a.localeCompare(b));
  return JSON.stringify(keys.map((key) => [key, format[key]]));
}

/** Character format that covers most of the paragraph's text. */
function dominantRunFormat(runs, merged, table) {
  const counts = new Map();
  for (const run of runs) {
    const format = runFormatOf(run, merged, table);
    const key = formatKey(format);
    const entry = counts.get(key) || { format, chars: 0 };
    entry.chars += run.text.length;
    counts.set(key, entry);
  }
  let best;
  for (const entry of counts.values()) {
    if (!best || entry.chars > best.chars) best = entry;
  }
  return best ? best.format : { ...merged };
}

function toMm(twips) {
  return Math.round((twips / TWIPS_PER_MM) * 10) / 10;
}

/** `w:pgSz` -> 纸张尺寸与方向(mm)。 */
function readPageSize(size) {
  const page = {};
  if (!size) return page;
  const width = intAttr(size, 'w:w');
  const height = intAttr(size, 'w:h');
  const orientation = attrOf(size, 'w:orient');
  if (width) page.widthMm = toMm(width);
  if (height) page.heightMm = toMm(height);
  if (orientation) page.orientation = orientation === 'landscape' ? '横向' : '纵向';
  return page;
}

/** `w:pgMar` -> 上下左右页边距 + 页眉/页脚距离(mm)。 */
function readMargins(margin) {
  const margins = {};
  if (!margin) return margins;
  for (const side of Object.keys(MARGIN_LABELS)) {
    const value = intAttr(margin, `w:${side}`);
    if (value !== undefined) margins[side] = toMm(value);
  }
  return margins;
}

/** 版心之外的距离:`w:pgMar` 的 header / footer(mm)。 */
function readEdgeDistances(margin) {
  const out = {};
  if (!margin) return out;
  const header = intAttr(margin, 'w:header');
  const footer = intAttr(margin, 'w:footer');
  if (header !== undefined) out.headerMm = toMm(header);
  if (footer !== undefined) out.footerMm = toMm(footer);
  return out;
}

/** Page size and margins from the body-level `w:sectPr`. */
function readPage(sectPr) {
  if (!sectPr) return undefined;
  const marginEl = childOf(sectPr, 'w:pgmar');
  const size = readPageSize(childOf(sectPr, 'w:pgsz'));
  const margins = readMargins(marginEl);
  const page = { ...size, ...readEdgeDistances(marginEl) };
  if (Object.keys(margins).length) page.marginsMm = margins;
  return page;
}

function buildParagraph(p, table, inTable) {
  // 直接格式(段落自己写的 w:ind / w:spacing / w:jc)优先于样式链
  const direct = readParagraphProps(childOf(p, 'w:ppr'));
  const styleId = direct.styleId || table.defaultParagraphId;
  const paragraph = resolveStyle(styleId, 'paragraph', table, table.defaults.paragraph);
  const runDefaults = resolveStyle(styleId, 'run', table, table.defaults.run);
  const format = dominantRunFormat(paragraphRuns(p), runDefaults.props, table);
  const props = { ...paragraph.props, ...direct };
  return {
    index: undefined, // 由 extractDocxFormat 统一编号
    text: trimEnd(paragraphText(p)),
    styleId: styleId || paragraph.name || undefined,
    styleName: paragraph.name,
    alignment: props.alignment,
    lineSpacing: props.lineSpacing,
    spaceBeforePt: props.spaceBeforePt,
    spaceAfterPt: props.spaceAfterPt,
    firstLineChars: props.firstLineChars,
    firstLinePt: props.firstLinePt,
    hangingChars: props.hangingChars,
    leftChars: props.leftChars,
    leftPt: props.leftPt,
    rightChars: props.rightChars,
    rightPt: props.rightPt,
    outlineLevel: props.outlineLevel,
    inTable,
    fontAscii: format.fontAscii,
    fontEastAsia: format.fontEastAsia,
    sizePt: format.sizePt,
    bold: format.bold === true,
    italic: format.italic === true,
    underline: format.underline === true,
  };
}

/**
 * Extract the formatting of a .docx.
 * @param documentXml `word/document.xml`
 * @param stylesXml `word/styles.xml` (optional but needed for inherited values)
 * @param themeXml `word/theme/theme1.xml` (optional; Word templates name fonts by theme)
 */
export function extractDocxFormat(documentXml, stylesXml, themeXml) {
  const table = readStyles(stylesXml, themeXml);
  const tree = parseHtml(documentXml);
  const body = findTag(tree, 'w:body') || tree;
  const sectPr = (body.children || []).findLast((c) => c.tag === 'w:sectpr');

  const paragraphs = [];
  const collect = (node, inTable) => {
    for (const child of node.children || []) {
      if (child.tag === 'w:p') paragraphs.push(buildParagraph(child, table, inTable));
      else if (child.tag === 'w:tbl') collect(child, true);
      else if (child.tag !== 'w:sectpr') collect(child, inTable);
    }
  };
  collect(body, false);
  paragraphs.forEach((paragraph, at) => {
    paragraph.index = at + 1;
  });

  return {
    page: readPage(sectPr),
    defaults: {
      fontAscii: table.defaults.run.fontAscii,
      fontEastAsia: table.defaults.run.fontEastAsia,
      sizePt: table.defaults.run.sizePt,
      lineSpacing: table.defaults.paragraph.lineSpacing,
    },
    paragraphs,
  };
}

/** `固定值 28.8pt` / `1.5 倍` / `单倍` */
export function describeLineSpacing(spacing) {
  if (!spacing) return '继承默认';
  if (spacing.rule !== 'auto') {
    const label = spacing.rule === 'exact' ? '固定值' : '最小值';
    return `${label} ${spacing.pt}pt`;
  }
  return spacing.lines === 1 ? '单倍' : `${spacing.lines} 倍`;
}

function describeIndent(paragraph) {
  if (paragraph.firstLineChars !== undefined) return `首行 ${paragraph.firstLineChars} 字符`;
  if (paragraph.firstLinePt !== undefined) return `首行 ${paragraph.firstLinePt}pt`;
  if (paragraph.hangingChars !== undefined) return `悬挂 ${paragraph.hangingChars} 字符`;
  return '无';
}

function fontSizeLabel(pt) {
  return pt === undefined ? '继承默认' : `${pt}pt`;
}

/** Count values and return them sorted by frequency, then by value. */
function tally(values) {
  const counts = new Map();
  for (const value of values) {
    if (value === undefined || value === '') continue;
    counts.set(value, (counts.get(value) || 0) + 1);
  }
  return [...counts.entries()].sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
}

function signature(paragraph) {
  const font = paragraph.fontEastAsia || paragraph.fontAscii || '默认';
  return `字体=${font}、字号=${fontSizeLabel(paragraph.sizePt)}、行距=${describeLineSpacing(paragraph.lineSpacing)}、缩进=${describeIndent(paragraph)}`;
}

/** `210×297 mm(纵向),页边距 上 37 / … mm` */
function pageSummary(page) {
  if (!page) return '';
  const size = page.widthMm && page.heightMm ? `${page.widthMm}×${page.heightMm} mm` : '未标注';
  const orientation = page.orientation ? `(${page.orientation})` : '';
  const margins = page.marginsMm || {};
  const parts = Object.keys(MARGIN_LABELS)
    .filter((side) => margins[side] !== undefined)
    .map((side) => `${MARGIN_LABELS[side]} ${margins[side]}`);
  const marginText = parts.length ? `,页边距 ${parts.join(' / ')} mm` : '';
  return `**页面**: ${size}${orientation}${marginText}`;
}

function defaultsSummary(defaults) {
  const d = defaults || {};
  const font = [d.fontAscii, d.fontEastAsia].filter(Boolean).join(' / ') || '未标注';
  return `**文档默认**: 字体 ${font},字号 ${fontSizeLabel(d.sizePt)},行距 ${describeLineSpacing(d.lineSpacing)}`;
}

function paragraphRow(paragraph) {
  const cell = (value) => String(value ?? '—').replaceAll('|', String.raw`\|`);
  const text = paragraph.text.length > 24 ? `${paragraph.text.slice(0, 24)}…` : paragraph.text;
  const cells = [
    paragraph.index,
    text,
    cell(paragraph.styleName || paragraph.styleId),
    cell(paragraph.alignment || '继承'),
    describeLineSpacing(paragraph.lineSpacing),
    describeIndent(paragraph),
    cell(paragraph.fontEastAsia),
    cell(paragraph.fontAscii),
    fontSizeLabel(paragraph.sizePt),
    paragraph.bold ? '是' : '—',
  ];
  return `| ${cells.join(' | ')} |`;
}

function distributionLines(body) {
  const groups = [
    ['中文字体', body.map((p) => p.fontEastAsia)],
    ['西文字体', body.map((p) => p.fontAscii)],
    ['字号', body.map((p) => fontSizeLabel(p.sizePt))],
    ['行距', body.map((p) => describeLineSpacing(p.lineSpacing))],
    ['首行缩进', body.map((p) => describeIndent(p))],
    ['对齐', body.map((p) => p.alignment)],
  ];
  const lines = [];
  for (const [label, values] of groups) {
    const top = tally(values).slice(0, 5);
    const summary = top.map(([value, count]) => `${value} ×${count}`).join('、');
    if (summary) lines.push(`- ${label}: ${summary}`);
  }
  return lines;
}

/** Paragraphs whose (font, size, spacing, indent) differs from the mainstream. */
function outliersOf(body) {
  const signatures = tally(body.map(signature));
  const mainstream = signatures.length ? signatures[0][0] : '';
  return { mainstream, list: body.filter((p) => signature(p) !== mainstream) };
}

function outlierLines(outliers) {
  const lines = outliers.list
    .slice(0, MAX_OUTLIERS)
    .map((p) => `- #${p.index}「${p.text.slice(0, 20)}」: ${signature(p)}`);
  if (outliers.list.length > MAX_OUTLIERS) lines.push(`- 其余 ${outliers.list.length - MAX_OUTLIERS} 段见 meta.formatting。`);
  return lines;
}

/**
 * Render a formatting report aimed at comparing against 行文规则: the raw table
 * plus a distribution of the dominant formats and the paragraphs that deviate.
 */
export function formatReport(doc) {
  const body = doc.paragraphs.filter((p) => p.text.trim());
  const lines = [];
  const page = pageSummary(doc.page);
  if (page) lines.push(page);
  const heading = `**段落格式**(共 ${doc.paragraphs.length} 段,有文字 ${body.length} 段):`;
  lines.push(defaultsSummary(doc.defaults), '', heading, '', TABLE_HEAD, TABLE_RULE);
  const shown = body.slice(0, MAX_REPORTED_PARAGRAPHS);
  lines.push(...shown.map(paragraphRow));
  if (body.length > shown.length) lines.push(`> 仅列出前 ${shown.length} 段,余下 ${body.length - shown.length} 段见 meta.formatting。`);
  lines.push('', '**格式分布**(主流值即行文规则比对基准):', ...distributionLines(body));

  const outliers = outliersOf(body);
  if (!outliers.list.length) {
    lines.push('', '**格式一致性**: 所有段落一致,无偏离。');
    return lines.join('\n');
  }
  lines.push('', `**偏离主流格式的段落**: ${outliers.list.length} 段(主流:${outliers.mainstream})`, ...outlierLines(outliers));
  return lines.join('\n');
}
