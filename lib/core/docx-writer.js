// HTML -> .docx built on docx@9.
//
// This replaces html-to-docx. That library pulled `image-size` (2 high DoS
// advisories with no fixed release) and its maintained fork pulled
// `probe-image-size` -> `needle` plus a `postinstall`, which made
// `dsh plugin add` stop and ask for build-script approval. Generating the
// document here keeps the dependency tree free of install scripts and of any
// network-capable image probing; `<img>` 由 image.js 按魔数读尺寸后嵌入
// (本地文件或 data: URL,不联网)。
import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  ExternalHyperlink,
  HeadingLevel,
  LevelFormat,
  HeightRule,
  LineRuleType,
  TableLayoutType,
  Packer,
  ImageRun,
  PageNumber,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TableOfContents,
  TextRun,
  WidthType,
  convertMillimetersToTwip,
} from 'docx';
import { attrOf, headingLevel, parseHtml } from './md.js';
import { OfficeError, isPlainObject } from './util.js';
import { fitImageSize, heightHintOf, loadImage, widthHintOf } from './image.js';
import { alignmentOf, autoColumnPercents, declarationsFrom, hexColor, paragraphStyleFromCss, runStyleFromCss, normalizeStyleSpec, isEmptyStyle } from './docx-style.js';

const HEADINGS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
];

const DEFAULT_FONT = '等线';
const CODE_FONT = 'Consolas';
const DEFAULT_SIZE = 21; // half-points -> 10.5pt
const DEFAULT_MARGIN_MM = 25;
const QUOTE_INDENT = 480;
const LIST_REFERENCE = { bullet: 'dsh-bullet', ordered: 'dsh-ordered' };

/** Dropped entirely: never rendered in a document. */
const SKIP_TAGS = new Set(['script', 'style', 'head', 'title', 'meta', 'link', 'svg', 'iframe', 'object', 'embed', 'template']);

/** Containers that only group blocks: their children are rendered directly. */
const TRANSPARENT_TAGS = new Set(['root', 'html', 'body', 'div', 'section', 'article', 'main', 'header', 'footer', 'nav', 'figure', 'details', 'summary', 'center']);

/** Handled by the block walker; dropped when met inside inline content. */
const BLOCK_ONLY_TAGS = new Set(['table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'ul', 'ol', 'li', 'blockquote', 'pre', 'p', 'hr']);

const TAG_RUN_STYLE = new Map([
  ['strong', { bold: true }],
  ['b', { bold: true }],
  ['em', { italics: true }],
  ['i', { italics: true }],
  ['u', { underline: {} }],
  ['ins', { underline: {} }],
  ['s', { strike: true }],
  ['strike', { strike: true }],
  ['del', { strike: true }],
  ['code', { font: CODE_FONT }],
  ['kbd', { font: CODE_FONT }],
  ['samp', { font: CODE_FONT }],
  ['tt', { font: CODE_FONT }],
  ['sub', { subScript: true }],
  ['sup', { superScript: true }],
]);

const RULE = { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 1 } };

/** 规范形式里的对齐 → docx 的 AlignmentType。 */
const ALIGNMENT = {
  both: AlignmentType.JUSTIFIED,
  center: AlignmentType.CENTER,
  left: AlignmentType.LEFT,
  right: AlignmentType.RIGHT,
};

function isSpaceChar(ch) {
  return ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r' || ch === '\f';
}

/** `style="..."` 属性 → 声明表。 */
function declsOf(node) {
  return declarationsFrom(attrOf(node, 'style'));
}

/** 规范形式里的字符样式 → docx@9 的 run 选项。 */
function toRunOptions(norm) {
  const out = {};
  if (norm.font || norm.fontAscii) {
    const west = norm.fontAscii || norm.font;
    out.font = { ascii: west, hAnsi: west, eastAsia: norm.font || west, cs: west };
  }
  if (norm.sizeHalfPt) out.size = norm.sizeHalfPt;
  if (norm.bold !== undefined) out.bold = norm.bold;
  if (norm.italics) out.italics = true;
  if (norm.underline) out.underline = {};
  if (norm.strike) out.strike = true;
  if (norm.color) out.color = norm.color;
  return out;
}

/** 框线模式 → docx@9 的表格边框选项。 */
function tableBorders(mode) {
  const line = { style: BorderStyle.SINGLE, size: 4, color: 'auto' };
  const thick = { style: BorderStyle.SINGLE, size: 12, color: 'auto' };
  const none = { style: BorderStyle.NONE, size: 0, color: 'auto' };
  if (mode === 'all') return { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line };
  if (mode === 'outline') return { top: line, bottom: line, left: line, right: line, insideHorizontal: none, insideVertical: none };
  if (mode === 'three-line') return { top: thick, bottom: thick, left: none, right: none, insideHorizontal: none, insideVertical: none };
  return { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };
}

/** 每列内容的最大字符数(用于按内容分配列宽)。 */
function columnLengths(rows, cols) {
  const lengths = new Array(cols).fill(0);
  rows.forEach((row) => {
    let at = 0;
    for (const cell of row.children || []) {
      if (cell.tag !== 'td' && cell.tag !== 'th') continue;
      if (at < cols) lengths[at] = Math.max(lengths[at], textOf(cell).trim().length);
      at += 1;
    }
  });
  return lengths;
}

/** 列宽百分比:manual 用给定值,auto 按内容长度分配。 */
function resolveTablePercents(norm, rows, cols) {
  if (!norm.columnMode) return null;
  if (norm.columnMode === 'manual') {
    const given = norm.columnPercents ?? [];
    return Array.from({ length: cols }, (_, i) => given[i] ?? Math.max(1, Math.round(100 / cols)));
  }
  return autoColumnPercents(columnLengths(rows, cols));
}

/** 规范形式里的段落样式 → docx@9 的段落选项。 */
function toParagraphOptions(norm) {
  const out = {};
  if (norm.align) out.alignment = ALIGNMENT[norm.align];
  const spacing = {};
  if (norm.line !== undefined) {
    spacing.line = norm.line;
    spacing.lineRule = norm.lineRule === 'exact' ? LineRuleType.EXACT : LineRuleType.AUTO;
  }
  if (norm.beforeTwips !== undefined) spacing.before = norm.beforeTwips;
  if (norm.afterTwips !== undefined) spacing.after = norm.afterTwips;
  if (Object.keys(spacing).length) out.spacing = spacing;
  if (norm.firstLineTwips !== undefined) out.indent = { firstLine: norm.firstLineTwips };
  return out;
}

/** 段落级选项的深合并(indent / spacing 各自是对象,不能整体覆盖)。 */
function mergeParagraphOptions(base, add) {
  const out = { ...base, ...add };
  if (base.indent && add.indent) out.indent = { ...base.indent, ...add.indent };
  if (base.spacing && add.spacing) out.spacing = { ...base.spacing, ...add.spacing };
  return out;
}

/** 元素的字符级内联样式(含底纹;tag 语义由调用方再叠加)。 */
function cssRunOptions(node) {
  const decls = declsOf(node);
  const out = toRunOptions(runStyleFromCss(decls));
  const fill = hexColor(decls['background-color'] || decls.background);
  if (fill) out.shading = { fill };
  return out;
}

/** 元素的段落级内联样式;em/ch 缩进按该元素(或继承)的字号折算。 */
function cssParagraphOptions(node, inherited) {
  const decls = declsOf(node);
  const run = runStyleFromCss(decls);
  const em = run.sizeHalfPt ?? inherited?.size ?? DEFAULT_SIZE;
  return toParagraphOptions(paragraphStyleFromCss(decls, em));
}

/** Merge tag semantics with the element's inline CSS on top of inherited style. */
function runStyleOf(node, inherited) {
  return { ...inherited, ...TAG_RUN_STYLE.get(node.tag), ...cssRunOptions(node) };
}

/** Split on newlines so they survive as real line breaks. */
function textRuns(text, style) {
  const lines = String(text).split('\n');
  const runs = [];
  lines.forEach((line, index) => {
    if (index > 0) runs.push(new TextRun({ ...style, break: 1 }));
    if (line) runs.push(new TextRun({ ...style, text: line }));
  });
  return runs;
}

/** All text below a node, newlines preserved (used for `<pre>`). */
function textOf(node) {
  if (node.tag === '#text') return node.text;
  let out = '';
  for (const child of node.children || []) out += textOf(child);
  return out;
}

/** Only well-known safe schemes become hyperlinks; text is kept either way. */
function safeHref(value) {
  const href = String(value ?? '').trim();
  if (!href) return '';
  const lower = href.toLowerCase();
  const banned = ['javascript:', 'data:', 'vbscript:', 'file:'];
  for (const scheme of banned) {
    if (lower.startsWith(scheme)) return '';
  }
  for (const ch of href.split('://')[0]) {
    if (ch === ':' || isSpaceChar(ch)) return '';
  }
  return href;
}

function altRuns(node, inherited) {
  const alt = attrOf(node, 'alt');
  return alt ? [new TextRun({ ...inherited, text: alt })] : [];
}

/** 已解析好的图片 → ImageRun(未解析的图片退回 alt 文字)。 */
function imageRuns(node) {
  const image = node.image;
  if (!image) return [];
  const alt = attrOf(node, 'alt');
  return [new ImageRun({
    type: image.type,
    data: image.data,
    transformation: { width: image.displayWidth, height: image.displayHeight },
    ...(alt ? { altText: { name: alt, description: alt, title: alt } } : {}),
  })];
}

function linkRuns(node, inherited) {
  const children = runsOf(node, runStyleOf(node, inherited));
  const href = safeHref(attrOf(node, 'href'));
  return href ? [new ExternalHyperlink({ children, link: href })] : children;
}

function inlineOf(node, inherited) {
  if (node.tag === '#text') return textRuns(node.text, inherited);
  if (SKIP_TAGS.has(node.tag) || BLOCK_ONLY_TAGS.has(node.tag)) return [];
  if (node.tag === 'br') return [new TextRun({ ...inherited, break: 1 })];
  if (node.tag === 'img') return node.image ? imageRuns(node) : altRuns(node, inherited);
  if (node.tag === 'a') return linkRuns(node, inherited);
  return runsOf(node, runStyleOf(node, inherited));
}

/** Inline content of a node as an array of TextRun / ExternalHyperlink. */
function runsOf(node, inherited = {}) {
  const out = [];
  for (const child of node.children || []) out.push(...inlineOf(child, inherited));
  return out;
}

/**
 * Wrap a node's inline children in a paragraph, applying the element's own
 * inline CSS on top of the inherited run style and the caller's block options.
 */
function paragraphFromNode(node, extra, runStyle = {}) {
  const options = mergeParagraphOptions(extra, cssParagraphOptions(node, runStyle));
  options.children = runsOf(node, { ...runStyle, ...cssRunOptions(node) });
  return new Paragraph(options);
}

/** 独占一段的 `<img>`:能嵌入就嵌图片,嵌不了就退回 alt 文字。 */
function imageBlock(node, extra, runStyle = {}) {
  const runs = node.image ? imageRuns(node) : altRuns(node, runStyle);
  return runs.length ? [new Paragraph({ ...extra, children: runs })] : [];
}

function preParagraph(node, extra) {
  const text = textOf(node).replaceAll('\r\n', '\n');
  return new Paragraph({ ...extra, children: textRuns(text, { font: CODE_FONT, size: DEFAULT_SIZE }) });
}

function textBlock(text, extra, runStyle = {}) {
  const trimmed = String(text).trim();
  return trimmed ? [new Paragraph({ ...extra, children: textRuns(trimmed, runStyle) })] : [];
}

function numberLevels(kind) {
  return [0, 1, 2, 3, 4, 5, 6, 7, 8].map((level) => ({
    level,
    format: kind === 'ordered' ? LevelFormat.DECIMAL : LevelFormat.BULLET,
    text: kind === 'ordered' ? `%${level + 1}.` : '•',
    alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 360 + level * 360, hanging: 360 } } },
  }));
}

const NUMBERING = {
  config: [
    { reference: LIST_REFERENCE.bullet, levels: numberLevels('bullet') },
    { reference: LIST_REFERENCE.ordered, levels: numberLevels('ordered') },
  ],
};

function listItems(node, depth, extra, runStyle = {}) {
  const reference = node.tag === 'ol' ? LIST_REFERENCE.ordered : LIST_REFERENCE.bullet;
  const level = Math.min(depth, 8);
  const out = [];
  for (const item of node.children || []) {
    if (item.tag !== 'li') continue;
    const options = mergeParagraphOptions(extra, cssParagraphOptions(item, runStyle));
    options.numbering = { reference, level };
    options.children = runsOf(item, { ...runStyle, ...cssRunOptions(item) });
    out.push(new Paragraph(options));
    for (const nested of item.children || []) {
      if (nested.tag === 'ul' || nested.tag === 'ol') out.push(...listItems(nested, depth + 1, extra, runStyle));
    }
  }
  return out;
}

/** Cell content as paragraphs: one per inner block, else the inline content. */
function cellParagraphs(cell, bold) {
  const out = [];
  for (const child of cell.children || []) {
    if (child.tag !== 'p' && child.tag !== 'div') continue;
    const runs = runsOf(child, bold);
    if (runs.length) out.push(new Paragraph({ children: runs }));
  }
  if (out.length) return out;
  const runs = runsOf(cell, bold);
  return [new Paragraph({ children: runs })];
}

const DEFAULT_HEADER_FILL = 'F2F2F2';

function tableRowNodes(node) {
  const rows = [];
  const walk = (parent) => {
    for (const child of parent.children || []) {
      if (child.tag === 'tr') rows.push(child);
      else if (child.tag !== '#text') walk(child);
    }
  };
  walk(node);
  return rows;
}

function rowCellNodes(row) {
  return (row.children || []).filter((cell) => cell.tag === 'td' || cell.tag === 'th');
}

/** 表头单元格:底纹与加粗按表格规格来(默认灰底 + 加粗,与旧行为一致)。 */
function headerCellOptions(cell, norm) {
  const options = { children: cellParagraphs(cell, norm.headerBold === false ? {} : { bold: true }) };
  const fill = norm.headerShading === undefined ? DEFAULT_HEADER_FILL : norm.headerShading;
  if (fill) options.shading = { fill };
  if (norm.borders === 'three-line') {
    options.borders = { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'auto' } };
  }
  return options;
}

function cellOptions(cell, norm, isHeader) {
  const options = isHeader ? headerCellOptions(cell, norm) : { children: cellParagraphs(cell, {}) };
  if (norm.cellVerticalAlign) options.verticalAlign = norm.cellVerticalAlign;
  return options;
}

function rowOf(row, norm, isHeader) {
  const cells = rowCellNodes(row).map((cell) => (
    cell.tag === 'th' ? new TableCell(cellOptions(cell, norm, true)) : new TableCell(cellOptions(cell, norm, false))
  ));
  const options = { children: cells };
  if (isHeader && norm.repeatHeader) options.tableHeader = true;
  if (norm.cantSplit) options.cantSplit = true;
  if (norm.rowHeightTwips) options.height = { value: norm.rowHeightTwips, rule: HeightRule.ATLEAST };
  return new TableRow(options);
}

function tableOf(node, table) {
  const rowNodes = tableRowNodes(node);
  if (!rowNodes.length) return new Paragraph('');
  const norm = table.norm ?? {};
  const cols = rowNodes.reduce((max, row) => Math.max(max, rowCellNodes(row).length), 1);
  const percents = resolveTablePercents(norm, rowNodes, cols);
  const options = {
    rows: rowNodes.map((row, index) => rowOf(row, norm, index === 0)),
    width: { size: 100, type: WidthType.PERCENTAGE },
  };
  if (percents) {
    options.columnWidths = percents.map((p) => Math.max(1, Math.round((p / 100) * table.textWidth)));
    options.layout = TableLayoutType.FIXED;
  }
  if (norm.align) options.alignment = ALIGNMENT[norm.align];
  if (norm.borders) options.borders = tableBorders(norm.borders);
  if (norm.cellMargins) options.margins = { marginUnitType: WidthType.DXA, ...norm.cellMargins };
  return new Table(options);
}

/**
 * 先把树里的 `<img>` 解析成可嵌入的图片(读文件、读尺寸、按正文宽度定显示大小)。
 * 没有 src 的图片保持原样,后面按 alt 文字处理。
 */
async function resolveImages(node, options) {
  for (const child of node.children || []) {
    if (child.tag !== 'img') {
      await resolveImages(child, options);
      continue;
    }
    const src = attrOf(child, 'src');
    if (!src) continue;
    const image = await loadImage(src, options.readImage);
    const declarations = declarationsFrom(attrOf(child, 'style'));
    const readAttr = (name) => attrOf(child, name);
    const size = fitImageSize(image, {
      widthHint: widthHintOf(readAttr, declarations, options.textWidthPx),
      heightHint: heightHintOf(readAttr, declarations),
      maxWidth: options.textWidthPx,
      maxHeight: options.maxHeightPx,
    });
    child.image = { ...image, displayWidth: size.width, displayHeight: size.height };
  }
}

function blockChildrenOf(node, depth, extra, runStyle = {}, table = {}) {
  const out = [];
  for (const child of node.children || []) out.push(...blockOf(child, depth, extra, runStyle, table));
  return out;
}

/**
 * `<body style="font-family:仿宋">` 这种写法要把样式传下去,否则内联样式只能写在
 * 每个 `<p>` 上 —— 所以透明容器(<body>/<div>/<section>…)把自己的 CSS 合并进上下文。
 */
function inheritedFrom(node, extra, runStyle) {
  return {
    extra: mergeParagraphOptions(extra, cssParagraphOptions(node, runStyle)),
    runStyle: { ...runStyle, ...cssRunOptions(node) },
  };
}

function blockOf(node, depth, extra, runStyle = {}, table = {}) {
  const tag = node.tag;
  if (tag === '#text') return textBlock(node.text, extra, runStyle);
  if (SKIP_TAGS.has(tag)) return [];
  const level = headingLevel(tag);
  if (level) {
    const options = mergeParagraphOptions(extra, cssParagraphOptions(node, runStyle));
    options.heading = HEADINGS[level - 1];
    options.children = runsOf(node, { ...runStyle, ...cssRunOptions(node) });
    return [new Paragraph(options)];
  }
  if (TRANSPARENT_TAGS.has(tag)) {
    const inherited = inheritedFrom(node, extra, runStyle);
    return blockChildrenOf(node, depth, inherited.extra, inherited.runStyle, table);
  }
  if (tag === 'ul' || tag === 'ol') return listItems(node, depth, extra, runStyle);
  if (tag === 'table') return [tableOf(node, table)];
  if (tag === 'blockquote') return blockChildrenOf(node, depth, mergeParagraphOptions(extra, { indent: { left: QUOTE_INDENT } }), runStyle, table);
  if (tag === 'pre') return [preParagraph(node, extra)];
  if (tag === 'hr') return [new Paragraph({ ...extra, border: RULE })];
  if (tag === 'img') return imageBlock(node, extra, runStyle);
  if (tag === 'br') return [];
  return [paragraphFromNode(node, extra, runStyle)];
}

/**
 * A4 with the requested margins. Width/height are always given in portrait
 * order: docx swaps them itself when the orientation is landscape.
 */
function pageOf(landscape, marginsMm) {
  const width = convertMillimetersToTwip(210);
  const height = convertMillimetersToTwip(297);
  const margin = convertMillimetersToTwip(Number.isFinite(marginsMm) && marginsMm > 0 ? marginsMm : DEFAULT_MARGIN_MM);
  return {
    page: {
      size: {
        orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
        width,
        height,
      },
      margin: { top: margin, bottom: margin, left: margin, right: margin },
    },
  };
}

const HEADING_STYLE_KEYS = ['heading1', 'heading2', 'heading3', 'heading4', 'heading5', 'heading6'];

// ---------------------------------------------------------------------------
// 页眉 / 页脚 / 页码 / 目录
// ---------------------------------------------------------------------------
function alignOption(value) {
  const align = alignmentOf(value);
  return align ? ALIGNMENT[align] : undefined;
}

function runSize(fontSizePt) {
  if (fontSizePt === undefined || fontSizePt === null || fontSizePt === '') return undefined;
  const pt = Number(fontSizePt);
  if (!Number.isFinite(pt) || pt <= 0 || pt > 200) {
    throw new OfficeError(`页眉页脚的 fontSizePt 需为 0–200 的磅值，收到 ${JSON.stringify(fontSizePt)}`, 'INVALID_ARGS');
  }
  return Math.round(pt * 2);
}

/** `第 {page} 页 共 {total} 页` → 文字与域交替的 run 序列。 */
function pageNumberRuns(template, base) {
  const text = template === true ? '第 {page} 页 共 {total} 页' : String(template);
  if (!text.includes('{page}') && !text.includes('{total}')) {
    throw new OfficeError('pageNumber 模板里要用 {page} 表示当前页、{total} 表示总页数，例如 "第 {page} 页 共 {total} 页"', 'INVALID_ARGS');
  }
  const out = [];
  for (const part of text.split(/(\{page\}|\{total\})/)) {
    if (!part) continue;
    if (part === '{page}') out.push(new TextRun({ ...base, children: [PageNumber.CURRENT] }));
    else if (part === '{total}') out.push(new TextRun({ ...base, children: [PageNumber.TOTAL_PAGES] }));
    else out.push(new TextRun({ ...base, text: part }));
  }
  return out;
}

/** 页眉/页脚规格 → Header / Footer(没给规格时返回 undefined)。 */
function headerFooterOf(spec, kind) {
  if (spec === undefined || spec === null) return undefined;
  if (!isPlainObject(spec)) {
    throw new OfficeError(`${kind} 需为对象，如 {"text":"XX公司文件","align":"center"}`, 'INVALID_ARGS');
  }
  const base = { bold: Boolean(spec.bold), ...(runSize(spec.fontSizePt) ? { size: runSize(spec.fontSizePt) } : {}) };
  const runs = [];
  if (spec.text) runs.push(new TextRun({ ...base, text: String(spec.text) }));
  if (spec.pageNumber) runs.push(...pageNumberRuns(spec.pageNumber, base));
  if (!runs.length) {
    throw new OfficeError(`${kind} 至少要给 text 或 pageNumber 之一`, 'INVALID_ARGS');
  }
  const paragraph = new Paragraph({ children: runs, alignment: alignOption(spec.align) });
  return kind === 'header' ? new Header({ children: [paragraph] }) : new Footer({ children: [paragraph] });
}

/** `toc: true` 或 `{title, levels}` → TableOfContents。 */
function tocOf(spec) {
  if (!spec) return undefined;
  const options = spec === true ? {} : spec;
  if (!isPlainObject(options)) throw new OfficeError('toc 需为 true 或对象，如 {"title":"目录","levels":3}', 'INVALID_ARGS');
  const levels = options.levels === undefined ? 3 : Number(options.levels);
  if (!Number.isInteger(levels) || levels < 1 || levels > 9) {
    throw new OfficeError('toc.levels 需为 1–9 的整数', 'INVALID_ARGS');
  }
  return new TableOfContents(options.title === undefined ? '目录' : String(options.title), {
    hyperlink: true,
    headingStyleRange: `1-${levels}`,
  });
}

/** 文档默认样式 + 各级标题样式(没有对应字段就沿用 docx 自己的默认)。 */
function defaultStyles(styleNorm) {
  const run = { font: DEFAULT_FONT, size: DEFAULT_SIZE, ...toRunOptions(styleNorm) };
  const document = { run, paragraph: toParagraphOptions(styleNorm) };
  if (!Object.keys(document.paragraph).length) delete document.paragraph;
  const out = { document };
  const headings = styleNorm.headings;
  if (!isEmptyStyle(headings)) {
    const headingRun = toRunOptions(headings);
    const headingParagraph = toParagraphOptions(headings);
    for (const key of HEADING_STYLE_KEYS) {
      out[key] = {
        ...(Object.keys(headingRun).length ? { run: headingRun } : {}),
        ...(Object.keys(headingParagraph).length ? { paragraph: headingParagraph } : {}),
      };
    }
  }
  return out;
}

/**
 * Build a .docx buffer from an HTML string.
 * @param html HTML fragment or full document
 * @param options { title?, landscape?, marginsMm?, style? } style 为文档级排版规格
 */
export async function htmlToDocxBuffer(html, options = {}) {
  const styleNorm = normalizeStyleSpec(options.style);
  const tableNorm = styleNorm.table ?? {}; // normalizeStyleSpec 已经规范化过,别再规范化一次
  const textWidth = Math.round(convertMillimetersToTwip(Math.max(20, 210 - 2 * (Number.isFinite(options.marginsMm) && options.marginsMm > 0 ? options.marginsMm : DEFAULT_MARGIN_MM))));
  const tree = parseHtml(html);
  // 正文宽度换算成 px(1px = 1/96 英寸):图片按它缩放,不会超出页面
  const textWidthPx = Math.max(60, Math.round(textWidth / 15));
  await resolveImages(tree, {
    readImage: options.readImage,
    textWidthPx,
    maxHeightPx: 900,
  });
  const children = blockChildrenOf(tree, 0, {}, {}, { norm: tableNorm, textWidth });
  const header = headerFooterOf(options.header, 'header');
  const footer = headerFooterOf(options.footer, 'footer');
  const toc = tocOf(options.toc);
  if (toc) children.unshift(toc);
  const doc = new Document({
    title: options.title,
    // 让 Word 打开时更新域:目录与页码会自动按最新排版刷新
    features: { updateFields: true },
    styles: { default: defaultStyles(styleNorm) },
    numbering: NUMBERING,
    sections: [
      {
        properties: pageOf(Boolean(options.landscape), options.marginsMm),
        ...(header ? { headers: { default: header } } : {}),
        ...(footer ? { footers: { default: footer } } : {}),
        children: children.length ? children : [new Paragraph('')],
      },
    ],
  });
  return Packer.toBuffer(doc);
}
