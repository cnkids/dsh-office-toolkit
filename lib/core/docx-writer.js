// HTML -> .docx built on docx@9.
//
// This replaces html-to-docx. That library pulled `image-size` (2 high DoS
// advisories with no fixed release) and its maintained fork pulled
// `probe-image-size` -> `needle` plus a `postinstall`, which made
// `dsh plugin add` stop and ask for build-script approval. Generating the
// document here keeps the dependency tree free of install scripts and of any
// network-capable image probing; `<img>` is simply rendered as its alt text.
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  LevelFormat,
  Packer,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  convertMillimetersToTwip,
} from 'docx';
import { attrOf, headingLevel, parseHtml } from './md.js';

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

function isHexDigit(ch) {
  return (ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'f');
}

function isSpaceChar(ch) {
  return ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r' || ch === '\f';
}

/** `#abc` / `abc` / `#AABBCC` -> `aabbcc`, anything else -> undefined. */
function hexColor(value) {
  const raw = String(value ?? '').trim().replace('#', '').toLowerCase();
  if (raw.length === 3) return raw[0] + raw[0] + raw[1] + raw[1] + raw[2] + raw[2];
  if (raw.length !== 6) return undefined;
  for (const ch of raw) {
    if (!isHexDigit(ch)) return undefined;
  }
  return raw;
}

/** CSS declarations of a `style="..."` attribute, lowercased keys. */
function styleMap(node) {
  const out = Object.create(null);
  const raw = attrOf(node, 'style');
  if (!raw) return out;
  for (const decl of String(raw).split(';')) {
    const at = decl.indexOf(':');
    if (at > 0) out[decl.slice(0, at).trim().toLowerCase()] = decl.slice(at + 1).trim();
  }
  return out;
}

/** `14px` / `10.5pt` / `14` -> half-points, else undefined. */
function fontSizeHalfPoints(value) {
  const raw = String(value ?? '').trim().toLowerCase();
  let scale = 1;
  let num = raw;
  if (raw.endsWith('px')) {
    scale = 0.75;
    num = raw.slice(0, -2);
  } else if (raw.endsWith('pt')) {
    num = raw.slice(0, -2);
  }
  const pt = Number.parseFloat(num);
  if (!Number.isFinite(pt) || pt <= 0) return undefined;
  return Math.min(Math.round(pt * scale * 2), 400);
}

function isBoldWeight(value) {
  const raw = String(value ?? '').trim().toLowerCase();
  if (raw === 'bold' || raw === 'bolder') return true;
  const num = Number.parseInt(raw, 10);
  return Number.isFinite(num) && num >= 600;
}

/** Run properties from the element's own `style` attribute. */
function cssRunStyle(node) {
  const css = styleMap(node);
  const style = {};
  const color = hexColor(css.color);
  if (color) style.color = color;
  const fill = hexColor(css['background-color'] || css.background);
  if (fill) style.shading = { fill };
  const size = fontSizeHalfPoints(css['font-size']);
  if (size) style.size = size;
  if (isBoldWeight(css['font-weight'])) style.bold = true;
  if (String(css['font-style'] ?? '').toLowerCase() === 'italic') style.italics = true;
  const decoration = String(css['text-decoration'] ?? '').toLowerCase();
  if (decoration.includes('underline')) style.underline = {};
  if (decoration.includes('line-through')) style.strike = true;
  return style;
}

/** Merge tag semantics with the element's inline CSS on top of inherited style. */
function runStyleOf(node, inherited) {
  const tagStyle = TAG_RUN_STYLE.get(node.tag) || {};
  return { ...inherited, ...tagStyle, ...cssRunStyle(node) };
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

function linkRuns(node, inherited) {
  const children = runsOf(node, runStyleOf(node, inherited));
  const href = safeHref(attrOf(node, 'href'));
  return href ? [new ExternalHyperlink({ children, link: href })] : children;
}

function inlineOf(node, inherited) {
  if (node.tag === '#text') return textRuns(node.text, inherited);
  if (SKIP_TAGS.has(node.tag) || BLOCK_ONLY_TAGS.has(node.tag)) return [];
  if (node.tag === 'br') return [new TextRun({ ...inherited, break: 1 })];
  if (node.tag === 'img') return altRuns(node, inherited);
  if (node.tag === 'a') return linkRuns(node, inherited);
  return runsOf(node, runStyleOf(node, inherited));
}

/** Inline content of a node as an array of TextRun / ExternalHyperlink. */
function runsOf(node, inherited = {}) {
  const out = [];
  for (const child of node.children || []) out.push(...inlineOf(child, inherited));
  return out;
}

/** `text-align` from a `style` attribute, else undefined. */
function alignmentOf(node) {
  const value = String(styleMap(node)['text-align'] ?? '').toLowerCase();
  if (value === 'center') return AlignmentType.CENTER;
  if (value === 'right' || value === 'end') return AlignmentType.RIGHT;
  if (value === 'justify') return AlignmentType.JUSTIFIED;
  if (value === 'left' || value === 'start') return AlignmentType.LEFT;
  return undefined;
}

/** Wrap a node's inline children in a paragraph. */
function paragraphFromNode(node, extra) {
  const alignment = alignmentOf(node);
  const paragraph = { ...extra, children: runsOf(node, cssRunStyle(node)) };
  if (alignment) paragraph.alignment = alignment;
  return new Paragraph(paragraph);
}

/** A standalone `<img>` at block level keeps its alt text as a paragraph. */
function imageBlock(node, extra) {
  const runs = altRuns(node, {});
  return runs.length ? [new Paragraph({ ...extra, children: runs })] : [];
}

function preParagraph(node, extra) {
  const text = textOf(node).replaceAll('\r\n', '\n');
  return new Paragraph({ ...extra, children: textRuns(text, { font: CODE_FONT, size: DEFAULT_SIZE }) });
}

function textBlock(text, extra) {
  const trimmed = String(text).trim();
  return trimmed ? [new Paragraph({ ...extra, children: textRuns(trimmed, {}) })] : [];
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

function listItems(node, depth, extra) {
  const reference = node.tag === 'ol' ? LIST_REFERENCE.ordered : LIST_REFERENCE.bullet;
  const level = Math.min(depth, 8);
  const out = [];
  for (const item of node.children || []) {
    if (item.tag !== 'li') continue;
    out.push(new Paragraph({ ...extra, numbering: { reference, level }, children: runsOf(item, {}) }));
    for (const nested of item.children || []) {
      if (nested.tag === 'ul' || nested.tag === 'ol') out.push(...listItems(nested, depth + 1, extra));
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

function rowOf(row) {
  const cells = [];
  for (const cell of row.children || []) {
    if (cell.tag !== 'td' && cell.tag !== 'th') continue;
    const header = cell.tag === 'th';
    cells.push(
      new TableCell({
        children: cellParagraphs(cell, header ? { bold: true } : {}),
        shading: header ? { fill: 'F2F2F2' } : undefined,
      })
    );
  }
  return new TableRow({ children: cells });
}

function tableOf(node) {
  const rows = [];
  const walk = (parent) => {
    for (const child of parent.children || []) {
      if (child.tag === 'tr') rows.push(rowOf(child));
      else if (child.tag !== '#text') walk(child);
    }
  };
  walk(node);
  if (!rows.length) return new Paragraph('');
  return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } });
}

function blockChildrenOf(node, depth, extra) {
  const out = [];
  for (const child of node.children || []) out.push(...blockOf(child, depth, extra));
  return out;
}

function blockOf(node, depth, extra) {
  const tag = node.tag;
  if (tag === '#text') return textBlock(node.text, extra);
  if (SKIP_TAGS.has(tag)) return [];
  const level = headingLevel(tag);
  if (level) return [new Paragraph({ ...extra, heading: HEADINGS[level - 1], children: runsOf(node, {}) })];
  if (TRANSPARENT_TAGS.has(tag)) return blockChildrenOf(node, depth, extra);
  if (tag === 'ul' || tag === 'ol') return listItems(node, depth, extra);
  if (tag === 'table') return [tableOf(node)];
  if (tag === 'blockquote') return blockChildrenOf(node, depth, { ...extra, indent: { left: QUOTE_INDENT } });
  if (tag === 'pre') return [preParagraph(node, extra)];
  if (tag === 'hr') return [new Paragraph({ ...extra, border: RULE })];
  if (tag === 'img') return imageBlock(node, extra);
  if (tag === 'br') return [];
  return [paragraphFromNode(node, extra)];
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

/**
 * Build a .docx buffer from an HTML string.
 * @param html HTML fragment or full document
 * @param options { title?, landscape?, marginsMm? }
 */
export async function htmlToDocxBuffer(html, options = {}) {
  const children = blockChildrenOf(parseHtml(html), 0, {});
  const doc = new Document({
    title: options.title,
    styles: { default: { document: { run: { font: DEFAULT_FONT, size: DEFAULT_SIZE } } } },
    numbering: NUMBERING,
    sections: [
      {
        properties: pageOf(Boolean(options.landscape), options.marginsMm),
        children: children.length ? children : [new Paragraph('')],
      },
    ],
  });
  return Packer.toBuffer(doc);
}
