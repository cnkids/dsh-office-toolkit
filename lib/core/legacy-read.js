// Pure-JS legacy Word readers (no external software, works on Windows/macOS/Linux):
//   .doc  -> word-extractor      .rtf -> built-in RTF parser      .odt -> zip + content.xml
import { readFile } from 'node:fs/promises';
import PizZip from 'pizzip';
import { OfficeError, extOf, isAsciiDigit, assertZipBudget } from './util.js';
import { attrIn, readTagAt } from './markup.js';

const SKIP_DESTS = new Set([
  'fonttbl', 'colortbl', 'stylesheet', 'info', 'pict', 'object', 'themedata', 'datastore',
  'latentstyles', 'rsidtbl', 'generator', 'listtable', 'listoverridetable', 'revtbl', 'xmlnstbl', 'filetbl',
]);

const RTF_STOP = /[^a-zA-Z]/;

// ---------------------------------------------------------------------------
// .doc via word-extractor
// ---------------------------------------------------------------------------
export async function readDocWithWordExtractor(path) {
  let WordExtractor;
  try {
    const mod = await import('word-extractor');
    WordExtractor = mod.default || mod;
  } catch (err) {
    throw new OfficeError('缺少 word-extractor 依赖: ' + (err?.message || err), 'MISSING_DEP');
  }
  const doc = await new WordExtractor().extract(path);
  const chunks = [doc.getBody()].filter((v) => v && String(v).trim());
  for (const [label, getter] of [['页眉/页脚', 'getHeaders'], ['脚注', 'getFootnotes'], ['尾注', 'getEndnotes'], ['文本框', 'getTextboxes']]) {
    const part = readOptionalSection(doc, getter);
    if (part) chunks.push(`【${label}】\n${part}`);
  }
  return chunks.join('\n\n');
}

function readOptionalSection(doc, getter) {
  try {
    const value = typeof doc[getter] === 'function' ? doc[getter]() : '';
    return value && String(value).trim() ? String(value).trim() : '';
  } catch {
    return '';
  }
}

// ---------------------------------------------------------------------------
// .rtf
// ---------------------------------------------------------------------------
function newRtfState() {
  return { out: '', depth: 0, ucSkip: 1, skipStack: [], pending: 0 };
}

function emit(state, text) {
  if (state.skipStack.length === 0) state.out += text;
}

function readControlWord(rtf, start) {
  let i = start + 1;
  if (i >= rtf.length) return { word: '', num: null, next: i };
  const symbol = rtf[i];
  if (symbol === '\\' || symbol === '{' || symbol === '}') return { word: symbol, num: null, next: i + 1, literal: true };
  if (symbol === "'") return { word: 'hex', hex: rtf.slice(i + 1, i + 3), next: i + 3 };
  if (!/[a-zA-Z]/.test(symbol)) return { word: '', num: null, next: i + 1 };
  let j = i;
  while (j < rtf.length && !RTF_STOP.test(rtf[j])) j += 1;
  const word = rtf.slice(i, j);
  const parsed = readSignedNumber(rtf, j);
  return { word, num: parsed.num, next: parsed.next };
}

function readSignedNumber(rtf, start) {
  let i = start;
  let sign = 1;
  if (rtf[i] === '-') { sign = -1; i += 1; }
  let digits = '';
  while (i < rtf.length && /\d/.test(rtf[i])) { digits += rtf[i]; i += 1; }
  if (rtf[i] === ' ') i += 1;
  return { num: digits === '' ? null : sign * Number(digits), next: i };
}

function handleRtfControl(rtf, i, state) {
  const cw = readControlWord(rtf, i);
  if (cw.literal) { emit(state, cw.word); return cw.next; }
  if (cw.word === 'hex') {
    if (/^[0-9a-fA-F]{2}$/.test(cw.hex || '')) emit(state, Buffer.from([Number.parseInt(cw.hex, 16)]).toString('latin1'));
    return cw.next;
  }
  if (cw.word === 'par' || cw.word === 'line' || cw.word === 'page') emit(state, '\n');
  else if (cw.word === 'tab') emit(state, '\t');
  else if (cw.word === 'uc') state.ucSkip = cw.num ?? 1;
  else if (cw.word === 'u' && cw.num !== null) emitUnicode(state, cw.num);
  else if (SKIP_DESTS.has(cw.word)) state.skipStack.push(state.depth + 1);
  return cw.next;
}

function emitUnicode(state, num) {
  emit(state, String.fromCodePoint(num < 0 ? num + 65536 : num));
  state.pending += state.ucSkip;
}

function handleRtfChar(rtf, i, state) {
  if (state.pending > 0) { state.pending -= 1; return i + 1; }
  emit(state, rtf[i]);
  return i + 1;
}

function closeRtfGroup(rtf, i, state) {
  state.depth -= 1;
  while (state.skipStack.length && state.skipStack[state.skipStack.length - 1] > state.depth) state.skipStack.pop();
  return i + 1;
}

/** Minimal RTF -> text (handles \par \line \tab \uN \'hh, skips font/color/pict destinations). */
export function rtfToText(rtf) {
  const state = newRtfState();
  let i = 0;
  while (i < rtf.length) {
    const ch = rtf[i];
    if (ch === '{') { state.depth += 1; i += 1; }
    else if (ch === '}') i = closeRtfGroup(rtf, i, state);
    else if (ch === '\\') i = handleRtfControl(rtf, i, state);
    else i = handleRtfChar(rtf, i, state);
  }
  return state.out.replaceAll('\r', '').replaceAll(/\n{3,}/g, '\n\n').trim();
}

/** Plain text -> minimal RTF (paragraphs only; CJK via \uN escapes). */
export function textToRtf(text) {
  const escaped = String(text)
    .replaceAll('\\', String.raw`\\`)
    .replaceAll('{', String.raw`\{`)
    .replaceAll('}', String.raw`\}`)
    .replaceAll(/[^\x00-\x7F]/g, (ch) => String.raw`\u` + ch.codePointAt(0) + '?');
  const body = escaped.split(/\n{2,}/)
    .map((p) => p.replaceAll('\n', String.raw`\line `))
    .map((p) => p + String.raw`\par`)
    .join('\n');
  return `{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0\\fnil\\fcharset134 Microsoft YaHei;}}\\f0\\fs21\n${body}\n}`;
}

// ---------------------------------------------------------------------------
// .odt
// ---------------------------------------------------------------------------
// Blocks whose content must never reach the output.
const ODT_SKIP_BLOCKS = new Set([
  'office:automatic-styles', 'office:font-face-decls', 'office:scripts', 'text:note',
]);
// Elements rewritten as an HTML open/close pair.
const ODT_PAIRS = new Map([
  ['text:p', ['<p>', '</p>']],
  ['text:list', ['<ul>', '</ul>']],
  ['text:list-item', ['<li>', '</li>']],
  ['table:table', ['<table>', '</table>']],
  ['table:table-row', ['<tr>', '</tr>']],
  ['table:table-cell', ['<td>', '</td>']],
  ['table:covered-table-cell', ['<td></td>', '']],
  ['text:span', ['', '']],
]);
const ODT_VOID_TAGS = new Set(['text:line-break', 'text:tab', 'text:s']);
const ODT_BODY_CLOSE = '</office:body>';
const MAX_ODT_SPACES = 1000;

function clampHeading(level) {
  return Math.min(6, Math.max(1, Number(level) || 1));
}

/** The `<office:body>...</office:body>` slice, or the whole input when absent. */
function odtBody(source) {
  const start = source.indexOf('<office:body');
  if (start === -1) return source;
  const end = source.lastIndexOf(ODT_BODY_CLOSE);
  return end < start ? source : source.slice(start, end + ODT_BODY_CLOSE.length);
}

function escAttr(value) {
  return String(value)
    .replaceAll('&', '&amp;').replaceAll('"', '&quot;')
    .replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

/** `<text:h text:outline-level="2">` -> 2 (single digit, clamped to 1..6). */
function odtHeadingLevel(tag) {
  const raw = String(attrIn(tag.attrs, 'text:outline-level') ?? '');
  const first = raw.length ? raw.codePointAt(0) : -1;
  return clampHeading(isAsciiDigit(first) ? first - 48 : 1);
}

/** `<text:s text:c="3"/>` -> three spaces. */
function odtSpaces(tag) {
  const count = Number.parseInt(attrIn(tag.attrs, 'text:c') ?? '1', 10);
  return ' '.repeat(Number.isFinite(count) && count > 0 && count <= MAX_ODT_SPACES ? count : 1);
}

function odtOpenHtml(tag, headingStack) {
  if (ODT_VOID_TAGS.has(tag.name)) return tag.selfClosing ? odtVoidHtml(tag) : '';
  if (tag.name === 'text:h') {
    const level = odtHeadingLevel(tag);
    headingStack.push(level);
    return `<h${level}>`;
  }
  if (tag.name === 'text:a') {
    const href = attrIn(tag.attrs, 'xlink:href');
    return href === undefined ? '' : `<a href="${escAttr(href)}">`;
  }
  const pair = ODT_PAIRS.get(tag.name);
  return pair ? pair[0] : '';
}

function odtVoidHtml(tag) {
  if (tag.name === 'text:line-break') return '<br/>';
  if (tag.name === 'text:tab') return '\t';
  return odtSpaces(tag);
}

function odtCloseHtml(name, headingStack) {
  if (name === 'text:h') return `</h${headingStack.pop() ?? 1}>`;
  if (name === 'text:a') return '</a>';
  const pair = ODT_PAIRS.get(name);
  return pair ? pair[1] : '';
}

/**
 * ODT content.xml -> HTML fragment (headings, paragraphs, lists, tables, links).
 * One forward pass: no tag regex, so no super-linear backtracking (S5852).
 */
function pushIfAny(out, text) {
  if (text) out.push(text);
}

function isSkipBlock(tag) {
  return ODT_SKIP_BLOCKS.has(tag.name);
}

/** Depth after a skip-block boundary tag (self-closing opens do not nest). */
function nextSkipDepth(depth, tag) {
  if (tag.closing) return Math.max(0, depth - 1);
  return tag.selfClosing ? depth : depth + 1;
}

function odtTagHtml(tag, headingStack) {
  return tag.closing ? odtCloseHtml(tag.name, headingStack) : odtOpenHtml(tag, headingStack);
}

/**
 * Consume one tag plus the text that precedes it. Text and tags inside a skip
 * block are dropped, which is how automatic-styles / notes are removed.
 */
function consumeOdtTag(state, chunk, tag) {
  if (state.skipDepth > 0) {
    if (isSkipBlock(tag)) state.skipDepth = nextSkipDepth(state.skipDepth, tag);
    return;
  }
  pushIfAny(state.out, chunk);
  if (isSkipBlock(tag)) {
    state.skipDepth = nextSkipDepth(0, tag);
    return;
  }
  state.out.push(odtTagHtml(tag, state.headingStack));
}

/**
 * ODT content.xml -> HTML fragment (headings, paragraphs, lists, tables, links).
 * One forward pass: no tag regex, so no super-linear backtracking (S5852).
 */
export function odtXmlToHtml(contentXml) {
  const xml = odtBody(String(contentXml));
  const state = { out: [], headingStack: [], skipDepth: 0 };
  let i = 0;
  let textStart = 0;
  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    if (lt === -1) break;
    const tag = readTagAt(xml, lt);
    if (!tag) { i = lt + 1; continue; }
    consumeOdtTag(state, xml.slice(textStart, lt), tag);
    i = tag.end;
    textStart = i;
  }
  if (state.skipDepth === 0) pushIfAny(state.out, xml.slice(textStart));
  return state.out.join('');
}

export function htmlEscape(text) {
  return String(text).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

export function textToHtml(text) {
  const paras = String(text).split(/\n{2,}/).map((p) => `<p>${htmlEscape(p).replaceAll('\n', '<br/>')}</p>`);
  return '<html><body>' + paras.join('\n') + '</body></html>';
}

export async function readOdt(path, { asHtml = false } = {}) {
  const buf = await readFile(path);
  await assertZipBudget(buf, 'odt');
  let zip;
  try {
    zip = new PizZip(buf);
  } catch (err) {
    throw new OfficeError('不是有效的 .odt 文件: ' + (err?.message || err), 'BAD_ODT');
  }
  const content = zip.file('content.xml')?.asText();
  if (!content) throw new OfficeError('.odt 缺少 content.xml', 'BAD_ODT');
  const html = '<html><body>' + odtXmlToHtml(content) + '</body></html>';
  if (asHtml) return html;
  const { htmlToMarkdown } = await import('./md.js');
  return htmlToMarkdown(html);
}

export async function readRtf(path, { asHtml = false } = {}) {
  const text = rtfToText((await readFile(path)).toString('latin1'));
  return asHtml ? textToHtml(text) : text;
}

/** Dispatch a legacy Word file to the pure-JS reader for its extension. */
export async function jsReadLegacy(path, { asHtml = false } = {}) {
  const ext = extOf(path);
  if (ext === 'doc') {
    const text = await readDocWithWordExtractor(path);
    return asHtml ? textToHtml(text) : text;
  }
  if (ext === 'rtf') return readRtf(path, { asHtml });
  if (ext === 'odt') return readOdt(path, { asHtml });
  throw new OfficeError(`纯 JS 解析不支持 .${ext}`, 'UNSUPPORTED_FORMAT');
}
