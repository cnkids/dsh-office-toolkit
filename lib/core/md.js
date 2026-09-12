// Markdown -> HTML (marked) and HTML -> Markdown-ish text (tiny tree walker),
// plus plain text -> HTML. The HTML side must handle what mammoth emits:
// h1-h6, p, ul/ol/li, table/tr/td/th, strong/em, a, br, img, blockquote, pre/code.
//
// The tree walker uses the linear scanners from markup.js rather than tag
// regexes: patterns such as /<[^>]+>/g or /[ \t]+\n/g backtrack super-linearly
// on crafted input and are a DoS vector (SonarQube S5852).
import { marked } from 'marked';
import { OfficeError } from './util.js';
import { attrIn, readTagAt } from './markup.js';

marked.setOptions({ gfm: true, breaks: false });

export function markdownToHtml(md) {
  if (typeof md !== 'string' || !md.trim()) throw new OfficeError('markdown 内容为空');
  return String(marked.parse(md, { async: false }));
}

const VOID_TAGS = new Set(['br', 'img', 'hr', 'input', 'meta', 'link']);
const INLINE_TAGS = new Set(['strong', 'b', 'em', 'i', 'u', 'code', 'a', 'span', 'sub', 'sup', 'del', 's', 'mark']);

const CODE_DIGIT_1 = 49;
const CODE_DIGIT_6 = 54;
const HEAD_CLOSE = '</head>';

export function decodeEntities(s) {
  return String(s)
    .replaceAll('&nbsp;', ' ')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&#39;', "'")
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&');
}

function escapePipes(s) {
  return String(s).replaceAll('|', String.raw`\|`);
}

// ---------------------------------------------------------------------------
// linear HTML tokenizer
// ---------------------------------------------------------------------------
const NOISE_PAIRS = [['<!--', '-->'], ['<![CDATA[', ']]>']];

/** End index just past a comment / CDATA / `<!...>` declaration, else -1. */
function markupNoiseEnd(source, at) {
  for (const [open, close] of NOISE_PAIRS) {
    if (!source.startsWith(open, at)) continue;
    const end = source.indexOf(close, at + open.length);
    return end === -1 ? -1 : end + close.length;
  }
  if (source[at + 1] !== '!') return -1;
  const end = source.indexOf('>', at + 2);
  return end === -1 ? -1 : end + 1;
}

function pushText(stack, text) {
  if (!text) return;
  stack.at(-1).children.push({ tag: '#text', text: decodeEntities(text) });
}

function openNode(stack, tag) {
  const node = { tag: tag.name.toLowerCase(), attrs: tag.attrs, children: [] };
  stack.at(-1).children.push(node);
  if (!VOID_TAGS.has(node.tag) && !tag.selfClosing) stack.push(node);
}

function closeNode(stack, name) {
  const tag = name.toLowerCase();
  for (let i = stack.length - 1; i > 0; i -= 1) {
    if (stack[i].tag === tag) { stack.length = i; return; }
  }
}

/** Parse an HTML fragment into a minimal node tree. */
export function parseHtml(html) {
  const source = String(html);
  const root = { tag: 'root', attrs: '', children: [] };
  const stack = [root];
  let i = 0;
  let textStart = 0;
  while (i < source.length) {
    if (source[i] !== '<') { i += 1; continue; }
    const noiseEnd = markupNoiseEnd(source, i);
    if (noiseEnd !== -1) {
      pushText(stack, source.slice(textStart, i));
      i = noiseEnd;
      textStart = i;
      continue;
    }
    const tag = readTagAt(source, i);
    if (!tag) { i += 1; continue; }
    pushText(stack, source.slice(textStart, i));
    if (tag.closing) closeNode(stack, tag.name);
    else openNode(stack, tag);
    i = tag.end;
    textStart = i;
  }
  pushText(stack, source.slice(textStart));
  return root;
}

export function attrOf(node, name) {
  const value = attrIn(node.attrs, name);
  return value === undefined ? '' : decodeEntities(value);
}

function collapse(s) {
  return String(s).replaceAll(/[ \t\r\n]+/g, ' ').trim();
}

/** Drop trailing newlines without a backtracking `/\n+$/` scan. */
function trimTrailingNewlines(s) {
  let end = s.length;
  while (end > 0 && s[end - 1] === '\n') end -= 1;
  return s.slice(0, end);
}

/** Drop spaces/tabs that sit immediately before a newline (`/[ \t]+\n/g`). */
function trimBeforeNewline(s) {
  const parts = [];
  let i = 0;
  let plainStart = 0;
  while (i < s.length) {
    if (s[i] !== ' ' && s[i] !== '\t') { i += 1; continue; }
    let j = i;
    while (j < s.length && (s[j] === ' ' || s[j] === '\t')) j += 1;
    if (s[j] === '\n') {
      parts.push(s.slice(plainStart, i));
      plainStart = j;
    }
    i = j;
  }
  parts.push(s.slice(plainStart));
  return parts.join('');
}

/** Collapse runs of 3+ newlines to exactly two (`/\n{3,}/g`). */
function collapseNewlines(s) {
  const parts = [];
  let i = 0;
  let plainStart = 0;
  while (i < s.length) {
    if (s[i] !== '\n') { i += 1; continue; }
    let j = i;
    while (j < s.length && s[j] === '\n') j += 1;
    if (j - i >= 3) {
      parts.push(s.slice(plainStart, i), '\n\n');
      plainStart = j;
    }
    i = j;
  }
  parts.push(s.slice(plainStart));
  return parts.join('');
}

/** `h1`..`h6` -> 1..6, anything else -> 0. */
export function headingLevel(tag) {
  if (tag.length !== 2 || tag[0] !== 'h') return 0;
  const cp = tag.codePointAt(1);
  return cp >= CODE_DIGIT_1 && cp <= CODE_DIGIT_6 ? cp - CODE_DIGIT_1 + 1 : 0;
}

/** Render inline content of a node (no block structure). */
function renderInline(node) {
  const out = [];
  for (const child of node.children || []) {
    if (child.tag === '#text') { out.push(child.text); continue; }
    const inner = renderInline(child);
    switch (child.tag) {
      case 'br': out.push('  \n'); break;
      case 'img': out.push(`![${attrOf(child, 'alt')}](${attrOf(child, 'src')})`); break;
      case 'strong': case 'b': out.push(`**${collapse(inner)}**`); break;
      case 'em': case 'i': out.push(`*${collapse(inner)}*`); break;
      case 'u': out.push(`_${collapse(inner)}_`); break;
      case 'del': case 's': out.push(`~~${collapse(inner)}~~`); break;
      case 'code': out.push(`\`${collapse(inner)}\``); break;
      case 'a': {
        const href = attrOf(child, 'href');
        out.push(href ? `[${collapse(inner)}](${href})` : collapse(inner));
        break;
      }
      case 'p': out.push(collapse(inner)); break;
      default: out.push(inner);
    }
  }
  return out.join('');
}

function tableRow(rowNode) {
  return (rowNode.children || [])
    .filter((x) => x.tag === 'td' || x.tag === 'th')
    .map((x) => escapePipes(collapse(renderInline(x))));
}

function renderTable(node) {
  const rows = [];
  const collect = (n) => {
    for (const c of n.children || []) {
      if (c.tag === 'tr') rows.push(tableRow(c));
      else if (c.tag !== '#text') collect(c);
    }
  };
  collect(node);
  if (!rows.length) return '';
  const width = Math.max(...rows.map((r) => r.length));
  const pad = (r) => {
    const cells = [...r];
    while (cells.length < width) { cells.push(''); }
    return cells;
  };
  const lines = ['| ' + pad(rows[0]).join(' | ') + ' |'];
  lines.push('| ' + Array.from({ length: width }, () => '---').join(' | ') + ' |');
  for (let i = 1; i < rows.length; i += 1) lines.push('| ' + pad(rows[i]).join(' | ') + ' |');
  return lines.join('\n') + '\n\n';
}

// ---------------------------------------------------------------------------
// block rendering
// ---------------------------------------------------------------------------
function flushPending(state) {
  if (!state.pending.length) return;
  const text = collapse(state.pending.map((c) => (c.tag === '#text' ? c.text : renderInline(c))).join(''));
  state.pending = [];
  if (text) state.out.push(text + '\n\n');
}

function pushBlock(state, text) {
  flushPending(state);
  state.out.push(text);
}

function pushHeading(state, child, level) {
  flushPending(state);
  state.out.push('#'.repeat(level) + ' ' + collapse(renderInline(child)) + '\n\n');
}

function pushParagraph(state, child) {
  flushPending(state);
  const text = collapse(renderInline(child));
  if (text) state.out.push(text + '\n\n');
}

function pushList(state, child, listDepth) {
  flushPending(state);
  let idx = 1;
  for (const li of child.children || []) {
    if (li.tag !== 'li') continue;
    const marker = child.tag === 'ol' ? `${idx++}. ` : '- ';
    state.out.push('  '.repeat(listDepth) + marker + collapse(renderInline(li)) + '\n');
    const nested = (li.children || []).filter((c) => c.tag === 'ul' || c.tag === 'ol');
    if (nested.length) state.out.push(renderBlocks({ children: nested }, listDepth + 1));
  }
  state.out.push('\n');
}

function pushQuote(state, child) {
  flushPending(state);
  const inner = renderBlocks(child).trim().split('\n').map((l) => '> ' + l).join('\n');
  state.out.push(inner + '\n\n');
}

function pushPre(state, child) {
  flushPending(state);
  state.out.push('```\n' + trimTrailingNewlines(renderInline(child)) + '\n```\n\n');
}

function pushImage(state, child) {
  pushBlock(state, `![${attrOf(child, 'alt')}](${attrOf(child, 'src')})\n\n`);
}

/** `head` is dropped wholesale; other containers recurse in place. */
function pushContainer(state, child, listDepth) {
  if (child.tag === 'head') { state.pending = []; return; }
  flushPending(state);
  state.out.push(renderBlocks(child, listDepth));
}

function noop() { /* dropped tag */ }

const BLOCK_HANDLERS = {
  p: pushParagraph,
  ul: pushList,
  ol: pushList,
  table: (state, child) => pushBlock(state, renderTable(child)),
  blockquote: pushQuote,
  pre: pushPre,
  hr: (state) => pushBlock(state, '---\n\n'),
  br: (state, child) => { state.pending.push(child); },
  img: pushImage,
  div: pushContainer,
  body: pushContainer,
  html: pushContainer,
  head: pushContainer,
  section: pushContainer,
  article: pushContainer,
  main: pushContainer,
  header: pushContainer,
  footer: pushContainer,
  style: noop,
  script: noop,
  title: noop,
  meta: noop,
  link: noop,
};

function renderUnknownBlock(state, child, listDepth) {
  if (INLINE_TAGS.has(child.tag)) { state.pending.push(child); return; }
  pushContainer(state, child, listDepth);
}

/** Render block content of a node. */
function renderBlocks(node, listDepth = 0) {
  const state = { out: [], pending: [] };
  for (const child of node.children || []) {
    if (child.tag === '#text') { state.pending.push(child); continue; }
    const level = headingLevel(child.tag);
    if (level) pushHeading(state, child, level);
    else {
      const handler = BLOCK_HANDLERS[child.tag];
      if (handler) handler(state, child, listDepth);
      else renderUnknownBlock(state, child, listDepth);
    }
  }
  flushPending(state);
  return state.out.join('');
}

/** Convert mammoth-style HTML to compact Markdown text. */
export function htmlToMarkdown(html) {
  if (!html) return '';
  const source = String(html);
  const headEnd = source.indexOf(HEAD_CLOSE);
  const body = headEnd === -1 ? source : source.slice(headEnd + HEAD_CLOSE.length);
  return collapseNewlines(trimBeforeNewline(renderBlocks(parseHtml(body)))).trim();
}

/** Convert a plain text string into minimal safe HTML paragraphs. */
export function plainTextToHtml(text) {
  const esc = String(text ?? '')
    .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
  const paras = esc.split(/\n{2,}/).map((p) => p.replaceAll('\n', '<br/>'));
  return '<html><body>' + paras.map((p) => `<p>${p}</p>`).join('\n') + '</body></html>';
}
