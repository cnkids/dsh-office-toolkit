// Markdown -> HTML (marked) and HTML -> Markdown-ish text (tiny tree walker),
// plus plain text -> HTML. The HTML side must handle what mammoth emits:
// h1-h6, p, ul/ol/li, table/tr/td/th, strong/em, a, br, img, blockquote, pre/code.
import { marked } from 'marked';
import { OfficeError } from './util.js';

marked.setOptions({ gfm: true, breaks: false });

export function markdownToHtml(md) {
  if (typeof md !== 'string' || !md.trim()) throw new OfficeError('markdown 内容为空');
  return String(marked.parse(md, { async: false }));
}

const VOID_TAGS = new Set(['br', 'img', 'hr', 'input', 'meta', 'link']);
const INLINE_TAGS = new Set(['strong', 'b', 'em', 'i', 'u', 'code', 'a', 'span', 'sub', 'sup', 'del', 's', 'mark']);

function decodeEntities(s) {
  return String(s)
    .replaceAll('&nbsp;', ' ')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll(/&#39;|&apos;/g, "'")
    .replaceAll('&amp;', '&');
}

function escapePipes(s) {
  return String(s).replaceAll('|', String.raw`\|`);
}

// Comments / CDATA / doctype carry no text and may contain '>', so they are
// stripped before the tag tokenizer runs.
const MARKUP_NOISE_RE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<![^>]*>/g;
const TOKEN_RE = /<(\/?)([a-zA-Z][a-zA-Z0-9]*)((?:"[^"]*"|'[^']*'|[^>"'])*)>/g;

/** Parse an HTML fragment into a minimal node tree. */
function parseHtml(html) {
  const source = String(html).replaceAll(MARKUP_NOISE_RE, '');
  const root = { tag: 'root', attrs: '', children: [] };
  const stack = [root];
  let cursor = 0;
  for (const m of source.matchAll(TOKEN_RE)) {
    pushText(stack, source.slice(cursor, m.index));
    cursor = m.index + m[0].length;
    if (m[1]) closeNode(stack, m[2].toLowerCase());
    else openNode(stack, m[2].toLowerCase(), m[3] || '', m[0].endsWith('/>'));
  }
  pushText(stack, source.slice(cursor));
  return root;
}

function closeNode(stack, tag) {
  for (let i = stack.length - 1; i > 0; i--) {
    if (stack[i].tag === tag) { stack.length = i; return; }
  }
}

function openNode(stack, tag, attrs, selfClosing) {
  const node = { tag, attrs, children: [] };
  stack.at(-1).children.push(node);
  if (!VOID_TAGS.has(tag) && !selfClosing) stack.push(node);
}

function pushText(stack, text) {
  if (!text) return;
  stack.at(-1).children.push({ tag: '#text', text: decodeEntities(text) });
}

function attrOf(node, name) {
  const pattern = String.raw`${name}\s*=\s*"([^"]*)"|${name}\s*=\s*'([^']*)'`;
  const m = new RegExp(pattern, 'i').exec(String(node.attrs || ''));
  return m ? decodeEntities(m[1] ?? m[2] ?? '') : '';
}

function collapse(s) {
  return String(s).replaceAll(/[ \t\r\n]+/g, ' ').trim();
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
  for (let i = 1; i < rows.length; i++) lines.push('| ' + pad(rows[i]).join(' | ') + ' |');
  return lines.join('\n') + '\n\n';
}

function tableRow(rowNode) {
  return (rowNode.children || [])
    .filter((x) => x.tag === 'td' || x.tag === 'th')
    .map((x) => escapePipes(collapse(renderInline(x))));
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
  state.out.push('```\n' + renderInline(child).replace(/\n+$/, '') + '\n```\n\n');
}

function pushImage(state, child) {
  pushBlock(state, `![${attrOf(child, 'alt')}](${attrOf(child, 'src')})\n\n`);
}

function pushContainer(state, child, listDepth) {
  if (child.tag === 'head') { state.pending = []; return; }
  flushPending(state);
  state.out.push(renderBlocks(child, listDepth));
}

function noop() { /* dropped tag */ }

const HEADING_RE = /^h([1-6])$/;
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
    const heading = HEADING_RE.exec(child.tag);
    if (heading) pushHeading(state, child, Number(heading[1]));
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
  let body = String(html);
  const headEnd = body.indexOf('</head>');
  if (headEnd !== -1) body = body.slice(headEnd + 7);
  body = body.replaceAll(/<head[\s\S]*?<\/head>/gi, '');
  const tree = parseHtml(body);
  const md = renderBlocks(tree)
    .replaceAll(/[ \t]+\n/g, '\n')
    .replaceAll(/\n{3,}/g, '\n\n')
    .trim();
  return md;
}

/** Convert a plain text string into minimal safe HTML paragraphs. */
export function plainTextToHtml(text) {
  const esc = String(text ?? '')
    .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
  const paras = esc.split(/\n{2,}/).map((p) => p.replaceAll('\n', '<br/>'));
  return '<html><body>' + paras.map((p) => `<p>${p}</p>`).join('\n') + '</body></html>';
}
