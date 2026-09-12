// Linear, backtracking-free scanners for HTML/XML-ish markup.
//
// Regex-based tag matching (e.g. /<sheet\b[^>]*>/g) has unbounded backtracking:
// on crafted input a failing scan re-walks the same characters, so the total
// work grows super-linearly (SonarQube S5852). Every scanner here is a single
// forward pass, so the cost is O(n) in the input length and cannot be used to
// stall the host.

import { isAsciiDigit, isAsciiLetter } from './util.js';

const CODE_COLON = 58;
const CODE_HYPHEN = 45;
const CODE_UNDERSCORE = 95;
const CODE_DOT = 46;

/** XML name characters: letters, digits, `:`, `-`, `_`, `.`. */
export function isNameChar(cp) {
  return isAsciiLetter(cp) || isAsciiDigit(cp)
    || cp === CODE_COLON || cp === CODE_HYPHEN || cp === CODE_UNDERSCORE || cp === CODE_DOT;
}

export function isSpaceChar(ch) {
  return ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r';
}

/**
 * Read one `<...>` markup starting at `at` (which must point at `<`).
 * Quoted attribute values may contain `>`, which is why the scan is manual.
 * @param text source text
 * @param at index of `<`
 * @returns parsed tag, or null when no complete tag starts there
 */
export function readTagAt(text, at) {
  const s = String(text);
  let i = at + 1;
  const closing = s[i] === '/';
  if (closing) i += 1;
  const nameStart = i;
  while (i < s.length && isNameChar(s.codePointAt(i))) i += 1;
  if (i === nameStart || !isAsciiLetter(s.codePointAt(nameStart))) return null;
  const name = s.slice(nameStart, i);
  const attrsStart = i;
  let quote = '';
  while (i < s.length) {
    const ch = s[i];
    if (quote) {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === '>') {
      break;
    }
    i += 1;
  }
  if (i >= s.length) return null;
  const raw = s.slice(attrsStart, i);
  const selfClosing = raw.endsWith('/');
  return { name, closing, selfClosing, attrs: selfClosing ? raw.slice(0, -1) : raw, end: i + 1 };
}

/** Index of the first non-closing `<tagName ...>`, or -1. */
export function indexOfTag(text, tagName, ignoreCase = false) {
  const s = String(text);
  const want = ignoreCase ? String(tagName).toLowerCase() : String(tagName);
  let i = 0;
  while (i < s.length) {
    const lt = s.indexOf('<', i);
    if (lt === -1) return -1;
    const tag = readTagAt(s, lt);
    if (tag) {
      if (!tag.closing && matchesName(tag.name, want, ignoreCase)) return lt;
      i = tag.end;
    } else {
      i = lt + 1;
    }
  }
  return -1;
}

function matchesName(name, want, ignoreCase) {
  return ignoreCase ? name.toLowerCase() === want : name === want;
}

/** Raw text of every non-closing `<tagName ...>` occurrence, in order. */
export function tagTexts(text, tagName, ignoreCase = false) {
  const s = String(text);
  const want = ignoreCase ? String(tagName).toLowerCase() : String(tagName);
  const out = [];
  let i = 0;
  while (i < s.length) {
    const lt = s.indexOf('<', i);
    if (lt === -1) break;
    const tag = readTagAt(s, lt);
    if (tag) {
      if (!tag.closing && matchesName(tag.name, want, ignoreCase)) out.push(s.slice(lt, tag.end));
      i = tag.end;
    } else {
      i = lt + 1;
    }
  }
  return out;
}

/** Raw text of the first non-closing `<tagName ...>`, or undefined. */
export function firstTag(text, tagName, ignoreCase = false) {
  const at = indexOfTag(text, tagName, ignoreCase);
  if (at === -1) return undefined;
  const tag = readTagAt(text, at);
  return tag ? String(text).slice(at, tag.end) : undefined;
}

function isAttrNameEnd(ch) {
  return ch === undefined || ch === '=' || isSpaceChar(ch);
}

function readAttrValue(s, from) {
  let i = from;
  while (i < s.length && isSpaceChar(s[i])) i += 1;
  if (s[i] !== '=') return undefined;
  i += 1;
  while (i < s.length && isSpaceChar(s[i])) i += 1;
  const quote = s[i];
  if (quote !== '"' && quote !== "'") return undefined;
  const end = s.indexOf(quote, i + 1);
  return end === -1 ? undefined : s.slice(i + 1, end);
}

/**
 * Value of attribute `name` inside a raw attribute list (case-insensitive,
 * tolerates whitespace around `=`). Returns undefined when absent.
 */
export function attrIn(attrs, name) {
  const raw = ` ${attrs ?? ''}`;
  const lower = raw.toLowerCase();
  const needle = String(name).toLowerCase();
  let from = 1;
  while (from < raw.length) {
    const at = lower.indexOf(needle, from);
    if (at === -1) return undefined;
    if (isSpaceChar(raw[at - 1]) && isAttrNameEnd(raw[at + needle.length])) {
      const value = readAttrValue(raw, at + needle.length);
      if (value !== undefined) return value;
    }
    from = at + 1;
  }
  return undefined;
}

/** Value of attribute `name` on a raw start tag, or undefined. */
export function attrValue(rawTag, name) {
  const tag = readTagAt(String(rawTag), 0);
  return tag ? attrIn(tag.attrs, name) : undefined;
}

/**
 * Insert `inner` right before the final `</tagName>`, but only when that close
 * tag (plus trailing whitespace) really ends the text.
 */
export function appendBeforeClose(xml, tagName, inner) {
  const s = String(xml);
  const close = `</${tagName}>`;
  const at = s.lastIndexOf(close);
  if (at === -1) return s;
  const tail = s.slice(at + close.length);
  if (tail.trim() !== '') return s;
  return `${s.slice(0, at)}${inner}${close}${tail}`;
}
