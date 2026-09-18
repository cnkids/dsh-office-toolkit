// 读侧:把 Word 的自动编号展开成文字,让 office_read 看到真实的层级。
//
// 段落的编号通常不在段落自己身上:要么由 numbering.xml 的 <w:lvl><w:pStyle> 关联到标题
// 样式,要么写在 styles.xml 的段落样式里。只取正文会丢掉「一、/（一）/1./（1）」这些层级,
// 而它们正是行文规则比对要认的东西。这里按 Word 的规则重现计数器,把编号文字以文本 run
// 的形式插到段落最前面(项目符号除外 —— 那是视觉装饰,mammoth 已经渲染成列表)。
import { attrIn, readTagAt } from './markup.js';
import { directChildren, elementInner, findElements } from './docx-xml.js';
import { normalizeEntryName, officePartOf, partOf } from './ooxml.js';

const MAX_LEVEL = 9;
const CN_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
const CN_UNITS = ['', '十', '百', '千'];
const CN_FORMATS = ['chineseCounting', 'chineseCountingThousand', 'ideographDigital', 'japaneseCounting', 'taiwaneseCounting'];
const ROMAN_TABLE = [
  [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
  [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
];
const CARDINALS = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten'];
const ORDINALS = ['zeroth', 'first', 'second', 'third', 'fourth', 'fifth', 'sixth', 'seventh', 'eighth', 'ninth', 'tenth'];
/** 编号与正文之间的分隔:tab(默认)按空格处理,标点型编号读起来更顺。 */
const SEPARATOR = { nothing: '', space: ' ', tab: ' ' };

/** 中文数字(1–9999,Word 的 chineseCounting 就是这个写法)。 */
function cnNumber(value) {
  if (value <= 0 || value > 9999) return String(value);
  const digits = String(value).split('').map(Number).reverse();
  let out = '';
  let pendingZero = false;
  for (let index = digits.length - 1; index >= 0; index -= 1) {
    const digit = digits[index];
    if (digit === 0) {
      pendingZero = out !== '';
      continue;
    }
    out += (pendingZero ? '零' : '') + CN_DIGITS[digit] + CN_UNITS[index];
    pendingZero = false;
  }
  return out.replace(/^一十/, '十');
}

function roman(value) {
  if (value <= 0 || value > 3999) return String(value);
  let rest = value;
  let out = '';
  for (const [amount, sign] of ROMAN_TABLE) {
    while (rest >= amount) {
      out += sign;
      rest -= amount;
    }
  }
  return out;
}

/** A…Z、AA…(Word 的 lowerLetter / upperLetter 是 26 进制进一位)。 */
function letters(value) {
  if (value <= 0) return String(value);
  let rest = value;
  let out = '';
  while (rest > 0) {
    const remainder = (rest - 1) % 26;
    out = String.fromCodePoint(65 + remainder) + out;
    rest = Math.floor((rest - 1) / 26);
  }
  return out;
}

const FORMATTERS = new Map([
  ['decimal', String],
  ['decimalZero', (value) => String(value).padStart(2, '0')],
  ['lowerLetter', (value) => letters(value).toLowerCase()],
  ['upperLetter', (value) => letters(value)],
  ['lowerRoman', (value) => roman(value)],
  ['upperRoman', (value) => roman(value).toUpperCase()],
  ['cardinalText', (value) => CARDINALS[value] ?? String(value)],
  ['ordinal', (value) => ORDINALS[value] ?? `${value}th`],
]);
for (const name of CN_FORMATS) FORMATTERS.set(name, cnNumber);

/** 某级计数器的显示文字。认不出的数字格式按十进制处理。 */
export function formatCounter(numFmt, value) {
  const format = FORMATTERS.get(numFmt);
  return format ? format(value) : String(value);
}

function tagAttr(raw, name) {
  const tag = readTagAt(raw, 0);
  return tag ? attrIn(tag.attrs, name) : undefined;
}

/** 元素内部某个直接子元素的 w:val。 */
function childVal(inner, name) {
  const found = directChildren(inner).find((child) => child.name === name);
  if (!found) return undefined;
  return tagAttr(inner.slice(found.start, found.end), 'w:val');
}

function asNumber(value, fallback) {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

/** numbering.xml → abstractNumId → ilvl → 级别定义。 */
function parseAbstracts(numberingXml) {
  const abstracts = new Map();
  for (const element of findElements(numberingXml, 'w:abstractNum')) {
    const raw = numberingXml.slice(element.start, element.end);
    const inner = elementInner(raw, 'w:abstractNum');
    const levels = new Map();
    for (const child of directChildren(inner).filter((item) => item.name === 'w:lvl')) {
      const lvlRaw = inner.slice(child.start, child.end);
      const lvlInner = elementInner(lvlRaw, 'w:lvl');
      levels.set(asNumber(tagAttr(lvlRaw, 'w:ilvl'), 0), {
        start: asNumber(childVal(lvlInner, 'w:start'), 1),
        numFmt: childVal(lvlInner, 'w:numFmt') ?? 'decimal',
        lvlText: childVal(lvlInner, 'w:lvlText') ?? '',
        pStyle: childVal(lvlInner, 'w:pStyle'),
        suff: childVal(lvlInner, 'w:suff') ?? 'tab',
      });
    }
    abstracts.set(asNumber(tagAttr(raw, 'w:abstractNumId'), 1), levels);
  }
  return abstracts;
}

/** numbering.xml → numId → { abstractNumId, overrides }。 */
function parseNums(numberingXml) {
  const nums = new Map();
  for (const element of findElements(numberingXml, 'w:num')) {
    const raw = numberingXml.slice(element.start, element.end);
    const inner = elementInner(raw, 'w:num');
    const overrides = new Map();
    for (const child of directChildren(inner).filter((item) => item.name === 'w:lvlOverride')) {
      const ovRaw = inner.slice(child.start, child.end);
      const start = childVal(elementInner(ovRaw, 'w:lvlOverride'), 'w:startOverride');
      if (start !== undefined) overrides.set(asNumber(tagAttr(ovRaw, 'w:ilvl'), 0), asNumber(start, 1));
    }
    nums.set(asNumber(tagAttr(raw, 'w:numId'), 0), {
      abstractNumId: asNumber(childVal(inner, 'w:abstractNumId'), 0),
      overrides,
    });
  }
  return nums;
}

/** 段落样式里的 <w:numPr>(样式可以直接带编号)与 <w:basedOn> 继承链。 */
function parseStyleNumbering(stylesXml) {
  const styles = new Map();
  if (!stylesXml) return styles;
  for (const element of findElements(stylesXml, 'w:style')) {
    const raw = stylesXml.slice(element.start, element.end);
    const id = tagAttr(raw, 'w:styleId');
    if (!id) continue;
    const inner = elementInner(raw, 'w:style');
    const entry = { basedOn: childVal(inner, 'w:basedOn'), numId: 0, ilvl: 0 };
    const pPr = directChildren(inner).find((child) => child.name === 'w:pPr');
    if (pPr) {
      const pPrInner = elementInner(inner.slice(pPr.start, pPr.end), 'w:pPr');
      const numPr = directChildren(pPrInner).find((child) => child.name === 'w:numPr');
      if (numPr) {
        const numPrInner = elementInner(pPrInner.slice(numPr.start, numPr.end), 'w:numPr');
        entry.numId = asNumber(childVal(numPrInner, 'w:numId'), 0);
        entry.ilvl = asNumber(childVal(numPrInner, 'w:ilvl'), 0);
      }
    }
    styles.set(id, entry);
  }
  return styles;
}

/** 按 numbering.xml 的 <w:lvl><w:pStyle> 反查「哪个样式绑在哪一级」。 */
function styleBindings(abstracts, nums) {
  const bindings = new Map();
  for (const { abstractNumId } of nums.values()) {
    for (const [ilvl, level] of abstracts.get(abstractNumId) ?? []) {
      if (level.pStyle && !bindings.has(level.pStyle)) bindings.set(level.pStyle, { numId: 0, ilvl });
    }
  }
  return bindings;
}

/** numId 补全:样式联动只给了 ilvl,得挑一个引用该 abstractNum 的 numId。 */
function bindingsWithNumIds(bindings, abstracts, nums) {
  const resolved = new Map();
  for (const [styleId, binding] of bindings) {
    if (!binding.numId) {
      const match = [...nums.entries()].find(([, num]) => (abstracts.get(num.abstractNumId) ?? new Map()).get(binding.ilvl)?.pStyle === styleId);
      if (!match) continue;
      binding.numId = match[0];
    }
    resolved.set(styleId, binding);
  }
  return resolved;
}

/** 段落没写 numPr 时,顺着样式链找编号(样式自己的 numPr 优先于样式联动)。 */
function styleNumberingOf(styleId, styles, bindings) {
  let current = styleId;
  const seen = new Set();
  while (current && !seen.has(current)) {
    seen.add(current);
    const style = styles.get(current);
    if (style?.numId > 0) return { numId: style.numId, ilvl: style.ilvl };
    const bound = bindings.get(current);
    if (bound?.numId > 0) return bound;
    current = style?.basedOn;
  }
  return null;
}

/** 段落 pPr 里的 <w:numPr>(numId=0 表示取消编号)。 */
function paragraphNumbering(pPrInner) {
  const numPr = directChildren(pPrInner).find((child) => child.name === 'w:numPr');
  if (!numPr) return null;
  const numPrInner = elementInner(pPrInner.slice(numPr.start, numPr.end), 'w:numPr');
  const numId = asNumber(childVal(numPrInner, 'w:numId'), 0);
  return numId > 0 ? { numId, ilvl: asNumber(childVal(numPrInner, 'w:ilvl'), 0) } : null;
}

export function parseNumberingContext(numberingXml, stylesXml) {
  const abstracts = parseAbstracts(numberingXml);
  const nums = parseNums(numberingXml);
  return {
    abstracts,
    nums,
    styles: parseStyleNumbering(stylesXml),
    bindings: bindingsWithNumIds(styleBindings(abstracts, nums), abstracts, nums),
  };
}

/** 走一个段落,推进计数器并算出它的编号文字。 */
function numberTextOf(context, numbering) {
  const num = context.nums.get(numbering.numId);
  const abstract = num && context.abstracts.get(num.abstractNumId);
  const level = abstract?.get(numbering.ilvl);
  if (!level || level.numFmt === 'bullet' || level.numFmt === 'none' || !level.lvlText) return '';
  const startAt = (index) => num.overrides.get(index) ?? abstract.get(index)?.start ?? 1;
  let state = context.counters.get(numbering.numId);
  if (!state) {
    state = { values: [], seen: [] };
    context.counters.set(numbering.numId, state);
  }
  state.values[numbering.ilvl] = state.seen[numbering.ilvl] ? state.values[numbering.ilvl] + 1 : startAt(numbering.ilvl);
  state.seen[numbering.ilvl] = true;
  for (let index = numbering.ilvl + 1; index < MAX_LEVEL; index += 1) state.seen[index] = false;
  // %N 用第 N 级自己的数字格式(如一级「一、」、三级「1.」),祖先还没出现过就用起始值。
  const text = level.lvlText.replaceAll(/%(\d)/g, (whole, digit) => {
    const index = Number(digit) - 1;
    if (index >= MAX_LEVEL) return whole;
    return formatCounter(abstract.get(index)?.numFmt ?? level.numFmt, state.seen[index] ? state.values[index] : startAt(index));
  });
  return text + (SEPARATOR[level.suff] ?? ' ');
}

function escapeXml(value) {
  return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

/** 段落里该往哪插:必须在 w:pPr 之后(schema 要求属性在前)。 */
function insertionPoint(documentXml, paragraph) {
  const innerStart = paragraph.openEnd;
  const pPr = directChildren(documentXml.slice(innerStart, paragraph.end - '</w:p>'.length))
    .find((child) => child.name === 'w:pPr');
  return pPr ? innerStart + pPr.end : innerStart;
}

function insertNumbers(documentXml, context) {
  context.counters = new Map();
  const pieces = [];
  let cursor = 0;
  let count = 0;
  for (const paragraph of findElements(documentXml, 'w:p')) {
    // 自闭合的 <w:p/> 没有内容可插
    if (paragraph.end - paragraph.openEnd < '</w:p>'.length) continue;
    const inner = elementInner(documentXml.slice(paragraph.start, paragraph.end), 'w:p');
    const pPr = directChildren(inner).find((child) => child.name === 'w:pPr');
    const pPrInner = pPr ? elementInner(inner.slice(pPr.start, pPr.end), 'w:pPr') : '';
    const numbering = paragraphNumbering(pPrInner)
      ?? styleNumberingOf(childVal(pPrInner, 'w:pStyle') ?? '', context.styles, context.bindings);
    if (!numbering) continue;
    const text = numberTextOf(context, numbering);
    if (!text) continue;
    const at = insertionPoint(documentXml, paragraph);
    pieces.push(documentXml.slice(cursor, at), `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`);
    cursor = at;
    count += 1;
  }
  if (!count) return null;
  pieces.push(documentXml.slice(cursor));
  return { documentXml: pieces.join(''), count };
}

function entryNameOf(zip, path) {
  if (zip.files[path]) return path;
  const wanted = normalizeEntryName(path).toLowerCase();
  return Object.keys(zip.files).find((name) => normalizeEntryName(name).toLowerCase() === wanted) ?? null;
}

/**
 * 把主文档里的自动编号展开成文字,写回内存中的 zip 并返回新字节。
 * 没有编号(或只有项目符号)时返回 null,调用方照旧用原字节。
 * @returns { buffer, count } 或 null
 */
export function materializeNumbering(zip) {
  const main = officePartOf(zip);
  if (main?.kind !== 'docx') return null;
  const numberingEntry = partOf(zip, 'word/numbering.xml');
  if (!numberingEntry) return null;
  const name = entryNameOf(zip, main.path);
  if (!name) return null;
  const numbering = parseNumberingContext(numberingEntry.asText(), partOf(zip, 'word/styles.xml')?.asText());
  const applied = insertNumbers(main.entry.asText(), numbering);
  if (!applied) return null;
  zip.file(name, applied.documentXml);
  return { buffer: zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }), count: applied.count };
}
