// numbering.xml 与多级编号:标题自动编号(1 / 1.1 / 1.1.1)、正文列表挂到当前标题层级之下。
//
// 关键点:OOXML 的多级编号靠一个 abstractNum 里的 9 个 <w:lvl> 表达,各级用 %1..%9 引用
// 祖先计数器;把某个 <w:lvl> 用 <w:pStyle> 关联到 HeadingN,该样式的段落就自动带编号。
import { OfficeError } from './util.js';
import { attrIn, readTagAt } from './markup.js';
import { directChildren, rootInner } from './docx-xml.js';
import { partOf } from './ooxml.js';

const NUMBERING_PART = 'word/numbering.xml';
const STYLES_PART = 'word/styles.xml';
const RELS_PART = 'word/_rels/document.xml.rels';
const CONTENT_TYPES = '[Content_Types].xml';
const NUMBERING_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const NUMBERING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAX_LEVEL = 9;

/** 第 index 级(0 起)的编号文字:一级 "%1."、二级 "%1.%2"、三级 "%1.%2.%3" … */
export function levelTextAt(index) {
  const parts = Array.from({ length: index + 1 }, (_, i) => `%${i + 1}`);
  const text = parts.join('.');
  return index === 0 ? `${text}.` : text;
}

/**
 * GB/T 9704 公文层次:一、→（一）→1.→（1），再深继续用阿拉伯数字。
 * 中文数字靠 `chineseCounting` 数字格式,全角括号直接写在 lvlText 里。
 */
function gongwenLevel(index) {
  if (index === 0) return { numFmt: 'chineseCounting', text: '%1、' };
  if (index === 1) return { numFmt: 'chineseCounting', text: '（%2）' };
  if (index === 2) return { numFmt: 'decimal', text: '%3.' };
  if (index === 3) return { numFmt: 'decimal', text: '（%4）' };
  return { numFmt: 'decimal', text: `%${index + 1}.` };
}

/**
 * 预置方案。每级返回 { numFmt, text }。
 * nsid 固定:同一方案重复调用时据此认出「我们自己那套」并原地替换,
 * 而不是每次追加一份新的 abstractNum(这是幂等的关键)。
 */
const PRESETS = {
  'multicol-1_1_1': {
    nsid: '0D5F0001',
    label: '1 / 1.1 / 1.1.1',
    levels: (index) => ({ numFmt: 'decimal', text: levelTextAt(index) }),
  },
  'gongwen-1_1_1_1': {
    nsid: '0D5F0002',
    label: '一、/（一）/1./（1）',
    levels: gongwenLevel,
  },
};

/** 可用的预置方案名。 */
export const NUMBERING_PRESETS = Object.keys(PRESETS);

export function requirePreset(value) {
  const name = value === undefined || value === null || value === '' ? NUMBERING_PRESETS[0] : String(value).trim();
  if (!NUMBERING_PRESETS.includes(name)) {
    const hints = NUMBERING_PRESETS.map((p) => `${p}（${PRESETS[p].label}）`).join(' / ');
    throw new OfficeError(`set_numbering 的 style 可为 ${hints}`, 'INVALID_ARGS');
  }
  return name;
}

/** 样式名 → 标题级别:Heading1 / 标题 1 / 1 都算;认不出返回 0。 */
export function headingIndexOf(value) {
  let rest = String(value ?? '').trim().toLowerCase().replaceAll(' ', '');
  if (!rest) return 0;
  if (rest.startsWith('heading')) rest = rest.slice(7);
  else if (rest.startsWith('标题')) rest = rest.slice(2);
  if (rest.length === 1 && rest >= '1' && rest <= '9') return Number(rest);
  return 0;
}

function styleLevelOf(id, name) {
  const fromId = headingIndexOf(id);
  return fromId || headingIndexOf(name);
}

function requireStartNumber(value, what) {
  const n = Number(value);
  if (!Number.isInteger(n) || n < 1 || n > 9999) throw new OfficeError(`${what} 需为 1–9999 的整数`, 'INVALID_ARGS');
  return n;
}

/**
 * 参与编号的标题链。
 * - `exclude`: 不参与编号的标题样式(Heading1–9 / 标题 N / 数字 id 都认)
 * - `startFrom`: `{样式名: 起始数字}`,该样式成为链上第 1 级,比它浅的标题样式自动不参与
 * @returns {{levels: Array<{styleId: string, start: number}>, headingToChainLevel: Map<number, number>}}
 */
function parseExclude(exclude) {
  const excluded = new Set();
  for (const name of exclude) {
    const index = headingIndexOf(name);
    if (!index) {
      throw new OfficeError(`set_numbering 的 exclude 里有认不出的标题样式「${name}」（可用 Heading1–Heading9）`, 'INVALID_ARGS');
    }
    excluded.add(index);
  }
  return excluded;
}

function parseStartFrom(startFrom) {
  if (startFrom === undefined || startFrom === null) return { fromIndex: 0, firstStart: 1 };
  const entries = Object.entries(startFrom);
  if (entries.length !== 1) {
    throw new OfficeError('set_numbering 的 startFrom 需形如 {"Heading2": 1}（只能给一个样式）', 'INVALID_ARGS');
  }
  const [name, rawStart] = entries[0];
  const fromIndex = headingIndexOf(name);
  if (!fromIndex) {
    throw new OfficeError(`set_numbering 的 startFrom 里有认不出的标题样式「${name}」（可用 Heading1–Heading9）`, 'INVALID_ARGS');
  }
  return { fromIndex, firstStart: requireStartNumber(rawStart, 'set_numbering 的 startFrom 起始数字') };
}

export function chainSpecs(styleIds, { exclude = [], startFrom, linkToHeading = true } = {}) {
  const excluded = parseExclude(exclude);
  const { fromIndex, firstStart } = parseStartFrom(startFrom);
  const levels = [];
  const headingToChainLevel = new Map();
  for (let index = 1; index <= MAX_LEVEL; index += 1) {
    if (excluded.has(index)) continue;
    if (fromIndex && index < fromIndex) continue;
    const chainLevel = levels.length + 1;
    const detected = styleIds.get(index);
    levels.push({
      styleId: linkToHeading ? (detected ?? `Heading${index}`) : '',
      start: chainLevel === 1 ? firstStart : 1,
    });
    headingToChainLevel.set(index, chainLevel);
  }
  if (!levels.length) {
    throw new OfficeError('set_numbering: exclude 把标题样式全排除了，没有可编号的级别', 'INVALID_ARGS');
  }
  return { levels, headingToChainLevel };
}

/**
 * styles.xml 里各级标题的 styleId(1 起)。
 * Word 中文版的 styleId 可能是 "1"/"标题 1",所以按 id 与 name 双重识别,优先 canonical 的 HeadingN。
 */
export function headingStyleIds(stylesXml) {
  const ids = new Map();
  if (!stylesXml) return ids;
  const inner = rootInner(stylesXml, 'w:styles');
  for (const child of directChildren(inner)) {
    if (child.name !== 'w:style') continue;
    const raw = inner.slice(child.start, child.end);
    const tag = readTagAt(raw, 0);
    if (!tag || attrIn(tag.attrs, 'w:type') !== 'paragraph') continue;
    const id = attrIn(tag.attrs, 'w:styleId');
    const styleInner = rootInner(raw, 'w:style');
    const nameTag = directChildren(styleInner).find((c) => c.name === 'w:name');
    const name = nameTag ? attrIn(styleInner.slice(nameTag.start, nameTag.end), 'w:val') : '';
    const level = styleLevelOf(id, name);
    if (!level) continue;
    const current = ids.get(level);
    if (!current || /^heading\d$/i.test(String(id))) ids.set(level, id);
  }
  return ids;
}

function maxNumericAttr(xml, tagName, attr) {
  let max = 0;
  let i = 0;
  while (i < xml.length) {
    const at = xml.indexOf(`<${tagName} `, i);
    if (at === -1) break;
    const tag = readTagAt(xml, at);
    if (!tag || tag.name !== tagName) {
      i = at + 1;
      continue;
    }
    const digits = String(attrIn(tag.attrs, attr) ?? '').replace(/^\D*/, '');
    const n = Number.parseInt(digits, 10);
    if (Number.isFinite(n) && n > max) max = n;
    i = tag.end;
  }
  return max;
}

/** 一个 <w:lvl>(schema 顺序:start, numFmt, pStyle, lvlText, lvlJc, pPr)。 */
function levelTag(index, spec, preset) {
  const level = preset.levels(index);
  const style = spec.styleId ? `<w:pStyle w:val="${spec.styleId}"/>` : '';
  return `<w:lvl w:ilvl="${index}">` +
    `<w:start w:val="${spec.start}"/>` +
    `<w:numFmt w:val="${level.numFmt}"/>` +
    style +
    `<w:lvlText w:val="${level.text}"/>` +
    '<w:lvlJc w:val="left"/>' +
    '<w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr>' +
    '</w:lvl>';
}

function abstractNumTag(abstractNumId, specs, preset) {
  const levels = specs.levels.map((spec, index) => levelTag(index, spec, preset)).join('');
  return `<w:abstractNum w:abstractNumId="${abstractNumId}">` +
    `<w:nsid w:val="${preset.nsid}"/><w:multiLevelType w:val="multilevel"/>${levels}</w:abstractNum>`;
}

/** 找出我们自己加的那份 abstractNum(靠 nsid 认)。 */
function findOurAbstractNum(xml, nsid) {
  let i = 0;
  while (i < xml.length) {
    const at = xml.indexOf('<w:abstractNum ', i);
    if (at === -1) return null;
    const tag = readTagAt(xml, at);
    if (tag?.name !== 'w:abstractNum') {
      i = at + 1;
      continue;
    }
    const closeAt = xml.indexOf('</w:abstractNum>', tag.end);
    const end = closeAt === -1 ? xml.length : closeAt + '</w:abstractNum>'.length;
    if (xml.slice(at, end).includes(`<w:nsid w:val="${nsid}"/>`)) {
      return { start: at, end, abstractNumId: Number.parseInt(attrIn(tag.attrs, 'w:abstractNumId') ?? '0', 10) };
    }
    i = end;
  }
  return null;
}

/** 引用了该 abstractNum 的 num 的 numId。 */
function numIdOfAbstract(xml, abstractNumId) {
  let i = 0;
  while (i < xml.length) {
    const at = xml.indexOf('<w:num ', i);
    if (at === -1) return 0;
    const tag = readTagAt(xml, at);
    if (tag?.name !== 'w:num') {
      i = at + 1;
      continue;
    }
    const closeAt = xml.indexOf('</w:num>', tag.end);
    const end = closeAt === -1 ? xml.length : closeAt + '</w:num>'.length;
    if (xml.slice(at, end).includes(`<w:abstractNumId w:val="${abstractNumId}"/>`)) {
      return Number.parseInt(attrIn(tag.attrs, 'w:numId') ?? '0', 10);
    }
    i = end;
  }
  return 0;
}

function numTag(numId, abstractNumId) {
  return `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractNumId}"/></w:num>`;
}

/** 文档根标签的命名空间集合复用到 numbering.xml,避免自己拼错 xmlns。 */
function numberingRoot(documentXml) {
  const at = String(documentXml).indexOf('<w:document');
  const tag = at === -1 ? null : readTagAt(documentXml, at);
  if (tag?.name === 'w:document') return documentXml.slice(at, tag.end).replace('<w:document', '<w:numbering');
  return '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
}

function addContentTypeOverride(xml, partName, contentType) {
  if (xml.includes(partName)) return xml;
  const override = `<Override PartName="${partName}" ContentType="${contentType}"/>`;
  const close = xml.lastIndexOf('</Types>');
  return close === -1 ? xml : `${xml.slice(0, close)}${override}${xml.slice(close)}`;
}

function addNumberingRelationship(relsXml) {
  if (!relsXml) return `<Relationships xmlns="${RELS_NS}"><Relationship Id="rIdNumbering" Type="${NUMBERING_REL_TYPE}" Target="numbering.xml"/></Relationships>`;
  if (relsXml.includes(NUMBERING_REL_TYPE)) return relsXml;
  const id = `rId${maxNumericAttr(relsXml, 'Relationship', 'Id') + 1}`;
  const rel = `<Relationship Id="${id}" Type="${NUMBERING_REL_TYPE}" Target="numbering.xml"/>`;
  const close = relsXml.lastIndexOf('</Relationships>');
  return close === -1 ? relsXml : `${relsXml.slice(0, close)}${rel}${relsXml.slice(close)}`;
}

/**
 * 确保压缩包里有一套多级编号,返回可用的 numId。
 * numbering.xml 不存在时会连同 [Content_Types].xml 与关系文件一起补出来。
 * @returns {{numId: number, abstractNumId: number, created: boolean, linkedLevels: number}}
 */
export function ensureNumbering(zip, { documentXml, preset, linkToHeading, exclude = [], startFrom }) {
  const presetName = requirePreset(preset);
  const spec = PRESETS[presetName];
  const link = Boolean(linkToHeading);
  const styleIds = headingStyleIds(partOf(zip, STYLES_PART)?.asText() ?? '');
  const specs = chainSpecs(styleIds, { exclude, startFrom, linkToHeading: link });
  const existing = partOf(zip, NUMBERING_PART);

  if (!existing) {
    const abstract = abstractNumTag(0, specs, spec);
    const xml = `${numberingRoot(documentXml)}${abstract}${numTag(1, 0)}</w:numbering>`;
    zip.file(NUMBERING_PART, xml);
    const ct = partOf(zip, CONTENT_TYPES);
    if (ct) zip.file(ct.name, addContentTypeOverride(ct.asText(), '/word/numbering.xml', NUMBERING_CONTENT_TYPE));
    const rels = partOf(zip, RELS_PART);
    if (rels) zip.file(rels.name, addNumberingRelationship(rels.asText()));
    else zip.file(RELS_PART, addNumberingRelationship(''));
    return { numId: 1, abstractNumId: 0, created: true, reused: false, specs };
  }

  const xml = existing.asText();
  const ours = findOurAbstractNum(xml, spec.nsid);
  if (ours) {
    // 幂等:同一份 abstractNum 原地换掉各级定义,numId 保持不变(段落里的引用不会失效)
    const rebuilt = abstractNumTag(ours.abstractNumId, specs, spec);
    const replaced = xml.slice(0, ours.start) + rebuilt + xml.slice(ours.end);
    const numId = numIdOfAbstract(replaced, ours.abstractNumId);
    if (numId) {
      zip.file(existing.name, replaced);
      return { numId, abstractNumId: ours.abstractNumId, created: false, reused: true, specs };
    }
    zip.file(existing.name, replaced.replace('</w:numbering>', `${numTag(maxNumericAttr(replaced, 'w:num', 'w:numId') + 1, ours.abstractNumId)}</w:numbering>`));
    return { numId: maxNumericAttr(replaced, 'w:num', 'w:numId') + 1, abstractNumId: ours.abstractNumId, created: false, reused: true, specs };
  }

  const abstractNumId = maxNumericAttr(xml, 'w:abstractNum', 'w:abstractNumId') + 1;
  const numId = maxNumericAttr(xml, 'w:num', 'w:numId') + 1;
  const abstract = abstractNumTag(abstractNumId, specs, spec);
  // abstractNum 必须排在所有 num 之前
  const firstNum = xml.indexOf('<w:num ');
  const withAbstract = firstNum === -1
    ? xml.replace('</w:numbering>', `${abstract}</w:numbering>`)
    : `${xml.slice(0, firstNum)}${abstract}${xml.slice(firstNum)}`;
  zip.file(existing.name, withAbstract.replace('</w:numbering>', `${numTag(numId, abstractNumId)}</w:numbering>`));
  return { numId, abstractNumId, created: false, reused: false, specs };
}
