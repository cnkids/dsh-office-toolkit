// OOXML 元素级操作的小工具:拆直接子元素、按 schema 顺序重建、取元素内部 XML。
// docx-edit.js(段落/表格手术)与 docx-table.js 共用,避免同一套扫描逻辑写两遍。
// 所有扫描都是单向前进(见 markup.js 的说明),不用正则,不存在回溯。
import { readTagAt } from './markup.js';

/** XML 属性值转义。 */
export function attrText(value) {
  return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');
}

/** 元素的直接子元素(不深入嵌套:同名元素取最近的一层)。 */
export function directChildren(inner) {
  const out = [];
  let i = 0;
  while (i < inner.length) {
    const at = inner.indexOf('<', i);
    if (at === -1) break;
    const tag = readTagAt(inner, at);
    if (!tag || tag.closing) {
      i = tag ? tag.end : at + 1;
      continue;
    }
    if (tag.selfClosing) {
      out.push({ name: tag.name, start: at, end: tag.end });
      i = tag.end;
      continue;
    }
    const closeAt = inner.indexOf(`</${tag.name}>`, tag.end);
    const end = closeAt === -1 ? inner.length : closeAt + tag.name.length + 3;
    out.push({ name: tag.name, start: at, end });
    i = end;
  }
  return out;
}

/** 收集所有 <name> 元素(允许嵌套,如同名不嵌套),用于给超链接里的 run 也套样式。 */
export function findElements(xml, name) {
  const out = [];
  let i = 0;
  while (i < xml.length) {
    const at = xml.indexOf(`<${name}`, i);
    if (at === -1) break;
    const tag = readTagAt(xml, at);
    if (!tag || tag.name !== name || tag.closing) {
      i = at + 1;
      continue;
    }
    if (tag.selfClosing) {
      out.push({ start: at, openEnd: tag.end, end: tag.end });
      i = tag.end;
      continue;
    }
    const closeAt = xml.indexOf(`</${name}>`, tag.end);
    out.push({ start: at, openEnd: tag.end, end: closeAt === -1 ? xml.length : closeAt + name.length + 3 });
    i = tag.end;
  }
  return out;
}

/** `<w:pPr …>…</w:pPr>` → 内部 XML。 */
export function elementInner(raw, name) {
  const openEnd = raw.indexOf('>') + 1;
  return raw.slice(openEnd, raw.length - name.length - 3);
}

function orderKey(name, order) {
  const at = order.indexOf(name);
  return at === -1 ? order.length : at;
}

/**
 * 重建元素的内部 XML:已有的同名子元素被替换,新子元素按 schema 顺序插入,
 * 其余子元素原样保留(顺序按 order 归一化,未知标签留在末尾且保持相对顺序)。
 */
export function rebuildChildren(inner, additions, order) {
  const items = directChildren(inner).map((c) => ({ name: c.name, raw: inner.slice(c.start, c.end) }));
  for (const [name, raw] of Object.entries(additions)) {
    const at = items.findIndex((it) => it.name === name);
    if (at >= 0) items[at].raw = raw;
    else items.push({ name, raw });
  }
  items.sort((a, b) => orderKey(a.name, order) - orderKey(b.name, order));
  return items.map((it) => it.raw).join('');
}

/** 把 `<tag …>` 形式的元素拆成开标签与内部 XML;不是成对标签时返回 null。 */
export function openAndInner(raw, name) {
  const tag = readTagAt(raw, 0);
  if (!tag || tag.name !== name || tag.closing || tag.selfClosing) return null;
  const closeAt = raw.lastIndexOf(`</${name}>`);
  if (closeAt === -1) return null;
  return { openTag: raw.slice(0, tag.end), inner: raw.slice(tag.end, closeAt) };
}

/**
 * 更新元素内部的属性容器(如 w:p 的 w:pPr、w:tc 的 w:tcPr、w:tbl 的 w:tblPr):
 * 已有则按 schema 顺序合并,没有则新建并放在最前(这几类容器在 schema 里都必须最前)。
 */
export function upsertProps(elementXml, elementName, propsName, additions, order) {
  const parts = openAndInner(elementXml, elementName);
  if (!parts) return elementXml;
  const at = directChildren(parts.inner).find((c) => c.name === propsName);
  const existing = at ? parts.inner.slice(at.start, at.end) : '';
  const merged = rebuildChildren(existing ? elementInner(existing, propsName) : '', additions, order);
  const props = `<${propsName}>${merged}</${propsName}>`;
  const rest = at ? parts.inner.slice(0, at.start) + parts.inner.slice(at.end) : parts.inner;
  return `${parts.openTag}${props}${rest}</${elementName}>`;
}

/** 在元素内部增删子元素后重新组装。 */
export function rebuildElement(raw, name, additions, order) {
  const parts = openAndInner(raw, name);
  if (!parts) return raw;
  return `${parts.openTag}${rebuildChildren(parts.inner, additions, order)}</${name}>`;
}

// ---------------------------------------------------------------------------
// run 级属性套用 / 文本提取(段落与表格共用)
// ---------------------------------------------------------------------------
/** w:rPr 子元素的 schema 顺序。 */
export const RPR_ORDER = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath'];

/** 给一个 run 套上字符样式(rPr 必须是 w:r 的第一个子元素)。 */
function restyleRun(runXml, openEndRel, tags) {
  const openTag = runXml.slice(0, openEndRel);
  const inner = runXml.slice(openEndRel, runXml.length - '</w:r>'.length);
  const existing = directChildren(inner).find((c) => c.name === 'w:rPr');
  const raw = existing ? inner.slice(existing.start, existing.end) : '';
  const rebuilt = rebuildChildren(raw ? elementInner(raw, 'w:rPr') : '', tags, RPR_ORDER);
  const rest = existing ? inner.slice(0, existing.start) + inner.slice(existing.end) : inner;
  return `${openTag}<w:rPr>${rebuilt}</w:rPr>${rest}</w:r>`;
}

/** 给一段 XML 里所有 run 套上字符样式(从后往前改,前面的偏移保持有效)。 */
export function restyleRuns(xml, tags) {
  if (!Object.keys(tags).length) return xml;
  let out = xml;
  for (const run of [...findElements(xml, 'w:r')].reverse()) {
    if (run.openEnd === run.end) continue; // 自闭合 <w:r/> 没有内容可套
    const rebuilt = restyleRun(out.slice(run.start, run.end), run.openEnd - run.start, tags);
    out = out.slice(0, run.start) + rebuilt + out.slice(run.end);
  }
  return out;
}

/** 一段 XML 里所有 <w:t> 的纯文本(用于按内容算列宽)。 */
export function elementText(xml) {
  let out = '';
  let i = 0;
  while (i < xml.length) {
    const at = xml.indexOf('<w:t', i);
    if (at === -1) break;
    const tag = readTagAt(xml, at);
    if (tag?.name !== 'w:t') {
      i = at + 1;
      continue;
    }
    if (tag.selfClosing) {
      i = tag.end;
      continue;
    }
    const closeAt = xml.indexOf('</w:t>', tag.end);
    if (closeAt === -1) break;
    out += xml.slice(tag.end, closeAt);
    i = closeAt + 6;
  }
  return out;
}

/** 取根元素内部 XML:跳过 <?xml …?> 声明,按名字找根标签(styles/numbering 等部件用)。 */
export function rootInner(xml, name) {
  const at = String(xml).indexOf(`<${name}`);
  if (at === -1) return '';
  const tag = readTagAt(xml, at);
  if (!tag) return '';
  const close = String(xml).lastIndexOf(`</${name}>`);
  return close <= tag.end ? '' : String(xml).slice(tag.end, close);
}
