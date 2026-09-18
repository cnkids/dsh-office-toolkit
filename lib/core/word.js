// .docx read/write/template core: mammoth (read), docx-writer (write, OOXML
// generated locally), docxtemplater (template fill) — zips go through ooxml.js.
import mammoth from 'mammoth';
import Docxtemplater from 'docxtemplater';
import { OfficeError, isWithinBytes, CAPS, roughStats, assertZipBudget } from './util.js';
import { htmlToDocxBuffer } from './docx-writer.js';
import { extractDocxFormat } from './docx-format.js';
import { openOoxml, officePartOf, partOf } from './ooxml.js';
import { materializeNumbering } from './docx-numbering-read.js';
import { PLUGIN_VERSION } from './version.js';

/** Read a .docx buffer -> { content, html, meta }. */
export async function readDocx(buf, { html = false, formatting = false } = {}) {
  if (!isWithinBytes(buf, CAPS.MAX_WORD_INPUT_BYTES)) {
    throw new OfficeError('docx 文件超过 40 MB 上限', 'OFFICE_TOO_LARGE');
  }
  await assertZipBudget(buf, 'docx');
  // 非标准 zip(反斜杠条目名 / 大小写不同)在这里被修正,mammoth 才能找到 word/document.xml
  const { zip, buffer, repaired } = await openOoxml(buf, 'docx');
  const format = formatting ? readFormattingFromZip(zip) : undefined;
  // 自动编号(多级标题 1.1 / 公文 一、（一）)展开成文字,否则正文里只剩标题文字
  const numbered = materializeNumbering(zip);
  const parsed = await parseWithFallback(zip, numbered?.buffer ?? buffer);
  const htmlText = parsed.html;
  const messages = numbered
    ? [`已把 ${numbered.count} 段自动编号展开成文字（如「一、」「（一）」）`, ...parsed.messages].slice(0, 5)
    : parsed.messages;
  const text = parsed.text;
  return {
    content: text,
    html: htmlText,
    formatting: format,
    meta: {
      words: roughStats(text).words,
      paragraphs: roughStats(text).paragraphs,
      messages: messages.slice(0, 5),
      repairedParts: repaired.length ? repaired : undefined,
      reader: parsed.reader,
    },
  };
}

/**
 * mammoth 认不出这个文件时的两类提示:
 *   "Could not find main document part…" / "Could not find the body element…"
 * 两种情况都值得换一种读法再试,而不是直接失败。
 */
function isMissingPart(err) {
  const message = String(err?.message || '').toLowerCase();
  return message.includes('main document part')
    || message.includes('valid .docx')
    || message.includes('body element')
    || message.includes('is a docx');
}

/**
 * 解析 .docx 正文,三级兜底,保证「能开 zip 就读得出字」:
 *   1. mammoth 直接解析(修正条目名与引用后的字节)
 *   2. mammoth 认不出时用内置解析器从 word/document.xml 取正文
 *      (表格/样式会丢,但不会读不出来)
 */
async function parseWithFallback(zip, buffer) {
  const direct = await tryMammoth(buffer);
  if (direct) return { ...direct, reader: 'mammoth' };
  const builtin = builtinText(zip);
  if (!builtin) {
    throw new OfficeError(
      `docx 解析失败:${describeContainer(zip, [])}。` +
        '请把这条报错发回以便定位;应急办法是用 Word 打开后另存为 .docx 再试',
      'BAD_DOCX'
    );
  }
  return { ...builtin, reader: '内置解析器(mammoth 无法解析,表格与样式可能丢失)' };
}

async function tryMammoth(buffer) {
  try {
    const htmlResult = await mammoth.convertToHtml({ buffer });
    const textResult = await mammoth.extractRawText({ buffer });
    return {
      html: htmlResult.value || '',
      messages: (htmlResult.messages || []).map((m) => m.message).slice(0, 5),
      text: (textResult.value || '').replaceAll('\r', ''),
    };
  } catch (err) {
    if (isMissingPart(err)) return null;
    throw new OfficeError('生成 docx 正文失败: ' + (err?.message || err), 'BAD_DOCX');
  }
}

/**
 * 找到主文档部件,顺序与真正的 OOXML 读取器一致:
 *   1. 包关系 `_rels/.rels` 里 Type 以 /officeDocument 结尾的那条 Target
 *   2. 规范名 word/document.xml
 *   3. 以 document.xml 结尾的任何部件(容忍奇怪的目录名与大小写)
 * @returns { entry, path } 或 null
 */
export function resolveMainPart(zip) {
  const main = officePartOf(zip);
  return main?.kind === 'docx' ? { entry: main.entry, path: main.path } : null;
}

/** 内置解析器兜底:按段落取文本(表格与样式会丢,但不会读不出来)。 */
function builtinText(zip) {
  const main = resolveMainPart(zip);
  if (!main) return null;
  try {
    const doc = extractDocxFormat(main.entry.asText());
    const paragraphs = doc.paragraphs.map((p) => p.text).filter((text) => text.trim());
    if (!paragraphs.length) return null;
    const escaped = paragraphs.map((text) => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;'));
    return {
      html: escaped.map((text) => `<p>${text}</p>`).join('\n'),
      messages: [`mammoth 无法解析该文件,已用内置解析器读取 ${main.path}(表格与样式可能丢失)`],
      text: paragraphs.join('\n'),
    };
  } catch {
    return null;
  }
}

/** 读取失败时把容器结构一并给出,便于一轮定位问题。 */
function describeContainer(zip, repaired) {
  const entries = zip ? Object.keys(zip.files).filter((name) => !name.endsWith('/')) : [];
  const main = zip ? resolveMainPart(zip) : null;
  const parts = [
    `插件 v${PLUGIN_VERSION}`,
    `zip 条目 ${entries.length} 个`,
    `主文档部件 ${main ? main.path : '未找到'}`,
    `已修正条目 ${repaired?.length || 0} 处`,
  ];
  if (entries.length) parts.push(`条目示例: ${entries.slice(0, 6).join('、')}`);
  return parts.join('; ');
}

/** 字体 / 字号 / 行距 / 缩进 / 页面设置(用于行文规则比对)。 */
function readFormattingFromZip(zip) {
  const documentXml = partOf(zip, 'word/document.xml');
  if (!documentXml) {
    throw new OfficeError(
      'docx 缺少 word/document.xml —— 这通常不是标准 Word 文件(可能是 WPS 或第三方工具生成后重命名)。请用 Word 打开后另存为 .docx 再试',
      'BAD_DOCX'
    );
  }
  const stylesXml = partOf(zip, 'word/styles.xml');
  const themeXml = partOf(zip, 'word/theme/theme1.xml');
  return extractDocxFormat(
    documentXml.asText(),
    stylesXml ? stylesXml.asText() : undefined,
    themeXml ? themeXml.asText() : undefined
  );
}

/**
 * Create a .docx buffer from HTML (preferred), markdown or plain text.
 * @param source { html?|markdown?|text? } exactly one
 * @param opts { title?, marginsMm?, landscape?, pageBreakBefore? }
 */
export async function writeDocx(source, opts = {}) {
  let html = '';
  if (typeof source.html === 'string') html = source.html;
  else if (typeof source.markdown === 'string') {
    const { markdownToHtml } = await import('./md.js');
    html = markdownToHtml(source.markdown);
  } else if (typeof source.text === 'string') {
    const { plainTextToHtml } = await import('./md.js');
    html = plainTextToHtml(source.text);
  } else {
    throw new OfficeError('需要提供 html / markdown / text 三者之一作为文档内容');
  }
  if (!html.trim()) throw new OfficeError('文档内容为空');
  const options = {
    title: String(opts.title || '文档'),
    landscape: Boolean(opts.landscape),
    marginsMm: Number.isFinite(opts.marginsMm) ? opts.marginsMm : undefined,
    style: opts.style,
    header: opts.header,
    footer: opts.footer,
    toc: opts.toc,
    readImage: opts.readImage,
  };
  try {
    const buf = await htmlToDocxBuffer(html, options);
    return Buffer.isBuffer(buf) ? buf : Buffer.from(buf);
  } catch (err) {
    // 参数校验类错误已经是 OfficeError,原样抛出保留错误码(便于工具层/agent 判断)
    if (err instanceof OfficeError) throw err;
    throw new OfficeError('生成 docx 失败: ' + (err?.message || err), 'DOCX_BUILD_FAILED');
  }
}

/**
 * Fill {{placeholders}} in an existing .docx buffer with data values.
 * Values may be string/number/boolean; nested objects reachable via a.b.
 */
export async function fillDocxTemplate(buf, data, { start = '{{', end = '}}' } = {}) {
  if (!isWithinBytes(buf, CAPS.MAX_WORD_INPUT_BYTES)) {
    throw new OfficeError('模板文件超过 40 MB 上限', 'OFFICE_TOO_LARGE');
  }
  await assertZipBudget(buf, 'docx 模板');
  const { zip } = await openOoxml(buf, 'docx 模板');
  let doc;
  try {
    doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start, end },
    });
    doc.render(data || {});
  } catch (err) {
    const detail = err?.properties?.explanation || err?.message || String(err);
    throw new OfficeError('模板渲染失败: ' + detail, 'TEMPLATE_RENDER_FAILED');
  }
  const out = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  return Buffer.from(out);
}
