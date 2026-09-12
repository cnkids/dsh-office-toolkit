// .docx read/write/template core: mammoth (read), docx-writer (write, OOXML
// generated locally), docxtemplater (template fill) — zips go through ooxml.js.
import mammoth from 'mammoth';
import Docxtemplater from 'docxtemplater';
import { OfficeError, isWithinBytes, CAPS, roughStats, assertZipBudget } from './util.js';
import { htmlToDocxBuffer } from './docx-writer.js';
import { extractDocxFormat } from './docx-format.js';
import { openOoxml, partOf } from './ooxml.js';

/** Read a .docx buffer -> { content, html, meta }. */
export async function readDocx(buf, { html = false, formatting = false } = {}) {
  if (!isWithinBytes(buf, CAPS.MAX_WORD_INPUT_BYTES)) {
    throw new OfficeError('docx 文件超过 40 MB 上限', 'OFFICE_TOO_LARGE');
  }
  await assertZipBudget(buf, 'docx');
  // 非标准 zip(条目名用反斜杠)在这里被修正,mammoth 才能找到 word/document.xml
  const { zip, buffer, repaired } = await openOoxml(buf, 'docx');
  const format = formatting ? readFormattingFromZip(zip) : undefined;
  const htmlResult = await mammoth.convertToHtml({ buffer });
  const htmlText = htmlResult.value || '';
  const messages = (htmlResult.messages || []).map((m) => m.message);
  const textResult = await mammoth.extractRawText({ buffer });
  const text = (textResult.value || '').replaceAll('\r', '');
  return {
    content: text,
    html: htmlText,
    formatting: format,
    meta: {
      words: roughStats(text).words,
      paragraphs: roughStats(text).paragraphs,
      messages: messages.slice(0, 5),
      repairedParts: repaired.length ? repaired : undefined,
    },
  };
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
  };
  try {
    const buf = await htmlToDocxBuffer(html, options);
    return Buffer.isBuffer(buf) ? buf : Buffer.from(buf);
  } catch (err) {
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
