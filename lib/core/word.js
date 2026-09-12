// .docx read/write/template core: mammoth (read), @turbodocx/html-to-docx (write),
// docxtemplater + PizZip (template fill).
import mammoth from 'mammoth';
import htmlToDocx from '@turbodocx/html-to-docx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { OfficeError, isWithinBytes, CAPS, roughStats, assertZipBudget } from './util.js';
import { readTagAt, attrValue } from './markup.js';
import { decodeEntities } from './md.js';

function escapeText(s) {
  return String(s)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

/** Image containers that carry no meaning of their own: drop the tag, keep the text. */
function isImageWrapper(name) {
  return name === 'figure' || name === 'figcaption';
}

/** `alt` of an `<img>` start tag, entity-decoded then re-escaped; '' when absent. */
function imageAltText(startTag) {
  const alt = attrValue(startTag, 'alt');
  return alt ? escapeText(decodeEntities(alt)) : '';
}

/**
 * Remove `<img>` tags (and unwrap `<figure>` / `<figcaption>`) before handing
 * HTML to html-to-docx. Image embedding is out of scope for this plugin, and
 * the image code path is the only one that reaches the network-capable
 * `probe-image-size` / `needle` stack — with no `<img>` this plugin still makes
 * no outbound request at all. `alt` text is kept so no information is lost and
 * no markup can be smuggled in.
 * @param html raw HTML
 * @returns html without images
 */
export function stripImages(html) {
  const s = String(html);
  let out = '';
  let i = 0;
  while (i < s.length) {
    const lt = s.indexOf('<', i);
    if (lt === -1) return out + s.slice(i);
    out += s.slice(i, lt);
    const tag = readTagAt(s, lt);
    if (!tag) {
      out += '&lt;';
      i = lt + 1;
      continue;
    }
    const name = tag.name.toLowerCase();
    if (name === 'img') out += tag.closing ? '' : imageAltText(s.slice(lt, tag.end));
    else if (!isImageWrapper(name)) out += s.slice(lt, tag.end);
    i = tag.end;
  }
  return out;
}

/** Read a .docx buffer -> { content, html, meta }. */
export async function readDocx(buf, { html = false } = {}) {
  if (!isWithinBytes(buf, CAPS.MAX_WORD_INPUT_BYTES)) {
    throw new OfficeError('docx 文件超过 40 MB 上限', 'OFFICE_TOO_LARGE');
  }
  await assertZipBudget(buf, 'docx');
  const htmlResult = await mammoth.convertToHtml({ buffer: buf });
  const htmlText = htmlResult.value || '';
  const messages = (htmlResult.messages || []).map((m) => m.message);
  const textResult = await mammoth.extractRawText({ buffer: buf });
  const text = (textResult.value || '').replaceAll('\r', '');
  return {
    content: text,
    html: htmlText,
    meta: {
      words: roughStats(text).words,
      paragraphs: roughStats(text).paragraphs,
      messages: messages.slice(0, 5),
    },
  };
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
  const title = String(opts.title || '文档');
  const options = {
    title,
    font: '等线',
    fontSize: 21, // half-points → 10.5pt
    pageSize: opts.landscape ? 'A4 Landscape' : 'A4',
    margins: opts.marginsMm
      ? {
          top: opts.marginsMm,
          bottom: opts.marginsMm,
          left: opts.marginsMm,
          right: opts.marginsMm,
        }
      : undefined,
  };
  try {
    const buf = await htmlToDocx(stripImages(html), null, options);
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
  let zip;
  try {
    zip = new PizZip(buf);
  } catch (err) {
    throw new OfficeError('不是有效的 docx 文件(无法解压): ' + (err?.message || err), 'BAD_DOCX');
  }
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
