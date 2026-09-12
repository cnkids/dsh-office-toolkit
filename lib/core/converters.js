// Public API of the cross-platform legacy-Word layer.
// Pure-JS readers first (no external software), external converters preferred
// when available for fidelity; writing legacy targets needs an external tool.
import { writeFile } from 'node:fs/promises';
import { OfficeError, extOf } from './util.js';
import { jsReadLegacy, textToRtf } from './legacy-read.js';
import {
  converterStatus, describeStatus, externalReadBackends, externalConvertBackends, hasExternalConverter,
} from './legacy-external.js';

export { rtfToText, textToRtf, odtXmlToHtml, textToHtml, readOdt, readRtf, readDocWithWordExtractor } from './legacy-read.js';
export { converterStatus, describeStatus, TEXTUTIL_TARGETS, WORD_COM_FORMATS } from './legacy-external.js';

const JS_WRITE_TARGETS = new Set(['txt', 'html', 'rtf', 'docx']);

async function tryBackends(backends, errors) {
  for (const [label, run] of backends) {
    try {
      const content = await run();
      if (content && String(content).trim()) return { content: String(content), backend: label };
      errors.push(`${label}: 结果为空`);
    } catch (err) {
      errors.push(`${label}: ${err?.message || err}`);
    }
  }
  return null;
}

/**
 * Read a legacy Word-family file (.doc/.rtf/.odt) to text or HTML.
 * @returns {Promise<{content: string, backend: string}>}
 */
export async function wordRead(path, { asHtml = false } = {}) {
  const errors = [];
  const external = await externalReadBackends(path, { asHtml });
  const hit = await tryBackends([...external, ['纯 JS 解析', () => jsReadLegacy(path, { asHtml })]], errors);
  if (hit) return hit;
  const ext = extOf(path);
  throw new OfficeError(
    `无法读取 .${ext} 文件。已尝试: ${errors.join(' | ')}。` +
    '可安装 LibreOffice(https://www.libreoffice.org/) 以获得最佳兼容性。',
    'LEGACY_READ_FAILED'
  );
}

/** Pure-JS rebuild fallback for legacy conversion (text-level fidelity). */
async function jsRebuild(inputPath, outputPath) {
  const target = extOf(outputPath);
  const { content } = await wordRead(inputPath, { asHtml: false });
  if (target === 'txt') {
    await writeFile(outputPath, content);
    return '纯 JS(文本重建)';
  }
  const { content: html } = await wordRead(inputPath, { asHtml: true });
  if (target === 'html') {
    await writeFile(outputPath, html);
    return '纯 JS(HTML 重建)';
  }
  if (target === 'rtf') {
    await writeFile(outputPath, textToRtf(content));
    return '纯 JS(RTF 生成)';
  }
  const word = await import('./word.js');
  await writeFile(outputPath, await word.writeDocx({ html }, { title: '转换文档' }));
  return '纯 JS(HTML→docx 重建)';
}

/**
 * Convert a Word-family file to a target extension.
 * @returns {Promise<{backend: string}>}
 */
export async function wordConvert(inputPath, outputPath) {
  const errors = [];
  const backends = await externalConvertBackends(inputPath, outputPath);
  for (const [label, run] of backends) {
    try {
      await run();
      return { backend: label };
    } catch (err) {
      errors.push(`${label}: ${err?.message || err}`);
    }
  }
  if (JS_WRITE_TARGETS.has(extOf(outputPath))) {
    try {
      return { backend: await jsRebuild(inputPath, outputPath) };
    } catch (err) {
      errors.push(`纯 JS 回退: ${err?.message || err}`);
    }
  }
  const srcExt = extOf(inputPath);
  const dstExt = extOf(outputPath);
  const hint = process.platform === 'win32'
    ? '建议安装 LibreOffice 或 Microsoft Word 后重试。'
    : '建议安装 LibreOffice 后重试。';
  throw new OfficeError(
    `无法把 .${srcExt} 转换为 .${dstExt}。已尝试: ${errors.join(' | ')}。${hint}`,
    'LEGACY_CONVERT_FAILED'
  );
}

/** Whether this machine can produce the given legacy target extension. */
export async function hasWordConverter(targetExt) {
  return hasExternalConverter(targetExt);
}

/** Backend summary string for diagnostics / tool output. */
export async function backendSummary() {
  return describeStatus(await converterStatus());
}
