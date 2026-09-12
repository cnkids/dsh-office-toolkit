// Office operation layer: pure operations on absolute paths + buffers.
// IO is injected ({ readBuf, writeBuf, stat }) so the DSH host adapter can
// route through ctx.fs / sandbox policy; the CLI selftest uses node:fs.
import { OfficeError, extOf, fmtBytes, readPath, writePath, truncate, CAPS, ZIP_OFFICE_EXTS, assertOfficeBinary } from './util.js';
import { tmpdir } from 'node:os';
import { randomBytes } from 'node:crypto';
import * as legacy from './legacy.js';
import { lazyModule } from './deps.js';
import { formatReport } from './docx-format.js';

// 重依赖懒加载:依赖装不全时插件仍能加载,且只有真正用到该模块的工具会失败,
// 报错会明确指出缺哪个包(见 deps.js)
const word = lazyModule('./word.js', 'Word 读写', ['readDocx', 'writeDocx', 'fillDocxTemplate']);
const excel = lazyModule('./excel.js', 'Excel 读写', ['readWorkbook', 'buildWorkbook', 'editWorkbook', 'workbookSummary']);
const converters = lazyModule('./converters.js', '旧 Word 格式转换', ['wordRead', 'wordConvert']);
const md = lazyModule('./md.js', 'Markdown 转换', ['htmlToMarkdown', 'markdownToHtml', 'plainTextToHtml']);

export const OFFICE_EXTENSIONS = {
  word: ['docx', 'doc', 'odt', 'rtf'],
  table: ['xlsx', 'xls', 'xlsb', 'ods', 'csv', 'tsv', 'txt'],
};

function defaultIo() {
  return {
    async readBuf(p) { return readPath(p); },
    async writeBuf(p, b, options) { return writePath(p, b, options); },
    async remove(p) {
      const { unlink } = await import('node:fs/promises');
      try { await unlink(p); } catch { /* ignore */ }
    },
    tmpFile(ext) {
      return joinTmp(tmpdir(), `dsh-office-${randomBytes(6).toString('hex')}${ext}`);
    },
    async stat(p) {
      const { stat } = await import('node:fs/promises');
      const s = await stat(p);
      return { size: s.size, type: s.isDirectory() ? 'dir' : 'file' };
    },
  };
}

function joinTmp(dir, name) {
  return dir.endsWith('/') ? dir + name : dir + '/' + name;
}

function normalizeIo(io) {
  const def = defaultIo();
  return {
    readBuf: io?.readBuf ? io.readBuf.bind(io) : def.readBuf,
    writeBuf: io?.writeBuf ? io.writeBuf.bind(io) : def.writeBuf,
    remove: io?.remove ? io.remove.bind(io) : def.remove,
    tmpFile: io?.tmpFile ? io.tmpFile.bind(io) : def.tmpFile,
    stat: io?.stat ? io.stat.bind(io) : def.stat,
  };
}

async function ensureRegular(io, path, caps) {
  let st;
  try {
    st = await io.stat(path);
  } catch {
    throw new OfficeError(`文件不存在: ${path}`, 'NOT_FOUND');
  }
  if (!st || st.type === 'dir') throw new OfficeError(`不是文件: ${path}`, 'NOT_A_FILE');
  const size = st.size || 0;
  if (caps && size > caps) throw new OfficeError(`文件过大(${fmtBytes(size)}), 上限 ${fmtBytes(caps)}`, 'OFFICE_TOO_LARGE');
  return st;
}

// ---------------------------------------------------------------------------
// office_read
// ---------------------------------------------------------------------------
export async function opRead(path, rawOpts, ioIn) {
  const opts = rawOpts || {};
  const io = normalizeIo(ioIn);
  const ext = extOf(path);
  if (ext === 'docx') return readDocxFile(io, path, opts);
  if (['doc', 'rtf', 'odt'].includes(ext)) return readLegacyWordFile(io, path, ext, opts);
  if (ext === 'xlsx') return readXlsxFile(io, path, opts);
  if (['xls', 'xlsb', 'ods', 'csv', 'tsv'].includes(ext)) return readLegacyTableFile(io, path, ext, opts);
  throw new OfficeError(`office_read 不支持的扩展名 .${ext}（支持: docx/doc/odt/rtf/xlsx/xls/xlsb/ods/csv/tsv）`, 'UNSUPPORTED_FORMAT');
}

async function readDocxFile(io, path, opts) {
  const needHtml = opts.format === 'html';
  const wantFormat = Boolean(opts.withFormatting);
  const buf = await io.readBuf(path);
  const r = await word.readDocx(buf, { html: needHtml, formatting: wantFormat });
  const content = needHtml ? r.html : ((await md.htmlToMarkdown(r.html)) || r.content);
  const capped = truncate(content, opts.maxChars ?? CAPS.DEFAULT_MAX_CHARS);
  const body = `已读取 Word 文档 ${path}\n- 大小: ${fmtBytes(buf.length)}\n- 词数约: ${r.meta.words}, 段落约: ${r.meta.paragraphs}\n\n` + capped;
  return {
    content: wantFormat && r.formatting ? `${body}\n\n## 格式报告\n\n${formatReport(r.formatting)}` : body,
    meta: {
      kind: 'word',
      format: 'docx',
      bytes: buf.length,
      words: r.meta.words,
      paragraphs: r.meta.paragraphs,
      formatting: wantFormat ? summarizeFormatting(r.formatting) : undefined,
    },
    html: needHtml ? r.html : undefined,
  };
}

/** 结构化格式信息,供比对工具直接消费(段落过多时截断)。 */
function summarizeFormatting(doc) {
  const paragraphs = doc.paragraphs.slice(0, CAPS.DEFAULT_MAX_ROWS);
  return {
    page: doc.page,
    defaults: doc.defaults,
    truncated: doc.paragraphs.length > paragraphs.length,
    paragraphs,
  };
}

async function readLegacyWordFile(io, path, ext, opts) {
  const needHtml = opts.format === 'html';
  const read = await converters.wordRead(path, { asHtml: needHtml });
  const content = needHtml ? read.content : ((await md.htmlToMarkdown(read.content)) || read.content);
  const capped = truncate(content, opts.maxChars ?? CAPS.DEFAULT_MAX_CHARS);
  const words = (capped.match(/[\u4e00-\u9fff]|[A-Za-z0-9]+/g) || []).length;
  const formatNote = opts.withFormatting ? '\n> 注: 字体/行距等格式提取目前仅支持 .docx\n' : '';
  return {
    content: `已读取 ${ext.toUpperCase()} 文档 ${path}\n> 解析方式: ${read.backend}，复杂版式可能丢失${formatNote}\n` + capped,
    meta: { kind: 'word', format: ext, via: read.backend, words },
  };
}

async function readXlsxFile(io, path, opts) {
  await ensureRegular(io, path, CAPS.MAX_EXCEL_INPUT_BYTES);
  const buf = await io.readBuf(path);
  const r = await excel.readWorkbook(buf, opts);
  return { content: `已读取 Excel 工作簿 ${path}（${r.meta.sheetNames.length} 个工作表）\n\n` + r.content, meta: r.meta };
}

async function readLegacyTableFile(io, path, ext, opts) {
  await ensureRegular(io, path, CAPS.MAX_EXCEL_INPUT_BYTES);
  const buf = await io.readBuf(path);
  const r = await legacy.readLegacySpreadsheet(buf, ext, opts);
  return { content: `已读取 ${ext.toUpperCase()} 表格 ${path}（${r.meta.sheets.length} 个工作表）\n\n` + r.content, meta: r.meta };
}

// ---------------------------------------------------------------------------
// office_write_docx
// ---------------------------------------------------------------------------
export async function opWriteDocx(path, params, ioIn) {
  const io = normalizeIo(ioIn);
  const ext = extOf(path);
  if (!['docx', 'doc', 'rtf', 'odt'].includes(ext)) {
    throw new OfficeError(`office_write_docx 输出扩展名应为 .docx（或 .doc/.rtf/.odt），实际 .${ext}`, 'UNSUPPORTED_FORMAT');
  }
  const source = {
    html: typeof params.html === 'string' ? params.html : undefined,
    markdown: typeof params.markdown === 'string' ? params.markdown : undefined,
    text: typeof params.text === 'string' ? params.text : undefined,
  };
  if (ext !== 'docx') return writeLegacyWord(io, path, ext, source, params);
  const buf = await word.writeDocx(source, {
    title: params.title || '文档',
    landscape: Boolean(params.landscape),
    marginsMm: params.marginsMm,
  });
  await writeOfficeBinary(io, path, buf, 'docx');
  const from = contentSourceLabel(source);
  return { content: `已生成 Word 文档: ${path}\n- 大小: ${fmtBytes(buf.length)}\n- 内容来源: ${from}`, meta: { format: 'docx', bytes: buf.length } };
}

/** Which input channel produced the document (for the result summary). */
function contentSourceLabel(source) {
  if (source.html) return 'html';
  if (source.markdown) return 'markdown';
  return 'text';
}

async function writeLegacyWord(io, path, ext, source, params) {
  const { markdownToHtml, plainTextToHtml } = await import('./md.js');
  let html = '';
  if (source.html) html = source.html;
  else if (source.markdown) html = markdownToHtml(source.markdown);
  else if (source.text) html = plainTextToHtml(source.text);
  else throw new OfficeError('需要提供 html / markdown / text 三者之一作为文档内容');
  const tmpHtml = io.tmpFile ? io.tmpFile('.html') : path + '.office.tmp.html';
  let converted;
  try {
    const title = String(params.title || '文档').replaceAll('<', '&lt;');
    // 临时文件落在公共 tmp 目录:含用户文档内容,显式 0600 只允许本人读
    await io.writeBuf(tmpHtml, Buffer.from(`<html><head><meta charset="utf-8"><title>${title}</title></head><body>${html}</body></html>`), { mode: 0o600 });
    converted = await converters.wordConvert(tmpHtml, path);
  } finally {
    if (io.remove) await io.remove(tmpHtml);
  }
  const st = await ensureRegular(io, path, CAPS.MAX_TEXT_BYTES * 8);
  return {
    content: `已生成 ${ext.toUpperCase()} 文档: ${path}\n- 大小: ${fmtBytes(st.size)}\n- 转换方式: ${converted.backend}` +
      `\n> 注: .doc/.rtf/.odt 为旧格式，版式支持有限；复杂样式建议用 .docx`,
    meta: { format: ext, via: converted.backend, bytes: st.size },
  };
}

// ---------------------------------------------------------------------------
// office_write_xlsx
// ---------------------------------------------------------------------------
export async function opWriteXlsx(path, spec, ioIn) {
  const io = normalizeIo(ioIn);
  const ext = extOf(path);
  if (ext !== 'xlsx') throw new OfficeError('office_write_xlsx 仅支持输出 .xlsx（旧 .xls 请先转存为 .xlsx）', 'UNSUPPORTED_FORMAT');
  if (!spec || !Array.isArray(spec.sheets) || !spec.sheets.length) {
    throw new OfficeError('需要 sheets 参数: [{name?, rows: 二维数组, header?}]');
  }
  const buf = await excel.buildWorkbook(spec);
  await writeOfficeBinary(io, path, buf, 'xlsx');
  const summary = await excel.workbookSummary(buf);
  const sheetsDesc = summary ? summary.sheets.map(sheetLabel).join('、') : '';
  return {
    content: `已生成 Excel 工作簿: ${path}\n- 大小: ${fmtBytes(buf.length)}\n- 工作表: ${sheetsDesc}`,
    meta: { format: 'xlsx', bytes: buf.length, sheets: summary?.sheets },
  };
}

/** Write an OOXML buffer, refusing to emit a text file under a binary extension. */
async function writeOfficeBinary(io, path, buf, ext) {
  const bytes = ZIP_OFFICE_EXTS.has(ext) ? assertOfficeBinary(buf, ext) : buf;
  await io.writeBuf(path, bytes);
  return bytes;
}

/** One sheet of the workbook summary as `名称(行×列)`. */
function sheetLabel(s) {
  return `${s.name}(${s.rows}行×${s.cols}列)`;
}

// ---------------------------------------------------------------------------
// office_edit_xlsx
// ---------------------------------------------------------------------------
export async function opEditXlsx(path, ops, ioIn) {
  const io = normalizeIo(ioIn);
  const ext = extOf(path);
  if (ext !== 'xlsx') {
    if (ext === 'xls') {
      throw new OfficeError('.xls 是旧格式，请先用 office_convert 转成 .xlsx 再编辑', 'NEED_CONVERT');
    }
    throw new OfficeError('office_edit_xlsx 仅支持 .xlsx', 'UNSUPPORTED_FORMAT');
  }
  const buf = await io.readBuf(path);
  const before = await excel.workbookSummary(buf);
  const { buf: out, changes } = await excel.editWorkbook(buf, ops);
  await writeOfficeBinary(io, path, out, 'xlsx');
  return {
    content: `已更新 Excel: ${path}\n- 执行操作: ${changes.join('、')}\n- 大小: ${fmtBytes(buf.length)} → ${fmtBytes(out.length)}`,
    meta: { format: 'xlsx', bytes: out.length, changes, before },
  };
}

// ---------------------------------------------------------------------------
// office_fill_docx_template
// ---------------------------------------------------------------------------
export async function opFillTemplate(templatePath, outputPath, data, ioIn) {
  const io = normalizeIo(ioIn);
  if (extOf(templatePath) !== 'docx' || extOf(outputPath) !== 'docx') {
    throw new OfficeError('模板填充仅支持 .docx → .docx（模板中占位符形如 {{变量名}}）', 'UNSUPPORTED_FORMAT');
  }
  const buf = await io.readBuf(templatePath);
  const out = await word.fillDocxTemplate(buf, data && typeof data === 'object' ? data : {});
  await writeOfficeBinary(io, outputPath, out, 'docx');
  const used = data && typeof data === 'object' ? Object.keys(data).length : 0;
  return {
    content: `已按模板生成文档: ${outputPath}\n- 模板: ${templatePath}\n- 填充变量数: ${used}\n- 大小: ${fmtBytes(out.length)}\n> 提示: 模板内请使用 {{变量名}} 占位（不支持循环/条件，如需批量套打可多次调用）`,
    meta: { format: 'docx', bytes: out.length, template: templatePath, variables: used },
  };
}

// ---------------------------------------------------------------------------
// office_convert
// ---------------------------------------------------------------------------
const WORD_SRC_EXTS = new Set(['doc', 'docx', 'rtf', 'odt', 'html', 'txt', 'md', 'markdown']);
const WORD_DST_EXTS = new Set(['docx', 'doc', 'rtf', 'odt', 'html', 'txt', 'pdf']);
const TABLE_SRC_EXTS = new Set(['xlsx', 'xls', 'xlsb', 'ods', 'csv', 'tsv']);
const TABLE_DST_EXTS = new Set(['xlsx', 'xls', 'xlsb', 'ods', 'csv', 'tsv', 'html']);
const LEGACY_WORD_EXTS = new Set(['doc', 'rtf', 'odt']);

export async function opConvert(srcPath, dstPath, ioIn) {
  const io = normalizeIo(ioIn);
  const srcExt = extOf(srcPath);
  const dstExt = extOf(dstPath);
  if (!dstExt) throw new OfficeError('转换目标需带扩展名, 如 out.xlsx');
  if (srcExt === dstExt) return sameFormat(io, srcPath, dstPath, srcExt);
  if (WORD_SRC_EXTS.has(srcExt)) return convertWordFamily(io, srcPath, srcExt, dstPath, dstExt);
  if (TABLE_SRC_EXTS.has(srcExt)) return convertTableFamily(io, srcPath, srcExt, dstPath, dstExt);
  throw new OfficeError(`不支持的转换组合: .${srcExt} → .${dstExt}`, 'UNSUPPORTED_CONVERSION');
}

/** Same extension: nothing to convert, but a different target needs the bytes copied. */
async function sameFormat(io, srcPath, dstPath, ext) {
  if (srcPath === dstPath) {
    return { content: `源与目标为同一文件(.${ext})，无需转换`, meta: { same: true, copied: false } };
  }
  await writeOfficeBinary(io, dstPath, await io.readBuf(srcPath), ext);
  return { content: `源与目标格式相同(.${ext})，已原样复制到 ${dstPath}`, meta: { same: true, copied: true } };
}

async function sourceToHtml(io, srcPath, srcExt) {
  const buf = await io.readBuf(srcPath);
  if (srcExt === 'docx') return (await word.readDocx(buf)).html;
  const text = buf.toString('utf8');
  if (srcExt === 'md' || srcExt === 'markdown') {
    const { markdownToHtml } = await import('./md.js');
    return markdownToHtml(text);
  }
  if (srcExt === 'txt') {
    const { plainTextToHtml } = await import('./md.js');
    return plainTextToHtml(text);
  }
  return text;
}

function convertResult(srcPath, srcExt, dstPath, dstExt, bytes, via, note) {
  return {
    content: `转换完成: ${srcPath} (.${srcExt}) → ${dstPath} (.${dstExt})\n- 大小: ${fmtBytes(bytes)}` +
      (via ? `\n- 转换方式: ${via}` : '') + (note ? `\n> 注: ${note}` : ''),
    meta: { from: srcExt, to: dstExt, via, bytes },
  };
}

async function convertViaExternal(io, srcPath, srcExt, dstPath, dstExt) {
  const { backend } = await converters.wordConvert(srcPath, dstPath);
  const st = await ensureRegular(io, dstPath, CAPS.MAX_EXCEL_INPUT_BYTES * 2);
  return convertResult(srcPath, srcExt, dstPath, dstExt, st.size, backend, '旧格式转换，复杂版式可能简化');
}

async function convertHtmlViaExternal(io, html, srcPath, srcExt, dstPath, dstExt) {
  const tmp = io.tmpFile ? io.tmpFile('.html') : dstPath + '.office.tmp.html';
  await io.writeBuf(tmp, Buffer.from('<html><head><meta charset="utf-8"></head><body>' + html + '</body></html>'), { mode: 0o600 });
  try {
    return await convertViaExternal(io, tmp, srcExt, dstPath, dstExt);
  } finally {
    if (io.remove) await io.remove(tmp);
  }
}

async function convertWordFamily(io, srcPath, srcExt, dstPath, dstExt) {
  if (TABLE_DST_EXTS.has(dstExt)) {
    throw new OfficeError(`不支持 ${srcExt} → ${dstExt}：Word/文本无法直接转成表格`, 'UNSUPPORTED_CONVERSION');
  }
  if (!WORD_DST_EXTS.has(dstExt)) {
    throw new OfficeError(`不支持的转换组合: .${srcExt} → .${dstExt}`, 'UNSUPPORTED_CONVERSION');
  }
  if (LEGACY_WORD_EXTS.has(srcExt)) return convertViaExternal(io, srcPath, srcExt, dstPath, dstExt);

  const html = await sourceToHtml(io, srcPath, srcExt);
  if (dstExt === 'html') {
    await io.writeBuf(dstPath, Buffer.from(html));
    const st = await ensureRegular(io, dstPath, CAPS.MAX_TEXT_BYTES * 8);
    return convertResult(srcPath, srcExt, dstPath, dstExt, st.size, 'HTML 直出');
  }
  if (dstExt === 'txt') {
    await io.writeBuf(dstPath, Buffer.from(await md.htmlToMarkdown(html)));
    const st = await ensureRegular(io, dstPath, CAPS.MAX_TEXT_BYTES * 8);
    return convertResult(srcPath, srcExt, dstPath, dstExt, st.size, '文本重建');
  }
  if (dstExt === 'docx') {
    const buf = await word.writeDocx({ html }, { title: '转换文档' });
    await writeOfficeBinary(io, dstPath, buf, 'docx');
    return convertResult(srcPath, srcExt, dstPath, dstExt, buf.length, 'HTML→docx 重建', '复杂版式可能简化');
  }
  return convertHtmlViaExternal(io, html, srcPath, srcExt, dstPath, dstExt);
}

async function convertTableFamily(io, srcPath, srcExt, dstPath, dstExt) {
  if (!TABLE_DST_EXTS.has(dstExt)) {
    throw new OfficeError(`不支持 ${srcExt} → ${dstExt}：表格只能转表格/CSV/HTML`, 'UNSUPPORTED_CONVERSION');
  }
  const buf = await io.readBuf(srcPath);
  const XLSX = await loadSheetjs();
  const wb = legacy.readSheetJs(XLSX, buf, srcExt);
  if (legacy.TEXT_TABLE_EXTS.has(srcExt)) legacy.neutralizeTextFormulas(wb);
  const out = writeTableBuffer(XLSX, wb, dstExt);
  await writeOfficeBinary(io, dstPath, out, dstExt);
  const st = await ensureRegular(io, dstPath, CAPS.MAX_EXCEL_INPUT_BYTES * 2);
  return convertResult(srcPath, srcExt, dstPath, dstExt, st.size, 'SheetJS');
}

function writeTableBuffer(XLSX, wb, dstExt) {
  if (dstExt === 'xls') {
    try {
      return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'biff8' }));
    } catch {
      throw new OfficeError('.xls 导出当前不可用(SheetJS 社区版限制)，请改输出 .xlsx 后用 WPS/Office 另存为 .xls', 'XLS_WRITE_UNSUPPORTED');
    }
  }
  if (dstExt === 'xlsx') return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', compression: true }));
  // SheetJS 没有 tsv bookType:按 CSV 写但把分隔符换成制表符
  if (dstExt === 'tsv') return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'csv', FS: '\t' }));
  return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: dstExt }));
}

let _sheetjsMod = null;
async function loadSheetjs() {
  if (!_sheetjsMod) _sheetjsMod = await import('@e965/xlsx');
  return _sheetjsMod;
}
