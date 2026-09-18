// Office operation layer: pure operations on absolute paths + buffers.
// IO is injected ({ readBuf, writeBuf, stat }) so the DSH host adapter can
// route through ctx.fs / sandbox policy; the CLI selftest uses node:fs.
import { OfficeError, extOf, fmtBytes, readPath, writePath, truncate, CAPS, ZIP_OFFICE_EXTS, assertOfficeBinary } from './util.js';
import { tmpdir } from 'node:os';
import { randomBytes } from 'node:crypto';
import * as legacy from './legacy.js';
import { openOoxml, officePartOf } from './ooxml.js';
import { PLUGIN_VERSION } from './version.js';
import { lazyModule } from './deps.js';
import { formatReport } from './docx-format.js';
import { joinNote, joinTables, planJoins } from './join.js';
import { buildTable, headerRowOf, queryTable } from './query.js';
import { editDocx } from './docx-edit.js';

// 重依赖懒加载:依赖装不全时插件仍能加载,且只有真正用到该模块的工具会失败,
// 报错会明确指出缺哪个包(见 deps.js)
const word = lazyModule('./word.js', 'Word 读写', ['readDocx', 'writeDocx', 'fillDocxTemplate']);
const excel = lazyModule('./excel.js', 'Excel 读写', ['readWorkbook', 'readMatrixSheets', 'buildWorkbook', 'editWorkbook', 'workbookSummary']);
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
/** 文件真实内容形态:改后缀是常见操作,不能只信扩展名。 */
/** office_read 支持的扩展名(读之前先按扩展名做一次白名单校验)。 */
const READABLE_EXTS = new Set(['docx', 'doc', 'odt', 'rtf', 'xlsx', 'xls', 'xlsb', 'ods', 'csv', 'tsv', 'txt']);

const WORD_LEGACY_EXTS = new Set(['doc', 'rtf', 'odt']);
/**
 * 纯文本表格嗅探:无 NUL 字节且含制表符/逗号,就按 .tsv / .csv 处理。
 * 只读前 64 KB,避免大文件全量扫描。
 */
function delimitedExtOf(buffer) {
  const head = buffer.subarray(0, 65536);
  if (head.includes(0)) return '';
  const text = head.toString('utf8');
  if (!text.trim()) return '';
  const tabs = (text.match(/\t/g) || []).length;
  const commas = (text.match(/,/g) || []).length;
  if (tabs >= 1 && tabs >= commas) return 'tsv';
  if (commas >= 1) return 'csv';
  return '';
}

/**
 * 嗅探真实格式:OLE2(d0cf11e0)是老 .doc/.xls;zip 则看主文档部件在哪个家族。
 * 返回 { ext, buffer, entries, note } —— ext 是探测出来的真实格式。
 * `preloaded` 允许调用方把已经读到的字节传进来,避免二次读盘。
 */
async function sniffFormat(io, path, ext, preloaded) {
  const buffer = preloaded ?? (await io.readBuf(path));
  const magic = buffer.subarray(0, 4);
  const isOle = magic[0] === 0xd0 && magic[1] === 0xcf && magic[2] === 0x11 && magic[3] === 0xe0;
  if (isOle) {
    const real = WORD_LEGACY_EXTS.has(ext) || OFFICE_EXTENSIONS.word.includes(ext) ? 'doc' : 'xls';
    return { ext: real, buffer, entries: [], note: real === ext ? '' : `文件内容其实是 OLE 老格式(.${real})` };
  }
  const isZip = magic[0] === 0x50 && magic[1] === 0x4b;
  if (!isZip) {
    // 导出成 CSV/TSV 却存成 .xlsx 很常见,按内容判断
    const text = delimitedExtOf(buffer);
    return { ext: text || ext, buffer, entries: [], note: text && text !== ext ? `文件内容其实是分隔符文本(.${text})` : '' };
  }
  const { zip, repaired } = await openOoxml(buffer, ext);
  const entries = Object.keys(zip.files).filter((name) => !name.endsWith('/'));
  // 主部件优先按包关系解析:有些文件把它放在非规范路径(word/main.xml)
  const main = officePartOf(zip);
  if (!main) {
    return {
      ext,
      buffer,
      entries,
      error: `这个 .${ext} 是个 zip,但里面没有 Office 主文档(插件 v${PLUGIN_VERSION};实际条目: ${entries.slice(0, 8).join('、') || '空'})。` +
        '多半是被改过后缀、或第三方工具生成的损坏文件 —— 请用 Word 打开后另存为 .docx 再试',
    };
  }
  const note = main.kind === ext ? '' : `文件内容其实是 .${main.kind}`;
  const repairNote = repaired.length ? `(自动修正了 ${repaired.length} 个非标准 zip 条目名)` : '';
  return { ext: main.kind, buffer, entries, note: `${note}${repairNote}` };
}

export async function opRead(path, rawOpts, ioIn) {
  const opts = rawOpts || {};
  const io = normalizeIo(ioIn);
  const declared = extOf(path);
  // 先保证错误语义不变:不支持的扩展名 / 文件不存在 / 传了目录
  if (!READABLE_EXTS.has(declared)) {
    throw new OfficeError(`office_read 不支持的扩展名 .${declared}（支持: docx/doc/odt/rtf/xlsx/xls/xlsb/ods/csv/tsv）`, 'UNSUPPORTED_FORMAT');
  }
  await ensureRegular(io, path, CAPS.MAX_EXCEL_INPUT_BYTES);
  // 再按真实内容判断(改后缀、非标准 zip 都在这里被纠正)
  const sniffed = await sniffFormat(io, path, declared);
  if (sniffed.error) throw new OfficeError(sniffed.error, 'BAD_CONTAINER');
  // 复用嗅探时读到的字节,避免二次读盘
  const readBuf = io.readBuf;
  let first = true;
  io.readBuf = async (p) => {
    if (first && p === path) {
      first = false;
      return sniffed.buffer;
    }
    return readBuf(p);
  };
  const ext = sniffed.ext;
  // 绝不在调用方传入的对象上写属性:DSH 宿主传进来的参数是冻结的
  const readOpts = { ...opts, formulaMode: formulaModeOf(opts), ...(sniffed.note ? { containerNote: sniffed.note } : {}) };
  if (ext === 'docx') return readDocxFile(io, path, readOpts);
  if (['doc', 'rtf', 'odt'].includes(ext)) return readLegacyWordFile(io, path, ext, readOpts);
  if (ext === 'xlsx') return readXlsxFile(io, path, readOpts);
  return readLegacyTableFile(io, path, ext, readOpts);
}

/** 公式显示模式:value(默认,计算值)/ formula(公式本体)/ both(两者)。true 当作 formula。 */
function formulaModeOf(opts) {
  const raw = opts.formulas;
  if (raw === undefined || raw === null || raw === false || raw === '') return 'value';
  const mode = raw === true ? 'formula' : String(raw);
  if (!['value', 'formula', 'both'].includes(mode)) {
    throw new OfficeError('office_read 的 formulas 可为 value（计算值，默认）/ formula（公式本体）/ both（公式 + 计算值）', 'INVALID_ARGS');
  }
  return mode;
}

// ---------------------------------------------------------------------------
// Word 分段阅读:超大文档按字符偏移续读,并可先取标题大纲
// ---------------------------------------------------------------------------
/** 行级线性扫描取 Markdown 标题(跳过围栏代码块)及其字符偏移。 */
export function outlineOf(text) {
  const items = [];
  let offset = 0;
  let fenced = false;
  for (const line of text.split('\n')) {
    if (line.startsWith('```')) fenced = !fenced;
    else if (!fenced) pushHeading(items, line, offset);
    offset += line.length + 1;
  }
  return items;
}

function pushHeading(items, line, offset) {
  let level = 0;
  while (level < line.length && line[level] === '#') level += 1;
  if (level < 1 || level > 6 || line[level] !== ' ') return;
  items.push({ level, title: line.slice(level + 1).trim(), offset });
}

/** 可作分页收尾的字符:换行与句末标点。Word 转出的 Markdown 段落常以 <br> 变空格,
 *  所以只认换行不够用,还得认句号类标点,否则每页都会切在半句话上。 */
const BOUNDARY_MARKS = ['\n', '。', '！', '？', '；', '!', '?', ';'];

/** 窗口内最靠后的分页边界(收尾位置须达到窗口 60%,否则宁可硬切)。 */
function boundaryOf(raw, min) {
  let best = -1;
  for (const mark of BOUNDARY_MARKS) {
    const end = raw.lastIndexOf(mark) + 1;
    if (end >= min && end > best) best = end;
  }
  return best;
}

/**
 * 分页窗口:从 offset 取 maxChars,尽量在句子/段落边界收尾。
 * 切片与下一段的 offset 严格衔接(边界字符归本段),逐段拼回来等于全文,不重不漏。
 */
export function pageOf(text, offset, maxChars) {
  const limit = Math.min(offset + maxChars, text.length);
  if (limit >= text.length) return { slice: text.slice(offset), end: text.length, more: false };
  const raw = text.slice(offset, limit);
  const cut = boundaryOf(raw, Math.floor(maxChars * 0.6));
  if (cut <= 0) return { slice: raw, end: limit, more: true };
  return { slice: raw.slice(0, cut), end: offset + cut, more: true };
}

/** 校验 offset / outline:html 模式不做分页(按字符切会切坏标签)。 */
function pageOffset(opts, asHtml) {
  const hasOffset = opts.offset !== undefined && opts.offset !== null && opts.offset !== '';
  if (asHtml && (hasOffset || opts.outline)) {
    throw new OfficeError('format: "html" 下不支持 offset / outline：按字符切会切坏标签。分段阅读请用默认的 text 模式', 'INVALID_ARGS');
  }
  if (!hasOffset) return 0;
  const n = Number(opts.offset);
  if (!Number.isInteger(n) || n < 0) throw new OfficeError('offset 需为 ≥ 0 的整数（正文起始字符位置）', 'INVALID_ARGS');
  return n;
}

const tsvCell = (s) => String(s).replaceAll('\t', '␉').replaceAll(/\r?\n/g, '⏎');

function outlineBlock(items) {
  if (!items.length) return '正文里没有 Markdown 标题（没有 # 开头的行），无法给大纲；可直接从 offset 0 分页读。';
  const lines = ['```tsv', '级别\t偏移\t标题'];
  for (const h of items) lines.push(`${h.level}\t${h.offset}\t${tsvCell(h.title)}`);
  lines.push('```');
  const sample = items[Math.min(1, items.length - 1)].offset;
  return lines.join('\n') + `\n> 用法: 传 offset 从某个标题开始读，例如 offset: ${sample}`;
}

/**
 * Word 正文的统一输出:`outline` 只给标题大纲,否则按 offset/maxChars 分页。
 * 两者都带上总字符数,让调用方知道还有没有下一段 —— 不再"每次都从头读"。
 */
function wordBody({ path, markdown, opts, offset, info, meta, appendix }) {
  const total = markdown.length;
  if (opts.outline) {
    const items = outlineOf(markdown);
    return {
      content: `已读取 Word 文档 ${path}${info}\n- 标题大纲: ${items.length} 条 / 正文共 ${total} 字符\n\n${outlineBlock(items)}`,
      meta: { ...meta, totalChars: total, headings: items.length, outline: items },
    };
  }
  if (offset >= total) {
    return {
      content: `已到文档末尾 ${path}${info}\n- 正文共 ${total} 字符，offset=${offset} 已越界；要重读请传更小的 offset`,
      meta: { ...meta, totalChars: total, offset, ended: true, truncated: false },
    };
  }
  const page = pageOf(markdown, offset, opts.maxChars ?? CAPS.DEFAULT_MAX_CHARS);
  const remain = page.more ? `（剩余 ${total - page.end}）` : '（已到文末）';
  const range = `- 正文: 第 ${offset}–${page.end} 字符 / 共 ${total} 字符${remain}`;
  const next = page.more ? `\n- 继续读: 再调用本工具并传 offset: ${page.end}` : '';
  return {
    content: `已读取 Word 文档 ${path}${info}\n${range}${next}\n\n${page.slice}${appendix || ''}`,
    meta: { ...meta, totalChars: total, offset, end: page.end, truncated: page.more },
  };
}

async function readDocxFile(io, path, opts) {
  const asHtml = opts.format === 'html';
  const wantFormat = Boolean(opts.withFormatting);
  const offset = pageOffset(opts, asHtml);
  const buf = await io.readBuf(path);
  const r = await word.readDocx(buf, { formatting: wantFormat });
  const info = `\n- 大小: ${fmtBytes(buf.length)}\n- 词数约: ${r.meta.words}, 段落约: ${r.meta.paragraphs}${containerNote(opts)}`;
  if (asHtml) return htmlWordBody(path, r.html, info, opts, { kind: 'word', format: 'docx' });
  const wantReport = wantFormat && r.formatting;
  const appendix = wantReport && offset === 0 ? `\n\n## 格式报告（整篇）\n\n${formatReport(r.formatting)}` : '';
  const lateReport = wantReport && offset > 0 ? '\n- 注: 格式报告只在 offset=0 时给出' : '';
  return wordBody({
    path,
    markdown: (await md.htmlToMarkdown(r.html)) || r.content,
    opts,
    offset,
    info: info + lateReport,
    meta: {
      kind: 'word',
      format: 'docx',
      bytes: buf.length,
      words: r.meta.words,
      paragraphs: r.meta.paragraphs,
      formatting: wantFormat ? summarizeFormatting(r.formatting) : undefined,
    },
    appendix,
  });
}

/** html 输出不分页:只按 maxChars 截断,并说明想分段该切回 text 模式。 */
function htmlWordBody(path, html, info, opts, meta) {
  const capped = truncate(html, opts.maxChars ?? CAPS.DEFAULT_MAX_CHARS);
  return {
    content: `已读取 Word 文档 ${path}${info}\n> 格式: html（不分页；要分段读长文请用默认 text 模式 + offset）\n\n${capped}`,
    meta: { ...meta, htmlChars: html.length },
    html,
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
  const asHtml = opts.format === 'html';
  const offset = pageOffset(opts, asHtml);
  const read = await converters.wordRead(path, { asHtml });
  const formatNote = opts.withFormatting ? '\n- 注: 字体/行距等格式提取目前仅支持 .docx' : '';
  const info = `\n- 解析方式: ${read.backend}，复杂版式可能丢失${formatNote}${containerNote(opts)}`;
  if (asHtml) {
    const capped = truncate(read.content, opts.maxChars ?? CAPS.DEFAULT_MAX_CHARS);
    return {
      content: `已读取 ${ext.toUpperCase()} 文档 ${path}${info}\n> 格式: html（不分页；要分段读长文请用默认 text 模式 + offset）\n\n${capped}`,
      meta: { kind: 'word', format: ext, via: read.backend },
    };
  }
  const markdown = (await md.htmlToMarkdown(read.content)) || read.content;
  return wordBody({
    path,
    markdown,
    opts,
    offset,
    info,
    meta: { kind: 'word', format: ext, via: read.backend, words: roughWords(markdown) },
  });
}

/** 词数粗估(取当前分页窗口,避免为了报数再扫全文)。 */
function roughWords(text) {
  return (text.match(/[\u4e00-\u9fff]|[A-Za-z0-9]+/g) || []).length;
}

async function readXlsxFile(io, path, opts) {
  await ensureRegular(io, path, CAPS.MAX_EXCEL_INPUT_BYTES);
  const buf = await io.readBuf(path);
  const r = await excel.readWorkbook(buf, opts);
  return { content: `已读取 Excel 工作簿 ${path}（${r.meta.sheetNames.length} 个工作表）${containerNote(opts)}\n\n` + r.content, meta: r.meta };
}

async function readLegacyTableFile(io, path, ext, opts) {
  await ensureRegular(io, path, CAPS.MAX_EXCEL_INPUT_BYTES);
  const buf = await io.readBuf(path);
  const r = await legacy.readLegacySpreadsheet(buf, ext, opts);
  return { content: `已读取 ${ext.toUpperCase()} 表格 ${path}（${r.meta.sheets.length} 个工作表）${containerNote(opts)}\n\n` + r.content, meta: r.meta };
}

// ---------------------------------------------------------------------------
// office_query —— 在文件内直接算:筛选 / 分组 / 聚合 / 排序 / 画像
// 只看结论,不把几万行搬进上下文;算不了的需求(跨文件 join、透视、建模)才需要写脚本。
// ---------------------------------------------------------------------------
const QUERYABLE_EXTS = new Set(['xlsx', 'xls', 'xlsb', 'ods', 'csv', 'tsv', 'txt']);

/** 按名称或从 1 开始的序号挑工作表。 */
function pickSheet(loaded, want) {
  const { sheets, sheetNames } = loaded;
  if (!sheets.length) throw new OfficeError('这个文件里没有可读的工作表', 'BAD_XLSX');
  if (want === undefined || want === null || want === '') return sheets[0];
  const index = Number(want);
  if (Number.isInteger(index) && index >= 1 && index <= sheets.length) return sheets[index - 1];
  const hit = sheets.find((s) => s.name === String(want));
  if (hit) return hit;
  throw new OfficeError(`找不到工作表「${want}」。可用：${(sheetNames || []).join('、')}（也可用序号，从 1 开始）`, 'UNKNOWN_SHEET');
}

export async function opQuery(path, params, ioIn) {
  const opts = params || {};
  const io = normalizeIo(ioIn);
  const declared = extOf(path);
  if (!READABLE_EXTS.has(declared)) {
    throw new OfficeError(`office_query 不支持的扩展名 .${declared}（支持表格: xlsx/xls/xlsb/ods/csv/tsv）`, 'UNSUPPORTED_FORMAT');
  }
  await ensureRegular(io, path, CAPS.MAX_EXCEL_INPUT_BYTES);
  const buffer = await io.readBuf(path);
  const sniffed = await sniffFormat(io, path, declared, buffer);
  if (sniffed.error) throw new OfficeError(sniffed.error, 'BAD_CONTAINER');
  if (!QUERYABLE_EXTS.has(sniffed.ext)) {
    throw new OfficeError(`office_query 只能算表格，这个文件其实是 .${sniffed.ext} 文档 —— 要读内容请用 office_read`, 'UNSUPPORTED_FORMAT');
  }
  const loaded = await loadQuerySheets(io, path, sniffed.ext, buffer);
  const raw = pickSheet(loaded, opts.sheet);
  const sheet = tableOfRaw(raw, opts.headerRow);
  const joins = planJoins(opts.join, path);
  const target = joins.length ? await applyJoins(io, sheet, joins) : sheet;
  const r = queryTable(target, opts);
  const others = loaded.sheetNames.length > 1
    ? `\n- 工作表: ${loaded.sheetNames.join('、')}（本次只算「${sheet.name}」）`
    : '';
  const note = sniffed.note ? `\n- 注: ${sniffed.note}` : '';
  return {
    content: `已计算表格 ${path}${note}${others}\n\n${r.text}`,
    meta: { kind: 'table-query', format: sniffed.ext, ...r.meta },
  };
}

/** 原始矩阵 + headerRow → 查询用的表(列名、数据行、截断与总行数)。 */
function tableOfRaw(raw, headerRow) {
  const matrix = Array.isArray(raw.matrix) ? raw.matrix : [];
  const row = headerRowOf(headerRow, matrix.length);
  const table = buildTable(matrix, row);
  return {
    name: raw.name,
    columns: table.columns,
    rows: table.rows,
    headerRow: row,
    truncated: Boolean(raw.truncated),
    totalRows: Number(raw.totalRows) || matrix.length,
  };
}

/** 读一个可查询文件的全部工作表原始矩阵(xlsx 走 ExcelJS,其余走 SheetJS)。 */
async function loadQuerySheets(io, path, ext, preloaded) {
  const buffer = preloaded ?? (await io.readBuf(path));
  const scan = { maxRows: CAPS.MAX_QUERY_ROWS, maxCols: CAPS.MAX_QUERY_COLS };
  return ext === 'xlsx' ? excel.readMatrixSheets(buffer, scan) : legacy.readMatrixSheets(buffer, ext, scan);
}

/** 连接用的右侧表:单独按扩展名与真实内容挑读取器,可以是另一个文件,也可以是同文件的另一张表。 */
async function loadQueryTable(io, path, spec) {
  const ext = extOf(path);
  if (!READABLE_EXTS.has(ext)) {
    throw new OfficeError(`join 的 path 不支持 .${ext}（支持表格: xlsx/xls/xlsb/ods/csv/tsv）`, 'UNSUPPORTED_FORMAT');
  }
  await ensureRegular(io, path, CAPS.MAX_EXCEL_INPUT_BYTES);
  const buffer = await io.readBuf(path);
  const sniffed = await sniffFormat(io, path, ext, buffer);
  if (sniffed.error) throw new OfficeError(sniffed.error, 'BAD_CONTAINER');
  if (!QUERYABLE_EXTS.has(sniffed.ext)) {
    throw new OfficeError(`join 的 path 指向的是 .${sniffed.ext} 文档，不是表格`, 'UNSUPPORTED_FORMAT');
  }
  const loaded = await loadQuerySheets(io, path, sniffed.ext, buffer);
  const raw = pickSheet(loaded, spec.sheet);
  const matrix = Array.isArray(raw.matrix) ? raw.matrix : [];
  const headerRow = headerRowOf(spec.headerRow === undefined || spec.headerRow === null ? 1 : spec.headerRow, matrix.length);
  const table = buildTable(matrix, headerRow);
  return { name: spec.label || raw.name, columns: table.columns, rows: table.rows, headerRow };
}

/** 依次连接(多条 join 时后一条接在前一条的结果上,列名以连接后的结果为准)。 */
async function applyJoins(io, first, joins) {
  let table = first;
  const notes = [];
  for (const spec of joins) {
    const right = await loadQueryTable(io, spec.path, spec);
    const result = joinTables(table, { ...right, name: spec.label || right.name }, spec);
    table = { name: table.name, columns: result.columns, rows: result.rows, headerRow: table.headerRow, truncated: table.truncated, totalRows: result.rows.length };
    notes.push(joinNote(spec, right, result));
  }
  return { ...table, notes };
}

// ---------------------------------------------------------------------------
// office_write_docx
// ---------------------------------------------------------------------------
export async function opWriteDocx(path, params, ioIn, opts = {}) {
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
    style: params.style,
    header: params.header,
    footer: params.footer,
    toc: params.toc,
    readImage: opts.readImage,
  });
  await writeOfficeBinary(io, path, buf, 'docx');
  const from = contentSourceLabel(source);
  return { content: `已生成 Word 文档: ${path}\n- 大小: ${fmtBytes(buf.length)}\n- 内容来源: ${from}`, meta: { format: 'docx', bytes: buf.length } };
}

/** 嗅探到的「扩展名与实际格式不符 / 已修复非标准 zip」提示。 */
function containerNote(opts) {
  return opts?.containerNote ? `\n- 注: ${opts.containerNote}` : '';
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
// office_edit_docx —— 就地改已有 .docx(只重写主文档部件,其余部件原样保留)
// ---------------------------------------------------------------------------
/** 结果首行:区分试运行 / 另存 / 就地覆盖。 */
function editHead({ dry, path, target, outputPath }) {
  if (dry) {
    const would = outputPath ? `（本应另存为 ${target}）` : '';
    return `已试运行(未写入): ${path}${would}`;
  }
  if (outputPath) return `已另存为 ${target}（原文件未改动）`;
  return `已更新 Word 文档: ${target}`;
}

/**
 * 就地(或另存)修改 .docx。
 * @param opts.dryRun 只算不写:返回改动摘要,原文件与 outputPath 都不动
 * @param opts.outputPath 另存到新文件,不覆盖原文件
 */
export async function opEditDocx(path, ops, ioIn, opts = {}) {
  const io = normalizeIo(ioIn);
  const ext = extOf(path);
  if (ext !== 'docx') {
    if (ext === 'doc') {
      throw new OfficeError('.doc 是旧格式，请先用 office_convert 转成 .docx 再编辑', 'NEED_CONVERT');
    }
    throw new OfficeError('office_edit_docx 仅支持 .docx', 'UNSUPPORTED_FORMAT');
  }
  const target = opts.outputPath || path;
  if (opts.outputPath && extOf(target) !== 'docx') {
    throw new OfficeError(`另存的目标应为 .docx，实际 .${extOf(target)}`, 'UNSUPPORTED_FORMAT');
  }
  await ensureRegular(io, path, CAPS.MAX_WORD_INPUT_BYTES);
  const buf = await io.readBuf(path);
  const { buf: out, changes, paragraphsBefore, paragraphsAfter } = await editDocx(buf, ops);
  const dry = Boolean(opts.dryRun);
  if (!dry) await writeOfficeBinary(io, target, out, 'docx');
  const head = editHead({ dry, path, target, outputPath: opts.outputPath });
  return {
    content: `${head}\n- 执行操作: ${changes.join('、')}\n- 段落数: ${paragraphsBefore} → ${paragraphsAfter}` +
      `\n- 大小: ${fmtBytes(buf.length)} → ${fmtBytes(out.length)}${dry ? '（预估）' : ''}` +
      '\n> 注: 只重写了正文部件，图片/页眉页脚/样式/批注等其余部件原样保留',
    meta: { format: 'docx', bytes: out.length, changes, paragraphsBefore, paragraphsAfter, dryRun: dry, outputPath: opts.outputPath || undefined },
  };
}

// ---------------------------------------------------------------------------
// office_fill_docx_template
// ---------------------------------------------------------------------------
// 注意:本文件返回给模型的文本里不要写 `{{变量名}}` 这种字面量 —— DSH 的提示词渲染器
// 会把 {{...}} 当成变量引用(见 lib/index.js 里 fillTemplateTool 的说明)。
export async function opFillTemplate(templatePath, outputPath, data, ioIn) {
  const io = normalizeIo(ioIn);
  if (extOf(templatePath) !== 'docx' || extOf(outputPath) !== 'docx') {
    throw new OfficeError('模板填充仅支持 .docx → .docx（模板里用两个半角花括号包住变量名作占位）', 'UNSUPPORTED_FORMAT');
  }
  const buf = await io.readBuf(templatePath);
  const out = await word.fillDocxTemplate(buf, data && typeof data === 'object' ? data : {});
  await writeOfficeBinary(io, outputPath, out, 'docx');
  const used = data && typeof data === 'object' ? Object.keys(data).length : 0;
  return {
    content: `已按模板生成文档: ${outputPath}\n- 模板: ${templatePath}\n- 填充变量数: ${used}\n- 大小: ${fmtBytes(out.length)}\n> 提示: 模板里用「两个左花括号 + 变量名 + 两个右花括号」作占位（不支持循环/条件，如需批量套打可多次调用）`,
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

export async function opConvert(srcPath, dstPath, ioIn, opts = {}) {
  const io = normalizeIo(ioIn);
  const srcExt = extOf(srcPath);
  const dstExt = extOf(dstPath);
  if (!dstExt) throw new OfficeError('转换目标需带扩展名, 如 out.xlsx');
  if (srcExt === dstExt) return sameFormat(io, srcPath, dstPath, srcExt);
  if (WORD_SRC_EXTS.has(srcExt)) return convertWordFamily(io, srcPath, srcExt, dstPath, dstExt, opts);
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

async function convertViaExternal(io, srcPath, srcExt, dstPath, dstExt, opts = {}) {
  const { backend } = await converters.wordConvert(srcPath, dstPath, opts);
  const st = await ensureRegular(io, dstPath, CAPS.MAX_EXCEL_INPUT_BYTES * 2);
  return convertResult(srcPath, srcExt, dstPath, dstExt, st.size, backend, '旧格式转换，复杂版式可能简化');
}

async function convertHtmlViaExternal(io, html, srcPath, srcExt, dstPath, dstExt, opts = {}) {
  const tmp = io.tmpFile ? io.tmpFile('.html') : dstPath + '.office.tmp.html';
  await io.writeBuf(tmp, Buffer.from('<html><head><meta charset="utf-8"></head><body>' + html + '</body></html>'), { mode: 0o600 });
  try {
    return await convertViaExternal(io, tmp, srcExt, dstPath, dstExt, opts);
  } finally {
    if (io.remove) await io.remove(tmp);
  }
}

async function convertWordFamily(io, srcPath, srcExt, dstPath, dstExt, opts = {}) {
  if (TABLE_DST_EXTS.has(dstExt)) {
    throw new OfficeError(`不支持 ${srcExt} → ${dstExt}：Word/文本无法直接转成表格`, 'UNSUPPORTED_CONVERSION');
  }
  if (!WORD_DST_EXTS.has(dstExt)) {
    throw new OfficeError(`不支持的转换组合: .${srcExt} → .${dstExt}`, 'UNSUPPORTED_CONVERSION');
  }
  if (LEGACY_WORD_EXTS.has(srcExt)) return convertViaExternal(io, srcPath, srcExt, dstPath, dstExt, opts);

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
    const buf = await word.writeDocx({ html }, { title: '转换文档', readImage: opts.readImage });
    await writeOfficeBinary(io, dstPath, buf, 'docx');
    return convertResult(srcPath, srcExt, dstPath, dstExt, buf.length, 'HTML→docx 重建', '复杂版式可能简化');
  }
  return convertHtmlViaExternal(io, html, srcPath, srcExt, dstPath, dstExt, opts);
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
