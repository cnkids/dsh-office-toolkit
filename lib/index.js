// dsh-office-toolkit — DeepSeek Harness host plugin (host-only, no client half).
// Registers Word/Excel tools on the agent tool registry. Binary IO goes through
// node:fs after ctx.fs resolution + a containment fence; reads use ctx.fs.readBytes
// when available, and fs/observed events are emitted so the session's file
// observation state (and therefore the official read/write/edit guards) stay coherent.
import { dirname, isAbsolute, join, resolve as resolvePath } from 'node:path';
import { realpathSync } from 'node:fs';
import { mkdir, readFile, stat as fsStat, unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { randomBytes } from 'node:crypto';
import * as office from './core/office.js';
import { OfficeError } from './core/util.js';
import { isInsideAny } from './core/path-guard.js';
import { CORE_DEPS, INSTALL_COMMAND, missingDeps } from './core/deps.js';
import { PLUGIN_VERSION } from './core/version.js';

export const name = 'office-toolkit';
export const inject = ['tools'];

const OUTPUT = {
  schema: {
    type: 'object',
    additionalProperties: false,
    required: ['content'],
    properties: { content: { type: 'string' } },
  },
  render: (_args, value) => [{ type: 'text', text: String(value?.content || '') }],
};

const PATH_DESC = '文件路径。可用绝对路径，或相对会话工作区的路径。';

// 通用 read/write/edit 工具只能处理 UTF-8 文本(读二进制会直接报 binary file),
// 而工具注册没有优先级机制 —— 模型唯一能看到的信号就是 description。因此把
// 「哪类文件必须用本插件」放在每个描述的最前面,避免模型先拿通用工具试错一轮。
const WORD_EXTS = '.docx .doc .rtf .odt .pdf';
const TABLE_EXTS = '.xlsx .xls .xlsb .ods .csv .tsv';
const WORD_FAMILY = 'doc/docx/rtf/odt/html/txt/md';
const TABLE_FAMILY = 'xlsx/xls/xlsb/ods/csv/tsv/html';
const BINARY_HINT = 'Word/Excel 是二进制格式,通用 read/write/edit 工具处理不了(会报 binary file),必须用本工具。';

const SHEET_SPECS_SCHEMA = {
  type: 'array',
  description: '工作表数组',
  items: {
    type: 'object',
    properties: {
      name: { type: 'string', description: '工作表名，默认 Sheet1/Sheet2…' },
      rows: { type: 'array', description: '二维数组，每行一个数组', items: { type: 'array', items: {} } },
      header: { type: 'boolean', description: '是否美化首行作为表头' },
      columnWidths: { type: 'array', items: { type: 'number' }, description: '各列宽度（字符数）' },
      rowHeights: { type: 'array', items: { type: 'number' }, description: '各行高度' },
    },
    required: ['rows'],
  },
};

const DOCX_TABLE_PROPS = {
  borders: { type: 'string', enum: ['all', 'none', 'outline', 'three-line'], description: '框线：all=全部框线、none=无框线、outline=仅外框、three-line=三线表（上下粗线 + 表头下细线）' },
  headerShading: { type: 'string', description: '表头底纹 RRGGBB（如 F2F2F2）；传 "none" 去掉底纹' },
  headerBold: { type: 'boolean', description: '表头是否加粗，默认 true' },
  columnWidthMode: { type: 'string', enum: ['auto', 'manual'], description: '列宽：auto=按内容自动分配（谁内容长谁宽，推荐）；manual=用 columnWidths' },
  columnWidths: { type: 'array', items: { type: 'number' }, description: '各列百分比，如 [30,40,30]（按比例理解，不必凑满 100）' },
  align: { type: 'string', enum: ['left', 'center', 'right'], description: '表格整体对齐' },
  cellVerticalAlign: { type: 'string', enum: ['top', 'center', 'bottom'], description: '单元格内容的垂直对齐' },
  rowHeightPt: { type: 'number', description: '每行最小行高（磅），配合自动换行可让长表格更整齐' },
  repeatHeader: { type: 'boolean', description: '表头行跨页重复（长表格翻页后仍能看到表头）' },
  cantSplit: { type: 'boolean', description: '禁止同一行被分页断开（true 时整行保持在同一页）' },
  cellMargins: {
    type: 'object',
    description: '单元格内边距（twip，1pt=20）：如 {"left":108,"right":108,"top":40,"bottom":40}；没给的边用 Word 默认值（左右 108、上下 0）',
    properties: {
      top: { type: 'number', description: '上内边距（twip）' },
      bottom: { type: 'number', description: '下内边距（twip）' },
      left: { type: 'number', description: '左内边距（twip），Word 默认 108' },
      right: { type: 'number', description: '右内边距（twip），Word 默认 108' },
    },
  },
};

const DOCX_STYLE_PROPS = {
  font: { type: 'string', description: '字体（中文与西文都用它；会写进 eastAsia，汉字才会真的用这个字体渲染）' },
  fontAscii: { type: 'string', description: '西文字体（可选；不传则与 font 相同）' },
  sizePt: { type: 'number', description: '字号（磅）。三号=16、四号=14、小四=12、二号=22、小二=18' },
  lineSpacingPt: { type: 'number', description: '行距固定值（磅），公文常用 28.8；与 lineSpacingMultiple 二选一' },
  lineSpacingMultiple: { type: 'number', description: '行距倍数（如 1.5）；与 lineSpacingPt 二选一' },
  firstLineIndentChars: { type: 'number', description: '首行缩进字符数（公文常用 2）；0 表示不缩进' },
  firstLineIndentPt: { type: 'number', description: '首行缩进磅值（与 firstLineIndentChars 二选一）' },
  align: { type: 'string', enum: ['both', 'center', 'left', 'right'], description: '对齐方式，both=两端对齐' },
  spacingBeforePt: { type: 'number', description: '段前距（磅）' },
  spacingAfterPt: { type: 'number', description: '段后距（磅）' },
  bold: { type: 'boolean', description: '是否加粗（改已有文档时用）' },
  color: { type: 'string', description: '字色 RRGGBB，如 FF0000' },
};

const DOCX_OP_LIST_SCHEMA = {
  type: 'array',
  description: '按顺序执行的操作列表',
  items: {
    type: 'object',
    properties: {
      op: { type: 'string', description: '操作类型：replace_text / set_paragraph / set_style / set_numbering / insert_paragraph / delete_paragraph / set_table / delete_table / insert_table_row / delete_table_row / insert_table_column / delete_table_column / merge_table_cells / unmerge_table_cells' },
      find: { type: 'string', description: 'replace_text：要查找的原文（Word 常把一句话拆进多个 run，本工具跨 run 也能匹配）' },
      replace: { type: 'string', description: 'replace_text：替换成什么；不传表示删除' },
      limit: { type: 'integer', description: 'replace_text：最多替换几处，默认不限' },
      paragraph: { type: 'integer', description: '段落序号，从 1 开始（按文档顺序编号，表格里的段落也算）' },
      match: { type: 'string', description: '用整段原文定位段落；必须唯一，否则报错并给出候选序号' },
      text: { type: 'string', description: 'set_paragraph / insert_paragraph：段落的新正文' },
      heading: { type: 'integer', description: 'insert_paragraph：套用 Word 内置标题样式 Heading1–Heading9（插入的段落会出现在标题大纲里）' },
      position: { type: 'string', enum: ['end', 'start', 'before', 'after'], description: 'insert_paragraph：插入位置，默认 end（文末，会自动落在页面设置之前）；before/after 需要配合 paragraph 或 match' },
      onlyEmpty: { type: 'boolean', description: 'delete_table：只删 0 行的空表格（删行删空后的残留，Word 不渲染但 XML 里还在），不用去数它是第几个表' },
      scope: { type: 'string', enum: ['all', 'body', 'headings', 'table'], description: 'set_style 按角色选段落：all=全部、body=正文（不含标题）、headings=标题、table=表格内段落，可与 from/to 区间叠加；表格类操作用 "all" 表示所有表格' },
      from: { type: 'integer', description: 'set_style：起始段落序号（含）' },
      to: { type: 'integer', description: 'set_style：结束段落序号（含）' },
      table: { type: 'integer', description: '表格序号，从 1 开始（文档里只有一个表格时可省略）；表格类操作用它选表' },
      at: { type: 'integer', description: '表格类操作：第几行/列（从 1 开始）。insert_* 表示插在它之前，省略表示追加到末尾；delete_* 表示从这里开始删' },
      count: { type: 'integer', description: '表格类操作：增删几行/几列，默认 1' },
      range: { type: 'string', description: 'merge_table_cells / unmerge_table_cells：矩形区域，如 "A1:B2"（合并保留左上角单元格的内容；取消合并把 gridSpan/vMerge 还原成独立单元格）' },
      style: { type: 'string', enum: ['multicol-1_1_1', 'gongwen-1_1_1_1'], description: 'set_numbering：预置多级编号方案。multicol-1_1_1 = 一级 1、二级 1.1、三级 1.1.1…（西式）；gongwen-1_1_1_1 = 一、/（一）/1./（1）（GB/T 9704 公文层次）' },
      linkToHeading: { type: 'boolean', description: 'set_numbering：是否把 1–9 级编号按样式挂到 Heading1–Heading9（默认 true，标题自动编号）；false 则只给落进选择范围的标题显式编号' },
      exclude: { type: 'array', items: { type: 'string' }, description: 'set_numbering：不参与编号的标题样式，如 ["Heading1"]；这些样式既不编号也不占层级' },
      startFrom: { type: 'object', additionalProperties: { type: 'number' }, description: 'set_numbering：从哪个标题样式开始编号，值是链上第 1 级的起始数字（默认 1），如 {"Heading2": 1}；比它浅的标题样式自动不参与' },
      ...DOCX_STYLE_PROPS,
      ...DOCX_TABLE_PROPS,
    },
    required: ['op'],
  },
};

const XLSX_OP_LIST_SCHEMA = {
  type: 'array',
  description: '按顺序执行的操作列表',
  items: {
    type: 'object',
    properties: {
      op: { type: 'string', description: '操作类型，见工具描述' },
      sheet: { description: '工作表名或从 1 开始的序号，默认 1' },
      ref: { type: 'string', description: '单元格引用，如 B3' },
      range: { type: 'string', description: '区域，如 A1:C10' },
      start: { type: 'string', description: 'set_cells 起始单元格，如 B2' },
      values: { type: 'array', items: { type: 'array', items: {} }, description: 'set_cells 的二维值' },
      value: { description: '单元格值（"=" 开头为公式，date:YYYY-MM-DD 为日期）' },
      formula: { type: 'string', description: '公式（不带等号也可）' },
      style: { type: 'object', additionalProperties: true, description: '样式对象' },
      at: { type: 'integer', description: '插入/删除行(列)的位置' },
      count: { type: 'integer', description: '插入/删除行(列)的数量' },
      name: { type: 'string', description: '工作表名' },
      col: { description: '列号或列字母' },
      width: { type: 'number', description: '列宽' },
      row: { type: 'integer', description: '行号' },
      height: { type: 'number', description: '行高' },
      rows: { type: 'integer', description: 'freeze 冻结行数' },
      cols: { type: 'integer', description: 'freeze 冻结列数' },
      header: { type: 'boolean', description: 'add_sheet 是否美化表头' },
      path: { type: 'string', description: 'add_image 图片路径' },
      cell: { type: 'string', description: 'add_image 锚点单元格' },
      rule: { type: 'object', additionalProperties: true, description: 'conditional_format 的一条规则 {type,...}：cellIs(operator/value/value2/style)、expression(formula)、colorScale(colors 2–3 个 RRGGBB)、dataBar(color)、iconSet(iconSet/showValue)、top10(rank/percent/bottom)、aboveAverage(aboveAverage)、containsText(operator/text)、timePeriod(timePeriod)；data_validation 的一条规则 {type,...}：list(values 候选数组 或 source 区域)、whole/decimal/date/textLength(operator/value/value2)、custom(formula)，可带 allowBlank/promptTitle/prompt/errorTitle/error/errorStyle' },
      rules: { type: 'array', items: { type: 'object', additionalProperties: true }, description: 'conditional_format：同一区域的多条规则（按顺序，优先级依次递增）' },
      chartType: { type: 'string', enum: ['bar', 'column', 'line', 'pie'], description: '图表类型' },
      categories: { type: 'string', description: '图表分类轴范围，如 A2:A6' },
      series: { type: 'array', items: {}, description: '图表数值系列，如 ["B2:B6"] 或 [{range:"B2:B6",label:"B1"}]' },
      title: { type: 'string', description: '图表标题' },
      anchor: { type: 'string', description: '图表左上角单元格，如 G2' },
      optional: { type: 'boolean', description: '该操作失败时跳过而不中断' },
    },
    required: ['op'],
  },
};

// ---------------------------------------------------------------------------
// path resolution / sandbox fence / observation
// ---------------------------------------------------------------------------
function containmentRoots(ctx, exec) {
  const roots = [];
  let mode = 'workspace-write';
  try {
    const policy = ctx.get('sandboxPolicy');
    const resolved = policy?.resolve
      ? policy.resolve({ ...(exec?.agent?.session ? { session: exec.agent.session } : {}) })
      : null;
    if (resolved?.mode) mode = resolved.mode;
    if (resolved?.workspaceRoot) roots.push(resolved.workspaceRoot);
  } catch { /* policy service unavailable */ }
  const sessionCwd = exec?.agent?.session?.header?.cwd;
  if (sessionCwd) roots.push(sessionCwd);
  // DSH's own sandbox-policy default workspaceRoot is process.cwd(); mirror it
  // so a missing/limited policy service still yields a sane writable root.
  roots.push(process.cwd(), tmpdir());
  return { mode, roots: [...new Set(roots.map((r) => resolvePath(String(r))))] };
}

async function resolveTarget(ctx, exec, raw, { forWrite = false, cwd } = {}) {
  if (typeof raw !== 'string' || !raw.trim()) throw new OfficeError('路径不能为空', 'INVALID_ARGS');
  const fsSvc = ctx.get('fs');
  const baseDir = cwd || exec?.agent?.session?.header?.cwd;
  let displayPath;
  let target = null;
  if (fsSvc?.resolve) {
    try {
      target = await fsSvc.resolve(raw, {
        ...(baseDir ? { cwd: baseDir } : {}),
        ...(exec?.signal ? { signal: exec.signal } : {}),
      });
      displayPath = target?.displayPath || raw;
    } catch (err) {
      throw new OfficeError(`无法解析路径 ${raw}: ${err?.message || err}`, 'PATH_RESOLVE_FAILED');
    }
  } else {
    displayPath = isAbsolute(raw) ? raw : resolvePath(baseDir || process.cwd(), raw);
  }
  if (forWrite) assertWritable(ctx, exec, displayPath);
  return { displayPath, target };
}

/**
 * 第二个文件(图片、join 的另一张表)的读取器:路径解析走与主文件同一套沙箱规则,
 * core 只拿到字节,不自己碰文件系统。baseDir 是相对路径的基准(图片按文档所在目录)。
 */
function makeSecondaryReader(ctx, exec, io, baseDir) {
  return async (raw) => {
    const { displayPath } = await resolveTarget(ctx, exec, raw, { cwd: baseDir });
    return io.readBuf(displayPath);
  };
}

/**
 * Resolve symlinks so the fence cannot be bypassed by a link that lives inside
 * an allowed root but points outside it. A not-yet-existing target falls back to
 * resolving its parent directory.
 */
function realizePath(p) {
  try {
    return realpathSync(p);
  } catch {
    const parent = dirname(p);
    try {
      return join(realpathSync(parent), p.slice(parent.length + 1));
    } catch {
      return p;
    }
  }
}

function assertWritable(ctx, exec, displayPath) {
  const { mode, roots } = containmentRoots(ctx, exec);
  if (mode === 'danger-full-access') return;
  // 词法判定 + 真实路径判定都要通过:前者挡住 ../ 越界,后者挡住符号链接绕行
  if (isInsideAny(displayPath, roots) && isInsideAny(realizePath(displayPath), roots.map(realizePath))) return;
  throw new OfficeError(
    `写入被拒绝：${displayPath} 不在会话工作区或临时目录内（当前文件权限模式 ${mode}）。` +
    `请把输出写到工作区路径，或请用户授权更高权限后重试。`,
    'FS_SANDBOX_DENIED'
  );
}

function emitObserved(ctx, exec, target, payload) {
  if (!target) return;
  try { ctx.emit('fs/observed', target, payload, exec); } catch { /* observers optional */ }
}

async function observePresent(ctx, exec, target) {
  if (!target) return;
  try {
    const info = await ctx.get('fs')?.stat?.(target, exec?.signal);
    emitObserved(ctx, exec, target, { kind: 'present', ...(info?.version === undefined ? {} : { version: info.version }) });
  } catch { /* ignore */ }
}

function makeIo(ctx, exec, { displayPath, target }) {
  return {
    async readBuf(p) {
      const fsSvc = ctx.get('fs');
      if (fsSvc?.readBytes && target && p === displayPath) {
        const bytes = await fsSvc.readBytes(target, exec?.signal);
        await observePresent(ctx, exec, target);
        return Buffer.from(bytes);
      }
      const buf = await readFile(p);
      if (p === displayPath) await observePresent(ctx, exec, target);
      return buf;
    },
    async writeBuf(p, data, options) {
      await mkdir(dirname(p), { recursive: true });
      await writeFile(p, data, options);
      if (p === displayPath) await observePresent(ctx, exec, target);
      return p;
    },
    async remove(p) {
      try { await unlink(p); } catch { /* ignore */ }
    },
    async stat(p) {
      try {
        const s = await fsStat(p);
        return { size: s.size, type: s.isDirectory() ? 'dir' : 'file' };
      } catch (err) {
        throw statFailure(err, p, { ctx, exec, displayPath, target });
      }
    },
    tmpFile(ext) {
      return join(tmpdir(), `dsh-office-${randomBytes(6).toString('hex')}${ext}`);
    },
  };
}

// ---------------------------------------------------------------------------
// tool definitions
// ---------------------------------------------------------------------------
function statFailure(err, p, io) {
  if (err?.code !== 'ENOENT') return err;
  if (p === io.displayPath) emitObserved(io.ctx, io.exec, io.target, { kind: 'absent' });
  return new OfficeError(`文件不存在：${p}`, 'NOT_FOUND');
}

function toolDefs(ctx) {
  return [readTool(ctx), queryTool(ctx), writeDocxTool(ctx), editDocxTool(ctx), writeXlsxTool(ctx), editXlsxTool(ctx), fillTemplateTool(ctx), convertTool(ctx)];
}

function readTool(ctx) {
  return {
    name: 'office_read',
    description:
      `读取 Word / PDF / Excel / 表格文件内容（${WORD_EXTS} ${TABLE_EXTS}）。` +
      '用户提到 Word、PDF、Excel、文档、表格、.docx、.pdf、.xlsx 等后缀时一律用本工具。' + BINARY_HINT +
      '解析方式自动选择：.docx 走 mammoth（Markdown 风格正文/表格）；.doc/.rtf/.odt 优先本机转换器' +
      '(macOS textutil / LibreOffice / Word)，不可用时回退纯 JS；.xlsx 走 ExcelJS；' +
      '.xls/.xlsb/.ods/.csv/.tsv 走 SheetJS（纯 JS，全平台可用）；' +
      '.pdf 抽文本层（优先本机 pdftotext / macOS PDFKit，都没有就用随包携带的 pdfjs，全平台可用）—— ' +
      '只给文本与页数，版式/表格不可靠，扫描件（图片型 PDF）读不出文字、需要 OCR。' +
      'Excel 以 TSV 代码块返回，默认只给前若干行；要做统计/筛选/分组/排序，用 office_query 一次算完，' +
      '要看单元格里的公式本体（而不是算出来的值）就传 formulas: "formula" 或 "both"，' +
      '不要靠反复翻页把整张表读进上下文。' +
      '长文档分段读：Word 正文超过 maxChars 时，返回里会写明总字符数与「继续读」该传的 offset，' +
      '下次带上它就接着往下读，不必从头重来；也可以先传 outline: true 拿标题大纲（级别 / 字符偏移 / 标题），再按偏移直接跳到某一章。' +
      '需要核对字体/字号/行距/缩进等排版格式时，对 .docx 传 withFormatting: true。' +
      'Word 的自动编号（多级标题、公文体例）读取时会展开成文字前缀（一、/（一）/1.1），行文规则比对才认得出层级；' +
      '项目符号保持列表结构不塞进正文。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: PATH_DESC },
        sheets: { type: 'array', items: { type: 'string' }, description: '仅 Excel：只读取这些工作表（名称；序号请写成字符串，如 "2"）' },
        range: { type: 'string', description: '仅 Excel：读取范围，如 A1:F50' },
        maxRows: { type: 'integer', description: '仅 Excel：最多读取行数，默认 400' },
        maxCols: { type: 'integer', description: '仅 Excel：最多读取列数，默认 60' },
        maxChars: { type: 'integer', description: 'Word / PDF 正文每段返回的字符上限（分页大小），默认 90000' },
        offset: { type: 'integer', description: '仅 Word / PDF：从正文第几个字符开始返回，默认 0。返回里会给出总字符数与下次该传的 offset，用于分段读长文档；offset 超出总长会明确提示已到末尾' },
        outline: { type: 'boolean', description: '仅 Word / PDF 的 text 模式：只返回标题大纲（Word 是 Markdown 标题，PDF 是页清单；TSV：级别 / 字符偏移 / 标题），可据此传 offset 直接跳读；与 format: "html" 不能同时用' },
        format: { type: 'string', enum: ['text', 'html'], description: 'Word / PDF 输出格式：text=Markdown 风格（默认），html=原始 HTML' },
        formulas: { type: 'string', enum: ['value', 'formula', 'both'], description: '仅表格：公式单元格显示什么。value=计算值（默认）、formula=公式本体（共享公式给出引用地址）、both=公式 → 计算值；xlsx 与 xls/xlsb/ods 都生效，传 true 等同 formula' },
        withFormatting: {
          type: 'boolean',
          description: '仅 .docx：额外返回格式报告 —— 每段的字体(中文/西文)、字号、行距(固定值/倍数)、首行缩进、对齐、样式名，以及页面尺寸与页边距；并给出格式分布(主流值)与偏离主流的段落。用于比对行文规则(如"正文三号仿宋、行距固定值 28.8 磅")。结构化数据在 meta.formatting。',
        },
      },
      required: ['path'],
    },
    timeoutMs: 120000,
    async execute(args, exec) {
      const { displayPath, target } = await resolveTarget(ctx, exec, args.path);
      const io = makeIo(ctx, exec, { displayPath, target });
      const r = await office.opRead(displayPath, args, io);
      return { content: r.content };
    },
  };
}

function queryTool(ctx) {
  return {
    name: 'office_query',
    description:
      `在表格文件内直接算（${TABLE_EXTS}）：全表扫描，只把结论返回，几万行不必读进上下文。` +
      '表格数据的求和/均值/计数/去重/分组/排序/条件筛选，优先用本工具而不是临时写脚本；' +
      '返回的内容是结果表，不是原始数据。不给 groupBy/aggregate 时返回「表结构画像」：每列的' +
      '类型、非空、空值、去重数、最小/最大/求和/均值、最高频取值 —— 一次调用就知道这张表长什么样。' +
      '多表连接（join：另一个文件或同一文件的另一张表，按列等值连接，inner/left/right/full）与透视表（pivot：rows × columns × values，可带合计）也在这里做，不必写脚本；' +
      '窗口函数、统计建模这类才需要另想办法（如 Python）。' +
      '列名取表头行（headerRow，默认第 1 行；表头不在第 1 行就传行号，没有表头传 0）；' +
      `条件算子：eq/ne/gt/gte/lt/lte/contains/startsWith/endsWith/in/notIn/isBlank/notBlank；` +
      '聚合函数：sum/avg/min/max/count/countDistinct（数值列能识别 ¥1,234.00、12.5% 这类写法）。' +
      '多个 where 条件是 AND 关系；结果默认按分组首次出现顺序，可用 orderBy 排序。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: PATH_DESC },
        sheet: { description: '工作表名或从 1 开始的序号，默认第 1 个工作表' },
        headerRow: { type: 'integer', description: '表头行号，默认 1；传 0 表示没有表头（列名自动取 A/B/C…）' },
        where: {
          type: 'array',
          description: '筛选条件（多个条件为 AND），如 [{"col":"地区","op":"eq","value":"华东"},{"col":"金额","op":"gt","value":1000}]',
          items: {
            type: 'object',
            properties: {
              col: { type: 'string', description: '列名' },
              op: { type: 'string', description: '算子，默认 eq' },
              value: { description: '比较值；in/notIn 传数组；isBlank/notBlank 不需要 value' },
            },
            required: ['col'],
          },
        },
        join: {
          description: '先连另一张表再算（可传对象或对象数组，数组时依次连接）。每项 {path?, sheet?, on, type?, suffix?, headerRow?, as?}：path 为另一张表的文件（默认同文件），sheet 选工作表，on 是等值连接键（列名字符串或 {left,right}，数组表示多列），type 取 inner（默认）/left/right/full，suffix 用于右表重名列（默认 _2），as 给该表起名（同时作为列前缀）。连接发生在 where/分组/透视之前；连接键空值不参与匹配，结果会写明匹配与未匹配行数',
          oneOf: [
            { type: 'object', additionalProperties: true },
            { type: 'array', items: { type: 'object', additionalProperties: true } },
          ],
        },
        pivot: {
          type: 'object',
          additionalProperties: true,
          description: '透视表：{rows:["地区"], columns:"产品", values:[{"col":"金额","fn":"sum"}], totals?:true}。rows 为行维度（可多列），columns 为列维度列名（取值自动展开成列，单指标时直接用取值当表头），values 为指标（同 aggregate），totals:true 追加「合计」行与列。与 groupBy/aggregate 二选一；可用 where/orderBy/limit',
        },
        groupBy: { type: 'array', items: { type: 'string' }, description: '按这些列分组，如 ["地区","产品"]；只给 groupBy 时输出每组的行数' },
        aggregate: {
          type: 'array',
          description: '汇总项，如 [{"col":"金额","fn":"sum","as":"销售额"}]；不给 groupBy 时是全表汇总',
          items: {
            type: 'object',
            properties: {
              col: { type: 'string', description: '要汇总的列名' },
              fn: { type: 'string', description: 'sum/avg/min/max/count/countDistinct' },
              as: { type: 'string', description: '结果列名，默认 fn(列名)' },
            },
            required: ['col', 'fn'],
          },
        },
        orderBy: {
          type: 'array',
          description: '排序，如 [{"col":"销售额","dir":"desc"}]（列名可用分组列或汇总结果列）',
          items: {
            type: 'object',
            properties: { col: { type: 'string' }, dir: { type: 'string', enum: ['asc', 'desc'] } },
            required: ['col'],
          },
        },
        limit: { type: 'integer', description: '结果行数上限，默认 200，最大 2000' },
      },
      required: ['path'],
    },
    timeoutMs: 120000,
    async execute(args, exec) {
      const { displayPath, target } = await resolveTarget(ctx, exec, args.path);
      const io = makeIo(ctx, exec, { displayPath, target });
      const r = await office.opQuery(displayPath, args, io);
      return { content: r.content };
    },
  };
}

function writeDocxTool(ctx) {
  return {
    name: 'office_write_docx',
    description:
      '新建 Word 文档（默认 .docx，也支持 .doc / .rtf / .odt）。' +
      '要生成 Word 文档时用本工具，不要用通用 write 工具（写不出合法的 .docx）。' +
      '页眉页脚用 header / footer（页码是 Word 域，如 {"pageNumber":"第 {page} 页 共 {total} 页"}），目录用 toc（插入 TOC 域，Word 打开时自动生成）。' +
      '内容三选一：html / markdown / text，支持标题、段落、粗体斜体、列表、表格、引用、图片等。' +
      '图片写 <img src="图片/a.png" alt="说明"> 或 Markdown 的 ![说明](图片/a.png)：src 取相对输出文件所在目录的路径（也认 ~/ 与绝对路径）或 data: URL，读取走与其它文件同一套沙箱规则，' +
      '只支持 PNG / JPEG / GIF / BMP 且单张不超过 8 MB，按真实像素等比缩放到正文宽度内（可用 width 或 CSS width 指定）；远程 http/https 地址会报错，本插件不联网。' +
      '排版：style 参数设整篇默认（字体 / 字号 / 行距 / 首行缩进 / 对齐 / 段前段后），' +
      '例如公文正文 {"font":"仿宋_GB2312","sizePt":16,"lineSpacingPt":28.8,"firstLineIndentChars":2}，' +
      '标题用 style.headings，表格用 style.table（框线 all/none/outline/three-line、表头底纹、列宽 auto 按内容分配或 manual 指定百分比、表格对齐）。' +
      'html 里的内联样式也认：font-family（写两个逗号分隔时，第一个当中文字体、第二个当西文字体）、' +
      'font-size、line-height（带单位=固定值，纯数字=倍数）、text-indent、margin、text-align、color、background-color，' +
      '且 &lt;body&gt; / &lt;div&gt; 上的样式会被子元素继承 —— 需要逐段不同样式时就写在 html 上。' +
      '.doc/.rtf/.odt 旧格式输出需要本机转换器（macOS textutil / LibreOffice / Microsoft Word），' +
      '不可用时会返回安装提示；图片等复杂元素建议用 .docx。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: PATH_DESC + ' 扩展名决定输出格式：.docx（推荐）/.doc/.rtf/.odt。' },
        html: { type: 'string', description: 'HTML 正文（与 markdown/text 三选一）' },
        markdown: { type: 'string', description: 'Markdown 正文（与 html/text 三选一）' },
        text: { type: 'string', description: '纯文本正文（与 html/markdown 三选一）' },
        title: { type: 'string', description: '文档标题（元数据）' },
        landscape: { type: 'boolean', description: '是否横向纸张（仅 .docx）' },
        marginsMm: { type: 'number', description: '页边距毫米（仅 .docx）' },
        header: {
          type: 'object',
          description: '页眉，如 {"text":"XX单位文件","align":"center","fontSizePt":14}；也可用 pageNumber 放页码',
          properties: {
            text: { type: 'string', description: '页眉文字' },
            align: { type: 'string', enum: ['left', 'center', 'right'], description: '对齐，默认左' },
            bold: { type: 'boolean', description: '是否加粗' },
            fontSizePt: { type: 'number', description: '字号（磅）' },
            pageNumber: { type: 'string', description: '页码模板，用 {page} 表示当前页、{total} 表示总页数' },
          },
        },
        footer: {
          type: 'object',
          description: '页脚，如 {"pageNumber":"第 {page} 页 共 {total} 页","align":"center"}；公文的「— 1 —」可写 {"pageNumber":"— {page} —"}',
          properties: {
            text: { type: 'string', description: '页脚文字（会排在页码模板前面）' },
            align: { type: 'string', enum: ['left', 'center', 'right'], description: '对齐，默认左' },
            bold: { type: 'boolean', description: '是否加粗' },
            fontSizePt: { type: 'number', description: '字号（磅）' },
            pageNumber: { type: 'string', description: '页码模板，用 {page} 表示当前页、{total} 表示总页数；页码是 Word 域，打开时自动更新' },
          },
        },
        toc: {
          type: 'object',
          description: '插入目录（Word 的 TOC 域，打开文档时按当前标题自动生成）。如 {"title":"目 录","levels":3}；只给 true 表示标题「目录」、3 级',
          properties: {
            title: { type: 'string', description: '目录标题，默认「目录」' },
            levels: { type: 'integer', description: '收录几级标题，1–9，默认 3' },
          },
        },
        style: {
          type: 'object',
          description: '整篇排版（字体 / 字号 / 行距 / 首行缩进 / 对齐 / 段前段后 / 标题 / 表格），不传则用内置默认。公文体例示例：{"font":"仿宋_GB2312","sizePt":16,"lineSpacingPt":28.8,"firstLineIndentChars":2,"align":"both"}',
          properties: {
            ...DOCX_STYLE_PROPS,
            headings: { type: 'object', additionalProperties: true, description: '各级标题（Heading1–6）的排版，写法同本对象，如 {"font":"黑体","sizePt":16}' },
            table: { type: 'object', description: '表格样式（框线 / 表头底纹 / 列宽 / 对齐），如 {"borders":"three-line","columnWidthMode":"auto"}', properties: DOCX_TABLE_PROPS },
          },
        },
      },
      required: ['path'],
    },
    timeoutMs: 120000,
    async execute(args, exec) {
      const { displayPath, target } = await resolveTarget(ctx, exec, args.path, { forWrite: true });
      const io = makeIo(ctx, exec, { displayPath, target });
      // <img src> 按输出文档所在目录解析,读取仍走沙箱
      const readImage = makeSecondaryReader(ctx, exec, io, dirname(displayPath));
      const r = await office.opWriteDocx(displayPath, args, io, { readImage });
      return { content: r.content };
    },
  };
}

function editDocxTool(ctx) {
  return {
    name: 'office_edit_docx',
    description:
      '就地修改已有 Word 文档（.docx）：查找替换、整段改写、改排版（字体/字号/行距/首行缩进/对齐）、多级自动编号、插入/删除段落、表格样式与增删行列、合并单元格。' +
      '改已有 Word 文档必须用本工具 —— 通用 edit 工具改不了二进制文档（会报 binary file），' +
      '而 office_write_docx 是整篇重建（原格式、图片、页眉页脚全丢）。' +
      '本工具只重写正文部件，其余部件原样保留，因此段落样式、图片、表格、页眉页脚、批注都不受影响。' +
      '默认就地覆盖；给 outputPath 可另存、给 dryRun:true 只报改动不写盘。' +
      '段落定位二选一：paragraph（段落序号，从 1 开始）或 match（整段原文，必须唯一）；' +
      '序号与内容可用 office_read 先看一眼。操作按顺序执行，后面的序号以执行后的状态为准。' +
      'replace_text 跨 run 匹配（Word 常把一句话拆进多个 run），但替换文字会沿用匹配起点所在 run 的字符格式。' +
      'set_style 改排版：scope 可按角色选 —— "body"（正文，不含标题，例如"把正文改成单倍行距"）、"headings"（标题）、"table"（表格内段落）、"all"（全部），' +
      '还能叠加 from/to 只改某段区间，或用 paragraph/match 改单段；它只覆盖你给出的属性，段落原有的加粗、字号等其它格式保持不动。' +
      '表格用 set_table 改样式（框线 / 表头底纹 / 列宽 / 单元格内边距 / 对齐 / 垂直对齐 / 行高 / 表头跨页重复 / 禁止断行），insert/delete_table_row·column 增删行列，merge_table_cells 合并单元格（保留左上角内容）、unmerge_table_cells 取消合并，' +
      'delete_table 删除整张表格（连 0 行的空表也能删，删完会自动保证正文不以表格结尾；传 onlyEmpty:true 可一次清掉所有 0 行残留，不必数序号）；' +
      '表格用 table 序号选（只有一个表时可省略）或 scope:"all"。' +
      'set_numbering 做多级自动编号：两种预置 —— style:"multicol-1_1_1"（1、1.1、1.1.1）与 style:"gongwen-1_1_1_1"（一、/（一）/1./（1），GB/T 9704 公文层次）；linkToHeading 默认 true —— 把编号按样式挂到 Heading1–9，' +
      '标题无需手写编号文字；同时把「本来就是列表项」的正文段落挂到当前标题层级之下（H1 下为 1.1、H2 下为 1.1.1），普通正文段落不动。编号由 Word 维护，改标题会自动重排。' +
      'exclude 可让某些标题样式不参与编号（如 ["Heading1"]）；startFrom 可指定从哪个标题样式起算、链上第 1 级的起始数字（如 {"Heading2": 1}），比它浅的标题样式自动不参与。' +
      '本操作幂等：重复调用原地替换已有编号定义，不会堆积重复的 abstractNum，段落里的 numId 也保持不变；拿不到正确编号的列表项会跳过并在结果里写明跳过了几段。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: PATH_DESC + ' 必须为 .docx（旧 .doc 先用 office_convert 转 .docx）。' },
        ops: DOCX_OP_LIST_SCHEMA,
        dryRun: { type: 'boolean', description: '只试运行：把将要执行的改动摘要报出来，但不写盘（原文件与 outputPath 都不动）。改动多、把握不准时建议先跑一次' },
        outputPath: { type: 'string', description: '另存到新文件（必须 .docx），原文件保持不变；不传则就地覆盖' },
      },
      required: ['path', 'ops'],
    },
    timeoutMs: 120000,
    async execute(args, exec) {
      const src = await resolveTarget(ctx, exec, args.path, { forWrite: !args.outputPath });
      if (!args.outputPath) {
        const io = makeIo(ctx, exec, { displayPath: src.displayPath, target: src.target });
        const r = await office.opEditDocx(src.displayPath, args.ops, io, { dryRun: args.dryRun });
        return { content: r.content };
      }
      const dst = await resolveTarget(ctx, exec, args.outputPath, { forWrite: true });
      const io = makeIo(ctx, exec, { displayPath: dst.displayPath, target: dst.target });
      const { readFile } = await import('node:fs/promises');
      io.readBuf = async (p) => readFile(p);
      const r = await office.opEditDocx(src.displayPath, args.ops, io, { dryRun: args.dryRun, outputPath: dst.displayPath });
      return { content: r.content };
    },
  };
}

function writeXlsxTool(ctx) {
  return {
    name: 'office_write_xlsx',
    description:
      '新建 Excel 工作簿（.xlsx），可一次写入多个工作表；要生成表格文件时用本工具。' +
      'rows 为二维数组；单元格值规则：数字/布尔按原样，字符串以 "=" 开头视为公式（如 "=SUM(B2:B9)"），' +
      '"date:2026-09-09" 写入日期。header:true 会给首行加粗底纹；columnWidths 可设列宽。' +
      '需要图表时先写数据，再用 office_edit_xlsx 的 add_chart。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: PATH_DESC + ' 必须为 .xlsx。' },
        sheets: SHEET_SPECS_SCHEMA,
      },
      required: ['path', 'sheets'],
    },
    timeoutMs: 180000,
    async execute(args, exec) {
      const { displayPath, target } = await resolveTarget(ctx, exec, args.path, { forWrite: true });
      const io = makeIo(ctx, exec, { displayPath, target });
      const r = await office.opWriteXlsx(displayPath, args, io);
      return { content: r.content };
    },
  };
}

function editXlsxTool(ctx) {
  return {
    name: 'office_edit_xlsx',
    description:
      '修改已有 .xlsx（按顺序执行 ops）。改表格必须用本工具，通用 edit 工具改不了二进制表格。支持操作：' +
      'set_value{sheet,ref,value,style}、set_cells{sheet,start,values,style}、set_formula{sheet,ref,formula}、' +
      'style_range{sheet,range,style}、merge/unmerge{sheet,range}、insert_rows/delete_rows{sheet,at,count}、' +
      'insert_cols/delete_cols{sheet,at,count}、add_sheet{name,rows,header}、rename_sheet{sheet,name}、' +
      'delete_sheet{sheet}、set_col_width{sheet,col,width}、set_row_height{sheet,row,height}、' +
      'freeze{sheet,rows,cols}、auto_filter{sheet,range}、add_image{sheet,path,cell|range}、' +
      'add_chart{sheet,chartType,categories,series,title,anchor}、' +
      'conditional_format{sheet,range,rule|rules}（cellIs / expression / colorScale / dataBar / iconSet / top10 / aboveAverage / containsText / timePeriod）、' +
      'data_validation{sheet,range,rule}（list 下拉 / whole·decimal·date·textLength 区间 / custom 公式）。' +
      'style 支持 {bold,italic,fontSize,color,fill,align,valign,wrap,numFmt,border}（颜色为 RRGGBB）。' +
      'add_chart 的 chartType 支持 bar/column/line/pie；categories 如 "A2:A6"，series 如 ["B2:B6"]（可带 label 单元格）。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: PATH_DESC + ' 必须为 .xlsx（旧 .xls 先用 office_convert 转 .xlsx）。' },
        ops: XLSX_OP_LIST_SCHEMA,
      },
      required: ['path', 'ops'],
    },
    timeoutMs: 180000,
    async execute(args, exec) {
      const { displayPath, target } = await resolveTarget(ctx, exec, args.path, { forWrite: true });
      const io = makeIo(ctx, exec, { displayPath, target });
      const r = await office.opEditXlsx(displayPath, args.ops, io);
      return { content: r.content };
    },
  };
}

function fillTemplateTool(ctx) {
  return {
    name: 'office_fill_docx_template',
    description:
      // 这里不能出现 `{{变量}}` 这类字面量:PTC 模式下工具描述会被嵌进 tools:sdk 提示词
      // 段落并参与变量插值,DSH 会把 {{...}} 当成变量引用而抛错,整个 prompt 组装都会失败。
      '用数据填充 .docx 模板里的占位符（docxtemplater）；模板套打用本工具。' +
      '占位符写法：变量名左右各加两个半角花括号 —— 例如变量名叫 orderNo，模板里就写两个左花括号、orderNo、两个右花括号。' +
      'data 的键即变量名，支持点号嵌套路径（如 a.b）。' +
      '支持数组循环与条件，语法写在模板文件里（不是写在本参数里）：把数组名用「双花括号 + 井号」包起来、' +
      '再用「双花括号 + 斜杠 + 同名」收尾，该段就会按数组逐项重复（可嵌套；传空数组时整块消失）；' +
      '条件用「双花括号 + 井号 + 标记名」表示标记为真时渲染，用「双花括号 + 尖号 + 标记名」表示标记为假时渲染。' +
      '所以「一份合同带 N 行明细」「加急/常规二选一」这类需求一次调用就够，不必逐份调用或写脚本。' +
      '（这些花括号只应出现在模板文件里：写进工具描述或 AGENTS.md 会被 DSH 当成提示词变量而报错。）',
    parameters: {
      type: 'object',
      properties: {
        templatePath: { type: 'string', description: '模板文件路径（.docx，内含上述双花括号占位符）' },
        outputPath: { type: 'string', description: '输出文件路径（.docx）' },
        data: { type: 'object', additionalProperties: true, description: '变量名→值 的对象' },
      },
      required: ['templatePath', 'outputPath', 'data'],
    },
    timeoutMs: 120000,
    async execute(args, exec) {
      const src = await resolveTarget(ctx, exec, args.templatePath);
      const dst = await resolveTarget(ctx, exec, args.outputPath, { forWrite: true });
      const io = makeIo(ctx, exec, { displayPath: src.displayPath, target: src.target });
      const r = await office.opFillTemplate(src.displayPath, dst.displayPath, args.data, io);
      return { content: r.content };
    },
  };
}

function convertTool(ctx) {
  return {
    name: 'office_convert',
    description:
      `格式转换（含旧格式与 PDF）。Word 家族：${WORD_FAMILY} 互转；` +
      `表格家族：${TABLE_FAMILY} 互转（含 .xls 导出）。` +
      '要把文件另存为另一种格式时用本工具，不要用通用读写工具搬运二进制内容。' +
      '目标 .pdf 需要 LibreOffice 或 Microsoft Word；docx 读取走 mammoth，其余走本机转换器或纯 JS 重建。' +
      '不支持 Word↔表格。',
    parameters: {
      type: 'object',
      properties: {
        sourcePath: { type: 'string', description: '源文件路径' },
        outputPath: { type: 'string', description: '目标文件路径（扩展名决定目标格式）' },
      },
      required: ['sourcePath', 'outputPath'],
    },
    timeoutMs: 180000,
    async execute(args, exec) {
      const src = await resolveTarget(ctx, exec, args.sourcePath);
      const dst = await resolveTarget(ctx, exec, args.outputPath, { forWrite: true });
      const io = makeIo(ctx, exec, { displayPath: dst.displayPath, target: dst.target });
      // reads of the source may go through node:fs; keep target bound to the output
      io.readBuf = async (p) => (await readFile(p));
      const readImage = makeSecondaryReader(ctx, exec, io, dirname(src.displayPath));
      const r = await office.opConvert(src.displayPath, dst.displayPath, io, { readImage });
      return { content: r.content };
    },
  };
}

// ---------------------------------------------------------------------------
// plugin apply
// ---------------------------------------------------------------------------
export function apply(ctx) {
  const tools = ctx.tools || ctx.get('tools');
  if (!tools || typeof tools.register !== 'function') {
    ctx.logger?.warn?.('[dsh-office-toolkit] tools 服务不可用，插件未注册');
    return;
  }
  const registerAll = () => {
    const disposers = [];
    for (const def of toolDefs(ctx)) {
      const tool = { ...def, output: OUTPUT };
      const dispose = tools.register(tool);
      disposers.push(dispose);
    }
    ctx.logger?.info?.(`[dsh-office-toolkit] v${PLUGIN_VERSION} 已注册 ${disposers.length} 个 Office 工具`);
    // 依赖装不全时不要等用户撞上报错:加载即自检并给出明确的修复命令
    const missing = missingDeps();
    if (missing.length) {
      const detail = missing.map((spec) => `${spec}(${CORE_DEPS.find(([s]) => s === spec)?.[1] || ''})`).join('、');
      ctx.logger?.warn?.(
        `[dsh-office-toolkit] v${PLUGIN_VERSION} 依赖安装不完整,缺少 ${detail};相关工具会报 MISSING_DEPENDENCY。` +
          `请重跑安装命令补全后重启 dsh web:${INSTALL_COMMAND}`
      );
    }
    return () => { for (const d of disposers) { try { typeof d === 'function' && d(); } catch { /* ignore */ } } };
  };
  if (typeof ctx.effect === 'function') {
    ctx.effect(registerAll, 'dsh-office-toolkit:tools');
  } else {
    registerAll();
  }
}

export { CAPS } from './core/util.js';
