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
const WORD_EXTS = '.docx .doc .rtf .odt';
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

async function resolveTarget(ctx, exec, raw, { forWrite = false } = {}) {
  if (typeof raw !== 'string' || !raw.trim()) throw new OfficeError('路径不能为空', 'INVALID_ARGS');
  const fsSvc = ctx.get('fs');
  const sessionCwd = exec?.agent?.session?.header?.cwd;
  let displayPath;
  let target = null;
  if (fsSvc?.resolve) {
    try {
      target = await fsSvc.resolve(raw, {
        ...(sessionCwd ? { cwd: sessionCwd } : {}),
        ...(exec?.signal ? { signal: exec.signal } : {}),
      });
      displayPath = target?.displayPath || raw;
    } catch (err) {
      throw new OfficeError(`无法解析路径 ${raw}: ${err?.message || err}`, 'PATH_RESOLVE_FAILED');
    }
  } else {
    displayPath = isAbsolute(raw) ? raw : resolvePath(sessionCwd || process.cwd(), raw);
  }
  if (forWrite) assertWritable(ctx, exec, displayPath);
  return { displayPath, target };
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
  return [readTool(ctx), queryTool(ctx), writeDocxTool(ctx), writeXlsxTool(ctx), editXlsxTool(ctx), fillTemplateTool(ctx), convertTool(ctx)];
}

function readTool(ctx) {
  return {
    name: 'office_read',
    description:
      `读取 Word / Excel / 表格文件内容（${WORD_EXTS} ${TABLE_EXTS}）。` +
      '用户提到 Word、Excel、文档、表格、.docx、.xlsx 等后缀时一律用本工具。' + BINARY_HINT +
      '解析方式自动选择：.docx 走 mammoth（Markdown 风格正文/表格）；.doc/.rtf/.odt 优先本机转换器' +
      '(macOS textutil / LibreOffice / Word)，不可用时回退纯 JS；.xlsx 走 ExcelJS；' +
      '.xls/.xlsb/.ods/.csv/.tsv 走 SheetJS（纯 JS，全平台可用）。' +
      'Excel 以 TSV 代码块返回，默认只给前若干行；要做统计/筛选/分组/排序，用 office_query 一次算完，' +
      '不要靠反复翻页把整张表读进上下文。' +
      '需要核对字体/字号/行距/缩进等排版格式时，对 .docx 传 withFormatting: true。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: PATH_DESC },
        sheets: { type: 'array', items: { type: 'string' }, description: '仅 Excel：只读取这些工作表（名称；序号请写成字符串，如 "2"）' },
        range: { type: 'string', description: '仅 Excel：读取范围，如 A1:F50' },
        maxRows: { type: 'integer', description: '仅 Excel：最多读取行数，默认 400' },
        maxCols: { type: 'integer', description: '仅 Excel：最多读取列数，默认 60' },
        maxChars: { type: 'integer', description: 'Word 正文返回字符上限，默认 90000' },
        format: { type: 'string', enum: ['text', 'html'], description: 'Word 输出格式：text=Markdown 风格（默认），html=原始 HTML' },
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
      '本工具覆盖范围之外的需求（跨文件 join、透视表、窗口函数、图表、统计建模）再另想办法（如 Python）。' +
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
      '内容三选一：html / markdown / text，支持标题、段落、粗体斜体、列表、表格、引用等。' +
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
      },
      required: ['path'],
    },
    timeoutMs: 120000,
    async execute(args, exec) {
      const { displayPath, target } = await resolveTarget(ctx, exec, args.path, { forWrite: true });
      const io = makeIo(ctx, exec, { displayPath, target });
      const r = await office.opWriteDocx(displayPath, args, io);
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
      'add_chart{sheet,chartType,categories,series,title,anchor}。' +
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
      'data 的键即变量名，支持点号嵌套路径（如 a.b）。不支持循环/条件语法；批量套打时多次调用即可。',
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
      const r = await office.opConvert(src.displayPath, dst.displayPath, io);
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
