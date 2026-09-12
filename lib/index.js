// dsh-office-toolkit — DeepSeek Harness host plugin (host-only, no client half).
// Registers Word/Excel tools on the agent tool registry. Binary IO goes through
// node:fs after ctx.fs resolution + a containment fence; reads use ctx.fs.readBytes
// when available, and fs/observed events are emitted so the session's file
// observation state (and therefore the official read/write/edit guards) stay coherent.
import { dirname, isAbsolute, join, resolve as resolvePath } from 'node:path';
import { mkdir, readFile, stat as fsStat, unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { randomBytes } from 'node:crypto';
import * as office from './core/office.js';
import { OfficeError } from './core/util.js';
import { isInsideAny } from './core/path-guard.js';

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

function assertWritable(ctx, exec, displayPath) {
  const { mode, roots } = containmentRoots(ctx, exec);
  if (mode === 'danger-full-access') return;
  if (isInsideAny(displayPath, roots)) return;
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
    async writeBuf(p, data) {
      await mkdir(dirname(p), { recursive: true });
      await writeFile(p, data);
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
  return [readTool(ctx), writeDocxTool(ctx), writeXlsxTool(ctx), editXlsxTool(ctx), fillTemplateTool(ctx), convertTool(ctx)];
}

function readTool(ctx) {
  return {
    name: 'office_read',
    description:
      '读取 Word/Excel 文件内容。按扩展名自动选择解析器：.docx 走 mammoth（返回 Markdown 风格正文/表格），' +
      '.doc/.rtf/.odt 优先用本机转换器(macOS textutil / LibreOffice / Word)，不可用时回退纯 JS 解析，' +
      '.xlsx 走 ExcelJS，.xls/.xlsb/.ods/.csv/.tsv 走 SheetJS（纯 JS，全平台可用）。' +
      'Excel 以 TSV 代码块返回，可用 sheets/range/maxRows 参数分段读取大表。',
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

function writeDocxTool(ctx) {
  return {
    name: 'office_write_docx',
    description:
      '新建 Word 文档（默认 .docx）。内容三选一：html / markdown / text。' +
      '支持标题、段落、粗体斜体、列表、表格、引用等。' +
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
      '新建 Excel 工作簿（.xlsx），可一次写入多个工作表。rows 为二维数组；单元格值规则：' +
      '数字/布尔按原样，字符串以 "=" 开头视为公式（如 "=SUM(B2:B9)"），"date:2026-09-09" 写入日期。' +
      'header:true 会给首行加粗底纹；columnWidths 可设列宽。需要图表时先写数据，再用 office_edit_xlsx 的 add_chart。',
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
      '修改已有 .xlsx（按顺序执行 ops）。支持操作：' +
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
      '用数据填充 .docx 模板中的 {{变量}} 占位符（docxtemplater）。data 的键即变量名，支持嵌套如 {{a.b}}。' +
      '不支持循环/条件语法；批量套打时多次调用即可。',
    parameters: {
      type: 'object',
      properties: {
        templatePath: { type: 'string', description: '模板文件路径（.docx，内含 {{变量}}）' },
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
      '格式转换。Word 家族：doc/docx/rtf/odt/html/txt/md 之间互转（docx 读取走 mammoth，其余走本机转换器，' +
      '不可用时用纯 JS 文本重建）；目标 .pdf 需要 LibreOffice 或 Microsoft Word。' +
      '表格家族：xlsx/xls/xlsb/ods/csv/tsv/html 之间互转（SheetJS，纯 JS，含 .xls 导出）。' +
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
    ctx.logger?.info?.(`[dsh-office-toolkit] 已注册 ${disposers.length} 个 Office 工具`);
    return () => { for (const d of disposers) { try { typeof d === 'function' && d(); } catch { /* ignore */ } } };
  };
  if (typeof ctx.effect === 'function') {
    ctx.effect(registerAll, 'dsh-office-toolkit:tools');
  } else {
    registerAll();
  }
}

export { CAPS } from './core/util.js';
