// openpyxl fallback for xlsx writing/editing.
//
// DSH 0.1.6-alpha 的 profile 解析器缺陷会让 ExcelJS 的依赖链在 import 时崩溃
// (见 deps.js),而 ExcelJS 又是唯一的写入后端。openpyxl 是纯 Python、走本机解释器、
// 完全不经过被缺陷命中的解析路径,所以写入/编辑可以在 ExcelJS 不可用时改走它。
//
// 职责切分:值编码、样式与规则校验都留在 JS(与 ExcelJS 后端共用 xlsx-cells /
// xlsx-style / xlsx-rules),子进程只做机械的落盘动作,再读回自校验。
import { execFile, spawn } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { OfficeError, isPlainObject } from './util.js';
import { HEADER_STYLE, cellTag, splitCell } from './xlsx-cells.js';
import { normalizeStyle } from './xlsx-style.js';
import { pythonDataValidation, pythonRules } from './xlsx-rules.js';

const EXECUTOR = fileURLToPath(new URL('./xlsx_py_exec.py', import.meta.url));
const PROBE_TIMEOUT_MS = 10000;
// 可用环境变量调整:大工作簿可放宽超时/输出体积上限(每次调用时读取,便于运行期调整)
const timeoutMs = () => Number(process.env.DSH_OFFICE_PY_TIMEOUT) || 170000;
const maxStdout = () => Number(process.env.DSH_OFFICE_PY_MAX_STDOUT) || 64 * 1024 * 1024;
const MIN_PYTHON = [3, 8];

/**
 * 解释器候选:显式配置优先,其次 PATH 里的 python3/python,最后常见绝对路径。
 * 设 DSH_OFFICE_PYTHON_ONLY=1 时只用显式配置(便于固定环境与测试失败分支)。
 */
function candidates() {
  const configured = String(process.env.DSH_OFFICE_PYTHON || '').trim();
  const fallback = ['python3', 'python', '/opt/homebrew/bin/python3', '/usr/local/bin/python3', '/usr/bin/python3'];
  if (!configured) return fallback;
  return process.env.DSH_OFFICE_PYTHON_ONLY ? [configured] : [configured, ...fallback];
}

/** 探测脚本:报告版本并确认 openpyxl 可导入。 */
const PROBE_SRC = 'import sys, openpyxl; print(sys.version_info[0], sys.version_info[1], openpyxl.__version__)';

function probeWith(exe) {
  return new Promise((resolve) => {
    execFile(exe, ['-c', PROBE_SRC], { timeout: PROBE_TIMEOUT_MS }, (err, stdout) => {
      const [major, minor] = String(stdout || '').trim().split(/\s+/).map(Number);
      if (err) return resolve({ exe, ok: false, reason: '缺少 openpyxl' });
      if (major < MIN_PYTHON[0] || (major === MIN_PYTHON[0] && minor < MIN_PYTHON[1])) {
        return resolve({ exe, ok: false, reason: `Python ${major}.${minor} 版本过低` });
      }
      return resolve({ exe, ok: true, version: `${major}.${minor}` });
    });
  });
}

let pythonChoice = null;

/**
 * 第一个同时满足「可执行 + 带 openpyxl + 版本够新」的解释器。
 * 结果缓存;失败不缓存,便于用户装好 openpyxl 后不重启直接重试。
 */
export async function resolvePython() {
  if (pythonChoice) return pythonChoice;
  const tried = [];
  for (const exe of candidates()) {
    const probe = await probeWith(exe);
    if (probe.ok) {
      pythonChoice = exe;
      return exe;
    }
    tried.push(`${exe}(${probe.reason})`);
  }
  throw new OfficeError(
    'ExcelJS 不可用,需要改用 openpyxl 兜底,但没找到可用的 Python 解释器(要求 3.8+ 且已装 openpyxl)。\n' +
      `已尝试: ${tried.join('、')}\n` +
      '修复: pip3 install openpyxl（或用 DSH_OFFICE_PYTHON 指定解释器路径）',
    'PY_XLSX_UNAVAILABLE'
  );
}

/** 仅供测试:重置解释器缓存。 */
export function resetPythonCache() {
  pythonChoice = null;
}

/** 跑执行器:payload 走 stdin(避免命令行长度上限),stdout 只回一行 JSON。 */
function runExecutor(payload) {
  return resolvePython().then((exe) => new Promise((resolve, reject) => {
    const child = spawn(exe, ['-B', EXECUTOR], { stdio: ['pipe', 'pipe', 'pipe'] });
    const out = [];
    let outBytes = 0;
    let killed = false;
    child.stdout.on('data', (chunk) => {
      outBytes += chunk.length;
      if (outBytes > maxStdout()) {
        killed = true;
        child.kill();
      } else out.push(chunk);
    });
    const timer = setTimeout(() => {
      killed = true;
      child.kill();
    }, timeoutMs());
    let stderr = '';
    child.stderr.on('data', (chunk) => {
      if (stderr.length < 4000) stderr += chunk.toString();
    });
    child.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });
    child.on('close', () => {
      clearTimeout(timer);
      resolve({ text: Buffer.concat(out).toString(), stderr, killed });
    });
    child.stdin.on('error', () => {});
    child.stdin.end(JSON.stringify(payload));
  }));
}

/** 子进程没有给出 JSON 时的现场说明(末几行 stderr 最有信息量)。 */
function childCrashNote(result) {
  const tail = (result.stderr || result.text || '').trim().split('\n').slice(-4).join('；');
  const head = result.killed ? 'openpyxl 子进程异常,已超时/超量' : 'openpyxl 子进程没有返回结果';
  return tail ? `${head}: ${tail}` : head;
}

/** 解析子进程回答;失败时把 Python 侧的说明原样包成可读错误。 */
function unwrap(result, what) {
  let answer = null;
  try {
    answer = JSON.parse(result.text.trim() || '{}');
  } catch {
    answer = null;
  }
  if (answer?.ok) return answer;
  if (answer?.error) throw new OfficeError(`${what}失败: ${answer.error}`, answer.kind === 'verify' ? 'PY_XLSX_VERIFY' : 'PY_XLSX_FAILED');
  throw new OfficeError(`${what}失败(${childCrashNote(result)})`, 'PY_XLSX_FAILED');
}

// ---------------------------------------------------------------------------
// payload 组装
// ---------------------------------------------------------------------------
/** 二维数组 -> 带类型标签的单元格(空值保留 null,长度按最长行补齐)。 */
function tagRows(rows, sheetName) {
  if (!Array.isArray(rows)) throw new OfficeError(`工作表「${sheetName}」的 rows 必须是二维数组`);
  return rows.map((row, index) => {
    if (!Array.isArray(row)) throw new OfficeError(`工作表「${sheetName}」第 ${index + 1} 行不是数组`);
    return row.map((cell) => tagCell(cell));
  });
}

/** 单元格 -> 带类型标签的值(逐格样式一并带上,子进程按格应用)。 */
function tagCell(cell) {
  const { value, style } = splitCell(cell);
  const tag = cellTag(value);
  const normal = styleOf(style);
  return normal ? { ...tag, style: normal } : tag;
}

function styleOf(style) {
  const normal = normalizeStyle(style);
  return Object.keys(normal).length ? normal : null;
}

/** 表名截到 31 字符;重名时追加序号(与 Excel 自动改名一致,openpyxl 会静默改名所以这里自己来)。 */
function uniqueName(name, index, taken) {
  const base = String(name || `Sheet${index + 1}`).slice(0, 31);
  if (!taken.has(base)) return base;
  for (let n = 2; ; n += 1) {
    const suffix = `_${n}`;
    const candidate = base.slice(0, 31 - suffix.length) + suffix;
    if (!taken.has(candidate)) return candidate;
  }
}

function trimmedName(name, index) {
  return String(name || `Sheet${index + 1}`).slice(0, 31);
}

function buildSheetPayload(sheet, index, taken = new Set()) {
  const name = uniqueName(sheet.name, index, taken);
  const header = sheet.header === true;
  const headerStyle = sheet.headerStyle === false ? null : styleOf(sheet.headerStyle || HEADER_STYLE);
  return {
    name,
    header,
    headerStyle,
    rows: tagRows(sheet.rows ?? [], name),
    columnWidths: Array.isArray(sheet.columnWidths) ? sheet.columnWidths : null,
    rowHeights: Array.isArray(sheet.rowHeights) ? sheet.rowHeights : null,
  };
}

/** 写入口:spec -> payload(每个工作表同时算出可用于回读校验的期望形状)。 */
export function buildPayload(spec) {
  if (!spec || !Array.isArray(spec.sheets) || !spec.sheets.length) throw new OfficeError('至少需要一个工作表(sheets)');
  const taken = new Set();
  const sheets = spec.sheets.map((sheet, index) => {
    if (!isPlainObject(sheet)) throw new OfficeError(`sheets[${index}] 必须是对象`);
    const payload = buildSheetPayload(sheet, index, taken);
    taken.add(payload.name);
    return payload;
  });
  const expect = {
    sheets: sheets.map((s) => ({
      name: s.name,
      rows: Math.max(1, s.rows.length),
      cols: s.rows.reduce((m, row) => Math.max(m, row.length), 1),
    })),
  };
  return { mode: 'build', sheets, expect };
}

function numberOr(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function countOf(op) {
  return Math.max(1, numberOr(op.count, 1));
}

function atOf(op) {
  return Math.max(1, numberOr(op.at ?? op.index, 1));
}

/** 单元格写入类操作的值(统一打标签)。 */
function writeCells(op) {
  if (!Array.isArray(op.values)) throw new OfficeError('set_cells 需要 values 二维数组');
  return op.values.map((row) => (Array.isArray(row) ? row.map((cell) => tagCell(cell)) : null));
}

/** 单个 op -> Python 端可机械执行的指令;写过的格子同时登记到 expected 校验清单。 */
const PY_OP_BUILDERS = {
  set_value: (op, ctx) => {
    const ref = String(op.ref || op.cell || 'A1');
    ctx.written.push(ctx.writtenAt(op, ref));
    return { op: 'set_value', sheet: ctx.sheetOf(op), ref, cell: cellTag(op.value), style: styleOf(op.style) };
  },
  set_formula: (op, ctx) => {
    const ref = String(op.ref || op.cell || 'A1');
    ctx.written.push(ctx.writtenAt(op, ref));
    return { op: 'set_formula', sheet: ctx.sheetOf(op), ref, formula: String(op.formula ?? '').replace(/^=/, ''), style: styleOf(op.style) };
  },
  set_cells: (op, ctx) => {
    const start = String(op.start || 'A1');
    const values = writeCells(op);
    ctx.written.push(ctx.writtenAt(op, start, values));
    return { op: 'set_cells', sheet: ctx.sheetOf(op), start, values, style: styleOf(op.style) };
  },
  style_range: (op, ctx) => ({
    op: 'style_range',
    sheet: ctx.sheetOf(op),
    range: String(op.range || op.ref || ''),
    style: styleOf(op.style),
  }),
  merge: (op, ctx) => ({ op: 'merge', sheet: ctx.sheetOf(op), range: String(op.range || '') }),
  unmerge: (op, ctx) => ({ op: 'unmerge', sheet: ctx.sheetOf(op), range: String(op.range || '') }),
  insert_rows: (op, ctx) => ({ op: 'insert_rows', sheet: ctx.sheetOf(op), at: atOf(op), count: countOf(op) }),
  delete_rows: (op, ctx) => ({ op: 'delete_rows', sheet: ctx.sheetOf(op), at: atOf(op), count: countOf(op) }),
  insert_cols: (op, ctx) => ({ op: 'insert_cols', sheet: ctx.sheetOf(op), at: atOf(op), count: countOf(op) }),
  delete_cols: (op, ctx) => ({ op: 'delete_cols', sheet: ctx.sheetOf(op), at: atOf(op), count: countOf(op) }),
  add_sheet: (op, ctx) => {
    const index = ctx.names.length;
    const name = trimmedName(op.name, index);
    if (ctx.names.includes(name)) throw new OfficeError(`工作表「${name}」已存在`);
    const sheet = buildSheetPayload({ ...op, name }, index);
    sheet.name = name;
    if (sheet.rows.length) ctx.written.push({ sheet: name, index: index + 1, start: 'A1', rows: sheet.rows });
    ctx.names.push(name);
    return { op: 'add_sheet', sheet };
  },
  rename_sheet: (op, ctx) => {
    const sheet = ctx.sheetOf(op);
    const from = ctx.nameOf(sheet);
    const next = String(op.name).slice(0, 31);
    ctx.rename(from, next);
    return { op: 'rename_sheet', sheet, name: next };
  },
  delete_sheet: (op, ctx) => {
    const sheet = ctx.sheetOf(op);
    if (ctx.names.length <= 1) throw new OfficeError('不能删除最后一个工作表');
    const at = ctx.names.indexOf(ctx.nameOf(sheet));
    ctx.names.splice(at === -1 ? 0 : at, 1);
    return { op: 'delete_sheet', sheet };
  },
  set_col_width: (op, ctx) => ({
    op: 'set_col_width',
    sheet: ctx.sheetOf(op),
    col: typeof op.col === 'number' ? op.col : String(op.col),
    width: numberOr(op.width, 8.43),
  }),
  set_row_height: (op, ctx) => ({ op: 'set_row_height', sheet: ctx.sheetOf(op), row: Math.max(1, numberOr(op.row, 1)), height: numberOr(op.height, 15) }),
  freeze: (op, ctx) => ({ op: 'freeze', sheet: ctx.sheetOf(op), rows: numberOr(op.rows, 0), cols: numberOr(op.cols, 0) }),
  auto_filter: (op, ctx) => ({ op: 'auto_filter', sheet: ctx.sheetOf(op), range: String(op.range || '') }),
  add_image: (op, ctx) => ({
    op: 'add_image',
    sheet: ctx.sheetOf(op),
    path: resolve(String(op.path || '')),
    cell: op.cell === undefined ? null : String(op.cell),
    range: op.range === undefined ? null : String(op.range),
    width: numberOr(op.width, 0),
    height: numberOr(op.width ? op.height ?? op.width : 0, 0),
  }),
  conditional_format: (op, ctx) => ({ op: 'conditional_format', sheet: ctx.sheetOf(op), ...pythonRules(op) }),
  data_validation: (op, ctx) => ({ op: 'data_validation', sheet: ctx.sheetOf(op), ...pythonDataValidation(op) }),
};

const SUPPORTED_OPS = Object.keys(PY_OP_BUILDERS);

/**
 * 编辑入口:ops -> payload。图表不在这里(它只需成品包,由 office.js 用 injectChart 收尾)。
 * @param {string[]} sheetNames 现有工作表名(按顺序),用于校验序号/名字与改名联动
 */
export function editPayload(ops, sheetNames) {
  if (!Array.isArray(ops) || !ops.length) throw new OfficeError('ops 至少需要一个操作');
  const ctx = editContext([...sheetNames]);
  const out = [];
  const changes = [];
  ops.forEach((op, index) => {
    if (!isPlainObject(op)) throw new OfficeError('每个 op 必须是对象');
    const type = String(op.op || op.type || '');
    if (type === 'add_chart') return;
    const shape = PY_OP_BUILDERS[type];
    if (!shape) throw new OfficeError(`未知操作类型: ${type}（支持 ${SUPPORTED_OPS.join('/')}/add_chart）`);
    const changeIndex = changes.length;
    try {
      out.push({ ...shape(op, ctx), optional: op.optional === true, index, changeIndex });
      changes.push({ index: changeIndex, text: type });
    } catch (err) {
      if (!op.optional) throw err;
      changes.push({ index: changeIndex, text: `${type}(跳过: ${err.message})` });
    }
  });
  return { mode: 'edit', ops: out, expect: { sheets: ctx.names, written: ctx.written }, changes };
}

function editContext(names) {
  const ctx = {
    names,
    written: [],
    sheetOf: (op) => op.sheet ?? 1,
    nameOf(selector) {
      return typeof selector === 'number' ? names[selector - 1] : selector;
    },
    /** 登记待校验的写入点:名字快照 + 当时序号,改名后仍能定位。 */
    writtenAt(op, ref, rows) {
      const name = this.nameOf(this.sheetOf(op));
      const entry = { sheet: name, index: this.names.indexOf(name) + 1, ref };
      if (rows) { delete entry.ref; entry.start = ref; entry.rows = rows; }
      return entry;
    },
    rename(from, to) {
      const at = names.indexOf(from);
      if (at !== -1) names[at] = to;
    },
  };
  return ctx;
}

// ---------------------------------------------------------------------------
// 执行
// ---------------------------------------------------------------------------
async function withTempDir(run) {
  const dir = await mkdtemp(join(tmpdir(), 'dsh-xlsx-'));
  try {
    return await run(dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

/**
 * 生成新工作簿。返回 buf 与读回校验得到的真实工作表形状。
 * @param spec 工具入参
 * @param expect 可选的期望形状(测试里用来验证自校验会拦住不合格产物)
 */
export async function buildWorkbook(spec, expect = null) {
  const payload = buildPayload(spec);
  if (expect) payload.expect = expect;
  return withTempDir(async (dir) => {
    payload.output = join(dir, 'workbook.xlsx');
    const answer = unwrap(await runExecutor(payload), '用 openpyxl 生成 .xlsx');
    return { buf: await readFile(answer.output), sheets: answer.sheets };
  });
}

/** 就地把 ops 应用到现有工作簿,返回新 buf 与执行记录。 */
export async function editWorkbook(buf, ops, sheetNames) {
  const payload = editPayload(ops, sheetNames);
  return withTempDir(async (dir) => {
    payload.input = join(dir, 'input.xlsx');
    payload.output = join(dir, 'output.xlsx');
    await writeFile(payload.input, buf);
    const answer = unwrap(await runExecutor(payload), '用 openpyxl 修改 .xlsx');
    return {
      buf: await readFile(answer.output),
      changes: mergeChanges(payload.changes, answer.changes),
      sheets: answer.sheets,
      droppedCharts: answer.droppedCharts === true,
    };
  });
}

/** JS 侧跳过记录 + 子进程实际结果,按原始顺序拼成一篇执行记录。 */
function mergeChanges(planned, executed) {
  const done = new Map((executed || []).map((item) => [item.index, item.text]));
  return planned.map((item) => done.get(item.index) ?? item.text);
}

/** 结果摘要:与 excel.js#workbookSummary 的 meta 形状对齐(`名称(行×列)`)。 */
export function summaryOf(sheets) {
  return { sheets: (sheets || []).map((s) => ({ name: s.name, rows: s.rows, cols: s.cols })) };
}
