// NOSONAR: 本文件用自建的 t() 断言跑用例,不依赖 node:test,扫描器认不出是测试文件
// openpyxl 兜底后端测试(ExcelJS 因 DSH 解析缺陷不可用时走这条路)。
// Usage: node test/xlsx-py.test.mjs   (输出进 test/out-py/)
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as office from '../lib/core/office.js';
import * as excel from '../lib/core/excel.js';
import * as py from '../lib/core/xlsx-py.js';
import { readWorkbook } from '../lib/core/excel.js';
import { assertOfficeBinary } from '../lib/core/util.js';
import { test } from 'node:test';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, 'out-py');
await mkdir(outDir, { recursive: true });

// 让扫描器把本文件认成测试文件(S2187):注册一条不依赖外部环境的冒烟用例
test('openpyxl 兜底模块可加载', async () => {
  const mod = await import('../lib/core/xlsx-py.js');
  if (typeof mod.buildWorkbook !== 'function') throw new Error('缺少 buildWorkbook');
});

const results = [];
async function t(name, fn) {
  try {
    const detail = await fn();
    results.push({ name, ok: true, detail: detail || '' });
    console.log(`✓ ${name}`);
  } catch (err) {
    results.push({ name, ok: false, detail: err.message });
    console.log(`✗ ${name}\n   ${err.message}`);
  }
}

const resolverBug = () => Object.assign(
  new Error('createRequire.resolve.paths is not a function or its return value is not iterable'),
  { code: 'DSH_RESOLVER_BUG' }
);

/** openpyxl 重新打开产物,供断言落盘内容。 */
async function pyOpen(path) {
  const { execFile } = await import('node:child_process');
  const script = [
    'import json,sys,openpyxl',
    'wb=openpyxl.load_workbook(sys.argv[1])',
    'out={"sheets":wb.sheetnames}',
    'ws=wb[wb.sheetnames[0]]',
    'out["cells"]={c.coordinate:c.value for row in ws.iter_rows() for c in row if c.value is not None}',
    'out["freeze"]=ws.freeze_panes',
    'out["filter"]=ws.auto_filter.ref',
    'out["cf"]=sum(len(cf.rules) for cf in ws.conditional_formatting)',
    'out["dv"]=len(ws.data_validations.dataValidation)',
    'out["bold"]=[c.coordinate for row in ws.iter_rows() for c in row if c.font and c.font.b]',
    'out["widthB"]=ws.column_dimensions["B"].width',
    'print(json.dumps(out, default=str))',
  ].join('\n');
  const exe = await py.resolvePython();
  const stdout = await new Promise((resolve, reject) => {
    execFile(exe, ['-c', script, path], (err, out, stderr) => {
      if (err) reject(new Error(stderr || err.message));
      else resolve(out);
    });
  });
  return JSON.parse(stdout);
}

function sameList(actual, expected, label) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) throw new Error(`${label} 不符: ${a} != ${e}`);
}

// 没有可用解释器(如精简 CI 镜像)时整套跳过,而不是 24 条一起误报失败
let pyReady = false;
try {
  await py.resolvePython();
  pyReady = true;
} catch (err) {
  console.log(`跳过: 本机没有装 openpyxl 的 Python(${err.message.split('\n')[0]})`);
}

if (pyReady) {
// ---------------------------------------------------------------------------
// 解释器探测
// ---------------------------------------------------------------------------
await t('探测: 找到的解释器能 import openpyxl', async () => {
  const exe = await py.resolvePython();
  if (!exe) throw new Error('未返回解释器');
  return exe;
});

await t('探测: 指定不存在的解释器时给出安装提示而不是崩溃', async () => {
  const old = process.env.DSH_OFFICE_PYTHON;
  process.env.DSH_OFFICE_PYTHON = '/definitely/not/python';
  process.env.DSH_OFFICE_PYTHON_ONLY = '1';
  py.resetPythonCache();
  try {
    await py.resolvePython();
    throw new Error('不存在的解释器却解析成功');
  } catch (err) {
    if (err.code !== 'PY_XLSX_UNAVAILABLE') throw err;
    if (!err.message.includes('pip3 install openpyxl')) throw new Error('缺少安装提示: ' + err.message);
    if (!err.message.includes('/definitely/not/python')) throw new Error('未列出尝试过的解释器');
  } finally {
    if (old === undefined) delete process.env.DSH_OFFICE_PYTHON;
    else process.env.DSH_OFFICE_PYTHON = old;
    delete process.env.DSH_OFFICE_PYTHON_ONLY;
    py.resetPythonCache();
  }
  return '报 PY_XLSX_UNAVAILABLE 并给出 pip3 install openpyxl';
});

// ---------------------------------------------------------------------------
// 写入
// ---------------------------------------------------------------------------
await t('写入: 多工作表 + 表头底纹 + 列宽 + 公式/日期/num 前缀都按类型落盘', async () => {
  const path = join(outDir, 'built.xlsx');
  const r = await py.buildWorkbook({
    sheets: [
      {
        name: '订单',
        header: true,
        columnWidths: [18, 10],
        rows: [['产品', '金额'], ['键盘', 'num:1,200'], ['鼠标', 'date:2026-10-23'], ['合计', '=B2+B3']],
      },
      { name: '备注', rows: [['说明'], ['测试']] },
    ],
  });
  assertOfficeBinary(r.buf, 'xlsx');
  if (r.sheets.length !== 2) throw new Error('工作表数不对: ' + r.sheets.length);
  await writeFile(path, r.buf);
  const seen = await pyOpen(path);
  sameList(seen.sheets, ['订单', '备注'], '工作表名');
  if (seen.cells.A2 !== '键盘') throw new Error('文本没写入: ' + seen.cells.A2);
  if (seen.cells.B2 !== 1200) throw new Error('num 前缀未转数字: ' + seen.cells.B2);
  if (!/2026-10-23/.test(String(seen.cells.B3))) throw new Error('date 前缀未转日期: ' + seen.cells.B3);
  if (String(seen.cells.B4) !== '=B2+B3') throw new Error('公式没写入: ' + seen.cells.B4);
  if (seen.cells.A1 !== '产品') throw new Error('表头文字丢失');
  if (!seen.bold.includes('A1') || !seen.bold.includes('B1')) throw new Error('表头未加粗: ' + seen.bold);
  if (seen.widthB !== 10) throw new Error('列宽未生效: ' + seen.widthB);
  return '值类型/表头样式/列宽全部落盘';
});

await t('写入: 空工作表、超长表名截断、缺省名补 SheetN', async () => {
  const r = await py.buildWorkbook({ sheets: [{ rows: [] }, { name: 'x'.repeat(40), rows: [['a']] }] });
  const path = join(outDir, 'names.xlsx');
  await writeFile(path, r.buf);
  const seen = await pyOpen(path);
  if (seen.sheets[0] !== 'Sheet1') throw new Error('缺省名不对: ' + seen.sheets[0]);
  if (seen.sheets[1].length !== 31) throw new Error('表名未截断到 31: ' + seen.sheets[1].length);
  return '缺省名 Sheet1,表名截断 31 字符';
});

await t('写入: 同名工作表自动去重(openpyxl 会静默改名,不自己处理会误判校验失败)', async () => {
  const r = await py.buildWorkbook({ sheets: [{ name: 'S', rows: [['a']] }, { name: 'S', rows: [['b']] }] });
  sameList(r.sheets.map((s) => s.name), ['S', 'S_2'], '去重后的表名');
  return '重名改成 S_2';
});

await t('写入: 自校验能拦住不合格产物(期望行数比实际多)', async () => {
  try {
    await py.buildWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] }, { sheets: [{ name: 'S', rows: 9, cols: 3 }] });
    throw new Error('期望不符却没报错');
  } catch (err) {
    if (err.code !== 'PY_XLSX_VERIFY') throw new Error(`错误码不对(${err.code}): ${err.message}`);
    if (!err.message.includes('落盘不全')) throw new Error('提示不明确: ' + err.message);
  }
  return 'PY_XLSX_VERIFY + 明确说明哪张表不完整';
});

await t('写入: 微妙权限下缺 openpyxl 的解释器给出可读提示', async () => {
  const old = process.env.DSH_OFFICE_PYTHON;
  process.env.DSH_OFFICE_PYTHON = '/usr/bin/python3';
  py.resetPythonCache();
  try {
    await py.buildWorkbook({ sheets: [{ rows: [['a']] }] });
    return '本机 /usr/bin/python3 也带 openpyxl';
  } catch (err) {
    if (err.code !== 'PY_XLSX_UNAVAILABLE') throw new Error(`错误码不对(${err.code}): ${err.message}`);
    if (!err.message.includes('缺少 openpyxl')) throw new Error('未说明缺 openpyxl: ' + err.message);
    return '报 PY_XLSX_UNAVAILABLE 并点明缺 openpyxl';
  } finally {
    if (old === undefined) delete process.env.DSH_OFFICE_PYTHON;
    else process.env.DSH_OFFICE_PYTHON = old;
    py.resetPythonCache();
  }
});

// ---------------------------------------------------------------------------
// 编辑:全部操作
// ---------------------------------------------------------------------------
await t('编辑: 全部 op 执行成功且逐项落盘', async () => {
  const { buf } = await py.buildWorkbook({ sheets: [{ name: '数据', header: true, rows: [['项目', '金额'], ['甲', 1], ['乙', 2]] }] });
  const ops = [
    { op: 'set_value', sheet: 1, ref: 'D1', value: '备注', style: { bold: true } },
    { op: 'set_cells', sheet: 1, start: 'A5', values: [['小计', '=SUM(B2:B3)']] },
    { op: 'set_formula', sheet: 1, ref: 'B6', formula: '=SUM(B2:B3)' },
    { op: 'style_range', sheet: 1, range: 'A1:B1', style: { italic: true, align: 'center' } },
    { op: 'merge', sheet: 1, range: 'A8:B8' },
    { op: 'unmerge', sheet: 1, range: 'A8:B8' },
    { op: 'insert_rows', sheet: 1, at: 4, count: 1 },
    { op: 'delete_rows', sheet: 1, at: 4, count: 1 },
    { op: 'insert_cols', sheet: 1, at: 4, count: 1 },
    { op: 'delete_cols', sheet: 1, at: 4, count: 1 },
    { op: 'add_sheet', name: '备注', rows: [['说明'], ['测试']], header: true },
    { op: 'rename_sheet', sheet: '备注', name: '备注2' },
    { op: 'set_col_width', sheet: 1, col: 'B', width: 18 },
    { op: 'set_row_height', sheet: 1, row: 1, height: 28 },
    { op: 'freeze', sheet: 1, rows: 1, cols: 1 },
    { op: 'auto_filter', sheet: 1, range: 'A1:B3' },
    { op: 'conditional_format', sheet: 1, range: 'B2:B3', rule: { type: 'colorScale', colors: ['FFFFFF', 'FF0000'] } },
    { op: 'conditional_format', sheet: 1, range: 'B2:B3', rules: [{ type: 'top10', rank: 2, style: { bold: true } }, { type: 'containsText', operator: 'beginsWith', text: '甲' }] },
    { op: 'data_validation', sheet: 1, range: 'C2:C10', rule: { type: 'list', values: ['甲', '乙'], error: '请选择' } },
    { op: 'delete_sheet', sheet: '备注2' },
  ];
  const r = await py.editWorkbook(buf, ops, ['数据']);
  const path = join(outDir, 'edited.xlsx');
  await writeFile(path, r.buf);
  if (r.changes.length !== ops.length) throw new Error('执行记录不完整: ' + r.changes.join('|'));
  const seen = await pyOpen(path);
  sameList(seen.sheets, ['数据'], '删表后剩余工作表');
  if (seen.cells.D1 !== '备注') throw new Error('set_value 没落盘');
  if (String(seen.cells.B6) !== '=SUM(B2:B3)') throw new Error('set_formula 没落盘');
  if (seen.cells.A5 !== '小计') throw new Error('set_cells 没落盘');
  if (seen.freeze !== 'B2') throw new Error('冻结窗格没落盘: ' + seen.freeze);
  if (seen.filter !== 'A1:B3') throw new Error('筛选没落盘: ' + seen.filter);
  if (seen.cf < 2) throw new Error('条件格式没落盘: ' + seen.cf);
  if (seen.dv < 1) throw new Error('数据验证没落盘: ' + seen.dv);
  if (seen.widthB !== 18) throw new Error('列宽没落盘: ' + seen.widthB);
  return `${ops.length} 个 op 全部生效`;
});

await t('编辑: optional 的操作失败被跳过,不影响其余操作', async () => {
  const { buf } = await py.buildWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] });
  const r = await py.editWorkbook(buf, [
    { op: 'add_image', sheet: 1, cell: 'B2', path: join(outDir, 'missing.png'), optional: true },
    { op: 'set_value', sheet: 1, ref: 'C3', value: '仍执行' },
  ], ['S']);
  if (!r.changes[0].includes('跳过')) throw new Error('未记录跳过: ' + r.changes[0]);
  if (r.changes[1] !== 'set_value') throw new Error('后续操作记录不对: ' + r.changes[1]);
  return r.changes[0];
});

await t('编辑: 非 optional 的失败报 PY_XLSX_FAILED 且说明原因', async () => {
  const { buf } = await py.buildWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] });
  try {
    await py.editWorkbook(buf, [{ op: 'set_value', sheet: '不存在的表', ref: 'A1', value: 1 }], ['S']);
    throw new Error('不存在的工作表却没报错');
  } catch (err) {
    if (err.code !== 'PY_XLSX_FAILED') throw new Error(`错误码不对(${err.code}): ${err.message}`);
    if (!err.message.includes('不存在')) throw new Error('原因不明确: ' + err.message);
  }
  return 'PY_XLSX_FAILED + 工作表不存在';
});

await t('编辑: 未知 op 在 JS 侧就被拦下(不进子进程)', async () => {
  const { buf } = await py.buildWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] });
  try {
    await py.editWorkbook(buf, [{ op: 'no_such_op', sheet: 1 }], ['S']);
    throw new Error('未知 op 却没报错');
  } catch (err) {
    if (!err.message.includes('未知操作类型')) throw err;
  }
  return 'JS 侧校验未知操作类型';
});

// ---------------------------------------------------------------------------
// 子进程异常现场(用假解释器注入)
// ---------------------------------------------------------------------------
const FAKE = join(here, 'fake-python.sh');

/** 用假解释器跑一次,返回错误对象。 */
async function withFake(opts, fn) {
  const { mode, timeout, maxStdout } = opts;
  const old = { py: process.env.DSH_OFFICE_PYTHON, only: process.env.DSH_OFFICE_PYTHON_ONLY, mode: process.env.FAKE_MODE, t: process.env.DSH_OFFICE_PY_TIMEOUT, m: process.env.DSH_OFFICE_PY_MAX_STDOUT };
  process.env.DSH_OFFICE_PYTHON = FAKE;
  process.env.DSH_OFFICE_PYTHON_ONLY = '1';
  process.env.FAKE_MODE = mode;
  if (timeout) process.env.DSH_OFFICE_PY_TIMEOUT = String(timeout);
  if (maxStdout) process.env.DSH_OFFICE_PY_MAX_STDOUT = String(maxStdout);
  py.resetPythonCache();
  try {
    return await fn();
  } finally {
    const put = (key, value) => (value === undefined ? delete process.env[key] : (process.env[key] = value));
    put('DSH_OFFICE_PYTHON', old.py);
    put('DSH_OFFICE_PYTHON_ONLY', old.only);
    put('FAKE_MODE', old.mode);
    put('DSH_OFFICE_PY_TIMEOUT', old.t);
    put('DSH_OFFICE_PY_MAX_STDOUT', old.m);
    py.resetPythonCache();
  }
}

await t('子进程: 垃圾输出(非 JSON)时报 PY_XLSX_FAILED 并附 stderr 尾部', async () => {
  await withFake({ mode: 'garbage' }, async () => {
    try {
      await py.buildWorkbook({ sheets: [{ rows: [['a']] }] });
      throw new Error('非 JSON 输出却没报错');
    } catch (err) {
      if (err.code !== 'PY_XLSX_FAILED') throw new Error('错误码不对: ' + err.code);
      if (!err.message.includes('没有返回结果')) throw new Error('缺少现场说明: ' + err.message);
      if (!err.message.includes('Traceback')) throw new Error('未附 stderr 尾部: ' + err.message);
    }
  });
  return '报 PY_XLSX_FAILED 并带 stderr 尾部';
});

await t('子进程: 输出超量被杀时说明已超量', async () => {
  await withFake({ mode: 'big', maxStdout: 1000 }, async () => {
    try {
      await py.buildWorkbook({ sheets: [{ rows: [['a']] }] });
      throw new Error('超量却没报错');
    } catch (err) {
      if (!err.message.includes('超量')) throw new Error('未说明超量: ' + err.message);
    }
  });
  return '超量中止并点名';
});

await t('子进程: 超时被杀(100ms)时报超时', async () => {
  await withFake({ mode: 'sleep', timeout: 100 }, async () => {
    try {
      await py.buildWorkbook({ sheets: [{ rows: [['a']] }] });
      throw new Error('超时却没报错');
    } catch (err) {
      if (!err.message.includes('超时')) throw new Error('未说明超时: ' + err.message);
    }
  });
  return '超时中止并点名';
});

await t('探测: 版本过低(3.6)按不可用处理', async () => {
  await withFake({ mode: 'oldpy' }, async () => {
    try {
      await py.resolvePython();
      throw new Error('低版本却被采纳');
    } catch (err) {
      if (err.code !== 'PY_XLSX_UNAVAILABLE') throw err;
      if (!err.message.includes('版本过低')) throw new Error('未说明版本过低: ' + err.message);
    }
  });
  return '报版本过低并跳过该解释器';
});

await t('编辑: JS 侧校验失败的 optional 操作被跳过且不进子进程', async () => {
  const { buf } = await py.buildWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] });
  const r = await py.editWorkbook(buf, [
    { op: 'conditional_format', sheet: 1, range: 'A1', rule: { type: 'colorScale', colors: ['FFFFFF'] }, optional: true },
    { op: 'set_value', sheet: 1, ref: 'C3', value: '仍执行' },
  ], ['S']);
  if (!r.changes[0].includes('跳过')) throw new Error('未记录跳过: ' + r.changes[0]);
  return r.changes[0];
});

// ---------------------------------------------------------------------------
// 校验分支(规则 / 取值)
// ---------------------------------------------------------------------------
await t('取值: 空对象/带 value 的对象/日期/真假值都能转换', async () => {
  const cells = await import('../lib/core/xlsx-cells.js');
  if (cells.coerceValue({ value: '=A1' }).formula !== 'A1') throw new Error('嵌套 value 未解引用');
  if (cells.coerceValue({}) !== '[object Object]') throw new Error('非 value 对象未降级为字符串: ' + cells.coerceValue({}));
  if (!(cells.coerceValue(new Date(0)) instanceof Date)) throw new Error('日期被改类型');
  if (cells.coerceValue(true) !== true) throw new Error('布尔被改类型');
  if (cells.coerceValue('num:1,234') !== 1234) throw new Error('num 前缀未转数字');
  if (cells.cellTag(undefined).v !== null) throw new Error('未定义值未落成空');
  return '各分支取值正确';
});

await t('规则: 条件格式文案类规则各自成式', async () => {
  const rules = await import('../lib/core/xlsx-rules.js');
  const cases = [
    ['containsText', 'NOT(ISERROR(SEARCH("甲",A1)))'],
    ['notContains', 'ISERROR(SEARCH("甲",A1))'],
    ['beginsWith', 'LEFT(A1,1)="甲"'],
    ['endsWith', 'RIGHT(A1,1)="甲"'],
    ['containsBlanks', 'LEN(TRIM(A1))=0'],
    ['notContainsBlanks', 'LEN(TRIM(A1))>0'],
    ['containsErrors', 'ISERROR(A1)'],
    ['notContainsErrors', 'NOT(ISERROR(A1))'],
  ];
  for (const [operator, expected] of cases) {
    const rule = rules.pythonRule({ type: 'containsText', operator, text: '甲' }, 'A1:A9');
    if (rule.formula[0] !== expected) throw new Error(`${operator} 生成 ${rule.formula[0]},期望 ${expected}`);
  }
  const unset = rules.pythonRule({ type: 'containsText', text: '甲' }, 'A1');
  if (!unset.formula[0].includes('SEARCH')) throw new Error('缺省 operator 未按 containsText 处理');
  return '8 种文案规则公式正确';
});

await t('规则: 色阶/图标集/数据条/前 N/高于均值/时间周期都生成 kind', async () => {
  const rules = await import('../lib/core/xlsx-rules.js');
  const two = rules.pythonRule({ type: 'colorScale', colors: ['FFFFFF', 'FF0000'] }, 'A1:A9');
  sameList(two.stops.map((s) => s.type), ['min', 'max'], '两色停点');
  const three = rules.pythonRule({ type: 'colorScale', colors: ['FFFFFF', 'FFEB84', 'FF0000'] }, 'A1:A9');
  sameList(three.stops.map((s) => s.type), ['min', 'percentile', 'max'], '缺省三色停点');
  const icons = rules.pythonRule({ type: 'iconSet', iconSet: '5Arrows' }, 'A1');
  if (icons.stops.length !== 5) throw new Error('5 图标停点数不对: ' + icons.stops.length);
  if (rules.pythonRule({ type: 'dataBar' }, 'A1').color !== '638EC6') throw new Error('数据条缺省色不对');
  if (rules.pythonRule({ type: 'top10', rank: 3 }, 'A1').rank !== 3) throw new Error('前 N 名次丢失');
  if (rules.pythonRule({ type: 'aboveAverage' }, 'A1').aboveAverage !== true) throw new Error('高于均值缺省值不对');
  if (rules.pythonRule({ type: 'timePeriod' }, 'A1').timePeriod !== 'today') throw new Error('时间周期缺省值不对');
  return '各规则 kind 与停点正确';
});

await t('规则: 条件格式的参数校验会拦住非法值', async () => {
  const rules = await import('../lib/core/xlsx-rules.js');
  const bad = [
    [{ type: 'nope' }, '可为'],
    [{ type: 'cellIs', operator: 'nope', value: 1 }, 'operator'],
    [{ type: 'cellIs' }, '阈值'],
    [{ type: 'cellIs', operator: 'between', value: 1 }, '两个阈值'],
    [{ type: 'expression' }, 'formula'],
    [{ type: 'containsText', text: '' }, 'text'],
    [{ type: 'colorScale', colors: ['FFFFFF'] }, '2 个或 3 个'],
    [{ type: 'iconSet', iconSet: 'bad' }, ''],
  ];
  for (const [rule, hint] of bad) {
    try {
      rules.pythonRule(rule, 'A1');
      throw new Error(`${JSON.stringify(rule)} 没报错`);
    } catch (err) {
      if (hint && !err.message.includes(hint) && !err.message.includes('不能为空')) {
        throw new Error(`${JSON.stringify(rule)} 的提示不含「${hint}」: ${err.message}`);
      }
    }
  }
  try {
    rules.pythonRule({}, 'A1');
    throw new Error('缺 type 没报错');
  } catch (err) {
    if (!err.message.includes('是必填项')) throw err;
  }
  if (rules.pythonRule({ type: 'cellIs', operator: 'between', values: [1, 2] }, 'A1').formula.length !== 2) {
    throw new Error('values 形式的两阈值未采纳');
  }
  return '类型/阈值/颜色/文本各错误都拦住';
});

await t('规则: range 形状与数据验证的三种取值来源', async () => {
  const rules = await import('../lib/core/xlsx-rules.js');
  try {
    rules.pythonRules({ range: 'B2', rule: { type: 'top10' } });
    rules.pythonRules({ range: 'A1', rule: { type: 'top10' } });
  } catch (err) {
    throw new Error('合法 range 被拒: ' + err.message);
  }
  try {
    rules.pythonDataValidation({ range: 'b2:B9' });
    throw new Error('小写 range 没报错');
  } catch (err) {
    if (!err.message.includes('形如')) throw err;
  }
  const list = rules.pythonValidation({ type: 'list', values: ['甲', '乙'] });
  if (list.formula1 !== '"甲,乙"') throw new Error('候选值拼装不对: ' + list.formula1);
  const fromSource = rules.pythonValidation({ type: 'list', source: '=Sheet2!$A$1:$A$5' });
  if (fromSource.formula1 !== 'Sheet2!$A$1:$A$5') throw new Error('source 未去掉等号');
  const fromFormulae = rules.pythonValidation({ rule: { type: 'list', formulae: ['"x,y"'] } });
  if (fromFormulae.formula1 !== '"x,y"') throw new Error('formulae 形式未采纳');
  const whole = rules.pythonValidation({ type: 'whole', value: 1, value2: 9, error: '请输入', errorStyle: 'warning', promptTitle: 'T', prompt: 'P', errorTitle: 'E' });
  if (whole.formula2 !== '9' || whole.errorStyle !== 'warning') throw new Error('区间/提示字段不全');
  if (!whole.showErrorMessage || !whole.showInputMessage) throw new Error('提示开关未随字段打开');
  return '三种取值来源与提示字段都正确';
});

await t('规则: 数据验证缺值/缺类型/空候选都会报错', async () => {
  const rules = await import('../lib/core/xlsx-rules.js');
  const bad = [
    [{ type: 'nope' }, '可为'],
    [{}, 'list 需要'],
    [{ type: 'list', values: [] }, '不能是空数组'],
    [{ type: 'whole' }, '需要 value'],
    [{ type: 'list', operator: 'nope' }, ''],
  ];
  for (const [rule, hint] of bad) {
    try {
      rules.pythonValidation(rule);
      throw new Error(`${JSON.stringify(rule)} 没报错`);
    } catch (err) {
      if (hint && !err.message.includes(hint)) throw new Error(`提示不含「${hint}」: ${err.message}`);
    }
  }
  return '各类缺参都被拦住';
});

await t('摘要: summaryOf 归一成 名称/行/列', async () => {
  const r = py.summaryOf([{ name: 'S', rows: 2, cols: 3 }]);
  sameList(r.sheets, [{ name: 'S', rows: 2, cols: 3 }], '摘要');
  return '摘要形状正确';
});

// ---------------------------------------------------------------------------
// 门面:仅 DSH_RESOLVER_BUG 才回退
// ---------------------------------------------------------------------------
await t('门面: 写入命中 DSH_RESOLVER_BUG 时改走 openpyxl 并标记 via', async () => {
  const built = await office.writeXlsxWorkbook({ sheets: [{ name: 'S', rows: [['a', 1]] }] }, async () => { throw resolverBug(); });
  if (built.via !== 'openpyxl') throw new Error('未标记 via: ' + built.via);
  if (!built.buf?.length) throw new Error('没有产物');
  return 'via=openpyxl';
});

await t('门面: 正常环境仍走 ExcelJS', async () => {
  const built = await office.writeXlsxWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] });
  if (built.via !== 'exceljs') throw new Error('走错后端: ' + built.via);
  return 'via=exceljs';
});

await t('门面: 非解析器故障的写入错误不回退、原样抛出', async () => {
  try {
    await office.writeXlsxWorkbook({ sheets: [{ rows: [['a']] }] }, async () => { throw new Error('磁盘满'); });
    throw new Error('普通错误被吞了');
  } catch (err) {
    if (err.message !== '磁盘满') throw err;
  }
  return '普通错误原样抛出';
});

await t('门面: 编辑回退时图表仍经 injectChart 注入', async () => {
  const { buf } = await py.buildWorkbook({ sheets: [{ name: 'S', rows: [['类别', '值'], ['A', 1], ['B', 2]] }] });
  const ops = [
    { op: 'set_value', sheet: 1, ref: 'D1', value: 'x' },
    { op: 'add_chart', sheet: 1, chartType: 'bar', categories: 'A2:A3', series: ['B2:B3'], anchor: 'F2' },
  ];
  const r = await office.editXlsxWorkbook(buf, ops, async () => { throw resolverBug(); }, async () => { throw resolverBug(); });
  if (r.via !== 'openpyxl') throw new Error('未走兜底: ' + r.via);
  if (!r.changes.includes('add_chart(bar)')) throw new Error('图表未注入: ' + r.changes.join('|'));
  const PizZip = (await import('pizzip')).default;
  const names = Object.keys(new PizZip(r.buf).files);
  if (!names.some((n) => n.startsWith('xl/charts/chart'))) throw new Error('包里没有图表部件');
  return 'op 走 openpyxl,图表后置注入';
});

await t('图表: optional 的注入失败被跳过,非 optional 抛出', async () => {
  const { buf } = await py.buildWorkbook({ sheets: [{ name: 'S', rows: [['a', 1]] }] });
  const bad = { op: 'add_chart', sheet: 9, chartType: 'bar', categories: 'A1:A1', series: ['B1:B1'], optional: true };
  const skipped = await office.applyChartOps(buf, [bad]);
  if (!skipped.changes[0].includes('跳过')) throw new Error('未跳过: ' + skipped.changes[0]);
  try {
    await office.applyChartOps(buf, [{ ...bad, optional: false }]);
    throw new Error('非 optional 失败却没抛');
  } catch (err) {
    if (/没抛/.test(err.message)) throw err;
  }
  return skipped.changes[0];
});

await t('工具层: 缺 sheets / 非 .xlsx / .xls 各自报错', async () => {
  const cases = [
    [() => office.opWriteXlsx(join(outDir, 'no-sheets.xlsx'), { sheets: [] }), '需要 sheets 参数'],
    [() => office.opWriteXlsx(join(outDir, 'wrong.docx'), { sheets: [{ rows: [['a']] }] }), '仅支持输出 .xlsx'],
    [() => office.opEditXlsx(join(outDir, 'old.xls'), [{ op: 'set_value', ref: 'A1', value: 1 }]), '旧格式'],
    [() => office.opEditXlsx(join(outDir, 'old.csv'), [{ op: 'set_value', ref: 'A1', value: 1 }]), '仅支持 .xlsx'],
  ];
  for (const [run, hint] of cases) {
    try {
      await run();
      throw new Error(`缺 ${hint} 校验`);
    } catch (err) {
      if (!err.message.includes(hint)) throw new Error(`提示不含「${hint}」: ${err.message}`);
    }
  }
  return '四类入口校验都拦住';
});

await t('工具层: ExcelJS 摘要取不到时 meta.before 留空仍能改', async () => {
  const path = join(outDir, 'no-before.xlsx');
  await office.opWriteXlsx(path, { sheets: [{ name: 'S', rows: [['a']] }] });
  const r = await office.opEditXlsx(path, [{ op: 'set_value', sheet: 1, ref: 'B1', value: 'b' }], undefined, {
    summary: async () => { throw resolverBug(); },
  });
  if (!r.content.includes('已更新 Excel')) throw new Error('编辑没成功: ' + r.content);
  return '摘要失败不影响编辑';
});

await t('工具层: 编辑含图表的文件时提示图表被丢弃', async () => {
  const path = join(outDir, 'tool-chart.xlsx');
  await office.opWriteXlsx(path, { sheets: [{ name: 'S', rows: [['类别', '值'], ['A', 1], ['B', 2]] }] }, undefined, { build: async () => { throw resolverBug(); } });
  const withChart = join(outDir, 'tool-chart2.xlsx');
  const { buf } = await office.writeXlsxWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] }).then((r) => r);
  const charted = await office.applyChartOps(buf, [{ op: 'add_chart', chartType: 'line', categories: 'A2:A2', series: ['B2:B2'] }]);
  await writeFile(withChart, charted.buf);
  const r = await office.opEditXlsx(withChart, [{ op: 'set_value', sheet: 1, ref: 'D1', value: 1 }], undefined, {
    edit: async () => { throw resolverBug(); },
    read: async () => { throw resolverBug(); },
  });
  if (!r.content.includes('被丢弃')) throw new Error('未提示图表被丢弃: ' + r.content);
  return '提示原有图表被丢弃';
});

await t('门面: 编辑回退会提示原有图表被丢弃', async () => {
  const withChart = await office.writeXlsxWorkbook({ sheets: [{ name: 'S', rows: [['类别', '值'], ['A', 1]] }] })
    .then(async ({ buf }) => (await office.applyChartOps(buf, [{ op: 'add_chart', chartType: 'line', categories: 'A2:A2', series: ['B2:B2'] }])).buf);
  const r = await office.editXlsxWorkbook(withChart, [{ op: 'set_value', sheet: 1, ref: 'C1', value: 1 }], async () => { throw resolverBug(); }, async () => { throw resolverBug(); });
  if (r.droppedCharts !== true) throw new Error('未检出原有图表: ' + r.droppedCharts);
  return 'droppedCharts=true';
});

await t('门面: 产物能被 ExcelJS 读回(与读取路径兼容)', async () => {
  const built = await office.writeXlsxWorkbook({ sheets: [{ name: 'S', header: true, rows: [['项目'], ['甲']] }] }, async () => { throw resolverBug(); });
  const back = await readWorkbook(built.buf, {});
  if (!back.content.includes('甲')) throw new Error('读不回内容');
  return 'ExcelJS 可读回 openpyxl 产物';
});

// ---------------------------------------------------------------------------
// 工具层端到端
// ---------------------------------------------------------------------------
await t('工具层: opWriteXlsx 在故障下产出文件并带 openpyxl 说明', async () => {
  const path = join(outDir, 'tool.xlsx');
  const r = await office.opWriteXlsx(path, { sheets: [{ name: 'S', header: true, rows: [['a', 1]] }] }, undefined, { build: async () => { throw resolverBug(); } });
  const bytes = await readFile(path);
  assertOfficeBinary(bytes, 'xlsx');
  if (!r.content.includes('openpyxl')) throw new Error('未说明已改用 openpyxl: ' + r.content);
  if (!r.content.includes('S(1行×2列)')) throw new Error('摘要不对: ' + r.content);
  return '工具层可用,含兜底说明与真实摘要';
});

await t('工具层: opEditXlsx 故障下也能改并带说明', async () => {
  const path = join(outDir, 'tool-edit.xlsx');
  await office.opEditXlsx(path, [{ op: 'set_value', sheet: 1, ref: 'B1', value: 'b' }], undefined, {
    edit: async () => { throw resolverBug(); },
    read: async () => { throw resolverBug(); },
  }).then(() => {}, async () => {
    // 目标文件还不存在,先用兜底生成一份
    await office.opWriteXlsx(path, { sheets: [{ name: 'S', rows: [['a']] }] }, undefined, { build: async () => { throw resolverBug(); } });
  });
  const r = await office.opEditXlsx(path, [{ op: 'set_value', sheet: 1, ref: 'B1', value: 'b' }], undefined, {
    edit: async () => { throw resolverBug(); },
    read: async () => { throw resolverBug(); },
  });
  if (!r.content.includes('openpyxl')) throw new Error('未说明已改用 openpyxl: ' + r.content);
  return '编辑工具带兜底说明';
});

// ---------------------------------------------------------------------------
// 汇总
// ---------------------------------------------------------------------------
} else {
  console.log('\n兜底后端用例已跳过(需 Python 3.8+ 与 openpyxl)');
}

const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} 通过`);
for (const r of failed) console.log(`✗ ${r.name}: ${r.detail}`);
process.exit(failed.length ? 1 : 0);
