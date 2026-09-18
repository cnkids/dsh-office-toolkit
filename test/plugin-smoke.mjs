// Plugin-adapter smoke test: drives lib/index.js with a fake DSH host ctx
// (no real cordis needed) and exercises every registered tool end to end.
import { mkdir, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { homedir } from 'node:os';
import { deflateSync } from 'node:zlib';

const here = dirname(fileURLToPath(import.meta.url));
// .doc 写出依赖外部转换器(macOS textutil / LibreOffice / Word);
// 没有时跳过相关断言,好让 CI 的 Linux runner 也能跑这一套
const { hasWordConverter } = await import('../lib/core/converters.js');
const canWriteDoc = await hasWordConverter('doc').catch(() => false);
const outDir = join(here, 'out-plugin');
await mkdir(outDir, { recursive: true });

/** 最小合法 PNG(仅用于冒烟:尺寸写在 IHDR 里)。 */
function pngFixture(width, height) {
  const crcOf = (buf) => {
    let c = 0xffffffff;
    for (const byte of buf) {
      c ^= byte;
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const head = Buffer.alloc(4);
    head.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crcOf(body));
    return Buffer.concat([head, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const raw = Buffer.concat(Array.from({ length: height }, () => Buffer.concat([Buffer.from([0]), Buffer.alloc(width * 3, 0x66)])));
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const registered = new Map();
const logLines = [];
const emitted = [];
// 先以"无 fs 提供者"运行(覆盖回退分支),后面再切到提供者分支
let useFsProvider = false;
const fakeFs = {
  async resolve(raw) { return { displayPath: raw, path: raw }; },
  async readBytes(target) { return readFile(target.path); },
  async stat() { return { version: 1 }; },
};
const ctx = {
  tools: {
    register(tool) {
      if (registered.has(tool.name)) throw new Error(`重复注册: ${tool.name}`);
      registered.set(tool.name, tool);
      return () => registered.delete(tool.name);
    },
  },
  get(name) {
    if (name === 'tools') return ctx.tools;
    if (name === 'fs' && useFsProvider) return fakeFs;
    return undefined;
  },
  emit(name, ...rest) { emitted.push([name, ...rest]); },
  effect(fn) { const d = fn(); return d; },
  logger: { info: (m) => logLines.push(m), warn: (m) => logLines.push(m) },
};

const mod = await import('../lib/index.js');
mod.apply(ctx);

// 模拟 DSH 宿主:传给工具的参数对象是冻结的,插件不得改动调用方参数
for (const tool of registered.values()) {
  const run = tool.execute;
  tool.execute = (args, exec) => run.call(tool, Object.freeze({ ...args }), exec);
}

const results = [];
function check(name, cond, detail = '') {
  results.push({ name, ok: Boolean(cond), detail });
  console.log(`${cond ? '✓' : '✗'} ${name}${detail ? ' — ' + detail : ''}`);
}

check('apply 注册 8 个工具', registered.size === 8, [...registered.keys()].join(', '));
for (const [n, t] of registered) {
  check(`工具 ${n} 结构完整`, t.parameters?.type === 'object' && t.output?.render && typeof t.execute === 'function' && t.timeoutMs > 0);
}

// Schema shape regression: the host validates real tool arguments against these
// schemas, so a mis-shaped parameter object silently breaks agent calls.
for (const [n, t] of registered) {
  const props = t.parameters.properties || {};
  const shell = Object.entries(props).filter(([, v]) => v && v.type === undefined && v.description === undefined && v.enum === undefined && v.oneOf === undefined);
  check(`工具 ${n} 参数无空壳字段`, shell.length === 0, shell.map(([k]) => k).join(', '));
  const required = t.parameters.required || [];
  check(`工具 ${n} required 已声明`, required.every((k) => k in props), JSON.stringify(required));
}

// PTC 模式下工具描述会被嵌进 tools:sdk 提示词段落,再由 DSH 做 {{变量}} 插值 ——
// 描述里出现 `{{变量}}` 这种字面量会被当成变量引用,直接让整个 prompt 组装抛错
// (malformed prompt variable reference)。所以面向模型的文本里一律不能有 {{...}}。
const VARIABLE_LITERAL = /\{\{[^{}]*\}\}/;
function collectText(value, out = []) {
  if (typeof value === 'string') out.push(value);
  else if (Array.isArray(value)) for (const item of value) collectText(item, out);
  else if (value && typeof value === 'object') for (const item of Object.values(value)) collectText(item, out);
  return out;
}
for (const [n, t] of registered) {
  const texts = [...collectText(t.description), ...collectText(t.parameters)];
  const bad = texts.filter((s) => VARIABLE_LITERAL.test(s));
  check(`工具 ${n} 描述不含 {{}} 变量字面量`, bad.length === 0, bad.join(' | ').slice(0, 70));
}
const xlsxSheets = registered.get('office_write_xlsx').parameters.properties.sheets;
check('office_write_xlsx.sheets 为数组 schema',
  xlsxSheets?.type === 'array' && xlsxSheets?.items?.type === 'object',
  JSON.stringify(xlsxSheets?.type));
const xlsxOps = registered.get('office_edit_xlsx').parameters.properties.ops;
check('office_edit_xlsx.ops 为数组 schema',
  xlsxOps?.type === 'array' && xlsxOps?.items?.properties?.op !== undefined,
  JSON.stringify(xlsxOps?.type));
check('office_edit_xlsx.ops.items.required 含 op', (xlsxOps?.items?.required || []).includes('op'));

const exec = {}; // no agent → relative paths resolve against process.cwd()

const docx = join(outDir, 'smoke.docx');
const xlsx = join(outDir, 'smoke.xlsx');
const tpl = join(outDir, 'tpl.docx');
const filled = join(outDir, 'filled.docx');
const conv = join(outDir, 'smoke.doc');

const r1 = await registered.get('office_write_docx').execute(
  { path: docx, markdown: '# 冒烟测试\n\n**插件** 工具链验证。\n\n| 项 | 值 |\n| --- | --- |\n| A | 1 |\n' }, exec);
check('office_write_docx', r1.content.includes('已生成 Word 文档'), r1.content.split('\n')[0]);

const r2 = await registered.get('office_read').execute({ path: docx }, exec);
check('office_read(docx)', r2.content.includes('冒烟测试') && r2.content.includes('| A'), r2.content.split('\n')[0]);

const r2p = await registered.get('office_read').execute({ path: docx, maxChars: 24 }, exec);
check('office_read 分段阅读: 给总长 + 下次的 offset',
  /共 \d+ 字符/.test(r2p.content) && /继续读: 再调用本工具并传 offset: \d+/.test(r2p.content),
  r2p.content.split('\n').slice(3, 5).join(' | '));
const r2o = await registered.get('office_read').execute({ path: docx, outline: true }, exec);
check('office_read 标题大纲', r2o.content.includes('标题大纲') && r2o.content.includes('冒烟测试'), r2o.content.split('\n')[3]);

const r3 = await registered.get('office_write_xlsx').execute({
  path: xlsx,
  sheets: [
    { name: '明细', header: true, rows: [['产品', '数量', '金额'], ['键盘', 10, '=B2*199'], ['鼠标', 20, '=B3*89']] },
    { name: '备注', rows: [['说明'], ['冒烟测试']] },
  ],
}, exec);
check('office_write_xlsx', r3.content.includes('已生成 Excel'), r3.content.split('\n')[0]);

const r4 = await registered.get('office_edit_xlsx').execute({
  path: xlsx,
  ops: [
    { op: 'set_value', sheet: '明细', ref: 'E1', value: '合计', style: { bold: true } },
    { op: 'set_formula', sheet: '明细', ref: 'C4', formula: '=SUM(C2:C3)' },
    { op: 'style_range', sheet: '明细', range: 'A1:C1', style: { fill: 'DDEBF7' } },
    { op: 'add_chart', sheet: '明细', chartType: 'column', title: '金额对比', categories: 'A2:A3', series: [{ range: 'C2:C3' }], anchor: 'G2' },
  ],
}, exec);
check('office_edit_xlsx(含 add_chart)', r4.content.includes('add_chart'), r4.content.split('\n')[1] || '');

const r5 = await registered.get('office_read').execute({ path: xlsx, sheets: ['明细'] }, exec);
check('office_read(xlsx)', r5.content.includes('键盘') && r5.content.includes('SUM'), r5.content.split('\n')[0]);

const r5q = await registered.get('office_query').execute({ path: xlsx }, exec);
check('office_query(画像)', r5q.content.includes('表结构画像') && r5q.content.includes('产品'), r5q.content.split('\n')[0]);
const r5b = await registered.get('office_query').execute({
  path: xlsx, groupBy: ['产品'], aggregate: [{ col: '数量', fn: 'sum', as: '数量合计' }],
}, exec);
check('office_query(分组聚合)', r5b.content.includes('键盘') && r5b.content.includes('数量合计'), r5b.content.split('\n')[0]);
const r5c = await registered.get('office_query').execute({ path: xlsx, where: [{ col: '数量', op: 'gt', value: 10 }] }, exec);
check('office_query(筛选)', r5c.content.includes('命中 1 行'), r5c.content.split('\n').find((l) => l.startsWith('###')));

// 图片嵌入:工具层要按沙箱规则解析 <img src> 并注入读取器(core 不碰文件系统)
const imgPath = join(outDir, 'logo.png');
await writeFile(imgPath, pngFixture(24, 12));
const docxImg = join(outDir, 'with-image.docx');
const r1i = await registered.get('office_write_docx').execute(
  { path: docxImg, markdown: `# 带图\n\n![徽标](${imgPath})\n\n<img alt="占位">` }, exec);
check('office_write_docx(嵌入本地图片)', r1i.content.includes('已生成 Word 文档'), r1i.content.split('\n')[0]);
// zip 的中央目录里条目名是明文,不必解压就能确认图片部件在
const imgBytes = await readFile(docxImg);
check('图片真的进了 docx', imgBytes.includes('word/media/'), String(imgBytes.length) + ' 字节');
const missing = await registered.get('office_write_docx').execute(
  { path: docxImg, markdown: '<img src="不存在的图.png">' }, exec).catch((err) => err);
check('缺失图片报错清晰', String(missing.message || missing).includes('读不到图片文件'), String(missing.message || missing).slice(0, 80));

// 多表 join + 透视:走真实工具入口(文件、参数、schema 一起验证)
const xlsx2 = join(outDir, 'smoke-join.xlsx');
await registered.get('office_write_xlsx').execute({
  path: xlsx2,
  sheets: [{ name: '价格', header: true, rows: [['产品', '单价'], ['键盘', 199], ['鼠标', 89]] }],
}, exec);
const r5j = await registered.get('office_query').execute({
  path: xlsx,
  join: { path: xlsx2, sheet: '价格', on: '产品', type: 'left' },
  groupBy: ['产品'],
  aggregate: [{ col: '单价', fn: 'max', as: '单价' }],
  orderBy: [{ col: '单价', dir: 'desc' }],
}, exec);
check('office_query(join 另一张表)',
  r5j.content.includes('连接：与「价格」') && r5j.content.includes('键盘\t199'),
  r5j.content.split('\n').find((l) => l.startsWith('###')));
const r5p = await registered.get('office_query').execute({
  path: xlsx,
  pivot: { rows: ['产品'], columns: '产品', values: [{ col: '数量', fn: 'sum' }], totals: true },
}, exec);
check('office_query(透视表)', r5p.content.includes('合计') && r5p.content.includes('键盘'), r5p.content.split('\n').find((l) => l.startsWith('###')));

const r2e = await registered.get('office_edit_docx').execute({
  path: docx,
  ops: [
    { op: 'replace_text', find: '冒烟测试', replace: '冒烟测试（已改）' },
    { op: 'insert_paragraph', text: '编辑时追加的段落', position: 'end' },
  ],
}, exec);
check('office_edit_docx(替换 + 插入段落)',
  r2e.content.includes('替换文字 1 处') && r2e.content.includes('插入段落') && r2e.content.includes('图片/页眉页脚/样式'),
  r2e.content.split('\n')[1]);
const r2e2 = await registered.get('office_read').execute({ path: docx }, exec);
check('office_edit_docx 结果可读回', r2e2.content.includes('冒烟测试（已改）') && r2e2.content.includes('编辑时追加的段落'), r2e2.content.split('\n')[0]);

// 排版:写的时候设整篇样式,读回来核对格式报告(端到端往返)
const styled = join(outDir, 'styled.docx');
const rStyle = await registered.get('office_write_docx').execute({
  path: styled,
  markdown: '# 通知标题\n\n这是正文第一段。\n\n这是正文第二段。',
  style: { font: '仿宋_GB2312', sizePt: 16, lineSpacingPt: 28.8, firstLineIndentChars: 2, align: 'both', headings: { font: '黑体', sizePt: 16 } },
}, exec);
check('office_write_docx(style 整篇排版)', rStyle.content.includes('已生成 Word 文档'), rStyle.content.split('\n')[0]);
const rStyleRead = await registered.get('office_read').execute({ path: styled, withFormatting: true }, exec);
const styleMarkers = [
  ['文档默认字体', rStyleRead.content.includes('字体 仿宋_GB2312')],
  ['三号 16pt', rStyleRead.content.includes('16pt')],
  ['行距固定值 28.8pt', rStyleRead.content.includes('固定值 28.8pt')],
  ['首行缩进 2 字符(32pt)', rStyleRead.content.includes('首行 32pt')],
  ['两端对齐', rStyleRead.content.includes('两端对齐')],
  ['标题用黑体', rStyleRead.content.includes('黑体')],
];
check('排版往返核对(写字号/行距/缩进/对齐)', styleMarkers.every(([, ok]) => ok), styleMarkers.filter(([, ok]) => !ok).map(([n]) => n).join(',') || '6 项全部命中');

// 改已有文档的排版
const rRestyle = await registered.get('office_edit_docx').execute({
  path: styled,
  ops: [{ op: 'set_style', scope: 'all', font: '楷体_GB2312', sizePt: 14, lineSpacingMultiple: 1.5 }],
}, exec);
check('office_edit_docx(set_style 全篇改排版)', rRestyle.content.includes('设置') && rRestyle.content.includes('段样式'), rRestyle.content.split('\n')[1]);
const rRestyleRead = await registered.get('office_read').execute({ path: styled, withFormatting: true }, exec);
check('改后排版可读回', rRestyleRead.content.includes('楷体_GB2312') && rRestyleRead.content.includes('1.5 倍'), (rRestyleRead.content.match(/行距[^|\n]*/) || [''])[0].slice(0, 40));

// 表格:写的时候指定三线表 + 自动列宽,读回来核对;再用 set_table 改框线、增删行列
const tableDoc = join(outDir, 'table.docx');
await registered.get('office_write_docx').execute({
  path: tableDoc,
  markdown: '# 人员表\n\n| 姓名 | 部门 | 备注 |\n| --- | --- | --- |\n| 张三 | 技术研发中心 | 负责人 |\n| 李四 | 财务 | 兼会计 |',
  style: { font: '仿宋_GB2312', sizePt: 16, table: { borders: 'three-line', headerShading: 'F2F2F2', headerBold: true, columnWidthMode: 'auto', align: 'center' } },
}, exec);
const tableRead = await registered.get('office_read').execute({ path: tableDoc }, exec);
check('office_write_docx(表格三线表) 可读回', tableRead.content.includes('张三') && tableRead.content.includes('技术研发中心'), tableRead.content.split('\n')[0]);

const rTable = await registered.get('office_edit_docx').execute({
  path: tableDoc,
  ops: [
    { op: 'set_table', table: 1, borders: 'all', headerShading: 'none', align: 'left' },
    { op: 'insert_table_row', table: 1, at: 2, count: 1 },
    { op: 'merge_table_cells', table: 1, range: 'A1:B1' },
  ],
}, exec);
check('office_edit_docx(表格样式 + 增行 + 合并)',
  rTable.content.includes('设置 1 个表格的样式') && rTable.content.includes('插入 1 行') && rTable.content.includes('合并 A1:B1'),
  rTable.content.split('\n')[1]);
const tableAfter = await registered.get('office_read').execute({ path: tableDoc }, exec);
check('表格改动后仍可读回', tableAfter.content.includes('张三') && tableAfter.content.includes('李四'), tableAfter.content.split('\n')[0]);

// 单元格内边距
const rMargins = await registered.get('office_edit_docx').execute({
  path: tableDoc,
  ops: [{ op: 'set_table', table: 1, cellMargins: { left: 108, right: 108, top: 40, bottom: 40 } }],
}, exec);
check('office_edit_docx(set_table cellMargins)', rMargins.content.includes('设置 1 个表格的样式'), rMargins.content.split('\n')[1]);

// 多级自动编号
const numDoc = join(outDir, 'numbering.docx');
await registered.get('office_write_docx').execute({
  path: numDoc,
  markdown: '# 概述\n\n- 要点一\n- 要点二\n\n## 实施细节\n\n- 细节点',
}, exec);
const rNum = await registered.get('office_edit_docx').execute({
  path: numDoc,
  ops: [{ op: 'set_numbering', scope: 'all', style: 'multicol-1_1_1', linkToHeading: true }],
}, exec);
check('office_edit_docx(set_numbering 多级编号)',
  rNum.content.includes('多级编号') && rNum.content.includes('正文列表 3 段') && rNum.content.includes('样式链接'),
  rNum.content.split('\n')[1]);
// 幂等 + exclude / startFrom
const rNum2 = await registered.get('office_edit_docx').execute({
  path: numDoc,
  ops: [{ op: 'set_numbering', scope: 'all', exclude: ['Heading1'], startFrom: { Heading2: 1 } }],
}, exec);
check('office_edit_docx(set_numbering 幂等 + exclude/startFrom)',
  rNum2.content.includes('更新已有编号定义') && rNum2.content.includes('多级编号'),
  rNum2.content.split('\n')[1]);

const numRead = await registered.get('office_read').execute({ path: numDoc, outline: true }, exec);
check('编号后标题大纲仍正常', numRead.content.includes('概述') && numRead.content.includes('实施细节'), numRead.content.split('\n')[3]);

// 删除整张表格
const rDropTable = await registered.get('office_edit_docx').execute({
  path: tableDoc,
  ops: [{ op: 'delete_table', table: 1 }],
}, exec);
check('office_edit_docx(delete_table 删除整表)', rDropTable.content.includes('删除表格 1'), rDropTable.content.split('\n')[1]);
const afterDrop = await registered.get('office_read').execute({ path: tableDoc }, exec);
check('删表后文档仍可读且正文保留', afterDrop.content.includes('人员表') && !afterDrop.content.includes('技术研发中心'), afterDrop.content.split('\n')[0]);

// onlyEmpty:先删空一张表的行,再一次清掉残留
const emptyDoc = join(outDir, 'empty-table.docx');
await registered.get('office_write_docx').execute({ path: emptyDoc, markdown: '正文\n\n| a |\n| --- |\n| 1 |\n\n结尾' }, exec);
await registered.get('office_edit_docx').execute({ path: emptyDoc, ops: [{ op: 'delete_table_row', table: 1, at: 1, count: 2 }] }, exec);
const rOnlyEmpty = await registered.get('office_edit_docx').execute({ path: emptyDoc, ops: [{ op: 'delete_table', onlyEmpty: true }] }, exec);
check('office_edit_docx(delete_table onlyEmpty)', rOnlyEmpty.content.includes('删除表格 1'), rOnlyEmpty.content.split('\n')[1]);

// 按角色改排版:只改正文,标题不动
const rRole = await registered.get('office_edit_docx').execute({
  path: tableDoc,
  ops: [{ op: 'set_style', scope: 'headings', sizePt: 22, font: '黑体', align: 'center' }],
}, exec);
check('office_edit_docx(set_style scope=headings)', rRole.content.includes('设置 1 段样式'), rRole.content.split('\n')[1]);

const tplDef = await registered.get('office_write_docx').execute({ path: tpl, markdown: '订单号：{{orderNo}}\n客户：{{customer}}' }, exec);
const r6 = await registered.get('office_fill_docx_template').execute(
  { templatePath: tpl, outputPath: filled, data: { orderNo: 'SO-2026-88', customer: '测试客户' } }, exec);
const r6b = await registered.get('office_read').execute({ path: filled }, exec);
check('office_fill_docx_template', r6.content.includes('已按模板生成') && r6b.content.includes('SO-2026-88'), r6.content.split('\n')[0]);
check('模板填充的结果文本不含 {{}} 字面量', !VARIABLE_LITERAL.test(r6.content), r6.content.split('\n').at(-1)?.slice(0, 70));

if (canWriteDoc) {
  const r7 = await registered.get('office_convert').execute({ sourcePath: docx, outputPath: conv }, exec);
  const r7b = await registered.get('office_read').execute({ path: conv }, exec);
  check('office_convert(docx→doc) + 读回', r7.content.includes('转换完成') && r7b.content.includes('冒烟测试'), r7.content.split('\n')[0]);
} else {
  check('office_convert(docx→doc) + 读回', true, '无外部转换器,跳过');
}

// containment fence: writes outside workspace/tmp must be denied
let denied = false;
try {
  await registered.get('office_write_docx').execute({ path: '/etc/dsh-office-should-not-exist.docx', markdown: '# x' }, exec);
} catch (err) {
  denied = /FS_SANDBOX_DENIED|写入被拒绝/.test(String(err?.message));
}
check('工作区外写入被拒绝', denied);

// error surfacing: reading a missing file must throw a clear error
let missingMsg = '';
try {
  await registered.get('office_read').execute({ path: join(outDir, 'nope.docx') }, exec);
} catch (err) {
  missingMsg = String(err?.message);
}
check('缺失文件报错清晰', /文件不存在|not found|ENOENT/.test(missingMsg), missingMsg.slice(0, 60));

// ---- ctx.fs 提供者分支:解析走 ctx.fs、读取走 readBytes,并发出 fs/observed ----
useFsProvider = true;
const viaFs = await registered.get('office_read').execute({ path: docx }, exec);
check('经 ctx.fs 提供者读取', viaFs.content.includes('冒烟测试'), viaFs.content.split('\n')[0]);
check('读取后发出 fs/observed', emitted.some(([n]) => n === 'fs/observed'), `事件 ${emitted.length} 条`);

const rendered = registered.get('office_read').output.render([], { content: 'x' });
check('工具 output.render 可用', Array.isArray(rendered) && rendered[0].text === 'x');

const beforeAbsent = emitted.length;
let absentMsg = '';
try {
  await registered.get('office_read').execute({ path: join(outDir, 'nope.xlsx') }, exec);
} catch (err) {
  absentMsg = String(err?.message);
}
const sawAbsent = emitted.slice(beforeAbsent).some(([n, , payload]) => n === 'fs/observed' && payload?.kind === 'absent');
check('缺失文件报错并标记 absent', /文件不存在/.test(absentMsg) && sawAbsent, absentMsg.slice(0, 40));

// 指向工作区之外的符号链接不得绕过路径围栏(围栏必须解析真实路径)
const escapeLink = join(outDir, 'escape-link');
await rm(escapeLink, { force: true });
await symlink(homedir(), escapeLink);
let linkDenied = false;
try {
  await registered.get('office_write_docx').execute({ path: join(escapeLink, 'pwned.docx'), markdown: '# x' }, exec);
} catch (err) {
  linkDenied = /FS_SANDBOX_DENIED|写入被拒绝/.test(String(err?.message));
}
check('符号链接绕行被拒绝', linkDenied, linkDenied ? '' : '❌ 围栏被绕过');
await rm(escapeLink, { force: true });

console.log('\n日志: ' + logLines.join(' | '));
const failed = results.filter((r) => !r.ok);
console.log(`\n===== 插件冒烟: ${results.length - failed.length}/${results.length} 通过 =====`);
if (failed.length) {
  console.log('失败项: ' + failed.map((f) => f.name).join(', '));
  process.exit(1);
}
