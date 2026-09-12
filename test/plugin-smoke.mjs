// Plugin-adapter smoke test: drives lib/index.js with a fake DSH host ctx
// (no real cordis needed) and exercises every registered tool end to end.
import { mkdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, 'out-plugin');
await mkdir(outDir, { recursive: true });

const registered = new Map();
let logLines = [];
const ctx = {
  tools: {
    register(tool) {
      if (registered.has(tool.name)) throw new Error(`重复注册: ${tool.name}`);
      registered.set(tool.name, tool);
      return () => registered.delete(tool.name);
    },
  },
  get(name) { return name === 'tools' ? ctx.tools : undefined; },
  emit() { /* no observers in smoke */ },
  effect(fn) { const d = fn(); return d; },
  logger: { info: (m) => logLines.push(m), warn: (m) => logLines.push(m) },
};

const mod = await import('../lib/index.js');
mod.apply(ctx);

const results = [];
function check(name, cond, detail = '') {
  results.push({ name, ok: Boolean(cond), detail });
  console.log(`${cond ? '✓' : '✗'} ${name}${detail ? ' — ' + detail : ''}`);
}

check('apply 注册 6 个工具', registered.size === 6, [...registered.keys()].join(', '));
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

const tplDef = await registered.get('office_write_docx').execute({ path: tpl, markdown: '订单号：{{orderNo}}\n客户：{{customer}}' }, exec);
const r6 = await registered.get('office_fill_docx_template').execute(
  { templatePath: tpl, outputPath: filled, data: { orderNo: 'SO-2026-88', customer: '测试客户' } }, exec);
const r6b = await registered.get('office_read').execute({ path: filled }, exec);
check('office_fill_docx_template', r6.content.includes('已按模板生成') && r6b.content.includes('SO-2026-88'), r6.content.split('\n')[0]);

const r7 = await registered.get('office_convert').execute({ sourcePath: docx, outputPath: conv }, exec);
const r7b = await registered.get('office_read').execute({ path: conv }, exec);
check('office_convert(docx→doc) + 读回', r7.content.includes('转换完成') && r7b.content.includes('冒烟测试'), r7.content.split('\n')[0]);

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

console.log('\n日志: ' + logLines.join(' | '));
const failed = results.filter((r) => !r.ok);
console.log(`\n===== 插件冒烟: ${results.length - failed.length}/${results.length} 通过 =====`);
if (failed.length) {
  console.log('失败项: ' + failed.map((f) => f.name).join(', '));
  process.exit(1);
}
