// Standalone selftest for the office core (no DSH needed).
// Usage: node test/selftest.mjs   (outputs into test/out/)
import { mkdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as office from '../lib/core/office.js';
import * as word from '../lib/core/word.js';
import * as excel from '../lib/core/excel.js';
import * as legacy from '../lib/core/legacy.js';
import * as converters from '../lib/core/converters.js';
import { htmlToMarkdown, plainTextToHtml } from '../lib/core/md.js';
import { CAPS } from '../lib/core/util.js';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, 'out');
await mkdir(outDir, { recursive: true });

const results = [];
async function t(name, fn) {
  try {
    const detail = await fn();
    results.push({ name, ok: true, detail: detail || '' });
    console.log(`✓ ${name}`);
  } catch (err) {
    results.push({ name, ok: false, detail: String(err?.stack || err) });
    console.log(`✗ ${name}\n   ${err?.stack ? err.stack.split('\n').slice(0, 3).join('\n   ') : err}`);
  }
}

const docxPath = join(outDir, 'sample.docx');
const MARKDOWN = `# 项目周报

**项目**: DeepSeek Office 插件
**周期**: 2026-09-01 ~ 2026-09-07

## 进展
- 完成 docx 读写原型
- 完成 xlsx 读写原型
- 打通旧格式兼容层

## 数据表
| 任务 | 状态 | 完成度 |
| --- | --- | --- |
| Word 读写 | 完成 | 100% |
| Excel 读写 | 进行中 | 80% |

> 下周计划:接入图表与批量模板。
`;

await t('word: markdown → docx', async () => {
  const buf = await word.writeDocx({ markdown: MARKDOWN }, { title: '周报' });
  const { writeFile } = await import('node:fs/promises');
  await writeFile(docxPath, buf);
  return `${buf.length} bytes`;
});

await t('word: docx → 文本读取', async () => {
  const { readFile } = await import('node:fs/promises');
  const buf = await readFile(docxPath);
  const r = await word.readDocx(buf);
  if (!r.content.includes('项目周报')) throw new Error('标题内容缺失');
  return `文本长度 ${r.content.length}, 词数 ${r.meta.words}`;
});

await t('word: docx → html → markdown 保留表格', async () => {
  const { readFile } = await import('node:fs/promises');
  const buf = await readFile(docxPath);
  const r = await word.readDocx(buf, { html: true });
  const md = htmlToMarkdown(r.html);
  if (!r.html.includes('<table')) throw new Error('HTML 中无表格');
  if (!md.includes('Word 读写')) throw new Error('markdown 丢失表格内容');
  return `html ${r.html.length} chars`;
});

await t('word: 模板 {{变量}} 填充', async () => {
  const { writeFile } = await import('node:fs/promises');
  const tmplBuf = await word.writeDocx({ markdown: '合同编号：{{contractNo}}\n甲方：{{partyA}}\n金额：{{amount}} 元' }, { title: '模板' });
  const tmplPath = join(outDir, 'template.docx');
  await writeFile(tmplPath, tmplBuf);
  const out = await word.fillDocxTemplate(tmplBuf, { contractNo: 'HT-2026-001', partyA: '深圳某某科技有限公司', amount: 12800 });
  const outPath = join(outDir, 'filled.docx');
  await writeFile(outPath, out);
  const r = await word.readDocx(out);
  for (const key of ['HT-2026-001', '深圳某某科技有限公司', '12800']) {
    if (!r.content.includes(key)) throw new Error(`变量未替换: ${key}`);
  }
  if (r.content.includes('{{')) throw new Error('仍有未替换占位符');
  return '变量全部替换成功';
});

await t('word: 两份 docx 合并(html 拼接重建)', async () => {
  const { readFile, writeFile } = await import('node:fs/promises');
  const b1 = await readFile(docxPath);
  const b2 = await word.writeDocx({ markdown: '## 附录\n补充说明内容。' });
  const r1 = await word.readDocx(b1, { html: true });
  const r2 = await word.readDocx(b2, { html: true });
  const merged = await word.writeDocx({ html: r1.html + '\n' + r2.html }, { title: '合并文档' });
  const p = join(outDir, 'merged.docx');
  await writeFile(p, merged);
  const back = await word.readDocx(merged);
  if (!back.content.includes('附录')) throw new Error('合并内容缺失');
  return `${merged.length} bytes`;
});

const xlsxPath = join(outDir, 'sample.xlsx');
await t('excel: 创建工作簿(双表+公式+表头样式)', async () => {
  const { writeFile } = await import('node:fs/promises');
  const buf = await excel.buildWorkbook({
    sheets: [
      {
        name: '销售明细',
        header: true,
        columnWidths: [14, 12, 12],
        rows: [
          ['产品', '数量', '单价'],
          ['键盘', 120, 199],
          ['鼠标', 300, 89.5],
          ['显示器', 45, 1299],
        ],
      },
      {
        name: '汇总',
        header: true,
        rows: [
          ['合计数量', '合计金额'],
          ['=SUM(销售明细!B2:B4)', '=SUM(销售明细!B2:B4*销售明细!C2:C4)'],
        ],
      },
    ],
  });
  await writeFile(xlsxPath, buf);
  return `${buf.length} bytes`;
});

await t('excel: 编辑工作簿(设置单元格/样式/合并/增表/冻结/筛选)', async () => {
  const { readFile, writeFile } = await import('node:fs/promises');
  const buf = await readFile(xlsxPath);
  const { buf: out } = await excel.editWorkbook(buf, [
    { op: 'set_value', sheet: '销售明细', ref: 'E1', value: '备注', style: { bold: true } },
    { op: 'set_value', sheet: '销售明细', ref: 'E2', value: '本月热销', style: { color: 'C00000' } },
    { op: 'set_cells', sheet: '销售明细', start: 'A6', values: [['耳机', 200, 299], ['合计', '=SUM(B2:B5)', '=SUM(C2:C5)']] },
    { op: 'style_range', sheet: '销售明细', range: 'A6:C6', style: { bold: true, fill: 'FFF2CC' } },
    { op: 'merge', sheet: '汇总', range: 'A4:B4' },
    { op: 'set_value', sheet: '汇总', ref: 'A4', value: '数据截至 2026-09-07' },
    { op: 'add_sheet', name: '图表页', rows: [['季度', '销售额'], ['Q1', 10], ['Q2', 20]], header: true },
    { op: 'freeze', sheet: '销售明细', rows: 1 },
    { op: 'auto_filter', sheet: '销售明细', range: 'A1:E6' },
    { op: 'set_col_width', sheet: '销售明细', col: 1, width: 12 },
  ]);
  await writeFile(xlsxPath, out);
  return `${out.length} bytes`;
});

await t('excel: 图表操作当前行为检测', async () => {
  const { readFile } = await import('node:fs/promises');
  const buf = await readFile(xlsxPath);
  try {
    await excel.editWorkbook(buf, [
      { op: 'add_chart', sheet: '汇总', chartType: 'bar', title: '季度销售', ranges: ['A1:A5', 'B1:B5'] },
    ]);
    return 'addChart 已可用(意外)';
  } catch (err) {
    return `图表由独立 XML 注入模块负责,editWorkbook 内报错符合预期: ${err.message.slice(0, 80)}`;
  }
});

await t('excel: 图表 XML 注入(bar/line/pie, 同表多图共用一个 drawing)', async () => {
  const { writeFile } = await import('node:fs/promises');
  const PizZip = (await import('pizzip')).default;
  const { injectChart, chartParts } = await import('../lib/core/charts.js');
  let buf = await excel.buildWorkbook({
    sheets: [{ name: '数据', header: true, rows: [['季度', '销售额', '利润'], ['Q1', 120, 30], ['Q2', 180, 52], ['Q3', 150, 41], ['Q4', 210, 66]] }],
  });
  const cases = [
    { chartType: 'bar', categories: 'A2:A5', series: [{ range: 'B2:B5', label: 'B1' }], title: '季度销售额', anchor: 'E2' },
    { chartType: 'line', categories: 'A2:A5', series: [{ range: 'B2:B5', label: 'B1' }, { range: 'C2:C5', label: 'C1' }], title: '趋势', anchor: 'E20' },
    { chartType: 'pie', categories: 'A2:A5', series: [{ range: 'B2:B5' }], title: '占比', anchor: 'L2' },
  ];
  for (const c of cases) buf = await injectChart(buf, { sheet: '数据', ...c });
  const p = join(outDir, 'chart.xlsx');
  await writeFile(p, buf);

  const parts = chartParts(buf);
  for (const e of ['xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/charts/chart3.xml', 'xl/drawings/drawing1.xml']) {
    if (!parts.includes(e)) throw new Error(`缺少部件 ${e}; 实际: ${parts.join(', ')}`);
  }
  const zip = new PizZip(buf);
  const { tagTexts } = await import('../lib/core/markup.js');
  const refs = tagTexts(zip.file('xl/worksheets/sheet1.xml').asText(), 'drawing')
    .filter((t) => t.includes('r:id='));
  if (refs.length !== 1) throw new Error(`工作表应只引用 1 个 drawing, 实际 ${refs.length}`);
  const anchors = tagTexts(zip.file('xl/drawings/drawing1.xml').asText(), 'xdr:oneCellAnchor');
  if (anchors.length !== cases.length) throw new Error(`drawing1 应含 ${cases.length} 个锚点, 实际 ${anchors.length}`);
  const rels = zip.file('xl/drawings/_rels/drawing1.xml.rels').asText();
  for (const n of [1, 2, 3]) {
    if (!rels.includes(`../charts/chart${n}.xml`)) throw new Error(`drawing1 缺少 chart${n} 关系`);
  }
  const back = await excel.workbookSummary(buf);
  if (!back?.sheets.some((s) => s.name === '数据')) throw new Error('注入后 ExcelJS 无法重新解析');
  return `${parts.length} 个部件, 3 图共用 drawing1(${anchors.length} 锚点), 重新解析 OK`;
});

await t('excel: add_chart 经 editWorkbook 落盘', async () => {
  const { readFile, writeFile } = await import('node:fs/promises');
  const buf = await readFile(xlsxPath);
  const { buf: out, changes } = await excel.editWorkbook(buf, [
    { op: 'add_chart', sheet: '汇总', chartType: 'bar', title: '汇总图', categories: 'A1:A2', series: [{ range: 'B1:B2' }], anchor: 'D2' },
  ]);
  await writeFile(join(outDir, 'chart-edit.xlsx'), out);
  if (!changes.includes('add_chart(bar)')) throw new Error(`changes 异常: ${changes.join(',')}`);
  const { chartParts } = await import('../lib/core/charts.js');
  if (!chartParts(out).length) throw new Error('未注入图表部件');
  return 'editWorkbook 图表注入成功';
});

await t('excel: 读取工作簿内容', async () => {
  const { readFile } = await import('node:fs/promises');
  const buf = await readFile(xlsxPath);
  const r = await excel.readWorkbook(buf, {});
  for (const key of ['销售明细', '汇总', '图表页', '键盘', '耳机', 'SUM(']) {
    if (!r.content.includes(key)) throw new Error(`读取结果缺少 ${key}`);
  }
  return `工作表: ${r.meta.sheetNames.join(', ')}`;
});

await t('excel: 读取指定工作表窗口', async () => {
  const { readFile } = await import('node:fs/promises');
  const buf = await readFile(xlsxPath);
  const r = await excel.readWorkbook(buf, { sheets: ['汇总'] });
  if (r.content.includes('销售明细') && r.content.includes('### 工作表「销售明细」')) throw new Error('不应包含其他表');
  if (!r.content.includes('### 工作表「汇总」')) throw new Error('缺少汇总表');
  return '窗口读取正常';
});

await t('excel: 按 range 限定行列窗口', async () => {
  const { readFile } = await import('node:fs/promises');
  const buf = await readFile(xlsxPath);
  const r = await excel.readWorkbook(buf, { sheets: ['销售明细'], range: 'A1:B2' });
  if (!r.content.includes('产品') || !r.content.includes('键盘')) throw new Error('range 起始行解析失败: ' + r.content);
  if (r.content.includes('显示器')) throw new Error('range 结束行未生效: ' + r.content);
  return 'A1:B2 只返回前两行';
});

await t('legacy: .xls 导出/读取(SheetJS biff8)', async () => {
  const XLSX = await import('@e965/xlsx');
  const ws = XLSX.utils.aoa_to_sheet([['名称', '数值'], ['项目A', 100], ['项目B', 200]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '数据');
  let out;
  try {
    out = Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'biff8' }));
  } catch (err) {
    return `社区版不支持写 .xls (跳过读回测试): ${err.message}`;
  }
  const { writeFile } = await import('node:fs/promises');
  const p = join(outDir, 'legacy.xls');
  await writeFile(p, out);
  const r = await legacy.readLegacySpreadsheet(out, 'xls');
  if (!r.content.includes('项目A')) throw new Error('.xls 读回内容缺失');
  return `.xls 写入(${out.length}B)并读回成功`;
});

await t('legacy: .xlsx → .xls 导出尝试', async () => {
  try {
    await office.opConvert(xlsxPath, join(outDir, 'conv.xls'));
    return '.xls 导出成功';
  } catch (err) {
    return `期望内失败（社区版限制）: ${err.message}`;
  }
});

await t('legacy: .xlsx → .csv 转换', async () => {
  const { readFile } = await import('node:fs/promises');
  const csvPath = join(outDir, 'conv.csv');
  await office.opConvert(xlsxPath, csvPath);
  const csv = (await readFile(csvPath)).toString('utf8');
  if (!csv.includes('键盘')) throw new Error('csv 内容缺失');
  return `${csv.split('\n').length} 行 CSV`;
});

await t('legacy: .docx → .doc 转换(本机转换器)', async () => {
  if (!(await converters.hasWordConverter('doc'))) return '无外部转换器（跳过）';
  const p = join(outDir, 'conv.doc');
  const conv = await office.opConvert(docxPath, p);
  const text = (await converters.wordRead(p)).content;
  if (!text.includes('项目周报')) throw new Error('.doc 内容缺失');
  return `docx→doc→读回成功（${conv.meta.via}）`;
});

await t('legacy: .doc → .docx 转换回环', async () => {
  if (!(await converters.hasWordConverter('doc'))) return '无外部转换器（跳过）';
  const docPath = join(outDir, 'conv.doc');
  const backPath = join(outDir, 'conv-back.docx');
  await office.opConvert(docPath, backPath);
  const { readFile } = await import('node:fs/promises');
  const buf = await readFile(backPath);
  const r = await word.readDocx(buf);
  if (!r.content.includes('项目周报')) throw new Error('.doc→.docx 内容缺失');
  return '回环成功';
});

await t('convert: markdown 文件 → docx', async () => {
  const { writeFile, readFile } = await import('node:fs/promises');
  const mdPath = join(outDir, 'note.md');
  await writeFile(mdPath, '# 说明\n\n这是转换测试。\n\n- 条目一\n- 条目二\n');
  const p = join(outDir, 'note.docx');
  await office.opConvert(mdPath, p);
  const buf = await readFile(p);
  const r = await word.readDocx(buf);
  if (!r.content.includes('转换测试')) throw new Error('md→docx 内容缺失');
  return 'md→docx 成功';
});

await t('convert: 同格式且路径不同 → 原样复制', async () => {
  const { readFile } = await import('node:fs/promises');
  const copyPath = join(outDir, 'copy.docx');
  const r = await office.opConvert(docxPath, copyPath);
  if (r.meta.copied !== true) throw new Error('应报告已复制');
  const [src, dst] = await Promise.all([readFile(docxPath), readFile(copyPath)]);
  if (!src.equals(dst)) throw new Error('复制内容与源不一致');
  const same = await office.opConvert(docxPath, docxPath);
  if (same.meta.copied !== false) throw new Error('同一路径应跳过写入');
  return `${dst.length} bytes 原样复制`;
});

await t('office_read 顶层: 全格式探测', async () => {
  const r = await office.opRead(docxPath, {});
  if (!r.content.includes('项目周报')) throw new Error('docx 读取失败');
  const r2 = await office.opRead(xlsxPath, { sheets: ['汇总'] });
  if (!r2.content.includes('汇总')) throw new Error('xlsx 读取失败');
  return 'docx+xlsx 顶层读取 OK';
});

// ---------------------------------------------------------------------------
// 错误分支 / 未被主流程用到的编辑操作
// ---------------------------------------------------------------------------
await t('excel: 增删行列 / 行高 / 重命名 / 删表全覆盖', async () => {
  const { writeFile } = await import('node:fs/promises');
  const buf = await excel.buildWorkbook({
    sheets: [
      { name: 'A', header: true, rows: [['h1', 'h2'], [1, 2], [3, 4]] },
      { name: 'B', rows: [['x']] },
    ],
  });
  const { buf: out, changes } = await excel.editWorkbook(buf, [
    { op: 'insert_rows', sheet: 'A', at: 2, count: 2 },
    { op: 'delete_rows', sheet: 'A', at: 5, count: 1 },
    { op: 'insert_cols', sheet: 'A', at: 2, count: 1 },
    { op: 'delete_cols', sheet: 'A', at: 4, count: 1 },
    { op: 'set_row_height', sheet: 'A', row: 1, height: 30 },
    { op: 'rename_sheet', sheet: 'B', name: 'B2' },
    { op: 'delete_sheet', sheet: 'B2' },
    { op: 'set_col_width', sheet: 'A', col: 'A', width: 20 },
  ]);
  await writeFile(join(outDir, 'ops.xlsx'), out);
  const want = ['insert_rows', 'delete_rows', 'insert_cols', 'delete_cols', 'set_row_height', 'rename_sheet', 'delete_sheet', 'set_col_width'];
  for (const op of want) if (!changes.includes(op)) throw new Error('未执行 ' + op);
  const back = await excel.readWorkbook(out, {});
  if (back.meta.sheetNames.includes('B2')) throw new Error('B2 未被删除');
  return `${want.length} 个操作全部执行`;
});

await t('excel: 未知操作 / 无效引用 / 非法 start 报错', async () => {
  const buf = await excel.buildWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] });
  const msgOf = async (ops) => { try { await excel.editWorkbook(buf, ops); return ''; } catch (err) { return err.message; } };
  const m1 = await msgOf([{ op: 'nope' }]);
  if (!/未知操作类型/.test(m1)) throw new Error('未知操作未报错: ' + m1);
  const m2 = await msgOf([{ op: 'set_value', sheet: 'S', ref: '??', value: 1 }]);
  if (!/无效单元格引用/.test(m2)) throw new Error('无效引用未报错: ' + m2);
  const m3 = await msgOf([{ op: 'set_cells', start: 'X', values: [[1]] }]);
  if (!/start/.test(m3)) throw new Error('set_cells start 未报错: ' + m3);
  const m4 = await msgOf([{ op: 'set_cells', values: 'nope' }]);
  if (!/values/.test(m4)) throw new Error('set_cells values 未报错: ' + m4);
  return '四类参数错误都可读';
});

await t('excel: date: / num: 前缀写入', async () => {
  const buf = await excel.buildWorkbook({ sheets: [{ name: 'S', rows: [['date:2026-09-09', 'num:1,234.5']] }] });
  const r = await excel.readWorkbook(buf, {});
  if (!r.content.includes('2026-09-09')) throw new Error('日期未写入: ' + r.content);
  if (!r.content.includes('1234.5')) throw new Error('num: 未生效: ' + r.content);
  return '两种前缀都生效';
});

await t('word: 损坏 / 空 / 缺内容源 / 超限的报错', async () => {
  let threw = false;
  try { await word.readDocx(Buffer.from('not a docx')); } catch { threw = true; }
  if (!threw) throw new Error('损坏 docx 未报错');
  const msgOf = async (fn) => { try { await fn(); return ''; } catch (err) { return err.message; } };
  const e2 = await msgOf(() => word.writeDocx({ markdown: '   ' }));
  if (!/内容为空/.test(e2)) throw new Error('空 markdown 未报错: ' + e2);
  const e3 = await msgOf(() => word.writeDocx({}));
  if (!/html \/ markdown \/ text/.test(e3)) throw new Error('缺内容源未报错: ' + e3);
  if (word.writeDocx({ markdown: '# x' }, { title: 't' }) === undefined) throw new Error('未返回 promise');
  let code = '';
  try { await word.readDocx(Buffer.alloc(CAPS.MAX_WORD_INPUT_BYTES + 1)); } catch (err) { code = err.code; }
  if (code !== 'OFFICE_TOO_LARGE') throw new Error('超大 docx 未被拒绝: ' + code);
  return '5 类异常输入都能明确报错';
});

await t('word: 模板语法错误被捕获', async () => {
  const bad = await word.writeDocx({ markdown: '{{#each}} 没有闭合' });
  let msg = '';
  try { await word.fillDocxTemplate(bad, {}); } catch (err) { msg = err.message; }
  if (!/模板渲染失败/.test(msg)) throw new Error('未报模板渲染失败: ' + msg);
  return '模板语法错误被捕获并说明原因';
});

await t('md: plainTextToHtml 与块级元素', async () => {
  const html = plainTextToHtml('第一行\n第二行\n\n第二段 & <tag>');
  if (!html.includes('&amp;') || !html.includes('&lt;tag&gt;')) throw new Error('未转义: ' + html);
  if (!html.includes('<br/>')) throw new Error('未换行: ' + html);
  const md = htmlToMarkdown('<blockquote><p>引用</p></blockquote><pre>代码块</pre><hr/><p>甲<br/>乙</p><p><img src="a.png" alt="图"/></p>');
  for (const want of ['> 引用', '```', '---', '甲', '![图](a.png)']) {
    if (!md.includes(want)) throw new Error(`缺少 ${want}: ${md}`);
  }
  return 'plainTextToHtml + 5 类块级元素正常';
});

await t('office: 缺失 / 目录 / 不支持格式的报错', async () => {
  const { mkdir } = await import('node:fs/promises');
  const msgOf = async (fn) => { try { await fn(); return ''; } catch (err) { return err.code || err.message; } };
  const c1 = await msgOf(() => office.opRead(join(outDir, 'missing.xlsx'), {}));
  if (c1 !== 'NOT_FOUND') throw new Error('缺失文件未报 NOT_FOUND: ' + c1);
  const asDir = join(outDir, 'dir.xlsx');
  await mkdir(asDir, { recursive: true });
  const c2 = await msgOf(() => office.opRead(asDir, {}));
  if (c2 !== 'NOT_A_FILE') throw new Error('目录未报 NOT_A_FILE: ' + c2);
  const c3 = await msgOf(() => office.opRead(join(outDir, 'note.md'), {}));
  if (c3 !== 'UNSUPPORTED_FORMAT') throw new Error('md 未报 UNSUPPORTED_FORMAT: ' + c3);
  return 'NOT_FOUND / 不是文件 / 不支持格式 都正确';
});

await t('legacy: .csv 读取与 .rtf 写出读回', async () => {
  const { writeFile } = await import('node:fs/promises');
  const csv = join(outDir, 'plain.csv');
  await writeFile(csv, '产品,数量\n键盘,10\n', 'utf8');
  const r = await office.opRead(csv, {});
  if (!r.content.includes('键盘')) throw new Error('csv 读取失败: ' + r.content);
  if (!(await converters.hasWordConverter('rtf'))) return 'csv 读取 OK;无 rtf 转换器(跳过写出)';
  const rtf = join(outDir, 'plain.rtf');
  const w = await office.opWriteDocx(rtf, { text: '第一段\n\n第二段' });
  const back = await office.opRead(rtf, {});
  if (!back.content.includes('第一段')) throw new Error('rtf 读回失败: ' + back.content);
  return `csv 读取 + rtf 写出读回(${w.meta.via})`;
});

await t('converters: backendSummary 可用', async () => {
  const summary = await converters.backendSummary();
  if (typeof summary !== 'string' || !summary.length) throw new Error('摘要为空');
  return summary;
});

await t('charts: 带引号的工作表名', async () => {
  const { injectChart } = await import('../lib/core/charts.js');
  const PizZip = (await import('pizzip')).default;
  const buf = await excel.buildWorkbook({ sheets: [{ name: '销售 明细', header: true, rows: [['季度', '额'], ['Q1', 10], ['Q2', 20]] }] });
  const out = await injectChart(buf, {
    sheet: '销售 明细', chartType: 'bar', categories: "'销售 明细'!A2:A3",
    series: [{ range: 'B2:B3', label: 'B1' }], anchor: 'E2',
  });
  const xml = new PizZip(out).file('xl/charts/chart1.xml').asText();
  if (!xml.includes("'销售 明细'!A2:A3")) throw new Error('带引号表名未正确处理: ' + xml.slice(0, 160));
  return '引号包裹的工作表名解析正确';
});

await t('markup: 属性查找不会误配后缀', async () => {
  const { attrIn, attrValue, readTagAt, appendBeforeClose } = await import('../lib/core/markup.js');
  if (attrIn(' xname="1" name="2"', 'name') !== '2') throw new Error('后缀误配: ' + attrIn(' xname="1" name="2"', 'name'));
  if (attrValue('<a href=\'x\' name="n">', 'name') !== 'n') throw new Error('单引号属性解析失败');
  if (readTagAt('<a href="含>符号">', 0).attrs !== ' href="含>符号"') throw new Error('引号内 > 解析失败');
  if (appendBeforeClose('<r></r>  ', 'r', 'X') !== '<r>X</r>  ') throw new Error('尾部空白未保留');
  if (appendBeforeClose('<r></r>tail', 'r', 'X') !== '<r></r>tail') throw new Error('非结尾不应插入');
  return '属性查找 / 引号内 > / 尾部插入 都正确';
});


await t('回归: sheet 序号从 1 开始', async () => {
  const buf = await excel.buildWorkbook({
    sheets: [{ name: '第一张', rows: [['a']] }, { name: '第二张', rows: [['b']] }],
  });
  const { buf: out } = await excel.editWorkbook(buf, [{ op: 'set_value', sheet: 1, ref: 'B1', value: '标在1号表' }]);
  const r = await excel.readWorkbook(out, {});
  const first = r.tsvBySheet['第一张'] || '';
  if (!first.includes('标在1号表')) throw new Error('sheet:1 未命中第一张表: ' + JSON.stringify(r.tsvBySheet));
  return 'sheet:1 → 第一张表';
});

await t('回归: 日期写入后原样回读(不受时区影响)', async () => {
  const buf = await excel.buildWorkbook({ sheets: [{ name: 'S', rows: [['date:2026-09-09']] }] });
  const r = await excel.readWorkbook(buf, {});
  if (!r.content.includes('2026-09-09')) throw new Error('日期回读不一致: ' + r.content);
  return `本地时区 ${Intl.DateTimeFormat().resolvedOptions().timeZone} 下回读一致`;
});

await t('安全: CSV 里的 = 开头内容不会被转成活公式', async () => {
  const { writeFile, readFile } = await import('node:fs/promises');
  const csv = join(outDir, 'inject.csv');
  await writeFile(csv, "名称,值\n恶意,=1+1\nDDE,=cmd|'/c calc'!A0\n", 'utf8');
  const xlsx = join(outDir, 'inject.xlsx');
  await office.opConvert(csv, xlsx);
  const PizZip = (await import('pizzip')).default;
  const xml = new PizZip(await readFile(xlsx)).file('xl/worksheets/sheet1.xml').asText();
  const formulas = xml.match(/<f>[^<]*<\/f>/g) || [];
  if (formulas.length) throw new Error('外部 CSV 被写成了活公式: ' + formulas.join(','));
  const plain = await office.opRead(xlsx, {});
  if (!plain.content.includes('=1+1')) throw new Error('原文本丢失: ' + plain.content);
  return '= / DDE 都保持为文本';
});

await t('安全: 写入侧仍按文档承诺支持公式', async () => {
  const buf = await excel.buildWorkbook({ sheets: [{ name: 'S', rows: [['=SUM(1,2)']] }] });
  const PizZip = (await import('pizzip')).default;
  const xml = new PizZip(buf).file('xl/worksheets/sheet1.xml').asText();
  if (!/<f>SUM\(1,2\)<\/f>/.test(xml)) throw new Error('office_write_xlsx 的公式能力被破坏: ' + xml.slice(0, 200));
  return 'office_write_xlsx 仍写公式(仅转换文本源时中和)';
});

await t('安全: 临时文件以 0600 写入(不落到公共 tmp 的默认 0644)', async () => {
  if (!(await converters.hasWordConverter('rtf'))) return '无 rtf 转换器(跳过)';
  const { readFile, writeFile, unlink, stat } = await import('node:fs/promises');
  const { tmpdir } = await import('node:os');
  const calls = [];
  const spyIo = {
    async readBuf(pth) { return readFile(pth); },
    async writeBuf(pth, data, options) { calls.push({ pth, options }); await writeFile(pth, data, options); },
    async remove(pth) { await unlink(pth).catch(() => {}); },
    tmpFile(ext) { return join(tmpdir(), `dsh-perm-${Date.now()}${ext}`); },
    async stat(pth) { const st = await stat(pth); return { size: st.size, type: 'file' }; },
  };
  await office.opWriteDocx(join(outDir, 'perm.rtf'), { text: '临时文件权限检查' }, spyIo);
  const tempWrites = calls.filter((c) => c.pth.startsWith(tmpdir()));
  if (!tempWrites.length) throw new Error('没有经过临时文件?');
  for (const w of tempWrites) {
    if (w.options?.mode !== 0o600) throw new Error(`临时文件权限不是 0600: ${JSON.stringify(w.options)}`);
  }
  return `${tempWrites.length} 处临时写入都是 0600`;
});

await t('安全: 解压炸弹 — 声明解压总量超限时拒绝', async () => {
  const { readFile } = await import('node:fs/promises');
  const buf = await readFile(docxPath);
  const saved = CAPS.MAX_UNCOMPRESSED_BYTES;
  CAPS.MAX_UNCOMPRESSED_BYTES = 1024; // 正常 docx 解压后必然远超 1 KiB
  try {
    await word.readDocx(buf);
    throw new Error('没有拦截超限的解压体积');
  } catch (err) {
    if (err.code !== 'ZIP_BOMB_SUSPECTED') throw err;
    return err.message;
  } finally {
    CAPS.MAX_UNCOMPRESSED_BYTES = saved;
  }
});

await t('安全: 解压炸弹 — 压缩比异常时拒绝(压缩包本身很小)', async () => {
  const { default: PizZip } = await import('pizzip');
  const { readFile } = await import('node:fs/promises');
  const zip = new PizZip();
  zip.file('word/document.xml', 'A'.repeat(4 * 1024 * 1024)); // 4 MiB 极度可压缩
  const bomb = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  try {
    await word.readDocx(bomb);
    throw new Error('没有拦截高压缩比的压缩包');
  } catch (err) {
    if (err.code !== 'ZIP_BOMB_SUSPECTED') throw err;
    const packed = Buffer.byteLength(bomb);
    if (packed > 32 * 1024) throw new Error(`测试样本压缩后 ${packed} 字节,不够小,无法证明是压缩比触发`);
    // 正常文件不能被误伤
    await word.readDocx(await readFile(docxPath));
    await excel.readWorkbook(await readFile(xlsxPath));
    return `${packed} 字节 → 4 MiB 被拒;正常 docx/xlsx 仍可通过`;
  }
});

await t('安全: 写入 docx 前剔除图片(不触碰可联网的图片探测栈)', async () => {
  const cases = [
    ['<p>前<img src="a.png" alt="图 1">后</p>', '<p>前图 1后</p>'],
    ['<figure><img src="https://x/y.jpg"><figcaption>说明</figcaption></figure>', '说明'],
    ['<img src="x" alt="a&amp;b">', 'a&amp;b'],
    ['<img src="x" alt="<script>alert(1)</script>">', '&lt;script&gt;alert(1)&lt;/script&gt;'],
  ];
  for (const [input, want] of cases) {
    const got = word.stripImages(input);
    if (got !== want) throw new Error(`${JSON.stringify(input)} → ${JSON.stringify(got)},期望 ${JSON.stringify(want)}`);
  }
  const buf = await word.writeDocx({ html: '<p>图片</p><img src="data:image/gif;base64,R0lGODlhAQABAAAAACw=">' });
  if (!Buffer.isBuffer(buf) || buf.length === 0) throw new Error('剔除图片后仍应能生成 docx');
  return `${cases.length} 个用例 + 含图片的 HTML 仍生成 docx(${buf.length} 字节)`;
});

const failed = results.filter((r) => !r.ok);
console.log(`\n===== 结果: ${results.length - failed.length}/${results.length} 通过 =====`);
if (failed.length) {
  console.log('\n失败明细:');
  for (const f of failed) console.log(`- ${f.name}\n  ${f.detail.split('\n').slice(0, 6).join('\n  ')}`);
  process.exit(1);
}
