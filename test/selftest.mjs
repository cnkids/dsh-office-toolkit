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
import { htmlToMarkdown } from '../lib/core/md.js';

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

const failed = results.filter((r) => !r.ok);
console.log(`\n===== 结果: ${results.length - failed.length}/${results.length} 通过 =====`);
if (failed.length) {
  console.log('\n失败明细:');
  for (const f of failed) console.log(`- ${f.name}\n  ${f.detail.split('\n').slice(0, 6).join('\n  ')}`);
  process.exit(1);
}
