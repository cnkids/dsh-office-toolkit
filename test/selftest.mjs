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
import { htmlToDocxBuffer } from '../lib/core/docx-writer.js';
import { resolveMainPart } from '../lib/core/word.js';
import { extractDocxFormat, formatReport } from '../lib/core/docx-format.js';
import { CAPS, assertOfficeBinary } from '../lib/core/util.js';
import { CORE_DEPS, depFailure, lazyModule, missingDeps } from '../lib/core/deps.js';

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

// ---------------------------------------------------------------------------
// office_query:在文件内直接算(大表只回结论,不把行搬进上下文)
// ---------------------------------------------------------------------------
/** 造一份 12000 行的订单表,验证"整表计算"这条路径。 */
async function writeBigXlsx() {
  const { writeFile } = await import('node:fs/promises');
  const regions = ['华东', '华南', '华北', '西南'];
  const rows = [['地区', '产品', '金额', '日期']];
  for (let i = 0; i < 12000; i++) {
    rows.push([regions[i % 4], `P${i % 5}`, `${(i % 997) + 0.5}`, `2025/${(i % 12) + 1}/1`]);
  }
  const path = join(outDir, 'orders.xlsx');
  await writeFile(path, await excel.buildWorkbook({ sheets: [{ name: '订单', rows, header: true }, { name: '备注', rows: [['说明'], ['测试']] }] }));
  return path;
}

await t('query: 12000 行分组聚合(ExcelJS 路径)', async () => {
  const path = await writeBigXlsx();
  const r = await office.opQuery(path, {
    groupBy: ['地区'],
    aggregate: [{ col: '金额', fn: 'sum', as: '销售额' }, { col: '金额', fn: 'avg' }, { col: '产品', fn: 'countDistinct', as: '产品数' }],
    orderBy: [{ col: '销售额', dir: 'desc' }],
  });
  for (const key of ['销售额', '产品数', '华东', '西南']) {
    if (!r.content.includes(key)) throw new Error(`结果缺少 ${key}: ${r.content}`);
  }
  if (!r.content.includes('12000 行')) throw new Error('摘要没有报出扫描行数: ' + r.content);
  if (r.meta.scannedRows !== 12000 || r.meta.matchedRows !== 12000) throw new Error('meta 行数不对: ' + JSON.stringify(r.meta));
  if (r.content.length > 1200) throw new Error('结果没有收敛,返回了过多内容: ' + r.content.length);
  // 降序:西南(最大)必须排在华东之前
  const order = r.content.split('```tsv')[1].split('```')[0].trim().split('\n').slice(1).map((l) => l.split('\t')[0]);
  if (order[0] !== '西南' || order.length !== 4) throw new Error('排序或分组数不对: ' + order.join(','));
  return `12000 行 → 4 行结论(${r.content.length} 字符)`;
});

await t('query: 表头不在第一行 + 条件筛选', async () => {
  const path = await writeBigXlsx();
  const r = await office.opQuery(path, {
    sheet: '订单',
    headerRow: 1,
    where: [{ col: '地区', op: 'eq', value: '华东' }, { col: '金额', op: 'gte', value: 500 }],
    aggregate: [{ col: '金额', fn: 'count', as: '单数' }, { col: '金额', fn: 'max', as: '最大' }],
  });
  if (!/命中 \d+ 行/.test(r.content)) throw new Error('摘要没写命中行数: ' + r.content.split('\n')[0]);
  if (!r.meta.matchedRows || r.meta.matchedRows >= 12000) throw new Error('筛选没生效: ' + r.meta.matchedRows);
  return `命中 ${r.meta.matchedRows} 行`;
});

await t('query: 不给汇总条件时返回表结构画像', async () => {
  const path = await writeBigXlsx();
  const r = await office.opQuery(path, { sheet: 2 });
  if (!r.content.includes('表结构画像')) throw new Error('没有进入画像模式: ' + r.content);
  if (!r.content.includes('备注')) throw new Error('序号选表没生效(应为第 2 张表)');
  return r.meta.sheet;
});

await t('query: CSV 走 SheetJS 路径', async () => {
  const { writeFile } = await import('node:fs/promises');
  const path = join(outDir, 'query.csv');
  await writeFile(path, '地区,金额\n华东,"1,200.50"\n华南,800\n华东,200\n');
  const r = await office.opQuery(path, { groupBy: ['地区'], aggregate: [{ col: '金额', fn: 'sum', as: '合计' }], orderBy: [{ col: '合计', dir: 'desc' }] });
  if (!r.content.includes('华东') || !r.content.includes('1400.5')) throw new Error('CSV 聚合结果不对: ' + r.content);
  return '华东 1400.5 / 华南 800';
});

await t('query: 选错工作表 / 算了 Word 文档时报错清晰', async () => {
  const path = await writeBigXlsx();
  const sheetErr = await office.opQuery(path, { sheet: '不存在的表' }).catch((e) => e);
  if (sheetErr.code !== 'UNKNOWN_SHEET' || !sheetErr.message.includes('订单')) throw new Error('选表报错不对: ' + sheetErr.message);
  const { writeFile } = await import('node:fs/promises');
  const docPath = join(outDir, 'query-not-table.docx');
  await writeFile(docPath, await word.writeDocx({ markdown: '# 不是表格' }, {}));
  const wordErr = await office.opQuery(docPath, {}).catch((e) => e);
  if (!/只能算表格/.test(wordErr.message) || !/office_read/.test(wordErr.message)) throw new Error('算 Word 的报错不对: ' + wordErr.message);
  return sheetErr.message.slice(0, 40);
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

await t('安全: 写入 docx 不解析图片(只留 alt 文本,不联网)', async () => {
  const buf = await word.writeDocx({
    html: '<p>正文</p><img src="https://example.com/x.jpg" alt="图 1">',
  });
  if (!Buffer.isBuffer(buf) || buf.length === 0) throw new Error('含图片的 HTML 仍应生成 docx');
  const back = await word.readDocx(buf);
  if (!back.content.includes('正文')) throw new Error('正文丢失');
  if (!back.content.includes('图 1')) throw new Error('alt 文本未保留');
  const xml = new (await import('pizzip')).default(buf).file('word/document.xml').asText();
  if (xml.includes('graphic') || xml.includes('pic:pic')) throw new Error('docx 里出现了图片部件');
  return `图片被跳过,alt 保留(${buf.length} 字节)`;
});

await t('docx-writer: 标题 / 内联 / 列表 / 表格 / 引用 / 代码 / 链接 / 分隔线', async () => {
  const html = [
    '<h1>一级标题</h1><h3>三级标题</h3>',
    '<p><strong>粗</strong><em>斜</em><u>下划</u><s>删</s><code>码</code><sub>下</sub><sup>上</sup></p>',
    '<p style="color:#f00;background-color:#EEE;font-size:20px;font-weight:bold;font-style:italic;text-decoration:underline line-through;text-align:center">样式</p>',
    '<ul><li>项目一<ul><li>子项</li></ul></li></ul><ol><li>第一</li></ol>',
    '<table><thead><tr><th>表头</th></tr></thead><tbody><tr><td><p>甲</p><p>乙</p></td></tr></tbody></table>',
    '<blockquote>引用</blockquote>',
    '<pre>行一\n行二</pre>',
    '<hr>',
    '<p><a href="https://example.com/x">正常链接</a></p>',
    '<custom>未知标签</custom>',
  ].join('');
  const buf = await htmlToDocxBuffer(html);
  const zip = new (await import('pizzip')).default(buf);
  const xml = zip.file('word/document.xml').asText();
  const numbering = zip.file('word/numbering.xml');
  const rels = zip.file('word/_rels/document.xml.rels').asText();
  const markers = [
    ['Heading1 样式', xml.includes('Heading1')],
    ['Heading3 样式', xml.includes('Heading3')],
    ['粗体', xml.includes('<w:b/>')],
    ['斜体', xml.includes('<w:i/>')],
    ['下划线', xml.includes('<w:u ')],
    ['删除线', xml.includes('<w:strike/>')],
    ['等宽字体', xml.includes('Consolas')],
    ['下标', xml.includes('subscript')],
    ['上标', xml.includes('superscript')],
    ['前景色 f00→ff0000', xml.includes('ff0000')],
    ['底纹 #EEE→eeeeee', xml.includes('eeeeee')],
    ['20px→30 半磅', xml.includes('w:sz w:val="30"')],
    ['居中', xml.includes('w:jc w:val="center"')],
    ['项目符号编号', xml.includes('<w:numPr>')],
    ['表格', xml.includes('<w:tbl>')],
    ['表头底纹', xml.includes('F2F2F2')],
    ['引用缩进', xml.includes('w:ind ')],
    ['分隔线边框', xml.includes('<w:pBdr>')],
    ['软换行', xml.includes('<w:br/>')],
    ['超链接', xml.includes('w:hyperlink')],
    ['链接写入 rels', rels.includes('https://example.com/x')],
  ];
  const bad = markers.filter(([, ok]) => !ok).map(([name]) => name);
  if (bad.length) throw new Error('缺失: ' + bad.join(', '));
  if (!numbering) throw new Error('缺少 numbering.xml');
  if (!numbering.asText().includes('w:val="bullet"')) throw new Error('缺少无序列表编号定义');
  if (!numbering.asText().includes('w:val="decimal"')) throw new Error('缺少有序列表编号定义');
  const text = (await word.readDocx(buf)).content;
  for (const needle of ['一级标题', '三级标题', '粗', '斜', '下划', '删', '码', '样式', '项目一', '子项', '第一', '表头', '甲', '乙', '引用', '行一', '行二', '正常链接', '未知标签']) {
    if (!text.includes(needle)) throw new Error(`回读文本缺少: ${needle}`);
  }
  return `${markers.length} 项标记 + 19 个文本回读全部通过`;
});

await t('docx-writer: 横向纸张与页边距生效', async () => {
  const pgOf = async (opts) => {
    const buf = await htmlToDocxBuffer('<p>x</p>', opts);
    return new (await import('pizzip')).default(buf).file('word/document.xml').asText();
  };
  const portrait = await pgOf({});
  const landscape = await pgOf({ landscape: true });
  const margin = await pgOf({ marginsMm: 40 });
  if (!portrait.includes('w:orient="portrait"')) throw new Error('默认不是纵向');
  if (!landscape.includes('w:orient="landscape"')) throw new Error('横向未生效');
  if (!landscape.includes('w:w="16837"') || !landscape.includes('w:h="11905"')) throw new Error('横向未交换宽高: ' + landscape.slice(0, 200));
  if (!portrait.includes('w:w="11905"') || !portrait.includes('w:h="16837"')) throw new Error('纵向尺寸不是 A4');
  if (!margin.includes('w:top="2267"')) throw new Error('40mm 页边距未生效');
  if (!portrait.includes('w:top="1417"')) throw new Error('默认 25mm 页边距不对');
  return 'A4 纵向 / 横向 / 页边距均正确';
});

await t('安全: docx 里的 javascript: 链接降级为纯文本', async () => {
  const buf = await htmlToDocxBuffer('<p><a href="javascript:alert(1)">点我</a><a href="https://ok.example/">正常</a></p>');
  const zip = new (await import('pizzip')).default(buf);
  const xml = zip.file('word/document.xml').asText();
  const rels = zip.file('word/_rels/document.xml.rels').asText();
  if (xml.includes('javascript') || rels.includes('javascript')) throw new Error('危险链接写进了文档');
  if (!xml.includes('点我')) throw new Error('链接文字丢失');
  if (!rels.includes('https://ok.example/')) throw new Error('正常链接被误伤');
  return 'javascript: 已丢弃,http(s) 保留';
});

await t('依赖: 缺包时报错点明包名、用途与修复命令', async () => {
  const raw = Object.assign(new Error("Cannot find package 'mammoth' imported from /x/y.js"), { code: 'ERR_MODULE_NOT_FOUND' });
  const err = depFailure(raw, 'Word 读取');
  if (err.code !== 'MISSING_DEPENDENCY') throw new Error('错误码不对: ' + err.code);
  if (!err.message.includes('mammoth')) throw new Error('未点明包名: ' + err.message);
  if (!err.message.includes('读取 .docx')) throw new Error('未说明该依赖的用途');
  if (!err.message.includes('dsh plugin --profile web add')) throw new Error('未给出修复命令');
  if (!err.message.includes("Cannot find package 'mammoth'")) throw new Error('未保留原始报错');
  const other = new Error('boom');
  if (depFailure(other, 'x') !== other) throw new Error('普通错误被改写');
  const mod = lazyModule('dsh-这个包不存在-测试用', '测试', ['whatever']);
  try {
    await mod.whatever();
    throw new Error('懒加载没有抛错');
  } catch (e) {
    if (e.code !== 'MISSING_DEPENDENCY') throw e;
  }
  return '包名 / 用途 / 修复命令 / 原样透传 / 懒加载 均正确';
});

await t('依赖: 加载时自检能列出缺失依赖(本机应为空)', async () => {
  const missing = missingDeps();
  if (missing.length) throw new Error('本机缺依赖: ' + missing.join(', '));
  if (CORE_DEPS.length < 8) throw new Error('依赖清单不完整');
  return `${CORE_DEPS.length} 个运行时依赖全部可解析`;
});

await t('安全: 二进制输出自检拒绝文本冒充 .xlsx/.docx/.odt', async () => {
  const fakes = [
    ['xlsx', Buffer.from('产品\t数量\n键盘\t10\n', 'utf8')],
    ['docx', Buffer.from('<html><body>x</body></html>', 'utf8')],
    ['odt', Buffer.from('PK 但后面不是 zip', 'utf8')],
  ];
  for (const [ext, buf] of fakes) {
    try {
      assertOfficeBinary(buf, ext);
      throw new Error(`未拦截文本冒充 .${ext}`);
    } catch (e) {
      if (e.code !== 'BAD_OUTPUT_FORMAT') throw e;
    }
  }
  assertOfficeBinary(await excel.buildWorkbook({ sheets: [{ rows: [['a', 1]] }] }), 'xlsx');
  assertOfficeBinary(await word.writeDocx({ text: '正常' }), 'docx');
  return `${fakes.length} 类文本冒充被拒,真 OOXML 通过`;
});

await t('回归: 表格转 .tsv 用制表符(此前写成逗号分隔)', async () => {
  const { readFile, writeFile } = await import('node:fs/promises');
  const csv = join(outDir, 'sep.csv');
  const tsv = join(outDir, 'sep.tsv');
  await writeFile(csv, Buffer.from('产品,数量\n键盘,10\n', 'utf8'));
  await office.opConvert(csv, tsv);
  const text = await readFile(tsv, 'utf8');
  if (!text.includes('\t')) throw new Error('没有制表符: ' + JSON.stringify(text.slice(0, 60)));
  if (text.includes(',')) throw new Error('仍是逗号分隔: ' + JSON.stringify(text.slice(0, 60)));
  // 逗号分隔的 csv 仍应是逗号
  const back = join(outDir, 'sep-back.csv');
  await office.opConvert(csv, back);
  if (!(await readFile(back, 'utf8')).includes(',')) throw new Error('csv 反而不是逗号了');
  return JSON.stringify(text.split('\n')[0]);
});

await t('格式: 提取字体/字号/行距/缩进(样式继承 + 直接格式优先 + run 覆盖)', async () => {
  const styles = `<w:styles>
    <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="等线"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:styleId="Body"><w:name w:val="正文"/><w:pPr><w:spacing w:line="576" w:lineRule="exact"/></w:pPr><w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/><w:sz w:val="32"/></w:rPr></w:style>
    <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="标题"/><w:basedOn w:val="Body"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:rFonts w:eastAsia="方正小标宋简体"/><w:sz w:val="44"/><w:b/></w:rPr></w:style>
  </w:styles>`;
  const doc = `<w:document><w:body>
    <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>关于印发某某办法的通知</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Body"/><w:ind w:firstLineChars="200"/></w:pPr><w:r><w:t>正文段落</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Body"/><w:spacing w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="200"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="28"/></w:rPr><w:t>小标题</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2098" w:bottom="1984" w:left="1587" w:right="1474"/></w:sectPr>
  </w:body></w:document>`;
  const f = extractDocxFormat(doc, styles);
  const [title, body, override] = f.paragraphs;
  const bad = [];
  if (title.fontEastAsia !== '方正小标宋简体' || title.sizePt !== 22 || title.bold !== true) bad.push('标题 run 格式(继承自 basedOn 链)');
  if (title.alignment !== '居中') bad.push('标题居中');
  if (body.fontEastAsia !== '仿宋_GB2312' || body.sizePt !== 16) bad.push('正文字体(来自样式)');
  if (body.lineSpacing?.rule !== 'exact' || body.lineSpacing.pt !== 28.8) bad.push('固定值行距 28.8pt');
  if (body.firstLineChars !== 2) bad.push('首行缩进 2 字符');
  if (override.fontEastAsia !== '黑体' || override.sizePt !== 14) bad.push('直接 run 覆盖');
  if (override.lineSpacing?.rule !== 'auto' || override.lineSpacing.lines !== 1.5) bad.push('段落直接改行距为 1.5 倍');
  if (f.page?.marginsMm?.top !== 37 || f.page?.widthMm !== 210) bad.push('页面与页边距');
  if (bad.length) throw new Error('提取错误: ' + bad.join('、'));
  if (!formatReport(f).includes('偏离主流格式的段落')) throw new Error('报告缺少偏离检查');
  return `${f.paragraphs.length} 段 + 页面,9 项断言通过`;
});

await t('格式: office_read(withFormatting) 返回报告与 meta.formatting', async () => {
  const { writeFile } = await import('node:fs/promises');
  const PizZip = (await import('pizzip')).default;
  const base = new PizZip(await htmlToDocxBuffer('<p>占位</p>'));
  const zip = new PizZip(base.generate({ type: 'nodebuffer' }));
  zip.file('word/styles.xml', `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="等线"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:styleId="Body"><w:name w:val="正文"/><w:pPr><w:spacing w:line="576" w:lineRule="exact"/></w:pPr><w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/><w:sz w:val="32"/></w:rPr></w:style></w:styles>`);
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:p><w:pPr><w:pStyle w:val="Body"/><w:ind w:firstLineChars="200"/></w:pPr><w:r><w:t>格式比对样例</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2098" w:bottom="1984" w:left="1587" w:right="1474"/></w:sectPr>
  </w:body></w:document>`);
  const p = join(outDir, 'format-check.docx');
  await writeFile(p, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  const r = await office.opRead(p, { withFormatting: true });
  if (!r.content.includes('格式报告')) throw new Error('正文没有格式报告');
  if (!r.content.includes('仿宋_GB2312') || !r.content.includes('固定值 28.8pt')) throw new Error('报告缺少字体/行距');
  const fm = r.meta.formatting;
  if (fm?.paragraphs.length !== 1) throw new Error('meta.formatting 缺失');
  if (fm.paragraphs[0].fontEastAsia !== '仿宋_GB2312' || fm.paragraphs[0].sizePt !== 16) throw new Error('结构化数据不对');
  if (fm.page.marginsMm.top !== 37) throw new Error('页边距缺失');
  if (fm.truncated !== false) throw new Error('截断标记不对');
  // 不传参数时不应出现格式报告
  const plain = await office.opRead(p, {});
  if (plain.content.includes('格式报告')) throw new Error('未传 withFormatting 却返回了格式报告');
  return `报告 + meta(${fm.paragraphs.length} 段 / 页边距 ${fm.page.marginsMm.top}mm) 正确`;
});

await t('格式: 主题字体与隐式默认样式(Word 默认模板的真实写法)', async () => {
  // Word 模板的典型写法:w:default="1" 的段落样式 + 主题字体引用
  const styles = `<w:styles>
    <w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/>
      <w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>
      <w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/><w:sz w:val="24"/></w:rPr></w:style>
    <w:style w:type="character" w:default="1" w:styleId="4"><w:name w:val="Default Paragraph Font"/></w:style>
  </w:styles>`;
  const theme = `<a:theme><a:themeElements><a:fontScheme>
    <a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:font script="Hans" typeface="黑体"/></a:majorFont>
    <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:font script="Hans" typeface="宋体"/></a:minorFont>
  </a:fontScheme></a:themeElements></a:theme>`;
  // 注意:段落没有 w:pStyle,靠 w:default="1" 隐式套用
  const doc = `<w:document><w:body><w:p><w:r><w:t>没有显式样式的段落</w:t></w:r></w:p></w:body></w:document>`;
  const f = extractDocxFormat(doc, styles, theme);
  const p = f.paragraphs[0];
  if (p.styleName !== 'Normal') throw new Error('隐式默认样式未生效: ' + p.styleName);
  if (p.fontAscii !== 'Calibri') throw new Error('西文主题字体未解析: ' + p.fontAscii);
  if (p.fontEastAsia !== '宋体') throw new Error('中文主题字体未解析(应取 script="Hans"): ' + p.fontEastAsia);
  if (p.sizePt !== 12) throw new Error('字号未继承默认样式: ' + p.sizePt);
  if (p.lineSpacing?.lines !== 1.5) throw new Error('行距未继承默认样式: ' + JSON.stringify(p.lineSpacing));
  // 没有主题文件时不应崩,且不回退成 undefined
  const noTheme = extractDocxFormat(doc, styles);
  if (noTheme.paragraphs[0].styleName !== 'Normal') throw new Error('无主题时默认样式仍应生效');
  return '默认样式 + Calibri/宋体(script=Hans) + 12pt + 1.5 倍行距 均正确';
});

// ---- 读取兼容性(真实文件里常见的各种「不标准」) ----
await t('兼容: 非标准 zip(条目名带反斜杠)也能读并提取格式', async () => {
  const PizZip = (await import('pizzip')).default;
  const { writeFile } = await import('node:fs/promises');
  const good = await word.writeDocx({ markdown: '# 标题\n\n正文内容' });
  const out = new PizZip();
  for (const [name, entry] of Object.entries(new PizZip(good).files)) {
    if (!entry.dir) out.file(name.replaceAll('/', '\\'), entry.asNodeBuffer());
  }
  const p = join(outDir, 'backslash.docx');
  await writeFile(p, out.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  const r = await office.opRead(p, { withFormatting: true });
  if (!String(r.content).includes('正文内容')) throw new Error('正文没读出来');
  if (!r.meta.formatting?.paragraphs?.length) throw new Error('格式未提取');
  if (!String(r.content).includes('自动修正')) throw new Error('未提示自动修复了条目名');
  return '反斜杠条目名:自动修复后可正常读取与提取格式';
});

await t('兼容: 改过后缀的文件按真实内容读取', async () => {
  const { writeFile } = await import('node:fs/promises');
  // 1) docx 内容,名字却是 .doc
  const docx = await word.writeDocx({ markdown: '# 真身是 docx' });
  const fakeDoc = join(outDir, 'really-docx.doc');
  await writeFile(fakeDoc, docx);
  const r1 = await office.opRead(fakeDoc, {});
  if (!String(r1.content).includes('真身是 docx')) throw new Error('未按 docx 内容读取');
  if (!String(r1.content).includes('文件内容其实是 .docx')) throw new Error('未提示实际格式');
  // 2) CSV 内容,名字却是 .xlsx
  const fakeXlsx = join(outDir, 'really-csv.xlsx');
  await writeFile(fakeXlsx, Buffer.from('产品,数量\n键盘,10\n', 'utf8'));
  const r2 = await office.opRead(fakeXlsx, {});
  if (!String(r2.content).includes('键盘')) throw new Error('未按 CSV 内容读取');
  if (!String(r2.content).includes('分隔符文本')) throw new Error('未提示分隔符文本');
  // 3) 制表符文本,名字是 .xlsx → 按 tsv
  const fakeTsv = join(outDir, 'really-tsv.xlsx');
  await writeFile(fakeTsv, Buffer.from('产品\t数量\n键盘\t10\n', 'utf8'));
  const r3 = await office.opRead(fakeTsv, {});
  if (!String(r3.content).includes('TSV')) throw new Error('制表符文本未识别为 TSV');
  return 'docx→.doc / CSV→.xlsx / TSV→.xlsx 均按真实内容读取';
});

await t('兼容: OLE 老格式伪装成 .docx 时按 .doc 处理', async () => {
  const { writeFile } = await import('node:fs/promises');
  const ole = Buffer.concat([
    Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
    Buffer.alloc(4096),
  ]);
  const p = join(outDir, 'really-doc.docx');
  await writeFile(p, ole);
  try {
    await office.opRead(p, {});
    throw new Error('内容无效却没报错');
  } catch (err) {
    if (/不是有效的 docx/.test(err.message)) throw new Error('仍按 docx 解压,未按 OLE 处理: ' + err.message);
    if (err.code === 'BAD_CONTAINER') throw new Error('不应报 BAD_CONTAINER');
  }
  return 'OLE 头被识别为老格式,不再当作 docx 解压';
});

await t('兼容: zip 里没有 Office 主文档时报错可读(列出实际条目)', async () => {
  const { writeFile } = await import('node:fs/promises');
  const PizZip = (await import('pizzip')).default;
  const zip = new PizZip();
  zip.file('hello.txt', 'not office');
  zip.file('[Content_Types].xml', '<Types/>');
  const p = join(outDir, 'not-office.docx');
  await writeFile(p, zip.generate({ type: 'nodebuffer' }));
  try {
    await office.opRead(p, {});
    throw new Error('损坏文件却读成功了');
  } catch (err) {
    if (err.code !== 'BAD_CONTAINER') throw err;
    if (!err.message.includes('hello.txt')) throw new Error('未列出实际条目: ' + err.message);
  }
  return '报 BAD_CONTAINER 并列出实际条目,便于判断文件真身';
});

await t('契约: 调用方参数被冻结时所有操作都不得出错(DSH 宿主如此传参)', async () => {
  const { writeFile } = await import('node:fs/promises');
  const F = Object.freeze;
  // 非标准 zip 会走「加注」分支 —— v0.3.17 正是在这里往冻结对象上写属性而崩溃
  const PizZip = (await import('pizzip')).default;
  const out = new PizZip();
  for (const [name, entry] of Object.entries(new PizZip(await word.writeDocx({ markdown: '# 冻结测试' })).files)) {
    if (!entry.dir) out.file(name.replaceAll('/', '\\'), entry.asNodeBuffer());
  }
  const weird = join(outDir, 'frozen-weird.docx');
  await writeFile(weird, out.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  const plain = join(outDir, 'frozen-plain.xlsx');
  await office.opWriteXlsx(plain, { sheets: [{ rows: [['a', 1]] }] });

  const calls = [
    ['opRead(带注的非标准 zip)', () => office.opRead(weird, F({}), F({}))],
    ['opRead(普通文件)', () => office.opRead(plain, F({}), F({}))],
    ['opWriteXlsx', () => office.opWriteXlsx(join(outDir, 'fz.xlsx'), F({ sheets: F([{ rows: F([F(['x'])]) }]) }), F({}))],
    ['opEditXlsx', () => office.opEditXlsx(plain, F([F({ op: 'set_value', ref: 'D1', value: 'y' })]), F({}))],
    ['opWriteDocx', () => office.opWriteDocx(join(outDir, 'fz.docx'), F({ markdown: '# x' }), F({}))],
    ['opFillTemplate', () => office.opFillTemplate(join(outDir, 'fz.docx'), join(outDir, 'fz-out.docx'), F({ a: 'b' }), F({}))],
    ['opConvert', () => office.opConvert(plain, join(outDir, 'fz.csv'), F({}))],
    ['opRead(withFormatting)', () => office.opRead(join(outDir, 'fz.docx'), F({ withFormatting: true }), F({}))],
  ];
  for (const [label, run] of calls) {
    try {
      await run();
    } catch (err) {
      throw new Error(`${label} 在冻结参数下失败: ${err.message}`);
    }
  }
  // 冻结的输入对象必须原封不动(不能被塞进 containerNote)
  const frozen = F({ withFormatting: true });
  await office.opRead(weird, frozen);
  if (Object.keys(frozen).join(',') !== 'withFormatting') throw new Error('调用方参数被改写了: ' + JSON.stringify(frozen));
  return `${calls.length} 个操作在冻结参数下全部正常,且参数未被改写`;
});

/** 按给定改写方式生成一个变体 docx。 */
async function variantDocx(base, cfg) {
  const PizZip = (await import('pizzip')).default;
  const out = new PizZip();
  for (const [name, entry] of Object.entries(base.files)) {
    if (entry.dir) continue;
    const isRef = name.endsWith('.rels') || name.endsWith('[Content_Types].xml');
    const content = cfg.rewriteRefs && isRef
      ? Buffer.from(cfg.rewriteRefs(entry.asText()), 'utf8')
      : entry.asNodeBuffer();
    out.file(cfg.rename(name), content);
  }
  return out.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

await t('兼容: 部件名大小写/反斜杠不同(Windows 常见)也能读', async () => {
  const { writeFile } = await import('node:fs/promises');
  const PizZip = (await import('pizzip')).default;
  const base = new PizZip(await word.writeDocx({ markdown: '# 标题\n\n大小写测试正文' }));
  // 重点覆盖「条目名与引用它的 .rels/Content_Types 写法一致但都不规范」:
  // 只改条目名会让两边对不上,反而更糟 —— 早先的版本就是漏了这种。
  const variants = {
    '全大写目录': { rename: (n) => n.replace(/^word\//, 'Word/').replace('document.xml', 'Document.xml') },
    '大小写+反斜杠混合': { rename: (n) => n.replace(/^word\//, 'Word\\').replace('document.xml', 'Document.XML') },
    '仅反斜杠': { rename: (n) => n.replaceAll('/', '\\') },
    '条目与引用均为反斜杠': {
      rename: (n) => n.replaceAll('/', '\\'),
      rewriteRefs: (text) => text.replaceAll('word/', 'word\\').replaceAll('xl/', 'xl\\'),
    },
    '条目与引用均为大写': {
      rename: (n) => n.replace(/^word\//, 'Word/'),
      rewriteRefs: (text) => text.replaceAll('word/', 'Word/'),
    },
  };
  let at = 0;
  for (const [label, cfg] of Object.entries(variants)) {
    at += 1;
    const p = join(outDir, `case-${at}.docx`);
    await writeFile(p, await variantDocx(base, cfg));
    const r = await office.opRead(p, { withFormatting: true });
    if (!String(r.content).includes('大小写测试正文')) throw new Error(`${label}: 正文没读出来`);
    if (!r.meta.formatting?.paragraphs?.length) throw new Error(`${label}: 格式未提取`);
    if (!String(r.content).includes('自动修正')) throw new Error(`${label}: 未提示修正条目名`);
  }
  return `${Object.keys(variants).length} 种写法(大写/混合/反斜杠/引用同步)均可读`;
});

await t('兼容: mammoth 解析不了时用内置解析器兜底(不硬失败)', async () => {
  const PizZip = (await import('pizzip')).default;
  const { writeFile } = await import('node:fs/promises');
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const zip = new PizZip(new PizZip(await word.writeDocx({ markdown: '占位' })).generate({ type: 'nodebuffer' }));
  // 没有 w:body: mammoth 抛 "Could not find the body element"
  zip.file('word/document.xml', `<w:document ${NS}><w:p><w:r><w:t>兜底正文内容</w:t></w:r></w:p></w:document>`);
  const p = join(outDir, 'nobody.docx');
  await writeFile(p, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  const r = await office.opRead(p, {});
  if (!String(r.content).includes('兜底正文内容')) throw new Error('兜底也没读出来: ' + String(r.content).slice(0, 80));
  if (r.meta?.reader !== undefined) throw new Error('meta.reader 不该出现在这个层级');
  const hr = await office.opRead(p, { format: 'html' });
  if (!String(hr.html || hr.content).includes('兜底正文内容')) throw new Error('html 输出缺失');
  return 'mammoth 失败后内置解析器成功兜底';
});

await t('兼容: 主文档部件不在规范路径时按包关系解析', async () => {
  const PizZip = (await import('pizzip')).default;
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const relsNs = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
  const docWith = (text) => `<w:document ${NS}><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`;
  const base = new PizZip(await word.writeDocx({ markdown: '占位' }));

  // 主部件放在 word/main.xml,由 _rels/.rels 指过去(规范名不存在)
  const custom = new PizZip(base.generate({ type: 'nodebuffer' }));
  custom.file('word/main.xml', docWith('关系解析到的正文'));
  custom.file('word/document.xml', '');           // 规范名存在但为空
  custom.file('_rels/.rels', `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/main.xml"/></Relationships>`);
  const main = resolveMainPart(new PizZip(custom.generate({ type: 'nodebuffer' })));
  if (main?.path !== 'word/main.xml') throw new Error('未按关系解析主部件: ' + JSON.stringify(main?.path));

  // 没有关系文件时退回规范名
  const plain = new PizZip(base.generate({ type: 'nodebuffer' }));
  plain.remove('_rels/.rels');
  const fallback = resolveMainPart(new PizZip(plain.generate({ type: 'nodebuffer' })));
  if (fallback?.path !== 'word/document.xml') throw new Error('未退回规范名: ' + JSON.stringify(fallback?.path));

  // 规范名缺失时按「以 document.xml 结尾」容忍查找
  const odd = new PizZip(base.generate({ type: 'nodebuffer' }));
  odd.remove('_rels/.rels');
  odd.remove('word/document.xml');
  odd.file('Word/Sub/Document.XML', docWith('容忍查找'));
  const tolerant = resolveMainPart(new PizZip(odd.generate({ type: 'nodebuffer' })));
  if (tolerant?.path !== 'Word/Sub/Document.XML') throw new Error('容忍查找失败: ' + JSON.stringify(tolerant?.path));
  return '关系解析 / 规范名回退 / 容忍查找 三条路径均正确';
});

await t('诊断: 无法识别的容器报错列出条目与插件版本', async () => {
  const { writeFile } = await import('node:fs/promises');
  const PizZip = (await import('pizzip')).default;
  const zip = new PizZip();
  zip.file('hello.txt', 'not office');
  zip.file('[Content_Types].xml', '<Types/>');
  const p = join(outDir, 'not-office-2.docx');
  await writeFile(p, zip.generate({ type: 'nodebuffer' }));
  try {
    await office.opRead(p, {});
    throw new Error('损坏文件却读成功了');
  } catch (err) {
    if (err.code !== 'BAD_CONTAINER') throw err;
    for (const needle of ['插件 v', 'hello.txt']) {
      if (!err.message.includes(needle)) throw new Error(`诊断缺少「${needle}」: ${err.message}`);
    }
  }
  return '报错含插件版本与实际条目名';
});

await t('兼容: 主部件由关系声明在非规范路径时也能读', async () => {
  const { writeFile } = await import('node:fs/promises');
  const PizZip = (await import('pizzip')).default;
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const relsNs = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
  const zip = new PizZip(new PizZip(await word.writeDocx({ markdown: '占位' })).generate({ type: 'nodebuffer' }));
  zip.file('word/main.xml', `<w:document ${NS}><w:body><w:p><w:r><w:t>关系声明的正文</w:t></w:r></w:p></w:body></w:document>`);
  zip.remove('word/document.xml');
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/main.xml"/></Relationships>`);
  const p = join(outDir, 'custom-main.docx');
  await writeFile(p, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  const r = await office.opRead(p, {});
  if (!String(r.content).includes('关系声明的正文')) throw new Error('未按关系读到正文: ' + String(r.content).slice(0, 80));
  return '主部件在 word/main.xml 也能正常读取';
});

const failed = results.filter((r) => !r.ok);
console.log(`\n===== 结果: ${results.length - failed.length}/${results.length} 通过 =====`);
if (failed.length) {
  console.log('\n失败明细:');
  for (const f of failed) console.log(`- ${f.name}\n  ${f.detail.split('\n').slice(0, 6).join('\n  ')}`);
  process.exit(1);
}
