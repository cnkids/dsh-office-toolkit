// Standalone selftest for the office core (no DSH needed).
// Usage: node test/selftest.mjs   (outputs into test/out/)
import { mkdir } from 'node:fs/promises';
import { dirname, isAbsolute, join } from 'node:path';
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
import { pageOf, outlineOf } from '../lib/core/office.js';
import { CORE_DEPS, depFailure, lazyModule, missingDeps } from '../lib/core/deps.js';
import { editDocx, scanParagraphs } from '../lib/core/docx-edit.js';
import { openOoxml } from '../lib/core/ooxml.js';
import { normalizeStyleSpec, normalizeTableSpec, autoColumnPercents, lengthToTwips, runStyleFromCss, paragraphStyleFromCss, declarationsFrom } from '../lib/core/docx-style.js';
import { scanTables } from '../lib/core/docx-table.js';
import { headingStyleIds, levelTextAt, NUMBERING_PRESETS } from '../lib/core/docx-numbering.js';
import { formatCounter } from '../lib/core/docx-numbering-read.js';
import { fitImageSize, loadImage, probeImage } from '../lib/core/image.js';
import { writeFileSync, readFileSync } from 'node:fs';
import { deflateSync } from 'node:zlib';

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

// ---------------------------------------------------------------------------
// Word 分段阅读:offset 续读 + 标题大纲
// ---------------------------------------------------------------------------
/** 造一份带标题的长文档,并返回路径。 */
async function writeLongDocx() {
  const { writeFile } = await import('node:fs/promises');
  const parts = ['# 项目总结报告'];
  for (let c = 1; c <= 12; c++) {
    parts.push('', `## 第 ${c} 章 主题 ${c}`);
    for (let i = 0; i < 120; i++) parts.push(`第 ${c}-${i + 1} 段：用于分页测试的正文内容，含中文字符与编号 ${i % 89}，把文档撑到足够大。`);
  }
  const path = join(outDir, 'long.docx');
  await writeFile(path, await word.writeDocx({ markdown: parts.join('\n') }, { title: '报告' }));
  return path;
}

/** 从 opRead 结果里取出正文(去掉前面的摘要行)。 */
const pageBody = (content) => content.split('\n\n').slice(1).join('\n\n');

await t('页眉页脚: 文字 / 页码域 / 每节引用 / updateFields', async () => {
  const buf = await word.writeDocx({ markdown: '# 通知\n\n正文一段。' }, {
    title: '通知',
    header: { text: 'XX单位文件', align: 'center', fontSizePt: 14, bold: true },
    footer: { text: '— ', pageNumber: '{page} — 共 {total} 页', align: 'center' },
  });
  const { zip } = await openOoxml(buf, 'docx');
  const doc = zip.file('word/document.xml').asText();
  const header = zip.file('word/header1.xml')?.asText() ?? '';
  const footer = zip.file('word/footer1.xml')?.asText() ?? '';
  const settings = zip.file('word/settings.xml')?.asText() ?? '';
  const markers = [
    ['生成了页眉部件', header.includes('XX单位文件')],
    ['页眉居中', header.includes('w:jc w:val="center"')],
    ['页眉加粗与字号', /<w:b\/>/.test(header) && /w:sz w:val="28"/.test(header)],
    ['页脚有 PAGE 域', /PAGE/.test(footer)],
    ['页脚有 NUMPAGES 域', /NUMPAGES/.test(footer)],
    ['页脚文字与页码拼接', footer.includes('—') && footer.includes('共')],
    ['小节引用页眉页脚', /headerReference/.test(doc) && /footerReference/.test(doc)],
    ['打开时更新域', /updateFields/.test(settings)],
    ['正文可读回', (await word.readDocx(buf, {})).content.includes('正文一段')],
  ];
  const bad = markers.filter(([, ok]) => !ok).map(([n]) => n);
  if (bad.length) throw new Error('缺失: ' + bad.join(', '));
  return `${markers.length} 项（含 PAGE/NUMPAGES 域）`;
});

await t('目录: 插入 TOC 域并指定收录级别', async () => {
  const buf = await word.writeDocx({ markdown: '# 一、总体\n\n正文\n\n## （一）细节\n\n正文\n\n### 三级\n\n正文' }, {
    title: '报告', toc: { title: '目　录', levels: 2 },
  });
  const doc = await partText(buf, 'document.xml');
  const markers = [
    ['有 TOC 指令域', /instrText[^<]*TOC/.test(doc)],
    ['目录标题写入', doc.includes('目　录')],
    ['收录 1-2 级', /1-2/.test(doc)],
    ['一次性不带目录时没有 TOC', !/instrText[^<]*TOC/.test(await partText(await word.writeDocx({ markdown: '# A' }, { title: 't' }), 'document.xml'))],
    ['正文可读回', (await word.readDocx(buf, {})).content.includes('三级')],
  ];
  const bad = markers.filter(([, ok]) => !ok).map(([n]) => n);
  if (bad.length) throw new Error('缺失: ' + bad.join(', '));
  return 'TOC 域 + 级别范围 + 标题';
});

await t('页眉页脚/目录: 参数错误清晰', async () => {
  const cases = [
    { spec: { header: {} }, want: '至少要给 text 或 pageNumber' },
    { spec: { footer: { pageNumber: '第 X 页' } }, want: '要用 {page}' },
    { spec: { header: { fontSizePt: 999 } }, want: '需为 0–200 的磅值' },
    { spec: { toc: { levels: 12 } }, want: 'toc.levels 需为 1–9' },
    { spec: { header: '页眉' }, want: 'header 需为对象' },
  ];
  for (const { spec, want } of cases) {
    let msg = '';
    try { await word.writeDocx({ markdown: 'x' }, { title: 't', ...spec }); } catch (e) { msg = `${e.code}|${e.message}`; }
    if (!msg.includes(want)) throw new Error(`${JSON.stringify(spec)} 的报错不对: ${msg}`);
    if (!msg.startsWith('INVALID_ARGS')) throw new Error(`${JSON.stringify(spec)} 应报 INVALID_ARGS: ${msg}`);
  }
  return `${cases.length} 种参数错误都带 INVALID_ARGS`;
});

await t('word 分段: 首段给出总长与「继续读」的 offset', async () => {
  const path = await writeLongDocx();
  const r = await office.opRead(path, { maxChars: 5000 });
  if (!/第 0–\d+ 字符 \/ 共 \d+ 字符/.test(r.content)) throw new Error('摘要没写本次区间与总长: ' + r.content.split('\n').slice(0, 5).join(' | '));
  if (!r.content.includes(`offset: ${r.meta.end}`)) throw new Error('没有给出下次该传的 offset');
  if (r.meta.truncated !== true || r.meta.totalChars <= r.meta.end) throw new Error('meta 不对: ' + JSON.stringify(r.meta));
  return `共 ${r.meta.totalChars} 字符，本段 [0, ${r.meta.end})`;
});

await t('word 分段: 逐段拼回来与一次读完逐字节相同(不重不漏)', async () => {
  const path = await writeLongDocx();
  const full = pageBody((await office.opRead(path, { maxChars: 10_000_000 })).content);
  let offset = 0;
  let joined = '';
  for (let i = 0; i < 200; i++) {
    const r = await office.opRead(path, { offset, maxChars: 3000 });
    joined += pageBody(r.content);
    if (!r.meta.truncated) break;
    offset = r.meta.end;
  }
  if (joined !== full) {
    const at = [...full].findIndex((ch, i) => ch !== joined[i]);
    throw new Error(`拼接结果与全文不一致(首个差异 @${at})：${JSON.stringify(joined.slice(at - 20, at + 20))} vs ${JSON.stringify(full.slice(at - 20, at + 20))}`);
  }
  return `${full.length} 字符逐字节一致`;
});

await t('word 分段: 分段边界落在句末/换行,不切半句话', async () => {
  const path = await writeLongDocx();
  const r = await office.opRead(path, { maxChars: 3000 });
  const body = pageBody(r.content);
  if (!/[。！？；\n]$/.test(body)) throw new Error('本段没有收在句子/段落边界: ' + JSON.stringify(body.slice(-30)));
  if (!body.includes('第 1-1 段')) throw new Error('首段内容不对');
  // 边界字符归本段,下一段紧接着往下(拼接无损由下一条测试保证)
  const nxt = pageBody((await office.opRead(path, { offset: r.meta.end, maxChars: 3000 })).content);
  if (body + nxt !== pageBody((await office.opRead(path, { maxChars: 10_000_000 })).content).slice(0, body.length + nxt.length)) {
    throw new Error('两段衔接处与全文不一致');
  }
  return `本段 ${body.length} 字符，收尾 ${JSON.stringify(body.slice(-1))}`;
});

await t('word 分段: pageOf 的边界与不重不漏(纯函数)', async () => {
  const text = '第一句。第二句！第三句？' + 'x'.repeat(60) + '。末尾';
  const first = pageOf(text, 0, 20);
  if (first.end !== text.indexOf('？') + 1 || first.more !== true) throw new Error('没有收在句末标点上: ' + JSON.stringify(first));
  if (first.slice !== text.slice(0, first.end)) throw new Error('切片与 end 不匹配');
  // 窗口里没有任何边界字符时硬切,不会为了凑边界只返回一点点
  const dense = 'y'.repeat(100);
  const hard = pageOf(dense, 0, 20);
  if (hard.end !== 20 || hard.slice.length !== 20) throw new Error('无边界时没有硬切: ' + JSON.stringify(hard));
  // 逐段拼回来等于原文
  let offset = 0;
  let joined = '';
  for (let i = 0; i < 100; i++) {
    const page = pageOf(text, offset, 7);
    joined += page.slice;
    if (!page.more) break;
    offset = page.end;
  }
  if (joined !== text) throw new Error('分页拼接与原文不一致: ' + JSON.stringify(joined));
  if (pageOf(text, text.length, 10).more !== false) throw new Error('读到末尾还标记 more');
  return `${text.length} 字符按 7 字符分页拼接无损`;
});

await t('word 分段: outlineOf 只认真标题(围栏/无空格/超 6 级都排除)', async () => {
  const md = ['# A', '', '正文', '', '```sh', '# 代码里的注释', '```', '', '### C ###', '####### 七个井号', '#没有空格', '###### 六级'].join('\n');
  const items = outlineOf(md);
  const titles = items.map((h) => `${h.level}:${h.title}`);
  if (titles.join(' | ') !== '1:A | 3:C ### | 6:六级') throw new Error('标题识别不对: ' + titles.join(' | '));
  if (items[1].offset !== md.indexOf('### C')) throw new Error('偏移不对: ' + items[1].offset);
  if (outlineOf('没有标题的正文').length !== 0) throw new Error('无标题正文不该有结果');
  return titles.join(' | ');
});

await t('word 分段: 标题大纲带级别/偏移/标题,可据此跳读', async () => {
  const path = await writeLongDocx();
  const r = await office.opRead(path, { outline: true });
  if (!r.content.includes('标题大纲: 13 条')) throw new Error('标题条数不对: ' + r.content.split('\n').slice(0, 5).join(' | '));
  const rows = r.content.split('```tsv')[1].split('```')[0].trim().split('\n').slice(1).map((l) => l.split('\t'));
  if (rows.length !== 13) throw new Error('大纲行数不对: ' + rows.length);
  if (rows[0][0] !== '1' || rows[0][1] !== '0' || rows[0][2] !== '项目总结报告') throw new Error('一级标题行不对: ' + rows[0].join(','));
  const ch5 = rows.find((x) => x[2] === '第 5 章 主题 5');
  if (!ch5) throw new Error('缺少第 5 章');
  // 用大纲里的偏移直接跳读,应该正好从该章标题开始
  const jump = pageBody((await office.opRead(path, { offset: Number(ch5[1]), maxChars: 2000 })).content);
  if (!jump.includes('第 5 章 主题 5')) throw new Error('按大纲偏移跳读没落在该章: ' + jump.slice(0, 60));
  if (jump.includes('第 4 章')) throw new Error('跳读越到了上一章');
  return `${rows.length} 条标题，第 5 章 @${ch5[1]}`;
});

await t('word 分段: offset 越界 / 非法值 / html 组合都有清晰报错', async () => {
  const msgOf = async (fn) => { try { await fn(); return ''; } catch (err) { return err.message; } };
  const path = await writeLongDocx();
  const beyond = await office.opRead(path, { offset: 9_999_999 });
  if (!beyond.content.includes('已到文档末尾') || beyond.meta.ended !== true) throw new Error('越界提示不对: ' + beyond.content.split('\n').at(-1));
  const bad = await msgOf(() => office.opRead(path, { offset: -5 }));
  if (!/offset 需为/.test(bad)) throw new Error('非法 offset 报错不清晰: ' + bad);
  const badType = await msgOf(() => office.opRead(path, { offset: 1.5 }));
  if (!/offset 需为/.test(badType)) throw new Error('小数 offset 没被拦: ' + badType);
  const htmlErr = await msgOf(() => office.opRead(path, { format: 'html', offset: 10 }));
  if (!/html.*不支持 offset/.test(htmlErr)) throw new Error('html+offset 报错不清晰: ' + htmlErr);
  const htmlOutline = await msgOf(() => office.opRead(path, { format: 'html', outline: true }));
  if (!/不支持 offset \/ outline/.test(htmlOutline)) throw new Error('html+outline 报错不清晰: ' + htmlOutline);
  return '越界/负数/小数/html 组合 5 种情况';
});

await t('word 分段: html 模式不分页但说明该切 text 模式', async () => {
  const path = await writeLongDocx();
  const r = await office.opRead(path, { format: 'html', maxChars: 400 });
  if (!/不分页；要分段读长文请用默认 text 模式/.test(r.content)) throw new Error('html 模式没给分页指引: ' + r.content.split('\n').slice(0, 5).join(' | '));
  if (typeof r.html !== 'string' || !r.html.includes('<h1>')) throw new Error('html 输出丢失');
  return `${r.html.length} 字符 html（未分页）`;
});

await t('word 分段: 格式报告只在第一段给出,不重复占位', async () => {
  const path = await writeLongDocx();
  const first = await office.opRead(path, { withFormatting: true, maxChars: 4000 });
  if (!first.content.includes('## 格式报告（整篇）')) throw new Error('第一段没有格式报告');
  const second = await office.opRead(path, { withFormatting: true, offset: first.meta.end, maxChars: 4000 });
  if (second.content.includes('## 格式报告')) throw new Error('后续段重复给了格式报告');
  if (!second.content.includes('格式报告只在 offset=0 时给出')) throw new Error('后续段没有说明格式报告去哪了');
  return '首段带报告,续读只给提示';
});

// ---------------------------------------------------------------------------
// 排版:字体 / 行距 / 首行缩进(单位换算 + 文档级样式 + 内联 CSS + 改已有文档)
// ---------------------------------------------------------------------------
/** 取出 docx 里的某个部件文本。 */
async function partText(buf, name) {
  const { zip } = await openOoxml(buf, 'docx');
  const entry = zip.file(name) || zip.file(`word/${name}`);
  return entry ? entry.asText() : '';
}

await t('排版: 单位换算与规格规范化(公文体例)', async () => {
  const norm = normalizeStyleSpec({ font: '仿宋_GB2312', sizePt: 16, lineSpacingPt: 28.8, firstLineIndentChars: 2, align: 'both', spacingBeforePt: 0, spacingAfterPt: 0 });
  if (norm.sizeHalfPt !== 32) throw new Error('三号应换算成 32 半磅: ' + norm.sizeHalfPt);
  if (norm.line !== 576 || norm.lineRule !== 'exact') throw new Error('28.8 磅固定行距应为 576 twip/exact: ' + JSON.stringify(norm));
  if (norm.firstLineTwips !== 640) throw new Error('三号 2 字符缩进应为 640 twip: ' + norm.firstLineTwips);
  if (norm.font !== '仿宋_GB2312' || norm.align !== 'both') throw new Error('字体/对齐不对: ' + JSON.stringify(norm));
  const multiple = normalizeStyleSpec({ lineSpacingMultiple: 1.5 });
  if (multiple.line !== 360 || multiple.lineRule !== 'auto') throw new Error('1.5 倍行距应为 360/auto: ' + JSON.stringify(multiple));
  if (lengthToTwips('32px', null) !== 480) throw new Error('32px 应按 96dpi 换算成 480');
  if (lengthToTwips('2em', 32) !== 640) throw new Error('2em(三号) 应为 640');
  const heads = normalizeStyleSpec({ font: '仿宋', headings: { font: '黑体', sizePt: 16 } });
  if (heads.headings.font !== '黑体') throw new Error('headings 没被解析');
  const noHeads = await (async () => { try { normalizeStyleSpec({ headings: { font: 'x' } }, { allowHeadings: false }); return ''; } catch (e) { return e.code; } })();
  if (noHeads !== 'INVALID_ARGS') throw new Error('set_style 不该接受 headings: ' + noHeads);
  return `三号 32 半磅 / 行距 576 / 缩进 640 twip`;
});

await t('排版: 规格非法时报错清晰', async () => {
  const cases = [
    { spec: { sizePt: 'x' }, want: 'sizePt 需为' },
    { spec: { align: '居中' }, want: 'align 可为' },
    { spec: { lineSpacingPt: 28, lineSpacingMultiple: 1.5 }, want: '只能给一个' },
    { spec: { color: 'red' }, want: 'color 需为 RRGGBB' },
    { spec: { font: '' }, want: 'font 不能为空' },
  ];
  for (const { spec, want } of cases) {
    let msg = '';
    try { normalizeStyleSpec(spec); } catch (e) { msg = e.message; }
    if (!msg.includes(want)) throw new Error(`${JSON.stringify(spec)} 的报错不对: ${msg}`);
  }
  // 非对象单独测:不要在同一个数组里混类型
  let typeMsg = '';
  try { normalizeStyleSpec('楷体'); } catch (e) { typeMsg = e.message; }
  if (!typeMsg.includes('style 需为对象')) throw new Error('非对象入参的报错不对: ' + typeMsg);
  return `${cases.length + 1} 种非法规格`;
});

await t('表格: 单元格内边距 cellMargins', async () => {
  const spec = normalizeTableSpec({ cellMargins: { left: 108, right: 108, top: 40, bottom: 40 } });
  if (spec.cellMargins.left !== 108 || spec.cellMargins.bottom !== 40) throw new Error('内边距没解析: ' + JSON.stringify(spec.cellMargins));
  const partial = normalizeTableSpec({ cellMargins: { left: 200 } }).cellMargins;
  if (partial.left !== 200 || partial.right !== 108 || partial.top !== 0) throw new Error('没给的边应补 Word 默认值: ' + JSON.stringify(partial));

  const md = '| A | B |\n| --- | --- |\n| 1 | 2 |';
  const written = await partText(await word.writeDocx({ markdown: md }, { title: 't', style: { table: { cellMargins: { left: 108, right: 108, top: 40, bottom: 40 } } } }), 'document.xml');
  const edited = await partText((await editDocx(await word.writeDocx({ markdown: md }, { title: 't' }), [
    { op: 'set_table', table: 1, cellMargins: { left: 108, right: 108, top: 40, bottom: 40 } },
  ])).buf, 'document.xml');
  const cellMar = /<w:tblCellMar>.*?<\/w:tblCellMar>/;
  // 属性顺序不影响语义:两种写法都接受(两个字面量,不动态构造正则)
  const SIDE_RE = {
    top: /<w:top [^>]*\/>/,
    bottom: /<w:bottom [^>]*\/>/,
    left: /<w:left [^>]*\/>/,
    right: /<w:right [^>]*\/>/,
  };
  const sideValue = (xml, side) => {
    const hit = xml.match(SIDE_RE[side]);
    if (!hit) return null;
    return { value: (hit[0].match(/w:w="(\d+)"/) || [])[1], type: (hit[0].match(/w:type="([a-z]+)"/) || [])[1] };
  };
  for (const [label, doc] of [['写文档', written], ['改文档', edited]]) {
    const hit = doc.match(cellMar);
    if (!hit) throw new Error(label + ' 没生成 tblCellMar');
    for (const [side, want] of [['top', '40'], ['bottom', '40'], ['left', '108'], ['right', '108']]) {
      const got = sideValue(hit[0], side);
      if (got?.value !== want || got?.type !== 'dxa') throw new Error(`${label} 的 ${side} 应为 ${want}dxa，实际 ${JSON.stringify(got)}`);
    }
  }
  const bad = [
    { spec: { cellMargins: {} }, want: '至少要给' },
    { spec: { cellMargins: { left: -1 } }, want: '需为 0–5000' },
    { spec: { cellMargins: '108' }, want: '需为对象' },
  ];
  for (const { spec: badSpec, want: msg } of bad) {
    let text = '';
    try { normalizeTableSpec(badSpec); } catch (e) { text = e.message; }
    if (!text.includes(msg)) throw new Error(JSON.stringify(badSpec) + ' 的报错不对: ' + text);
  }
  return '左右 108 / 上下 40 两条路径都落到 tblCellMar';
});

await t('性能: 万段文档的批量改样式/编号必须线性(不退回 O(n²))', async () => {
  // 逐段拼字符串的写法在 1 万段上要 2 秒以上且随段数平方增长;这里卡一个宽松上限
  const parts = ['# 压测文档'];
  for (let i = 0; i < 2500; i += 1) parts.push('', `## 第 ${i} 节`, '', `第 ${i} 节正文。`, '', '- 列表甲', '- 列表乙');
  const buf = await word.writeDocx({ markdown: parts.join('\n') }, { title: 'perf' });
  const paragraphs = scanParagraphs(await partText(buf, 'document.xml')).length;
  if (paragraphs < 10000) throw new Error('测试前提不成立: 段落数只有 ' + paragraphs);
  const timeOf = async (ops) => { const t0 = Date.now(); await editDocx(buf, ops); return Date.now() - t0; };
  const style = await timeOf([{ op: 'set_style', scope: 'all', sizePt: 14 }]);
  const numbering = await timeOf([{ op: 'set_numbering', scope: 'all' }]);
  const replace = await timeOf([{ op: 'replace_text', find: '节正文', replace: '节正文内容' }]);
  const budget = 900;
  const slow = [['set_style', style], ['set_numbering', numbering], ['replace_text', replace]].filter(([, ms]) => ms > budget);
  if (slow.length) throw new Error(`${paragraphs} 段超过 ${budget}ms: ` + slow.map(([n, ms]) => `${n} ${ms}ms`).join('、'));
  return `${paragraphs} 段：改样式 ${style}ms / 编号 ${numbering}ms / 替换 ${replace}ms`;
});

await t('编号: 预置方案与标题样式识别', async () => {
  if (NUMBERING_PRESETS.join() !== 'multicol-1_1_1,gongwen-1_1_1_1') throw new Error('预置方案变了: ' + NUMBERING_PRESETS.join());
  if (levelTextAt(0) !== '%1.' || levelTextAt(2) !== '%1.%2.%3') throw new Error('各级编号文字不对: ' + levelTextAt(0) + '/' + levelTextAt(2));
  const styles = (await openOoxml(await word.writeDocx({ markdown: '# 标题' }, { title: 't' }), 'docx')).zip.file('word/styles.xml').asText();
  const ids = headingStyleIds(styles);
  if (ids.get(1) !== 'Heading1' || ids.get(3) !== 'Heading3') throw new Error('标题样式识别不对: ' + JSON.stringify([...ids]));
  return `1→${ids.get(1)}、3→${ids.get(3)}；${levelTextAt(0)} / ${levelTextAt(1)} / ${levelTextAt(2)}`;
});

await t('编号: 标题挂样式链接 + 正文列表按当前标题层级挂下一级', async () => {
  const md = ['# 概述', '', '本章说明。', '', '- 要点一', '- 要点二', '', '## 细节', '', '- 细节点', '', '# 结论', '', '- 结论点'].join('\n');
  const buf = await word.writeDocx({ markdown: md }, { title: 't' });
  const { buf: out, changes } = await editDocx(buf, [{ op: 'set_numbering', scope: 'all', style: 'multicol-1_1_1', linkToHeading: true }]);
  if (!/正文列表 4 段/.test(changes[0])) throw new Error('列表段计数不对: ' + changes[0]);
  // 只看那一段自己的 XML 里有没有 numPr(避免匹配到后面列表段)
  const plainParagraphHasNoNumbering = (xml) => {
    const target = scanParagraphs(xml).find((p) => p.text === '本章说明。');
    return Boolean(target) && !/<w:numPr>/.test(xml.slice(target.start, target.end));
  };
  const doc = await partText(out, 'document.xml');
  const numbering = await partText(out, 'numbering.xml');
  const markers = [
    ['新建了多级 abstractNum', /<w:multiLevelType w:val="multilevel"\/>/.test(numbering)],
    ['一级挂 Heading1', /<w:lvl w:ilvl="0">[\s\S]*?<w:pStyle w:val="Heading1"\/>[\s\S]*?<w:lvlText w:val="%1\."\/>/.test(numbering)],
    ['二级挂 Heading2 且文字为 %1.%2', /<w:lvl w:ilvl="1">[\s\S]*?<w:pStyle w:val="Heading2"\/>[\s\S]*?<w:lvlText w:val="%1\.%2"\/>/.test(numbering)],
    ['正文列表用 ilvl=1', /<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="\d+"\/><\/w:numPr>/.test(doc)],
    ['H2 下的列表用 ilvl=2', /<w:numPr><w:ilvl w:val="2"\/><w:numId w:val="\d+"\/><\/w:numPr>/.test(doc)],
    ['普通正文没被编号', plainParagraphHasNoNumbering(doc)],
    ['标题段落没写 numPr(靠样式链接)', !/w:pStyle w:val="Heading1"\/><w:numPr>/.test(doc)],
  ];
  const bad = markers.filter(([, ok]) => !ok).map(([n]) => n);
  if (bad.length) throw new Error('缺失: ' + bad.join(', '));
  const back = await word.readDocx(out, {});
  if (!back.content.includes('概述')) throw new Error('文档读不回来了');
  return changes[0];
});

await t('编号: 公文式预置(一、/（一）/1./（1）)', async () => {
  const md = ['# 一、总体情况', '', '## （一）主要进展', '', '- 完成事项甲', '', '### 1. 具体做法', '', '- 做法细节'].join('\n');
  const buf = await word.writeDocx({ markdown: md }, { title: 't' });
  const { buf: out, changes } = await editDocx(buf, [{ op: 'set_numbering', scope: 'all', style: 'gongwen-1_1_1_1' }]);
  if (!changes[0].includes('gongwen-1_1_1_1')) throw new Error('结果没写方案名: ' + changes[0]);
  const numbering = await partText(out, 'numbering.xml');
  const block = numbering.slice(numbering.indexOf('<w:nsid w:val="0D5F0002"'));
  if (!block.startsWith('<w:nsid') && !block.includes('<w:nsid w:val="0D5F0002"/>')) throw new Error('公文方案的 nsid 没写对');
  const levels = (block.match(/<w:lvl w:ilvl="\d+">[\s\S]*?<\/w:lvl>/g) || []).slice(0, 4).map((b) => ({
    fmt: (b.match(/<w:numFmt w:val="([^"]+)"\/>/) || [])[1],
    text: (b.match(/<w:lvlText w:val="([^"]+)"\/>/) || [])[1],
    style: (b.match(/<w:pStyle w:val="([^"]+)"\/>/) || [])[1],
  }));
  const want = [
    { fmt: 'chineseCounting', text: '%1、', style: 'Heading1' },
    { fmt: 'chineseCounting', text: '（%2）', style: 'Heading2' },
    { fmt: 'decimal', text: '%3.', style: 'Heading3' },
    { fmt: 'decimal', text: '（%4）', style: 'Heading4' },
  ];
  for (let i = 0; i < want.length; i += 1) {
    const got = levels[i];
    if (got?.fmt !== want[i].fmt || got?.text !== want[i].text || got?.style !== want[i].style) {
      throw new Error(`第 ${i + 1} 级不对: 期望 ${JSON.stringify(want[i])}，实际 ${JSON.stringify(got)}`);
    }
  }
  // 正文列表挂在「当前标题层级的下一级」:H2 下 → ilvl=2(显示「1.」)、H3 下 → ilvl=3(显示「（1）」)
  const doc = await partText(out, 'document.xml');
  const listIlvls = scanParagraphs(doc)
    .filter((p) => /<w:numPr>/.test(doc.slice(p.start, p.end)))
    .map((p) => (doc.slice(p.start, p.end).match(/<w:ilvl w:val="(\d+)"\/>/) || [])[1]);
  if (listIlvls.join() !== '2,3') throw new Error('正文列表的层级没按标题链挂: ' + listIlvls.join());
  if (!(await word.readDocx(out, {})).content.includes('完成事项甲')) throw new Error('文档读不回来了');
  const err = await (async () => { try { await editDocx(buf, [{ op: 'set_numbering', style: 'x' }]); return ''; } catch (e) { return e.message; } })();
  if (!err.includes('gongwen-1_1_1_1')) throw new Error('style 报错没列出公文预置: ' + err);
  return `一、/（一）/1./（1）四级都对；${changes[0].replace(/^.*?：/, '')}`;
});

await t('编号: 读取时展开成文字(公文 一、/（一）/1.)', async () => {
  const md = ['# 总体要求', '', '正文甲', '', '## 工作目标', '', '正文乙', '', '## 重点任务', '', '### 具体措施', '', '# 保障措施'].join('\n');
  const buf = await editDocx(await word.writeDocx({ markdown: md }, { title: 't' }), [
    { op: 'set_numbering', style: 'gongwen-1_1_1_1' },
  ]).then((r) => r.buf);
  const read = await word.readDocx(buf, {});
  const lines = read.content.split('\n').filter((line) => line.trim());
  const expected = ['一、 总体要求', '正文甲', '（一） 工作目标', '正文乙', '（二） 重点任务', '1. 具体措施', '二、 保障措施'];
  if (lines.join('|') !== expected.join('|')) throw new Error('展开结果不对: ' + lines.join('|'));
  // 兄弟项目 official-doc-rules 的行首规则要能认出来
  const RULES = [/^[一二三四五六七八九十]+、/, /^（[一二三四五六七八九十]+）/, /^\d+\./];
  const levels = lines.map((line) => RULES.findIndex((re) => re.test(line)));
  if (levels.join() !== '0,-1,1,-1,1,2,0') throw new Error('行首规则识别不对: ' + levels.join());
  if (!read.html.includes('<h1>一、 总体要求</h1>')) throw new Error('HTML 里没有编号: ' + read.html);
  if (!read.meta.messages[0]?.includes('5 段自动编号')) throw new Error('没提示展开了多少段: ' + read.meta.messages.join());
  // 没有编号的文件不该被改动
  const plain = await word.readDocx(await word.writeDocx({ markdown: md }, { title: 't' }), {});
  if (/^[一二三四五六七八九十]+、/m.test(plain.content) || plain.meta.messages.length) {
    throw new Error('无编号文件被改动了: ' + JSON.stringify(plain.content));
  }
  return `5 段标题 → ${expected.join(' / ')}`;
});

await t('编号: 展开时认得 exclude / startFrom / 西式方案', async () => {
  const md = ['# 甲', '', '## 甲一', '', '### 甲一一', '', '# 乙'].join('\n');
  const base = await word.writeDocx({ markdown: md }, { title: 't' });
  const readWith = async (spec) => (await word.readDocx((await editDocx(base, [spec])).buf, {})).content.split('\n').filter((l) => l.trim());
  const multicol = await readWith({ op: 'set_numbering', style: 'multicol-1_1_1' });
  if (multicol.join('|') !== '1. 甲|1.1 甲一|1.1.1 甲一一|2. 乙') throw new Error('西式方案不对: ' + multicol.join('|'));
  const excluded = await readWith({ op: 'set_numbering', style: 'gongwen-1_1_1_1', exclude: ['Heading3'] });
  if (excluded.join('|') !== '一、 甲|（一） 甲一|甲一一|二、 乙') throw new Error('exclude 不对: ' + excluded.join('|'));
  const fromH2 = await readWith({ op: 'set_numbering', style: 'gongwen-1_1_1_1', startFrom: { Heading2: 1 } });
  if (fromH2.join('|') !== '甲|一、 甲一|（一） 甲一一|乙') throw new Error('startFrom 不对: ' + fromH2.join('|'));
  return '1. / 1.1 / 1.1.1、exclude、startFrom 都对';
});

await t('编号: 项目符号不展开成文字,编号列表展开', async () => {
  const buf = await word.writeDocx({ markdown: '- 甲\n- 乙\n\n1. 丙\n2. 丁' }, { title: 't' });
  const read = await word.readDocx(buf, {});
  if (read.content.includes('●') || read.content.includes('•')) throw new Error('项目符号被写进正文了: ' + JSON.stringify(read.content));
  if (!read.content.includes('1. 丙')) throw new Error('编号列表没展开: ' + JSON.stringify(read.content));
  if (!read.html.includes('<ul><li>甲</li>')) throw new Error('项目符号列表的 HTML 结构变了: ' + read.html);
  return '● 保持列表结构，1. 展开成文字';
});

await t('编号: formatCounter 覆盖中文/罗马/字母等数字格式', async () => {
  const cases = [
    ['decimal', 7, '7'], ['decimalZero', 7, '07'],
    ['chineseCounting', 1, '一'], ['chineseCounting', 10, '十'], ['chineseCounting', 11, '十一'],
    ['chineseCounting', 20, '二十'], ['chineseCounting', 21, '二十一'], ['chineseCounting', 101, '一百零一'],
    ['chineseCounting', 110, '一百一十'], ['ideographDigital', 12, '十二'],
    ['lowerLetter', 1, 'a'], ['upperLetter', 26, 'Z'], ['upperLetter', 27, 'AA'],
    ['lowerRoman', 4, 'iv'], ['upperRoman', 9, 'IX'],
    ['ordinal', 3, 'third'], ['cardinalText', 5, 'five'],
  ];
  for (const [fmt, value, want] of cases) {
    const got = formatCounter(fmt, value);
    if (got !== want) throw new Error(`${fmt}(${value}) = ${got}，应为 ${want}`);
  }
  return `${cases.length} 种数字格式`;
});

// ---------------------------------------------------------------------------
// 图片嵌入(PNG/JPEG/GIF/BMP 按魔数读尺寸,本地文件或 data: URL,不联网)
// ---------------------------------------------------------------------------
const imgDir = join(outDir, 'img');
await mkdir(imgDir, { recursive: true });

function crc32(buf) {
  let crc = 0xffffffff;
  for (const byte of buf) {
    crc ^= byte;
    for (let k = 0; k < 8; k += 1) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function pngChunk(type, data) {
  const head = Buffer.alloc(4);
  head.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([head, body, crc]);
}
function pngBuffer(w, h) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const raw = Buffer.concat(Array.from({ length: h }, () => Buffer.concat([Buffer.from([0]), Buffer.alloc(w * 3, 0x80)])));
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr), pngChunk('IDAT', deflateSync(raw)), pngChunk('IEND', Buffer.alloc(0)),
  ]);
}
/** 图片写入需要注入读取器：core 不碰文件系统，沙箱路径解析由工具层负责。 */
const testReadImage = async (src) => readFileSync(isAbsolute(src) ? src : join(process.cwd(), src));

const smallPng = join(imgDir, 'small.png');
const widePng = join(imgDir, 'wide.png');
const photoJpg = join(imgDir, 'photo.jpg');
const animGif = join(imgDir, 'anim.gif');
writeFileSync(smallPng, pngBuffer(40, 20));
writeFileSync(widePng, pngBuffer(1600, 400));
// SOI + SOF0(高 300、宽 200) + EOI
writeFileSync(photoJpg, Buffer.concat([Buffer.from([0xff, 0xd8]), Buffer.from([0xff, 0xc0, 0x00, 0x11, 0x08, 0x01, 0x2c, 0x00, 0xc8, 0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x00, 0x03, 0x11, 0x00]), Buffer.from([0xff, 0xd9])]));
const gifBytes = Buffer.alloc(20);
gifBytes.write('GIF89a', 0, 'latin1');
gifBytes.writeUInt16LE(10, 6);
gifBytes.writeUInt16LE(6, 8);
writeFileSync(animGif, gifBytes);

await t('图片: 按魔数读尺寸(PNG/JPEG/GIF/BMP)', async () => {
  const cases = [[smallPng, 'png', 40, 20], [widePng, 'png', 1600, 400], [photoJpg, 'jpg', 200, 300], [animGif, 'gif', 10, 6]];
  for (const [file, type, width, height] of cases) {
    const got = probeImage(readFileSync(file));
    if (got?.type !== type || got.width !== width || got.height !== height) {
      throw new Error(`${file} 读成 ${JSON.stringify(got)}，应为 ${type} ${width}×${height}`);
    }
  }
  if (probeImage(Buffer.from('这不是图片'))) throw new Error('文本被误认成图片');
  if (probeImage(Buffer.alloc(0))) throw new Error('空 buffer 被误认成图片');
  return 'PNG / JPEG / GIF 尺寸与魔数都对';
});

await t('图片: 尺寸换算(等比 + 缩到页内)', async () => {
  const image = { width: 1600, height: 400 };
  const capped = fitImageSize(image, { maxWidth: 600, maxHeight: 900 });
  if (capped.width !== 600 || capped.height !== 150) throw new Error('等比缩放不对: ' + JSON.stringify(capped));
  const byWidth = fitImageSize({ width: 200, height: 300 }, { widthHint: 100 });
  if (byWidth.width !== 100 || byWidth.height !== 150) throw new Error('只给宽度时没按比例算高度: ' + JSON.stringify(byWidth));
  const tall = fitImageSize({ width: 100, height: 4000 }, { maxWidth: 600, maxHeight: 900 });
  if (tall.height !== 900 || tall.width !== 23) throw new Error('超高图片没缩到页高内: ' + JSON.stringify(tall));
  const zero = fitImageSize({ width: 100, height: 100 }, { widthHint: 0.1 });
  if (zero.width < 1 || zero.height < 1) throw new Error('尺寸被算成 0: ' + JSON.stringify(zero));
  return '600×150 / 100×150 / 23×900';
});

await t('图片: <img> 嵌成真图片(不是 alt 文字)', async () => {
  const md = `标题\n\n![小图](${smallPng})\n\n<img src="${photoJpg}" alt="证照" width="100">\n\n<img src="${widePng}">\n\n<img src="${smallPng}" style="width:80px">\n\n<img src="data:image/png;base64,${readFileSync(smallPng).toString('base64')}">`;
  const buf = await word.writeDocx({ markdown: md }, { title: 't', readImage: testReadImage });
  const doc = await partText(buf, 'document.xml');
  const px = (emu) => Math.round(Number(emu) / 9525);
  const sizes = (doc.match(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g) || [])
    .map((m) => m.match(/\d+/g).map(px).join('x'));
  if (sizes.length !== 5) throw new Error(`嵌进去 ${sizes.length} 张，应为 5 张: ${sizes.join(' ')}`);
  if (sizes[0] !== '40x20') throw new Error('原始尺寸没保留: ' + sizes[0]);
  if (sizes[1] !== '100x150') throw new Error('width=100 没按比例算高度: ' + sizes[1]);
  if (!sizes[2].startsWith('60') || Number(sizes[2].split('x')[1]) > 900) throw new Error('超宽图片没缩到页内: ' + sizes[2]);
  if (sizes[3] !== '80x40') throw new Error('CSS width:80px 没生效: ' + sizes[3]);
  const { zip } = await openOoxml(buf, 'docx');
  const media = Object.keys(zip.files).filter((name) => /^word\/media\/.+\.(png|jpg)$/.test(name));
  // 同一张图用了几次只存一份(尺寸不同也只是显示参数不同)
  if (media.length !== 3) throw new Error('媒体部件数不对(同图应复用): ' + media.join());
  const rels = zip.file('word/_rels/document.xml.rels').asText();
  if ((rels.match(/media\//g) || []).length !== 3) throw new Error('关系数不对: ' + rels);
  if (!/descr="证照"/.test(doc) && !/name="证照"/.test(doc)) throw new Error('alt 文字没写进图片属性: ' + doc.slice(doc.indexOf('graphicFrame') - 200, doc.indexOf('graphicFrame')));
  if (!(await word.readDocx(buf, {})).content.includes('标题')) throw new Error('带图片的文档读不回来了');
  return `${sizes.join(' / ')}，媒体部件 ${media.length} 个`;
});

await t('图片: 行内图片与无 src 的 alt 退回', async () => {
  const buf = await word.writeDocx({ html: `<p>前<img src="${smallPng}" alt="图标">后</p>` }, { title: 't', readImage: testReadImage });
  const doc = await partText(buf, 'document.xml');
  if (!/<w:drawing>/.test(doc)) throw new Error('行内图片没嵌入');
  const order = ['前', 'w:drawing', '后'].map((key) => doc.indexOf(key));
  if (!(order[0] < order[1] && order[1] < order[2])) throw new Error('行内图片位置不对: ' + order.join());
  const fallback = await word.writeDocx({ html: '<p>前<img alt="占位说明">后</p>' }, { title: 't' });
  if (!(await word.readDocx(fallback, {})).content.includes('前占位说明后')) throw new Error('没有 src 时没退回 alt 文字');
  return '行内嵌入 + 无 src 退回 alt';
});

await t('图片: 远程/缺失/不支持格式/超大 都给出明确报错', async () => {
  const expect = async (html, code, needle) => {
    try {
      await word.writeDocx({ html }, { title: 't', readImage: testReadImage });
    } catch (err) {
      if (err.code !== code) throw new Error(`${code} 变成了 ${err.code}: ${err.message}`);
      if (!err.message.includes(needle)) throw new Error(`${code} 的说明里没有「${needle}」: ${err.message}`);
      return;
    }
    throw new Error(`${code} 没有报错`);
  };
  await expect('<img src="https://x.test/a.png">', 'IMAGE_REMOTE', '不联网');
  await expect(`<img src="${join(imgDir, '不存在.png')}">`, 'IMAGE_NOT_FOUND', '读不到图片文件');
  await expect('<img src="README.md">', 'IMAGE_UNSUPPORTED', '只支持 PNG');
  await expect('<img src="data:image/webp;base64,AAAA">', 'IMAGE_UNSUPPORTED', 'data URL');
  const huge = join(imgDir, 'huge.png');
  writeFileSync(huge, Buffer.concat([readFileSync(smallPng), Buffer.alloc(9 * 1024 * 1024)]));
  await expect(`<img src="${huge}">`, 'IMAGE_TOO_LARGE', '8 MB');
  const loaded = await loadImage(smallPng, testReadImage);
  if (loaded.type !== 'png') throw new Error('loadImage 没返回类型');
  return '4 类错误码 + 上限提示';
});

await t('表格: 读取公式本体(value / formula / both)', async () => {
  const ExcelJS = (await import('@wekanteam/exceljs')).default;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.addRow(['项目', '金额']);
  ws.addRow(['甲', 1]);
  ws.addRow(['乙', 2]);
  ws.getCell('A4').value = '合计';
  ws.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 3 };
  const path = join(outDir, 'formula.xlsx');
  writeFileSync(path, Buffer.from(await wb.xlsx.writeBuffer()));
  const bodyOf = async (opts) => (await office.opRead(path, opts, {})).content.split('```tsv')[1].split('```')[0].trim().split('\n').at(-1);
  /** @type {Array<[object, string]>} */
  const cases = [[{}, '合计\t3'], [{ formulas: 'formula' }, '合计\t=SUM(B2:B3)'], [{ formulas: 'both' }, '合计\t=SUM(B2:B3) → 3'], [{ formulas: true }, '合计\t=SUM(B2:B3)']];
  for (const [opts, want] of cases) {
    const got = await bodyOf(opts);
    if (got !== want) throw new Error(`${JSON.stringify(opts)} 读成 ${got}，应为 ${want}`);
  }
  let message = '';
  try { await office.opRead(path, { formulas: 'x' }, {}); } catch (err) { message = `${err.code}|${err.message}`; }
  if (!message.startsWith('INVALID_ARGS')) throw new Error('非法 formulas 没报错: ' + message);
  // 走 SheetJS 的老格式/二进制格式也认公式(.ods 会保留公式记录)
  const XLSX = await import('@e965/xlsx');
  const sheet = XLSX.utils.aoa_to_sheet([['项目', '数值'], ['甲', 100], ['乙', 200]]);
  sheet.A4 = { t: 's', v: '合计' };
  sheet.B4 = { t: 'n', f: 'SUM(B2:B3)', v: 300 };
  sheet['!ref'] = 'A1:B4';
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, '数据');
  const odsPath = join(outDir, 'formula.ods');
  writeFileSync(odsPath, Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'ods' })));
  const legacy = (await office.opRead(odsPath, { formulas: 'both' }, {})).content;
  if (!legacy.includes('=SUM(B2:B3) → 300')) throw new Error('.ods 的公式没读出来: ' + legacy.split('```tsv')[1]);
  return 'value / formula / both / true 简写 / .ods 都对';
});

await t('表格: .xlsb 能识别并交给 SheetJS(不再误判成损坏文件)', async () => {
  const XLSX = await import('@e965/xlsx');
  const sheet = XLSX.utils.aoa_to_sheet([['项目', '数值'], ['甲', 100], ['乙', 200]]);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, '数据');
  const path = join(outDir, 'binary.xlsb');
  writeFileSync(path, Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsb' })));
  const read = await office.opRead(path, {}, {});
  if (read.meta.reported.length !== 1 || !read.content.includes('乙\t200')) throw new Error('.xlsb 没读出来: ' + JSON.stringify(read.content));
  if (read.content.includes('0 个工作表')) throw new Error('.xlsb 被当成空工作簿了: ' + read.content);
  const query = await office.opQuery(path, {}, {});
  if (!query.content.includes('乙')) throw new Error('.xlsb 没算出来: ' + JSON.stringify(query.content).slice(0, 200));
  // 二进制工作簿的主部件是 xl/workbook.bin,不能被当成 .xlsx 交给 ExcelJS
  const { officePartOf } = await import('../lib/core/ooxml.js');
  const { zip } = await openOoxml(readFileSync(path), 'xlsb');
  const main = officePartOf(zip);
  if (main.kind !== 'xlsb' || main.path !== 'xl/workbook.bin') throw new Error('主部件识别不对: ' + JSON.stringify(main && { kind: main.kind, path: main.path }));
  return '主部件 xl/workbook.bin → SheetJS，读取与查询都正常';
});

await t('表格: 条件格式(9 种规则 + 差异样式)', async () => {
  const buf = await excel.buildWorkbook({ sheets: [{ name: 'S', rows: [['项目', '金额'], ['甲', 120], ['乙', 30]] }] });
  const out = await excel.editWorkbook(buf, [
    { op: 'conditional_format', sheet: 1, range: 'B2:B3', rules: [
      { type: 'cellIs', operator: 'greaterThan', value: 50, style: { color: '9C0006', fill: 'FFC7CE', bold: true } },
      { type: 'expression', formula: 'MOD(ROW(),2)=0', style: { italic: true } },
      { type: 'colorScale', colors: ['F8696B', 'FFEB84', '63BE7B'] },
      { type: 'dataBar', color: '63BE7B' },
      { type: 'iconSet', iconSet: '3TrafficLights1' },
      { type: 'top10', rank: 3, bottom: true },
      { type: 'aboveAverage', aboveAverage: false },
      { type: 'containsText', text: '甲', style: { fill: 'FFFF00' } },
      { type: 'timePeriod', timePeriod: 'last7Days' },
    ] },
  ]);
  const xml = await partText(out.buf, 'xl/worksheets/sheet1.xml');
  const ruleTypes = [...xml.matchAll(/<cfRule type="([A-Za-z0-9]+)"/g)].map((m) => m[1]);
  const want = ['cellIs', 'expression', 'colorScale', 'dataBar', 'iconSet', 'top10', 'aboveAverage', 'containsText', 'timePeriod'];
  if (ruleTypes.join() !== want.join()) throw new Error('规则类型不对: ' + ruleTypes.join());
  const priorities = (xml.match(/priority="(\d+)"/g) || []).map((m) => Number(m.replaceAll(/\D/g, '')));
  if (priorities.join() !== '1,2,3,4,5,6,7,8,9') throw new Error('优先级不是依次递增: ' + priorities.join());
  if (!/sqref="B2:B3"/.test(xml)) throw new Error('作用区域不对');
  if (!/<colorScale>[\s\S]*?<cfvo type="percentile" val="50"\/>/.test(xml)) throw new Error('三色阶没写百分位中点');
  if (!/<dxfs count="3">/.test(await partText(out.buf, 'xl/styles.xml'))) throw new Error('差异样式(dxf)数量不对');
  // 差异格式的填充走 bgColor(dxf 约定)
  const styles = await partText(out.buf, 'xl/styles.xml');
  if (!/<dxf>[\s\S]*?<bgColor rgb="FFFFC7CE"\/>/.test(styles)) throw new Error('dxf 填充没写 bgColor: ' + styles.slice(styles.indexOf('<dxfs'), styles.indexOf('</dxfs>')));
  return `${ruleTypes.length} 种规则，优先级 1–9，dxf 3 个`;
});

await t('表格: 条件格式的参数校验', async () => {
  const buf = await excel.buildWorkbook({ sheets: [{ name: 'S', rows: [['项目', '金额'], ['甲', 1]] }] });
  /** @type {Array<[object, string]>} */
  const bad = [
    [{ op: 'conditional_format', range: 'B2', rule: { type: 'unknown' } }, 'conditional_format 的 type'],
    [{ op: 'conditional_format', range: 'B2', rule: { type: 'cellIs', operator: 'bigger', value: 1 } }, 'operator'],
    [{ op: 'conditional_format', range: 'B2', rule: { type: 'cellIs', operator: 'between', value: 1 } }, '两个阈值'],
    [{ op: 'conditional_format', range: 'B2', rule: { type: 'cellIs' } }, '需要 value'],
    [{ op: 'conditional_format', range: 'B2', rule: { type: 'colorScale', colors: ['FF0000'] } }, '2 个或 3 个颜色'],
    [{ op: 'conditional_format', range: 'B2:B', rule: { type: 'expression', formula: 'A1>0' } }, 'range 形如'],
    [{ op: 'conditional_format', range: 'B2', rule: { type: 'containsText' } }, 'text'],
  ];
  for (const [op, needle] of bad) {
    let message = '';
    try { await excel.editWorkbook(buf, [op]); } catch (err) { message = err.message; }
    if (!message.includes(needle)) throw new Error(`${JSON.stringify(op)} 的报错不对: ${message}`);
  }
  return `${bad.length} 种参数错误都有明确提示`;
});

await t('表格: 数据验证(下拉 / 区间 / 公式)', async () => {
  const buf = await excel.buildWorkbook({ sheets: [{ name: 'S', rows: [['项目', '金额', '状态'], ['甲', 1, '']] }] });
  const out = await excel.editWorkbook(buf, [
    { op: 'data_validation', sheet: 1, range: 'C2:C100', rule: { type: 'list', values: ['待办', '进行中', '已完成'], promptTitle: '选择状态', prompt: '从下拉里选', errorTitle: '值不对', error: '只能选下拉里的值' } },
    { op: 'data_validation', sheet: 1, range: 'D2:D9', rule: { type: 'list', source: '=Sheet2!$A$1:$A$5' } },
    { op: 'data_validation', sheet: 1, range: 'B2:B100', rule: { type: 'whole', operator: 'between', value: 1, value2: 10, error: '只能 1-10' } },
    { op: 'data_validation', sheet: 1, range: 'E2:E9', rule: { type: 'custom', formula: 'ISNUMBER(E2)' } },
  ]);
  const xml = await partText(out.buf, 'xl/worksheets/sheet1.xml');
  if (!/<dataValidations count="4">/.test(xml)) throw new Error('验证条数不对: ' + (xml.match(/<dataValidations[^>]*>/) || [])[0]);
  if (!/<formula1>&quot;待办,进行中,已完成&quot;<\/formula1>/.test(xml)) throw new Error('下拉候选没写成字面量列表');
  if (!/<formula1>Sheet2!\$A\$1:\$A\$5<\/formula1>/.test(xml)) throw new Error('区域来源没去掉等号');
  if (!/<formula1>1<\/formula1><formula2>10<\/formula2>/.test(xml)) throw new Error('区间上下限没写全');
  if (!/promptTitle="选择状态"|errorTitle="值不对"/.test(xml)) throw new Error('提示语没写进去');
  /** @type {Array<[object, string]>} */
  const bad = [
    [{ op: 'data_validation', range: 'C2', rule: { type: 'list' } }, 'values'],
    [{ op: 'data_validation', range: 'C2', rule: { type: 'list', values: [] } }, '空数组'],
    [{ op: 'data_validation', range: 'C2', rule: { type: 'unknown' } }, 'data_validation 的 type'],
    [{ op: 'data_validation', range: 'C2', rule: { type: 'whole' } }, '需要 value'],
  ];
  for (const [op, needle] of bad) {
    let message = '';
    try { await excel.editWorkbook(buf, [op]); } catch (err) { message = err.message; }
    if (!message.includes(needle)) throw new Error(`${JSON.stringify(op)} 的报错不对: ${message}`);
  }
  // 再改两次值:条件格式与数据验证要还在,而且不能越写越多(exceljs 会把区域拆散再合并错)
  let again = out.buf;
  for (const cell of ['A3', 'A4', 'A5']) {
    again = (await excel.editWorkbook(again, [{ op: 'set_value', sheet: 1, ref: cell, value: '乙' }])).buf;
  }
  const xml2 = await partText(again, 'xl/worksheets/sheet1.xml');
  if (!/<dataValidations count="4">/.test(xml2)) throw new Error('后续编辑把数据验证弄丢了/变多了: ' + (xml2.match(/<dataValidations[^>]*>/) || [])[0]);
  if (!/sqref="C2:C100"/.test(xml2)) throw new Error('区域被写成了别的形状: ' + (xml2.match(/sqref="[^"]*"/g) || []).join());
  return '4 条验证 + 4 类参数错误，反复编辑不丢不涨';
});

await t('编号: 幂等——重复调用原地替换,不新增重复定义', async () => {
  const md = ['# 概述', '', '- 要点一', '', '## 细节', '', '- 细节点'].join('\n');
  let buf = await word.writeDocx({ markdown: md }, { title: 't' });
  const snapshot = async () => {
    const { zip } = await openOoxml(buf, 'docx');
    const numbering = zip.file('word/numbering.xml').asText();
    const doc = zip.file('word/document.xml').asText();
    return {
      abstract: (numbering.match(/<w:abstractNum /g) || []).length,
      nums: (numbering.match(/<w:num /g) || []).length,
      numIds: [...new Set(doc.match(/<w:numId w:val="\d+"\/>/g) || [])].join(),
      hasNsid: /<w:nsid w:val="0D5F0001"\/>/.test(numbering),
    };
  };
  const first = await editDocx(buf, [{ op: 'set_numbering', scope: 'all' }]);
  buf = first.buf;
  const a = await snapshot();
  if (!a.hasNsid) throw new Error('我们那份 abstractNum 没带 nsid,无法识别与替换');
  if (first.changes[0].includes('更新已有编号定义')) throw new Error('第一次不该说「更新已有」: ' + first.changes[0]);
  for (let i = 0; i < 2; i += 1) buf = (await editDocx(buf, [{ op: 'set_numbering', scope: 'all' }])).buf;
  const b = await snapshot();
  if (b.abstract !== a.abstract || b.nums !== a.nums) throw new Error(`重复调用新增了重复定义: ${JSON.stringify(a)} → ${JSON.stringify(b)}`);
  if (b.numIds !== a.numIds) throw new Error('numId 变了,段落里的引用会失效: ' + a.numIds + ' → ' + b.numIds);
  const third = await editDocx(buf, [{ op: 'set_numbering', scope: 'all' }]);
  if (!third.changes[0].includes('更新已有编号定义')) throw new Error('幂等替换没有在结果里说明: ' + third.changes[0]);
  if (a.abstract === 0) throw new Error('测试前提不成立');
  return `3 次调用后仍是 abstractNum ${b.abstract} / num ${b.nums}，numId 稳定 ${b.numIds}`;
});

await t('编号: exclude 指定不参与的标题样式', async () => {
  const md = ['# 一、总体情况', '', '## （一）进展', '', '- 细节甲', '', '## （二）问题', '', '- 问题甲'].join('\n');
  const buf = await word.writeDocx({ markdown: md }, { title: 't' });
  const { buf: out, changes } = await editDocx(buf, [{ op: 'set_numbering', scope: 'all', exclude: ['Heading1'], startFrom: { Heading2: 1 } }]);
  const numbering = await partText(out, 'numbering.xml');
  const chain = (numbering.slice(numbering.indexOf('<w:nsid w:val="0D5F0001"')).match(/<w:lvl w:ilvl="\d+">[\s\S]*?<\/w:lvl>/g) || [])
    .map((block) => (block.match(/<w:pStyle w:val="([^"]+)"\/>/) || [])[1]);
  if (chain[0] !== 'Heading2' || chain[1] !== 'Heading3') throw new Error('链没有从 Heading2 开始: ' + chain.slice(0, 3).join(','));
  if (chain.includes('Heading1')) throw new Error('Heading1 被 exclude 了却还在链上: ' + chain.join(','));
  if (!/正文列表 2 段/.test(changes[0])) throw new Error('H2 下的列表应被编号: ' + changes[0]);
  const back = await word.readDocx(out, {});
  if (!back.content.includes('细节甲')) throw new Error('文档读不回来了');
  return `链首 = ${chain[0]}；${changes[0].replace(/^.*?：/, '')}`;
});

await t('编号: startFrom 指定起始样式与起始数字', async () => {
  const md = ['# 一、总体情况', '', '- 要点甲', '', '## （一）进展', '', '- 细节甲'].join('\n');
  const buf = await word.writeDocx({ markdown: md }, { title: 't' });
  // 起始数字:第一级从 5 开始
  const from = await editDocx(buf, [{ op: 'set_numbering', scope: 'all', startFrom: { Heading1: 5 } }]);
  const numbering = await partText(from.buf, 'numbering.xml');
  const block = numbering.slice(numbering.indexOf('<w:nsid w:val="0D5F0001"'));
  const firstLevel = (block.match(/<w:lvl w:ilvl="0">[\s\S]*?<\/w:lvl>/) || [])[0];
  if (!/<w:start w:val="5"\/>/.test(firstLevel)) throw new Error('第一级起始数字没生效: ' + firstLevel);
  if (!/<w:lvlText w:val="%1\."\/>/.test(firstLevel)) throw new Error('第一级编号文字不对: ' + firstLevel);
  // 从 Heading2 起:比它浅的标题样式不参与,其下的列表被跳过并如实报出
  const skip = await editDocx(buf, [{ op: 'set_numbering', scope: 'all', startFrom: { Heading2: 1 } }]);
  if (!/跳过 1 段/.test(skip.changes[0])) throw new Error('H1 下的列表应被跳过并报出: ' + skip.changes[0]);
  const doc = await partText(skip.buf, 'document.xml');
  if (!/<w:numPr><w:ilvl w:val="1"\/>/.test(doc)) throw new Error('H2 下的列表应挂 ilvl=1: ' + doc.slice(0, 200));
  return `${from.changes[0].replace(/^.*?：/, '')}；${skip.changes[0].replace(/^.*?：/, '')}`;
});

await t('编号: exclude / startFrom 的参数错误', async () => {
  const buf = await word.writeDocx({ markdown: '# 概述\n\n- 要点一' }, { title: 't' });
  const fail = async (ops) => { try { await editDocx(buf, ops); return { code: '', message: '' }; } catch (e) { return { code: e.code, message: e.message }; } };
  const cases = [
    { label: 'exclude 认不出', spec: { exclude: ['MyStyle'] }, want: 'INVALID_ARGS', text: '认不出' },
    { label: 'exclude 非数组', spec: { exclude: 'Heading1' }, want: 'INVALID_ARGS', text: '需为数组' },
    { label: 'startFrom 两个键', spec: { startFrom: { Heading1: 1, Heading2: 1 } }, want: 'INVALID_ARGS', text: '只能给一个样式' },
    { label: 'startFrom 认不出', spec: { startFrom: { 大标题: 1 } }, want: 'INVALID_ARGS', text: '认不出' },
    { label: 'startFrom 数字非法', spec: { startFrom: { Heading1: 0 } }, want: 'INVALID_ARGS', text: '需为 1–9999' },
    { label: 'exclude 排除所有标题', spec: { exclude: ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6', 'Heading7', 'Heading8', 'Heading9'] }, want: 'INVALID_ARGS', text: '没有可编号的级别' },
  ];
  for (const c of cases) {
    const got = await fail([{ op: 'set_numbering', scope: 'all', ...c.spec }]);
    if (got.code !== c.want || !got.message.includes(c.text)) {
      throw new Error(`${c.label} 期望 ${c.want}/「${c.text}」，实际 ${got.code || '没报错'}/${got.message}`);
    }
  }
  return `${cases.length} 种参数错误`;
});

await t('编号: numbering.xml 缺失时自动补出部件与关系', async () => {
  const buf = await word.writeDocx({ markdown: '# 概述\n\n- 要点一\n- 要点二' }, { title: 't' });
  const { zip } = await openOoxml(buf, 'docx');
  zip.remove('word/numbering.xml');
  const stripped = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  const { buf: out } = await editDocx(stripped, [{ op: 'set_numbering', scope: 'all' }]);
  const { zip: after } = await openOoxml(out, 'docx');
  const numbering = after.file('word/numbering.xml');
  if (!numbering) throw new Error('numbering.xml 没被补出来');
  if (!numbering.asText().includes('<w:abstractNum')) throw new Error('补出来的 numbering.xml 里没有 abstractNum');
  if (!/numbering\.xml/.test(after.file('word/_rels/document.xml.rels').asText())) throw new Error('关系没补');
  if (!/word\/numbering\.xml/.test(after.file('[Content_Types].xml').asText())) throw new Error('Content_Types 没补');
  if (!(await word.readDocx(out, {})).content.includes('要点一')) throw new Error('文档读不回来');
  return '部件 + 关系 + Content_Types 三件套';
});

await t('编号: 参数错误与无可编号段落', async () => {
  const buf = await word.writeDocx({ markdown: '# 概述\n\n普通正文\n\n- 要点一' }, { title: 't' });
  const codeOf = async (ops) => { try { await editDocx(buf, ops); return ''; } catch (e) { return e.code; } };
  const msgOf = async (ops) => { try { await editDocx(buf, ops); return ''; } catch (e) { return e.message; } };
  const cases = [
    { label: 'style 不在预置里', ops: [{ op: 'set_numbering', style: 'multicol-1_1_1_1' }], want: 'INVALID_ARGS' },
    { label: 'linkToHeading 非布尔', ops: [{ op: 'set_numbering', linkToHeading: 'yes' }], want: 'INVALID_ARGS' },
    { label: 'scope 非法', ops: [{ op: 'set_numbering', scope: 'each' }], want: 'INVALID_ARGS' },
    { label: '表格内没有可编号段落', ops: [{ op: 'set_numbering', scope: 'table' }], want: 'PARAGRAPH_NOT_FOUND' },
  ];
  for (const { label, ops, want } of cases) {
    const got = await codeOf(ops);
    if (got !== want) throw new Error(label + ' 期望 ' + want + '，实际 ' + (got || '没报错'));
  }
  const styleMsg = await msgOf([{ op: 'set_numbering', style: 'x' }]);
  if (!styleMsg.includes('multicol-1_1_1')) throw new Error('style 报错没列出预置值: ' + styleMsg);
  // linkToHeading=false:标题显式编号
  const { buf: out, changes } = await editDocx(buf, [{ op: 'set_numbering', scope: 'headings', linkToHeading: false }]);
  const doc = await partText(out, 'document.xml');
  if (!/<w:numPr><w:ilvl w:val="0"\/>/.test(doc)) {
    throw new Error('linkToHeading=false 时标题应显式编号: ' + changes.join());
  }
  return `${cases.length} 种错误 + linkToHeading=false 路径`;
});

await t('角色选择: body 只改正文,headings 只改标题,table 只改表格内', async () => {
  const md = [
    '# 标题一', '', '正文第一段。', '',
    '| 项目 | 金额 |', '| --- | --- |', '| 收入 | 1200 |', '',
    '## 标题二', '', '正文第二段。',
  ].join('\n');
  const buf = await word.writeDocx({ markdown: md }, { title: 't' });
  const paragraphsOf = async (b) => {
    const doc = await partText(b, 'document.xml');
    return scanParagraphs(doc).map((p) => ({ text: p.text, heading: p.headingLevel, inTable: p.inTable }));
  };
  const all = await paragraphsOf(buf);
  const headings = all.filter((p) => p.heading).length;
  const inTable = all.filter((p) => p.inTable).length;
  if (headings !== 2 || inTable < 4) throw new Error(`测试前提不成立: 标题 ${headings}、表格内 ${inTable}`);

  const bodyOnly = await editDocx(buf, [{ op: 'set_style', scope: 'body', sizePt: 14, font: '仿宋' }]);
  if (!bodyOnly.changes[0].includes(`${all.length - headings} 段`)) {
    throw new Error(`body 应命中 ${all.length - headings} 段: ${bodyOnly.changes[0]}`);
  }
  const headingsOnly = await editDocx(buf, [{ op: 'set_style', scope: 'headings', sizePt: 22, font: '黑体' }]);
  if (!headingsOnly.changes[0].includes(`${headings} 段`)) throw new Error(`headings 应命中 ${headings} 段: ${headingsOnly.changes[0]}`);
  const tableOnly = await editDocx(buf, [{ op: 'set_style', scope: 'table', sizePt: 12, font: '楷体' }]);
  if (!tableOnly.changes[0].includes(`${inTable} 段`)) throw new Error(`table 应命中 ${inTable} 段: ${tableOnly.changes[0]}`);

  // 角色 + 区间叠加:第 7–8 段是「标题二 + 正文第二段」,scope=body 应只命中其中 1 段
  const lastBody = all.length;
  const scoped = await editDocx(buf, [{ op: 'set_style', scope: 'body', from: lastBody - 1, to: lastBody, sizePt: 13 }]);
  if (!scoped.changes[0].includes('1 段')) throw new Error('角色 + 区间叠加没生效: ' + scoped.changes[0]);

  // 只改正文时,标题的 Heading 样式必须原封不动
  const bodyDoc = await partText(bodyOnly.buf, 'document.xml');
  if (!/w:pStyle w:val="Heading1"/.test(bodyDoc)) throw new Error('改正文把标题样式抹掉了');
  const bodyText = scanParagraphs(bodyDoc).filter((p) => p.headingLevel > 0).map((p) => p.text);
  if (bodyText.join() !== '标题一,标题二') throw new Error('标题内容被改了: ' + bodyText.join());
  return `全部 ${all.length} 段 = 标题 ${headings} + 正文 ${all.length - headings}（含表格内 ${inTable}）`;
});

await t('表格: 框线与列宽规格(含按内容自动分配)', async () => {
  const spec = normalizeTableSpec({ borders: 'three-line', headerShading: 'F2F2F2', headerBold: false, columnWidthMode: 'auto', align: 'center' });
  if (spec.borders !== 'three-line' || spec.headerShading !== 'f2f2f2' || spec.headerBold !== false) throw new Error('表格规格不对: ' + JSON.stringify(spec));
  if (spec.columnMode !== 'auto' || spec.align !== 'center') throw new Error('列宽模式/对齐不对: ' + JSON.stringify(spec));
  const manual = normalizeTableSpec({ columnWidths: [3, 5, 2] });
  if (manual.columnMode !== 'manual') throw new Error('给了 columnWidths 应自动切到 manual');
  if (manual.columnPercents.join() !== '30,50,20') throw new Error('百分比应按比例归一: ' + manual.columnPercents.join());
  const auto = autoColumnPercents([2, 10, 6]);
  if (auto.length !== 3 || auto.reduce((a, b) => a + b, 0) !== 100) throw new Error('自动列宽应为和为 100 的数组: ' + auto);
  if (!(auto[1] > auto[2] && auto[2] > auto[0])) throw new Error('应按内容长度排序: ' + auto);
  if (autoColumnPercents([500, 1, 1])[0] > 92) throw new Error('超长列不该把其他列挤没: ' + autoColumnPercents([500, 1, 1]));
  if (normalizeTableSpec({ headerShading: 'none' }).headerShading !== null) throw new Error('headerShading:"none" 应表示去掉底纹');
  const bad = [
    { spec: { borders: 'double' }, want: 'borders 可为' },
    { spec: { columnWidthMode: 'manual' }, want: '需要同时给 columnWidths' },
    { spec: { columnWidths: [] }, want: 'columnWidths 需为' },
    { spec: { headerShading: 'red' }, want: 'headerShading 需为' },
    { spec: { align: 'middle' }, want: 'align 可为' },
  ];
  for (const { spec: spec2, want } of bad) {
    let msg = '';
    try { normalizeTableSpec(spec2); } catch (e) { msg = e.message; }
    if (!msg.includes(want)) throw new Error(JSON.stringify(spec2) + ' 的报错不对: ' + msg);
  }
  return '三线表 / 自动列宽 / 5 种非法规格';
});

await t('表格: 扫描行列与内容', async () => {
  const buf = await word.writeDocx({ markdown: '| A | B |\n| --- | --- |\n| 1 | 2 |\n\n中间段落\n\n| C |\n| --- |\n| 3 |' }, { title: 't' });
  const { zip } = await openOoxml(buf, 'docx');
  const tables = scanTables(zip.file('word/document.xml').asText());
  if (tables.length !== 2) throw new Error('应扫到两个表格: ' + tables.length);
  if (tables[0].rows !== 2 || tables[0].cols !== 2) throw new Error('第一个表应为 2×2: ' + JSON.stringify(tables[0]));
  if (tables[1].rows !== 2 || tables[1].cols !== 1) throw new Error('第二个表应为 2×1: ' + JSON.stringify(tables[1]));
  return tables.map((t) => t.rows + '×' + t.cols).join(' / ');
});

await t('表格: set_table 三线表 / 表头底纹 / 居中', async () => {
  const buf = await word.writeDocx({ markdown: '| 姓名 | 部门 | 备注 |\n| --- | --- | --- |\n| 张三 | 技术研发中心 | 负责人 |' }, { title: 't' });
  const { buf: out, changes } = await editDocx(buf, [
    { op: 'set_table', borders: 'three-line', headerShading: 'F2F2F2', headerBold: true, align: 'center' },
  ]);
  const doc = await partText(out, 'document.xml');
  const markers = [
    ['上下粗线', /<w:top w:val="single" w:sz="12"/.test(doc) && /<w:bottom w:val="single" w:sz="12"/.test(doc)],
    ['无竖线', /<w:insideV w:val="none"/.test(doc) && /<w:left w:val="none"/.test(doc)],
    ['表头下细线', /<w:tcBorders><w:bottom w:val="single" w:sz="6"/.test(doc)],
    ['表头底纹', doc.includes('w:fill="f2f2f2"')],
    ['表头加粗', doc.includes('<w:b/>')],
    ['表格居中', /<w:jc w:val="center"\/>/.test(doc)],
  ];
  const bad = markers.filter(([, ok]) => !ok).map(([n]) => n);
  if (bad.length) throw new Error('缺失: ' + bad.join(', '));
  if (!changes[0].includes('1 个表格')) throw new Error('结果描述不对: ' + changes[0]);
  const md = htmlToMarkdown((await word.readDocx(out, {})).html);
  if (!md.includes('张三')) throw new Error('内容被破坏');
  return changes[0];
});

await t('表格: 列宽 auto 按内容 / manual 百分比', async () => {
  const md = '| 姓名 | 部门 | 备注 |\n| --- | --- | --- |\n| 张三 | 技术研发中心 | 负责人 |';
  const widthsOf = async (tableSpec) => {
    const buf = await word.writeDocx({ markdown: md }, { title: 't', style: { table: tableSpec } });
    const doc = await partText(buf, 'document.xml');
    return (doc.match(/<w:gridCol w:w="(\d+)"\/>/g) || []).map((s) => Number(s.replaceAll(/\D/g, '')));
  };
  const autoWidths = await widthsOf({ columnWidthMode: 'auto' });
  if (autoWidths.length !== 3) throw new Error('列数不对: ' + autoWidths);
  if (!(autoWidths[1] > autoWidths[2] && autoWidths[2] > autoWidths[0])) throw new Error('auto 没按内容分配: ' + autoWidths);
  const manualWidths = await widthsOf({ columnWidths: [20, 50, 30] });
  const sum = manualWidths.reduce((a, b) => a + b, 0);
  const ratios = manualWidths.map((w) => Math.round((w / sum) * 100));
  if (ratios.join() !== '20,50,30') throw new Error('manual 比例不对: ' + ratios.join());
  if (!(await partText(await word.writeDocx({ markdown: md }, { title: 't', style: { table: { columnWidths: [50, 25, 25] } } }), 'document.xml')).includes('w:tblLayout w:type="fixed"')) {
    throw new Error('指定列宽时应用固定布局');
  }
  return `auto ${autoWidths.join('/')} · manual ${ratios.join('/')}`;
});

await t('表格: 增删行列与合并单元格', async () => {
  const buf = await word.writeDocx({ markdown: '| A | B | C |\n| --- | --- | --- |\n| a1 | b1 | c1 |\n| a2 | b2 | c2 |' }, { title: 't' });
  const sizeOf = async (b) => {
    const doc = await partText(b, 'document.xml');
    const t = scanTables(doc)[0];
    return `${t.rows}×${t.cols}`;
  };
  const r1 = await editDocx(buf, [{ op: 'insert_table_row', at: 2, count: 1 }]);
  if (await sizeOf(r1.buf) !== '4×3') throw new Error('插行后应为 4×3: ' + await sizeOf(r1.buf));
  const r2 = await editDocx(r1.buf, [{ op: 'delete_table_row', at: 1, count: 1 }]);
  if (await sizeOf(r2.buf) !== '3×3') throw new Error('删行后应为 3×3: ' + await sizeOf(r2.buf));
  const r3 = await editDocx(r2.buf, [{ op: 'insert_table_column', at: 2, count: 2 }]);
  if (await sizeOf(r3.buf) !== '3×5') throw new Error('插列后应为 3×5: ' + await sizeOf(r3.buf));
  const r4 = await editDocx(r3.buf, [{ op: 'delete_table_column', at: 3, count: 2 }]);
  if (await sizeOf(r4.buf) !== '3×3') throw new Error('删列后应为 3×3: ' + await sizeOf(r4.buf));
  const r5 = await editDocx(r4.buf, [{ op: 'merge_table_cells', range: 'A1:B2' }]);
  const doc = await partText(r5.buf, 'document.xml');
  if (!/<w:gridSpan w:val="2"\/>/.test(doc) || !/<w:vMerge w:val="restart"\/>/.test(doc) || !/<w:vMerge\/>/.test(doc)) {
    throw new Error('合并单元格的 gridSpan/vMerge 不对');
  }
  if ((await sizeOf(r5.buf)) !== '3×3') throw new Error('合并后网格列数不该变: ' + await sizeOf(r5.buf));
  // 合并保留左上角内容:单独用一份干净表格验证,避免被前面的增删行影响
  const clean = await word.writeDocx({ markdown: '| 甲 | 乙 | 丙 |\n| --- | --- | --- |\n| 1 | 2 | 3 |' }, { title: 't' });
  const merged = await editDocx(clean, [{ op: 'merge_table_cells', range: 'A1:B1' }]);
  const mergedMd = htmlToMarkdown((await word.readDocx(merged.buf, {})).html);
  if (!mergedMd.includes('甲')) throw new Error('合并后左上角内容丢了: ' + mergedMd);
  if (mergedMd.includes('乙')) throw new Error('被合并掉的单元格内容还在: ' + mergedMd);
  if (!mergedMd.includes('丙')) throw new Error('区域外的单元格被动了: ' + mergedMd);
  return `${await sizeOf(buf)} → 插行/删行/插列/删列/合并 各一步`;
});

await t('表格: 表头跨页重复 / 垂直对齐 / 行高 / 禁止断行', async () => {
  const md = '| 项目 | 说明 |\n| --- | --- |\n| 甲 | 说明甲 |\n| 乙 | 说明乙 |';
  // 写路径
  const written = await partText(await word.writeDocx({ markdown: md }, {
    title: 't',
    style: { table: { repeatHeader: true, cellVerticalAlign: 'center', rowHeightPt: 24, cantSplit: true } },
  }), 'document.xml');
  // 改路径
  const edited = await partText((await editDocx(await word.writeDocx({ markdown: md }, { title: 't' }), [
    { op: 'set_table', table: 1, repeatHeader: true, cellVerticalAlign: 'bottom', rowHeightPt: 30, cantSplit: true },
  ])).buf, 'document.xml');
  const checks = [
    ['写:表头重复', /<w:tblHeader\/>/.test(written)],
    ['写:垂直居中', /<w:vAlign w:val="center"\/>/.test(written)],
    ['写:行高 24pt=480', /<w:trHeight w:val="480" w:hRule="atLeast"\/>/.test(written)],
    ['写:禁止断行', /<w:cantSplit\/>/.test(written)],
    ['改:表头重复', /<w:tblHeader\/>/.test(edited)],
    ['改:底部对齐', /<w:vAlign w:val="bottom"\/>/.test(edited)],
    ['改:行高 30pt=600', /<w:trHeight w:val="600"/.test(edited)],
  ];
  const bad = checks.filter(([, ok]) => !ok).map(([n]) => n);
  if (bad.length) throw new Error('缺失: ' + bad.join(', '));
  // tblHeader 只加在表头行:第一行有、最后一行没有
  const rows = written.match(/<w:tr>[\s\S]*?<\/w:tr>/g) || [];
  if (rows.length !== 3 || !/<w:tblHeader\/>/.test(rows[0]) || /<w:tblHeader\/>/.test(rows[2])) {
    throw new Error('tblHeader 不该出现在非表头行');
  }
  const bad2 = [
    { spec: { cellVerticalAlign: 'middle' }, want: 'cellVerticalAlign 可为' },
    { spec: { rowHeightPt: 9999 }, want: 'rowHeightPt 需为' },
    { spec: { repeatHeader: 'yes' }, want: 'repeatHeader 需为' },
  ];
  for (const { spec, want } of bad2) {
    let message = '';
    try { normalizeTableSpec(spec); } catch (e) { message = e.message; }
    if (!message.includes(want)) throw new Error(JSON.stringify(spec) + ' 的报错不对: ' + message);
  }
  return `${checks.length} 项（写/改两条路径）`;
});

await t('表格: unmerge_table_cells 取消合并', async () => {
  const md = '| A | B | C |\n| --- | --- | --- |\n| a1 | b1 | c1 |\n| a2 | b2 | c2 |';
  const buf = await word.writeDocx({ markdown: md }, { title: 't' });
  const merged = await editDocx(buf, [{ op: 'merge_table_cells', table: 1, range: 'A1:B2' }]);
  const mergedDoc = await partText(merged.buf, 'document.xml');
  if (!/gridSpan w:val="2"/.test(mergedDoc) || !/vMerge/.test(mergedDoc)) throw new Error('测试前提不成立: 没合并成功');
  const before = scanTables(mergedDoc)[0];
  const unmerged = await editDocx(merged.buf, [{ op: 'unmerge_table_cells', table: 1, range: 'A1:B2' }]);
  if (!unmerged.changes[0].includes('取消合并 A1:B2')) throw new Error('结果描述不对: ' + unmerged.changes[0]);
  const doc = await partText(unmerged.buf, 'document.xml');
  if (/gridSpan|vMerge/.test(doc)) throw new Error('取消合并后仍有 gridSpan/vMerge');
  const after = scanTables(doc)[0];
  if (after.rows !== before.rows || after.cols !== before.cols) {
    throw new Error(`网格变了: ${before.rows}×${before.cols} → ${after.rows}×${after.cols}`);
  }
  // 区域外的内容不受影响
  const mdOut = htmlToMarkdown((await word.readDocx(unmerged.buf, {})).html);
  for (const keep of ['c1', 'a2', 'b2', 'c2']) {
    if (!mdOut.includes(keep)) throw new Error('取消合并误伤了 ' + keep + ': ' + mdOut);
  }
  // 没合并过的区域要报错,而不是静默成功
  let message = '';
  try { await editDocx(buf, [{ op: 'unmerge_table_cells', table: 1, range: 'A1:B1' }]); } catch (e) { message = `${e.code}|${e.message}`; }
  if (!message.startsWith('TABLE_NOT_MERGED')) throw new Error('未合并区域的报错不对: ' + message);
  return `${before.rows}×${before.cols} 网格保持不变，区域外内容完好`;
});

await t('表格: delete_table 删除整表(含删空后的 0 行残留)', async () => {
  const md = ['# 附录 A', '', '| 项目 | 说明 |', '| --- | --- |', '| 结构 | 说明文字 |', '',
    '# 附录 B', '', '| 列1 | 列2 |', '| --- | --- |', '| r1 | x |', '| r2 | y |', '', '结语段落。'].join('\n');
  const buf = await word.writeDocx({ markdown: md }, { title: 't' });
  const rowsOf = async (b, index) => {
    const doc = await partText(b, 'document.xml');
    const t = scanTables(doc)[index - 1];
    return t ? t.rows : -1;
  };
  // 先把附录 B 的表格行删光(3 行),留下 0 行空表 —— 复现真实场景
  const emptied = await editDocx(buf, [{ op: 'delete_table_row', table: 2, at: 1, count: 3 }]);
  if (await rowsOf(emptied.buf, 2) !== 0) throw new Error('测试前提不成立: 行没删光');
  // 0 行表上删行/删列都应给出「用 delete_table」的提示
  for (const op of ['delete_table_row', 'delete_table_column']) {
    let message = '';
    try { await editDocx(emptied.buf, [{ op, table: 2, at: 1 }]); } catch (e) { message = e.message; }
    if (!message.includes('delete_table')) throw new Error(`${op} 在 0 行表上的报错应提示 delete_table: ${message}`);
  }
  // delete_table 清掉残留
  const { buf: out, changes } = await editDocx(emptied.buf, [{ op: 'delete_table', table: 2 }]);
  if (!changes[0].includes('删除表格 2')) throw new Error('结果描述不对: ' + changes[0]);
  if (await rowsOf(out, 2) !== -1) throw new Error('残留表格没删掉');
  const mdOut = htmlToMarkdown((await word.readDocx(out, {})).html);
  if (!mdOut.includes('结构')) throw new Error('附录 A 的表格被误删');
  if (!mdOut.includes('结语段落。')) throw new Error('结语段落丢失');
  if (!mdOut.includes('附录 A')) throw new Error('标题丢失');
  return changes[0];
});

await t('表格: delete_table 的 onlyEmpty 一次清掉所有 0 行残留', async () => {
  const md = ['# A', '', '| x |', '| --- |', '| 1 |', '', '正文段落', '', '# B', '', '| y |', '| --- |', '| 2 |', '', '# C', '', '| z |', '| --- |', '| 3 |', '', '结语段落'].join('\n');
  const buf = await word.writeDocx({ markdown: md }, { title: 't' });
  // 把 B、C 两张表的行删光,留 0 行残留
  const emptied = await editDocx(buf, [
    { op: 'delete_table_row', table: 2, at: 1, count: 2 },
    { op: 'delete_table_row', table: 3, at: 1, count: 2 },
  ]);
  const rowsOf = async (b) => scanTables(await partText(b, 'document.xml')).map((t) => t.rows).join(',');
  if (await rowsOf(emptied.buf) !== '2,0,0') throw new Error('测试前提不成立: ' + await rowsOf(emptied.buf));
  const cleaned = await editDocx(emptied.buf, [{ op: 'delete_table', scope: 'all', onlyEmpty: true }]);
  if (!cleaned.changes[0].includes('删除表格 2、3')) throw new Error('结果描述不对: ' + cleaned.changes[0]);
  if (await rowsOf(cleaned.buf) !== '2') throw new Error('清理后应只剩 1 张表: ' + await rowsOf(cleaned.buf));
  const mdOut = htmlToMarkdown((await word.readDocx(cleaned.buf, {})).html);
  for (const keep of ['**x**', '| 1 |', '正文段落', '结语段落', '# A']) {
    if (!mdOut.includes(keep)) throw new Error('清理误伤了 ' + keep + ': ' + mdOut);
  }
  // 已经没有空表时给出清晰报错,而不是静默成功
  let again = '';
  try { await editDocx(cleaned.buf, [{ op: 'delete_table', onlyEmpty: true }]); } catch (e) { again = e.message; }
  if (!again.includes('没有 0 行的空表格')) throw new Error('再次清理的报错不对: ' + again);
  return cleaned.changes[0];
});

await t('表格: delete_table 的 scope:all 与正文结尾保护', async () => {
  const two = await word.writeDocx({ markdown: '# A\n\n| x |\n| --- |\n| 1 |\n\n中间段落\n\n# B\n\n| y |\n| --- |\n| 2 |\n\n结尾段落' }, { title: 't' });
  const all = await editDocx(two, [{ op: 'delete_table', scope: 'all' }]);
  if (!all.changes[0].includes('删除表格 1、2')) throw new Error('scope:all 描述不对: ' + all.changes[0]);
  if (scanTables(await partText(all.buf, 'document.xml')).length !== 0) throw new Error('还有表格残留');
  const text = htmlToMarkdown((await word.readDocx(all.buf, {})).html);
  if (!text.includes('中间段落') || !text.includes('结尾段落')) throw new Error('正文被误删: ' + text);

  // 末尾是表格:删完不能让正文以表格结尾;只剩表格时还得补一个空段落
  const tail = await word.writeDocx({ markdown: '正文段落。\n\n| x |\n| --- |\n| 1 |' }, { title: 't' });
  const cut = await editDocx(tail, [{ op: 'delete_table', table: 1 }]);
  const only = await editDocx(await word.writeDocx({ markdown: '| x |\n| --- |\n| 1 |' }, { title: 't' }), [{ op: 'delete_table', table: 1 }]);
  for (const [label, b] of [['末尾表格', cut.buf], ['只剩表格', only.buf]]) {
    const doc = await partText(b, 'document.xml');
    const body = doc.slice(doc.indexOf('<w:body'), doc.lastIndexOf('<w:sectPr'));
    if (body.trimEnd().endsWith('</w:tbl>')) throw new Error(`${label}: 正文仍以表格结尾`);
    if (!/<w:p[\s/>]/.test(body)) throw new Error(`${label}: 正文里没有段落了`);
    if (!(await word.readDocx(b, {})).meta) throw new Error(`${label}: 文档读不回来`);
  }
  return 'scope:all 删两张；正文结尾两种情况都兜住';
});

await t('表格: delete_table 的错误分支', async () => {
  const withTable = await word.writeDocx({ markdown: '| A |\n| --- |\n| 1 |' }, { title: 't' });
  const noTable = await word.writeDocx({ markdown: '只有正文' }, { title: 't' });
  const fail = async (buf, ops) => { try { await editDocx(buf, ops); return { code: '', message: '' }; } catch (e) { return { code: e.code, message: e.message }; } };
  const cases = [
    { buf: noTable, ops: [{ op: 'delete_table' }], want: 'TABLE_NOT_FOUND', label: '文档里没有表格' },
    { buf: withTable, ops: [{ op: 'delete_table', table: 9 }], want: 'TABLE_NOT_FOUND', label: '序号越界' },
    { buf: withTable, ops: [{ op: 'delete_table', scope: 'each' }], want: 'INVALID_ARGS', label: 'scope 非法' },
  ];
  for (const c of cases) {
    const got = await fail(c.buf, c.ops);
    if (got.code !== c.want) throw new Error(`${c.label} 期望 ${c.want}，实际 ${got.code || '没报错'}`);
  }
  return `${cases.length} 种错误分支`;
});

await t('表格: 错误分支(选表 / 序号 / 区间)', async () => {
  const withTable = await word.writeDocx({ markdown: '| A |\n| --- |\n| 1 |' }, { title: 't' });
  const withoutTable = await word.writeDocx({ markdown: '只有正文' }, { title: 't' });
  const twoTables = await word.writeDocx({ markdown: '| A |\n| --- |\n| 1 |\n\nx\n\n| B |\n| --- |\n| 2 |' }, { title: 't' });
  const fail = async (buf, ops) => { try { await editDocx(buf, ops); return { code: '', message: '' }; } catch (e) { return { code: e.code, message: e.message }; } };
  const cases = [
    { buf: withoutTable, ops: [{ op: 'set_table', borders: 'all' }], want: 'TABLE_NOT_FOUND', label: '没有表格' },
    { buf: twoTables, ops: [{ op: 'set_table', borders: 'all' }], want: 'INVALID_ARGS', label: '多个表格却没给序号' },
    { buf: withTable, ops: [{ op: 'set_table', table: 9, borders: 'all' }], want: 'TABLE_NOT_FOUND', label: '序号越界' },
    { buf: withTable, ops: [{ op: 'set_table', table: 1 }], want: 'INVALID_ARGS', label: 'set_table 没给属性' },
    { buf: withTable, ops: [{ op: 'delete_table_row', at: 9 }], want: 'INVALID_ARGS', label: '删行越界' },
    { buf: withTable, ops: [{ op: 'delete_table_column', at: 9 }], want: 'INVALID_ARGS', label: '删列越界' },
    { buf: withTable, ops: [{ op: 'merge_table_cells', range: 'A1:B9' }], want: 'INVALID_ARGS', label: '合并越界' },
    { buf: withTable, ops: [{ op: 'merge_table_cells', range: 'A1' }], want: 'INVALID_ARGS', label: 'range 格式错' },
    { buf: twoTables, ops: [{ op: 'set_table', scope: 'each', borders: 'all' }], want: 'INVALID_ARGS', label: 'scope 非法' },
    { buf: withTable, ops: [{ op: 'set_style', scope: 'headings', sizePt: 12 }], want: 'PARAGRAPH_NOT_FOUND', label: '角色没匹配到段落' },
    { buf: withTable, ops: [{ op: 'set_style', scope: 'body', table: { borders: 'all' } }], want: 'INVALID_ARGS', label: 'set_style 里塞 table' },
  ];
  for (const c of cases) {
    const got = await fail(c.buf, c.ops);
    if (got.code !== c.want) throw new Error(`${c.label} 期望 ${c.want}，实际 ${got.code || '没报错'}: ${got.message}`);
  }
  return `${cases.length} 种错误分支`;
});

await t('表格: 按内容分配列宽时超长单元不吞掉其他列', async () => {
  const md = `| 短 | 很长的列内容需要占据更多宽度但也不能把别的列吃掉 | 中 |\n| --- | --- | --- |\n| 1 | ${'长'.repeat(80)} | 2 |`;
  const buf = await word.writeDocx({ markdown: md }, { title: 't', style: { table: { columnWidthMode: 'auto' } } });
  const doc = await partText(buf, 'document.xml');
  const widths = (doc.match(/<w:gridCol w:w="(\d+)"\/>/g) || []).map((s) => Number(s.replaceAll(/\D/g, '')));
  const sum = widths.reduce((a, b) => a + b, 0);
  const shares = widths.map((w) => Math.round((w / sum) * 100));
  if (shares[1] > 92) throw new Error('超长列占得太多: ' + shares.join());
  if (shares[0] < 3 || shares[2] < 3) throw new Error('其他列被挤没了: ' + shares.join());
  return shares.join('/');
});

await t('排版: 超长畸形 line-height 必须线性失败(不变 O(n²))', () => {
  // `\d*\.?\d+` 这类重叠量词在 6 万位数字串上要 2 秒;line-height 来自不可信的 HTML,必须挡住
  const bad = '1'.repeat(60000) + 'X';
  const t0 = Date.now();
  const sa = paragraphStyleFromCss(declarationsFrom(`line-height:${bad}`), 32);
  const sb = paragraphStyleFromCss(declarationsFrom(`line-height:${bad}pt`), 32);
  const ms = Date.now() - t0;
  if (sa.line !== undefined || sb.line !== undefined) throw new Error('畸形行高不该被解析出数值');
  if (ms > 300) throw new Error(`解析 ${bad.length} 字符耗时 ${ms}ms，疑似回退到超线性匹配`);
  const ok = paragraphStyleFromCss(declarationsFrom('line-height:1.5'), 32);
  if (ok.line !== 360 || ok.lineRule !== 'auto') throw new Error('正常倍数行高解析错了: ' + JSON.stringify(ok));
  return `6 万字符 ${ms}ms 内失败`;
});

await t('排版: 内联 CSS 解析(font-family / line-height / text-indent / margin)', async () => {
  const decls = declarationsFrom('font-family: 仿宋_GB2312, Times New Roman; font-size: 16pt; line-height: 28.8pt; text-indent: 2em; text-align: justify; margin: 12pt 0');
  const run = runStyleFromCss(decls);
  if (run.font !== '仿宋_GB2312' || run.fontAscii !== 'Times New Roman') throw new Error('中/西文字体没分开: ' + JSON.stringify(run));
  if (run.sizeHalfPt !== 32) throw new Error('字号没解析: ' + run.sizeHalfPt);
  const para = paragraphStyleFromCss(decls, 32);
  if (para.line !== 576 || para.lineRule !== 'exact') throw new Error('行距没解析: ' + JSON.stringify(para));
  if (para.firstLineTwips !== 640) throw new Error('缩进没解析: ' + para.firstLineTwips);
  if (para.align !== 'both') throw new Error('对齐没解析: ' + para.align);
  if (para.beforeTwips !== 240 || para.afterTwips !== 240) throw new Error('段前段后没解析: ' + JSON.stringify(para));
  return '中文字体/西文字体/固定行距/缩进/段前后都解析出来了';
});

await t('排版: 文档级 style 写进 docDefaults(整篇默认)', async () => {
  const buf = await word.writeDocx({ html: '<p>正文</p>' }, {
    title: 't',
    style: { font: '仿宋_GB2312', sizePt: 16, lineSpacingPt: 28.8, firstLineIndentChars: 2, align: 'both', headings: { font: '黑体', sizePt: 16 } },
  });
  const styles = await partText(buf, 'styles.xml');
  const defaults = styles.slice(styles.indexOf('<w:docDefaults'), styles.indexOf('</w:docDefaults>') + 16);
  const markers = [
    ['中文字体进 docDefaults', defaults.includes('仿宋_GB2312') && defaults.includes('w:eastAsia="仿宋_GB2312"')],
    ['三号(32 半磅)', defaults.includes('w:sz w:val="32"')],
    ['行距固定值 576', defaults.includes('w:line="576"') && defaults.includes('w:lineRule="exact"')],
    ['首行缩进 640', defaults.includes('w:firstLine="640"')],
    ['两端对齐', defaults.includes('w:jc w:val="both"')],
    ['标题样式用黑体', /w:styleId="Heading1"[\s\S]*?黑体/.test(styles)],
  ];
  const bad = markers.filter(([, ok]) => !ok).map(([n]) => n);
  if (bad.length) throw new Error('缺失: ' + bad.join(', '));
  return `${markers.length} 项默认排版都落到 styles.xml`;
});

await t('排版: 内联样式落到段落与 run,并能被容器继承', async () => {
  const html = '<div style="font-family: 楷体_GB2312, Times New Roman; font-size: 16pt; line-height: 28.8pt; text-indent: 2em">'
    + '<p>div 里继承下来的段落</p></div>'
    + '<p style="font-family:黑体;font-size:16pt;text-align:center;margin:12pt 0">单独指定的一段</p>';
  const buf = await word.writeDocx({ html }, { title: 't' });
  const doc = await partText(buf, 'document.xml');
  const markers = [
    ['div 的字体传给子段落', doc.includes('w:eastAsia="楷体_GB2312"')],
    ['中西文字体分开', doc.includes('w:ascii="Times New Roman"') && doc.includes('w:eastAsia="楷体_GB2312"')],
    ['div 的行距传给子段落', doc.includes('w:line="576"') && doc.includes('w:lineRule="exact"')],
    ['div 的首行缩进传给子段落', doc.includes('w:firstLine="640"')],
    ['段落自身 style 生效(黑体+居中)', doc.includes('w:eastAsia="黑体"') && doc.includes('w:jc w:val="center"')],
    ['段前段后 12 磅', doc.includes('w:before="240"') && doc.includes('w:after="240"')],
  ];
  const bad = markers.filter(([, ok]) => !ok).map(([n]) => n);
  if (bad.length) throw new Error('缺失: ' + bad.join(', '));
  return `${markers.length} 项内联样式都生效`;
});

await t('排版: 改已有文档 set_style(全篇 + 单段 + 区间)', async () => {
  const buf = await word.writeDocx({ markdown: '# 合同标题\n\n第一条 金额 100 元。\n\n第二条 期限 30 天。\n\n第三条 附则。' }, { title: 't' });
  const { buf: out, changes } = await editDocx(buf, [
    { op: 'set_style', scope: 'all', font: '仿宋_GB2312', sizePt: 16, lineSpacingPt: 28.8, firstLineIndentChars: 2, align: 'both' },
    { op: 'set_style', match: '合同标题', font: '黑体', sizePt: 22, lineSpacingPt: 33, align: 'center', firstLineIndentChars: 0 },
    { op: 'set_style', from: 3, to: 3, spacingBeforePt: 6 },
  ]);
  if (!changes[0].includes('7 段') && !/设置 \d+ 段样式/.test(changes[0])) throw new Error('全篇改样式的结果描述不对: ' + changes[0]);
  const doc = await partText(out, 'document.xml');
  const md = htmlToMarkdown((await word.readDocx(out, {})).html);
  const markers = [
    ['正文仿宋三号', doc.includes('w:eastAsia="仿宋_GB2312"') && doc.includes('w:sz w:val="32"')],
    ['正文固定行距+首行缩进', doc.includes('w:line="576"') && doc.includes('w:firstLine="640"')],
    ['标题黑体二号+居中', doc.includes('w:eastAsia="黑体"') && doc.includes('w:sz w:val="44"') && doc.includes('w:jc w:val="center"')],
    ['标题不缩进', doc.includes('w:firstLine="0"')],
    ['标题的 Heading1 样式没被抹掉', doc.includes('w:pStyle w:val="Heading1"')],
    ['区间段前距 6 磅', doc.includes('w:before="120"')],
    ['标签配对完好', (doc.match(/<w:r>/g) || []).length === (doc.match(/<\/w:r>/g) || []).length],
    ['正文内容没被改', md.includes('第一条 金额 100 元。') && md.includes('第三条 附则。')],
  ];
  const bad = markers.filter(([, ok]) => !ok).map(([n]) => n);
  if (bad.length) throw new Error('缺失: ' + bad.join(', '));
  return changes.join(' / ');
});

await t('排版: set_style 不覆盖没提到的格式(加粗/字号保留)', async () => {
  const buf = await word.writeDocx({ markdown: '普通一段\n\n**加粗**与普通混排' }, { title: 't' });
  const { buf: out } = await editDocx(buf, [{ op: 'set_style', scope: 'all', font: '仿宋_GB2312', firstLineIndentChars: 2 }]);
  const doc = await partText(out, 'document.xml');
  if (!doc.includes('<w:b/>')) throw new Error('原有的加粗被抹掉了');
  if (doc.includes('w:sz w:val=')) throw new Error('没提到字号却写入了 sz');
  if (doc.includes('w:line=')) throw new Error('没提到行距却写入了 line');
  return '只写入了 font 与 firstLine,加粗保留';
});

await t('排版: set_style 的错误分支', async () => {
  const buf = await word.writeDocx({ markdown: '第一段\n\n第二段' }, { title: 't' });
  const codeOf = async (ops) => { try { await editDocx(buf, ops); return ''; } catch (e) { return e.code; } };
  const cases = [
    { label: '没给样式属性', ops: [{ op: 'set_style' }], want: 'INVALID_ARGS' },
    { label: 'scope 非法', ops: [{ op: 'set_style', scope: 'each', sizePt: 12 }], want: 'INVALID_ARGS' },
    { label: '字号非法', ops: [{ op: 'set_style', scope: 'all', sizePt: 0 }], want: 'INVALID_ARGS' },
    { label: 'from > to', ops: [{ op: 'set_style', from: 2, to: 1, sizePt: 12 }], want: 'INVALID_ARGS' },
    { label: '区间越界', ops: [{ op: 'set_style', from: 9, to: 12, sizePt: 12 }], want: 'PARAGRAPH_NOT_FOUND' },
    { label: 'match 找不到', ops: [{ op: 'set_style', match: '不存在', sizePt: 12 }], want: 'PARAGRAPH_NOT_FOUND' },
    { label: '对齐值非法', ops: [{ op: 'set_style', align: 'middle', scope: 'all' }], want: 'INVALID_ARGS' },
    { label: '两个行距参数', ops: [{ op: 'set_style', scope: 'all', lineSpacingPt: 28, lineSpacingMultiple: 1.5 }], want: 'INVALID_ARGS' },
  ];
  for (const { label, ops, want } of cases) {
    const got = await codeOf(ops);
    if (got !== want) throw new Error(label + '(' + JSON.stringify(ops) + ') 期望 ' + want + '，实际 ' + (got || '没报错'));
  }
  return `${cases.length} 种错误分支`;
});

// ---------------------------------------------------------------------------
// Word 局部修改:office_edit_docx
// ---------------------------------------------------------------------------
const CONTRACT_MD = [
  '# 采购合同', '',
  '甲方：某某科技有限公司', '',
  '**金额**为 100 元。', '',
  '| 项目 | 数量 |', '| --- | --- |', '| 键盘 | 10 |', '',
  '乙方：某某贸易有限公司', '',
  '附则：本合同一式两份。',
].join('\n');

/** 造一份带标题/加粗/表格的合同文档。 */
async function writeContract(name = 'contract.docx') {
  const { writeFile } = await import('node:fs/promises');
  const path = join(outDir, name);
  await writeFile(path, await word.writeDocx({ markdown: CONTRACT_MD }, { title: '合同' }));
  return path;
}

/** zip 里每个部件的文本快照(用来证明只动了该动的部件)。 */
async function partSnapshot(buf) {
  const { zip } = await openOoxml(buf, 'docx');
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  return new Map(names.map((n) => [n, zip.files[n].asText()]));
}

await t('word 编辑: 只重写正文部件,其余部件逐字节不变', async () => {
  const { readFile } = await import('node:fs/promises');
  const path = await writeContract();
  const buf = await readFile(path);
  const before = await partSnapshot(buf);
  const r = await editDocx(buf, [
    { op: 'replace_text', find: '金额为 100', replace: '金额为 200' },
    { op: 'insert_paragraph', text: '签署日期：2026-01-01', heading: 2, position: 'end' },
    { op: 'delete_paragraph', paragraph: 2 },
  ]);
  const after = await partSnapshot(r.buf);
  const changed = [...after.keys()].filter((n) => before.get(n) !== after.get(n));
  if (changed.join() !== 'word/document.xml') throw new Error('被改动的部件不止正文: ' + changed.join(', '));
  const untouched = [...before.keys()].filter((n) => !changed.includes(n));
  if (!untouched.every((n) => before.get(n) === after.get(n))) throw new Error('有其它部件被改写');
  const back = await word.readDocx(r.buf, {});
  const md = htmlToMarkdown(back.html);
  if (!md.includes('金额为 200')) throw new Error('替换没生效: ' + md);
  if (md.includes('100 元')) throw new Error('旧值还在: ' + md);
  if (md.includes('甲方：某某科技')) throw new Error('删除段落没生效');
  if (!md.includes('乙方：某某贸易')) throw new Error('不该动的段落被改了: ' + md);
  if (!md.includes('键盘')) throw new Error('表格被破坏');
  if (!md.includes('## 签署日期')) throw new Error('插入的标题段落没生效');
  return `${before.size} 个部件中仅正文变化`;
});

await t('word 编辑: 跨 run 查找替换(Word 常把一句话拆进多个 run)', async () => {
  const { readFile } = await import('node:fs/promises');
  const path = await writeContract('cross-run.docx');
  const buf = await readFile(path);
  // `**金额**为 100 元。` 被拆成加粗 run「金额」+ 普通 run「为 100 元。」
  const paras = scanParagraphs((await openOoxml(buf, 'docx')).zip.file('word/document.xml').asText());
  const target = paras.find((p) => p.text.includes('金额为 100'));
  if (!target || target.segs.length < 2) throw new Error('测试前提不成立: 这句话没有被拆进多个 run');
  const { buf: out, changes } = await editDocx(buf, [{ op: 'replace_text', find: '金额为 100 元', replace: '金额为 200 元' }]);
  if (!changes[0].includes('1 处')) throw new Error('替换计数不对: ' + changes.join());
  const md = htmlToMarkdown((await word.readDocx(out, {})).html);
  if (!md.includes('200 元') || md.includes('100 元')) throw new Error('跨 run 替换结果不对: ' + md);
  return `${target.segs.length} 个 run 之间完成替换`;
});

await t('word 编辑: 整段改写保留段落样式与字符格式', async () => {
  const { readFile } = await import('node:fs/promises');
  const path = await writeContract('set-para.docx');
  const buf = await readFile(path);
  const { buf: out } = await editDocx(buf, [
    { op: 'set_paragraph', match: '乙方：某某贸易有限公司', text: '乙方：某某物流有限公司' },
    { op: 'set_paragraph', paragraph: 1, text: '采购合同（修订版）' },
  ]);
  const read = await word.readDocx(out, {});
  const md = htmlToMarkdown(read.html);
  if (!md.startsWith('# 采购合同（修订版）')) throw new Error('标题样式/内容不对: ' + md.slice(0, 40));
  if (!md.includes('乙方：某某物流有限公司')) throw new Error('整段改写没生效');
  const outline = await office.opRead(path, { outline: true });
  if (outline.meta.headings < 1) throw new Error('测试文档本应有标题');
  const edited = await office.opRead(path, {});
  if (!edited.content.includes('采购合同')) throw new Error('原文件不该被这里改动(editDocx 只返回字节)');
  return '标题仍是标题，正文已替换';
});

await t('word 编辑: 插入段落(含标题)与删除段落', async () => {
  const { readFile } = await import('node:fs/promises');
  const path = await writeContract('ins-del.docx');
  const buf = await readFile(path);
  const { buf: out, paragraphsBefore, paragraphsAfter } = await editDocx(buf, [
    { op: 'insert_paragraph', text: '第一章 总则', heading: 2, position: 'start' },
    { op: 'insert_paragraph', text: '插入在附则之后', position: 'after', match: '附则：本合同一式两份。' },
    { op: 'delete_paragraph', match: '甲方：某某科技有限公司' },
  ]);
  // 两个 insert + 一个 delete → 净增 1
  if (paragraphsAfter !== paragraphsBefore + 1) throw new Error(`段落数应为 ${paragraphsBefore} → ${paragraphsBefore + 1}(插 2 删 1)：实际 ${paragraphsAfter}`);
  const md = htmlToMarkdown((await word.readDocx(out, {})).html);
  if (!md.includes('第一章 总则')) throw new Error('文首插入失败: ' + md.slice(0, 60));
  if (!md.includes('插入在附则之后')) throw new Error('按 match 插入失败');
  if (md.includes('甲方：某某科技')) throw new Error('删除段落失败');
  if (md.indexOf('第一章 总则') > md.indexOf('采购合同')) throw new Error('文首插入位置不对');
  if (md.indexOf('插入在附则之后') < md.indexOf('附则：本合同一式两份。')) throw new Error('after 位置不对');
  const { zip } = await openOoxml(out, 'docx');
  const xml = zip.file('word/document.xml').asText();
  if (!xml.trimEnd().endsWith('</w:document>')) throw new Error('正文 XML 结构被破坏');
  if (xml.indexOf('<w:sectPr') < xml.lastIndexOf('<w:p>')) throw new Error('页面设置节点没被留在最后');
  return `${paragraphsBefore} 段 → 插入 2 删 1 → ${paragraphsAfter} 段`;
});

await t('word 编辑: 文本中的 & < > 与首尾空格正确转义', async () => {
  const { readFile } = await import('node:fs/promises');
  const path = await writeContract('escape.docx');
  const buf = await readFile(path);
  const tricky = '  A & B < C > D  ';
  const { buf: out } = await editDocx(buf, [
    { op: 'set_paragraph', match: '附则：本合同一式两份。', text: tricky },
    { op: 'replace_text', find: '100', replace: '200 & 300' },
  ]);
  const { zip } = await openOoxml(out, 'docx');
  const xml = zip.file('word/document.xml').asText();
  if (xml.includes('< C >')) throw new Error('尖括号没有转义,文档结构会被破坏');
  if (!xml.includes('xml:space="preserve"')) throw new Error('首尾空格没加 xml:space');
  if ((xml.match(/xml:space="preserve"/g) || []).length > 40) throw new Error('xml:space 被重复写入');
  const md = htmlToMarkdown((await word.readDocx(out, {})).html);
  if (!md.includes('A & B < C > D')) throw new Error('实体解码回来不对: ' + md);
  if (!md.includes('200 & 300')) throw new Error('替换文本里的 & 不对: ' + md);
  return 'A & B < C > D 与首尾空格都对';
});

await t('word 编辑: 表格里的段落也能定位与改写', async () => {
  const { readFile } = await import('node:fs/promises');
  const path = await writeContract('table-edit.docx');
  const buf = await readFile(path);
  const { zip } = await openOoxml(buf, 'docx');
  const paras = scanParagraphs(zip.file('word/document.xml').asText());
  const cell = paras.find((p) => p.text === '键盘');
  if (!cell) throw new Error('没找到表格里的段落');
  const { buf: out } = await editDocx(buf, [{ op: 'set_paragraph', match: '键盘', text: '机械键盘' }]);
  const md = htmlToMarkdown((await word.readDocx(out, {})).html);
  if (!md.includes('机械键盘')) throw new Error('表格单元格没改写成功: ' + md);
  if (!md.includes('项目') || !md.includes('数量')) throw new Error('表头被破坏: ' + md);
  return `表格内第 ${cell.index} 段改写成功`;
});

await t('word 编辑: limit 限制替换处数', async () => {
  const { writeFile, readFile } = await import('node:fs/promises');
  const path = join(outDir, 'limit.docx');
  await writeFile(path, await word.writeDocx({ markdown: 'AAA 待办\n\nBBB 待办\n\nCCC 待办' }, { title: 'x' }));
  const buf = await readFile(path);
  const { buf: out, changes } = await editDocx(buf, [{ op: 'replace_text', find: '待办', replace: '完成', limit: 2 }]);
  if (!changes[0].includes('2 处')) throw new Error('limit 没生效: ' + changes.join());
  const md = htmlToMarkdown((await word.readDocx(out, {})).html);
  if ((md.match(/完成/g) || []).length !== 2 || (md.match(/待办/g) || []).length !== 1) {
    throw new Error('替换处数不对: ' + md);
  }
  const all = await editDocx(buf, [{ op: 'replace_text', find: '待办', replace: '完成' }]);
  if (!all.changes[0].includes('3 处')) throw new Error('不限次数时应替换 3 处: ' + all.changes.join());
  return 'limit=2 只改前两处，不限时改 3 处';
});

await t('word 编辑: 定位失败 / 参数非法都有清晰报错', async () => {
  const { readFile, writeFile } = await import('node:fs/promises');
  const path = await writeContract('errors.docx');
  const buf = await readFile(path);
  const fail = async (ops) => { try { await editDocx(buf, ops); return { code: '', message: '' }; } catch (e) { return { code: e.code, message: e.message }; } };

  const CASES = [
    { ops: [], want: 'INVALID_ARGS' },
    { ops: [{ op: 'unknown_op' }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'delete_paragraph' }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'delete_paragraph', paragraph: 999 }], want: 'PARAGRAPH_NOT_FOUND' },
    { ops: [{ op: 'delete_paragraph', paragraph: 0 }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'delete_paragraph', paragraph: 1.5 }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'delete_paragraph', match: '不存在的段落' }], want: 'PARAGRAPH_NOT_FOUND' },
    { ops: [{ op: 'set_paragraph', paragraph: 1, text: '' }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'insert_paragraph', text: 'x', heading: 12 }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'insert_paragraph', text: 'x', position: 'middle' }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'insert_paragraph', text: 'x', position: 'before' }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'replace_text', find: '' }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'replace_text', find: 'x', limit: 0 }], want: 'INVALID_ARGS' },
    { ops: [{ op: 'replace_text', find: 'x', limit: 'abc' }], want: 'INVALID_ARGS' },
    { ops: Array.from({ length: 201 }, () => ({ op: 'delete_paragraph', paragraph: 1 })), want: 'INVALID_ARGS' },
  ];
  for (const { ops, want } of CASES) {
    const got = await fail(ops);
    if (got.code !== want) {
      throw new Error(JSON.stringify(ops).slice(0, 58) + ' 期望 ' + want + '，实际 ' + (got.code || '没报错'));
    }
  }
  const unknown = await fail([{ op: 'unknown_op' }]);
  if (!unknown.message.includes('replace_text')) throw new Error('未知操作的报错没列出可用操作: ' + unknown.message);

  // 唯一性:同一段文字出现两次时必须改用序号
  const dup = join(outDir, 'dup.docx');
  await writeFile(dup, await word.writeDocx({ markdown: '重复段\n\n重复段' }, { title: 'x' }));
  const dupFail = await (async () => { try { await editDocx(await readFile(dup), [{ op: 'delete_paragraph', match: '重复段' }]); return ''; } catch (e) { return e.code; } })();
  if (dupFail !== 'AMBIGUOUS_MATCH') throw new Error('重复段没有报歧义: ' + dupFail);

  // 扩展名与文件不存在
  const badExt = await office.opEditDocx(join(outDir, 'note.txt'), [], null).then(() => '', (e) => e.code);
  if (badExt !== 'UNSUPPORTED_FORMAT') throw new Error('.txt 没被拦: ' + badExt);
  const docExt = await office.opEditDocx(join(outDir, 'whatever.doc'), [{ op: 'delete_paragraph', paragraph: 1 }], null).then(() => '', (e) => e.code);
  if (docExt !== 'NEED_CONVERT') throw new Error('.doc 应提示先转 .docx: ' + docExt);
  const missing = await office.opEditDocx(join(outDir, 'nope.docx'), [{ op: 'delete_paragraph', paragraph: 1 }], null).then(() => '', (e) => e.code);
  if (missing !== 'NOT_FOUND') throw new Error('文件不存在时错误码不对: ' + missing);
  return `${CASES.length} 种非法输入 + 歧义匹配 + 3 种路径/格式错误`;
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

await t('安全: 远程图片一律拒绝(不联网,也不静默丢图)', async () => {
  let failure;
  try {
    await word.writeDocx({ html: '<p>正文</p><img src="https://example.com/x.jpg" alt="图 1">' });
  } catch (err) {
    failure = err;
  }
  if (!failure) throw new Error('远程图片应该直接报错,而不是悄悄丢掉');
  if (failure.code !== 'IMAGE_REMOTE') throw new Error('错误码不对: ' + failure.code);
  if (!failure.message.includes('不联网')) throw new Error('报错没说明不联网: ' + failure.message);
  // 无 src 时仍按 alt 文字处理,不会为了图片去联网或读盘
  const buf = await word.writeDocx({ html: '<p>正文</p><img alt="图 1">' });
  const back = await word.readDocx(buf);
  if (!back.content.includes('正文') || !back.content.includes('图 1')) throw new Error('alt 文本没保留: ' + back.content);
  const xml = new (await import('pizzip')).default(buf).file('word/document.xml').asText();
  if (xml.includes('pic:pic')) throw new Error('docx 里出现了图片部件');
  return `远程被拒(${failure.code}),无 src 退回 alt`;
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
