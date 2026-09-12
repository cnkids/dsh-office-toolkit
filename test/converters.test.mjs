// Unit tests for the cross-platform layer: pure-JS legacy readers, RTF/ODT
// conversion helpers, and the Windows-aware path guard.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import PizZip from 'pizzip';
import { isPathInside, isInsideAny } from '../lib/core/path-guard.js';
import { rtfToText, textToRtf, odtXmlToHtml, readOdt, readRtf, readDocWithWordExtractor } from '../lib/core/legacy-read.js';
import { converterStatus, describeStatus, hasWordConverter, wordRead } from '../lib/core/converters.js';
import { htmlToMarkdown } from '../lib/core/md.js';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, 'out-converters');
await mkdir(outDir, { recursive: true });

const results = [];
async function t(name, fn) {
  try {
    const detail = await fn();
    results.push({ name, ok: true });
    console.log(`✓ ${name}${detail ? ' — ' + detail : ''}`);
  } catch (err) {
    results.push({ name, ok: false, detail: String(err?.message || err) });
    console.log(`✗ ${name}\n   ${err?.message || err}`);
  }
}

function assert(cond, message) {
  if (!cond) throw new Error(message);
}

// ---------------------------------------------------------------------------
// RTF
// ---------------------------------------------------------------------------
await t(String.raw`RTF: 中文 \u 转义与段落/制表符`, async () => {
  const rtf = String.raw`{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset134 宋体;}}\f0\fs21 合同\u32534?\u21495?\par 第一行\line 第二行\tab 制表}`;
  const text = rtfToText(rtf);
  assert(text.includes('合同编号'), String.raw`中文 \u 转义未还原: ` + text);
  assert(text.includes('第一行\n第二行'), String.raw`\line 未换行: ` + JSON.stringify(text));
  assert(text.includes('\t制表'), String.raw`\tab 未还原`);
  assert(!text.includes('宋体'), 'fonttbl 未被跳过');
  return text.replaceAll('\n', '⏎');
});

await t('RTF: 跳过 pict/info/stylesheet 目的地', async () => {
  const rtf = String.raw`{\rtf1{\info{\title 标题元数据}}正文{\stylesheet{\s0 样式}}继续{\pict\pngblip 0011}结尾}`;
  const text = rtfToText(rtf);
  assert(text.includes('正文') && text.includes('继续') && text.includes('结尾'), '正文缺失: ' + text);
  assert(!text.includes('标题元数据') && !text.includes('样式') && !text.includes('0011'), '目的地内容未跳过: ' + text);
  return text;
});

await t('RTF: textToRtf → rtfToText 回环(含中文)', async () => {
  const source = '第一段中文内容\n\n第二段 with English';
  const back = rtfToText(textToRtf(source));
  assert(back.includes('第一段中文内容'), '第一段丢失: ' + back);
  assert(back.includes('第二段 with English'), '第二段丢失: ' + back);
  return '回环一致';
});

await t('RTF: readRtf 读取文件(文本/HTML 两种模式)', async () => {
  const path = join(outDir, 'sample.rtf');
  await writeFile(path, textToRtf('你好世界\n\n第二段'), 'utf8');
  const text = await readRtf(path);
  const html = await readRtf(path, { asHtml: true });
  assert(text.includes('你好世界'), '文本模式失败');
  assert(html.includes('<p>') && html.includes('你好世界'), 'HTML 模式失败');
  return '两种模式正常';
});

// ---------------------------------------------------------------------------
// ODT
// ---------------------------------------------------------------------------
const ODT_CONTENT = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">
<office:automatic-styles><style:style xmlns:style="x"/></office:automatic-styles>
<office:body><office:text>
<text:h text:outline-level="1">季度报告</text:h>
<text:h text:outline-level="2">二级标题</text:h>
<text:p>这是第一段<text:s text:c="3"/>带空格。</text:p>
<text:list><text:list-item><text:p>条目一</text:p></text:list-item><text:list-item><text:p>条目二</text:p></text:list-item></text:list>
<table:table><table:table-row><table:table-cell><text:p>产品</text:p></table:table-cell><table:table-cell><text:p>数量</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell><text:p>键盘</text:p></table:table-cell><table:table-cell><text:p>10</text:p></table:table-cell></table:table-row></table:table>
<text:p><text:a xlink:href="https://example.com">链接文字</text:a></text:p>
</office:text></office:body></office:document-content>`;

await t('ODT: content.xml → HTML(标题/段落/列表/表格/链接)', async () => {
  const html = odtXmlToHtml(ODT_CONTENT);
  assert(html.includes('<h1>季度报告</h1>'), '标题映射失败: ' + html);
  assert(html.includes('<h2>二级标题</h2>'), 'outline-level 映射失败: ' + html);
  assert(html.includes('这是第一段   带空格。'), 'text:c 空格数未还原: ' + html);
  assert(html.includes('<li><p>条目一</p></li>'), '列表映射失败: ' + html);
  assert(html.includes('<td><p>键盘</p></td>'), '表格映射失败: ' + html);
  assert(html.includes('<a href="https://example.com">链接文字</a>'), '链接映射失败');
  assert(!html.includes('style:style'), 'automatic-styles 未清理');
  return `${html.length} 字符`;
});

await t('ODT: readOdt 读取真实 zip 包', async () => {
  const zip = new PizZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text');
  zip.file('content.xml', ODT_CONTENT);
  zip.file('META-INF/manifest.xml', '<manifest/>');
  const path = join(outDir, 'sample.odt');
  await writeFile(path, zip.generate({ type: 'nodebuffer' }));
  const markdown = await readOdt(path);
  assert(markdown.includes('季度报告'), 'Markdown 缺少标题: ' + markdown);
  assert(markdown.includes('| 键盘 | 10 |'), 'Markdown 表格异常: ' + markdown);
  return markdown.split('\n')[0];
});

await t('ODT: 空/损坏文件报错清晰', async () => {
  const bad = join(outDir, 'broken.odt');
  await writeFile(bad, Buffer.from('not a zip'));
  let message = '';
  try {
    await readOdt(bad);
  } catch (err) {
    message = String(err?.message);
  }
  const clear = message.includes('不是有效的 .odt') || message.includes('缺少 content.xml');
  assert(clear, '报错不清晰: ' + message);
  return message.slice(0, 40);
});

// ---------------------------------------------------------------------------
// .doc via word-extractor
// ---------------------------------------------------------------------------
await t('.doc: word-extractor 纯 JS 读取', async () => {
  const fixture = join(here, 'out', 'conv.doc');
  try {
    await readFile(fixture);
  } catch {
    return '跳过（先运行 selftest.mjs 生成 test/out/conv.doc）';
  }
  const text = await readDocWithWordExtractor(fixture);
  assert(text.includes('项目周报'), '未读到预期内容: ' + text.slice(0, 80));
  return `${text.length} 字符`;
});

await t('.doc: converters.wordRead 自动选择后端', async () => {
  const fixture = join(here, 'out', 'conv.doc');
  try {
    await readFile(fixture);
  } catch {
    return '跳过（缺少 .doc 样本）';
  }
  const hit = await wordRead(fixture);
  assert(hit.content.includes('项目周报'), '读取失败');
  assert(typeof hit.backend === 'string' && hit.backend.length > 0, '未报告后端');
  return `后端: ${hit.backend}`;
});

// ---------------------------------------------------------------------------
// backend detection
// ---------------------------------------------------------------------------
await t('后端探测: 状态与摘要可用', async () => {
  const status = await converterStatus();
  assert(typeof status.textutil === 'boolean', 'textutil 字段缺失');
  assert('soffice' in status && 'wordCom' in status, '字段缺失');
  const summary = describeStatus(status);
  assert(summary.includes('纯 JS'), '摘要未包含纯 JS 回退: ' + summary);
  const canDoc = await hasWordConverter('doc');
  assert(typeof canDoc === 'boolean', 'hasWordConverter 返回值异常');
  return `platform=${status.platform}, ${summary}`;
});

// ---------------------------------------------------------------------------
// path guard (Windows semantics tested on any OS)
// ---------------------------------------------------------------------------
await t('路径围栏: Windows 大小写不敏感', async () => {
  const options = { platform: 'win32' };
  assert(isPathInside(String.raw`C:\Users\Me\work\a.docx`, String.raw`C:\Users\Me\work`, options), '同盘同目录应通过');
  assert(isPathInside(String.raw`c:\users\me\WORK\a.docx`, String.raw`C:\Users\Me\work`, options), '大小写不同应通过');
  assert(!isPathInside(String.raw`C:\Users\Me\workspace\a.docx`, String.raw`C:\Users\Me\work`, options), '同前缀的兄弟目录必须拒绝');
  assert(!isPathInside(String.raw`D:\data\a.docx`, String.raw`C:\Users\Me\work`, options), '不同盘必须拒绝');
  return 'win32 语义正确';
});

await t('路径围栏: POSIX 大小写敏感', async () => {
  const options = { platform: 'linux' };
  assert(isPathInside('/home/me/work/a.docx', '/home/me/work', options), '同目录应通过');
  assert(!isPathInside('/home/me/Work/a.docx', '/home/me/work', options), 'POSIX 下大小写不同必须拒绝');
  assert(!isPathInside('/home/me/work2/a.docx', '/home/me/work', options), '兄弟目录必须拒绝');
  return 'posix 语义正确';
});

await t('路径围栏: 多根目录判定', async () => {
  // 刻意避开 /tmp 这类公共可写目录(sonar S5443),断言语义与根目录取值无关
  const roots = ['/workspace', '/srv/office-data'];
  assert(isInsideAny('/srv/office-data/x.docx', roots), '第二个根目录应通过');
  assert(isInsideAny('/workspace/sub/y.xlsx', roots), 'workspace 子目录应通过');
  assert(!isInsideAny('/etc/passwd', roots), '/etc 必须拒绝');
  return '多根判定正确';
});

// ---------------------------------------------------------------------------
// regex backtracking safety (S5852 regression): malformed input must stay linear
// ---------------------------------------------------------------------------
await t('安全: 畸形超长输入不会触发超线性回溯', async () => {
  const n = 200000;
  const cases = [
    `<${'a'.repeat(n)}`,
    `<office:body${'x'.repeat(n)}</office:body>`,
    `<!--${'x'.repeat(n)}`,
    `<text:p ${'a="b" '.repeat(20000)}`,
    'A'.repeat(n),
  ];
  const started = Date.now();
  for (const evil of cases) {
    odtXmlToHtml(evil);
    htmlToMarkdown(evil);
  }
  const ms = Date.now() - started;
  if (ms > 3000) throw new Error(`${cases.length} 组畸形输入耗时 ${ms}ms，疑似回溯`);
  return `${cases.length} 组 × ODT/HTML 两路，共 ${ms}ms`;
});

await t('安全: 超长畸形图表引用快速失败', async () => {
  const { buildWorkbook } = await import('../lib/core/excel.js');
  const { injectChart } = await import('../lib/core/charts.js');
  const buf = await buildWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] });
  const started = Date.now();
  let threw = false;
  try {
    await injectChart(buf, { sheet: 'S', chartType: 'bar', categories: `A${'A'.repeat(200000)}`, series: ['B1:B2'] });
  } catch {
    threw = true;
  }
  const ms = Date.now() - started;
  if (!threw) throw new Error('畸形引用应当报错');
  if (ms > 2000) throw new Error(`畸形引用耗时 ${ms}ms，疑似回溯`);
  return `超长畸形引用 ${ms}ms 内失败`;
});

const failed = results.filter((r) => !r.ok);
console.log(`\n===== 跨平台层: ${results.length - failed.length}/${results.length} 通过 =====`);
if (failed.length) {
  console.log('失败项: ' + failed.map((f) => `${f.name}(${f.detail})`).join('; '));
  process.exit(1);
}
