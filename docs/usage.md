# 用法与参数

[← 返回 README](../README.md)

## 读取兼容性

真实环境里的 Word 文件常常「不标准」，插件在读取前会按**文件真实内容**判断，而不是只信扩展名：

| 情况 | 表现 |
| --- | --- |
| zip 条目名用反斜杠（部分国产工具、或在 Windows 上手工重打包过） | **自动把条目名与引用它的部件（`[Content_Types].xml`、`*.rels`）一起改成规范写法**后读取，并提示修正了几处；mammoth / ExcelJS 本来会直接报「不是有效的 docx」 |
| 部件名大小写不同（`Word/Document.xml`、`WORD\DOCUMENT.XML`） | 同上按规范名归一化后读取。Windows 文件系统不分大小写，那里的工具常这么写 |
| mammoth 认不出这个文件（找不到主文档部件或没有 body） | 先用 PizZip 重新打包再试；仍失败则用**内置解析器**从 `word/document.xml` 取正文，保证「能开 zip 就读得出字」，并在 `meta.reader` 标明用了哪种读法 |
| 后缀 `.doc`，内容其实是 `.docx` | 按真实内容当 docx 读，并提示「文件内容其实是 .docx」 |
| 后缀 `.docx`，内容其实是老的 OLE `.doc` | 按老格式读取，提示实际格式，不再报「docx 无法解压」 |
| 后缀 `.xlsx`，内容其实是 CSV / TSV 文本 | 按分隔符文本读取，提示实际格式 |
| 主文档部件不在规范路径（如由 `_rels/.rels` 声明为 `word/main.xml`） | 按**包关系**解析主部件后读取，规范名不存在也不影响 |
| zip 里没有 Office 主文档 | 报 `BAD_CONTAINER`，并列出**实际条目名与插件版本**，便于判断文件真身 |

仍会明确报错、不做猜测的两种情况：文件不存在（`NOT_FOUND`）、传入的是目录（`NOT_A_FILE`）。

## 读取

Word 正文默认以 Markdown 风格文本返回；Excel 以 TSV 代码块返回，可用 `sheets` / `range` / `maxRows` / `maxCols` 分段读取大表。

**只看内容**用 `office_read`；**要算**（求和 / 分组 / 筛选 / 去重 / 排序）用 [`office_query`](#计算office_query) —— 后者不会把整张表读进上下文。

**长 Word 文档分段读**（不会每次都从头开始）：

```json
{ "path": "标书.docx", "maxChars": 20000 }
{ "path": "标书.docx", "offset": 20000, "maxChars": 20000 }
{ "path": "标书.docx", "outline": true }
```

- `maxChars` 是**每段**的字符上限（默认 90000）。返回里会写明总字符数、本次区间，以及还有下一页时「继续读」该传的 `offset`。
- 分页边界尽量收在句末或换行处，不会把半句话、半张表格切开；**逐段拼回来与一次读完逐字节相同**，不重不漏。
- `offset` 超出总长会直接告知「已到文档末尾」，不是报错。
- `outline: true` 只返回标题大纲（TSV：级别 / 字符偏移 / 标题），可据此传 `offset` 直接跳到某一章。大纲按 Markdown 标题（`#` 开头）提取，围栏代码块里的 `#` 会跳过。
- `format: "html"` 下不支持 `offset` / `outline`（按字符切会切坏标签）；html 输出仍按 `maxChars` 截断，并提示想分段就切回 text 模式。
- `withFormatting` 的格式报告只在 `offset=0` 那一段给出，后续段只留一行提示，避免每页重复。

```json
{ "path": "报告.docx" }
{ "path": "报告.docx", "format": "html" }
{ "path": "销售.xlsx", "sheets": ["明细"], "range": "A1:F100" }
{ "path": "明细.csv", "maxRows": 50 }
```

## 计算（office_query）

在文件内做筛选 / 分组 / 聚合 / 排序，**只把结论返回**。全表扫描，不受 `office_read` 的行窗口限制；支持 `.xlsx .xls .xlsb .ods .csv .tsv`。

```json
{ "path": "订单.xlsx", "groupBy": ["地区"], "aggregate": [{"col": "金额", "fn": "sum", "as": "销售额"}], "orderBy": [{"col": "销售额", "dir": "desc"}] }
{ "path": "订单.xlsx", "where": [{"col": "地区", "op": "in", "value": ["华东", "华南"]}, {"col": "金额", "op": "gt", "value": 1000}], "aggregate": [{"col": "金额", "fn": "avg"}, {"col": "订单号", "fn": "countDistinct", "as": "订单数"}] }
{ "path": "订单.xlsx" }
```

| 参数 | 说明 |
| --- | --- |
| `sheet` | 工作表名或从 1 开始的序号，默认第 1 个 |
| `headerRow` | 表头行号，默认 1；表头不在第一行就传行号，没有表头传 `0`（列名取 `A`/`B`/`C`…） |
| `where` | 条件数组，多条件为 **AND**；算子 `eq ne gt gte lt lte contains startsWith endsWith in notIn isBlank notBlank` |
| `groupBy` | 分组列，如 `["地区","产品"]`；只给 `groupBy` 时输出每组行数 |
| `aggregate` | 汇总项 `[{col, fn, as}]`，`fn` 取 `sum avg min max count countDistinct`；不给 `groupBy` 时是一行全表汇总 |
| `orderBy` | 排序 `[{col, dir}]`，`col` 可用分组列或汇总结果列 |
| `limit` | 结果行数上限，默认 200，最大 2000；超出会写明「共 N 组，仅显示前 M 组」 |

**不给 `groupBy` / `aggregate` 时返回表结构画像**：每列的类型（数字 / 日期 / 文本 / 混合 / 空）、非空、空值、去重数、最小 / 最大 / 求和 / 均值、最高频的 3 个取值与次数 —— 一次调用就知道这张表长什么样，不用先读一遍再想怎么问。画像同样接受 `where`，可以只看子集。

数值识别做的是"尽力而为"：`1,234.00`、`¥88`、`12.5%`（→ `0.125`）、全角空格都能认；日期把 `2025/1/1`、`2025年1月1日`、带时间的写法统一成 `YYYY-MM-DD` 再比较。反过来，**认不出就报错或跳过，绝不猜**：

- `sum` / `avg` 跳过非数值单元格，并在结果下方写明跳过了几个（不会静默当成 0）。
- 某一组一个数值都没有时，`sum` / `avg` 返回空单元格而不是 `0`。
- 一边能解析成数字、另一边不能（如「金额」列里混着"待定"）时判为**不可比**，`>` / `=` 一律不成立 —— 否则字典序会让"待定" > 1000 静默成立。
- 列名写错、算子写错都会直接报错，并列出全部可用列名 / 合法算子。
- 文本排序按 Unicode 码位（结果确定，不依赖系统的排序规则）；中文不等于拼音序。

### 什么时候该退回脚本

本工具只做**声明式**操作。以下需求它覆盖不了，请照常写 Python 或别的办法：

- 跨文件 / 跨表的 join、关联、比对；
- 透视表、窗口函数、累计值、同比环比；
- 中位数、分位数、标准差、回归等统计量；
- 图表、可视化、导出报告。

一句话：**表格的筛选与汇总先问 `office_query`，它答不了再写脚本。**

## 新建

文档内容三选一：`html` / `markdown` / `text`。`office_write_docx` 支持标题、段落、粗体、斜体、下划线、删除线、等宽、上下标、列表（含嵌套）、表格、引用、代码块、分隔线、软换行，以及 A4 纵向 / 横向与自定义页边距。

```json
{ "path": "周报.docx", "markdown": "# 周报\n\n- 完成 A\n- 完成 B\n", "title": "周报" }
{ "path": "横向.docx", "markdown": "# 宽表", "landscape": true, "marginsMm": 15 }
```

工作簿支持多工作表、表头美化、列宽行高、公式与日期。

```json
{ "path": "台账.xlsx", "sheets": [
  { "name": "明细", "header": true, "columnWidths": [14, 10, 12],
    "rows": [["产品", "数量", "金额"], ["键盘", 10, "=B2*199"], ["鼠标", 20, "=B3*89"]] },
  { "name": "备注", "rows": [["说明"], ["公式在 Excel/WPS 打开时计算"]] } ] }
```

## 编辑已有工作簿

`ops` 按顺序执行，一个 `ops` 数组内可以混用任意操作。

```json
{ "path": "台账.xlsx", "ops": [
  { "op": "set_value", "sheet": "明细", "ref": "E1", "value": "合计", "style": { "bold": true } },
  { "op": "set_formula", "sheet": "明细", "ref": "E2", "formula": "SUM(C2:C3)" },
  { "op": "style_range", "sheet": "明细", "range": "A1:C1", "style": { "fill": "DDEBF7", "align": "center" } },
  { "op": "merge", "sheet": "明细", "range": "A5:C5" },
  { "op": "freeze", "sheet": "明细", "rows": 1, "cols": 1 },
  { "op": "auto_filter", "sheet": "明细", "range": "A1:C3" },
  { "op": "set_col_width", "sheet": "明细", "col": "C", "width": 18 },
  { "op": "add_chart", "sheet": "明细", "chartType": "column", "title": "金额对比",
    "categories": "A2:A3", "series": [{ "range": "C2:C3", "label": "金额" }], "anchor": "G2" } ] }
```

`op` 取值：`set_value`、`set_cells`、`set_formula`、`style_range`、`merge`、`unmerge`、
`insert_rows`、`delete_rows`、`insert_cols`、`delete_cols`、`add_sheet`、`rename_sheet`、
`delete_sheet`、`set_col_width`、`set_row_height`、`freeze`、`auto_filter`、`add_image`、
`add_chart`。单个操作失败时加 `"optional": true` 可跳过而不中断整批。

## 模板套打

模板里写 `{{变量}}`，支持 `{{a.b}}` 形式的嵌套取值与中文变量名；不支持循环 / 条件，批量套打多次调用即可。

```json
{ "templatePath": "合同模板.docx", "outputPath": "合同-001.docx",
  "data": { "合同编号": "HT-2026-001", "甲方": "某某公司", "金额": "12800" } }
```

## 转换

Word 家族（`doc` / `docx` / `rtf` / `odt` / `html` / `txt` / `md`）与表格家族（`xlsx` / `xls` / `xlsb` / `ods` / `csv` / `tsv` / `html`）各自内部互转，**不支持跨家族**。源与目标扩展名相同时：路径不同则复制、路径相同则直接返回。

```json
{ "sourcePath": "旧表.xls", "outputPath": "新表.xlsx" }
{ "sourcePath": "报告.docx", "outputPath": "报告.pdf" }
{ "sourcePath": "合同.docx", "outputPath": "合同.odt" }
```

## 参数

- **路径**：绝对路径，或相对会话工作区的路径；写入默认只允许会话工作区与系统临时目录，越界返回 `FS_SANDBOX_DENIED`。
- **单元格值**：数字 / 布尔原样；`=` 开头视为公式；`date:2026-09-09` 写入日期；`num:1,234.5` 强制数字。
- **样式**：`{bold, italic, fontSize, color, fill, align, valign, wrap, numFmt, border}`（颜色为 `RRGGBB`）。
- **截断**：大表默认 400 行 × 60 列（分段用 `range` / `maxRows`，要全表统计用 `office_query`）；Word 正文默认每段 90000 字符，用 `offset` 接着读。`office_query` 扫描上限 20 万行 × 200 列，超出会明确告知只统计了前 N 行。
- **结果规模**：`office_query` 的结果默认最多 200 行、6 位有效小数内取整；`limit` 上限 2000。
- **输入体积**：`.docx` 40 MB、`.xlsx` 60 MB；三者都是 zip，解压后总量超过 **1 GiB** 或压缩比超过 **150:1** 会按「疑似压缩炸弹」拒绝（`ZIP_BOMB_SUSPECTED`）。

## 限制

- 图表支持 `bar` / `column` / `line` / `pie`（以 OOXML 注入实现，数据由 Excel / WPS 打开时计算）；组合图、双轴等请手动调整。一个工作表只能有一个 drawing 部件，因此同表多图表共用一个 drawing（各自独立锚点，可分别拖动）。
- `.doc` 纯 JS 解析不保留表格与版式；写出 `.doc` / `.odt`、转 `.pdf` 需要 LibreOffice 或 Word。
- 图片只读不写：写入 `.docx` 时不嵌入图片，只把 `<img>` 的 `alt` 文本留在正文里。
- CSV / TSV / TXT 按 **UTF-8** 解码（这类格式不带编码信息），GBK 等其他编码请先转码。
- `javascript:` / `data:` / `vbscript:` / `file:` 链接在写入 `.docx` 时会降级为纯文本。
- 写入 `.xlsx` / `.docx` / `.odt` 时会校验输出必须是真正的 OOXML（`PK` 头），否则中止并报 `BAD_OUTPUT_FORMAT` —— 不会留下「扩展名是 Office、内容却是文本」的假文件。
- `.tsv` 按制表符写出，`.csv` 按逗号写出。

## 读取排版格式（行文规则比对）

对 `.docx` 传 `withFormatting: true`，会在正文之后附一份格式报告，并在 `meta.formatting` 里给出结构化数据：

- **每段**：字体（中文 / 西文）、字号、行距（`固定值 28.8pt` / `1.5 倍` / `单倍`）、首行缩进（`2 字符` / 磅值）、对齐、样式名、是否加粗
- **页面**：纸张尺寸与上下左右页边距（mm）
- **文档默认**：`docDefaults` 里的字体、字号、行距
- **格式分布**：字体 / 字号 / 行距 / 缩进 / 对齐各自的出现次数 —— 主流值即比对基准
- **偏离主流格式的段落**：逐条列出，便于快速定位不合规处

值和样式都做了继承解析，按 Word 的实际优先级合并：

1. `docDefaults`（文档默认）
2. **隐式默认样式** —— `w:default="1"` 的段落样式（通常是 `Normal` / `正文`），Word 会套到没写 `w:pStyle` 的段落上
3. `basedOn` 样式链
4. 段落直接格式（`w:ind` / `w:spacing` / `w:jc`）
5. run 级覆盖（一个段落取覆盖文字最多的 run 格式）

**主题字体也会解析**：Word 默认模板用 `w:eastAsiaTheme="minorEastAsia"` 这类引用而不是字体名，插件会读 `word/theme/theme1.xml`，并按 `<a:font script="Hans">` 取出中文实际字体（例如 `宋体` / `等线`），所以「样式给字体、段落自己改行距」这类公文常见写法能读对。

```json
{ "path": "通知.docx", "withFormatting": true }
```

报告节选：

```text
**页面**: 210×297 mm,页边距 上 37 / 下 35 / 左 28 / 右 26 mm
**文档默认**: 字体 Times New Roman / 等线,字号 10.5pt,行距 继承默认

| # | 段落文字 | 样式 | 对齐 | 行距 | 缩进 | 中文字体 | 字号 | 加粗 |
| 1 | 关于印发某某管理办法的通知 | 标题 | 居中 | 固定值 28pt | 无 | 方正小标宋简体 | 22pt | 是 |
| 2 | 各处室、各直属单位: | 正文 | 继承 | 固定值 28pt | 首行 2 字符 | 仿宋_GB2312 | 16pt | — |

**格式分布**(主流值即行文规则比对基准):
- 中文字体: 仿宋_GB2312 ×2、方正小标宋简体 ×1、黑体 ×1
- 行距: 固定值 28pt ×5
```

**限制**：格式提取目前只支持 `.docx`（`.doc` / `.rtf` / `.odt` 会返回一行提示）。表格内的段落也会被列出，并带 `inTable` 标记。

## 工具参数

以下是 7 个工具注册到 DSH 的完整参数表（由工具 schema 自动导出）。

### `office_read`

读取 Word / Excel / 表格文件内容（.docx .doc .rtf .odt .xlsx .xls .xlsb .ods .csv .tsv）。用户提到 Word、Excel、文档、表格、.docx、.xlsx 等后缀时一律用本工具。Word/Excel 是二进制格式,通用 read/write/edit 工具处理不了(会报 binary file),必须用本工具。解析方式自动选择：.docx 走 mammoth（Markdown 风格正文/表格）；.doc/.rtf/.odt 优先本机转换器(macOS textutil / LibreOffice / Word)，不可用时回退纯 JS；.xlsx 走 ExcelJS；.xls/.xlsb/.ods/.csv/.tsv 走 SheetJS（纯 JS，全平台可用）。Excel 以 TSV 代码块返回，可用 sheets/range/maxRows 分段读取大表。需要核对字体/字号/行距/缩进等排版格式时，对 .docx 传 withFormatting: true。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 |
| `sheets` | array<string> | — | 仅 Excel：只读取这些工作表（名称；序号请写成字符串，如 "2"） |
| `range` | string | — | 仅 Excel：读取范围，如 A1:F50 |
| `maxRows` | integer | — | 仅 Excel：最多读取行数，默认 400 |
| `maxCols` | integer | — | 仅 Excel：最多读取列数，默认 60 |
| `maxChars` | integer | — | Word 正文每段返回的字符上限（分页大小），默认 90000 |
| `offset` | integer | — | 仅 Word：从正文第几个字符开始返回，默认 0。返回里会给出总字符数与下次该传的 offset，用于分段读长文档；offset 超出总长会明确提示已到末尾 |
| `outline` | boolean | — | 仅 Word 的 text 模式：只返回标题大纲（TSV：级别 / 字符偏移 / 标题），可据此传 offset 直接跳到某一章；与 format: "html" 不能同时用 |
| `format` | enum: `text` / `html` | — | Word 输出格式：text=Markdown 风格（默认），html=原始 HTML |
| `withFormatting` | boolean | — | 仅 .docx：额外返回格式报告 —— 每段的字体(中文/西文)、字号、行距(固定值/倍数)、首行缩进、对齐、样式名，以及页面尺寸与页边距；并给出格式分布(主流值)与偏离主流的段落。用于比对行文规则(如"正文三号仿宋、行距固定值 28.8 磅")。结构化数据在 meta.formatting。 |

### `office_query`

在表格文件内直接算（.xlsx .xls .xlsb .ods .csv .tsv）：全表扫描，只把结论返回，几万行不必读进上下文。表格数据的求和/均值/计数/去重/分组/排序/条件筛选，优先用本工具而不是临时写脚本；返回的内容是结果表，不是原始数据。不给 groupBy/aggregate 时返回**表结构画像**。本工具覆盖范围之外的需求（跨文件 join、透视表、窗口函数、图表、统计建模）再另想办法。列名取表头行（headerRow，默认第 1 行）；条件算子：eq/ne/gt/gte/lt/lte/contains/startsWith/endsWith/in/notIn/isBlank/notBlank；聚合函数：sum/avg/min/max/count/countDistinct（数值列能识别 ¥1,234.00、12.5% 这类写法）。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 |
| `sheet` | — | — | 工作表名或从 1 开始的序号，默认第 1 个工作表 |
| `headerRow` | integer | — | 表头行号，默认 1；传 0 表示没有表头（列名自动取 A/B/C…） |
| `where` | array<object> | — | 筛选条件（多个条件为 AND），如 `[{"col":"地区","op":"eq","value":"华东"}]`；`col` 必填，`op` 默认 eq，`value` 为比较值（`in`/`notIn` 传数组，`isBlank`/`notBlank` 不需要） |
| `groupBy` | array<string> | — | 按这些列分组，如 `["地区","产品"]`；只给 `groupBy` 时输出每组的行数 |
| `aggregate` | array<object> | — | 汇总项 `[{col, fn, as}]`；`fn` 取 sum/avg/min/max/count/countDistinct，`as` 为结果列名（默认 `fn(列名)`） |
| `orderBy` | array<object> | — | 排序如 `[{"col":"销售额","dir":"desc"}]`，`dir` 取 asc/desc（默认 asc） |
| `limit` | integer | — | 结果行数上限，默认 200，最大 2000 |

### `office_write_docx`

新建 Word 文档（默认 .docx，也支持 .doc / .rtf / .odt）。要生成 Word 文档时用本工具，不要用通用 write 工具（写不出合法的 .docx）。内容三选一：html / markdown / text，支持标题、段落、粗体斜体、列表、表格、引用等。.doc/.rtf/.odt 旧格式输出需要本机转换器（macOS textutil / LibreOffice / Microsoft Word），不可用时会返回安装提示；图片等复杂元素建议用 .docx。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 扩展名决定输出格式：.docx（推荐）/.doc/.rtf/.odt。 |
| `html` | string | — | HTML 正文（与 markdown/text 三选一） |
| `markdown` | string | — | Markdown 正文（与 html/text 三选一） |
| `text` | string | — | 纯文本正文（与 html/markdown 三选一） |
| `title` | string | — | 文档标题（元数据） |
| `landscape` | boolean | — | 是否横向纸张（仅 .docx） |
| `marginsMm` | number | — | 页边距毫米（仅 .docx） |

### `office_write_xlsx`

新建 Excel 工作簿（.xlsx），可一次写入多个工作表；要生成表格文件时用本工具。rows 为二维数组；单元格值规则：数字/布尔按原样，字符串以 "=" 开头视为公式（如 "=SUM(B2:B9)"），"date:2026-09-09" 写入日期。header:true 会给首行加粗底纹；columnWidths 可设列宽。需要图表时先写数据，再用 office_edit_xlsx 的 add_chart。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 必须为 .xlsx。 |
| `sheets` | array<object> | ✅ | 工作表数组 |

### `office_edit_xlsx`

修改已有 .xlsx（按顺序执行 ops）。改表格必须用本工具，通用 edit 工具改不了二进制表格。支持操作：set_value{sheet,ref,value,style}、set_cells{sheet,start,values,style}、set_formula{sheet,ref,formula}、style_range{sheet,range,style}、merge/unmerge{sheet,range}、insert_rows/delete_rows{sheet,at,count}、insert_cols/delete_cols{sheet,at,count}、add_sheet{name,rows,header}、rename_sheet{sheet,name}、delete_sheet{sheet}、set_col_width{sheet,col,width}、set_row_height{sheet,row,height}、freeze{sheet,rows,cols}、auto_filter{sheet,range}、add_image{sheet,path,cell|range}、add_chart{sheet,chartType,categories,series,title,anchor}。style 支持 {bold,italic,fontSize,color,fill,align,valign,wrap,numFmt,border}（颜色为 RRGGBB）。add_chart 的 chartType 支持 bar/column/line/pie；categories 如 "A2:A6"，series 如 ["B2:B6"]（可带 label 单元格）。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 必须为 .xlsx（旧 .xls 先用 office_convert 转 .xlsx）。 |
| `ops` | array<object> | ✅ | 按顺序执行的操作列表 |

### `office_fill_docx_template`

用数据填充 .docx 模板中的 {{变量}} 占位符（docxtemplater）；模板套打用本工具。data 的键即变量名，支持嵌套如 {{a.b}}。不支持循环/条件语法；批量套打时多次调用即可。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `templatePath` | string | ✅ | 模板文件路径（.docx，内含 {{变量}}） |
| `outputPath` | string | ✅ | 输出文件路径（.docx） |
| `data` | object | ✅ | 变量名→值 的对象 |

### `office_convert`

格式转换（含旧格式与 PDF）。Word 家族：doc/docx/rtf/odt/html/txt/md 互转；表格家族：xlsx/xls/xlsb/ods/csv/tsv/html 互转（含 .xls 导出）。要把文件另存为另一种格式时用本工具，不要用通用读写工具搬运二进制内容。目标 .pdf 需要 LibreOffice 或 Microsoft Word；docx 读取走 mammoth，其余走本机转换器或纯 JS 重建。不支持 Word↔表格。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `sourcePath` | string | ✅ | 源文件路径 |
| `outputPath` | string | ✅ | 目标文件路径（扩展名决定目标格式） |
