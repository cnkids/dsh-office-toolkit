# 用法与参数

[← 返回 README](../README.md)

## 安装与更新的可靠性

- `dsh plugin` 是 **pnpm 的薄转发**，用的是 PATH 上的 pnpm；**pnpm < 11.10 会把 URL 形式的依赖直接当作本地缓存**（同一个 `latest` 直链再 `add`，服务端收不到任何请求，`--force` / `remove` + `add` 都无效）。先 `pnpm -v`，低于 11.10 请升级。
- `releases/latest/download/...` 的内容会随发版变化，任何缓存层都可能留下旧内容。要绝对可靠就用**带版本号的直链**。
- 装完用启动日志（`[dsh-office-toolkit] vX.Y.Z …`）或 profile 里那份 `package.json` 的 `version` 确认，避免白折腾。

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

```json
{ "path": "报告.docx" }
{ "path": "报告.docx", "format": "html" }
{ "path": "销售.xlsx", "sheets": ["明细"], "range": "A1:F100" }
{ "path": "明细.csv", "maxRows": 50 }
```

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
- **截断**：大表默认 400 行 × 60 列、Word 正文默认 90000 字符，超出会截断并提示用参数分段读。
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

以下是 6 个工具注册到 DSH 的完整参数表（由工具 schema 自动导出）。

### `office_read`

读取 Word / Excel / 表格文件内容（.docx .doc .rtf .odt .xlsx .xls .xlsb .ods .csv .tsv）。用户提到 Word、Excel、文档、表格、.docx、.xlsx 等后缀时一律用本工具。Word/Excel 是二进制格式,通用 read/write/edit 工具处理不了(会报 binary file),必须用本工具。解析方式自动选择：.docx 走 mammoth（Markdown 风格正文/表格）；.doc/.rtf/.odt 优先本机转换器(macOS textutil / LibreOffice / Word)，不可用时回退纯 JS；.xlsx 走 ExcelJS；.xls/.xlsb/.ods/.csv/.tsv 走 SheetJS（纯 JS，全平台可用）。Excel 以 TSV 代码块返回，可用 sheets/range/maxRows 分段读取大表。需要核对字体/字号/行距/缩进等排版格式时，对 .docx 传 withFormatting: true。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 |
| `sheets` | array<string> | — | 仅 Excel：只读取这些工作表（名称；序号请写成字符串，如 "2"） |
| `range` | string | — | 仅 Excel：读取范围，如 A1:F50 |
| `maxRows` | integer | — | 仅 Excel：最多读取行数，默认 400 |
| `maxCols` | integer | — | 仅 Excel：最多读取列数，默认 60 |
| `maxChars` | integer | — | Word 正文返回字符上限，默认 90000 |
| `format` | enum: `text` / `html` | — | Word 输出格式：text=Markdown 风格（默认），html=原始 HTML |
| `withFormatting` | boolean | — | 仅 .docx：额外返回格式报告 —— 每段的字体(中文/西文)、字号、行距(固定值/倍数)、首行缩进、对齐、样式名，以及页面尺寸与页边距；并给出格式分布(主流值)与偏离主流的段落。用于比对行文规则(如"正文三号仿宋、行距固定值 28.8 磅")。结构化数据在 meta.formatting。 |

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
