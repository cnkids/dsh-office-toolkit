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
- **公式默认给算出来的值**，要看公式本体就传 `formulas: "formula"`（或 `"both"` 显示「公式 → 值」，`true` 等同 `"formula"`）；共享公式给出引用地址，`.xls` / `.xlsb` / `.ods` 同样支持。
- **自动编号会展开成文字**：Word 里的多级编号（样式联动或段落 `numPr`）读出来是 `一、`、`（一）`、`1.1` 这样的文字前缀，行文规则比对才认得出层级；项目符号（`●`/`•`）保持列表结构、不塞进正文。展开段数会写在 `meta.messages` 里。

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

**多表连接**（`join`）在聚合之前执行：另一张表可以在别的文件里，也可以是同一文件的另一张工作表。

```json
{ "path": "订单.xlsx", "join": { "path": "客户.xlsx", "sheet": "客户", "on": {"left": "客户编号", "right": "编号"}, "type": "left" },
  "groupBy": ["地区"], "aggregate": [{"col": "金额", "fn": "sum"}] }
```

| `join` 字段 | 说明 |
| --- | --- |
| `path` | 另一张表所在文件（省略表示同一文件） |
| `sheet` | 那张文件里的工作表（默认第 1 个） |
| `on` | 等值连接键：列名（两表同名）或 `{left, right}`；数组表示多列组合键 |
| `type` | `inner`（默认，只留两边都匹配的）/ `left` / `right` / `full` |
| `suffix` | 右表重名列的后缀，默认 `_2`（左右同名的**连接键**只保留一份） |
| `headerRow` | 右表的表头行号，默认 1 |
| `as` | 给右表起名（同时作为列前缀，避免重名） |

`join` 也可以传数组，按顺序依次连接。连接键的 `1` 与 `"1"`、`2025/1/1` 与 `2025-01-01` 视为同一个值；**空键不参与匹配**（与 SQL 一致）。结果下方会写明匹配了多少行、哪边有多少行没配上 —— 对不上时先看这行，多半是编号格式或前后空格的问题。

**透视表**（`pivot`）把行维度 × 列维度 × 指标摆成矩阵：

```json
{ "path": "订单.xlsx", "pivot": { "rows": ["地区"], "columns": "产品", "values": [{"col": "金额", "fn": "sum"}], "totals": true } }
```

- `rows` 行维度（可多列）；`columns` 列维度（取值自动展开成列，按首次出现顺序；单指标时直接用取值当表头）；`values` 指标（写法同 `aggregate`，可给多个）。
- `totals: true` 追加「合计」行与「合计」列。
- 没有数据的格子留空（不是 0）。列维度取值超过 60 个会报错，提示先用 `where` 收窄或把它放进 `rows`。
- 与 `groupBy` / `aggregate` 二选一；`where`、`orderBy`、`limit` 都照常生效。

**不给 `groupBy` / `aggregate` 时返回表结构画像**：每列的类型（数字 / 日期 / 文本 / 混合 / 空）、非空、空值、去重数、最小 / 最大 / 求和 / 均值、最高频的 3 个取值与次数 —— 一次调用就知道这张表长什么样，不用先读一遍再想怎么问。画像同样接受 `where`，可以只看子集。

数值识别做的是"尽力而为"：`1,234.00`、`¥88`、`12.5%`（→ `0.125`）、全角空格都能认；日期把 `2025/1/1`、`2025年1月1日`、带时间的写法统一成 `YYYY-MM-DD` 再比较。反过来，**认不出就报错或跳过，绝不猜**：

- `sum` / `avg` 跳过非数值单元格，并在结果下方写明跳过了几个（不会静默当成 0）。
- 某一组一个数值都没有时，`sum` / `avg` 返回空单元格而不是 `0`。
- 一边能解析成数字、另一边不能（如「金额」列里混着"待定"）时判为**不可比**，`>` / `=` 一律不成立 —— 否则字典序会让"待定" > 1000 静默成立。
- 列名写错、算子写错都会直接报错，并列出全部可用列名 / 合法算子。
- 文本排序按 Unicode 码位（结果确定，不依赖系统的排序规则）；中文不等于拼音序。

### 什么时候该退回脚本

本工具只做**声明式**操作。以下需求它覆盖不了，请照常写 Python 或别的办法：

- 窗口函数、累计值、同比环比；
- 中位数、分位数、标准差、回归等统计量；
- 图表、可视化、导出报告。

一句话：**表格的筛选与汇总先问 `office_query`，它答不了再写脚本。**

## 新建

文档内容三选一：`html` / `markdown` / `text`。`office_write_docx` 支持标题、段落、粗体、斜体、下划线、删除线、等宽、上下标、列表（含嵌套）、表格、引用、代码块、分隔线、软换行、**图片**，以及 A4 纵向 / 横向与自定义页边距。

**图片**写 `<img src="图片/logo.png" alt="徽标">` 或 Markdown 的 `![徽标](图片/logo.png)`：`src` 是相对**输出文件所在目录**的路径（`~/` 与绝对路径也认），也可以是 `data:image/png;base64,…`；读取与主文件走同一套沙箱路径解析（core 只拿字节，不自己碰文件系统）。尺寸按图片真实像素等比放进正文宽度内，`width="120"` 或 CSS `width:50%` 可指定宽度。转换（`.md`/`.html` → `.docx`）时相对路径按**源文件所在目录**解析。不支持远程 URL（报 `IMAGE_REMOTE`），只支持 PNG / JPEG / GIF / BMP、单张 ≤ 8 MB。

### 页眉 / 页脚 / 页码 / 目录

```json
{ "path": "通知.docx", "markdown": "# 关于……的通知\n\n正文",
  "header": { "text": "XX单位文件", "align": "center", "fontSizePt": 14, "bold": true },
  "footer": { "pageNumber": "第 {page} 页 共 {total} 页", "align": "center", "fontSizePt": 12 },
  "toc": { "title": "目　录", "levels": 3 } }
```

| 参数 | 说明 |
| --- | --- |
| `header` / `footer` | `{ text?, align?, bold?, fontSizePt?, pageNumber? }`；`align` 取 `left`/`center`/`right` |
| `pageNumber` | 页码**模板**：`{page}` 是当前页、`{total}` 是总页数，如 `"第 {page} 页 共 {total} 页"`；公文页码写 `"— {page} —"` |
| `toc` | `true`（标题「目录」、收录 3 级）或 `{ title?, levels? }`，`levels` 取 1–9 |

页码与目录都是 **Word 域**：生成时只写入域代码，Word 打开时按当前排版算出真实页数与目录条目（插件已设置 `updateFields`，所以打开即刷新；若没刷新，按 F9 手动更新一次）。

### 排版：字体 / 行距 / 首行缩进

两种写法，可以混用：**整篇默认**用 `style` 参数，**逐段差异**写在 html 的内联样式上。

```json
{ "path": "通知.docx", "markdown": "# 关于……的通知\n\n正文第一段。",
  "style": { "font": "仿宋_GB2312", "sizePt": 16, "lineSpacingPt": 28.8, "firstLineIndentChars": 2, "align": "both",
             "headings": { "font": "黑体", "sizePt": 16 } } }
```

| `style` 字段 | 说明 |
| --- | --- |
| `font` / `fontAscii` | 字体；`font` 会写进 `w:eastAsia`（**不写这个，Word 不会用中文字体渲染汉字**），`fontAscii` 单独设西文，不传则同 `font` |
| `sizePt` | 字号（磅）：三号=16、四号=14、小四=12、二号=22、小二=18 |
| `lineSpacingPt` | 行距**固定值**（磅），公文常用 28.8；与下面二选一 |
| `lineSpacingMultiple` | 行距**倍数**，如 1.5 |
| `firstLineIndentChars` / `firstLineIndentPt` | 首行缩进（字符数 / 磅值），公文常用 2 字符；`0` 表示不缩进 |
| `align` | `both`（两端对齐）/ `center` / `left` / `right` |
| `spacingBeforePt` / `spacingAfterPt` | 段前 / 段后距（磅） |
| `headings` | 各级标题（Heading1–6）的排版，写法同本对象 |
| `table` | 表格样式，见下 |

这些都会写进文档默认样式（`docDefaults`），所以正文段落不用逐段声明。

表格样式写在 `style.table` 里（markdown 的表格、html 的 `<table>` 都适用）：

```json
{ "path": "人员表.docx", "markdown": "| 姓名 | 部门 |\n| --- | --- |\n| 张三 | 技术研发中心 |",
  "style": { "table": { "borders": "three-line", "headerShading": "F2F2F2", "headerBold": true,
                        "columnWidthMode": "auto", "align": "center" } } }
```

| `style.table` 字段 | 说明 |
| --- | --- |
| `borders` | `all` 全部框线 / `none` 无 / `outline` 仅外框 / `three-line` 三线表（上下粗线 + 表头下细线） |
| `headerShading` | 表头底纹 `RRGGBB`；传 `"none"` 去掉底纹 |
| `headerBold` | 表头是否加粗，默认 `true` |
| `columnWidthMode` | `auto` **按内容自动分配**（谁内容长谁宽，推荐）/ `manual` 用 `columnWidths` |
| `columnWidths` | 各列百分比，如 `[30, 40, 30]`（按比例理解，不必凑满 100） |
| `align` | 表格整体对齐：`left` / `center` / `right` |
| `cellMargins` | 单元格内边距（twip，1pt=20）：`{"left":108,"right":108,"top":40,"bottom":40}`；没给的边用 Word 默认值（左右 108、上下 0） |
| `cellVerticalAlign` | 单元格内容垂直对齐：`top` / `center` / `bottom` |
| `rowHeightPt` | 每行最小行高（磅） |
| `repeatHeader` | `true` 时表头行跨页重复（长表格翻页后仍看得到表头） |
| `cantSplit` | `true` 时禁止同一行被分页断开 |

自动列宽会给每列至少 4%，并做归一化 —— 某一列内容特别长时也不会把其他列挤没。指定列宽时会同时切到固定布局（`w:tblLayout fixed`），Word 里拖列宽也不会跑。

内联样式同样生效，而且 `<body>` / `<div>` 上的样式会被子元素继承：

```html
<body style="font-family: 仿宋_GB2312, Times New Roman; font-size: 16pt; line-height: 28.8pt; text-indent: 2em">
  <p>正文段落（继承上面的字体与行距）</p>
  <p style="font-family:黑体;font-size:16pt;text-align:center;margin:12pt 0">单独指定的一段</p>
</body>
```

| 内联属性 | 说明 |
| --- | --- |
| `font-family` | 写两个逗号分隔时，**第一个当中文字体、第二个当西文字体** |
| `font-size` | `pt` / `px`（按 96dpi） |
| `line-height` | 带单位=固定值（如 `28.8pt`）；纯数字=倍数（如 `1.5`） |
| `text-indent` | `em`（按该段字号）/ `pt` / `px`；无单位按 pt |
| `margin` / `margin-top` / `margin-bottom` | 段前段后（支持 1/2/3/4 值简写） |
| `text-align` | `justify`（两端对齐）/ `center` / `left` / `right` / `start` / `end` |
| `color` / `background-color` / `font-weight` / `font-style` / `text-decoration` | 同前 |

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

## 编辑已有 Word 文档

改已有 `.docx` 的正文用 `office_edit_docx`（**不要**用 `office_write_docx` 覆盖原文件——那是整篇重建，原格式、图片、页眉页脚全丢）。只重写正文部件，其余部件原样保留。

```json
{ "path": "合同.docx", "ops": [{ "op": "replace_text", "find": "100 元", "replace": "200 元" }] }
{ "path": "合同.docx", "ops": [{ "op": "set_paragraph", "match": "乙方：某某贸易有限公司", "text": "乙方：某某物流有限公司" }] }
{ "path": "合同.docx", "ops": [{ "op": "insert_paragraph", "text": "第三条 补充条款", "heading": 3, "position": "end" }] }
{ "path": "合同.docx", "ops": [{ "op": "delete_paragraph", "paragraph": 6 }] }
{ "path": "合同.docx", "ops": [{ "op": "set_style", "scope": "all", "font": "仿宋_GB2312", "sizePt": 16, "lineSpacingPt": 28.8, "firstLineIndentChars": 2, "align": "both" }] }
{ "path": "合同.docx", "ops": [{ "op": "set_style", "scope": "body", "lineSpacingMultiple": 1 }] }
{ "path": "合同.docx", "ops": [{ "op": "set_style", "match": "合同标题", "font": "黑体", "sizePt": 22, "align": "center", "firstLineIndentChars": 0 }] }
{ "path": "合同.docx", "ops": [{ "op": "set_table", "table": 1, "borders": "three-line", "columnWidthMode": "auto" }] }
```

改已有文档的表格：

```json
{ "path": "人员表.docx", "ops": [
  { "op": "set_table", "table": 1, "borders": "all", "headerShading": "F2F2F2", "align": "center" },
  { "op": "delete_table", "table": 2 },
  { "op": "insert_table_row", "table": 1, "at": 2, "count": 1 },
  { "op": "delete_table_column", "table": 1, "at": 3 },
  { "op": "merge_table_cells", "table": 1, "range": "A1:B1" } ] }
```

`set_style` 的作用范围三选一：`scope: "all"`（全篇所有段落，含表格内）、`from`/`to`（段落区间，含两端）、`paragraph`/`match`（单段）。它写的是**段落直接格式**（优先级高于样式表默认），所以一定生效；文档默认块（`docDefaults`）保持原样。

| 操作 | 参数 | 说明 |
| --- | --- | --- |
| `replace_text` | `find`、`replace`、`limit` | 正文内查找替换。Word 常把一句话拆进多个 run，本工具**跨 run 匹配**；`replace` 不传表示删除，`limit` 限制最多替换几处 |
| `set_paragraph` | `text` + 定位 | 整段改写：保留段落自己的样式（标题还是标题）与该段首个 run 的字符格式 |
| `insert_paragraph` | `text`、`heading?`、`position?` | 插入段落。`position` 取 `end`（默认，文末，自动落在页面设置之前）/ `start` / `before` / `after`（后两者配合定位）；`heading` 传 1–9 套用 Word 内置样式 `Heading1`–`Heading9`，插入后即出现在标题大纲里 |
| `set_style` | 定位 + 样式字段 | 改排版：字体 / 字号 / 行距 / 首行缩进 / 对齐 / 段前段后 / 加粗 / 字色，字段与上面的 `style` 相同。只覆盖你给出的属性，段落原有的加粗、字号等其它格式保持不动 |
| `set_table` | `table`/`scope` + `style.table` 的字段 | 改表格样式：框线 / 表头底纹 / 表头加粗 / 列宽 / 单元格内边距 / 表格对齐 |
| `delete_table` | `table`/`scope`/`onlyEmpty` | **删除整张表格**（连 0 行的空表也能删）。`onlyEmpty: true` 一次清掉所有 0 行残留，不必数序号。删完会自动保证正文不为空、且不以表格结尾（否则 Word 会提示修复） |
| `insert_table_row` / `delete_table_row` | `table`/`scope`、`at`、`count?` | 增删行。`insert` 的 `at` 表示插在第几行**之前**（省略则追加到末尾）；`delete` 的 `at` 表示从这里开始删 |
| `insert_table_column` / `delete_table_column` | 同上 | 增删列（含 `w:tblGrid` 同步） |
| `merge_table_cells` | `table`/`scope`、`range` | 合并矩形区域，如 `"A1:B2"`（含跨行跨列），**保留左上角单元格的内容**（被覆盖单元格的内容会丢弃） |
| `unmerge_table_cells` | `table`/`scope`、`range` | 取消合并：去掉 `gridSpan`/`vMerge` 并把占位补回成独立空单元格；区域未合并时报 `TABLE_NOT_MERGED` |
| `set_numbering` | 定位/`scope`、`style`、`linkToHeading?`、`exclude?`、`startFrom?` | 多级自动编号：一级 `1.`、二级 `1.1`、三级 `1.1.1`（`style: "multicol-1_1_1"`）。默认 `linkToHeading: true`，把 1–9 级编号按样式挂到 `Heading1–Heading9`，标题不用再手写编号文字；同时把**本来就是列表项**的正文段落挂到当前标题层级之下 |
| `delete_paragraph` | 定位 | 删除整段 |

**段落定位**二选一，三种操作都用这一套：

- `paragraph`：段落序号，从 **1** 开始，按文档顺序编号（表格单元格里的段落也计入）。
- `match`：整段原文（首尾空格忽略）。必须唯一，否则报错并给出候选序号，提示你改用序号或写更长的 match。

**`set_style` 的段落选择有「角色」维度**（这是「只把正文改成单倍行距、标题不动」的正解）：

| `scope` | 命中 |
| --- | --- |
| `body` | 正文段落（**不含标题**；表格内的段落也算正文） |
| `headings` | 标题段落（按 `pStyle` 识别 Heading1–9 / 标题 1 / 1，或 `outlineLvl`） |
| `table` | 表格内的段落 |
| `all` | 全部段落 |

`scope` 还能与 `from`/`to` 区间叠加（如 `scope: "body"` + `from: 3, to: 9`），或者用 `paragraph`/`match` 只改一段。都不写就是单段定位。

表格类操作用 `table` 序号选表（从 1 开始；文档里只有一个表格时可以省略），或 `scope: "all"` 改所有表格。

序号与内容先用 `office_read`（或 `outline: true`）看一眼最稳。多个操作**按顺序执行**，后面的 `paragraph` / `table` 序号以执行后的文档状态为准。

**改之前可以先试运行，也可以另存不覆盖原文件**：

```json
{ "path": "合同.docx", "dryRun": true, "ops": [{ "op": "set_style", "scope": "body", "lineSpacingMultiple": 1 }] }
{ "path": "合同.docx", "outputPath": "合同-改后.docx", "ops": [{ "op": "delete_table", "table": 2 }] }
```

`dryRun: true` 只把「将要执行的操作」摘要报出来、**不写盘**（原文件与 `outputPath` 都不动）；`outputPath` 另存为新文件、原文件保持不变。改动多或把握不准时，先 `dryRun` 看一眼再落盘。

两点如实说明：替换文字会沿用**匹配起点所在 run** 的字符格式（跨 run 替换时，加粗范围可能跟着变，例如 `**金额**为 100` 换成 `金额为 200` 后加粗覆盖到整个替换串）；`office_edit_docx` 只处理正文，页眉页脚、脚注、文本框里的文字不在范围内。

### 多级编号（标题自动编号）

```json
{ "path": "方案.docx", "ops": [{ "op": "set_numbering", "scope": "all", "style": "multicol-1_1_1", "linkToHeading": true }] }
```

一次调用做两件事：

1. **标题**用 Word 多级自动编号（预置见下表）：`1.` / `1.1` / `1.1.1`，或公文式的 `一、` / `（一）` / `1.`（`numbering.xml` 里新增一个 9 级 `abstractNum`，1–9 级分别 `<w:pStyle>` 挂到 `Heading1–Heading9`）—— 所以标题段落**不需要**写 `<w:numPr>`，也不用手写「一、」「1.1」这类编号文字（已有的手写编号用 `replace_text` 去掉即可）。改标题、调顺序时 Word 会自动重排。
2. **正文列表**（本来就是列表项、带 `numPr` 的段落）挂到**当前标题层级的下一个级别**：`Heading1` 下是 `1.1 / 1.2`，`Heading2` 下是 `1.1.1 / …`。普通正文段落不动，不会被变成列表。

| 参数 | 说明 |
| --- | --- |
| `style` | 预置方案：`multicol-1_1_1`（一级 `1.`、二级 `1.1`、三级 `1.1.1`…，西式）；`gongwen-1_1_1_1`（`一、` → `（一）` → `1.` → `（1）`，GB/T 9704 公文层次，中文数字用 `chineseCounting` 数字格式） |
| `linkToHeading` | 默认 `true`：编号按样式挂到 `Heading1–Heading9`（标题自动编号）。传 `false` 则不给样式加链接，只把落进选择范围的标题段落**显式**写 `numPr` |
| `exclude` | 不参与编号的标题样式，如 `["Heading1"]` —— 既不编号也不占层级（链会自动往上收） |
| `startFrom` | `{"Heading2": 1}`：从该标题样式起算，比它浅的标题样式自动不参与；值是**链上第 1 级的起始数字**（默认 1，给 `{"Heading1": 5}` 就从“5.”开始排） |
| 定位 | 与 `set_style` 相同：`scope`（`all` / `body` / `headings` / `table`）± `from`/`to`，或 `paragraph`/`match`；都不给就是全篇 |

`numbering.xml` 不存在时（少见）会连同关系文件与 `[Content_Types].xml` 里的 Override 一起补出来。

**表格删空后想彻底清掉**：把行删光的表格会留下 0 行 `<w:tbl>`（Word 通常不渲染，但 XML 里还在），此时 `delete_table_row` / `delete_table_column` 会报错并提示改用 `delete_table`：

```json
{ "op": "delete_table", "table": 2 }
{ "op": "delete_table", "scope": "all", "onlyEmpty": true }
```

空表不渲染、肉眼数不出序号，所以清残留推荐用下面那条（`onlyEmpty: true`）。

**幂等**：本工具加的那份编号带固定的 `w:nsid`，重复调用会**原地替换**各级定义（`numId` 保持不变，段落里的引用不会失效），不会像以前那样每次堆积一份新的 `abstractNum`。

**跳过的情况会如实报出**：如果列表项上方还没有「参与编号的标题」（例如 `Heading1` 被 `exclude` 掉了、而该列表直接挂在 `Heading1` 下），它拿不到正确的编号，工具会跳过它并在结果里写「跳过 N 段」，而不是给出错误编号。

编号级别的缩进统一设成 0（`w:ind left=0 hanging=0`），不改变文档原有的缩进/首行缩进设置；编号文字直接接在段落开头。

**一点如实说明**：多级编号里同一级别共用一个计数器 —— 如果文档同时用了 `Heading2`，那么「`Heading1` 下的正文列表项」与「`Heading2`」都在第 1 级（都显示成 `1.x`）。这是 Word 多级列表的机制，不是本工具的限制；对「一级标题 + 其下正文列表」这种最常见结构，编号完全正确。

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
`add_chart`、`conditional_format`、`data_validation`。单个操作失败时加 `"optional": true` 可跳过而不中断整批。

### 条件格式与数据验证

```json
{ "op": "conditional_format", "sheet": "明细", "range": "D2:D100", "rules": [
  { "type": "cellIs", "operator": "greaterThan", "value": 1000, "style": { "color": "9C0006", "fill": "FFC7CE", "bold": true } },
  { "type": "colorScale", "colors": ["F8696B", "FFEB84", "63BE7B"] },
  { "type": "dataBar", "color": "638EC6" } ] }
{ "op": "data_validation", "sheet": "明细", "range": "C2:C100",
  "rule": { "type": "list", "values": ["待办", "进行中", "已完成"], "promptTitle": "选择状态", "prompt": "从下拉里选", "errorTitle": "值不对", "error": "只能选下拉里的值" } }
```

| `conditional_format` 的 `rule.type` | 参数 |
| --- | --- |
| `cellIs` | `operator`（`equal notEqual greaterThan lessThan greaterThanOrEqual lessThanOrEqual between notBetween`，默认 greaterThan）、`value` / `value2`、`style` |
| `expression` | `formula`（如 `MOD(ROW(),2)=0`）、`style` |
| `colorScale` | `colors`（2 或 3 个 RRGGBB，默认白→蓝；3 个时中间是百分位 50%） |
| `dataBar` | `color`（默认 638EC6） |
| `iconSet` | `iconSet`（如 `3TrafficLights1`、`5Arrows`）、`showValue` |
| `top10` | `rank`、`percent`、`bottom`（>N 或后 N 名） |
| `aboveAverage` | `aboveAverage`（false = 低于平均值） |
| `containsText` | `operator`（`containsText notContains beginsWith endsWith containsBlanks notContainsBlanks containsErrors notContainsErrors`）、`text`、`style` |
| `timePeriod` | `timePeriod`（`today yesterday tomorrow last7Days thisWeek lastWeek nextWeek thisMonth lastMonth nextMonth`） |

`style` 用与 `style_range` 相同的字段（`bold italic fontSize color fill numFmt border borderColor`），会写进差异格式（dxf）。同一区域多条规则用 `rules: [...]`（优先级按顺序递增），单条也可以用 `rule`。`data_validation` 同理：`type` 取 `list`（`values` 候选数组，或 `source: "=Sheet2!$A$1:$A$5"` 引用区域）/ `whole` / `decimal` / `date` / `textLength`（`operator` + `value`/`value2`）/ `custom`（`formula`），可选 `allowBlank`、`promptTitle`、`prompt`、`errorTitle`、`error`、`errorStyle`。

## 模板套打

模板里写 `{{变量}}`，支持 `{{a.b}}` 形式的嵌套取值与中文变量名。

```json
{ "templatePath": "合同模板.docx", "outputPath": "合同-001.docx",
  "data": { "合同编号": "HT-2026-001", "甲方": "某某公司", "金额": "12800" } }
```

**数组循环与条件**（docxtemplater 原生支持，写在模板文件里）：

| 模板写法 | 效果 |
| --- | --- |
| `{{#items}}…{{/items}}` | 该段按 `items` 数组重复；可嵌套；传空数组时整块消失 |
| `{{#flag}}…{{/flag}}` | `flag` 为真时渲染这一块 |
| `{{^flag}}…{{/flag}}` | `flag` 为假（或空数组）时渲染这一块 |

```json
{ "templatePath": "验收单模板.docx", "outputPath": "验收单-001.docx",
  "data": {
    "项目名称": "办公楼改造",
    "items": [ { "名称": "水泥", "数量": 120 }, { "名称": "钢筋", "数量": 80 } ],
    "加急": false
  } }
```

模板里写 `{{#items}}{{名称}} × {{数量}}{{/items}}`，就会一次生成任意行数的明细；`{{^加急}}按常规流程办理{{/加急}}` 只在非加急时出现。注意：这些双花括号**只能出现在模板文件里** —— 写进工具描述或 `AGENTS.md` 会被 DSH 当成提示词变量而报错。

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
- 图片嵌入只认**本地文件**与 `data:` URL（PNG / JPEG / GIF / BMP，单张 ≤ 8 MB）：按魔数读真实像素尺寸，超出正文宽就等比缩到页内；`<img width="120">` 或 CSS `width:50%` 可指定宽度。远程地址（`http/https`）直接报 `IMAGE_REMOTE` —— 本插件不联网，请先下载到本地；缺失文件、不支持的格式（SVG / WebP / TIFF）也各自报错，不会静默丢图。
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

以下是 8 个工具注册到 DSH 的完整参数表（由工具 schema 自动导出）。

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
| `formulas` | string | — | 仅表格：`value`（默认，计算值）/ `formula`（公式本体）/ `both`（公式 → 值）；传 `true` 等同 `formula` |
| `withFormatting` | boolean | — | 仅 .docx：额外返回格式报告 —— 每段的字体(中文/西文)、字号、行距(固定值/倍数)、首行缩进、对齐、样式名，以及页面尺寸与页边距；并给出格式分布(主流值)与偏离主流的段落。用于比对行文规则(如"正文三号仿宋、行距固定值 28.8 磅")。结构化数据在 meta.formatting。 |

### `office_query`

在表格文件内直接算（.xlsx .xls .xlsb .ods .csv .tsv）：全表扫描，只把结论返回，几万行不必读进上下文。表格数据的求和/均值/计数/去重/分组/排序/条件筛选，优先用本工具而不是临时写脚本；返回的内容是结果表，不是原始数据。不给 groupBy/aggregate 时返回**表结构画像**。多表连接（join：另一个文件或同一文件的另一张表，按列等值连接，inner/left/right/full）与透视表（pivot：rows × columns × values，可带合计）也在这里做，不必写脚本；窗口函数、统计建模这类才需要另想办法。列名取表头行（headerRow，默认第 1 行）；条件算子：eq/ne/gt/gte/lt/lte/contains/startsWith/endsWith/in/notIn/isBlank/notBlank；聚合函数：sum/avg/min/max/count/countDistinct（数值列能识别 ¥1,234.00、12.5% 这类写法）。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 |
| `sheet` | — | — | 工作表名或从 1 开始的序号，默认第 1 个工作表 |
| `headerRow` | integer | — | 表头行号，默认 1；传 0 表示没有表头（列名自动取 A/B/C…） |
| `where` | array<object> | — | 筛选条件（多个条件为 AND），如 `[{"col":"地区","op":"eq","value":"华东"}]`；`col` 必填，`op` 默认 eq，`value` 为比较值（`in`/`notIn` 传数组，`isBlank`/`notBlank` 不需要） |
| `join` | object \| array | — | 先连另一张表再算：`{path?, sheet?, on, type?, suffix?, headerRow?, as?}`，数组表示依次连接多条 |
| `pivot` | object | — | 透视表：`{rows, columns, values, totals?}`；与 `groupBy`/`aggregate` 二选一 |
| `groupBy` | array<string> | — | 按这些列分组，如 `["地区","产品"]`；只给 `groupBy` 时输出每组的行数 |
| `aggregate` | array<object> | — | 汇总项 `[{col, fn, as}]`；`fn` 取 sum/avg/min/max/count/countDistinct，`as` 为结果列名（默认 `fn(列名)`） |
| `orderBy` | array<object> | — | 排序如 `[{"col":"销售额","dir":"desc"}]`，`dir` 取 asc/desc（默认 asc） |
| `limit` | integer | — | 结果行数上限，默认 200，最大 2000 |

### `office_write_docx`

新建 Word 文档（默认 .docx，也支持 .doc / .rtf / .odt）。要生成 Word 文档时用本工具，不要用通用 write 工具（写不出合法的 .docx）。内容三选一：html / markdown / text，支持标题、段落、粗体斜体、列表、表格、引用、图片等（图片写 `<img src="图片/a.png" alt="说明">` 或 Markdown 的 `![说明](图片/a.png)`：本地文件或 `data:` URL，PNG/JPEG/GIF/BMP，单张 ≤ 8 MB，按真实像素等比缩放到正文宽内；远程地址会报错，不联网）。.doc/.rtf/.odt 旧格式输出需要本机转换器（macOS textutil / LibreOffice / Microsoft Word），不可用时会返回安装提示；图片等复杂元素建议用 .docx。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 扩展名决定输出格式：.docx（推荐）/.doc/.rtf/.odt。 |
| `html` | string | — | HTML 正文（与 markdown/text 三选一） |
| `markdown` | string | — | Markdown 正文（与 html/text 三选一） |
| `text` | string | — | 纯文本正文（与 html/markdown 三选一） |
| `title` | string | — | 文档标题（元数据） |
| `landscape` | boolean | — | 是否横向纸张（仅 .docx） |
| `marginsMm` | number | — | 页边距毫米（仅 .docx） |
| `header` / `footer` | object | — | 页眉 / 页脚：`{text, align, bold, fontSizePt, pageNumber}`；`pageNumber` 是模板，`{page}`=当前页、`{total}`=总页数（Word 域） |
| `toc` | object | — | 目录：`true` 或 `{title, levels}`；插入 Word 的 TOC 域，打开时自动生成 |
| `style` | object | — | 整篇排版：字体 / 字号 / 行距 / 首行缩进 / 对齐 / 段前段后 / `headings`（标题）/ `table`（表格框线·表头底纹·列宽·单元格内边距·对齐）。公文体例 `{"font":"仿宋_GB2312","sizePt":16,"lineSpacingPt":28.8,"firstLineIndentChars":2,"align":"both","table":{"borders":"three-line","columnWidthMode":"auto"}}` |

### `office_edit_docx`

就地修改已有 Word 文档（.docx）的正文：查找替换、整段改写、插入段落、删除段落。改已有 Word 文档必须用本工具 —— 通用 edit 工具改不了二进制文档（会报 binary file），而 office_write_docx 是整篇重建（原格式、图片、页眉页脚全丢）。本工具只重写正文部件，其余部件原样保留。段落定位二选一：paragraph（段落序号，从 1 开始）或 match（整段原文，必须唯一）。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。必须为 .docx（旧 .doc 先用 office_convert 转 .docx）。 |
| `ops` | array<object> | ✅ | 操作列表，逐项含 `op` 与各自参数：`replace_text`(`find`/`replace`/`limit`)、`set_paragraph`(`text`+定位)、`set_style`(角色 `scope` = all/body/headings/table，或 `from`-`to`、`paragraph`/`match`)；`set_numbering`(多级自动编号，`style` + `linkToHeading`)；表格 `set_table`(含 `cellMargins`)、`insert_table_row`/`delete_table_row`、`insert_table_column`/`delete_table_column`(`table` 序号 + `at`/`count`)、`merge_table_cells`/`unmerge_table_cells`(range 如 "A1:B2")；`insert_paragraph`(`text`/`heading`/`position`)、`delete_paragraph`(定位) |

### `office_write_xlsx`

新建 Excel 工作簿（.xlsx），可一次写入多个工作表；要生成表格文件时用本工具。rows 为二维数组；单元格值规则：数字/布尔按原样，字符串以 "=" 开头视为公式（如 "=SUM(B2:B9)"），"date:2026-09-09" 写入日期。header:true 会给首行加粗底纹；columnWidths 可设列宽。需要图表时先写数据，再用 office_edit_xlsx 的 add_chart。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 必须为 .xlsx。 |
| `sheets` | array<object> | ✅ | 工作表数组 |

### `office_edit_xlsx`

修改已有 .xlsx（按顺序执行 ops）。改表格必须用本工具，通用 edit 工具改不了二进制表格。支持操作：set_value{sheet,ref,value,style}、set_cells{sheet,start,values,style}、set_formula{sheet,ref,formula}、style_range{sheet,range,style}、merge/unmerge{sheet,range}、insert_rows/delete_rows{sheet,at,count}、insert_cols/delete_cols{sheet,at,count}、add_sheet{name,rows,header}、rename_sheet{sheet,name}、delete_sheet{sheet}、set_col_width{sheet,col,width}、set_row_height{sheet,row,height}、freeze{sheet,rows,cols}、auto_filter{sheet,range}、add_image{sheet,path,cell|range}、add_chart{sheet,chartType,categories,series,title,anchor}、conditional_format{sheet,range,rule|rules}、data_validation{sheet,range,rule}。style 支持 {bold,italic,fontSize,color,fill,align,valign,wrap,numFmt,border}（颜色为 RRGGBB）。add_chart 的 chartType 支持 bar/column/line/pie；categories 如 "A2:A6"，series 如 ["B2:B6"]（可带 label 单元格）。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `path` | string | ✅ | 文件路径。可用绝对路径，或相对会话工作区的路径。 必须为 .xlsx（旧 .xls 先用 office_convert 转 .xlsx）。 |
| `ops` | array<object> | ✅ | 按顺序执行的操作列表 |

### `office_fill_docx_template`

用数据填充 .docx 模板里的占位符（docxtemplater）；模板套打用本工具。占位符写法：变量名左右各加两个半角花括号。data 的键即变量名，支持点号嵌套路径（如 a.b）。支持数组循环（{{#items}} … {{/items}}）与条件（{{#flag}}…{{/flag}}、{{^flag}}…{{/flag}}），写在模板文件里。注意双花括号只能出现在模板文件里，写进工具描述或 AGENTS.md 会被 DSH 当成提示词变量而报错。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `templatePath` | string | ✅ | 模板文件路径（.docx，内含 {{变量}}） |
| `outputPath` | string | ✅ | 输出文件路径（.docx） |
| `data` | object | ✅ | 变量名→值 的对象；值为数组时配合模板里的 `{{#数组名}}…{{/数组名}}` 做循环，布尔值配合 `{{#标记}}`/`{{^标记}}` 做条件 |

### `office_convert`

格式转换（含旧格式与 PDF）。Word 家族：doc/docx/rtf/odt/html/txt/md 互转；表格家族：xlsx/xls/xlsb/ods/csv/tsv/html 互转（含 .xls 导出）。要把文件另存为另一种格式时用本工具，不要用通用读写工具搬运二进制内容。目标 .pdf 需要 LibreOffice 或 Microsoft Word；docx 读取走 mammoth，其余走本机转换器或纯 JS 重建。不支持 Word↔表格。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `sourcePath` | string | ✅ | 源文件路径 |
| `outputPath` | string | ✅ | 目标文件路径（扩展名决定目标格式） |
