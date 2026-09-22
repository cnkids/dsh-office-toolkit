<h1 align="center">dsh-office-toolkit</h1>

<p align="center">
  <strong>让 DSH 智能体自己读写 Word / Excel</strong><br>
  纯 JavaScript，不装 Office 也能跑。
</p>

<p align="center">
  <a href="https://github.com/cnkids/dsh-office-toolkit/releases/latest"><img src="https://img.shields.io/github/v/release/cnkids/dsh-office-toolkit?style=flat&amp;label=release&amp;color=4D6BFE" alt="Latest release"></a>
  <a href="https://github.com/cnkids/dsh-office-toolkit/releases"><img src="https://img.shields.io/github/downloads/cnkids/dsh-office-toolkit/total?style=flat&amp;label=downloads&amp;color=4D6BFE" alt="Total downloads"></a>
  <img src="https://img.shields.io/badge/node-%E2%89%A520-339933?style=flat" alt="Node.js 20 or newer">
  <img src="https://img.shields.io/badge/macOS%20%7C%20Windows%20%7C%20Linux-4493F8?style=flat" alt="Supported platforms: macOS, Windows and Linux">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-2EA44F?style=flat" alt="MIT License"></a>
</p>

<p align="center">
  <a href="#安装"><strong>安装</strong></a> ·
  <a href="#快速上手">快速上手</a> ·
  <a href="#工具">工具</a> ·
  <a href="#文档">文档</a> ·
  <a href="#常见问题">常见问题</a> ·
  <a href="#贡献">贡献</a> ·
  <a href="https://github.com/cnkids/dsh-office-toolkit/releases/latest">Releases</a>
</p>

---

> 装上这个宿主插件，智能体就能直接读报告、写台账、改表格、套合同模板、转换旧格式 —— 不用再手动搬文件，也不会因为通用 `read` 读不了二进制而卡住。

## 特性

- **纯 JS，开箱可用** —— 不装 Office / LibreOffice 就能读写 `.docx` `.xlsx` `.xls` `.csv`，也能读 `.pdf` 文本（本机有 `pdftotext` / macOS PDFKit 就用它，否则用随包携带的 pdfjs），Windows 上同样零依赖。
- **读写改一条龙** —— 新建文档、新建工作簿、编辑已有工作簿（样式 / 合并 / 冻结 / 筛选 / 图片 / 图表）、`{{变量}}` 模板套打、跨格式转换。
- **8 个工具自动注册** —— 见[工具](#工具)；描述里前置了文件类型关键字，智能体一次就能选对。
- **大表直接算，不用先搬数据** —— `office_query` 在文件内做筛选 / 分组 / 求和 / 计数 / 去重 / 排序，几万行也只返回结论；不给条件时返回表结构画像。跨文件的**多表连接**（`join`：inner/left/right/full，列名相同或左右对应）与**透视表**（`pivot`：行 × 列 × 指标，可带合计）也在这里做，不用写脚本。
- **排版可指定** —— `style` 参数设整篇字体 / 字号 / 行距 / 首行缩进 / 对齐（公文体例：仿宋_GB2312 三号、行距固定值 28.8 磅、首行缩进 2 字符），标题用 `style.headings`，表格用 `style.table`（三线表 / 表头底纹 / **按内容自动分配列宽**）；html 内联样式也认 `font-family`（含中文 `eastAsia`）/ `line-height` / `text-indent` / `margin`，`<body>`、`<div>` 上的样式会被继承。
- **改之前能先看、改之后能留底** —— `office_edit_docx` 传 `dryRun: true` 只报改动不写盘，传 `outputPath` 另存为新文件而原文件不变。
- **改已有文档能分段、分角色** —— `set_style` 的 `scope` 支持 `body`（正文，不含标题）/ `headings` / `table` / `all`，可叠加段落区间；例如「只把正文改成单倍行距」一句就够。表格另有 `set_table`（框线 / 底纹 / 列宽 / 单元格内边距 / 垂直对齐 / 行高 / 表头跨页重复 / 禁止断行）、增删行列、合并与取消合并，以及 `delete_table` 整表删除（含删空后的 0 行残留）。
- **多级自动编号** —— `set_numbering` 把标题按样式挂上 Word 多级编号（`1.` / `1.1` / `1.1.1`），正文列表挂到当前标题层级之下；支持 `exclude`（某些标题不参与）与 `startFrom`（从哪级起算、起始数字）；**幂等**，重复调用原地替换不堆积定义。编号由 Word 维护，改标题会自动重排，不用手写「一、」「1.1」。
- **页眉页脚、页码、目录** —— `header` / `footer`（页码是 Word 域，如「第 X 页 共 Y 页」或公文式「— 1 —」）、`toc` 插入目录域；插件已设 `updateFields`，Word 打开即按当前排版刷新。
- **图片与编号都落到实处** —— `<img>` 指向本地文件或 `data:` URL 就嵌成真图片（PNG / JPEG / GIF / BMP，按魔数读尺寸、按正文宽等比缩放；不联网，远程地址与超 8 MB 直接报错）；读 Word 时把自动编号展开成文字（`一、` / `（一）` / `1.1`），行文规则比对看得到层级。
- **公式与日期是真的** —— `=` 开头写公式，`date:2026-09-09` 写日期，日期回读不受时区影响。
- **排版格式可读** —— `.docx` 传 `withFormatting: true` 即返回每段的字体 / 字号 / 行距 / 首行缩进 / 对齐与页面页边距（含 `basedOn` 继承链、隐式默认样式与主题字体），并给出格式分布和偏离主流的段落，用于比对行文规则。见[用法与参数](docs/usage.md#读取排版格式行文规则比对)。
- **默认安全** —— 写入经 DSH 沙箱围栏（含 realpath 校验），文本来源中和公式注入，解析全走线性扫描（无 ReDoS），不联网、不常驻后台、依赖树无安装脚本。见[安全说明](docs/security.md)。
- **`npm audit` 0 条** —— 依赖取舍见[设计说明](docs/design.md#依赖与已知告警)。
- **可选的更高保真** —— 装了 LibreOffice / Word / macOS `textutil` 时，旧格式转换自动升级。

## 工具

| 工具 | 作用 |
| --- | --- |
| `office_read` | 读取 `.docx .doc .rtf .odt .pdf .xlsx .xls .xlsb .ods .csv .tsv`（Excel 以 TSV 返回、可切换公式本体 / 计算值；Word / PDF 长文可 `offset` 分段续读、`outline` 取标题大纲或页清单，Word 自动编号展开成文字；`.docx` 可选返回排版格式报告） |
| `office_query` | 在表格文件内计算：条件筛选、分组、求和/均值/计数/去重、排序、表结构画像、多表连接、透视表 —— 只返回结论 |
| `office_write_docx` | 新建文档（html / markdown / text，可嵌本地图片） |
| `office_edit_docx` | **改已有 .docx 正文**：查找替换（跨 run）/ 整段改写 / 插入段落（可套标题样式）/ 删除段落；其余部件原样保留 |
| `office_write_xlsx` | 新建工作簿（多表、表头、公式、日期、样式） |
| `office_edit_xlsx` | 编辑已有工作簿（单元格 / 样式 / 合并 / 行列 / 冻结 / 筛选 / 图片 / 图表 / 条件格式 / 数据验证） |
| `office_fill_docx_template` | `{{变量}}` 模板填充（合同、通知、批量套打） |
| `office_convert` | 格式互转（Word 家族 / 表格家族内部、`.xls` 导出、`.pdf` 输出） |

每个工具的参数表见[用法与参数](docs/usage.md#工具参数)。

## 安装

```sh
dsh plugin --profile web add dsh-office-toolkit
```

装完必须**重启 `dsh web`（或重开桌面端）并新建会话**才会加载。

<details>
<summary>其它安装方式（GitHub 直链 / 完全离线 / 本地源码）</summary>

| 场景 | 命令 |
| --- | --- |
| GitHub 直链（不经过 npm） | `dsh plugin --profile web add https://github.com/cnkids/dsh-office-toolkit/releases/latest/download/dsh-office-toolkit.tgz` |
| 完全离线 | 解压 `dsh-office-toolkit-offline.zip` 后 `dsh plugin --profile web add link:C:/dsh-office-toolkit` |
| 本地源码 | `dsh plugin --profile web add link:/path/to/dsh-office-toolkit` |

离线包内含完整 `node_modules`，全程零下载。目标机器需 **Node ≥ 20** 与 pnpm。

```sh
dsh --profile web --dump-config | grep -A1 dsh-office-toolkit   # 校验是否加载
dsh plugin --profile web remove dsh-office-toolkit              # 卸载
```

</details>

### 更新

```sh
dsh plugin --profile web update
```

装完重启 `dsh web` 并新建会话。确认版本：启动日志里的 `[dsh-office-toolkit] vX.Y.Z …`，或 profile 里那份 `package.json` 的 `version`。

`link:` 安装不需要 pnpm，源码目录 `git pull` 即最新；GitHub 直链的内容会随发版变化，更新时改用上面这条（或先 `remove` 再 `add`）。

## 快速上手

读取（Excel 可用 `sheets` / `range` 限定窗口）：

```json
{ "path": "报告.docx" }
{ "path": "销售.xlsx", "sheets": ["明细"], "range": "A1:F100" }
```

新建文档与工作簿：

```json
{ "path": "周报.docx", "markdown": "# 周报\n\n- 完成 A\n- 完成 B\n" }
{ "path": "台账.xlsx", "sheets": [
  { "name": "明细", "header": true, "columnWidths": [14, 10, 12],
    "rows": [["产品", "数量", "金额"], ["键盘", 10, "=B2*199"], ["鼠标", 20, "=B3*89"]] } ] }
```

编辑已有工作簿（按顺序执行 `ops`，可加图表）：

```json
{ "path": "台账.xlsx", "ops": [
  { "op": "set_value", "sheet": "明细", "ref": "E1", "value": "合计", "style": { "bold": true } },
  { "op": "style_range", "sheet": "明细", "range": "A1:C1", "style": { "fill": "DDEBF7", "align": "center" } },
  { "op": "add_chart", "sheet": "明细", "chartType": "column", "title": "金额对比",
    "categories": "A2:A3", "series": [{ "range": "C2:C3" }], "anchor": "G2" } ] }
```

模板套打（模板里写 `{{合同编号}}`，变量名支持中文）与格式转换：

```json
{ "templatePath": "合同模板.docx", "outputPath": "合同-001.docx",
  "data": { "合同编号": "HT-2026-001", "甲方": "某某公司", "金额": "12800" } }
{ "sourcePath": "旧表.xls", "outputPath": "新表.xlsx" }
{ "sourcePath": "报告.docx", "outputPath": "报告.pdf" }
```

全部操作、参数与限制见[用法与参数](docs/usage.md)。

## 让智能体优先使用本插件

DSH 的工具注册**没有优先级设置**，模型只依据每个工具的 `description` 选择；而内置 `read` / `write` / `edit` 只处理 UTF-8 文本，读 `.docx` 会直接返回 `Error: binary file`，也不会提示该换哪个工具。本插件已把「哪类文件必须用本工具」写在每个描述最前面。

想更保险，在**用户级指令文件** `~/.dsh/AGENTS.md`（对所有项目与会话生效）里加一句：Office 文件一律用 `office_read` / `office_query` / `office_write_*` / `office_edit_xlsx` / `office_fill_docx_template` / `office_convert`，不要用通用 `read`/`write`/`edit`；表格的统计与筛选先用 `office_query`，别为几行汇总临时写脚本。该文件热加载，保存即生效。

## 常见问题

**装完没反应？** 必须重启 `dsh web`（或重开桌面端）并**新建会话** —— 插件与工具列表只在启动时加载。

**Windows 提示 `'git' 不是内部或外部命令`？** `github:` 形式要用 git；改用固定直链即可绕开，或 `winget install --id Git.Git -e`。

**内网 / 无外网怎么装？** 用 `-offline.zip`，解压后 `link:` 安装，全程零下载。

**离线包自带依赖吗？** 是。`.offline.zip` 内含完整 `node_modules`，且发布前会用 `node test/offline-check.mjs` 校验每个依赖都落在包内（不允许借用 profile 目录）、依赖树无安装脚本；解压后 `link:` 安装全程零下载。

**更新后版本没变？** 跑 `dsh plugin --profile web update`；若用的是 GitHub 直链，先 `dsh plugin --profile web remove dsh-office-toolkit` 再 `add`。

**怎么确认插件版本 / 更新生效了？** 插件启动时会往 DSH 日志打印一行 `[dsh-office-toolkit] vX.Y.Z 已注册 7 个 Office 工具`；也可以直接看 profile 里那份 `package.json` 的 `version`（`dsh --profile web --dump-config` 能看到 profile 目录）。报错信息里也会带插件版本，便于排查。

**某份文件读不出来？** 插件已按**真实内容**判断格式：非标准 zip（条目名含反斜杠）、改过后缀的文件（`.doc` 里其实是 docx、`.xlsx` 里其实是 CSV 等）都会自动按实际格式读取并给出提示。若报 `BAD_CONTAINER`，错误信息会列出文件里的实际条目，便于判断它到底是什么。详见[用法与参数](docs/usage.md#读取兼容性)。

**`.xlsx` 读写全失败，报 `createRequire.resolve.paths is not a function` 或 `Cannot find module 'process/'`？** 这是 **DSH 0.1.6-alpha.1 / .2 的模块解析缺陷**，不是插件缺依赖 —— 重装、重启都没用；只有走 ExcelJS 的 xlsx 路径会中招，`.csv` / `.docx` 正常。插件侧无法绕开，请升级到修复版 DSH。报错现在会直接说明这一点（错误码 `DSH_RESOLVER_BUG`），不再误导你重装依赖；DSH 侧分析见 [discussions/7377](https://github.com/deepseek-ai/deepseek-harness/discussions/7377#discussion-10858263)。

**几万行的表怎么统计？** 用 `office_query`，它在文件内算完只回结论，不需要把整张表读进上下文。例如按地区求和、按金额倒序取前 20、只看某段时间的数据，都是一次调用；先不带条件跑一次还能拿到每列的类型、空值、去重数与高频值（表结构画像）。跨文件 join、透视表、统计建模这类插件覆盖不了的，再照常写脚本。

**几百页的 Word 怎么读？** 一次 `office_read` 只返回 `maxChars`（默认 90000 字符）那么长，返回里会写明总字符数和「继续读」该传的 `offset`，下一次带上它就从断点接着读，不必从头重来；也可以先传 `outline: true` 拿标题大纲（级别 / 字符偏移 / 标题），直接跳到某一章。

**改了 Word 文档后格式全没了？** 别用 `office_write_docx` 覆盖原文件——它是整篇重建。改已有文档用 `office_edit_docx`：按段落序号（或整段原文）定位，只改那几段，段落样式、表格、图片、页眉页脚、批注都保留。

**怎么一次生成多行明细的合同？** `office_fill_docx_template` 支持模板循环：模板里写两个花括号加 `#明细数组名` 包住那一行，再以两个花括号加斜杠收尾，传数组进去就会按行展开（可嵌套，空数组时整块消失）；条件用 `#标记` / `^标记`。模板文件里的双花括号是 docxtemplater 语法，和 DSH 的提示词变量无关。

**怎么只把正文改成单倍行距？** `office_edit_docx` 传 `{ "op": "set_style", "scope": "body", "lineSpacingMultiple": 1 }` —— `scope: "body"` 只命中非标题段落，标题与表格文字不动；要动标题就换 `"headings"`，只动表格内文字用 `"table"`。

**怎么排出公文体例？** `office_write_docx` 传 `style`：`{"font":"仿宋_GB2312","sizePt":16,"lineSpacingPt":28.8,"firstLineIndentChars":2,"align":"both","headings":{"font":"黑体","sizePt":16}}`（三号仿宋、行距固定值 28.8 磅、首行缩进 2 字符）。改已有文档用 `office_edit_docx` 的 `set_style`，配 `scope: "all"` 一次改全篇。

**转 PDF 报错？** `.pdf` 输出依赖本机 LibreOffice 或 Microsoft Word，纯 JS 不提供 PDF 渲染；写出 `.doc` / `.odt` 同理。各格式保真度见[平台支持](docs/platform.md)。

**智能体还是用了通用 `read`？** 见上一节，补一条用户级指令即可稳定命中。

## 文档

| 文档 | 内容 |
| --- | --- |
| [用法与参数](docs/usage.md) | 全部示例、`office_edit_xlsx` 操作速查、参数与限制 |
| [平台支持](docs/platform.md) | 各格式保真度、外部转换器探测、工作方式 |
| [安全说明](docs/security.md) | 已修漏洞、边界与假设、残余风险、私密报告入口 |
| [设计说明](docs/design.md) | 为什么不用正则解析、目录结构、依赖取舍 |
| [开发与发版](docs/development.md) | 跑测试、覆盖率、SonarQube、npm 发布、发版流程 |
| [版本记录](docs/changelog.md) | 完整变更历史 |

## 版本记录

| 版本 | 变更 |
| --- | --- |
| **0.3.31** | 命中 DSH 0.1.6-alpha.1 / .2 的模块解析缺陷（xlsx 路径全失败）时不再误报「缺依赖」，改为说明这是宿主缺陷并给出 DSH 侧报告链接（新错误码 `DSH_RESOLVER_BUG`）；README 补一条已知问题 |
| **0.3.30** | 修 PDF 读取在 Windows 上必失败：内置 pdfjs 起 fake worker 用的是 `await import(workerSrc)`，Node 的 ESM 加载器只认 `file:` / `data:` / `node:`，Windows 的 `C:\…` 裸路径会抛「Only URLs with a scheme in: file, data, and node」（POSIX 绝对路径碰巧能过，所以之前只在 Mac 上验证过）。现在 `workerSrc` 一律给 `file://` URL；同时把内置 pdfjs 从 6.3.289 换成 **4.10.38**（前者 `engines: node >= 22.13` 且无 polyfill，会在 Node 20/22 上抛 `Promise.withResolvers is not a function`；后者声明 `>= 20` 并自带 polyfill），插件的 Node ≥ 20 承诺才成立 |
| **0.3.29** | `office_read` 支持 `.pdf`：只抽文本层（页数 + 文本，多页带 `--- 第 N 页 ---` 标记），`offset` / `maxChars` 分段读、`outline` 给页清单、`format: "html"` 也是文本重建；后端链为 `pdftotext`（poppler）→ macOS 自带 PDFKit → 随包携带的 pdfjs，任一可用即可读，缺工具不会读不出来。图片型（扫描件）没有文本层，会明确提示需要 OCR。随包新增 `vendor/pdfjs`（Apache-2.0，仅文本抽取，不渲染、不联网、`isEvalSupported: false`），离线包依旧零外部依赖、零安装脚本 |
| **0.3.28** | ① 新增 `office_edit_docx`：就地改已有 Word（查找替换 / 整段改写 / 改排版 / 多级编号 / 插入删除段落 / 表格样式与增删行列 / 合并单元格），只重写正文部件。② 排版：`office_write_docx` 的 `style` 设字体·字号·行距·首行缩进·对齐·段前后（标题 `headings`、表格 `table` 含按内容自动列宽与单元格内边距），html 内联样式补齐 `font-family`（含中文 `eastAsia`）/`line-height`/`text-indent`/`margin` 并支持容器继承。③ `set_style` 按角色选段落（正文 / 标题 / 表格内 / 全部 + 区间）。④ `set_numbering` 多级自动编号（标题挂样式链接 + 正文列表挂当前标题下一级，含公文体例预置）。⑤ `office_read` 读取时把 Word 自动编号展开成文字（`一、`/`（一）`/`1.1`）。⑥ `office_write_docx` / `office_convert` 支持 `<img>` 嵌入本地图片（PNG/JPEG/GIF/BMP，按正文宽等比缩放，不联网）。⑦ 表格 `set_table` 增 `repeatHeader`/`cellVerticalAlign`/`rowHeightPt`/`cantSplit` 与 `unmerge_table_cells`。⑧ `office_query` 支持多表 `join`（inner/left/right/full）与 `pivot` 透视表（含合计）。⑨ `office_edit_xlsx` 增 `conditional_format`（9 种规则）与 `data_validation`（下拉 / 区间 / 公式）。⑩ `office_read` 可按 `formulas` 切换公式本体与计算值；修 `.xlsb` / `.ods` 的家族判定（此前分别被当成损坏文件与 Word 文档） |
| **0.3.27** | Word 长文档可分段读：`offset` 续读（不再每次从头开始）+ `outline` 标题大纲跳读；逐段拼回与全文逐字节一致 |
| **0.3.26** | 修 `office_fill_docx_template` 描述里的字面量占位符导致 DSH 提示词组装报 `malformed prompt variable reference`：面向模型的文本一律不再出现双花括号变量写法，并加测试门禁 |
| **0.3.25** | 新增 `office_query`：在表格文件内筛选 / 分组 / 求和 / 计数 / 去重 / 排序，几万行只回结论；不给条件时返回表结构画像。数值与日期按真实写法识别，算不了的需求（跨文件 join、透视、建模）照常写脚本 |
| **0.3.24** | 安装改用包名 `dsh-office-toolkit`（不再依赖直链），其它方式折叠收起 |
| **0.3.23** | 发布自动化:推 `v*` tag 由 GitHub Actions 发布 npm 并上传 Release 附件 |
| **0.3.22** | 文档精简：更新说明只留结论与命令 |
| **0.3.21** | 文档：更新请用带版本号的直链（`latest` 直链可能被 pnpm 复用缓存） |
| **0.3.20** | 主文档部件按包关系解析（非规范路径也能读）；启动日志与报错带上插件版本 |
| **0.3.19** | 部件名大小写不同也能读（连同引用一起归一化）；mammoth 认不出时重打包重试、再不行用内置解析器兜底 |
| **0.3.18** | 修 0.3.17 回归：读取非标准文件时报 `Cannot add property containerNote`（宿主参数是冻结的）；改为不改写调用方参数 |
| **0.3.17** | 读取兼容性：非标准 zip（条目名含反斜杠）自动修正、按真实内容识别改过后缀的文件；修正默认样式解析 bug |
| **0.3.16** | `.docx` 可读出排版格式（字体 / 字号 / 行距 / 缩进 / 对齐 / 页边距），`withFormatting: true` |
| **0.3.15** | 依赖装不全时点名缺哪个包并给出修复命令；二进制输出自检；离线包完整性门禁 |
| **0.3.14** | 文档重组：README 收敛为入口，细节移入 `docs/` |

更早版本见[版本记录](docs/changelog.md)。

## 贡献

- 问题与建议走 [GitHub Issues](https://github.com/cnkids/dsh-office-toolkit/issues)；**安全问题请用[私密漏洞报告](https://github.com/cnkids/dsh-office-toolkit/security/advisories/new)**，不要发在公开 Issue 里。
- 欢迎 PR：改完请跑 `npm test`，并保证 SonarQube 无新增问题（见[开发与发版](docs/development.md)）。
- 提交信息格式：`v<版本> <类型>: <说明>`；每次版本变更都要同步本文档与[版本记录](docs/changelog.md)。

## 许可证

[MIT](LICENSE)
