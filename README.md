# dsh-office-tools

DeepSeek Harness(DSH)**宿主插件**:给 AI 智能体增加读写 Word / Excel 的工具。**跨平台**(macOS / Windows / Linux),核心功能纯 JavaScript 实现,不需要安装 Office 或 LibreOffice。

> 版本 0.2.1 — 修复**同一工作表写多个图表只有第一个可见**的问题(OOXML 规定一个工作表只能有一个 drawing 部件,现已改为多图表共用一个 drawing);同时做了一轮代码质量治理(0 缺陷 / 0 代码异味,可靠性 · 安全性 · 可维护性均 A 级)。

<details>
<summary>历史版本</summary>

- **0.2.0** — 新增 Windows 支持:`.doc`/`.rtf`/`.odt` 读取改为纯 JS 优先,外部转换器按平台自动择优,并修正 Windows 路径大小写不敏感判定。

</details>

## 能力一览

| 工具 | 作用 | 支持格式 |
| --- | --- | --- |
| `office_read` | 读取文档/表格内容 | `.docx` `.doc` `.rtf` `.odt` `.xlsx` `.xls` `.xlsb` `.ods` `.csv` `.tsv` |
| `office_write_docx` | 新建 Word 文档(HTML / Markdown / 纯文本) | `.docx`(推荐)`.doc` `.rtf` `.odt` |
| `office_write_xlsx` | 新建工作簿(多工作表、表头样式、公式、日期) | `.xlsx` |
| `office_edit_xlsx` | 修改已有工作簿:单元格/样式/合并/增删行列/冻结/筛选/图片/图表 | `.xlsx` |
| `office_fill_docx_template` | `{{变量}}` 模板填充(合同、通知、批量套打) | `.docx` |
| `office_convert` | 格式互转 | Word 家族与表格家族内互转,含 `.xls` 导出与 `.pdf` 输出* |

\* `.pdf` 需要本机有 LibreOffice 或 Microsoft Word。

## 平台支持

分两层,纯 JS 层保证任何平台开箱可用,外部转换器只用于提升旧格式保真度:

| 能力 | 纯 JS(全平台) | 外部转换器 |
| --- | --- | --- |
| `.docx` 读 / 写 | ✅ mammoth / html-to-docx | — |
| `.xlsx` 读 / 写 / 编辑 / 图表 | ✅ exceljs + 自研 OOXML 注入 | — |
| `.xls` `.xlsb` `.ods` `.csv` 读 / 写 / 转换 | ✅ SheetJS(含 `.xls` 导出) | — |
| `.doc` 读 | ✅ word-extractor | macOS `textutil` / LibreOffice / Word COM |
| `.rtf` `.odt` 读 | ✅ 内置解析器 | 同上(更高保真) |
| `.doc` `.rtf` `.odt` 写 | ⚠ 仅 `.rtf` 可纯 JS 生成 | 需要 textutil / LibreOffice / Word |
| 转 `.pdf` | ❌ | 需要 LibreOffice 或 Word |

外部转换器自动探测,按平台择优:

| 平台 | 首选 | 备注 |
| --- | --- | --- |
| macOS | `/usr/bin/textutil`(系统自带) | 零安装,支持 doc/docx/rtf/odt/html/txt |
| Windows | LibreOffice `soffice.exe` | 装 Word 时自动改用 Word COM(PowerShell 调用) |
| Linux | LibreOffice `soffice` | 也可用 `libreoffice` 命令 |

Windows 上**不装任何 Office 软件**也能:读 `.doc/.rtf/.odt`、读写 `.docx/.xlsx/.xls/.csv`、编辑表格、加图表、模板填充。只有**写出旧格式或转 PDF** 才需要 LibreOffice / Word。

## 安装

插件包位于工作区 `dsh-office-tools/`,用 DSH 自带命令安装到 web profile:

```sh
dsh plugin --profile web add file:/Users/cnkids/Project/ElectronProj/dsh-office-tools
```

该命令会在 `~/.dsh/profiles/web` 里执行 `pnpm add`,并把 `dsh-office-tools` 自动追加到 `dsh.profile.bundles`。

**必须重启 `dsh web`(或重开 DeepSeek Harness 桌面端),然后新建会话**,新工具才会出现在智能体的工具列表里 —— `bundles` 只在启动时读取。

等价的纯手工步骤:

```sh
cd ~/.dsh/profiles/web
pnpm add file:/Users/cnkids/Project/ElectronProj/dsh-office-tools
# 再把 "dsh-office-tools" 追加到 package.json 的 dsh.profile.bundles 数组
```

验证(不会占用端口,可与运行中的实例并存):

```sh
dsh --profile web --dump-config | grep -A1 office
```

应能看到插件行 `id: dsh-office-tools / name: dsh-office-tools`。

卸载:

```sh
dsh plugin --profile web remove dsh-office-tools
```

## 智能体用法示例

读取文档:

```json
{ "path": "报告.docx" }
```

读取某个工作表的前 100 行:

```json
{ "path": "销售.xlsx", "sheets": ["明细"], "range": "A1:F100" }
```

新建 Word 文档:

```json
{ "path": "周报.docx", "markdown": "# 周报\n\n## 进展\n- 完成 A\n- 完成 B\n\n| 指标 | 值 |\n| --- | --- |\n| 覆盖率 | 92% |\n" }
```

新建工作簿:

```json
{
  "path": "台账.xlsx",
  "sheets": [
    { "name": "明细", "header": true, "columnWidths": [14, 10, 12],
      "rows": [["产品", "数量", "金额"], ["键盘", 10, "=B2*199"], ["鼠标", 20, "=B3*89"]] }
  ]
}
```

修改工作簿并加图表:

```json
{
  "path": "台账.xlsx",
  "ops": [
    { "op": "set_value", "sheet": "明细", "ref": "E1", "value": "合计", "style": { "bold": true } },
    { "op": "style_range", "sheet": "明细", "range": "A1:C1", "style": { "fill": "DDEBF7", "align": "center" } },
    { "op": "add_chart", "sheet": "明细", "chartType": "column", "title": "金额对比",
      "categories": "A2:A3", "series": [{ "range": "C2:C3" }], "anchor": "G2" }
  ]
}
```

模板填充:

```json
{ "templatePath": "合同模板.docx", "outputPath": "合同-001.docx",
  "data": { "合同编号": "HT-2026-001", "甲方": "某某公司", "金额": "12800" } }
```

模板里写 `{{合同编号}}` 即可;变量名可以是中文。

转换格式(含旧格式与 PDF):

```json
{ "sourcePath": "旧表.xls", "outputPath": "新表.xlsx" }
{ "sourcePath": "报告.docx", "outputPath": "报告.pdf" }
```

## 参数要点

- **路径**:绝对路径,或相对当前会话工作区的路径;Windows 下盘符大小写不敏感。
- **写入范围**:默认仅允许写入会话工作区与系统临时目录(与 DSH 沙箱一致);越界会返回 `FS_SANDBOX_DENIED` 提示。
- **单元格值**:数字/布尔原样;字符串以 `=` 开头视为公式;`date:2026-09-09` 写入日期;`num:1,234.5` 强制数字。
- **样式对象**:`{bold, italic, fontSize, color, fill, align, valign, wrap, numFmt, border}`(颜色 `RRGGBB`)。
- **大表读取**:默认 400 行 × 60 列,超限会提示用 `range`/`maxRows` 分段读。
- **内容上限**:Word 正文默认返回 90000 字符,超出截断并提示。

## 已知限制

- 图表支持 `bar` / `column` / `line` / `pie`,通过 OOXML 注入实现,打开时由 Excel/WPS 自动计算数据(不写缓存值)。复杂图表(组合图、双轴、趋势线)请用 Excel/WPS 手动调整。
- 一个工作表最多只有一个 drawing 部件(OOXML 规定),因此**同一工作表的多个图表共用一个 drawing**,各自是独立锚点,可分别拖动/调整;`add_image` 与 `add_chart` 混用同一工作表也走这条路径。
- `.doc` 纯 JS 解析(word-extractor)能拿到正文/页眉/脚注文字,**不保留表格与版式**;需要高保真时请装 LibreOffice 或 Word。
- 写出 `.doc` / `.odt` 必须有外部转换器;`.rtf` 可用纯 JS 生成(仅段落文本)。
- `.pdf` 输出需要 LibreOffice 或 Microsoft Word。
- `office_fill_docx_template` 只做 `{{变量}}` 替换,不支持循环 / 条件(需要时多次调用或改用脚本)。
- 文档内嵌图片:读取时会以 Markdown 图片语法提示,写入 `.docx` 暂不嵌入本地图片。
- `office_convert` 源与目标扩展名相同时不做转换(路径不同时会原样复制;路径相同则直接返回)。

## 开发与测试

```sh
cd dsh-office-tools
node test/selftest.mjs        # 核心库 20 项:docx/xlsx/xls/doc/模板/图表/转换端到端
node test/converters.test.mjs # 跨平台层 13 项:RTF/ODT/word-extractor/路径围栏/后端探测
node test/plugin-smoke.mjs    # 插件适配层 31 项:注册、schema、6 个工具调用、沙箱拒绝、报错
```

共 **64 项测试**,全部为纯 Node 脚本,不需要测试框架。其中「同表多图表」用例会校验三个图表落在同一个 drawing 部件里、工作表只引用它一次。

修改代码后的生效方式:

- profile 里当前是 `link:`(符号链接)安装,源码改动**不需要重新安装**,但宿主插件只在启动时加载,所以**改完都要重启 `dsh web`**(或重开桌面端)并新建会话才生效:

  ```sh
  dsh plugin --profile web add link:/Users/cnkids/Project/ElectronProj/dsh-office-tools
  ```

- 若改用 `file:` 安装,pnpm 会**复制**一份到 profile,改完必须重新执行一次安装命令刷新副本:

  ```sh
  dsh plugin --profile web add file:/Users/cnkids/Project/ElectronProj/dsh-office-tools
  ```

- HMR 只监听 `cordis.patch.yml` 的行变化,不监听插件源码。

## 代码质量

SonarQube 项目 `dsh-office-toll`(https://so.rclandy.com),扫描范围 `lib/` + `test/`:

```sh
export SONAR_TOKEN=<你的 token>
sonar-scanner        # 读取仓库根目录的 sonar-project.properties
```

`sonar-project.properties` 按项目要求不入库(`.gitignore` 已排除)。最近一次分析:0 Bug / 0 漏洞 / 0 代码异味,可靠性 · 安全性 · 可维护性均 A 级。

## 目录结构

```
dsh-office-tools/
├── package.json          # dsh.bundle.patch 指向 cordis.patch.yml
├── cordis.patch.yml      # 把插件行插入 profile 插件列表
├── lib/
│   ├── index.js          # 宿主适配层:工具注册、路径解析、沙箱围栏、fs/observed 事件
│   └── core/
│       ├── office.js     # 六个操作(读/写/编辑/模板/转换)编排
│       ├── word.js       # docx 读(mammoth)/写(html-to-docx)/模板(docxtemplater)
│       ├── excel.js      # xlsx 读/建/编辑(exceljs)
│       ├── charts.js     # OOXML 图表注入(bar/line/pie)
│       ├── legacy.js     # 旧表格格式(SheetJS:xls/ods/csv)
│       ├── converters.js # 旧 Word 格式公共 API(读/写/探测)
│       ├── legacy-read.js     # 纯 JS 解析:doc(word-extractor)/rtf/odt
│       ├── legacy-external.js # 外部转换器:textutil / LibreOffice / Word COM
│       ├── path-guard.js # 跨平台路径围栏(Windows 大小写不敏感)
│       ├── md.js         # Markdown↔HTML 与 HTML→Markdown 转换
│       └── util.js       # 错误类型、体积上限、截断工具
└── test/                 # 三个测试脚本与样例输出
```
