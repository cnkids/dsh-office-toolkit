# dsh-office-toolkit

> DSH(DeepSeek Harness)宿主插件:让智能体自己读写 Word / Excel —— 不装 Office 也能跑。

装上这个插件,智能体就能直接读报告、写台账、改表格、套合同模板、转换旧格式,不用你手动搬文件。核心能力是纯 JavaScript(macOS / Windows / Linux 行为一致),只有写出 `.doc` / `.odt` 和转 `.pdf` 才可选地借用本机 LibreOffice 或 Word。

**安装**:`dsh plugin --profile web add github:cnkids/dsh-office-toolkit`,重启 `dsh web` 后新建会话。

> **0.3.8** 补充「更新」说明:重跑同一条 `add` 即可,三种安装来源均已实测。**0.3.7** 强化工具描述,让智能体优先用本插件读写 Office 文件。**0.3.6** 把 `xlsx` 换成无已知 high 漏洞的 `@e965/xlsx@0.20.3`。**0.3.5** 补齐 npm 元数据(repository / homepage / bugs)并把发布源固定为官方 registry。**0.3.4** 新增永不过期的安装直链 `releases/latest/download/dsh-office-toolkit.tgz`。**0.3.3** 补充无 git / 完全离线机器的安装方式。**0.3.2** 精简文档。**0.3.1** 仓库名与包名统一为 `dsh-office-toolkit`。**0.3.0** 安全热点清零(见「设计说明」)。**0.2.x** 修复同表多图表丢失、新增 Windows 支持。

## 工具

| 工具 | 作用 | 格式 |
| --- | --- | --- |
| `office_read` | 读取内容(Excel 以 TSV 返回) | `.docx .doc .rtf .odt .xlsx .xls .xlsb .ods .csv .tsv` |
| `office_write_docx` | 新建文档(html / markdown / text) | `.docx`(推荐)`.doc .rtf .odt` |
| `office_write_xlsx` | 新建工作簿(多表、表头、公式、日期) | `.xlsx` |
| `office_edit_xlsx` | 改已有工作簿:单元格 / 样式 / 合并 / 行列 / 冻结 / 筛选 / 图片 / 图表 | `.xlsx` |
| `office_fill_docx_template` | `{{变量}}` 模板填充(合同、通知、批量套打) | `.docx` |
| `office_convert` | 格式互转(含 `.xls` 导出、`.pdf` 输出) | Word 家族 / 表格家族内部 |

## 平台支持

纯 JS 层保证开箱可用,外部转换器只用于提升旧格式保真度。

| 能力 | 纯 JS | 外部转换器 |
| --- | --- | --- |
| docx / xlsx 读写、编辑、图表 | ✅ | — |
| xls / xlsb / ods / csv 读、写、转换 | ✅ SheetJS | — |
| doc / rtf / odt 读取 | ✅ 内置解析器 | 更高保真 |
| doc / rtf / odt 写出 | 仅 rtf | 必需 |
| 转 pdf | ❌ | 必需 |

外部转换器自动探测:macOS 用系统自带 `textutil`;Windows / Linux 用 LibreOffice `soffice`,Windows 装了 Word 时改用 Word COM。**Windows 上不装任何 Office 也能读写 docx/xlsx/xls/csv、编辑表格、加图表、套模板。**

## 安装

包名与仓库名均为 `dsh-office-toolkit`。**没装 git 就用第一条**,它不需要 git、不需要手动下载,也永远指向最新版。

| 场景 | 命令 |
| --- | --- |
| **无 git(推荐)** | `dsh plugin --profile web add https://github.com/cnkids/dsh-office-toolkit/releases/latest/download/dsh-office-toolkit.tgz` |
| 有 git · 跟随 main | `dsh plugin --profile web add github:cnkids/dsh-office-toolkit` |
| 有 git · 锁定版本 | `dsh plugin --profile web add github:cnkids/dsh-office-toolkit#v0.3.4` |
| 完全离线 | 解压 `releases/latest/download/dsh-office-toolkit-offline.zip` 后 `dsh plugin --profile web add link:C:/dsh-office-toolkit` |
| 本地源码 | `dsh plugin --profile web add link:/path/to/dsh-office-toolkit` |

`releases/latest/download/...` 是 GitHub 的固定别名,每个 Release 都会同时上传**带版本号**和**不带版本号**两份附件,所以这条命令不用随版本改。`.tgz` 只有几十 KB;离线包内含完整 `node_modules`,解压后 `link:` 安装零下载。

`dsh plugin` 会在 profile 目录执行 `pnpm add`,并自动把包名追加进 `dsh.profile.bundles`(本包声明了 `dsh.bundle`,无需手改)。装完**必须重启 `dsh web`(或重开桌面端)并新建会话**才会加载。

```sh
dsh --profile web --dump-config | grep -A1 dsh-office-toolkit   # 校验
dsh plugin --profile web remove dsh-office-toolkit              # 卸载
```

### 更新

**重跑当初那条 `add` 就是更新** —— pnpm 会重新解析并拉取,`file:` / 固定直链 / `github:` 三种来源都已实测(`file:` 同名文件内容变了会重装;固定直链的 URL 内容变了也会重装;`github:` 分支前进后会解析到新提交)。也可一次更新 profile 内全部依赖:

```sh
dsh plugin --profile web add <当初用的 spec>   # 重新安装 = 更新
dsh plugin --profile web update                # 更新 profile 内全部依赖
```

- **`link:` 安装不需要 pnpm**:源码目录 `git pull` 即最新。
- 装了**带版本号**的 spec(`#v0.3.7` 或 `.../releases/download/v0.3.7/...tgz`)不会自动前进,更新时要换版本号 —— 这也是推荐用**不带版本号的固定直链**的原因。
- **更新后必须重启 `dsh web`(或重开桌面端)并新建会话**,否则看起来像没更新:工具列表与描述只在启动时读取。

目标机器需要 Node ≥ 18 和 pnpm;`.tgz` 不含依赖,仍需能访问 npm registry(内网先配镜像)。裸包名 `add dsh-office-toolkit` 需先发布到 npm(见「开发 → 发布到 npm」,暂未发布)。

## 让智能体优先使用本插件

DSH 的工具注册**没有优先级设置**,模型只依据每个工具的 `description` 选择;而内置 `read` / `write` / `edit` 只处理 UTF-8 文本,读 `.docx` 会直接返回 `Error: binary file`,且不会提示该换哪个工具。本插件已把「哪类文件必须用本工具」写在每个描述最前面,并把 `Word / Excel / 文档 / 表格 / .docx / .xlsx` 等词作为触发提示,通常一次就能选对。

想更保险,在**用户级指令文件** `~/.dsh/AGENTS.md` 里加一段(该文件对所有项目、所有会话生效):

```markdown
## 文件读写
- 读取 .docx/.doc/.rtf/.odt/.xlsx/.xls/.xlsb/.ods/.csv/.tsv 一律用 `office_read`。
- 创建/修改用 `office_write_docx` / `office_write_xlsx` / `office_edit_xlsx` / `office_fill_docx_template`,格式转换用 `office_convert`。
- 不要用通用 `read`/`write`/`edit` 处理这些文件 —— 它们只能处理 UTF-8 文本。
```

改完重启 `dsh web` 并新建会话生效。

## 用法示例

读取(Excel 可用 `sheets` / `range` 限定窗口):

```json
{ "path": "报告.docx" }
{ "path": "销售.xlsx", "sheets": ["明细"], "range": "A1:F100" }
```

新建:

```json
{ "path": "周报.docx", "markdown": "# 周报\n\n- 完成 A\n- 完成 B\n" }
{ "path": "台账.xlsx", "sheets": [
  { "name": "明细", "header": true, "columnWidths": [14, 10, 12],
    "rows": [["产品", "数量", "金额"], ["键盘", 10, "=B2*199"], ["鼠标", 20, "=B3*89"]] } ] }
```

改工作簿(按顺序执行 `ops`):

```json
{ "path": "台账.xlsx", "ops": [
  { "op": "set_value", "sheet": "明细", "ref": "E1", "value": "合计", "style": { "bold": true } },
  { "op": "style_range", "sheet": "明细", "range": "A1:C1", "style": { "fill": "DDEBF7", "align": "center" } },
  { "op": "add_chart", "sheet": "明细", "chartType": "column", "title": "金额对比",
    "categories": "A2:A3", "series": [{ "range": "C2:C3" }], "anchor": "G2" } ] }
```

模板填充(模板里写 `{{合同编号}}`,变量名支持中文):

```json
{ "templatePath": "合同模板.docx", "outputPath": "合同-001.docx",
  "data": { "合同编号": "HT-2026-001", "甲方": "某某公司", "金额": "12800" } }
```

转换:

```json
{ "sourcePath": "旧表.xls", "outputPath": "新表.xlsx" }
{ "sourcePath": "报告.docx", "outputPath": "报告.pdf" }
```

## 参数与限制

**参数**

- 路径:绝对路径或相对会话工作区;写入默认只允许会话工作区与系统临时目录,越界返回 `FS_SANDBOX_DENIED`。
- 单元格值:数字 / 布尔原样;`=` 开头视为公式;`date:2026-09-09` 写入日期;`num:1,234.5` 强制数字。
- 样式:`{bold, italic, fontSize, color, fill, align, valign, wrap, numFmt, border}`(颜色 `RRGGBB`)。
- 大表默认 400 行 × 60 列、Word 正文默认 90000 字符,超出会截断并提示用参数分段读。

**限制**

- 图表支持 `bar` / `column` / `line` / `pie`,以 OOXML 注入实现,打开时由 Excel/WPS 计算数据;组合图、双轴等复杂图表请手动调整。
- 一个工作表只能有一个 drawing 部件,因此同表多图表共用一个 drawing(各自独立锚点,可分别拖动)。
- `.doc` 纯 JS 解析不保留表格与版式;写出 `.doc` / `.odt`、转 `.pdf` 需要 LibreOffice 或 Word。
- 模板填充只做 `{{变量}}` 替换,不支持循环 / 条件(批量套打多次调用即可)。
- 文档内嵌图片只读不写。
- `office_convert` 源与目标扩展名相同时:路径不同则原样复制,路径相同则直接返回。

## 开发

```sh
node test/selftest.mjs        # 核心库 21 项:读写编辑、模板、图表、转换端到端
node test/converters.test.mjs # 跨平台层 15 项:RTF/ODT/word-extractor/路径围栏/回溯安全
node test/plugin-smoke.mjs    # 插件适配层 31 项:注册、schema、6 个工具、沙箱拒绝、报错
```

共 67 项,纯 Node 脚本,不需要测试框架。profile 当前是 `link:` 安装,改完源码重启 `dsh web` 即生效;若用 `file:` 安装需重新执行一次 `add` 刷新副本(HMR 不监听插件源码)。

代码质量走 SonarQube(项目 `dsh-office-toll`;`sonar-project.properties` 不入库):

```sh
export SONAR_TOKEN=<token> && sonar-scanner
```

当前 0 缺陷 / 0 漏洞 / 0 代码异味 / 0 安全热点,可靠性 · 安全性 · 可维护性均 A 级。

### 发布到 npm(维护者)

```sh
npm login --registry https://registry.npmjs.org/   # 首次,需要 npm 账号 + 2FA
npm publish
```

`package.json` 的 `publishConfig.registry` 已固定为官方源,所以 `npm publish` **不受 `~/.npmrc` 里国内镜像的影响**(镜像只能读不能发,不加这行会报 `ENEEDAWAUTH` / 403)。发布后裸包名 `dsh plugin --profile web add dsh-office-toolkit` 即可用。

## 设计说明

解析 OOXML / HTML / ODT 时没有用 `/<[^>]+>/g`、`/([A-Za-z]+\d+)/` 这类正则:它们带无界量词且分支可重叠,在构造输入上会让回溯引擎反复重扫同一段文本,耗时随输入超线性增长(ReDoS,SonarQube `S5852`)。现在统一走 `lib/core/markup.js` 的一次前向扫描,每个字符只被访问有限次,**复杂度是输入长度的 O(n) 上界,与输入内容无关**。`test/converters.test.mjs` 覆盖 20 万字符级畸形输入,要求毫秒级返回或快速报错。

## 依赖与已知告警

- **`xlsx` 用的是 `@e965/xlsx@0.20.3`**:npm 上的 `xlsx` 停在 `0.18.5`,带 Prototype Pollution(`GHSA-4r6h-8v6p-xvw6`)与 ReDoS(`GHSA-5pgg-2g8v-p4x9`)两个 high;SheetJS 早已停止在 npm 发布,修复版只在其官方 CDN。但官方 CDN 的 URL 形式依赖会被 pnpm 的 `blockExoticSubdeps` 判为 exotic 子依赖而拒绝安装,所以改用 npm 上该官方构建的自动转发包(月下载 300 万+,仓库 [sheetjs-npm-publisher](https://github.com/e965/sheetjs-npm-publisher))。包名不同,故代码里 `import('@e965/xlsx')`。
- **其余 `npm audit` 告警无法修复**:`image-size`(经 html-to-docx;受影响 `<=2.0.2`,而 npm 最新就是 2.0.2,上游暂无修复版)、`uuid`(经 exceljs;漏洞路径是 v3/v5/v6 带 `buf` 参数,exceljs 只用 v4)。
- **pnpm 提示的 deprecated 子依赖**(`fstream` `glob` `inflight` `lodash.isequal` `rimraf`,有时还有 `uuid`)全部来自 `exceljs@4.4.0` —— npm 上的最新稳定版(2023-10)。逐个核对公告:`inflight` / `rimraf` / `lodash.isequal` 零公告;`glob@7.2.3` 不在其公告范围(公告针对 CLI 10.2–10.4 / 11.0);`fstream@1.0.12` 本身即修复版;`uuid` 的漏洞路径(见上)不可达。**exceljs 没有更新版本,插件侧也无法用 overrides 干预**(pnpm/npm 的 overrides 只在根项目生效,DSH 的根是 profile 目录),所以这些告警无法消除,也不影响安全。
- 想让 pnpm 不再打印这些告警,安装时加 `--loglevel=error`,或在 `~/.dsh/profiles/web/.npmrc` 写一行 `loglevel=error`(会连进度输出一起隐藏)。

## 目录结构

```
lib/index.js                 宿主适配层:工具注册、路径解析、沙箱围栏、fs/observed 事件
lib/core/office.js           六个操作的编排
lib/core/word.js             docx 读 / 写 / 模板(mammoth / html-to-docx / docxtemplater)
lib/core/excel.js            xlsx 读 / 建 / 编辑(exceljs)
lib/core/charts.js           OOXML 图表注入
lib/core/legacy.js           旧表格格式(SheetJS)
lib/core/legacy-read.js      纯 JS 解析 doc / rtf / odt
lib/core/legacy-external.js  外部转换器(textutil / LibreOffice / Word COM)
lib/core/converters.js       旧 Word 格式公共 API
lib/core/markup.js           HTML / XML 线性扫描器
lib/core/md.js               Markdown ↔ HTML
lib/core/path-guard.js       跨平台路径围栏
lib/core/util.js             错误类型、体积上限、截断
test/                        三个测试脚本
```
