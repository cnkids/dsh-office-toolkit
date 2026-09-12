<h1 align="center">dsh-office-toolkit</h1>

<p align="center">
  <strong>让 DSH 智能体自己读写 Word / Excel</strong><br>
  读报告、写台账、改表格、套合同模板、转换旧格式 —— 不装 Office 也能跑。
</p>

<p align="center"><sub>DSH(DeepSeek Harness)宿主插件 · 纯 JavaScript · macOS / Windows / Linux 行为一致</sub></p>

<p align="center">
  <a href="https://github.com/cnkids/dsh-office-toolkit/releases/latest"><img src="https://img.shields.io/github/v/release/cnkids/dsh-office-toolkit?style=flat&amp;label=release&amp;color=4D6BFE" alt="Latest release"></a>
  <a href="https://github.com/cnkids/dsh-office-toolkit/releases"><img src="https://img.shields.io/github/downloads/cnkids/dsh-office-toolkit/total?style=flat&amp;label=downloads&amp;color=4D6BFE" alt="Total downloads"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-2EA44F?style=flat" alt="MIT License"></a>
  <img src="https://img.shields.io/badge/node-%E2%89%A518-339933?style=flat" alt="Node.js 18 or newer">
  <img src="https://img.shields.io/badge/macOS%20%7C%20Windows%20%7C%20Linux-4493F8?style=flat" alt="Supported platforms: macOS, Windows and Linux">
  <img src="https://img.shields.io/badge/dsh-plugin-4D6BFE?style=flat" alt="DSH plugin">
</p>

<p align="center">
  <a href="#安装"><strong>安装</strong></a>
  ·
  <a href="#工具">工具</a>
  ·
  <a href="#用法示例">用法示例</a>
  ·
  <a href="#常见问题">常见问题</a>
  ·
  <a href="https://github.com/cnkids/dsh-office-toolkit/releases/latest">Releases</a>
</p>

---

给 DSH 装上这个宿主插件,智能体就能直接读文档、写台账、改表格、套合同模板、转换旧格式,不用再手动搬文件。核心能力全部由纯 JavaScript 实现,**不需要安装 Office 或 LibreOffice**;只有写出 `.doc` / `.odt` 和转 `.pdf` 时才可选地借用本机转换器。

插件只做文档读写这一件事:不访问网络、不常驻后台、不修改 DSH 上游代码;所有写入都要先通过 DSH 的沙箱围栏。

## 工具

注册 6 个工具,由智能体按需调用:

| 工具 | 作用 | 格式 |
| --- | --- | --- |
| `office_read` | 读取内容(Excel 以 TSV 返回) | `.docx .doc .rtf .odt .xlsx .xls .xlsb .ods .csv .tsv` |
| `office_write_docx` | 新建文档(html / markdown / text) | `.docx`(推荐)`.doc .rtf .odt` |
| `office_write_xlsx` | 新建工作簿(多表、表头、公式、日期) | `.xlsx` |
| `office_edit_xlsx` | 改已有工作簿:单元格 / 样式 / 合并 / 行列 / 冻结 / 筛选 / 图片 / 图表 | `.xlsx` |
| `office_fill_docx_template` | `{{变量}}` 模板填充(合同、通知、批量套打) | `.docx` |
| `office_convert` | 格式互转(含 `.xls` 导出、`.pdf` 输出) | Word 家族 / 表格家族内部 |

## 平台支持

纯 JS 层保证开箱可用,外部转换器只用于提升旧格式保真度。需要 **Node ≥ 20**。

| 能力 | 纯 JS | 外部转换器 |
| --- | --- | --- |
| docx / xlsx 读写、编辑、图表 | ✅ | — |
| xls / xlsb / ods / csv 读、写、转换 | ✅ SheetJS | — |
| doc / rtf / odt 读取 | ✅ 内置解析器 | 更高保真 |
| doc / rtf / odt 写出 | 仅 rtf | 必需 |
| 转 pdf | ❌ | 必需 |

外部转换器按平台自动探测:macOS 用系统自带 `textutil`;Windows / Linux 用 LibreOffice `soffice`,Windows 装了 Word 时自动改用 Word COM。**Windows 上不装任何 Office 也能读写 docx/xlsx/xls/csv、编辑表格、加图表、套模板。**

## 安装

包名与仓库名均为 `dsh-office-toolkit`。**没装 git 就用第一条** —— 它不需要 git、不需要手动下载,也永远指向最新版。

| 场景 | 命令 |
| --- | --- |
| **无 git(推荐)** | `dsh plugin --profile web add https://github.com/cnkids/dsh-office-toolkit/releases/latest/download/dsh-office-toolkit.tgz` |
| 有 git · 跟随 main | `dsh plugin --profile web add github:cnkids/dsh-office-toolkit` |
| 有 git · 锁定版本 | `dsh plugin --profile web add github:cnkids/dsh-office-toolkit#v0.3.11` |
| 完全离线 | 解压 `releases/latest/download/dsh-office-toolkit-offline.zip` 后 `dsh plugin --profile web add link:C:/dsh-office-toolkit` |
| 本地源码 | `dsh plugin --profile web add link:/path/to/dsh-office-toolkit` |

`releases/latest/download/...` 是 GitHub 的固定别名,每个 Release 都会同时上传**带版本号**和**不带版本号**两份附件,所以这条命令不用随版本改。`.tgz` 只有几十 KB;离线包内含完整 `node_modules`,解压后 `link:` 安装零下载。

`dsh plugin` 会在 profile 目录执行 `pnpm add`,并自动把包名追加进 `dsh.profile.bundles`(本包声明了 `dsh.bundle`,无需手改)。装完**必须重启 `dsh web`(或重开桌面端)并新建会话**才会加载。

```sh
dsh --profile web --dump-config | grep -A1 dsh-office-toolkit   # 校验
dsh plugin --profile web remove dsh-office-toolkit              # 卸载
```

目标机器需要 Node ≥ 18 和 pnpm;`.tgz` 不含依赖,仍需能访问 npm registry(内网先配镜像)。裸包名 `add dsh-office-toolkit` 需先发布到 npm。

### 更新

**重跑当初那条 `add` 就是更新** —— pnpm 会重新解析并拉取:

| 安装方式 | 更新行为 |
| --- | --- |
| 固定直链(推荐) | URL 内容变更后重跑 `add` 即装新版 |
| 本地 `.tgz`(`file:`) | 同名文件内容变更后重跑 `add` 即重装 |
| `github:` / git 分支 | 重跑 `add`,或 `dsh plugin --profile web update` |

一次更新 profile 内全部依赖:`dsh plugin --profile web update`。

- **`link:` 安装不需要 pnpm**:源码目录 `git pull` 即最新。
- 装了**带版本号**的 spec(`#v0.3.11` 或 `.../releases/download/v0.3.11/...tgz`)不会自动前进,更新时要换版本号 —— 这也是推荐用固定直链的原因。
- **更新后必须重启 `dsh web` 并新建会话**,否则看起来像没更新:工具列表与描述只在启动时读取。

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

## 让智能体优先使用本插件

DSH 的工具注册**没有优先级设置**,模型只依据每个工具的 `description` 选择;而内置 `read` / `write` / `edit` 只处理 UTF-8 文本,读 `.docx` 会直接返回 `Error: binary file`,且不会提示该换哪个工具。本插件已把「哪类文件必须用本工具」写在每个描述最前面,并把 `Word / Excel / 文档 / 表格 / .docx / .xlsx` 等词作为触发提示,通常一次就能选对。

想更保险,在**用户级指令文件** `~/.dsh/AGENTS.md` 里加一段(该文件对所有项目、所有会话生效):

```markdown
## 文件读写
- 读取 .docx/.doc/.rtf/.odt/.xlsx/.xls/.xlsb/.ods/.csv/.tsv 一律用 `office_read`。
- 创建/修改用 `office_write_docx` / `office_write_xlsx` / `office_edit_xlsx` / `office_fill_docx_template`,格式转换用 `office_convert`。
- 不要用通用 `read`/`write`/`edit` 处理这些文件 —— 它们只能处理 UTF-8 文本。
```

指令文件是热加载的,保存即生效;插件本身的改动仍需重启 `dsh web`。

## 工作方式

```text
用户 / 智能体
   │
   ├─ office_read · office_write_docx · office_write_xlsx
   │  office_edit_xlsx · office_fill_docx_template · office_convert
   │
   ├─ docx / xlsx ─────────── 纯 JS:mammoth · @turbodocx/html-to-docx · @wekanteam/exceljs · 自研 OOXML 图表注入
   ├─ xls / xlsb / ods / csv ─ 纯 JS:SheetJS(@e965/xlsx 0.20.3)
   └─ doc / rtf / odt ──────── 内置线性解析器优先
                                 └─ 不可用时 → 本机转换器
                                    macOS textutil / LibreOffice / Word COM
   │
   └─ 写入前经过 DSH 沙箱围栏:仅会话工作区与系统临时目录
```

## 参数与限制

**参数**

- 路径:绝对路径或相对会话工作区;写入默认只允许会话工作区与系统临时目录,越界返回 `FS_SANDBOX_DENIED`。
- 单元格值:数字 / 布尔原样;`=` 开头视为公式;`date:2026-09-09` 写入日期;`num:1,234.5` 强制数字。
- 样式:`{bold, italic, fontSize, color, fill, align, valign, wrap, numFmt, border}`(颜色 `RRGGBB`)。
- 大表默认 400 行 × 60 列、Word 正文默认 90000 字符,超出会截断并提示用参数分段读。
- 输入体积:`.docx` 40 MB、`.xlsx` 60 MB;三者都是 zip,解压后总量超过 **1 GiB** 或压缩比超过 **150:1** 会按「疑似压缩炸弹」拒绝(`ZIP_BOMB_SUSPECTED`)。

**限制**

- 图表支持 `bar` / `column` / `line` / `pie`,以 OOXML 注入实现,打开时由 Excel/WPS 计算数据;组合图、双轴等复杂图表请手动调整。
- 一个工作表只能有一个 drawing 部件,因此同表多图表共用一个 drawing(各自独立锚点,可分别拖动)。
- `.doc` 纯 JS 解析不保留表格与版式;写出 `.doc` / `.odt`、转 `.pdf` 需要 LibreOffice 或 Word。
- 模板填充只做 `{{变量}}` 替换,不支持循环 / 条件(批量套打多次调用即可)。
- 文档内嵌图片只读不写:写入 `.docx` 时会移除 `<img>`(有 `alt` 文本则保留为正文),不会嵌入图片。
- `office_convert` 源与目标扩展名相同时:路径不同则原样复制,路径相同则直接返回。
- CSV / TSV / TXT 按 **UTF-8** 解码(这类格式不自带编码信息);GBK 等其他编码请先转成 UTF-8。

## 常见问题

**Windows 上提示 `'git' 不是内部或外部命令`?**
`github:` 形式的依赖必须调用 git 拉取。改用固定直链那条命令即可绕开,或执行 `winget install --id Git.Git -e` 装上 git。

**内网 / 无外网机器怎么装?**
下载 `dsh-office-toolkit-offline.zip`(内含完整 `node_modules`),解压后 `link:` 安装,全程零下载。Windows 路径记得用正斜杠:`dsh plugin --profile web add link:C:/dsh-office-toolkit`。

**智能体还是先用了通用 `read`?**
见「让智能体优先使用本插件」。工具描述已强化,再加一条用户级指令即可稳定命中。

**更新完感觉没变化?**
插件在 DSH 启动时加载,必须**重启 `dsh web`(或重开桌面端)并新建会话**。校验:`dsh --profile web --dump-config | grep -A1 dsh-office-toolkit`。

**转 PDF 报错?**
`.pdf` 输出依赖本机 LibreOffice 或 Microsoft Word,纯 JS 不提供 PDF 渲染。

**装完提示 `Ignored build scripts`?**
依赖里只有 `@turbodocx/html-to-docx` 带一个 `postinstall`,内容仅仅是打印一条推广文案(读本地 `messages.json`,不下载、不写盘、不执行外部命令)。pnpm 默认就不执行依赖的构建脚本,所以这条提示可以忽略;想少看日志就加 `--loglevel=error`。

## 依赖与已知告警

- **`xlsx` 用的是 `@e965/xlsx@0.20.3`**:npm 上的 `xlsx` 停在 `0.18.5`,带 Prototype Pollution(`GHSA-4r6h-8v6p-xvw6`)与 ReDoS(`GHSA-5pgg-2g8v-p4x9`)两个 high;SheetJS 早已停止在 npm 发布,修复版只在其官方 CDN。但官方 CDN 的 URL 形式依赖会被 pnpm 的 `blockExoticSubdeps` 判为 exotic 子依赖而拒绝安装,所以改用 npm 上该官方构建的自动转发包(月下载 300 万+,仓库 [sheetjs-npm-publisher](https://github.com/e965/sheetjs-npm-publisher))。包名不同,故代码里写 `import('@e965/xlsx')`。
- **换用两个维护中的 fork,`npm audit` 归零**(0.3.12):
  - `exceljs@4.4.0` → **`@wekanteam/exceljs@4.7.3`**([Wekan](https://github.com/wekan/exceljs) 维护的同线 fork)。原版锁定 `uuid@^8.3.0`,而 `uuid` 的 [GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq) 只在 `11.1.1` 修复;fork 已升到 `uuid@^14`。4.x 同线,API 不变。
  - `html-to-docx@1.8.0` → **`@turbodocx/html-to-docx@1.23.1`**([TurboDocx](https://github.com/TurboDocx/html-to-docx) 维护的 fork,周下载 9 万+)。原版依赖 `image-size`,而它有 2 个 high DoS 且 **npm 上全部版本都受影响、没有修复版**;fork 换成了 `probe-image-size`。
  - 两个 fork 都在 npm 与国内镜像上可直接安装,依赖树**反而更小**(207 → 204 包,83.5 → 81.8 MB)。

## 设计说明

解析 OOXML / HTML / ODT 时没有用 `/<[^>]+>/g`、`/([A-Za-z]+\d+)/` 这类正则:它们带无界量词且分支可重叠,在构造输入上会让回溯引擎反复重扫同一段文本,耗时随输入超线性增长(ReDoS,SonarQube `S5852`)。现在统一走 `lib/core/markup.js` 的一次前向扫描,每个字符只被访问有限次,**复杂度是输入长度的 O(n) 上界,与输入内容无关**。`test/converters.test.mjs` 覆盖 20 万字符级畸形输入,要求毫秒级返回或快速报错。

## 本地开发

```sh
npm test           # 三个测试脚本:核心库 40 · 跨平台层 16 · 插件适配层 36,共 92 项
npm run coverage   # 同上并统计覆盖率,写出 coverage/lcov.info
```

纯 Node 脚本,不需要测试框架。覆盖率(c8):**语句 91.8% / 分支 68.7% / 函数 93.7%**。未覆盖的主要是 `legacy-external.js` 的 Word COM / LibreOffice 分支与 `converters.js` 的纯 JS 回退 —— 它们只在没有 textutil / LibreOffice 的机器上才会走到。

profile 当前是 `link:` 安装,改完源码重启 `dsh web` 即生效;若用 `file:` 安装需重新执行一次 `add` 刷新副本(HMR 不监听插件源码)。

代码质量走 SonarQube(项目 `dsh-office-toll`;`sonar-project.properties` 不入库):

```sh
export SONAR_TOKEN=<token> && sonar-scanner   # 会读取 coverage/lcov.info
```

当前 0 缺陷 / 0 漏洞 / 0 代码异味 / 0 安全热点;整体覆盖率 85.4%、新代码覆盖率 88.9%(门槛 80%),质量门通过;可靠性 · 安全性 · 可维护性均 A 级。

### 发布到 npm(维护者)

```sh
npm login --registry https://registry.npmjs.org/   # 首次,需要 npm 账号 + 2FA
npm publish
```

`package.json` 的 `publishConfig.registry` 已固定为官方源,所以 `npm publish` **不受 `~/.npmrc` 里国内镜像的影响**(镜像只能读不能发)。发布后裸包名 `dsh plugin --profile web add dsh-office-toolkit` 即可用。

### 发版(维护者)

1. 改 `package.json` 版本号并同步 README「版本记录」;
2. 跑上面三个测试脚本;
3. 提交并打 tag:`git tag -a vX.Y.Z -m "..." && git push origin main vX.Y.Z`;
4. `npm pack` 产出 `.tgz`;把目录连同 `node_modules` 一起打包成 `-offline.zip`;
5. 建 Release 并上传 **4 个附件**:带版本号与不带版本号各一份 `.tgz` 与 `-offline.zip`(不带版本号的那两份供 `releases/latest/download/` 固定别名使用)。

> 上传前确认离线包确实含 `node_modules`:`unzip -l ...-offline.zip | grep -c node_modules/` 应有数千条。

## 版本记录

| 版本 | 变更 |
| --- | --- |
| **0.3.12** | 安全收口:`npm audit` **归零**(换用 `@wekanteam/exceljs` / `@turbodocx/html-to-docx` 两个维护 fork)、防解压炸弹(解压总量 / 压缩比上限)、写入 `.docx` 前剥离图片 |
| **0.3.11** | 安全审计:修公式注入 / 符号链接绕过围栏 / 临时文件全局可读,并公开剩余风险 |
| **0.3.10** | 接入 c8 覆盖率(语句 91.6%);修 3 个 bug:CSV 读出乱码、`office_edit_xlsx` 的 `sheet` 序号基准、日期回读差一天 |
| **0.3.9** | README 改版:居中标题与徽章、导航、常见问题、工作方式图、版本记录 |
| **0.3.8** | 补充「更新」说明:重跑同一条 `add` 即可更新 |
| **0.3.7** | 强化工具描述,让智能体优先用本插件读写 Office 文件 |
| **0.3.6** | `xlsx` 换用 `@e965/xlsx@0.20.3`,消除 2 个 high 漏洞 |
| **0.3.5** | 补齐 npm 元数据并把发布源固定为官方 registry |
| **0.3.4** | 新增永不过期的安装直链 `releases/latest/download/...` |
| **0.3.3** | 补充无 git / 完全离线机器的安装方式 |
| **0.3.2** | 精简文档并补仓库简介 |
| **0.3.1** | 仓库名与包名统一为 `dsh-office-toolkit` |
| **0.3.0** | 安全热点清零:标签与引用解析改为线性扫描 |
| **0.2.x** | 修复同一工作表多图表丢失;新增 Windows 支持 |

## 目录结构

```
lib/index.js                 宿主适配层:工具注册、路径解析、沙箱围栏、fs/observed 事件
lib/core/office.js           六个操作的编排
lib/core/word.js             docx 读 / 写 / 模板(mammoth / @turbodocx/html-to-docx / docxtemplater)
lib/core/excel.js            xlsx 读 / 建 / 编辑(@wekanteam/exceljs)
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

## 安全说明

### 已经修掉的

| 问题 | 影响 | 处理 |
| --- | --- | --- |
| **公式注入** | 把外部 CSV/TSV/TXT 转成 `.xlsx` 时,`=` 开头的内容被写成**活公式**(`=cmd\|'/c calc'!A0`、`=HYPERLINK(...)`),用户在 Excel 里打开就可能触发 DDE / 外链 | 文本格式来源的表格一律中和公式(保持字符串、清掉公式字段);`office_write_xlsx` 自己的公式能力不受影响 |
| **符号链接绕过写入围栏** | 围栏此前只做词法比较:工作区内一个指向外部的符号链接,能让写入落到沙箱之外 | 围栏同时校验**真实路径**(realpath),绕行会被拒绝 |
| **临时文件全局可读** | 转换旧格式时会把文档内容写进公共临时目录的 `0644` 文件,同机其他用户可读 | 临时文件显式 `0600` |
| **正则回溯(ReDoS)** | 构造输入可让宿主进程长时间卡住 | 0.3.0 起所有标签/引用解析改为线性扫描,见「设计说明」 |
| **解压炸弹** | `.docx` / `.xlsx` / `.odt` 都是 zip,此前只限制压缩包体积,一个几十 KB 的文件可解压出几十 GB,把宿主进程撑爆 | 0.3.12 起解析前先读 zip 中央目录,声明解压总量超过 1 GiB 或压缩比超过 150:1 直接拒绝 |
| **图片探测 DoS 与外链抓取** | 写 `.docx` 时 html-to-docx 会去探测 `<img>` 的尺寸:旧版走 `image-size`(ICNS / JXL / HEIF 解析有无上限循环,`GHSA-w3rx-r6r6-pgpr`、`GHSA-5p2g-fcmc-qvqq`,**上游至今没有修复版**),新版走 `probe-image-size` → `needle`(会按 `<img src>` 真的发起 HTTP 请求) | 0.3.12 起写入前先剥掉 `<img>` / `<figure>`(`alt` 文本保留) —— 两条路径都不会被触发,插件始终不访问网络 |
| **`xlsx` 已知漏洞** | Prototype Pollution 与 ReDoS,解析不可信表格时可达 | 0.3.6 起换用 `@e965/xlsx@0.20.3` |
| **依赖链上的 4 条 `npm audit` 告警** | `image-size` 2 个 high(经 html-to-docx,上游无修复版)+ `uuid` 1 个 moderate(经 exceljs,修复版只在 11.x) | 0.3.12 起换用维护中的 fork(`@turbodocx/html-to-docx`、`@wekanteam/exceljs`),**`npm audit` 归零** |

### 边界与假设

- **写入围栏由插件自己实现**(`lib/core/path-guard.js` + `lib/index.js`)。`.docx` / `.xlsx` 是二进制,而 DSH 的 `ctx.fs` 只提供 `writeText`,插件只能用 `node:fs` 落盘 —— 所以**这个围栏就是真正的边界**,不存在宿主写入沙箱兜底。
- 围栏允许写入:DSH 策略给出的 `workspaceRoot`、会话 cwd、`process.cwd()`、系统临时目录。其中 `process.cwd()` 是镜像 DSH 默认沙箱根的兜底 —— 若 `dsh web` 从很宽的目录(例如用户主目录)启动,可写范围会随之变宽,**建议从工作区目录启动**。
- **读取不做围栏**(与内置 `read` 工具一致):只按扩展名区分,不限制目录;`office_convert` 的源文件同理。
- 不访问网络:插件自身不发起任何请求,写入 `.docx` 前也会剥掉 `<img>`,因此 `@turbodocx/html-to-docx` 里那套可联网的图片探测栈(`probe-image-size` / `needle`)不会被触发。
- 不常驻后台。本包自身没有 `prepare` / `postinstall`;依赖树里唯一的安装脚本是 `@turbodocx/html-to-docx` 的 `postinstall`,只打印推广文案(见「常见问题」),且 pnpm 默认不执行依赖构建脚本。

### 已知且暂不修复

- **zip 声明的解压体积可以伪造**:本插件的防护基于 zip 中央目录里的 `uncompressedSize`(以及压缩比)。蓄意构造的压缩包可以把该字段写小,此时仍会在真实解压时膨胀 —— 这层防护抬高的是门槛而**不是硬边界**;真要处理完全不可信的输入,请在独立进程 / 容器里跑并限制内存。
- **供应链**:`@e965/xlsx` 是 SheetJS 官方构建在 npm 上的第三方转发(月下载 300 万+),不是官方 publisher;`package-lock.json` 已锁定 integrity。若对此敏感,可改用官方 CDN 的 URL 依赖 —— 但 pnpm 默认的 `blockExoticSubdeps` 会拒绝这种形式。`@wekanteam/exceljs`(Wekan)与 `@turbodocx/html-to-docx`(TurboDocx)同样是「官方包停更后的第三方维护 fork」,它们换来的是一条完全干净的 `npm audit`;两者的发布方都是有长期公开仓库的组织,但**这与上游官方包并非同一 publisher**,请自行评估后决定是否接受。

### 报告安全问题

请通过 [GitHub 私密漏洞报告](https://github.com/cnkids/dsh-office-toolkit/security/advisories/new) 提交,不要发在公开 Issue 里。

本项目基于 [MIT License](LICENSE) 开源。
