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
  <img src="https://img.shields.io/badge/node-%E2%89%A520-339933?style=flat" alt="Node.js 20 or newer">
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

核心能力全部纯 JavaScript 实现,**不装 Office 也能用**;只有写出 `.doc` / `.odt` 和转 `.pdf` 时才可选地借用本机转换器。插件只做文档读写这一件事:不联网、不常驻后台、不修改 DSH 上游代码。

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

包名与仓库名均为 `dsh-office-toolkit`。**没装 git 就用第一条**,它不需要 git,也永远指向最新版。

| 场景 | 命令 |
| --- | --- |
| **无 git(推荐)** | `dsh plugin --profile web add https://github.com/cnkids/dsh-office-toolkit/releases/latest/download/dsh-office-toolkit.tgz` |
| 有 git · 跟随 main | `dsh plugin --profile web add github:cnkids/dsh-office-toolkit` |
| 完全离线 | 解压 `dsh-office-toolkit-offline.zip` 后 `dsh plugin --profile web add link:C:/dsh-office-toolkit` |
| 本地源码 | `dsh plugin --profile web add link:/path/to/dsh-office-toolkit` |

`releases/latest/download/...` 是固定别名(每个 Release 都传带版本号与不带版本号两份附件),所以命令不用随版本改。`.tgz` 只有几十 KB,依赖走 npm;离线包含完整 `node_modules`,零下载。

`dsh plugin` 在 profile 目录执行 `pnpm add`,并自动把包名写进 `dsh.profile.bundles`(本包声明了 `dsh.bundle`)。**装完必须重启 `dsh web` 并新建会话**才会加载。

```sh
dsh --profile web --dump-config | grep -A1 dsh-office-toolkit   # 校验
dsh plugin --profile web remove dsh-office-toolkit              # 卸载
```

目标机器需 **Node ≥ 20** 与 pnpm;`.tgz` 不含依赖,需要能访问 npm registry(内网先配镜像)。

### 更新

**重跑当初那条 `add` 就是更新**;一次更新 profile 内全部依赖用 `dsh plugin --profile web update`。

- `link:` 安装不需要 pnpm,源码目录 `git pull` 即最新。
- 装了**带版本号**的 spec(`#vX.Y.Z`、`.../releases/download/vX.Y.Z/...tgz`)不会自动前进 —— 这也是推荐固定直链的原因。
- **更新后必须重启 `dsh web` 并新建会话**,否则看起来像没更新(工具列表与描述只在启动时读取)。

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

想更保险,在**用户级指令文件** `~/.dsh/AGENTS.md`(对所有项目与会话生效)里加一句:Office 文件一律用 `office_read` / `office_write_*` / `office_edit_xlsx` / `office_fill_docx_template` / `office_convert`,不要用通用 `read`/`write`/`edit`。该文件热加载,保存即生效。

## 工作方式

```text
用户 / 智能体
   │
   ├─ office_read · office_write_docx · office_write_xlsx
   │  office_edit_xlsx · office_fill_docx_template · office_convert
   │
   ├─ docx / xlsx ─────────── 纯 JS:mammoth 读取 · 自研 OOXML 生成(docx)· @wekanteam/exceljs · 自研图表注入
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

- 图表支持 `bar` / `column` / `line` / `pie`(OOXML 注入,数据由 Excel/WPS 打开时计算);组合图、双轴请手动调整。同表多图表共用一个 drawing(一个工作表只能有一个)。
- `.doc` 纯 JS 解析不保留表格与版式;写出 `.doc` / `.odt`、转 `.pdf` 需要 LibreOffice 或 Word。
- 模板填充只做 `{{变量}}` 替换,不支持循环 / 条件(批量套打多次调用)。
- 图片只读不写:写入 `.docx` 时不嵌入图片,只把 `<img>` 的 `alt` 文本留在正文里。
- `office_convert` 源与目标扩展名相同时,路径不同则复制、路径相同则直接返回。
- CSV / TSV / TXT 按 **UTF-8** 解码(这类格式不带编码信息),GBK 请先转码。

## 常见问题

**Windows 提示 `'git' 不是内部或外部命令`?** `github:` 形式要用 git,改用固定直链即可绕开,或 `winget install --id Git.Git -e`。

**内网 / 无外网怎么装?** 用 `-offline.zip`(内含完整 `node_modules`)解压后 `link:` 安装,零下载。

**智能体还是先用了通用 `read`?** 见「让智能体优先使用本插件」。

**更新完没变化 / 转 PDF 报错?** 更新后必须重启 `dsh web` 并新建会话;`.pdf` 输出依赖本机 LibreOffice 或 Word,纯 JS 不提供 PDF 渲染。

## 依赖与已知告警

**`npm audit` 0 条,整棵依赖树没有任何安装脚本**(0.3.13 起)。三处取舍:

- **`.docx` 生成是自己写的**(`lib/core/docx-writer.js`,基于 `docx@9`),不再用 `html-to-docx`:它依赖的 `image-size` 有 2 个 high DoS 且 **npm 上所有版本都受影响、没有修复版**;而唯一清掉它的维护 fork 又带 `postinstall` + `axios`/`needle`,会让 `dsh plugin add` 停下来要求批准构建脚本。
- **`exceljs` 换成 `@wekanteam/exceljs`**([Wekan](https://github.com/wekan/exceljs) 维护的 4.x 同线 fork):上游锁 `uuid@^8.3.0`,而 `uuid` 的告警只在 `11.1.1` 修复,fork 已升到 `uuid@^14`。
- **`xlsx` 用 `@e965/xlsx@0.20.3`**:npm 上的 `xlsx` 停在 `0.18.5`(Prototype Pollution + ReDoS 两个 high),修复版只在 SheetJS 自建 CDN;而 URL 形式依赖会被 pnpm 的 `blockExoticSubdeps` 拒绝,故改用该官方构建的 npm 转发包(月下载 300 万+,[sheetjs-npm-publisher](https://github.com/e965/sheetjs-npm-publisher))。

后两个都不是上游官方 publisher(`@e965` 是转发,`@wekanteam` 是他人维护的 fork),`package-lock.json` 已锁 integrity;介意请自行评估。

## 设计说明

解析 OOXML / HTML / ODT 不用 `/<[^>]+>/g` 这类正则:带无界量词与可重叠分支,构造输入会让回溯反复重扫同一段文本(ReDoS,`S5852`)。现在统一走 `lib/core/markup.js` 的一次前向扫描,复杂度与输入内容无关,是长度的 O(n) 上界;`test/converters.test.mjs` 用 20 万字符畸形输入验证毫秒级返回。

## 本地开发

```sh
npm test           # 三个测试脚本:核心库 43 · 跨平台层 16 · 插件适配层 36,共 95 项
npm run coverage   # 同上并统计覆盖率,写出 coverage/lcov.info
```

纯 Node 脚本,无需测试框架。覆盖率(c8):**语句 92.5% / 分支 70.8% / 函数 93.9%**;未覆盖的主要是 `legacy-external.js` 的 Word COM / LibreOffice 分支与 `converters.js` 的纯 JS 回退。

`link:` 安装改完源码重启 `dsh web` 即生效;`file:` / `github:` 安装需重跑一次 `add` 刷新副本(HMR 不监听插件源码)。

代码质量走 SonarQube(项目 `dsh-office-toll`;`sonar-project.properties` 不入库):`export SONAR_TOKEN=<token> && sonar-scanner`(读 `coverage/lcov.info`)。当前 0 缺陷 / 0 漏洞 / 0 代码异味 / 0 安全热点,整体覆盖率 86.7%、新代码覆盖率 93.9%(门槛 80%),质量门通过。

### 发布到 npm(维护者)

```sh
npm login --registry https://registry.npmjs.org/   # 首次,需要 npm 账号 + 2FA
npm publish
```

`publishConfig.registry` 已固定为官方源,所以**不受 `~/.npmrc` 国内镜像影响**(镜像只能读不能发)。发布后裸包名 `add dsh-office-toolkit` 即可用。

### 发版(维护者)

1. 改 `package.json` 版本号并同步 README「版本记录」;
2. 跑三个测试脚本 + Sonar;
3. 提交、`git tag -a vX.Y.Z -m "..."`、push;
4. `npm pack` 出 `.tgz`;目录连同 `node_modules`(`npm install --omit=dev` 后)打包成 `-offline.zip`,上传前用 `unzip -l ...-offline.zip | grep -c node_modules/` 确认为数千条;
5. 建 Release,上传 **4 个附件**:带版本号与不带版本号各一份 `.tgz` 与 `-offline.zip`(不带版本号的两份供 `releases/latest/download/` 固定别名使用)。

## 版本记录

| 版本 | 变更 |
| --- | --- |
| **0.3.13** | `.docx` 生成改为自研(`docx@9`),去掉带 `postinstall` 的依赖 —— 安装不再被 pnpm 打断;精简 README,修正 Node ≥ 20 与覆盖率数字 |
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
lib/core/word.js             docx 读 / 写 / 模板(mammoth / docxtemplater)
lib/core/docx-writer.js      HTML → OOXML 文档生成(自研,基于 docx@9)
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
| 公式注入 | 外部 CSV/TSV/TXT 转 `.xlsx` 时 `=` 开头内容成为活公式(`=cmd\|'/c calc'!A0`),打开可能触发 DDE / 外链 | 文本来源一律中和公式(`office_write_xlsx` 自己的公式不受影响) |
| 符号链接绕过围栏 | 围栏只做词法比较,工作区内一个指向外部的软链即可把写入带出沙箱 | 同时校验 realpath |
| 临时文件全局可读 | 旧格式转换的 `0644` 临时文件同机可读 | 显式 `0600` |
| 正则回溯(ReDoS) | 构造输入可让宿主长时间卡住 | 所有标签 / 引用解析改为线性扫描(0.3.0) |
| 解压炸弹 | 只限制压缩包体积,几十 KB 的 zip 可解压出几十 GB | 解压总量 > 1 GiB 或压缩比 > 150:1 直接拒绝(0.3.12) |
| 图片探测 DoS / 外链抓取 | `html-to-docx` 用 `image-size` 量图片尺寸,而它有 2 个 high DoS(ICNS / JXL / HEIF)且**无修复版**;其维护 fork 改用 `probe-image-size` → `needle`,会按 `<img src>` 真发 HTTP 请求 | 0.3.13 起 `.docx` 由自研生成器产出(`docx@9`),不解析图片、只保留 `alt`,图片与网络两条路径都不存在 |
| `xlsx` 已知漏洞 | Prototype Pollution + ReDoS | 0.3.6 起换 `@e965/xlsx@0.20.3` |
| 依赖链 4 条 `npm audit` 告警 | `uuid`、`image-size` 上游均不可修 | 换 `@wekanteam/exceljs` + 自研 docx 生成,`npm audit` 归零;整棵依赖树无安装脚本(0.3.12 / 0.3.13) |

### 边界与假设

- **写入围栏是插件自己实现的**(`path-guard.js`):`.docx` / `.xlsx` 是二进制,而 `ctx.fs` 只提供 `writeText`,只能用 `node:fs` 落盘 —— **没有宿主沙箱兜底**。可写范围:策略给出的 `workspaceRoot`、会话 cwd、`process.cwd()`、系统临时目录;其中 `process.cwd()` 是兜底,**建议从工作区目录启动 `dsh web`**,否则可写范围会随启动目录变宽。
- **读取不做围栏**(与内置 `read` 一致),`office_convert` 的源文件同理。
- 不联网、不常驻后台;**本包与整棵依赖树都没有 `preinstall` / `install` / `postinstall` 脚本**,`dsh plugin add` 不会停下来要求批准构建脚本。

### 残余风险

- zip 头里的 `uncompressedSize` 可以伪造,解压炸弹防护抬高的是门槛而**不是硬边界**;处理完全不可信的输入请在独立进程 / 容器里跑并限制内存。
- `@e965/xlsx` 与 `@wekanteam/exceljs` 不是上游官方 publisher,见「依赖与已知告警」。

### 报告安全问题

请通过 [GitHub 私密漏洞报告](https://github.com/cnkids/dsh-office-toolkit/security/advisories/new) 提交,不要发在公开 Issue 里。

本项目基于 [MIT License](LICENSE) 开源。
