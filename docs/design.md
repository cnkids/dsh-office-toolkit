# 设计说明

[← 返回 README](../README.md)

## 为什么不用正则解析

解析 OOXML / HTML / ODT 时没有用 `/<[^>]+>/g`、`/([A-Za-z]+\d+)/` 这类正则：它们带无界量词且分支可重叠，在构造输入上会让回溯引擎反复重扫同一段文本，耗时随输入超线性增长（ReDoS，SonarQube `S5852`）。现在统一走 `lib/core/markup.js` 的一次前向扫描，每个字符只被访问有限次，**复杂度是输入长度的 O(n) 上界，与输入内容无关**。`test/converters.test.mjs` 覆盖 20 万字符级畸形输入，要求毫秒级返回或快速报错；`office_query` 的数值 / 日期识别同样只用分支互斥、无重叠量词的正则（`\d+\.?\d*` 这类写法在超长数字串上会退化成 O(n²)），`test/query.test.mjs` 用 6 万位畸形数字串守住这条线。

`.docx` 的生成同样是自研的（`lib/core/docx-writer.js`）：把 HTML 解析成节点树后直接产出 OOXML，而不是拼接 HTML 字符串再交给第三方库 —— 这样才可能做到「不解析图片、依赖树无安装脚本」。

## 目录结构

```text
lib/index.js                 宿主适配层：工具注册、路径解析、沙箱围栏、fs/observed 事件
lib/core/office.js           七个操作的编排
lib/core/word.js             docx 读 / 写 / 模板（mammoth / docxtemplater）
lib/core/docx-writer.js      HTML → OOXML 文档生成（自研，基于 docx@9）
lib/core/excel.js            xlsx 读 / 建 / 编辑（@wekanteam/exceljs）
lib/core/query.js            表格计算：筛选 / 分组 / 聚合 / 排序 / 画像（纯函数，不碰 IO）
lib/core/charts.js           OOXML 图表注入
lib/core/legacy.js           旧表格格式（SheetJS）
lib/core/legacy-read.js      纯 JS 解析 doc / rtf / odt
lib/core/legacy-external.js  外部转换器（textutil / LibreOffice / Word COM）
lib/core/converters.js       旧 Word 格式公共 API
lib/core/markup.js           HTML / XML 线性扫描器
lib/core/md.js               Markdown ↔ HTML
lib/core/path-guard.js       跨平台路径围栏
lib/core/util.js             错误类型、体积上限、截断
test/                        四个测试脚本
```

## 依赖与已知告警

**`npm audit` 0 条，整棵依赖树没有任何安装脚本**（0.3.13 起）。三处取舍：

- **`.docx` 生成是自己写的**（`lib/core/docx-writer.js`，基于 `docx@9`），不再用 `html-to-docx`：它依赖的 `image-size` 有 2 个 high DoS 且 **npm 上所有版本都受影响、没有修复版**；而唯一清掉它的维护 fork 又带 `postinstall` + `axios`/`needle`，会让 `dsh plugin add` 停下来要求批准构建脚本。
- **`exceljs` 换成 `@wekanteam/exceljs`**（[Wekan](https://github.com/wekan/exceljs) 维护的 4.x 同线 fork）：上游锁 `uuid@^8.3.0`，而 `uuid` 的告警只在 `11.1.1` 修复，fork 已升到 `uuid@^14`。
- **`xlsx` 用 `@e965/xlsx@0.20.3`**：npm 上的 `xlsx` 停在 `0.18.5`（Prototype Pollution + ReDoS 两个 high），修复版只在 SheetJS 自建 CDN；而 URL 形式依赖会被 pnpm 的 `blockExoticSubdeps` 拒绝，故改用该官方构建的 npm 转发包（月下载 300 万+，[sheetjs-npm-publisher](https://github.com/e965/sheetjs-npm-publisher)）。代码里写 `import('@e965/xlsx')`。

后两个都不是上游官方 publisher（`@e965` 是转发，`@wekanteam` 是他人维护的 fork），`package-lock.json` 已锁 integrity；介意请自行评估。
