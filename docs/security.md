# 安全说明

[← 返回 README](../README.md)

## 已经修掉的

| 问题 | 影响 | 处理 |
| --- | --- | --- |
| 公式注入 | 外部 CSV/TSV/TXT 转 `.xlsx` 时 `=` 开头内容成为活公式（`=cmd\|'/c calc'!A0`），打开可能触发 DDE / 外链 | 文本来源一律中和公式（`office_write_xlsx` 自己的公式不受影响） |
| 符号链接绕过围栏 | 围栏只做词法比较，工作区内一个指向外部的软链即可把写入带出沙箱 | 同时校验 realpath |
| 临时文件全局可读 | 旧格式转换的 `0644` 临时文件同机可读 | 显式 `0600` |
| 正则回溯（ReDoS） | 构造输入可让宿主长时间卡住 | 所有标签 / 引用解析改为线性扫描（0.3.0） |
| 解压炸弹 | 只限制压缩包体积，几十 KB 的 zip 可解压出几十 GB | 解压总量 > 1 GiB 或压缩比 > 150:1 直接拒绝（0.3.12） |
| 图片探测 DoS / 外链抓取 | `html-to-docx` 用 `image-size` 量图片尺寸，而它有 2 个 high DoS（ICNS / JXL / HEIF）且**无修复版**；其维护 fork 改用 `probe-image-size` → `needle`，会按 `<img src>` 真发 HTTP 请求 | 0.3.13 起 `.docx` 由自研生成器产出（`docx@9`），不解析图片、只保留 `alt`，图片与网络两条路径都不存在 |
| `xlsx` 已知漏洞 | Prototype Pollution + ReDoS | 0.3.6 起换 `@e965/xlsx@0.20.3` |
| 依赖链 4 条 `npm audit` 告警 | `uuid`、`image-size` 上游均不可修 | 换 `@wekanteam/exceljs` + 自研 docx 生成，`npm audit` 归零；整棵依赖树无安装脚本（0.3.12 / 0.3.13） |
| 危险链接 | `javascript:` / `data:` / `vbscript:` / `file:` 链接可被写进文档与关系文件 | 写入 `.docx` 时降级为纯文本（0.3.13） |

## 边界与假设

- **写入围栏是插件自己实现的**（`lib/core/path-guard.js` + `lib/index.js`）：`.docx` / `.xlsx` 是二进制，而 DSH 的 `ctx.fs` 只提供 `writeText`，插件只能用 `node:fs` 落盘 —— 所以**这个围栏就是真正的边界**，不存在宿主写入沙箱兜底。可写范围：DSH 策略给出的 `workspaceRoot`、会话 cwd、`process.cwd()`、系统临时目录；其中 `process.cwd()` 是镜像 DSH 默认沙箱根的兜底，**建议从工作区目录启动 `dsh web`**，否则可写范围会随启动目录变宽。
- **读取不做围栏**（与内置 `read` 工具一致）：只按扩展名区分，不限制目录；`office_convert` 的源文件同理。
- 不联网、不常驻后台；**本包与整棵依赖树都没有 `preinstall` / `install` / `postinstall` 脚本**，`dsh plugin add` 不会停下来要求批准构建脚本。

## 残余风险

- **zip 声明的解压体积可以伪造**：本插件的防护基于 zip 中央目录里的 `uncompressedSize`（以及压缩比）。蓄意构造的压缩包可以把该字段写小，此时仍会在真实解压时膨胀 —— 这层防护抬高的是门槛而**不是硬边界**；真要处理完全不可信的输入，请在独立进程 / 容器里跑并限制内存。
- **供应链**：`@e965/xlsx` 与 `@wekanteam/exceljs` 不是上游官方 publisher，详见[设计说明 → 依赖与已知告警](design.md#依赖与已知告警)。

## 报告安全问题

请通过 [GitHub 私密漏洞报告](https://github.com/cnkids/dsh-office-toolkit/security/advisories/new) 提交，不要发在公开 Issue 里。
