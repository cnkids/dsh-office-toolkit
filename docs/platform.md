# 平台支持

[← 返回 README](../README.md)

纯 JS 层保证开箱可用，外部转换器只用于提升旧格式保真度。需要 **Node ≥ 20**。

| 能力 | 纯 JS | 外部转换器 |
| --- | --- | --- |
| docx / xlsx 读写、编辑、图表 | ✅ | — |
| xls / xlsb / ods / csv 读、写、转换 | ✅ SheetJS | — |
| doc / rtf / odt 读取 | ✅ 内置解析器 | 更高保真 |
| doc / rtf / odt 写出 | 仅 rtf | 必需 |
| 转 pdf | ❌ | 必需 |

外部转换器按平台自动探测：macOS 用系统自带 `textutil`；Windows / Linux 用 LibreOffice `soffice`，Windows 装了 Word 时自动改用 Word COM。**Windows 上不装任何 Office 也能读写 docx/xlsx/xls/csv、编辑表格、加图表、套模板。**

## 各格式保真度

- **`.docx` / `.xlsx`**：读取走 mammoth 与 `@wekanteam/exceljs`；写入由内置 OOXML 生成器（`docx-writer.js`）与 `charts.js` 完成 —— 样式、公式、合并、冻结、图表都是原生 OOXML，Excel / WPS / Word 打开即用，不需要任何本机 Office。
- **`.doc` / `.rtf` / `.odt`**：内置纯 JS 解析器保证可读，但不保留表格与版式；本机装了转换器时自动升级为更高保真。
- **`.xls` / `.xlsb` / `.ods` / `.csv` / `.tsv`**：走 SheetJS，全平台纯 JS。
- **写 `.doc` / `.odt`、转 `.pdf`**：必须依赖 LibreOffice 或 Microsoft Word，纯 JS 不提供渲染。

## 工作方式

```text
用户 / 智能体
   │
   ├─ office_read · office_write_docx · office_write_xlsx
   │  office_edit_xlsx · office_fill_docx_template · office_convert
   │
   ├─ docx / xlsx ─────────── 纯 JS：mammoth 读取 · 自研 OOXML 生成（docx）· @wekanteam/exceljs · 自研图表注入
   ├─ xls / xlsb / ods / csv ─ 纯 JS：SheetJS（@e965/xlsx 0.20.3）
   └─ doc / rtf / odt ──────── 内置线性解析器优先
                                 └─ 不可用时 → 本机转换器
                                    macOS textutil / LibreOffice / Word COM
   │
   └─ 写入前经过 DSH 沙箱围栏：仅会话工作区与系统临时目录
```
