# 版本记录

[← 返回 README](../README.md)

| 版本 | 变更 |
| --- | --- |
| **0.3.31** | 修「误导性报错」:**根因在 DSH,不在本插件**,本版只做插件侧能做的两件事 —— 改报错文案、记已知问题。DSH 0.1.6-alpha.1 / .2 的 profile 模块解析器把 `process/` 这类「内建名 + 尾斜杠」说明符切掉子路径后没重新判断内建,再对 `resolve.paths()` 返回的 `null` 做 `for...of`,于是抛 TypeError;ExcelJS 的依赖 `readable-stream` 内部的 `require('process/')`(尾斜杠本意是走 npm polyfill 包)正好命中,所以**只有走 ExcelJS 的 xlsx 路径**(`office_read` 读 `.xlsx`、`office_query` 算 `.xlsx`、`office_write_xlsx`、`office_edit_xlsx`)全失败,`.csv` / `.docx` 不受影响。
此前 `depFailure()` 把它判成缺包,让用户重跑 `dsh plugin add` 补依赖 —— 依赖其实完好,照做无效。现在先识别宿主故障再判断缺包:命中 `resolve.paths is not a function`,或说明符是「`isBuiltin` 判定的内建名 + 尾斜杠」(`process/`、`string_decoder/`),就报新错误码 `DSH_RESOLVER_BUG`,说明这是 DSH 缺陷、重装/重启无效、插件侧无法绕开,并附 DSH 侧报告链接;真实缺包仍是 `MISSING_DEPENDENCY`,非内建名的尾斜杠说明符不会被误伤。README 常见问题同步补一条 |
| **0.3.30** | 修 PDF 读取在 **Windows 上必然失败**:内置 pdfjs 起 fake worker 走 `await import(workerSrc)`,Node 的 ESM 加载器只接受 `file:`/`data:`/`node:` 三种 scheme,Windows 的 `C:\…` 裸路径直接抛 `Only URLs with a scheme in: file, data, and node`(POSIX 绝对路径碰巧能当模块说明符,所以此前只在 macOS 上验证过 —— PDFKit/pdftotext 又把兜底路径挡住了);现在 `workerSrc` 一律用 `pathToFileURL()` 给 `file://` URL,`cMapUrl`/`standardFontDataUrl` 保持文件系统路径(pdfjs 在 Node 侧用 `fs.readFile` 读,写成 URL 反而 ENOENT),两类地址形态都有断言守;
内置 pdfjs 从 6.3.289 换成 **4.10.38**:6.x 的 `engines` 是 `>= 22.13.0 \|\| >= 24` 且不含 polyfill,在 Node 20/22 上会抛 `Promise.withResolvers is not a function`(Promise.try 更要 Node 24),与插件声明的 `node >= 20` 冲突;4.10.38 声明 `>= 20` 并自带这两个 API 的 polyfill,新增「删掉 Promise.withResolvers / Promise.try 后仍能抽文本」的模拟用例守住这条线;
过滤 pdfjs 在 Node 下的渲染类噪音警告(缺 canvas/DOMMatrix/Path2D、standardFontDataUrl),只按消息前缀滤掉这些,其它日志照常输出 | `office_read` 支持 `.pdf`:`pdf.js` 只抽文本层 —— 返回页数与正文(多页插入 `--- 第 N 页 ---` 标记),`offset`/`maxChars` 与 Word 一致可分段续读,`outline` 给页清单(单页也给「第 1 页」),`format:"html"` 是文本重建,`withFormatting` 明确说明 PDF 没有排版格式报告;后端链 `pdftotext`(poppler,stdin/stdout 不落临时文件) → macOS 自带 PDFKit(经 `osascript` 的 ObjC 桥读 `PDFDocument`,零安装) → 随包 `vendor/pdfjs`(Apache-2.0 精简构建:`pdf.min.mjs` + worker + `cmaps` + `standard_fonts`,不渲染、不读系统字体、`isEvalSupported:false`),任一后端失败自动换下一个,并在结果里写明实际用了哪个;加密/有密码的 PDF 报 `PDF_ENCRYPTED`,扫描件(无文本层)明确提示需要 OCR;
`office_read` 的文件内容嗅探新增 PDF 魔数(`%PDF-`),改后缀的 PDF 也能按真实内容读;`office_query` 依旧只算表格,对 PDF 给出明确拒绝;`verify:offline` 增加随包 pdfjs 完整性检查(缺文件即视为离线不可用) |
| **0.3.28** | 未发布前合并:① 新增 `office_edit_docx`(`replace_text` 跨 run 替换 / `set_paragraph` / `set_style` / `set_numbering` / `insert_paragraph` / `delete_paragraph`,以及表格 `set_table` / `insert_table_row`·`delete_table_row` / `insert_table_column`·`delete_table_column` / `merge_table_cells`);只重写 `word/document.xml`,其余 zip 条目逐字节不变;② 排版:`office_write_docx` 的 `style`(字体/字号/行距/首行缩进/对齐/段前后,`headings` 标题,`table` 表格)写进 `docDefaults`;html 内联样式补齐 `font-family`(含 `w:eastAsia`)、`line-height`、`text-indent`、`margin` 并支持容器继承;③ `set_style` 按角色选段落(scope = all/body/headings/table,可叠加区间);④ 表格:框线(all/none/outline/三线表)、表头底纹与加粗、对齐、列宽(按内容自动分配或手动百分比)、单元格内边距 `cellMargins`(→ `w:tblCellMar`);⑤ 表格新增 `delete_table`(按序号或 scope:"all" 删整表,0 行空表也能删;`onlyEmpty:true` 一次清掉所有 0 行残留;删完保证正文不为空且不以表格结尾;0 行表上用增删行列会提示改用 delete_table);`numbering.xml` 里新增 9 级 `abstractNum`,1–9 级按 `w:pStyle` 挂 Heading1–9(标题自动编号 1/1.1/1.1.1),正文列表项按「当前标题层级 +1」写 `numPr`,普通正文不动;缺 `numbering.xml` 时连同关系与 Content_Types 一起补出;⑥ 模板套打:修正「不支持循环/条件」的错误说明 —— docxtemplater 的数组循环与条件本来就可用,现在写进工具描述与文档;
⑦ `office_write_docx` 新增 `header`/`footer`(含 `{page}`/`{total}` 页码域模板)与 `toc`(目录域),并设置 `updateFields` 让 Word 打开时刷新;参数校验错误保留 `INVALID_ARGS` 错误码;
⑧ 表格细化:`set_table` 新增 `repeatHeader`(表头跨页重复)、`cellVerticalAlign`、`rowHeightPt`、`cantSplit`,并新增 `unmerge_table_cells` 取消合并;写文档路径同步支持;
⑨ `office_edit_docx` 新增 `dryRun`(只报改动不写盘)与 `outputPath`(另存不覆盖原文件);
⑩ 新增公文式编号预置 `gongwen-1_1_1_1`(一、→（一）→1.→（1）,中文数字用 chineseCounting);
⑪ 性能:批量操作从逐段拼接改为一次性重建,1 万段文档 set_style 2121ms → 96ms、set_numbering 781ms → 54ms,并加回归测试;
⑫ 读侧:读取 Word 时把自动编号展开成文字(`一、`/`（一）`/`1.1`,项目符号不动),样式联动编号与 `numPr` 都认,`exclude`/`startFrom` 也一致;
⑬ 图片:新建与转换时 `<img>` 嵌成真图片 —— 本地文件或 `data:` URL,按魔数识别 PNG/JPEG/GIF/BMP 并读像素尺寸,按正文宽等比缩放到页内,`width` 属性与 CSS 宽度可作为提示;远程地址、缺失文件、不支持的格式、超过 8 MB 都有明确错误码(`IMAGE_REMOTE`/`IMAGE_NOT_FOUND`/`IMAGE_UNSUPPORTED`/`IMAGE_TOO_LARGE`),全程不联网;
⑭ `office_query` 新增多表 `join`(`on` 支持列名/`{left,right}`/多列数组,`type` 取 inner/left/right/full,同名列只保留一份、非键重名列加后缀,空键不参与匹配,结果写明匹配与未匹配行数;`join` 可传数组依次连接)与 `pivot` 透视表(`rows` × `columns` × `values`,单指标直接用取值当表头,`totals:true` 追加合计行列,列维度取值上限 60);
⑮ `office_edit_xlsx` 新增 `conditional_format`(cellIs/expression/colorScale/dataBar/iconSet/top10/aboveAverage/containsText/timePeriod,`rules` 多条按序定优先级,`style` 写进 dxf)与 `data_validation`(list 下拉 `values` 或 `source` 区域 / whole·decimal·date·textLength 区间 / custom 公式,`range` 整段生效);
⑯ `office_read` 新增 `formulas`(`value`/`formula`/`both`,`true` 等同 `formula`):公式单元格可以看公式本体而不是只给算出来的值,共享公式给出引用地址,`.xls`/`.xlsb`/`.ods` 同样支持;本插件自己写出的公式没有缓存值时,默认回读也带上等号;
⑰ 修两个表格家族判定 bug:`.xlsb` 的主部件是 `xl/workbook.bin`,此前既不认二进制主部件、又被 `_rels/.rels` 里 `extended-properties` 抢先匹配 `officeDocument` 子串,导致合法文件被判成「损坏或改过后缀」;`.ods` 与 `.odt` 的主部件都叫 `content.xml`,此前一律当 Word 文档读(表格结构全丢),现在按 `mimetype` 区分。两条都加了回归测试;
⑱ 修 exceljs 数据验证写回时的重复:`DataValidationsXform` 读入时把 `C2:C100` 拆成逐格、写出时按地址字符串排序贪心合并,会留下 `C10:C100` + `C2:C100` 两条重叠规则;现在写盘前按「相同规则 + 逐格矩形」自己合并,反复读写不增不减;
⑲ 图片读取改由工具层注入:core 不再直接读文件系统,而是由 `lib/index.js` 按沙箱规则解析路径后把字节交给 core(相对路径以输出文档目录为基准,转换时以源文件目录为基准),`.doc`/`.rtf`/`.odt` 回退重建同样接上;顺手去掉一处 `\d+%$` 正则(慢正则热点,改用后缀判断 + `Number()` 解析)与测试里的动态 `RegExp`/`.*(.+).*`;
⑳ 模块划分:单位换算与规格在 `lib/core/docx-style.js`,元素级 XML 手术在 `lib/core/docx-xml.js`,表格在 `lib/core/docx-table.js`,编号在 `lib/core/docx-numbering.js`(读侧展开在 `docx-numbering-read.js`),图片在 `lib/core/image.js`,连接在 `lib/core/join.js`,计算在 `lib/core/query.js` |
| **0.3.27** | Word 长文档支持分段阅读:新增 `offset`(字符偏移续读)与 `outline`(标题大纲,含级别/偏移/标题);分页边界收在句末或换行处,切片与下一段严格衔接(逐段拼回等于全文);offset 越界、非法值、与 `format: "html"` 组合都有明确报错;html 输出不分页但给出切回 text 模式的指引 |
| **0.3.26** | 修 `office_fill_docx_template` 的工具描述/参数说明含 `{{变量}}` 字面量 —— PTC 模式下工具声明会被嵌进 `tools:sdk` 提示词段落并参与变量插值,DSH 直接抛 `malformed prompt variable reference` 导致整个 prompt 组装失败。改为文字描述「两个半角花括号 + 变量名」,结果提示与错误信息同步清理,`test/plugin-smoke.mjs` 递归扫描全部工具 schema 做门禁 |
| **0.3.25** | 新增 `office_query`：表格内直接算(筛选/分组/sum·avg·min·max·count·countDistinct/排序)，全表扫描但只返回结论，不给条件则输出表结构画像；数值识别 `1,234.00`/`¥88`/`12.5%`，日期归一化后比较；类型不可比时判为不可比而非退回字典序。数值识别的正则改为无重叠量词(修掉一处可被超长数字串触发的 O(n²) 回溯)。`office_read` 的截断提示改为指向本工具 |
| **0.3.24** | 安装/更新改用裸包名 `dsh-office-toolkit`;GitHub 直链、离线包、本地源码收进折叠块;发布流程改为 Trusted Publisher 免 token |
| **0.3.23** | 加 GitHub Actions:推 `v*` tag 自动校验版本、跑测试、发布 npm(Trusted Publishing,无需 token)并上传 Release 附件 |
| **0.3.22** | 精简 README 更新说明与版本记录,去掉排查过程叙述 |
| **0.3.21** | 文档:更新说明改用带版本号的直链,并注明 `latest` 直链可能被 pnpm 复用缓存(需 pnpm ≥ 11.10) |
| **0.3.20** | 主文档部件按包关系解析(非规范路径也能读);启动日志与报错带插件版本 |
| **0.3.19** | 非标准 zip 的条目名与引用一起归一化(大小写/反斜杠);mammoth 认不出时用内置解析器兜底 |
| **0.3.18** | 修 0.3.17 回归:不再改写宿主传入的冻结参数 |
| **0.3.17** | 读取兼容性:非标准 zip、改过后缀的文件按真实内容读取;修默认样式解析 |
| **0.3.16** | `.docx` 可读出排版格式(字体/字号/行距/缩进/对齐/页边距),`withFormatting: true` |
| **0.3.15** | 缺依赖时报出包名与修复命令;二进制输出自检;离线包完整性门禁 |
| **0.3.14** | 文档重组：README 收敛为入口（特性 / 安装 / 快速上手 / 文档索引），细节移入 `docs/` |
| **0.3.13** | `.docx` 生成改为自研（`docx@9`），去掉带 `postinstall` 的依赖 —— 安装不再被 pnpm 打断；危险链接降级为纯文本；README 精简、修正 Node ≥ 20 |
| **0.3.12** | 安全收口：`npm audit` **归零**（换用 `@wekanteam/exceljs` / `@turbodocx/html-to-docx` 两个维护 fork）、防解压炸弹（解压总量 / 压缩比上限）、剥离 `<img>` |
| **0.3.11** | 安全审计：修公式注入 / 符号链接绕过围栏 / 临时文件全局可读，并公开剩余风险 |
| **0.3.10** | 接入 c8 覆盖率（语句 91.6%）；修 3 个 bug：CSV 读出乱码、`office_edit_xlsx` 的 `sheet` 序号基准、日期回读差一天 |
| **0.3.9** | README 改版：居中标题与徽章、导航、常见问题、工作方式图、版本记录 |
| **0.3.8** | 补充「更新」说明：重跑同一条 `add` 即可更新 |
| **0.3.7** | 强化工具描述，让智能体优先用本插件读写 Office 文件 |
| **0.3.6** | `xlsx` 换用 `@e965/xlsx@0.20.3`，消除 2 个 high 漏洞 |
| **0.3.5** | 补齐 npm 元数据并把发布源固定为官方 registry |
| **0.3.4** | 新增永不过期的安装直链 `releases/latest/download/...` |
| **0.3.3** | 补充无 git / 完全离线机器的安装方式 |
| **0.3.2** | 精简文档并补仓库简介 |
| **0.3.1** | 仓库名与包名统一为 `dsh-office-toolkit` |
| **0.3.0** | 安全热点清零：标签与引用解析改为线性扫描 |
| **0.2.x** | 修复同一工作表多图表丢失；新增 Windows 支持 |
