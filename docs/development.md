# 开发与发版

[← 返回 README](../README.md)

## 本地开发

```sh
npm test           # 四个测试脚本：核心库 65 · 跨平台层 16 · 计算层 39 · 插件适配层 42，共 162 项
npm run coverage   # 同上并统计覆盖率，写出 coverage/lcov.info
```

纯 Node 脚本，不需要测试框架。覆盖率（c8）：**语句 94.3% / 分支 76.7% / 函数 95.5%**；未覆盖的主要是 `legacy-external.js` 的 Word COM / LibreOffice 分支与 `converters.js` 的纯 JS 回退 —— 它们只在没有 textutil / LibreOffice 的机器上才会走到。

`link:` 安装改完源码重启 `dsh web` 即生效；`file:` / `github:` 安装需重跑一次 `add` 刷新副本（HMR 不监听插件源码）。

## 代码质量

走 SonarQube（项目 `dsh-office-toll`；`sonar-project.properties` 不入库）：

```sh
export SONAR_TOKEN=<token> && sonar-scanner   # 会读取 coverage/lcov.info
```

当前 0 缺陷 / 0 漏洞 / 0 代码异味 / 0 安全热点，整体覆盖率 86.7%、新代码覆盖率 93.9%（门槛 80%），质量门通过。认知复杂度按 SonarQube 推荐阈值控制在 15 以内。

## 发布到 npm（维护者）

```sh
npm login --registry https://registry.npmjs.org/   # 首次，需要 npm 账号 + 2FA
npm publish
```

`publishConfig.registry` 已固定为官方源，所以**不受 `~/.npmrc` 国内镜像影响**（镜像只能读不能发）。发布后裸包名 `dsh plugin --profile web add dsh-office-toolkit` 即可用。

## 发版（维护者）

推 `v*` tag 即自动完成 —— `.github/workflows/release.yml` 会：校验 tag 与 `package.json` 版本一致 → 跑三套测试与依赖完整性校验 → 发布到 npm（Trusted Publishing / OIDC，**无需任何 token**）→ 建 Release 并上传 4 个附件。

```sh
# 1. 改 package.json 版本号,并同步 README「版本记录」与 docs/changelog.md
# 2. 本地先过一遍:npm test && npm run verify:offline && sonar-scanner
git commit -am "vX.Y.Z <类型>: <说明>"
git tag -a vX.Y.Z -m "vX.Y.Z ..."
git push origin main vX.Y.Z     # 推 tag 触发 Actions
```

**npm 侧已配置完毕**：Trusted Publisher 为 `cnkids` / `dsh-office-toolkit` / `release.yml`，发布走 OIDC，**不需要任何 token**（换仓库名或改 workflow 文件名时要同步改这里）。

发布后：`npm view dsh-office-toolkit version` 复核；用户侧安装与更新改为裸包名 `dsh plugin --profile web add dsh-office-toolkit` 与 `dsh plugin --profile web update`。

**离线包与 `.tgz` 也由 Actions 构建上传**（`git archive` + `npm install --omit=dev` + zip），本地无需再手工打包。

## 离线包完整性（发布门禁）

离线安装要「开箱即用」，就必须自带完整且**自包含**的依赖 —— 不能靠 profile 目录兜底，否则在别人机器上就是运行时缺包。把离线包解压到任意空目录后执行：

```sh
node test/offline-check.mjs <解压目录>     # 等价于 npm run verify:offline -- <目录>
node test/selftest.mjs && node test/plugin-smoke.mjs   # 再跑一遍真实读写
```

校验器会逐包核对：本包声明的每个依赖、以及包内每个依赖包声明的依赖，都必须解析到**包目录以内**；同时要求依赖树里没有任何 `preinstall`/`install`/`postinstall`（否则 `dsh plugin add` 会被 pnpm 的构建脚本审批打断）。任一条不满足即退出码非 0。

`.tgz` 线路同理：上传前用干净目录验证一次依赖能装全（`pnpm add <tgz>` 后不应出现缺包或构建脚本报错）。
