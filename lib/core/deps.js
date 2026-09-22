// Lazy loading of the heavy runtime dependencies.
//
// A DSH profile can end up with a partly installed dependency tree: a build
// script pnpm refused to run, a stale bundled `node_modules` left over from an
// offline install, or a scoped package that only landed in the profile root.
// Importing the heavy modules lazily keeps the plugin itself loadable, limits
// the damage to the tools that really need the missing package, and turns a
// bare `ERR_MODULE_NOT_FOUND` into an error that names the package and says how
// to repair the install. Without that, the agent only sees an opaque failure and
// starts improvising — writing TSV text into a `.xlsx` is how that ends.
import { createRequire, isBuiltin } from 'node:module';
import { OfficeError } from './util.js';

/** Install command (same stable link as the README) used in repair hints. */
export const INSTALL_COMMAND =
  'dsh plugin --profile web add https://github.com/cnkids/dsh-office-toolkit/releases/latest/download/dsh-office-toolkit.tgz';

/**
 * DSH 0.1.6-alpha.1 / .2 的 profile 模块解析缺陷(根因在宿主,不在本插件)。
 *
 * 解析器把 `process/` 这类「内建名 + 尾斜杠」说明符切掉子路径后没有重新判断内建,
 * 再对 `createRequire(parent).resolve.paths(name)` 的返回值做 `for...of` —— 内建名
 * 返回的是 null,于是抛 TypeError。`readable-stream`(ExcelJS 的依赖)内部刻意写的
 * `require('process/')` / `require('string_decoder/')`(尾斜杠是为了强制走 npm polyfill
 * 包而不是内建)正好命中,所以只有懒加载 ExcelJS 的 xlsx 路径会失败,csv / docx 不受影响。
 * 插件侧无法绕开(触发点在依赖内部),只能把误导性的「缺少依赖」换成指向宿主的说明。
 *
 * @see https://github.com/deepseek-ai/deepseek-harness/discussions/7377
 */
export const DSH_RESOLVER_ISSUE =
  'https://github.com/deepseek-ai/deepseek-harness/discussions/7377#discussion-10858263';

/** 宿主解析器抛出的 TypeError 里稳定出现的片段(变量名可能不同,只认方法名)。 */
const RESOLVER_TYPE_ERROR_MARK = 'resolve.paths is not a function';

/** Runtime dependencies, each with what stops working when it is missing. */
export const CORE_DEPS = [
  ['mammoth', '读取 .docx'],
  ['@wekanteam/exceljs', '读写 .xlsx'],
  ['docx', '生成 .docx'],
  ['docxtemplater', '填充 .docx 模板'],
  ['pizzip', '解压 / 打包 OOXML'],
  ['marked', 'Markdown 转换'],
  ['@e965/xlsx', '读写 .xls/.xlsb/.ods/.csv/.tsv'],
  ['word-extractor', '读取 .doc'],
];

const PURPOSE_BY_SPEC = new Map(CORE_DEPS);
const resolved = createRequire(import.meta.url);
const cache = new Map();

/** Specs that cannot be resolved right now; cheap, no module is executed. */
export function missingDeps() {
  const missing = [];
  for (const [spec] of CORE_DEPS) {
    try {
      resolved.resolve(spec);
    } catch {
      missing.push(spec);
    }
  }
  return missing;
}

/** Pull the package name out of a Node resolution error (no regex needed). */
function specFromError(message) {
  const text = String(message || '');
  for (const marker of ["Cannot find package '", "Cannot find module '"]) {
    const at = text.indexOf(marker);
    if (at === -1) continue;
    const end = text.indexOf("'", at + marker.length);
    if (end > at) return text.slice(at + marker.length, end);
  }
  return '';
}

/** `process/` / `string_decoder/` 这类「内建名 + 尾斜杠」的说明符。 */
function isSlashBuiltinSpec(spec) {
  if (!spec.endsWith('/')) return false;
  const name = spec.slice(0, -1);
  return name.length > 0 && isBuiltin(name);
}

/** 是否命中 DSH 0.1.6-alpha 的解析器缺陷(而不是本插件缺依赖)。 */
function isHostResolverBug(err, spec) {
  if (String(err?.message || '').includes(RESOLVER_TYPE_ERROR_MARK)) return true;
  return isSlashBuiltinSpec(spec);
}

/** 把宿主解析器故障写成明确的「不是你的依赖问题」,并指向 DSH 侧报告。 */
function resolverBugError(err, what) {
  return new OfficeError(
    '这是 DSH 0.1.6-alpha.1 / .2 的模块解析缺陷,不是本插件缺少依赖 —— 重装依赖、重启 dsh web 都不会解决' +
      `(受影响功能:${what})。\n` +
      "触发点:ExcelJS 的依赖 readable-stream 内部会 require('process/') / require('string_decoder/')" +
      '(尾斜杠是为了走 npm polyfill 包),DSH 的 profile 解析器把它当成内建名、' +
      '再对 resolve.paths() 返回的 null 做 for...of,于是抛错;只影响走 ExcelJS 的 xlsx 路径(.csv / .docx 正常)。' +
      '插件侧无法绕开(触发点在依赖内部),请升级到修复版 DSH。\n' +
      `DSH 侧报告与补丁建议:${DSH_RESOLVER_ISSUE}\n` +
      `原始报错: ${err.message}`,
    'DSH_RESOLVER_BUG'
  );
}

/** Turn a failed import into a repair-oriented OfficeError. */
export function depFailure(err, purpose) {
  const found = specFromError(err?.message);
  if (isHostResolverBug(err, found)) {
    return resolverBugError(err, PURPOSE_BY_SPEC.get(found) || purpose || '该功能');
  }
  const code = err?.code;
  if (code !== 'ERR_MODULE_NOT_FOUND' && code !== 'MODULE_NOT_FOUND') return err;
  const spec = found || '未知包';
  const what = PURPOSE_BY_SPEC.get(spec) || purpose || '该功能';
  return new OfficeError(
    `缺少依赖 ${spec}(用于:${what})。依赖安装不完整,请重跑下面这条命令补全,然后重启 dsh web:\n${INSTALL_COMMAND}\n原始报错: ${err.message}`,
    'MISSING_DEPENDENCY'
  );
}

/** Import a runtime dependency, wrapping failures with repair advice. */
export async function loadCore(spec, purpose) {
  if (cache.has(spec)) return cache.get(spec);
  const pending = import(spec).then(
    (mod) => mod,
    (err) => {
      throw depFailure(err, purpose);
    }
  );
  cache.set(spec, pending);
  return pending;
}

/**
 * A lazily imported module namespace with an explicit member list. Each member
 * returns the real function's promise, so plain `await mod.fn(...)` call sites
 * keep working while a missing dependency only breaks the tools that use it.
 * An explicit list (rather than a Proxy) keeps the shape obvious and typed.
 */
export function lazyModule(spec, purpose, members) {
  const ns = {};
  for (const name of members) {
    ns[name] = async (...args) => {
      const mod = await loadCore(spec, purpose);
      return mod[name](...args);
    };
  }
  return ns;
}
