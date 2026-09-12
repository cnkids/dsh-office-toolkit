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
import { createRequire } from 'node:module';
import { OfficeError } from './util.js';

/** Install command (same stable link as the README) used in repair hints. */
export const INSTALL_COMMAND =
  'dsh plugin --profile web add https://github.com/cnkids/dsh-office-toolkit/releases/latest/download/dsh-office-toolkit.tgz';

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

/** Turn a failed import into a repair-oriented OfficeError. */
export function depFailure(err, purpose) {
  const code = err?.code;
  if (code !== 'ERR_MODULE_NOT_FOUND' && code !== 'MODULE_NOT_FOUND') return err;
  const spec = specFromError(err.message) || '未知包';
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
