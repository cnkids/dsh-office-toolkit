// Offline-package completeness check.
//
// Guarantees the "out of the box" promise for a self-contained package (the
// `-offline.zip`): every declared dependency — and every dependency of every
// package shipped inside it — must resolve to something *inside the package
// directory*. A resolution that lands outside means the package would silently
// borrow a module from the surrounding profile, which is how a half-installed
// plugin ends up failing at runtime on someone else's machine.
//
// It also asserts that no shipped package runs an install script, so
// `dsh plugin add` cannot be interrupted by pnpm's build-script approval.
//
// Usage: node test/offline-check.mjs [packageDir]   (default: repo root)
import { readFileSync, readdirSync, existsSync, realpathSync } from 'node:fs';
import { builtinModules } from 'node:module';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(process.argv[2] || join(here, '..'));
// 用真实路径比较:macOS 上 /tmp 是 /private/tmp 的符号链接,否则全是误判
const rootReal = realpathSync(root);
const INSTALL_SCRIPTS = ['preinstall', 'install', 'postinstall'];
const BUILTINS = new Set([...builtinModules, ...builtinModules.map((m) => `node:${m}`)]);

function readJson(path) {
  try {
    return JSON.parse(readFileSync(path, 'utf8'));
  } catch {
    return null;
  }
}

function listDirs(dir) {
  try {
    return readdirSync(dir, { withFileTypes: true }).filter((e) => e.isDirectory() && !e.name.startsWith('.'));
  } catch {
    return [];
  }
}

/** Every package directory below `root/node_modules`, nested trees included. */
function collectPackages(packageRoot) {
  const out = [];
  const walkModules = (modulesDir) => {
    for (const entry of listDirs(modulesDir)) {
      if (entry.name.startsWith('@')) {
        for (const sub of listDirs(join(modulesDir, entry.name))) {
          const pkgDir = join(modulesDir, entry.name, sub.name);
          out.push(pkgDir);
          walkModules(join(pkgDir, 'node_modules'));
        }
        continue;
      }
      const pkgDir = join(modulesDir, entry.name);
      out.push(pkgDir);
      walkModules(join(pkgDir, 'node_modules'));
    }
  };
  walkModules(join(packageRoot, 'node_modules'));
  return out;
}

/**
 * Is `spec` installed for `fromDir`? Walks `node_modules` upwards exactly like
 * Node does, but checks for the package *directory* rather than resolving an
 * entry file — type-only packages such as `@types/node` ship no entry at all.
 */
function findPackageDir(spec, fromDir) {
  let dir = fromDir;
  for (;;) {
    const candidate = join(dir, 'node_modules', spec);
    if (existsSync(join(candidate, 'package.json'))) return candidate;
    const parent = dirname(dir);
    if (parent === dir) return null;
    dir = parent;
  }
}

const inside = (target) => {
  let real;
  try {
    real = realpathSync(target);
  } catch {
    return false;
  }
  const rel = relative(rootReal, real);
  return rel === '' || (!rel.startsWith('..' + sep) && rel !== '..');
};

const problems = [];
const notes = [];
const scriptHits = [];
let checked = 0;

function checkPackage(pkgDir, label) {
  const manifest = readJson(join(pkgDir, 'package.json'));
  if (!manifest) return;
  const scripts = manifest.scripts || {};
  for (const key of INSTALL_SCRIPTS) {
    if (scripts[key]) scriptHits.push(`${manifest.name || label} → ${key}: ${scripts[key]}`);
  }
  for (const spec of Object.keys(manifest.dependencies || {})) {
    if (BUILTINS.has(spec)) continue; // Node 内置模块,无需安装
    checked += 1;
    const hit = findPackageDir(spec, pkgDir);
    if (!hit) {
      problems.push(`${label} 缺少依赖 ${spec}`);
      continue;
    }
    if (!inside(hit)) problems.push(`${label} 的依赖 ${spec} 落在了包外(靠 profile 兜底): ${relative(rootReal, hit) || hit}`);
  }
  for (const spec of Object.keys(manifest.optionalDependencies || {})) {
    const hit = findPackageDir(spec, pkgDir);
    if (!hit) notes.push(`${label} 未装可选依赖 ${spec}(平台相关)`);
    else if (!inside(hit)) notes.push(`${label} 的可选依赖 ${spec} 在包外`);
  }
}

console.log(`检查目录: ${root}`);
const rootManifest = readJson(join(root, 'package.json'));
if (!rootManifest) {
  console.error(`✗ ${root} 下没有 package.json`);
  process.exit(1);
}
console.log(`包: ${rootManifest.name}@${rootManifest.version}`);

if (!existsSync(join(root, 'node_modules'))) {
  console.error('✗ 没有 node_modules —— 离线包必须自带完整依赖');
  process.exit(1);
}

checkPackage(root, '本包');
const packages = collectPackages(root);
for (const pkgDir of packages) checkPackage(pkgDir, relative(root, pkgDir));

console.log(`顶层依赖检查: ${checked} 处,随包依赖包: ${packages.length} 个`);
if (notes.length) console.log(`可选依赖提示: ${notes.length} 条`);
if (scriptHits.length) {
  console.log('\n✗ 依赖树里存在安装脚本(会让 dsh plugin add 被 pnpm 打断):');
  for (const hit of scriptHits) console.log(`  - ${hit}`);
}
if (problems.length) {
  console.log('\n✗ 依赖不完整:');
  for (const p of problems.slice(0, 40)) console.log(`  - ${p}`);
  if (problems.length > 40) console.log(`  ...还有 ${problems.length - 40} 条`);
}
const ok = !problems.length && !scriptHits.length;
console.log(ok ? '\n✅ 依赖完整且自包含,可离线开箱即用' : `\n❌ 未通过:${problems.length} 个缺失/越界,${scriptHits.length} 个安装脚本`);
process.exit(ok ? 0 : 1);
