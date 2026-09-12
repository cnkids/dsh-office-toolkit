// 插件版本:集中一处,便于启动日志、报错诊断等处统一引用。
// 单独放一个文件是为了避免 core 层反向依赖宿主适配层(lib/index.js)。
import { createRequire } from 'node:module';

const pkg = createRequire(import.meta.url)('../../package.json');

/** 当前插件版本,例如 `0.3.20`。 */
export const PLUGIN_VERSION = pkg.version;
