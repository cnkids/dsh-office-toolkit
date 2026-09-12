// Path containment guard used to fence plugin writes to the session workspace.
// Windows semantics are case-insensitive and separator-aware; the platform can
// be injected so the logic is unit-testable on any OS.
import { posix, win32 } from 'node:path';

function normalizePath(api, value, foldCase) {
  const resolved = api.resolve(String(value));
  return foldCase ? resolved.toLowerCase() : resolved;
}

/**
 * Whether `target` is `root` itself or lives under it.
 * @param {string} target candidate path
 * @param {string} root allowed root
 * @param {{platform?: string}} [options] platform override (defaults to process.platform)
 */
export function isPathInside(target, root, options = {}) {
  const platform = options.platform || process.platform;
  const api = platform === 'win32' ? win32 : posix;
  const foldCase = platform === 'win32';
  const a = normalizePath(api, target, foldCase);
  const b = normalizePath(api, root, foldCase);
  const withSep = b.endsWith(api.sep) ? b : b + api.sep;
  return a === b || a.startsWith(withSep);
}

/** Whether `target` is inside any of the given roots. */
export function isInsideAny(target, roots, options = {}) {
  return roots.some((root) => isPathInside(target, root, options));
}
