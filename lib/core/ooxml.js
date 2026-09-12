// OOXML containers with non-standard entry names.
//
// Word's own files use forward slashes (`word/document.xml`), but some
// generators — and anything repacked on Windows by hand or by a third-party
// tool — write backslashes (`word\document.xml`). mammoth and ExcelJS then look
// up `word/document.xml`, find nothing, and report the misleading
// "Could not find main document part. Are you sure this is a valid .docx file?".
//
// openOoxml() detects that and rebuilds the archive with normalized names, so
// such documents read like any other.
import PizZip from 'pizzip';
import { OfficeError } from './util.js';

/** `word\document.xml`, `./word//document.xml` -> `word/document.xml`. */
export function normalizeEntryName(name) {
  const slashed = String(name).replaceAll('\\', '/');
  const isDir = slashed.endsWith('/');
  const parts = slashed.split('/').filter((part) => part && part !== '.');
  const joined = parts.join('/');
  return isDir && joined ? `${joined}/` : joined;
}

/**
 * Open an OOXML zip, repairing non-standard entry names when needed.
 * @param buf archive bytes
 * @param label used in error messages (e.g. `docx`)
 * @returns { zip, buffer, repaired } — `buffer` is the repaired archive, or the
 *   original bytes when nothing had to change; `repaired` lists the renames.
 */
export async function openOoxml(buf, label = 'OOXML') {
  const bytes = Buffer.isBuffer(buf) ? buf : Buffer.from(buf || []);
  let zip;
  try {
    zip = new PizZip(bytes);
  } catch (err) {
    throw new OfficeError(`不是有效的 ${label}（无法解压）: ${err?.message || err}`, 'BAD_ZIP');
  }
  const renames = [];
  for (const name of Object.keys(zip.files)) {
    const fixed = normalizeEntryName(name);
    if (fixed && fixed !== name) renames.push([name, fixed]);
  }
  if (!renames.length) return { zip, buffer: bytes, repaired: [] };

  const out = new PizZip();
  for (const [name, entry] of Object.entries(zip.files)) {
    const fixed = normalizeEntryName(name);
    if (!fixed) continue;
    if (entry.dir) out.folder(fixed);
    else out.file(fixed, entry.asNodeBuffer(), { binary: true });
  }
  return {
    zip: out,
    buffer: out.generate({ type: 'nodebuffer', compression: 'DEFLATE' }),
    repaired: renames.map(([from, to]) => `${from} → ${to}`),
  };
}

/** Look up a part by its standard name, ignoring case as a last resort. */
export function partOf(zip, path) {
  const direct = zip.file(path);
  if (direct) return direct;
  const wanted = normalizeEntryName(path).toLowerCase();
  for (const name of Object.keys(zip.files)) {
    if (normalizeEntryName(name).toLowerCase() === wanted) return zip.files[name];
  }
  return null;
}
