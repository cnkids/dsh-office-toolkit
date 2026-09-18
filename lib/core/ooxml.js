// OOXML containers with non-standard entry names.
//
// Word's own files use lowercase, forward-slash part names (`word/document.xml`),
// but real-world files break that in several ways:
//   * backslashes    — `word\document.xml` (third-party generators, repacking)
//   * different case — `Word/Document.xml` (Windows is case-insensitive, so tools
//                      there happily write it and Word still opens it)
//   * mixed          — `Word\Document.XML`
// mammoth and ExcelJS look up `word/document.xml` literally, find nothing, and
// report the misleading "Could not find main document part. Are you sure this is
// a valid .docx file?".
//
// openOoxml() normalizes the entry names — by patching the zip's name fields in
// place, without decompressing any entry — and rewrites the handful of small
// parts that reference those names (`[Content_Types].xml`, `*.rels`). A rename
// without the matching reference rewrite would leave the file *more* broken than
// before, because the two sides stop agreeing.
import PizZip from 'pizzip';
import { OfficeError } from './util.js';
import { attrIn, readTagAt } from './markup.js';

/** Canonical spelling of the OOXML parts that get referenced by name. */
const CANONICAL_PARTS = new Map([
  ['[content_types].xml', '[Content_Types].xml'],
  ['_rels/.rels', '_rels/.rels'],
  ['word/document.xml', 'word/document.xml'],
  ['word/styles.xml', 'word/styles.xml'],
  ['word/settings.xml', 'word/settings.xml'],
  ['word/numbering.xml', 'word/numbering.xml'],
  ['word/fonttable.xml', 'word/fontTable.xml'],
  ['word/websettings.xml', 'word/webSettings.xml'],
  ['word/theme/theme1.xml', 'word/theme/theme1.xml'],
  ['word/_rels/document.xml.rels', 'word/_rels/document.xml.rels'],
  ['word/footnotes.xml', 'word/footnotes.xml'],
  ['word/endnotes.xml', 'word/endnotes.xml'],
  ['word/comments.xml', 'word/comments.xml'],
  ['docprops/core.xml', 'docProps/core.xml'],
  ['docprops/app.xml', 'docProps/app.xml'],
  ['docprops/custom.xml', 'docProps/custom.xml'],
  ['xl/workbook.xml', 'xl/workbook.xml'],
  ['xl/styles.xml', 'xl/styles.xml'],
  ['xl/sharedstrings.xml', 'xl/sharedStrings.xml'],
  ['xl/_rels/workbook.xml.rels', 'xl/_rels/workbook.xml.rels'],
]);

/** 关系/内容类型里引用部件时可能出现的四种写法。 */
const SEPARATOR_VARIANTS = [
  (path) => path,
  (path) => `/${path}`,
  (path) => path.replaceAll('/', '\\'),
  (path) => `/${path.replaceAll('/', '\\')}`,
];

/** 大小写不敏感的整体替换(不用正则,线性扫描)。 */
function replaceAllInsensitive(text, needle, replacement) {
  if (!needle) return text;
  const lowerText = text.toLowerCase();
  const lowerNeedle = needle.toLowerCase();
  let out = '';
  let from = 0;
  for (;;) {
    const at = lowerText.indexOf(lowerNeedle, from);
    if (at === -1) return out + text.slice(from);
    out += text.slice(from, at) + replacement;
    from = at + needle.length;
  }
}

/** `word\document.xml`, `./Word//Document.XML` -> `word/document.xml`. */
export function normalizeEntryName(name) {
  const slashed = String(name).replaceAll('\\', '/');
  const isDir = slashed.endsWith('/');
  const parts = slashed.split('/').filter((part) => part && part !== '.');
  const joined = parts.join('/');
  const canonical = CANONICAL_PARTS.get(joined.toLowerCase());
  const fixed = canonical || joined;
  return isDir && fixed ? `${fixed}/` : fixed;
}

/** 把关系/内容类型里指向旧名字的引用改写成新名字。 */
function rewriteReferences(xml, renames) {
  let out = String(xml);
  for (const [from, to] of renames) {
    for (const makeVariant of SEPARATOR_VARIANTS) {
      const variant = makeVariant(from);
      if (variant.toLowerCase() === to.toLowerCase() || variant === `/${to}`) continue;
      out = replaceAllInsensitive(out, variant, to);
    }
  }
  return out;
}

/** 需要改名的条目(旧名 → 规范名)。 */
function listRenames(zip) {
  const renames = [];
  for (const name of Object.keys(zip.files)) {
    const fixed = normalizeEntryName(name);
    if (fixed && fixed !== name) renames.push([name, fixed]);
  }
  return renames;
}

// ---------------------------------------------------------------------------
// zip 名字字段的原地改写(不解压任何条目)
// ---------------------------------------------------------------------------
const LOCAL_SIG = [0x50, 0x4b, 0x03, 0x04];
const CENTRAL_SIG = [0x50, 0x4b, 0x01, 0x02];
const EOCD_SIG = [0x50, 0x4b, 0x05, 0x06];
const MAX_COMMENT = 65535;

function signatureAt(buf, at, sig) {
  return at >= 0 && buf[at] === sig[0] && buf[at + 1] === sig[1] && buf[at + 2] === sig[2] && buf[at + 3] === sig[3];
}

function u16(buf, at) {
  return buf[at] | (buf[at + 1] << 8);
}

function u32(buf, at) {
  return (buf[at] | (buf[at + 1] << 8) | (buf[at + 2] << 16) | (buf[at + 3] << 24)) >>> 0;
}

/** 从尾部找中央目录结束记录(允许 64KB 注释)。 */
function findEocd(buf) {
  const from = Math.max(0, buf.length - 22 - MAX_COMMENT);
  for (let at = buf.length - 22; at >= from; at -= 1) {
    if (signatureAt(buf, at, EOCD_SIG)) return at;
  }
  return -1;
}

/**
 * 原地改写条目名:只动中央目录与本地文件头的名字字段,不碰任何数据。
 * 仅适用于新旧名字字节数相同的情况 —— 反斜杠→斜杠、大小写归一化都满足,
 * 而它们正是实际遇到的两种。
 * @returns 改写后的新 Buffer,或 null(结构异常 / 长度会变 → 交给其它路径)
 */
function patchEntryNames(buf, renameMap) {
  const out = Buffer.from(buf);
  const eocd = findEocd(out);
  if (eocd === -1) return null;
  const count = u16(out, eocd + 10);
  let at = u32(out, eocd + 16);
  for (let i = 0; i < count; i += 1) {
    if (!signatureAt(out, at, CENTRAL_SIG)) return null;
    const nameLen = u16(out, at + 28);
    const extraLen = u16(out, at + 30);
    const commentLen = u16(out, at + 32);
    const localAt = u32(out, at + 42);
    const name = out.subarray(at + 46, at + 46 + nameLen).toString('utf8');
    const fixed = renameMap.get(name);
    if (fixed && Buffer.byteLength(fixed) === nameLen) {
      out.write(fixed, at + 46, 'utf8');
      if (signatureAt(out, localAt, LOCAL_SIG) && u16(out, localAt + 26) === nameLen) {
        out.write(fixed, localAt + 30, 'utf8');
      }
    }
    at += 46 + nameLen + extraLen + commentLen;
  }
  return out;
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

/** 引用部件(内容类型 + 关系)都很小,单独取出来改写即可。 */
function rewriteReferenceParts(zip, renames) {
  const targets = ['[Content_Types].xml', '_rels/.rels', 'word/_rels/document.xml.rels', 'xl/_rels/workbook.xml.rels'];
  let touched = 0;
  for (const path of targets) {
    const entry = partOf(zip, path);
    if (!entry) continue;
    const before = entry.asText();
    const after = rewriteReferences(before, renames);
    if (after !== before) {
      zip.file(path, after);
      touched += 1;
    }
  }
  return touched;
}

/**
 * 找出一份 OOXML 的主部件并判断家族。
 * 顺序与真正的读取器一致:包关系 `_rels/.rels` 里的 officeDocument 关系 → 规范名
 * → 以已知主部件文件名结尾的任何部件(容忍奇怪的目录名与大小写)。
 * @returns { kind, path, entry } 或 null;kind 为 docx / xlsx / odt
 */
export function officePartOf(zip) {
  const byTarget = officeTargetFromRels(zip);
  if (byTarget) {
    const entry = partOf(zip, byTarget);
    const kind = mainKindByName(byTarget) || kindByDirectory(byTarget);
    if (entry && kind) return { kind: odfKindOf(zip, kind), path: byTarget, entry };
  }
  for (const [name, kind] of KNOWN_MAIN_PARTS) {
    const entry = partOf(zip, name);
    if (entry) return { kind: odfKindOf(zip, kind), path: name, entry };
  }
  for (const name of Object.keys(zip.files)) {
    const entry = zip.files[name];
    if (entry.dir) continue;
    const kind = mainKindByName(name);
    if (kind) return { kind: odfKindOf(zip, kind), path: name, entry };
  }
  return null;
}

/**
 * `.odt` 与 `.ods` 的主部件都叫 content.xml,只看文件名会把表格读成文档 —— 靠 mimetype
 * 区分(没有 mimetype 就退回 manifest.xml 里的声明)。
 */
function odfKindOf(zip, kind) {
  if (kind !== 'odt') return kind;
  const declared = String(partOf(zip, 'mimetype')?.asText() || '').trim().toLowerCase();
  const hint = declared || String(partOf(zip, 'META-INF/manifest.xml')?.asText() || '').toLowerCase();
  return hint.includes('spreadsheet') ? 'ods' : kind;
}

/**
 * `_rels/.rels` 里 officeDocument 关系的 Target(规范化后的相对路径)。
 * 逐个 <Relationship> 看 Type 是否以 /officeDocument 结尾 —— 不能用 indexOf('officeDocument'),
 * 那会先撞上 extended-properties 的 `officeDocument/2006/relationships/...`,拿错 Target
 * (`.xlsb` 就是这样被误判成损坏文件的)。
 */
function officeTargetFromRels(zip) {
  const rels = partOf(zip, '_rels/.rels');
  if (!rels) return '';
  const xml = rels.asText();
  let at = 0;
  while (at < xml.length) {
    const start = xml.indexOf('<Relationship', at);
    if (start === -1) return '';
    const tag = readTagAt(xml, start);
    if (tag?.name !== 'Relationship') {
      at = start + 1;
      continue;
    }
    if (String(attrIn(tag.attrs, 'Type') || '').endsWith('/officeDocument')) {
      const raw = String(attrIn(tag.attrs, 'Target') || '').replaceAll('\\', '/');
      let begin = 0;
      while (raw[begin] === '.' || raw[begin] === '/') begin += 1;
      return normalizeEntryName(raw.slice(begin));
    }
    at = tag.end;
  }
  return '';
}

const KNOWN_MAIN_PARTS = [
  ['word/document.xml', 'docx'],
  ['xl/workbook.xml', 'xlsx'],
  ['xl/workbook.bin', 'xlsb'],
  ['content.xml', 'odt'],
];

/** 按主部件文件名判断家族 —— 只认真正的名字,不会把 styles.xml 之类误当主部件。 */
function mainKindByName(path) {
  const lower = normalizeEntryName(path).toLowerCase();
  if (lower.endsWith('/document.xml')) return 'docx';
  if (lower.endsWith('/workbook.xml')) return 'xlsx';
  // 二进制工作簿(.xlsb):ExcelJS 读不了,要交给 SheetJS,所以家族单独标出来
  if (lower.endsWith('/workbook.bin')) return 'xlsb';
  if (lower === 'content.xml' || lower.endsWith('/content.xml')) return 'odt';
  return '';
}

/** 按目录判断家族(关系声明的 Target 可能是不常见的部件名)。 */
function kindByDirectory(path) {
  const lower = normalizeEntryName(path).toLowerCase();
  if (lower.startsWith('word/')) return 'docx';
  if (lower.startsWith('xl/')) return 'xlsx';
  return '';
}

/**
 * Open an OOXML zip, repairing non-standard entry names (and the parts that
 * reference them) when needed.
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
  const renames = listRenames(zip);
  if (!renames.length) return { zip, buffer: bytes, repaired: [] };

  const lengthChanging = renames.some(([from, to]) => Buffer.byteLength(from) !== Buffer.byteLength(to));
  if (lengthChanging) {
    // 少见(如 word//document.xml):名字长度会变,不做字节改写。
    // 读取侧仍有大小写/斜杠容忍的查找与内置兜底,不会读不出来。
    return { zip, buffer: bytes, repaired: [] };
  }
  const patched = patchEntryNames(bytes, new Map(renames));
  if (!patched) return { zip, buffer: bytes, repaired: [] };

  const fixed = new PizZip(patched);
  const rewritten = rewriteReferenceParts(fixed, renames);
  return {
    zip: fixed,
    buffer: rewritten ? fixed.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) : patched,
    repaired: renames.map(([from, to]) => `${from} → ${to}`),
  };
}
