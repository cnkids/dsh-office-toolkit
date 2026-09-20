// PDF 读取(只抽文本层):优先借本机能力,兜底用随包的 pdfjs。
//
// 后端顺序:pdftotext(poppler) → macOS 自带 PDFKit → 内置 pdfjs。
// 都是「能读出字」的兜底链:任一后端失败就换下一个,不因为缺工具而读不出来。
//
// 边界(如实说明):只保证文本与页数 —— 版式、表格、图片不可靠,多栏可能串行;
// 扫描件是图片、没有文本层,需要 OCR,这里不做。内置 pdfjs 是 vendored 的
// Apache-2.0 代码(见 vendor/pdfjs/LICENSE),不联网、不解析 JS、不用系统字体。
import { execFile, spawn } from 'node:child_process';
import { promisify } from 'node:util';
import { access } from 'node:fs/promises';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { dirname, join } from 'node:path';
import { OfficeError, CAPS, isWithinBytes } from './util.js';

const execFileP = promisify(execFile);
const IS_MAC = process.platform === 'darwin';
const VENDOR = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'vendor', 'pdfjs');
const PDF_MAGIC = '%PDF-';
const MAX_PAGES = 5000;
/** 单次抽取的文本上限:PDF 正文比 HTML 输入长,给 4 倍 textutil 的量级。 */
const MAX_TEXT_CHARS = CAPS.MAX_TEXT_BYTES * 4;

/** 把字节喂给子进程的 stdin 并收 stdout(execFile 没有 input 选项)。 */
function runWithInput(exe, args, input, { timeout = 120000, maxBuffer = 16 * 1024 * 1024 } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(exe, args, { stdio: ['pipe', 'pipe', 'pipe'] });
    let out = '';
    let err = '';
    let size = 0;
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`执行超时(${Math.round(timeout / 1000)}s)`));
    }, timeout);
    const finish = (fn, value) => {
      clearTimeout(timer);
      fn(value);
    };
    child.stdout.on('data', (chunk) => {
      size += chunk.length;
      if (size > maxBuffer) {
        child.kill('SIGKILL');
        finish(reject, new Error('输出超过上限'));
        return;
      }
      out += chunk;
    });
    child.stderr.on('data', (chunk) => { err += chunk; });
    child.on('error', (e) => finish(reject, e));
    child.on('close', (code) => {
      if (code === 0) finish(resolve, out);
      else finish(reject, new Error(String(err || `退出码 ${code}`)));
    });
    child.stdin.on('error', () => {});
    child.stdin.end(input);
  });
}

/** `%PDF-` 开头的字节(PDF 允许文件头前有少量垃圾字节,这里只看前 1 KiB)。 */
export function looksLikePdf(buf) {
  const bytes = Buffer.isBuffer(buf) ? buf : Buffer.from(buf ?? []);
  return bytes.subarray(0, 1024).includes(PDF_MAGIC);
}

/**
 * 只在固定的可信目录里找外部工具 —— 不扫 PATH。
 * PATH 可以被污染(攻击者往前面塞一个同名可执行文件),而这里缺工具也不会读不出来:
 * 内置 pdfjs 兜底,所以宁可少找一个后端也不引入 PATH 投毒面。
 */
const TRUSTED_BIN_DIRS = IS_MAC
  ? ['/opt/homebrew/bin', '/usr/local/bin', '/opt/local/bin', '/usr/bin', '/bin']
  : ['/usr/local/bin', '/usr/bin', '/bin', '/snap/bin'];

async function findExecutable(name) {
  for (const dir of TRUSTED_BIN_DIRS) {
    const full = join(dir, name);
    try {
      await access(full);
      return full;
    } catch {
      /* 继续找 */
    }
  }
  return '';
}

// ---------------------------------------------------------------------------
// 文本整理
// ---------------------------------------------------------------------------
/** 行尾空白手工裁掉:不对「空格串」跑无锚点正则(那种写法在长空格串上是 O(n²))。 */
function trimLineEnd(line) {
  let end = line.length;
  while (end > 0 && (line[end - 1] === ' ' || line[end - 1] === '\t')) end -= 1;
  return line.slice(0, end);
}

/** 统一换行、去掉 NUL 与行尾空白,并把连续空行压到一个空行。 */
export function normalizePdfText(raw) {
  const unified = String(raw ?? '')
    .replaceAll('\r\n', '\n')
    .replaceAll('\r', '\n')
    .replaceAll('\0', '');
  return unified.split('\n').map(trimLineEnd).join('\n').replaceAll(/\n{3,}/g, '\n\n').trim();
}

/** 每页文本 → 正文:多页时插入页码标记,便于按页定位。 */
export function joinPages(pages) {
  const kept = pages.map((text) => normalizePdfText(text));
  if (kept.length <= 1) return kept[0] || '';
  return kept
    .map((text, index) => `--- 第 ${index + 1} 页 ---\n${text}`)
    .join('\n\n');
}

function clampText(text) {
  if (text.length <= MAX_TEXT_CHARS) return { text, truncated: false };
  return { text: text.slice(0, MAX_TEXT_CHARS), truncated: true };
}

// ---------------------------------------------------------------------------
// 后端:pdftotext(poppler)
// ---------------------------------------------------------------------------
async function viaPdftotext(buf, exe) {
  // `-` 表示从 stdin 读、往 stdout 写:不落临时文件;默认按阅读顺序抽,比 -layout 干净
  const stdout = await runWithInput(exe, ['-enc', 'UTF-8', '-', '-'], buf, {
    maxBuffer: MAX_TEXT_CHARS * 4,
  });
  return { pages: splitFormFeeds(String(stdout)), backend: 'pdftotext' };
}

/** poppler 用换页符分页,末尾那个多余的换页符不算一页。 */
function splitFormFeeds(raw) {
  const pages = String(raw).split('\f');
  while (pages.length > 1 && !normalizePdfText(pages.at(-1))) pages.pop();
  return pages;
}

// ---------------------------------------------------------------------------
// 后端:macOS 自带 PDFKit(经 osascript 的 ObjC 桥,零安装)
// ---------------------------------------------------------------------------
const PDFKIT_SCRIPT = [
  'ObjC.import("Quartz");',
  'function run(argv) {',
  '  const url = $.NSURL.fileURLWithPath(argv[0]);',
  '  const doc = $.PDFDocument.alloc.initWithURL(url);',
  '  if (!doc.js || doc.isLocked) return JSON.stringify({ error: "打开失败或已加密" });',
  '  const pages = [];',
  '  const count = doc.pageCount;',
  '  for (let i = 0; i < count; i++) pages.push(doc.pageAtIndex(i).string.js || "");',
  '  return JSON.stringify({ pages: pages });',
  '}',
].join('\n');

async function viaPdfKit(buf, path) {
  if (!path) throw new Error('PDFKit 需要文件路径');
  const { stdout } = await execFileP('/usr/bin/osascript', ['-l', 'JavaScript', '-e', PDFKIT_SCRIPT, '--', path], {
    timeout: 180000,
    maxBuffer: MAX_TEXT_CHARS * 8,
  });
  const parsed = JSON.parse(String(stdout).trim());
  if (parsed.error) throw new Error(parsed.error);
  return { pages: parsed.pages || [], backend: 'macOS PDFKit' };
}

// ---------------------------------------------------------------------------
// 后端:内置 pdfjs(随包携带,离线可用)
// ---------------------------------------------------------------------------
let pdfjsPromise = null;

/**
 * pdfjs 的 Node 构建在 import 时会为「渲染」缺 canvas / DOMMatrix 打三条警告 ——
 * 我们只抽文本、不渲染,临时静音(import 完立刻恢复),免得插件日志里出现误导性警告。
 */
async function importPdfjs() {
  const warn = console.warn;
  console.warn = () => {};
  try {
    // Windows 下不能直接 import 盘符路径,要转成 file:// URL
    return await import(pathToFileURL(join(VENDOR, 'pdf.min.mjs')).href);
  } finally {
    console.warn = warn;
  }
}

function loadPdfjs() {
  pdfjsPromise = pdfjsPromise || importPdfjs();
  return pdfjsPromise;
}

async function viaPdfjs(buf) {
  const pdfjs = await loadPdfjs();
  pdfjs.GlobalWorkerOptions.workerSrc = join(VENDOR, 'pdf.worker.min.mjs');
  const task = pdfjs.getDocument({
    data: new Uint8Array(buf),
    // 不解析 PDF 内嵌 JS、不读系统字体、不做字体渲染:只抽文本,攻击面最小
    isEvalSupported: false,
    useSystemFonts: false,
    disableFontFace: true,
    cMapUrl: join(VENDOR, 'cmaps') + '/',
    cMapPacked: true,
    standardFontDataUrl: join(VENDOR, 'standard_fonts') + '/',
    verbosity: 0,
  });
  const doc = await task.promise;
  const pages = [];
  const count = Math.min(doc.numPages, MAX_PAGES);
  for (let index = 1; index <= count; index += 1) {
    const page = await doc.getPage(index);
    const content = await page.getTextContent();
    pages.push(content.items.map((item) => `${item.str}${item.hasEOL ? '\n' : ''}`).join(''));
  }
  return { pages, backend: '内置 pdfjs', totalPages: doc.numPages };
}

// ---------------------------------------------------------------------------
// 入口
// ---------------------------------------------------------------------------
function encryptedError(err) {
  const message = String(err?.message || err);
  return /password|encrypted|encryption|isLocked/i.test(message);
}

/** 校验并规范化入参:大小上限与 PDF 魔数。 */
function pdfBytesOf(buf) {
  const bytes = Buffer.isBuffer(buf) ? buf : Buffer.from(buf ?? []);
  if (!isWithinBytes(bytes, CAPS.MAX_EXCEL_INPUT_BYTES)) {
    throw new OfficeError('PDF 超过 60 MB 上限', 'OFFICE_TOO_LARGE');
  }
  if (!looksLikePdf(bytes)) {
    throw new OfficeError('这不是有效的 PDF（文件头没有 %PDF-）—— 可能被改过后缀，或文件损坏', 'BAD_PDF');
  }
  return bytes;
}

/**
 * 后端执行顺序。backend 只用于测试与排查:auto 是正常路径,
 * pdfjs / pdftotext / pdfkit 各自强制单一后端。
 */
async function backendPlan(bytes, { path, backend }) {
  const plan = [];
  const poppler = backend === 'pdfjs' || backend === 'pdfkit' ? '' : await findExecutable('pdftotext');
  if (poppler) plan.push(['pdftotext', () => viaPdftotext(bytes, poppler)]);
  if (backend !== 'pdfjs' && backend !== 'pdftotext' && IS_MAC && path) plan.push(['macOS PDFKit', () => viaPdfKit(bytes, path)]);
  if (backend !== 'pdftotext' && backend !== 'pdfkit') plan.push(['内置 pdfjs', () => viaPdfjs(bytes)]);
  return plan;
}

/** 跑一个后端:成功返回结果,失败把原因记进 attempts 后返回 null(加密则直接抛)。 */
async function tryBackend(label, run, attempts, messages) {
  try {
    const result = await run();
    const { text, truncated } = clampText(joinPages(result.pages));
    if (!text) throw new Error('没有抽取到文本（可能是扫描件，需要 OCR）');
    if (truncated) messages.push(`文本超过 ${MAX_TEXT_CHARS} 字符，已截断`);
    if (attempts.length) messages.push(`前面的后端不可用(${attempts.join('、')})，已改用 ${result.backend}`);
    return { text, pages: Number(result.totalPages) || result.pages.length, backend: result.backend, messages, truncated };
  } catch (err) {
    if (encryptedError(err)) {
      throw new OfficeError(`这个 PDF 已加密或有打开密码，读不出文本: ${String(err?.message || err).slice(0, 120)}`, 'PDF_ENCRYPTED');
    }
    attempts.push(`${label}: ${String(err?.stderr || err?.message || err).slice(0, 80)}`);
    return null;
  }
}

/**
 * 抽取 PDF 文本。
 * @param {Buffer} buf
 * @param {{ path?: string, backend?: string }} options path 用于系统级后端(缺省时只能走纯 JS)
 * @returns {Promise<{ text: string, pages: number, backend: string, messages: string[], truncated: boolean }>}
 */
export async function readPdfText(buf, { path, backend = 'auto' } = {}) {
  const bytes = pdfBytesOf(buf);
  const attempts = [];
  const messages = [];
  for (const [label, run] of await backendPlan(bytes, { path, backend })) {
    const outcome = await tryBackend(label, run, attempts, messages);
    if (outcome) return outcome;
  }
  throw new OfficeError(`PDF 文本抽取失败。已尝试: ${attempts.join(' | ')}`, 'PDF_READ_FAILED');
}

/** 供 office_read 的提示文案复用:这台机器上哪些后端可用。 */
export async function pdfBackendHint() {
  const poppler = await findExecutable('pdftotext');
  const parts = [];
  if (poppler) parts.push('pdftotext');
  if (IS_MAC) parts.push('macOS PDFKit');
  parts.push('内置 pdfjs');
  return parts.join(' → ');
}
