// External legacy-Word converters, selected per platform:
//   macOS   textutil (built-in)   any OS  LibreOffice soffice   Windows  Word COM
// Every backend is optional; callers fall back to the pure-JS layer.
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { access, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';
import { OfficeError, CAPS, extOf } from './util.js';

const execFileP = promisify(execFile);
const IS_MAC = process.platform === 'darwin';
const IS_WIN = process.platform === 'win32';
const TEXTUTIL = '/usr/bin/textutil';

export const TEXTUTIL_TARGETS = ['txt', 'html', 'rtf', 'rtfd', 'doc', 'docx', 'odt', 'wordml'];
const SOFFICE_FILTERS = new Set(['pdf', 'doc', 'docx', 'rtf', 'odt', 'txt', 'html', 'xls', 'xlsx', 'csv']);

// Microsoft Word SaveAs format codes (wdFormat*).
export const WORD_COM_FORMATS = { doc: 0, txt: 2, rtf: 6, html: 8, docx: 16, odt: 23, pdf: 17 };

async function exists(path) {
  try { await access(path); return true; } catch { return false; }
}

/** Short, human-readable reason for a failed child-process call. */
function errDetail(err) {
  return String(err?.stderr || err?.message || err).slice(0, 300);
}

// ---------------------------------------------------------------------------
// detection
// ---------------------------------------------------------------------------
let statusCache = null;

export async function converterStatus(force = false) {
  if (statusCache && !force) return statusCache;
  statusCache = {
    platform: process.platform,
    textutil: IS_MAC && (await exists(TEXTUTIL)),
    soffice: await findSoffice(),
    wordCom: IS_WIN && (await detectWordCom()),
  };
  return statusCache;
}

export function describeStatus(status) {
  const parts = [];
  if (status.textutil) parts.push('macOS textutil');
  if (status.soffice) parts.push('LibreOffice');
  if (status.wordCom) parts.push('Word COM');
  parts.push('内置 RTF/ODT/doc 纯 JS 解析');
  return parts.join('、');
}

function sofficeCandidates() {
  const list = [];
  if (IS_MAC) list.push('/Applications/LibreOffice.app/Contents/MacOS/soffice');
  if (IS_WIN) list.push(...winSofficeCandidates());
  list.push('/usr/bin/soffice', '/usr/local/bin/soffice', '/snap/bin/libreoffice', '/usr/bin/libreoffice');
  return list;
}

function winSofficeCandidates() {
  const bases = [process.env.ProgramFiles, process.env['ProgramFiles(x86)'], process.env.LOCALAPPDATA].filter(Boolean);
  return bases.flatMap((base) => [
    join(base, 'LibreOffice', 'program', 'soffice.exe'),
    join(base, 'Programs', 'LibreOffice', 'program', 'soffice.exe'),
  ]);
}

async function findSoffice() {
  for (const candidate of sofficeCandidates()) {
    if (await exists(candidate)) return candidate;
  }
  for (const name of ['soffice', 'libreoffice']) {
    try {
      await execFileP(name, ['--version'], { timeout: 20000 });
      return name;
    } catch { /* not on PATH */ }
  }
  return null;
}

const WORD_COM_PROBE = [
  '$ErrorActionPreference = "Stop"',
  'try {',
  '  $w = New-Object -ComObject Word.Application',
  '  $w.Quit()',
  '  [Console]::Out.Write("YES")',
  '} catch { [Console]::Out.Write("NO") }',
].join('\n');

async function detectWordCom() {
  try {
    const { stdout } = await execFileP('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', WORD_COM_PROBE], { timeout: 30000 });
    return stdout.includes('YES');
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// textutil (macOS)
// ---------------------------------------------------------------------------
async function runTextutil(args) {
  try {
    const { stdout } = await execFileP(TEXTUTIL, args, { maxBuffer: CAPS.MAX_TEXT_BYTES * 4, timeout: 120000 });
    return stdout;
  } catch (err) {
    throw new OfficeError(`textutil 失败: ${errDetail(err)}`, 'TEXTUTIL_FAILED');
  }
}

export async function textutilRead(path, { asHtml = false } = {}) {
  const fmt = asHtml ? 'html' : 'txt';
  const args = ['-convert', fmt, '-stdout'];
  if (fmt === 'txt') args.push('-encoding', 'UTF-8');
  args.push(path);
  return runTextutil(args);
}

export async function textutilConvert(inputPath, outputPath) {
  const outExt = extOf(outputPath);
  if (!TEXTUTIL_TARGETS.includes(outExt)) throw new OfficeError(`textutil 不支持输出 .${outExt}`, 'UNSUPPORTED_FORMAT');
  await runTextutil(['-convert', outExt, '-output', outputPath, inputPath]);
  return outputPath;
}

// ---------------------------------------------------------------------------
// LibreOffice
// ---------------------------------------------------------------------------
async function runSoffice(soffice, args) {
  try {
    return await execFileP(soffice, args, { timeout: 180000, maxBuffer: 4 * 1024 * 1024 });
  } catch (err) {
    throw new OfficeError(`LibreOffice 调用失败: ${errDetail(err)}`, 'SOFFICE_FAILED');
  }
}

async function sofficeProduce(soffice, inputPath, outDir, target) {
  if (!SOFFICE_FILTERS.has(target)) throw new OfficeError(`LibreOffice 不支持输出 .${target}`, 'UNSUPPORTED_FORMAT');
  await runSoffice(soffice, ['--headless', '--norestore', '--convert-to', target, '--outdir', outDir, inputPath]);
  const produced = join(outDir, basename(inputPath).replace(/\.[^.]+$/, '') + '.' + target);
  if (!(await exists(produced))) throw new OfficeError('LibreOffice 未生成目标文件', 'SOFFICE_NO_OUTPUT');
  return produced;
}

export async function sofficeRead(soffice, path, { asHtml = false } = {}) {
  const dir = await mkdtemp(join(tmpdir(), 'dsh-office-lo-'));
  try {
    return await readFile(await sofficeProduce(soffice, path, dir, asHtml ? 'html' : 'txt'), 'utf8');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

export async function sofficeConvert(soffice, inputPath, outputPath) {
  const dir = await mkdtemp(join(tmpdir(), 'dsh-office-lo-'));
  try {
    const produced = await sofficeProduce(soffice, inputPath, dir, extOf(outputPath));
    await writeFile(outputPath, await readFile(produced));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
  return outputPath;
}

// ---------------------------------------------------------------------------
// Microsoft Word COM (Windows)
// ---------------------------------------------------------------------------
const WORD_COM_SCRIPT = `param([string]$Source, [string]$Target, [int]$Format)
$ErrorActionPreference = "Stop"
$word = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($Source, $false, $true)
  try { $doc.SaveAs([ref]$Target, [ref]$Format) } finally { $doc.Close(0) }
  [Console]::Out.Write("OK")
} catch {
  [Console]::Error.Write($_.Exception.Message)
  exit 1
} finally {
  if ($word -ne $null) { try { $word.Quit() } catch {} }
}
`;

export async function wordComConvert(inputPath, outputPath) {
  const format = WORD_COM_FORMATS[extOf(outputPath)];
  if (format === undefined) throw new OfficeError(`Word COM 不支持输出 .${extOf(outputPath)}`, 'UNSUPPORTED_FORMAT');
  const dir = await mkdtemp(join(tmpdir(), 'dsh-office-word-'));
  const scriptPath = join(dir, 'convert.ps1');
  try {
    await writeFile(scriptPath, WORD_COM_SCRIPT, 'utf8');
    const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
      '-Source', inputPath, '-Target', outputPath, '-Format', String(format)];
    const { stdout } = await execFileP('powershell.exe', args, { timeout: 180000, maxBuffer: 4 * 1024 * 1024 });
    if (!String(stdout).includes('OK')) throw new OfficeError('Word COM 未返回成功标记', 'WORD_COM_FAILED');
  } catch (err) {
    if (err instanceof OfficeError) throw err;
    throw new OfficeError(`Word COM 转换失败: ${errDetail(err)}`, 'WORD_COM_FAILED');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
  return outputPath;
}

export async function wordComRead(path, { asHtml = false } = {}) {
  const dir = await mkdtemp(join(tmpdir(), 'dsh-office-word-'));
  try {
    const out = join(dir, asHtml ? 'out.html' : 'out.txt');
    await wordComConvert(path, out);
    return await readFile(out, 'utf8');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

/** Ordered external read backends for the current platform. */
export async function externalReadBackends(path, { asHtml = false } = {}) {
  const status = await converterStatus();
  const ext = extOf(path);
  const backends = [];
  if (status.textutil) backends.push(['macOS textutil', () => textutilRead(path, { asHtml })]);
  if (status.soffice) backends.push(['LibreOffice', () => sofficeRead(status.soffice, path, { asHtml })]);
  if (status.wordCom && ext === 'doc') backends.push(['Word COM', () => wordComRead(path, { asHtml })]);
  return backends;
}

/** Ordered external convert backends for the target extension. */
export async function externalConvertBackends(inputPath, outputPath) {
  const status = await converterStatus();
  const target = extOf(outputPath);
  const backends = [];
  if (status.textutil && TEXTUTIL_TARGETS.includes(target)) {
    backends.push(['macOS textutil', () => textutilConvert(inputPath, outputPath)]);
  }
  if (status.soffice) backends.push(['LibreOffice', () => sofficeConvert(status.soffice, inputPath, outputPath)]);
  if (status.wordCom && WORD_COM_FORMATS[target] !== undefined) {
    backends.push(['Word COM', () => wordComConvert(inputPath, outputPath)]);
  }
  return backends;
}

export async function hasExternalConverter(targetExt) {
  const status = await converterStatus();
  if (targetExt === 'pdf') return Boolean(status.soffice || status.wordCom);
  return Boolean(status.textutil || status.soffice || status.wordCom);
}
