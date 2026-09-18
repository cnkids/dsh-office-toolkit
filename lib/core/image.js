// 图片嵌入:按魔数识别 PNG/JPEG/GIF/BMP 并读出真实像素尺寸,把 <img> 变成真正的图片。
//
// 刻意不做网络请求,也不引入 image-size / probe-image-size 这类依赖(它们曾把
// `dsh plugin add` 拖进 postinstall 审批);更不直接碰文件系统 —— 本地图片的字节由
// 调用方(工具层,具备沙箱路径解析)通过 readBytes 注入,core 只处理字节。
// 尺寸不够用时按正文宽度等比缩到页内。
import { OfficeError } from './util.js';

/** 单张图片上限:超过就报错,避免文档或内存被撑爆。 */
export const MAX_IMAGE_BYTES = 8 * 1024 * 1024;
const MAX_DIMENSION = 20000;
const MIME_TYPES = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/x-ms-bmp': 'bmp' };

function probePng(data) {
  if (data.length < 24 || data.readUInt32BE(0) !== 0x89504e47 || data.readUInt32BE(4) !== 0x0d0a1a0a) return null;
  if (data.toString('latin1', 12, 16) !== 'IHDR') return null;
  return { type: 'png', width: data.readUInt32BE(16), height: data.readUInt32BE(20) };
}

function probeJpeg(data) {
  if (data.length < 4 || data[0] !== 0xff || data[1] !== 0xd8) return null;
  let at = 2;
  while (at + 9 < data.length) {
    if (data[at] !== 0xff) {
      at += 1;
      continue;
    }
    const marker = data[at + 1];
    // 无长度字段的标记(D8/D9/01/RSTn)直接跳过
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd9)) {
      at += 2;
      continue;
    }
    const size = data.readUInt16BE(at + 2);
    if (size < 2) return null;
    const isStartOfFrame = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isStartOfFrame) return { type: 'jpg', height: data.readUInt16BE(at + 5), width: data.readUInt16BE(at + 7) };
    at += 2 + size;
  }
  return null;
}

function probeGif(data) {
  if (data.length < 10 || data.toString('latin1', 0, 3) !== 'GIF') return null;
  return { type: 'gif', width: data.readUInt16LE(6), height: data.readUInt16LE(8) };
}

function probeBmp(data) {
  if (data.length < 26 || data.toString('latin1', 0, 2) !== 'BM') return null;
  return { type: 'bmp', width: Math.abs(data.readInt32LE(18)), height: Math.abs(data.readInt32LE(22)) };
}

/**
 * 按魔数识别图片并读出像素尺寸(docx 只支持 png/jpg/gif/bmp)。
 * @returns { type, width, height } 或 null
 */
export function probeImage(data) {
  const bytes = Buffer.isBuffer(data) ? data : Buffer.from(data ?? []);
  for (const probe of [probePng, probeJpeg, probeGif, probeBmp]) {
    const found = probe(bytes);
    if (found && found.width > 0 && found.height > 0 && found.width <= MAX_DIMENSION && found.height <= MAX_DIMENSION) return found;
  }
  return null;
}

function tooLarge(what, size) {
  return new OfficeError(`${what} ${(size / 1048576).toFixed(1)} MB，超过单张图片 8 MB 上限；请先压缩或用 office_convert 缩小`, 'IMAGE_TOO_LARGE');
}

/** data URL 手工切分:不对正文跑正则,结构上不可能回溯。 */
function fromDataUrl(src) {
  const comma = src.indexOf(',');
  if (comma < 0) return null;
  const header = src.slice('data:'.length, comma);
  const base64 = /;base64$/i.test(header);
  const mime = header.replace(/;base64$/i, '').trim().toLowerCase();
  if (!MIME_TYPES[mime]) {
    throw new OfficeError(`<img> 的 data URL 类型 ${mime || '(未标明)'} 不支持（只支持 PNG / JPEG / GIF / BMP）`, 'IMAGE_UNSUPPORTED');
  }
  const body = src.slice(comma + 1);
  const data = base64 ? Buffer.from(body, 'base64') : Buffer.from(decodeURIComponent(body), 'latin1');
  if (data.length > MAX_IMAGE_BYTES) throw tooLarge('data URL 图片', data.length);
  return data;
}

/** 本地图片:字节由调用方注入(工具层已按沙箱规则解析路径并读盘)。 */
async function readLocalFile(src, readBytes) {
  let data;
  try {
    data = await readBytes(src);
  } catch (err) {
    throw new OfficeError(`<img> 读不到图片文件 ${src}: ${err?.code || err?.message || err}`, 'IMAGE_NOT_FOUND');
  }
  if (data.length > MAX_IMAGE_BYTES) throw tooLarge(`图片 ${src}`, data.length);
  return data;
}

/**
 * `<img src="…">` → 图片字节 + 尺寸。src 只接受本地路径或 data: URL。
 * @param {string} src
 * @param {(path: string) => Promise<Buffer>} readBytes 本地图片的读取器(由调用方注入)
 */
export async function loadImage(src, readBytes) {
  const value = String(src ?? '').trim();
  if (!value) throw new OfficeError('<img> 缺少 src', 'INVALID_ARGS');
  if (/^https?:/i.test(value)) {
    throw new OfficeError(`不支持远程图片（${value.slice(0, 60)}）—— 本插件不联网，请先下载到本地再用相对路径引用`, 'IMAGE_REMOTE');
  }
  if (!/^data:/i.test(value) && typeof readBytes !== 'function') {
    throw new OfficeError('<img> 需要图片读取器才能嵌入本地图片：请用 office_write_docx / office_convert 生成文档，或去掉 src 只留 alt 文字', 'IMAGE_UNSUPPORTED');
  }
  const data = /^data:/i.test(value) ? fromDataUrl(value) : await readLocalFile(value, readBytes);
  const probed = probeImage(data);
  if (!probed) {
    throw new OfficeError(
      `<img> 的图片格式不支持：${value.startsWith('data:') ? value.slice(0, 24) + '…' : value}。只支持 PNG / JPEG / GIF / BMP（SVG / WebP / TIFF 请先转成 PNG）`,
      'IMAGE_UNSUPPORTED'
    );
  }
  return { ...probed, data };
}

/**
 * `120` / `120px` / `50%` → 像素值,认不出返回 0。
 * 用后缀判断 + Number() 解析,不对用户给的串跑正则。
 */
function lengthHint(raw, textWidthPx) {
  const value = String(raw ?? '').trim().toLowerCase();
  if (!value) return 0;
  if (value.endsWith('%')) {
    const percent = Number(value.slice(0, -1).trim());
    return Number.isFinite(percent) && percent > 0 ? Math.round((percent / 100) * textWidthPx) : 0;
  }
  const parsed = Number(value.endsWith('px') ? value.slice(0, -2).trim() : value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

/** `<img width="120">` 或 CSS `width:120px / 50%` → 目标宽度(px);CSS 优先于属性。 */
export function widthHintOf(readAttr, declarations, textWidthPx) {
  return lengthHint(declarations.width, textWidthPx) || lengthHint(readAttr('width'), 0);
}

export function heightHintOf(readAttr, declarations) {
  return lengthHint(declarations.height, 0) || lengthHint(readAttr('height'), 0);
}

/**
 * 目标显示尺寸(px):显式宽高优先,其余等比；再整体缩到正文宽内、页高以内。
 * docx 的 transformation 以 px 计(内部按 9525 EMU/px 换算)。
 */
export function fitImageSize(image, { widthHint = 0, heightHint = 0, maxWidth = 600, maxHeight = 900 } = {}) {
  const natural = { width: image.width, height: image.height };
  let width = widthHint > 0 ? widthHint : natural.width;
  let height = heightHint > 0 ? heightHint : natural.height;
  if (widthHint > 0 && heightHint <= 0) height = Math.round((natural.height / natural.width) * width);
  if (heightHint > 0 && widthHint <= 0) width = Math.round((natural.width / natural.height) * height);
  const scale = Math.min(1, maxWidth / width, maxHeight / height);
  return {
    width: Math.max(1, Math.round(width * scale)),
    height: Math.max(1, Math.round(height * scale)),
  };
}
