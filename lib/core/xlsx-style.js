// Cell style normalisation for the openpyxl fallback.
// Mirrors excel.js#applyCellStyle field for field, but emits plain JSON instead of
// mutating an ExcelJS cell, so the child process never has to guess at defaults.
import { argb } from './xlsx-cells.js';
import { isPlainObject } from './util.js';

const BOOL_FLAGS = ['bold', 'italic', 'wrap', 'border'];
const TEXT_FIELDS = ['color', 'fill', 'numFmt', 'align', 'valign', 'borderColor'];
// openpyxl 的垂直对齐没有 ExcelJS 的 middle,统一折成 center
const VALIGN_ALIAS = { middle: 'center', top: 'top', bottom: 'bottom', center: 'center' };
const ALIGN_ALIAS = { left: 'left', center: 'center', right: 'right', justify: 'justify', distributed: 'distributed' };

function pickBool(style, out) {
  for (const flag of BOOL_FLAGS) {
    if (style[flag]) out[flag] = true;
  }
}

function pickText(style, out) {
  for (const field of TEXT_FIELDS) {
    if (style[field] === undefined || style[field] === null || style[field] === '') continue;
    if (field === 'color' || field === 'fill' || field === 'borderColor') out[field] = argb(style[field]);
    else out[field] = String(style[field]);
  }
}

function pickEnums(style, out) {
  const align = ALIGN_ALIAS[style.align];
  if (align) out.align = align;
  const valign = VALIGN_ALIAS[style.valign];
  if (valign) out.valign = valign;
}

function pickFontSize(style, out) {
  const size = Number(style.fontSize);
  if (Number.isFinite(size) && size > 0) out.size = size;
}

/** `{bold...,fill,align...}` -> a normalised style object (empty object when nothing usable). */
export function normalizeStyle(style) {
  const out = {};
  if (!isPlainObject(style)) return out;
  pickBool(style, out);
  pickFontSize(style, out);
  pickText(style, out);
  pickEnums(style, out);
  return out;
}
