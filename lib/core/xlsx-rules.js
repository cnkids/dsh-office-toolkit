// Conditional-formatting and data-validation rules, normalised for the openpyxl
// fallback. The ExcelJS builders live in excel.js; this module emits the same
// semantics as plain JSON (openpyxl needs `kind` + explicit formulas), so a broken
// ExcelJS import can still produce equivalent rules.
import { OfficeError, isPlainObject } from './util.js';

const CF_TYPES = new Set(['cellIs', 'expression', 'colorScale', 'dataBar', 'iconSet', 'top10', 'aboveAverage', 'containsText', 'timePeriod']);
const CF_OPERATORS = new Set(['equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'between', 'notBetween']);
const TEXT_OPERATORS = new Set(['containsText', 'notContains', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors']);
const TIME_PERIODS = new Set(['today', 'yesterday', 'tomorrow', 'last7Days', 'thisWeek', 'lastWeek', 'nextWeek', 'thisMonth', 'lastMonth', 'nextMonth']);
const DV_TYPES = new Set(['list', 'whole', 'decimal', 'date', 'textLength', 'custom']);
const DV_OPERATORS = new Set(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']);
const ICON_STOPS = { 3: [0, 33, 67], 4: [0, 25, 50, 75], 5: [0, 20, 40, 60, 80] };
const RANGE_RE = /^[A-Z]+\d+(?::[A-Z]+\d+)?$/;

/** 单元格区域是否为 `A1` / `A1:C10` 形状。 */
export function isRangeSpec(ref) {
  return RANGE_RE.test(String(ref || ''));
}

function requireChoice(value, allowed, what, fallback) {
  if (value === undefined || value === null || value === '') {
    if (fallback !== undefined) return fallback;
    throw new OfficeError(`${what} 是必填项，可为 ${[...allowed].join(' / ')}`);
  }
  const picked = String(value);
  if (!allowed.has(picked)) throw new OfficeError(`${what} 可为 ${[...allowed].join(' / ')}，收到「${picked}」`);
  return picked;
}

function requireText(value, what) {
  const text = value === undefined || value === null ? '' : String(value);
  if (!text) throw new OfficeError(`conditional_format 的 ${what} 是必填项`);
  return text;
}

/** 区域左上角单元格(文本类规则要写相对引用公式)。 */
export function topLeftOf(range) {
  const head = String(range).split(':')[0];
  return /^[A-Z]+\d+$/.exec(head)?.[0] ?? 'A1';
}

/** Excel 字符串字面量:内部双引号翻倍。 */
function literal(text) {
  return `"${String(text).replaceAll('"', '""')}"`;
}

/** 阈值统一成字符串:数字按数值写法,文本加引号转义。 */
function quoted(value) {
  return typeof value === 'number' ? String(value) : literal(value);
}

function cellIsThresholds(rule, operator) {
  const list = Array.isArray(rule.values) ? rule.values : [rule.value, rule.value2].filter((v) => v !== undefined && v !== null);
  if (operator === 'between' || operator === 'notBetween') {
    if (list.length !== 2) throw new OfficeError('conditional_format 的 between / notBetween 需要 value 与 value2 两个阈值');
    return list.map(quoted);
  }
  if (!list.length) throw new OfficeError('conditional_format 需要 value（cellIs 等的阈值）');
  return [quoted(list[0])];
}

/** 文本类规则 -> 等价表达式公式(openpyxl 无 containsText 专用类)。 */
function textFormula(operator, text, cell) {
  const needle = literal(text);
  const search = `SEARCH(${needle},${cell})`;
  switch (operator) {
    case 'notContains': return `ISERROR(${search})`;
    case 'beginsWith': return `LEFT(${cell},${String(text).length})=${literal(text)}`;
    case 'endsWith': return `RIGHT(${cell},${String(text).length})=${literal(text)}`;
    case 'containsBlanks': return `LEN(TRIM(${cell}))=0`;
    case 'notContainsBlanks': return `LEN(TRIM(${cell}))>0`;
    case 'containsErrors': return `ISERROR(${cell})`;
    case 'notContainsErrors': return `NOT(ISERROR(${cell}))`;
    default: return `NOT(ISERROR(${search}))`;
  }
}

function scaleColors(rule) {
  const given = Array.isArray(rule.colors) ? rule.colors : [];
  if (given.length && given.length !== 2 && given.length !== 3) {
    throw new OfficeError('conditional_format 的 colorScale 需要 2 个或 3 个颜色（colors: ["F8696B","63BE7B"]）');
  }
  return given.length ? given : ['FFFFFF', '638EC6'];
}

/** 2 色 -> min/max；3 色 -> min/percentile/max(与 excel.js 的色阶停点一致)。 */
function scaleStops(count) {
  if (count <= 2) return [{ type: 'min' }, { type: 'max' }];
  return [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }];
}

function iconStops(iconSet) {
  const count = Number(String(iconSet).charAt(0)) || 3;
  return ICON_STOPS[count] ?? ICON_STOPS[3];
}

/**
 * 一条条件格式规则 -> `{kind, ...}`(kind 供 Python 端分派,样式放 dxf)。
 * @param {object} rule 工具入参里的一条规则
 * @param {string} range 规则所在区域(文本类规则取左上角做相对引用)
 */
export function pythonRule(rule, range) {
  const type = requireChoice(rule.type, CF_TYPES, 'conditional_format 的 type');
  const dxf = rule.style;
  const cell = topLeftOf(range);
  switch (type) {
    case 'cellIs': {
      const operator = requireChoice(rule.operator, CF_OPERATORS, 'conditional_format 的 operator', 'greaterThan');
      return { kind: 'cellIs', operator, formula: cellIsThresholds(rule, operator), dxf };
    }
    case 'expression':
      return { kind: 'expression', formula: [requireText(rule.formula ?? rule.value, 'formula')], dxf };
    case 'colorScale': {
      const colors = scaleColors(rule);
      return {
        kind: 'colorScale',
        colors: colors.map((c) => String(c).replace(/^#/, '')),
        stops: scaleStops(colors.length),
      };
    }
    case 'dataBar':
      return { kind: 'dataBar', color: String(rule.color || '638EC6').replace(/^#/, '') };
    case 'iconSet': {
      const iconSet = String(rule.iconSet || '3TrafficLights1');
      return { kind: 'iconSet', iconSet, showValue: rule.showValue !== false, stops: iconStops(iconSet) };
    }
    case 'top10':
      return { kind: 'top10', rank: Number(rule.rank ?? 10), percent: Boolean(rule.percent), bottom: Boolean(rule.bottom), dxf };
    case 'aboveAverage':
      return { kind: 'aboveAverage', aboveAverage: rule.aboveAverage !== false, dxf };
    case 'containsText': {
      const operator = requireChoice(rule.operator, TEXT_OPERATORS, 'conditional_format 的 operator', 'containsText');
      return { kind: 'expression', formula: [textFormula(operator, requireText(rule.text, 'text'), cell)], dxf };
    }
    default:
      return {
        kind: 'timePeriod',
        timePeriod: requireChoice(rule.timePeriod, TIME_PERIODS, 'conditional_format 的 timePeriod', 'today'),
        dxf,
      };
  }
}

/** conditional_format 的 range + rule/rules -> `{range, rules}`。 */
export function pythonRules(op) {
  const range = String(op.range || op.ref || '');
  if (!isRangeSpec(range)) throw new OfficeError(`conditional_format 的 range 形如 A1:A20（收到「${range}」）`);
  const rules = (Array.isArray(op.rules) ? op.rules : [isPlainObject(op.rule) ? op.rule : op]).map((r) => pythonRule(r, range));
  return { range, rules };
}

/** list 校验的公式:候选值、区域引用或公式。 */
function validationFormulae(spec, type) {
  if (type === 'list') {
    if (Array.isArray(spec.values)) {
      if (!spec.values.length) throw new OfficeError('data_validation 的 values 不能是空数组');
      return [literal(spec.values.map((v) => String(v).replaceAll('"', '')).join(','))];
    }
    if (spec.source !== undefined) return [String(spec.source).replace(/^=/, '')];
    if (Array.isArray(spec.formulae)) return spec.formulae;
    throw new OfficeError('data_validation 的 list 需要 values: ["甲","乙"] 或 source: "=Sheet2!$A$1:$A$5"');
  }
  const first = spec.value === undefined ? spec.formula : spec.value;
  const list = Array.isArray(spec.values) ? spec.values : [first, spec.value2].filter((v) => v !== undefined && v !== null);
  if (!list.length) throw new OfficeError('data_validation 需要 value（或 values / formula）');
  return list.map((v) => (typeof v === 'number' ? v : String(v)));
}

/** 一条数据验证 -> openpyxl DataValidation 的构造参数。 */
export function pythonValidation(op) {
  const spec = isPlainObject(op.rule) ? op.rule : op;
  const type = requireChoice(spec.type, DV_TYPES, 'data_validation 的 type', 'list');
  const formulae = validationFormulae(spec, type);
  return {
    type,
    formula1: String(formulae[0]),
    formula2: formulae[1] === undefined ? null : String(formulae[1]),
    allowBlank: spec.allowBlank !== false,
    operator: type === 'list' || type === 'custom' ? null : requireChoice(spec.operator, DV_OPERATORS, 'data_validation 的 operator', 'between'),
    showErrorMessage: spec.showErrorMessage ?? spec.error !== undefined,
    showInputMessage: spec.showInputMessage ?? (spec.prompt !== undefined || spec.promptTitle !== undefined),
    errorTitle: spec.errorTitle === undefined ? null : String(spec.errorTitle),
    error: spec.error === undefined ? null : String(spec.error),
    errorStyle: spec.errorStyle === undefined ? null : String(spec.errorStyle),
    promptTitle: spec.promptTitle === undefined ? null : String(spec.promptTitle),
    prompt: spec.prompt === undefined ? null : String(spec.prompt),
  };
}

/** data_validation 的 range + rule -> openpyxl 用的 `{range, rule}`。 */
export function pythonDataValidation(op) {
  const range = String(op.range || op.ref || '');
  if (!isRangeSpec(range)) throw new OfficeError(`data_validation 的 range 形如 C2:C100（收到「${range}」）`);
  return { range, validation: pythonValidation(op) };
}
