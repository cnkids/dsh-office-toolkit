// 表格计算层:筛选 / 分组 / 聚合 / 排序 / 画像。纯函数,不碰 IO —— 这是 office_query 的「算」。
// 设计约束:没有 eval、没有表达式解析,只有固定的算子与聚合函数,结构上不存在注入面。
import { OfficeError } from './util.js';

export const FILTER_OPS = ['eq', 'ne', 'gt', 'gte', 'lt', 'lte', 'contains', 'startsWith', 'endsWith', 'in', 'notIn', 'isBlank', 'notBlank'];
export const AGG_FNS = ['sum', 'avg', 'min', 'max', 'count', 'countDistinct'];

const TEXT_OPS = new Set(['contains', 'startsWith', 'endsWith']);
const MAX_GROUPS = 5000;
const MAX_RESULT_ROWS = 2000;
const DEFAULT_RESULT_ROWS = 200;
const TOP_VALUES = 3;
const FREQ_CAP = 20000;
const COL_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

// ---------------------------------------------------------------------------
// 值归一化 —— 真实表格里 "1,234.00"、"12.5%"、"¥88"、"2025/1/1" 都算数
// ---------------------------------------------------------------------------
const CURRENCY_WS = /[¥￥$€£\s\u00a0\u3000]/g;
// 注意 `\d+\.?\d*` 是重叠量词:`\.?` 可以匹配空串,遇到超长数字串会退化成 O(n²)
// —— 6 万位的单元格内容就能把主线程卡住十几秒。改成 `\d+(?:\.\d*)?`,分支首字符互斥,恒为线性。
const PLAIN_NUMBER = /^[-+]?(?:\d+(?:\.\d*)?|\.\d+)$/;
const GROUPED_NUMBER = /^[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/;
const DATE_LIKE = /^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:[ T].*)?$/;

const text = (v) => (v === null || v === undefined ? '' : String(v));
const isBlank = (v) => v === null || v === undefined || v === '';

/** 尽力转数字:`¥1,234.00`→1234、`12.5%`→0.125;不像数字就返回 null(不猜)。 */
export function toNumber(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v !== 'string') return null;
  let s = v.replaceAll(CURRENCY_WS, '');
  if (!s) return null;
  let scale = 1;
  if (s.endsWith('%')) { scale = 0.01; s = s.slice(0, -1).replaceAll(CURRENCY_WS, ''); }
  if (s.includes(',')) {
    if (!GROUPED_NUMBER.test(s)) return null;
    s = s.replaceAll(',', '');
  }
  if (!PLAIN_NUMBER.test(s)) return null;
  const n = Number(s) * scale;
  return Number.isFinite(n) ? n : null;
}

/** 日期归一化成 YYYY-MM-DD(`2025/1/1`、`2025年1月1日`、带时间的都认);不是日期返回 null。 */
export function toDay(v) {
  if (typeof v !== 'string') return null;
  const m = DATE_LIKE.exec(v);
  if (!m) return null;
  return `${m[1]}-${String(Number(m[2])).padStart(2, '0')}-${String(Number(m[3])).padStart(2, '0')}`;
}

/** 分组/去重用的归一化键:1 与 "1" 同组,`2025/1/1` 与 `2025-01-01` 同组。 */
export function valueKey(v) {
  const n = toNumber(v);
  if (n !== null) return `n:${n}`;
  const d = toDay(v);
  return d === null ? `s:${text(v)}` : `d:${d}`;
}

/**
 * 排序/比较:两边都是数字按数值,都是日期按日期,否则按字符串。
 * 一边是数字、另一边不是(如「金额」列里混着"待定")时返回 null —— 不可比,
 * 绝不能退回字典序,否则 "待定" > 1000 会成立,筛选结果会静默出错。
 */
function textCompare(a, b) {
  const sa = text(a);
  const sb = text(b);
  if (sa === sb) return 0;
  return sa < sb ? -1 : 1;
}

/**
 * 同一族内比较:两边都能被 `parse` 解析才比,否则返回 null(不可比)。
 * 两边都解析不出时返回 undefined,表示"这一类不适用",交给下一族。
 */
function compareSameKind(a, b, parse) {
  const pa = parse(a);
  const pb = parse(b);
  if (pa === null && pb === null) return undefined;
  if (pa === null || pb === null) return null;
  if (pa === pb) return 0;
  return pa < pb ? -1 : 1;
}

function compare(a, b) {
  const asNumber = compareSameKind(a, b, toNumber);
  if (asNumber !== undefined) return asNumber;
  const asDay = compareSameKind(a, b, toDay);
  return asDay === undefined ? textCompare(a, b) : asDay;
}

function round(n) {
  if (!Number.isFinite(n)) return null;
  return Number(n.toPrecision(12));
}

function fmtCell(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return String(round(v));
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function escapeTsv(v) {
  return fmtCell(v).replaceAll('\t', '␉').replaceAll(/\r?\n/g, '⏎');
}

// ---------------------------------------------------------------------------
// 条件
// ---------------------------------------------------------------------------
const toList = (value) => (Array.isArray(value) ? value : [value]);

function textMatch(cell, needle, op) {
  const s = text(cell).toLowerCase();
  const n = text(needle).toLowerCase();
  if (op === 'contains') return s.includes(n);
  return op === 'startsWith' ? s.startsWith(n) : s.endsWith(n);
}

function compareWith(op, c) {
  // 类型不可比时,除 ne 外一律不成立(宁可少命中,也不给静默错误的结果)
  if (c === null) return op === 'ne';
  if (op === 'eq') return c === 0;
  if (op === 'ne') return c !== 0;
  if (op === 'gt') return c > 0;
  if (op === 'gte') return c >= 0;
  if (op === 'lt') return c < 0;
  return c <= 0;
}

/** 排序用:类型不可比时退回文本序,保证排序结果稳定(Sort 不允许返回 null)。 */
function sortValue(a, b) {
  const c = compare(a, b);
  return c === null ? textCompare(a, b) : c;
}

/** 单个条件求值。`op` 缺失按 eq 处理(写法宽容一点,少一次来回)。 */
export function evalCondition(cell, cond) {
  const op = text(cond?.op) || 'eq';
  if (op === 'isBlank') return isBlank(cell);
  if (op === 'notBlank') return !isBlank(cell);
  if (op === 'in') return toList(cond.value).some((x) => compare(cell, x) === 0);
  if (op === 'notIn') return !toList(cond.value).some((x) => compare(cell, x) === 0);
  if (TEXT_OPS.has(op)) return textMatch(cell, cond.value, op);
  return compareWith(op, compare(cell, cond.value));
}

// ---------------------------------------------------------------------------
// 表结构
// ---------------------------------------------------------------------------
function colLetter(index) {
  let i = index + 1;
  let out = '';
  while (i > 0) {
    out = COL_LETTERS[(i - 1) % 26] + out;
    i = Math.floor((i - 1) / 26);
  }
  return out;
}

function uniqueName(raw, index, used) {
  const base = text(raw).trim() || colLetter(index);
  let name = base;
  let n = 2;
  while (used.has(name)) { name = `${base}_${n}`; n += 1; }
  used.add(name);
  return name;
}

function headerRowOf(raw, rowCount) {
  if (raw === undefined || raw === null || raw === '') return 1;
  const n = Number(raw);
  if (!Number.isInteger(n) || n < 0) throw new OfficeError('headerRow 需为 ≥ 0 的整数（0 表示该表没有表头）', 'INVALID_ARGS');
  if (n > rowCount) throw new OfficeError(`headerRow=${n} 超出范围：该工作表只有 ${rowCount} 行`, 'INVALID_ARGS');
  return n;
}

/** 建立列名(表头为空或重名都有兜底)与数据行。 */
export function buildTable(matrix, headerRow) {
  const width = Math.max(matrix.reduce((w, row) => Math.max(w, row.length), 0), 1);
  const header = headerRow > 0 ? (matrix[headerRow - 1] || []) : [];
  const used = new Set();
  const columns = [];
  for (let c = 0; c < width; c++) columns.push(uniqueName(header[c], c, used));
  return { columns, rows: headerRow > 0 ? matrix.slice(headerRow) : matrix };
}

function columnIndex(columns, ref) {
  const key = text(ref).trim();
  const at = columns.indexOf(key);
  if (at >= 0) return at;
  const lower = key.toLowerCase();
  const ci = columns.findIndex((c) => c.toLowerCase() === lower);
  if (ci >= 0) return ci;
  throw new OfficeError(`找不到列「${key}」。可用列名：${columns.join('、')}`, 'UNKNOWN_COLUMN');
}

// ---------------------------------------------------------------------------
// 聚合累加器
// ---------------------------------------------------------------------------
function newAcc(fn) {
  return { fn, sum: 0, num: 0, count: 0, min: null, max: null, txtMin: null, txtMax: null, skipped: 0, freq: new Map(), capped: false };
}

function trackText(st, s) {
  if (st.txtMin === null || s < st.txtMin) st.txtMin = s;
  if (st.txtMax === null || s > st.txtMax) st.txtMax = s;
}

function addValue(st, v) {
  if (isBlank(v)) return;
  st.count += 1;
  if (st.fn === 'countDistinct') return addDistinct(st, v);
  if (st.fn === 'count') return undefined;
  const n = toNumber(v);
  if (n === null) {
    st.skipped += 1;
    // 日期按归一化后的 YYYY-MM-DD 比较,否则 "2025/1/3" 会排在 "2024/12/31" 前面
    if (st.fn === 'min' || st.fn === 'max') trackText(st, toDay(v) ?? text(v));
    return undefined;
  }
  st.num += 1;
  st.sum += n;
  st.min = st.min === null ? n : Math.min(st.min, n);
  st.max = st.max === null ? n : Math.max(st.max, n);
  return undefined;
}

function addDistinct(st, v) {
  if (st.capped) return;
  st.freq.set(valueKey(v), true);
  if (st.freq.size >= FREQ_CAP) st.capped = true;
}

/** @returns {{ value: *, note: string }} 聚合结果(数值列不够时返回 null 而不是 0)。 */
function accResult(st) {
  if (st.fn === 'count') return { value: st.count, note: '' };
  if (st.fn === 'countDistinct') {
    return st.capped ? { value: `≥${FREQ_CAP}`, note: '' } : { value: st.freq.size, note: '' };
  }
  if (st.num === 0) return emptyNumericResult(st);
  const note = st.skipped ? skippedNote(st) : '';
  if (st.fn === 'sum') return { value: round(st.sum), note };
  if (st.fn === 'avg') return { value: round(st.sum / st.num), note };
  return { value: st.fn === 'min' ? st.min : st.max, note };
}

/** 该组/该列一个数值都没有时:min/max 退回文本比较,sum/avg 明确说没数据而不是给 0。 */
function emptyNumericResult(st) {
  if (st.fn === 'min' || st.fn === 'max') {
    return { value: st.fn === 'min' ? st.txtMin : st.txtMax, note: st.skipped ? '按文本/日期比较' : '' };
  }
  return { value: null, note: st.skipped ? `${st.skipped} 个单元格都不是数值` : '无数据' };
}

function skippedNote(st) {
  return st.fn === 'min' || st.fn === 'max'
    ? `${st.skipped} 个非数值未参与(按数值比较)`
    : `${st.skipped} 个非数值单元格已跳过`;
}

// ---------------------------------------------------------------------------
// 查询计划
// ---------------------------------------------------------------------------
function compileWhere(where, columns) {
  if (where === undefined || where === null) return [];
  if (!Array.isArray(where)) throw new OfficeError('where 需为条件数组，如 [{"col":"地区","op":"eq","value":"华东"}]', 'INVALID_ARGS');
  return where.map((cond) => {
    const op = text(cond?.op) || 'eq';
    if (!FILTER_OPS.includes(op)) {
      throw new OfficeError(`不支持的筛选算子「${op}」。可用：${FILTER_OPS.join('、')}`, 'INVALID_ARGS');
    }
    if (op !== 'isBlank' && op !== 'notBlank' && cond.value === undefined) {
      throw new OfficeError(`条件 ${op} 需要 value（如 {"col":"金额","op":"gt","value":1000}）`, 'INVALID_ARGS');
    }
    return { index: columnIndex(columns, cond.col), cond: { ...cond, op } };
  });
}

function compileGroups(groupBy, columns) {
  if (groupBy === undefined || groupBy === null) return [];
  if (!Array.isArray(groupBy)) throw new OfficeError('groupBy 需为列名数组，如 ["地区","产品"]', 'INVALID_ARGS');
  return groupBy.map((name) => ({ name: text(name), index: columnIndex(columns, name) }));
}

function compileAggs(aggregate, columns) {
  if (aggregate === undefined || aggregate === null) return [];
  if (!Array.isArray(aggregate) || !aggregate.length) {
    throw new OfficeError('aggregate 需为非空数组，如 [{"col":"金额","fn":"sum","as":"销售额"}]', 'INVALID_ARGS');
  }
  return aggregate.map((spec) => {
    const fn = text(spec?.fn);
    if (!AGG_FNS.includes(fn)) throw new OfficeError(`不支持的聚合函数「${fn}」。可用：${AGG_FNS.join('、')}`, 'INVALID_ARGS');
    const index = columnIndex(columns, spec.col);
    return { index, fn, as: text(spec.as).trim() || `${fn}(${columns[index]})` };
  });
}

function compileOrder(orderBy, outColumns) {
  if (orderBy === undefined || orderBy === null) return [];
  if (!Array.isArray(orderBy)) throw new OfficeError('orderBy 需为数组，如 [{"col":"销售额","dir":"desc"}]', 'INVALID_ARGS');
  return orderBy.map((spec) => ({
    index: columnIndex(outColumns, spec.col),
    desc: text(spec.dir).toLowerCase() === 'desc',
  }));
}

function clampLimit(raw) {
  if (raw === undefined || raw === null || raw === '') return DEFAULT_RESULT_ROWS;
  const n = Number(raw);
  if (!Number.isInteger(n) || n < 1) throw new OfficeError('limit 需为 ≥ 1 的整数', 'INVALID_ARGS');
  return Math.min(n, MAX_RESULT_ROWS);
}

// ---------------------------------------------------------------------------
// 查询:筛选 / 分组 / 聚合 / 排序
// ---------------------------------------------------------------------------
function filterRows(rows, conditions) {
  if (!conditions.length) return rows;
  return rows.filter((row) => conditions.every((c) => evalCondition(row[c.index], c.cond)));
}

function groupOf(row, groups) {
  const labels = [];
  const keys = [];
  for (const g of groups) {
    labels.push(fmtCell(row[g.index]));
    keys.push(valueKey(row[g.index]));
  }
  return { labels, key: keys.join('\u0000') };
}

/** 分组并累加,返回 [{ labels, values }]。没有 groupBy 时就是一个全局桶(全表汇总)。 */
function bucketize(rows, groups, aggs) {
  const buckets = new Map();
  if (!groups.length) buckets.set('', { labels: [], accs: aggs.map((a) => newAcc(a.fn)) });
  for (const row of rows) {
    const g = groupOf(row, groups);
    let bucket = buckets.get(g.key);
    if (!bucket) {
      if (buckets.size >= MAX_GROUPS) {
        throw new OfficeError(`分组数超过 ${MAX_GROUPS}（分组列可能取值过散）。请先用 where 缩小范围，或换一列分组`, 'TOO_MANY_GROUPS');
      }
      bucket = { labels: g.labels, accs: aggs.map((a) => newAcc(a.fn)) };
      buckets.set(g.key, bucket);
    }
    for (let i = 0; i < aggs.length; i++) addValue(bucket.accs[i], row[aggs[i].index]);
  }
  return [...buckets.values()];
}

function finishBuckets(buckets, aggs) {
  const notes = [];
  const rows = buckets.map((b) => {
    const values = [...b.labels];
    b.accs.forEach((acc, i) => {
      const r = accResult(acc);
      if (r.note) notes.push(`「${aggs[i].as}」${r.note}`);
      values.push(r.value);
    });
    return values;
  });
  return { rows, notes };
}

function sortRows(rows, order) {
  if (!order.length) return rows;
  return [...rows].sort((a, b) => {
    for (const o of order) {
      const c = sortValue(a[o.index], b[o.index]);
      if (c !== 0) return o.desc ? -c : c;
    }
    return 0;
  });
}

function tsvBlock(header, rows) {
  const lines = [header.map(escapeTsv).join('\t')];
  for (const row of rows) lines.push(row.map(escapeTsv).join('\t'));
  return '```tsv\n' + lines.join('\n') + '\n```';
}

function summarize(sheet, matched, extra) {
  const size = `${sheet.rows.length} 行 × ${sheet.columns.length} 列`;
  const filter = matched === sheet.rows.length ? '' : ` → 命中 ${matched} 行`;
  return `工作表「${sheet.name}」${size}${filter}${extra ? ' → ' + extra : ''}`;
}

function notesBlock(notes) {
  const uniq = [...new Set(notes.filter(Boolean))];
  return uniq.length ? '\n' + uniq.map((n) => `> 注：${n}`).join('\n') : '';
}

function aggregateBlock(sheet, rows, spec) {
  const groups = compileGroups(spec.groupBy, sheet.columns);
  const aggs = compileAggs(spec.aggregate, sheet.columns);
  if (!aggs.length) aggs.push({ index: groups[0].index, fn: 'count', as: '行数' });
  const outColumns = [...groups.map((g) => g.name), ...aggs.map((a) => a.as)];
  const limit = clampLimit(spec.limit);
  const { rows: finished, notes } = finishBuckets(bucketize(rows, groups, aggs), aggs);
  const sorted = sortRows(finished, compileOrder(spec.orderBy, outColumns));
  const shown = sorted.slice(0, limit);
  const extra = `${sorted.length} 组`;
  const truncated = sorted.length > shown.length;
  return {
    text: `### ${summarize(sheet, rows.length, extra)}\n\n` + tsvBlock(outColumns, shown) +
      (truncated ? `\n> 共 ${sorted.length} 组，仅显示前 ${shown.length} 组（可用 orderBy 排序取前 N，或用 where 收窄）` : '') +
      notesBlock(notes),
    meta: { kind: 'aggregate', groups: sorted.length, shown: shown.length, columns: outColumns, truncated },
  };
}

// ---------------------------------------------------------------------------
// 画像:一次看清每列长什么样
// ---------------------------------------------------------------------------
function newProfileState() {
  return { blank: 0, numbers: 0, dates: 0, texts: 0, num: 0, sum: 0, min: null, max: null, txtMin: null, txtMax: null, freq: new Map(), capped: false };
}

function addNumeric(st, n) {
  st.numbers += 1;
  st.num += 1;
  st.sum += n;
  st.min = st.min === null ? n : Math.min(st.min, n);
  st.max = st.max === null ? n : Math.max(st.max, n);
}

function bumpFreq(st, v) {
  const key = valueKey(v);
  const hit = st.freq.get(key);
  if (hit) { hit.n += 1; return; }
  if (st.capped) return;
  st.freq.set(key, { n: 1, sample: fmtCell(v) });
  if (st.freq.size >= FREQ_CAP) st.capped = true;
}

function addProfileText(st, v) {
  const day = toDay(v);
  if (day === null) { st.texts += 1; trackText(st, text(v)); return; }
  st.dates += 1;
  trackText(st, day);
}

function addProfileCell(st, v) {
  if (isBlank(v)) { st.blank += 1; return; }
  const n = toNumber(v);
  if (n === null) addProfileText(st, v);
  else addNumeric(st, n);
  bumpFreq(st, v);
}

const byCountDesc = (a, b) => b.n - a.n;

function topValues(st) {
  let best = [];
  for (const e of st.freq.values()) {
    best.push(e);
    if (best.length > TOP_VALUES) best = best.toSorted(byCountDesc).slice(0, TOP_VALUES);
  }
  return best.toSorted(byCountDesc).map((e) => `${e.sample}(${e.n})`).join(' ');
}

function typeLabel(st) {
  const kinds = [st.numbers > 0, st.dates > 0, st.texts > 0].filter(Boolean).length;
  if (!kinds) return '空';
  if (kinds > 1) return '混合';
  if (st.numbers) return '数字';
  return st.dates ? '日期' : '文本';
}

/** 最小/最大值:有数值就按数值,否则用文本/日期极值;整列为空则留空。 */
function boundCell(st, which) {
  const isMin = which === 'min';
  if (st.numbers > 0) return round(isMin ? st.min : st.max);
  if (!st.dates && !st.texts) return '';
  return isMin ? st.txtMin : st.txtMax;
}

function profileRow(name, st) {
  const distinct = st.capped ? `≥${FREQ_CAP}` : String(st.freq.size);
  const isNum = st.numbers > 0;
  return [
    name,
    typeLabel(st),
    st.numbers + st.dates + st.texts,
    st.blank,
    distinct,
    boundCell(st, 'min'),
    boundCell(st, 'max'),
    isNum ? round(st.sum) : '',
    isNum ? round(st.sum / st.num) : '',
    topValues(st),
  ];
}

const PROFILE_HEADER = ['列', '类型', '非空', '空', '去重', '最小', '最大', '求和', '均值', `高频值(前${TOP_VALUES})`];

function profileBlock(sheet, rows) {
  const states = sheet.columns.map(() => newProfileState());
  for (const row of rows) {
    for (let c = 0; c < states.length; c++) addProfileCell(states[c], row[c]);
  }
  const body = sheet.columns.map((name, c) => profileRow(name, states[c]));
  const blankCols = states.filter((s) => typeLabel(s) === '空').length;
  const blankNote = blankCols ? `；${blankCols} 列整列为空` : '';
  const headerText = sheet.headerRow === 0 ? '（无表头，列名为 A/B/C…）' : '';
  return {
    text: `### ${summarize(sheet, rows.length, '表结构画像')}\n` +
      `> 表头：第 ${sheet.headerRow} 行${headerText}${blankNote}\n\n` + tsvBlock(PROFILE_HEADER, body),
    meta: { kind: 'profile', columns: sheet.columns, blankColumns: blankCols, truncated: false },
  };
}

// ---------------------------------------------------------------------------
// 入口
// ---------------------------------------------------------------------------
/**
 * @param {{name: string, matrix: Array<Array<*>>, truncated?: boolean, totalRows?: number}} raw 单个工作表
 * @param {object} spec office_query 的参数
 * @returns {{text: string, meta: object}} 正文与机器可读元信息
 */
export function querySheet(raw, spec = {}) {
  const matrix = Array.isArray(raw.matrix) ? raw.matrix : [];
  const headerRow = headerRowOf(spec.headerRow, matrix.length);
  const table = buildTable(matrix, headerRow);
  const sheet = {
    name: raw.name,
    columns: table.columns,
    rows: table.rows,
    headerRow,
    truncated: Boolean(raw.truncated),
    totalRows: Number(raw.totalRows) || table.rows.length,
  };
  if (!table.rows.length) {
    return { text: `工作表「${raw.name}」没有数据行（共 ${matrix.length} 行）。可用 office_read 查看原始内容`, meta: { kind: 'empty' } };
  }
  const rows = filterRows(table.rows, compileWhere(spec.where, table.columns));
  const wantsAgg = Boolean(spec.groupBy?.length || spec.aggregate?.length);
  // 画像模式只有「一列一行」这一种形态,orderBy/limit 无从作用 —— 明确报错,别静默忽略
  if (!wantsAgg && (spec.orderBy !== undefined || spec.limit !== undefined)) {
    throw new OfficeError('orderBy / limit 只在分组聚合时有效；要排序就先给 groupBy 或 aggregate', 'INVALID_ARGS');
  }
  const body = wantsAgg ? aggregateBlock(sheet, rows, spec) : profileBlock(sheet, rows);
  const warn = sheet.truncated ? `\n> 注：表实际有 ${sheet.totalRows} 行，受上限限制只统计了前 ${table.rows.length} 行` : '';
  return {
    text: body.text + warn,
    meta: { ...body.meta, sheet: raw.name, scannedRows: table.rows.length, matchedRows: rows.length },
  };
}
