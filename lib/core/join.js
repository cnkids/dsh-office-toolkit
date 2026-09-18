// office_query 的多表 join:在内存里按等值键做哈希连接。
//
// 只做等值连接(Excel 里跨表对照 99% 是「按编号/名称对齐」),key 用 valueKey 归一化,
// 这样 1 与 "1"、2025/1/1 与 2025-01-01 能对上。join 发生在聚合之前,结果再交给
// query.js 做筛选/分组/透视。最多两张表一次(需要三表就串两次,spec 支持数组)。
import { OfficeError } from './util.js';
import { valueKey } from './query.js';

const JOIN_TYPES = ['inner', 'left', 'right', 'full'];
/** 一次 join 最多落多少行,避免笛卡尔积把内存撑爆。 */
const MAX_JOINED_ROWS = 2_000_000;

const text = (v) => (v === null || v === undefined ? '' : String(v));

// ---------------------------------------------------------------------------
// 参数解析
// ---------------------------------------------------------------------------
/** `on` 可以是列名、{left,right} 或它们的数组。 */
function joinKeys(on) {
  if (on === undefined || on === null || on === '') {
    throw new OfficeError('join 需要 on（两表用来对齐的列，如 {"left":"客户编号","right":"编号"}）', 'INVALID_ARGS');
  }
  const list = Array.isArray(on) ? on : [on];
  if (!list.length) throw new OfficeError('join 需要 on（两表用来对齐的列，如 {"left":"客户编号","right":"编号"}）', 'INVALID_ARGS');
  return list.map((item) => {
    if (typeof item === 'string') return { left: item, right: item };
    const left = text(item?.left ?? item?.col).trim();
    const right = text(item?.right ?? item?.col2 ?? left).trim();
    if (!left || !right) throw new OfficeError('join 的 on 需形如 "客户编号" 或 {"left":"客户编号","right":"编号"}', 'INVALID_ARGS');
    return { left, right };
  });
}

/** join 规格(单条或数组)→ 规范化后的列表。 */
export function planJoins(join, mainPath) {
  if (join === undefined || join === null) return [];
  const list = Array.isArray(join) ? join : [join];
  return list.map((item) => {
    if (!item || typeof item !== 'object') throw new OfficeError('join 需为对象或对象数组', 'INVALID_ARGS');
    const type = text(item.type).trim() || 'inner';
    if (!JOIN_TYPES.includes(type)) {
      throw new OfficeError(`join 的 type 可为 ${JOIN_TYPES.join(' / ')}，收到「${type}」`, 'INVALID_ARGS');
    }
    const path = text(item.path).trim() || mainPath;
    if (!path) throw new OfficeError('join 需要 path（另一张表所在文件）', 'INVALID_ARGS');
    return {
      path,
      sheet: item.sheet,
      headerRow: item.headerRow,
      keys: joinKeys(item.on),
      type,
      suffix: item.suffix === undefined || item.suffix === null ? '_2' : String(item.suffix),
      label: item.as === undefined ? '' : String(item.as),
    };
  });
}

// ---------------------------------------------------------------------------
// 连接
// ---------------------------------------------------------------------------
function findColumn(columns, name, side) {
  const at = columns.indexOf(name);
  if (at >= 0) return at;
  const lower = name.toLowerCase();
  const ci = columns.findIndex((c) => c.toLowerCase() === lower);
  if (ci < 0) throw new OfficeError(`${side}表里找不到列「${name}」。可用列名：${columns.join('、')}`, 'UNKNOWN_COLUMN');
  return ci;
}

/**
 * 找出被连接的列序号,并决定结果里保留右表的哪些列。
 * 左右同名的连接键只保留一份(与 SQL 的 USING、pandas 的 on= 一致),其余重名列加后缀。
 */
export function resolveJoinColumns(leftCols, rightCols, spec) {
  const pairs = spec.keys.map((k) => {
    const left = findColumn(leftCols, k.left, '左');
    const right = findColumn(rightCols, k.right, '右');
    return { left, right, same: leftCols[left] === rightCols[right] };
  });
  const dropped = new Set(pairs.filter((p) => p.same).map((p) => p.right));
  const used = new Set(leftCols);
  const keep = [];
  const names = [];
  rightCols.forEach((name, index) => {
    if (dropped.has(index)) return;
    const base = spec.label ? `${spec.label}·${name}` : name;
    let out = base;
    let n = 2;
    while (used.has(out)) { out = `${base}${spec.suffix}${n > 2 ? n : ''}`; n += 1; }
    used.add(out);
    keep.push(index);
    names.push(out);
  });
  return { pairs, keep, names };
}

function keyOf(row, pairs, side) {
  const parts = [];
  for (const pair of pairs) {
    const value = row[pair[side]];
    if (value === null || value === undefined || value === '') return null; // 空键不参与连接(与 SQL 一致)
    parts.push(valueKey(value));
  }
  return parts.join('\u0000');
}

function indexRight(rows, pairs) {
  const index = new Map();
  for (const row of rows) {
    const key = keyOf(row, pairs, 'right');
    if (key === null) continue;
    const bucket = index.get(key);
    if (bucket) bucket.push(row);
    else index.set(key, [row]);
  }
  return index;
}

/**
 * 等值连接两张表。
 * @param {{name,columns,rows}} left
 * @param {{name,columns,rows}} right
 * @param {object} spec planJoins() 里的一条
 * @returns {{columns, rows, notes, matched, unmatchedLeft, unmatchedRight}}
 */
function pushJoined(ctx, rows, leftRow, rightRow) {
  if (rows.length >= MAX_JOINED_ROWS) {
    throw new OfficeError(`连接结果超过 ${MAX_JOINED_ROWS} 行（连接键可能不唯一，出现了笛卡尔积）。请先用 where / 聚合收窄，或换一个更唯一的键`, 'TOO_MANY_ROWS');
  }
  const leftCells = leftRow ?? backfillKeys(ctx, rightRow);
  const rightCells = ctx.keep.map((at) => (rightRow ? rightRow[at] : null));
  rows.push([...leftCells, ...rightCells]);
}

/** 右表孤儿行的连接键回填到左表那一列(同名列只保留一份,不回填就会丢键值)。 */
function backfillKeys({ leftWidth, pairs }, rightRow) {
  const filled = new Array(leftWidth).fill(null);
  for (const pair of pairs) if (pair.same) filled[pair.left] = rightRow[pair.right];
  return filled;
}

/** 左表逐行走一遍:匹配上的展开成多行,没匹配的按连接类型决定补空还是丢弃。 */
function joinLeftSide(leftRows, index, seen, ctx, push) {
  const keepLeft = ctx.type === 'left' || ctx.type === 'full';
  let matched = 0;
  let unmatchedLeft = 0;
  for (const row of leftRows) {
    const key = keyOf(row, ctx.pairs, 'left');
    const hits = key === null ? undefined : index.get(key);
    if (hits) {
      matched += 1;
      for (const hit of hits) {
        seen.add(hit);
        push(row, hit);
      }
      continue;
    }
    unmatchedLeft += 1;
    if (keepLeft) push(row, null);
  }
  return { matched, unmatchedLeft };
}

export function joinTables(left, right, spec) {
  const { pairs, keep, names } = resolveJoinColumns(left.columns, right.columns, spec);
  const ctx = { pairs, keep, leftWidth: left.columns.length, type: spec.type };
  const seen = new Set();
  const rows = [];
  const push = (leftRow, rightRow) => pushJoined(ctx, rows, leftRow, rightRow);
  const { matched, unmatchedLeft } = joinLeftSide(left.rows, indexRight(right.rows, pairs), seen, ctx, push);
  let unmatchedRight = 0;
  if (spec.type === 'right' || spec.type === 'full') {
    for (const row of right.rows) {
      if (seen.has(row)) continue;
      unmatchedRight += 1;
      push(null, row);
    }
  }
  return { columns: [...left.columns, ...names], rows, matched, unmatchedLeft, unmatchedRight };
}

/** 连接结果的说明文字(放在结果里,便于核对有没有漏配)。 */
export function joinNote(spec, right, result) {
  const how = { inner: '内连接(只留两边都匹配的)', left: '左连接(保留左表全部)', right: '右连接(保留右表全部)', full: '全连接(两边都保留)' }[spec.type];
  const name = spec.label || right.name;
  const parts = [`与「${name}」${how}`, `匹配 ${result.matched} 行`];
  if (result.unmatchedLeft) parts.push(`左表 ${result.unmatchedLeft} 行没配上`);
  if (result.unmatchedRight) parts.push(`右表 ${result.unmatchedRight} 行没配上`);
  return parts.join('、');
}
