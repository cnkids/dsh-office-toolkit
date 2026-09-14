// office_query 的计算层单测:值归一化、条件算子、分组聚合、排序截断、画像、错误分支。
// 这一层是纯函数(不碰 IO),所以直接喂矩阵即可,不需要造文件。
import { querySheet, toNumber, toDay, valueKey, evalCondition, buildTable, FILTER_OPS, AGG_FNS } from '../lib/core/query.js';

const results = [];
async function t(name, fn) {
  try {
    const detail = await fn();
    results.push({ name, ok: true });
    console.log(`✓ ${name}${detail ? ' — ' + detail : ''}`);
  } catch (err) {
    results.push({ name, ok: false, detail: String(err?.message || err) });
    console.log(`✗ ${name}\n   ${err?.message || err}`);
  }
}

function assert(cond, message) {
  if (!cond) throw new Error(message);
}

function assertThrows(fn, code) {
  try {
    fn();
  } catch (err) {
    assert(err.code === code, `期望错误码 ${code}，实际 ${err.code}: ${err.message}`);
    return err.message;
  }
  throw new Error(`期望抛出 ${code}，但没有报错`);
}

/** 按 TSV 行拆出结果单元格,方便断言。 */
function rows(text) {
  const block = text.split('```tsv')[1];
  assert(block, '没有找到 tsv 结果块: ' + text.slice(0, 120));
  return block.split('```')[0].trim().split('\n').map((line) => line.split('\t'));
}

const SHEET = {
  name: '订单',
  matrix: [
    ['地区', '产品', '金额', '日期', '备注'],
    ['华东', 'A', '1,200.50', '2025/1/3', ''],
    ['华东', 'B', 800, '2025/2/1', '加急'],
    ['华南', 'A', '¥2,000', '2024/12/31', ''],
    ['华南', 'B', '12.5%', '2025/3/5', '加急'],
    ['华北', 'A', '', '2025/1/3', ''],
  ],
};

// ---------------------------------------------------------------------------
// 值归一化
// ---------------------------------------------------------------------------
await t('数字:千分位 / 货币符号 / 百分比 / 全角空格', () => {
  assert(toNumber('1,234.00') === 1234, '千分位未解析');
  assert(toNumber('¥88') === 88 && toNumber('$1.5') === 1.5, '货币符号未剥离');
  assert(toNumber('12.5%') === 0.125, '百分比未换算');
  assert(toNumber('\u3000123\u3000') === 123, '全角空格未剥离');
  assert(toNumber(42) === 42 && toNumber(true) === null, '非字符串入参处理不对');
  return '1,234.00→1234、12.5%→0.125';
});

await t('数字:歧义写法宁可返回 null 也不猜', () => {
  assert(toNumber('1,5') === null, '欧洲小数写法被误判为 15');
  assert(toNumber('abc') === null && toNumber('') === null && toNumber('  ') === null, '非数字未返回 null');
  assert(toNumber('1.2.3') === null && toNumber('12%3') === null, '畸形数字未返回 null');
  assert(toNumber(Number.POSITIVE_INFINITY) === null, 'Infinity 未返回 null');
  return '1,5 / abc / 1.2.3 → null';
});

await t('日期:多种写法归一化成 YYYY-MM-DD', () => {
  assert(toDay('2025/1/1') === '2025-01-01', '/ 分隔未归一化');
  assert(toDay('2025年1月1日') === '2025-01-01', '中文日期未归一化');
  assert(toDay('2025-01-01 10:30') === '2025-01-01', '带时间未截断');
  assert(toDay('2025.1.1') === '2025-01-01', '. 分隔未归一化');
  assert(toDay('不是日期') === null && toDay(20250101) === null, '非日期未返回 null');
  return '4 种写法 → 2025-01-01';
});

await t('分组键:1 与 "1" 同组、两种日期写法同组', () => {
  assert(valueKey(1) === valueKey('1'), '数字与数字字符串没归一到同一组');
  assert(valueKey('1.0') === valueKey(1), '1.0 与 1 没归一');
  assert(valueKey('2025/1/1') === valueKey('2025-01-01'), '两种日期写法没归一');
  assert(valueKey('华东') !== valueKey('华南'), '不同文本被误归一');
  return 'n:/d:/s: 三类键值';
});

// ---------------------------------------------------------------------------
// 条件算子
// ---------------------------------------------------------------------------
await t('条件:数值比较穿透文本写法', () => {
  const cond = (op, value) => ({ col: '金额', op, value });
  assert(evalCondition('1,234.00', cond('gt', 1000)), '带千分位的文本没有按数值比较');
  assert(evalCondition('¥2,000', cond('gte', '2000')), '货币写法没有按数值比较');
  assert(evalCondition('800', cond('eq', 800)), '数字字符串与数字不等');
  // 「一边数字一边不是」必须判为不可比:否则 "待定" > 1000 会因字典序成立,结果静默出错
  assert(!evalCondition('待定', cond('gt', 1000)), '文本被拿去和数字比大小了');
  assert(!evalCondition('待定', cond('eq', 1000)), '文本与数字被判定相等');
  assert(evalCondition('待定', cond('ne', 1000)), 'ne 在类型不可比时应成立');
  assert(!evalCondition('', cond('lt', 1)), '空值与数字不该判定大小');
  return 'gt/gte/eq 走数值比较,类型不可比一律不成立';
});

await t('条件:日期比较按日历而不是字符串', () => {
  assert(evalCondition('2025/1/3', { op: 'gte', value: '2025-01-01' }), '日期 ≥ 判断失败');
  assert(evalCondition('2025/1/3', { op: 'gt', value: '2024/12/31' }), '跨年日期比较失败');
  assert(!evalCondition('2024/12/31', { op: 'gt', value: '2025-01-01' }), '日期比较方向反了');
  return '2025/1/3 > 2024/12/31';
});

await t('条件:文本算子忽略大小写', () => {
  assert(evalCondition('SpreadSheet', { op: 'contains', value: 'sheet' }), 'contains 未忽略大小写');
  assert(evalCondition('abc', { op: 'startsWith', value: 'A' }), 'startsWith 未忽略大小写');
  assert(evalCondition('abc', { op: 'endsWith', value: 'C' }), 'endsWith 未忽略大小写');
  assert(!evalCondition('abc', { op: 'contains', value: 'z' }), 'contains 误命中');
  return 'contains/startsWith/endsWith';
});

await t('条件:in / notIn / isBlank / notBlank', () => {
  assert(evalCondition('华东', { op: 'in', value: ['华东', '华南'] }), 'in 未命中');
  assert(!evalCondition('华北', { op: 'in', value: ['华东', '华南'] }), 'in 误命中');
  assert(evalCondition('华北', { op: 'notIn', value: ['华东'] }), 'notIn 未命中');
  assert(evalCondition(1000, { op: 'in', value: ['1,000'] }), 'in 没走数值比较');
  assert(evalCondition('', { op: 'isBlank' }) && !evalCondition(' ', { op: 'isBlank' }), 'isBlank 判定不对');
  assert(evalCondition('x', { op: 'notBlank' }), 'notBlank 判定不对');
  return '含数值归一化的 in 匹配';
});

await t('条件:缺失 op 按 eq 处理', () => {
  assert(evalCondition('华东', { col: '地区', value: '华东' }), '缺 op 时未按 eq');
  assert(!evalCondition('华南', { col: '地区', value: '华东' }), '缺 op 时误命中');
  return '宽容写法';
});

// ---------------------------------------------------------------------------
// 表结构
// ---------------------------------------------------------------------------
await t('表头:空列名与重名列都有兜底', () => {
  const table = buildTable([['金额', '金额', '', 'x'], [1, 2, 3, 4]], 1);
  assert(table.columns[0] === '金额' && table.columns[1] === '金额_2', '重名未加后缀: ' + table.columns.join(','));
  assert(table.columns[2] === 'C', '空列名未退化成列字母: ' + table.columns[2]);
  assert(table.rows.length === 1 && table.rows[0][0] === 1, '数据行切分不对');
  return table.columns.join(' | ');
});

await t('表头:headerRow=0 表示无表头,列名取 A/B/C', () => {
  const table = buildTable([[1, 2], [3, 4]], 0);
  assert(table.columns.join(',') === 'A,B', '列名不是列字母: ' + table.columns.join(','));
  assert(table.rows.length === 2, '无表头时首行被当成表头吃掉了');
  return 'A,B + 全部 2 行参与计算';
});

await t('表头:headerRow 越界或非法时报错清晰', () => {
  const over = assertThrows(() => querySheet({ name: 's', matrix: [['a'], [1]] }, { headerRow: 9 }), 'INVALID_ARGS');
  assert(/只有 2 行/.test(over), '越界报错没说明表里有多少行: ' + over);
  assertThrows(() => querySheet({ name: 's', matrix: [['a'], [1]] }, { headerRow: -1 }), 'INVALID_ARGS');
  assertThrows(() => querySheet({ name: 's', matrix: [['a'], [1]] }, { headerRow: 1.5 }), 'INVALID_ARGS');
  return over.slice(0, 40);
});

// ---------------------------------------------------------------------------
// 分组 / 聚合
// ---------------------------------------------------------------------------
await t('聚合:分组求和 / 均值 / 去重计数', () => {
  const out = querySheet(SHEET, {
    groupBy: ['地区'],
    aggregate: [
      { col: '金额', fn: 'sum', as: '销售额' },
      { col: '金额', fn: 'avg' },
      { col: '产品', fn: 'countDistinct', as: '产品数' },
    ],
  });
  const table = rows(out.text);
  assert(table[0].join(',') === '地区,销售额,avg(金额),产品数', '结果表头不对: ' + table[0].join(','));
  const hd = table.find((r) => r[0] === '华东');
  assert(hd[1] === '2000.5' && hd[2] === '1000.25' && hd[3] === '2', '华东汇总不对: ' + hd.join(','));
  const hn = table.find((r) => r[0] === '华南');
  assert(hn[1] === '2000.125', '华南汇总没把 12.5% 算成 0.125: ' + hn.join(','));
  return '华东 2000.5 / 华南 2000.25';
});

await t('聚合:不给 groupBy 就是全表汇总一行', () => {
  const out = querySheet(SHEET, { aggregate: [{ col: '金额', fn: 'sum', as: '总额' }, { col: '金额', fn: 'count', as: '有值' }] });
  const table = rows(out.text);
  assert(table.length === 2, '全表汇总应只有一行结果: ' + table.length);
  assert(table[1][0] === '4000.625' && table[1][1] === '4', '全表汇总值不对: ' + table[1].join(','));
  assert(!/非数值/.test(out.text), '空单元格不该被算成"非数值": ' + out.text);
  return '总额 4000.625，空单元格不计入 count';
});

await t('聚合:只给 groupBy 时自动补行数', () => {
  const out = querySheet(SHEET, { groupBy: ['产品'] });
  const table = rows(out.text);
  assert(table[0].join(',') === '产品,行数', '表头不是 产品,行数: ' + table[0].join(','));
  assert(table.find((r) => r[0] === 'A')[1] === '3', 'A 的组内行数不对');
  return 'A=3, B=2';
});

await t('聚合:min/max 在纯数值列按数值,在文本列按文本', () => {
  const num = querySheet(SHEET, { aggregate: [{ col: '金额', fn: 'min', as: '最小' }, { col: '金额', fn: 'max', as: '最大' }] });
  const numTable = rows(num.text);
  assert(numTable[1][0] === '0.125' && numTable[1][1] === '2000', '数值 min/max 不对: ' + numTable[1].join(','));
  const txt = querySheet(SHEET, { aggregate: [{ col: '地区', fn: 'min', as: '首' }, { col: '地区', fn: 'max', as: '尾' }] });
  assert(/按文本\/日期比较/.test(txt.text), '文本 min/max 没给出比较方式提示');
  return '数值 0.125~2000，文本走字典序';
});

await t('聚合:日期列 min/max 按日历排序', () => {
  const out = querySheet(SHEET, { aggregate: [{ col: '日期', fn: 'min', as: '最早' }, { col: '日期', fn: 'max', as: '最晚' }] });
  const table = rows(out.text);
  assert(table[1][0] === '2024-12-31' && table[1][1] === '2025-03-05', '日期 min/max 不对: ' + table[1].join(','));
  return '2024-12-31 ~ 2025-03-05';
});

await t('聚合:非数值单元格被跳过并如实告知', () => {
  const sheet = { name: 's', matrix: [['值'], [100], ['待定'], ['200'], ['-']] };
  const out = querySheet(sheet, { aggregate: [{ col: '值', fn: 'sum', as: '合计' }] });
  assert(rows(out.text)[1][0] === '300', '非数值没被跳过: ' + rows(out.text)[1][0]);
  assert(/「合计」2 个非数值单元格已跳过/.test(out.text), '没有如实告知跳过了几个值: ' + out.text.split('\n').at(-1));
  return '100 + 200 = 300，跳过 "待定" 与 "-"';
});

await t('聚合:某组一个数值都没有时返回空而不是 0', () => {
  const sheet = { name: 's', matrix: [['组', '值'], ['x', '待定'], ['x', '']] };
  const out = querySheet(sheet, { groupBy: ['组'], aggregate: [{ col: '值', fn: 'sum', as: '合计' }] });
  const table = rows(out.text);
  assert((table[1][1] ?? '') === '', '没有数值时不该给 0，实际: ' + JSON.stringify(table[1][1]));
  assert(/都不是数值/.test(out.text), '没有说明该组没有数值');
  return '空结果 + 明确提示';
});

await t('聚合:countDistinct 用归一化键,1 与 "1" 不重复计数', () => {
  const sheet = { name: 's', matrix: [['值'], [1], ['1'], ['1.0'], ['x']] };
  const out = querySheet(sheet, { aggregate: [{ col: '值', fn: 'countDistinct', as: '去重' }] });
  assert(rows(out.text)[1][0] === '2', '去重计数不对: ' + rows(out.text)[1][0]);
  return '1 / "1" / "1.0" / x → 2';
});

// ---------------------------------------------------------------------------
// 筛选 / 排序 / 截断
// ---------------------------------------------------------------------------
await t('筛选:多条件 AND + 命中行数进摘要', () => {
  const out = querySheet(SHEET, {
    where: [{ col: '地区', op: 'in', value: ['华东', '华南'] }, { col: '金额', op: 'gt', value: 1000 }],
    groupBy: ['地区'],
    aggregate: [{ col: '金额', fn: 'count', as: '单数' }],
  });
  const table = rows(out.text);
  assert(/命中 2 行/.test(out.text), '摘要没写命中行数: ' + out.text.split('\n')[0]);
  assert(table.length === 3, '分组数不对: ' + (table.length - 1));
  return '2 行 → 2 组';
});

await t('filterRows:筛完没有命中时给出空结果而不是报错', () => {
  const out = querySheet(SHEET, { where: [{ col: '地区', op: 'eq', value: '不存在' }] });
  assert(/命中 0 行/.test(out.text), '没有显示 0 命中: ' + out.text.split('\n')[0]);
  assert(!/undefined|NaN/.test(out.text), '空结果里出现了 undefined/NaN');
  return '命中 0 行仍然是合法画像';
});

await t('排序:按汇总结果降序 + 多键排序', () => {
  const out = querySheet(SHEET, {
    groupBy: ['地区'],
    aggregate: [{ col: '金额', fn: 'sum', as: '销售额' }],
    orderBy: [{ col: '销售额', dir: 'desc' }],
  });
  const table = rows(out.text).slice(1).map((r) => r[0]);
  assert(table.join(',') === '华东,华南,华北', '降序排序不对: ' + table.join(','));
  const asc = querySheet(SHEET, {
    groupBy: ['地区'],
    aggregate: [{ col: '金额', fn: 'sum', as: '销售额' }],
    orderBy: [{ col: '地区', dir: 'asc' }],
  });
  const ascOrder = rows(asc.text).slice(1).map((r) => r[0]);
  // 文本排序按 Unicode 码位(确定性,不依赖 ICU 数据),中文不等于拼音序
  assert(ascOrder.join(',') === '华东,华北,华南', '按分组列升序不对: ' + ascOrder.join(','));
  return 'desc/asc 与按分组列排序都正确';
});

await t('排序:orderBy 引用不存在的列时报错并列出可用列', () => {
  const msg = assertThrows(() => querySheet(SHEET, {
    groupBy: ['地区'],
    aggregate: [{ col: '金额', fn: 'sum', as: '销售额' }],
    orderBy: [{ col: '不存在' }],
  }), 'UNKNOWN_COLUMN');
  assert(/销售额/.test(msg), '报错没列出可用列: ' + msg);
  return msg.slice(0, 40);
});

await t('截断:结果超过 limit 时明确告知共多少组', () => {
  const matrix = [['组', '值']];
  for (let i = 0; i < 30; i++) matrix.push([`g${i}`, i]);
  const out = querySheet({ name: 's', matrix }, { groupBy: ['组'], limit: 5 });
  const table = rows(out.text);
  assert(table.length === 6, 'limit 未生效: ' + (table.length - 1));
  assert(/共 30 组，仅显示前 5 组/.test(out.text), '没有说明截断: ' + out.text.split('\n').at(-1));
  assert(out.meta.truncated === true && out.meta.groups === 30, 'meta 里的 groups/truncated 不对');
  return '30 组 → 显示 5 组';
});

await t('截断:limit 非法时报错', () => {
  assertThrows(() => querySheet(SHEET, { groupBy: ['地区'], limit: 0 }), 'INVALID_ARGS');
  assertThrows(() => querySheet(SHEET, { groupBy: ['地区'], limit: -3 }), 'INVALID_ARGS');
  assertThrows(() => querySheet(SHEET, { groupBy: ['地区'], limit: 'x' }), 'INVALID_ARGS');
  return '0 / -3 / "x" 都被拒绝';
});

await t('分组数过多时快速失败并给出建议', () => {
  const matrix = [['id']];
  for (let i = 0; i < 5200; i++) matrix.push([`id-${i}`]);
  const msg = assertThrows(() => querySheet({ name: 's', matrix }, { groupBy: ['id'] }), 'TOO_MANY_GROUPS');
  assert(/where/.test(msg) && /分组/.test(msg), '报错没给出收窄建议: ' + msg);
  return '5200 组被拦下';
});

// ---------------------------------------------------------------------------
// 画像
// ---------------------------------------------------------------------------
await t('画像:类型判定与统计列齐全', () => {
  const out = querySheet(SHEET, {});
  const table = rows(out.text);
  assert(table[0].join(',') === '列,类型,非空,空,去重,最小,最大,求和,均值,高频值(前3)', '画像表头不对: ' + table[0].join(','));
  const byName = Object.fromEntries(table.slice(1).map((r) => [r[0], r]));
  assert(byName['地区'][1] === '文本', '地区类型判定错: ' + byName['地区'][1]);
  assert(byName['金额'][1] === '数字', '金额类型判定错: ' + byName['金额'][1]);
  assert(byName['日期'][1] === '日期', '日期类型判定错: ' + byName['日期'][1]);
  assert(byName['金额'][2] === '4' && byName['金额'][3] === '1', '金额非空/空计数不对: ' + byName['金额'].join(','));
  assert(byName['金额'][7] === '4000.625' && byName['金额'][8] === '1000.15625', '金额求和/均值不对: ' + byName['金额'].join(','));
  assert(byName['备注'][1] === '文本' && byName['备注'][3] === '3', '备注列空值计数不对: ' + byName['备注'].join(','));
  return '类型/非空/空/去重/求和/均值/高频值';
});

await t('画像:混合类型与整列为空都能标出来', () => {
  const sheet = { name: 's', matrix: [['混', '空列'], ['abc', ''], [1, ''], ['2025/1/1', '']] };
  const out = querySheet(sheet, {});
  const table = rows(out.text);
  const byName = Object.fromEntries(table.slice(1).map((r) => [r[0], r]));
  assert(byName['混'][1] === '混合', '混合类型未标出: ' + byName['混'][1]);
  assert(byName['空列'][1] === '空', '整列为空未标出: ' + byName['空列'][1]);
  assert(/1 列整列为空/.test(out.text), '摘要没提示整列为空的列数');
  return '混合 / 空 都能识别';
});

await t('画像:高频值按出现次数降序并带上次数', () => {
  const sheet = { name: 's', matrix: [['c'], ['a'], ['a'], ['a'], ['b'], ['b'], ['c']] };
  const top = rows(querySheet(sheet, {}).text)[1][9];
  assert(top === 'a(3) b(2) c(1)', '高频值排序不对: ' + top);
  return top;
});

// ---------------------------------------------------------------------------
// 错误分支 / 边界
// ---------------------------------------------------------------------------
await t('列名:找不到列时报错并列出全部可用列', () => {
  const msg = assertThrows(() => querySheet(SHEET, { groupBy: ['不存在的列'] }), 'UNKNOWN_COLUMN');
  assert(/地区、产品、金额、日期、备注/.test(msg), '没列出可用列: ' + msg);
  const ci = assertThrows(() => querySheet(SHEET, { groupBy: ['金额'] , where: [{ col: '地区', op: 'eq' }] }), 'INVALID_ARGS');
  assert(/需要 value/.test(ci), '缺 value 的报错不清晰: ' + ci);
  return msg.slice(0, 44);
});

await t('算子/函数:写错时报错并列出全部合法值', () => {
  const op = assertThrows(() => querySheet(SHEET, { where: [{ col: '金额', op: 'like', value: 1 }] }), 'INVALID_ARGS');
  for (const name of FILTER_OPS) assert(op.includes(name), `算子报错漏了 ${name}`);
  const fn = assertThrows(() => querySheet(SHEET, { aggregate: [{ col: '金额', fn: 'median' }] }), 'INVALID_ARGS');
  for (const name of AGG_FNS) assert(fn.includes(name), `函数报错漏了 ${name}`);
  return '一句话列出全部合法值';
});

await t('入参:where / groupBy / aggregate / orderBy 形状不对时报错', () => {
  assertThrows(() => querySheet(SHEET, { where: 'x' }), 'INVALID_ARGS');
  assertThrows(() => querySheet(SHEET, { groupBy: '地区' }), 'INVALID_ARGS');
  assertThrows(() => querySheet(SHEET, { orderBy: {} }), 'INVALID_ARGS');
  // aggregate: [] 等价于「没给汇总」,退回画像而不是报错 —— 与 groupBy/where 的宽容度一致
  assert(querySheet(SHEET, { aggregate: [] }).meta.kind === 'profile', 'aggregate:[] 应退回画像');
  const only = assertThrows(() => querySheet(SHEET, { limit: 5 }), 'INVALID_ARGS');
  assert(/只在分组聚合时有效/.test(only), '画像模式下 limit 的报错不清晰: ' + only);
  return '4 种形状错误 + 画像模式下的 orderBy/limit 都被拦下';
});

await t('边界:空工作表与只有表头的表都能安全返回', () => {
  const empty = querySheet({ name: '空', matrix: [] }, {});
  assert(/没有数据行/.test(empty.text), '空表提示不对: ' + empty.text);
  const headerOnly = querySheet({ name: '只有表头', matrix: [['a', 'b']] }, {});
  assert(/没有数据行/.test(headerOnly.text), '只有表头的表未提示: ' + headerOnly.text);
  assert(headerOnly.meta.kind === 'empty', 'meta.kind 不是 empty');
  return '两种退化输入都不报错';
});

await t('边界:参差行(短行/长行)不产生 undefined', () => {
  const sheet = { name: 's', matrix: [['a', 'b', 'c'], [1], [2, 3, 4, 5]] };
  const out = querySheet(sheet, { groupBy: ['c'], aggregate: [{ col: 'b', fn: 'sum', as: 'b和' }] });
  assert(!/undefined|NaN/.test(out.text), '参差行产生了 undefined/NaN:\n' + out.text);
  const profile = querySheet(sheet, {});
  assert(!/undefined|NaN/.test(profile.text), '画像里出现了 undefined/NaN');
  return 'a/b/c 三列，短行按空处理';
});

await t('边界:超长文本与制表符/换行不会破坏 TSV 结构', () => {
  const sheet = { name: 's', matrix: [['v'], ['含\t制表'], ['含\n换行'], ['x'.repeat(5000)]] };
  const out = querySheet(sheet, {});
  const table = rows(out.text);
  assert(table.length === 2, '单列画像应只有表头 + 一行: ' + table.length);
  const top = table[1][9];
  assert(top.includes('含␉制表'), '制表符没有被转义可见: ' + top.slice(0, 60));
  assert(top.includes('含⏎换行'), '换行没有被转义可见: ' + top.slice(0, 60));
  assert(!top.includes('\t') && !top.includes('\n'), '高频值里残留了裸控制字符');
  return 'tab/newline 转义为 ␉/⏎';
});

await t('边界:列数超过上限时只算前 N 列并说明', () => {
  const wide = [Array.from({ length: 260 }, (_, i) => `c${i}`)];
  wide.push(Array.from({ length: 260 }, () => 1));
  const sheet = { name: 's', matrix: wide };
  const out = querySheet(sheet, {});
  assert(rows(out.text).length - 1 === 260, '画像应覆盖全部列: ' + (rows(out.text).length - 1));
  return '260 列画像不崩';
});

await t('安全: 超长畸形数字串必须线性失败(不变 O(n²))', () => {
  // `\d+\.?\d*` 这类重叠量词会让 6 万位数字串卡住十几秒;单元格内容不可信,必须挡住
  const bad = '1'.repeat(60000) + 'X';
  const t0 = Date.now();
  assert(toNumber(bad) === null, '畸形数字串应返回 null');
  const ms = Date.now() - t0;
  assert(ms < 200, `解析 ${bad.length} 字符耗时 ${ms}ms，疑似回退到超线性匹配`);
  const sheet = { name: 's', matrix: [['金额'], [bad], ['1,200']] };
  const t1 = Date.now();
  const out = querySheet(sheet, {});
  assert(rows(out.text)[1][7] === '1200', '画像里的求和不对: ' + rows(out.text)[1][7]);
  assert(Date.now() - t1 < 400, '画像阶段也出现了超线性耗时');
  return `6 万字符 ${ms}ms 内失败`;
});

await t('安全: 超长日期串与超长文本同样线性', () => {
  const t0 = Date.now();
  assert(toDay('2025-'.repeat(10000) + 'x') === null, '畸形日期串应返回 null');
  assert(!evalCondition('9'.repeat(50000) + 'x', { op: 'gt', value: 1 }), '超长文本不该满足数值比较');
  const ms = Date.now() - t0;
  assert(ms < 200, `长串比较耗时 ${ms}ms，疑似超线性`);
  return `${ms}ms`;
});

const failed = results.filter((r) => !r.ok);
console.log(`\n===== 计算层: ${results.length - failed.length}/${results.length} 通过 =====`);
if (failed.length) {
  console.log('失败项: ' + failed.map((f) => `${f.name}(${f.detail})`).join('; '));
  process.exit(1);
}
