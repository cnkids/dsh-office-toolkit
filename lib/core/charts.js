// OOXML chart injection: adds bar/column/line/pie charts into an .xlsx produced
// by ExcelJS, by wiring zip parts the way openpyxl does (no cached values —
// Excel/WPS rebuild them on open). Pure JS via PizZip.
import PizZip from 'pizzip';
import { OfficeError } from './util.js';

const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

const CHART_TYPES = new Set(['bar', 'column', 'line', 'pie']);
const COL_W_EMU = 640000;   // ~1 default column
const ROW_H_EMU = 300000;   // ~1 default row
const PX2EMU = 9525;

function escXml(s) {
  return String(s ?? '')
    .replaceAll('&', '&amp;').replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;').replaceAll('"', '&quot;');
}
function quoteSheet(name) {
  return `'${String(name).replaceAll("'", "''")}'`;
}
function a1ToIndexes(ref) {
  const m = /^([A-Z]+)(\d+)$/i.exec(String(ref));
  if (!m) throw new OfficeError(`无效单元格引用: ${ref}`);
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.codePointAt(0) - 64);
  return { col: col - 1, row: Number(m[2]) - 1 };
}
function rangeRef(sheetName, range) {
  // Accept "A2:A6", "销售明细!A2:A6", or "销售明细!$A$2:$A$6"
  const m = /^(?:(.*?)!)?\$?([A-Za-z]+\$?\d+):\$?([A-Za-z]+\$?\d+)$/.exec(String(range).trim());
  if (!m) throw new OfficeError(`无效范围: ${range}（应为 A2:A6 或 工作表!A2:A6）`);
  const sheet = m[1] ? m[1].replaceAll(/^'|'$/g, '') : sheetName;
  return `${quoteSheet(sheet)}!${m[2].toUpperCase()}:${m[3].toUpperCase()}`;
}

function buildSeriesXml(idx, labelRange, catRange, valRange) {
  const ser = [
    '<c:ser>',
    `<c:idx val="${idx}"/>`,
    `<c:order val="${idx}"/>`,
    '<c:tx><c:strRef>',
    `<c:f>${escXml(labelRange)}</c:f>`,
    '<c:strCache><c:ptCount val="0"/></c:strCache>',
    '</c:strRef></c:tx>',
    '<c:cat><c:strRef>',
    `<c:f>${escXml(catRange)}</c:f>`,
    '<c:strCache><c:ptCount val="0"/></c:strCache>',
    '</c:strRef></c:cat>',
    '<c:val><c:numRef>',
    `<c:f>${escXml(valRange)}</c:f>`,
    '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="0"/></c:numCache>',
    '</c:numRef></c:val>',
    '</c:ser>',
  ];
  return ser.join('');
}

function seriesRefs(spec, sheetName, s, i) {
  if (typeof s === 'string') {
    return { valRange: rangeRef(sheetName, s), labelRange: `${quoteSheet(sheetName)}!${singleCellOfRange(s)}` };
  }
  if (!s.range) throw new OfficeError(`series[${i}] 需要 range`);
  const label = s.label ? `${quoteSheet(sheetName)}!${s.label}` : `${quoteSheet(sheetName)}!${singleCellOfRange(s.range)}`;
  return { valRange: rangeRef(sheetName, s.range), labelRange: label };
}

function chartXml(spec) {
  const type = String(spec.chartType || 'bar').toLowerCase() === 'column' ? 'bar' : String(spec.chartType || 'bar').toLowerCase();
  if (!CHART_TYPES.has(type === 'column' ? 'bar' : type)) throw new OfficeError(`图表类型仅支持 bar/column/line/pie，实际 ${spec.chartType}`);
  const sheetName = spec.sheetName;
  const catRange = rangeRef(sheetName, spec.categories);
  const series = [];
  (spec.series || []).forEach((s, i) => {
    const { valRange, labelRange } = seriesRefs(spec, sheetName, s, i);
    series.push(buildSeriesXml(i, labelRange, catRange, valRange));
  });
  if (!series.length) throw new OfficeError('图表至少需要一个 series.range');
  const title = spec.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>${escXml(spec.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '';

  let plot = '';
  const axisIds = ['1001', '1002'];
  if (type === 'pie') {
    plot =
      '<c:pieChart><c:varyColors val="1"/>' + series.join('') + '</c:pieChart>';
  } else if (type === 'line') {
    plot =
      '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
      series.join('') +
      '<c:marker val="1"/>' +
      `<c:axId val="${axisIds[0]}"/><c:axId val="${axisIds[1]}"/>` +
      '</c:lineChart>' +
      catAx(axisIds[0], axisIds[1]) + valAx(axisIds[1], axisIds[0]);
  } else {
    const dir = String(spec.chartType).toLowerCase() === 'column' ? 'col' : 'bar';
    plot =
      `<c:barChart><c:barDir val="${dir}"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
      series.join('') +
      '<c:gapWidth val="150"/>' +
      `<c:axId val="${axisIds[0]}"/><c:axId val="${axisIds[1]}"/>` +
      '</c:barChart>' +
      catAx(axisIds[0], axisIds[1]) + valAx(axisIds[1], axisIds[0]);
  }

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${NS_REL}">` +
    '<c:lang val="zh-CN"/>' +
    '<c:chart>' +
    title +
    '<c:autoTitleDeleted val="0"/>' +
    '<c:plotArea><c:layout/>' +
    plot +
    '</c:plotArea>' +
    '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>' +
    '</c:chart>' +
    '</c:chartSpace>'
  );
}

function catAx(id, cross) {
  return (
    `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>` +
    `<c:crossAx val="${cross}"/></c:catAx>`
  );
}
function valAx(id, cross) {
  return (
    `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>` +
    `<c:crossAx val="${cross}"/></c:valAx>`
  );
}

function singleCellOfRange(range) {
  const m = /([A-Za-z]+\d+)/.exec(String(range));
  return m ? `$${m[1].toUpperCase()}` : '';
}

const EMPTY_DRAWING =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}"></xdr:wsDr>`;

/** One chart frame; `shapeId` must be unique inside its drawing part. */
function anchorXml(spec, chartRid, shapeId) {
  const { col, row } = a1ToIndexes(spec.anchor || 'A1');
  const cx = Math.round((spec.widthPx || 600) * PX2EMU);
  const cy = Math.round((spec.heightPx || 340) * PX2EMU);
  return (
    '<xdr:oneCellAnchor>' +
    `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>${Math.round(COL_W_EMU * 0.1)}</xdr:colOff>` +
    `<xdr:row>${row}</xdr:row><xdr:rowOff>${Math.round(ROW_H_EMU * 0.1)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/>` +
    '<xdr:graphicFrame macro="">' +
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="${shapeId}" name="${escXml(spec.title || '图表')}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm>` +
    '<a:graphic><a:graphicData uri="' + C_NS + '">' +
    `<c:chart xmlns:c="${C_NS}" xmlns:r="${NS_REL}" r:id="${chartRid}"/>` +
    '</a:graphicData></a:graphic>' +
    '</xdr:graphicFrame>' +
    '<xdr:clientData/>' +
    '</xdr:oneCellAnchor>'
  );
}

// ---------------------------------------------------------------------------
// OOXML package plumbing
// ---------------------------------------------------------------------------
/** Attribute lookup that never matches a longer name ending in `name`. */
function attrValue(tag, name) {
  const m = new RegExp(String.raw`(?:^|\s)${name}="([^"]*)"`).exec(String(tag));
  return m ? m[1] : undefined;
}

function relTags(relsXml) {
  return [...String(relsXml || '').matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0]);
}

function nextRelId(relsXml) {
  const ids = [...String(relsXml || '').matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  return `rId${(ids.length ? Math.max(...ids) : 0) + 1}`;
}

function addRelationship(relsXml, id, type, target) {
  const entry = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  if (String(relsXml || '').trim()) return String(relsXml).replace(/<\/Relationships>\s*$/, entry + '</Relationships>');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS_REL}">${entry}</Relationships>`;
}

/** `xl/worksheets/sheet1.xml` -> `xl/worksheets/_rels/sheet1.xml.rels`. */
function relsPathFor(partPath) {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
}

/** Resolve a relationship target (relative or absolute) against its owner part. */
function resolvePart(ownerPart, target) {
  if (String(target).startsWith('/')) return String(target).slice(1);
  const dir = ownerPart.slice(0, ownerPart.lastIndexOf('/'));
  const out = [];
  for (const seg of `${dir}/${target}`.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}

function relTargetOf(relsXml, rid, type) {
  const tag = relTags(relsXml).find((t) => attrValue(t, 'Id') === rid && attrValue(t, 'Type') === type);
  return tag ? attrValue(tag, 'Target') : undefined;
}

function nextPartNumber(zip, pattern) {
  const numbers = [...zip.file(pattern)].map((f) => Number(/\d+/.exec(f.name.split('/').pop() || '')?.[0] || 0));
  return (numbers.length ? Math.max(...numbers) : 0) + 1;
}

function nextShapeId(drawingXml) {
  const ids = [...String(drawingXml).matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]));
  return (ids.length ? Math.max(...ids) : 1) + 1;
}

function openZip(buf) {
  try {
    return new PizZip(buf);
  } catch (err) {
    throw new OfficeError('无法解压 xlsx: ' + (err?.message || err), 'BAD_XLSX');
  }
}

function sheetTags(wbXml) {
  return [...wbXml.matchAll(/<sheet\b[^>]*\/?>/g)].map((m) => m[0]);
}

function locateSheetIndex(entries, selector) {
  const index = typeof selector === 'number'
    ? selector - 1
    : entries.findIndex((t) => attrValue(t, 'name') === String(selector));
  if (index < 0 || index >= entries.length) {
    throw new OfficeError(`工作表不存在: ${selector}（可用: ${entries.map((t) => attrValue(t, 'name')).join('、')}）`);
  }
  return index;
}

function resolveSheet(zip, spec) {
  const wbXml = zip.file('xl/workbook.xml')?.asText();
  const wbRels = zip.file('xl/_rels/workbook.xml.rels')?.asText();
  if (!wbXml || !wbRels) throw new OfficeError('缺少 xl/workbook.xml 或 rels', 'BAD_XLSX');
  const entries = sheetTags(wbXml);
  const index = locateSheetIndex(entries, spec.sheet);
  const fallback = `xl/worksheets/sheet${index + 1}.xml`;
  const rawTarget = relTargetOf(wbRels, attrValue(entries[index], 'r:id'), `${NS_REL}/worksheet`);
  const path = rawTarget ? resolvePart('xl/workbook.xml', rawTarget) : fallback;
  const file = zip.file(path);
  if (!file) throw new OfficeError(`找不到工作表部件 ${path}`, 'BAD_XLSX');
  return { path, name: attrValue(entries[index], 'name'), file };
}

/** Reuse the sheet's drawing part when it already has one. */
function existingDrawing(zip, sheet) {
  const rid = /<drawing\b[^>]*r:id="([^"]+)"/.exec(sheet.file.asText())?.[1];
  if (!rid) return null;
  const target = relTargetOf(zip.file(relsPathFor(sheet.path))?.asText(), rid, `${NS_REL}/drawing`);
  if (!target) return null;
  const path = resolvePart(sheet.path, target);
  const file = zip.file(path);
  if (!file) return null;
  return { path, relsPath: relsPathFor(path), xml: file.asText(), linked: true };
}

function newDrawing(zip, sheet) {
  const no = nextPartNumber(zip, /^xl\/drawings\/drawing\d+\.xml$/);
  return {
    path: `xl/drawings/drawing${no}.xml`,
    relsPath: `xl/drawings/_rels/drawing${no}.xml.rels`,
    xml: EMPTY_DRAWING,
    linked: false,
    sheetRid: nextRelId(zip.file(relsPathFor(sheet.path))?.asText()),
    relTarget: `../drawings/drawing${no}.xml`,
  };
}

function resolveDrawing(zip, sheet) {
  return existingDrawing(zip, sheet) || newDrawing(zip, sheet);
}

/** Point the sheet at its drawing part (only needed for a freshly made one). */
function linkSheetToDrawing(zip, sheet, drawing) {
  const relsPath = relsPathFor(sheet.path);
  zip.file(relsPath, addRelationship(zip.file(relsPath)?.asText(), drawing.sheetRid, `${NS_REL}/drawing`, drawing.relTarget));
  let xml = sheet.file.asText();
  if (!/<worksheet[^>]*xmlns:r=/.test(xml)) xml = xml.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS_REL}"`);
  zip.file(sheet.path, xml.replace(/<\/worksheet>\s*$/, `<drawing r:id="${drawing.sheetRid}"/></worksheet>`));
}

function addContentTypes(zip, entries) {
  const path = '[Content_Types].xml';
  let xml = zip.file(path)?.asText();
  if (!xml) throw new OfficeError('缺少 [Content_Types].xml', 'BAD_XLSX');
  for (const [part, type] of entries) {
    if (xml.includes(`PartName="/${part}"`)) continue;
    xml = xml.replace(/<\/Types>\s*$/, `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  }
  zip.file(path, xml);
}

/**
 * Inject a chart into an xlsx buffer.
 *
 * OOXML allows at most one `<drawing>` per worksheet, so every chart on the
 * same sheet is appended as an extra anchor inside that one drawing part
 * (creating the part only when the sheet has none).
 *
 * @param buf  xlsx bytes (exceljs output)
 * @param spec { sheet: sheetNameOrIndex(1-based), chartType: bar|column|line|pie,
 *   title?, categories: 'A2:A6', series: ['B2:B6'] or [{range,label?}],
 *   anchor: 'B2', widthPx?, heightPx? }
 * @returns new xlsx buffer
 */
export async function injectChart(buf, spec) {
  const zip = openZip(Buffer.isBuffer(buf) ? buf : Buffer.from(buf));
  const sheet = resolveSheet(zip, spec);
  const full = { ...spec, sheetName: sheet.name };
  const chartPath = `xl/charts/chart${nextPartNumber(zip, /^xl\/charts\/chart\d+\.xml$/)}.xml`;
  zip.file(chartPath, chartXml({ ...full, chartType: String(spec.chartType || 'bar').toLowerCase() }));

  const drawing = resolveDrawing(zip, sheet);
  const chartRid = nextRelId(zip.file(drawing.relsPath)?.asText());
  zip.file(drawing.relsPath, addRelationship(
    zip.file(drawing.relsPath)?.asText(), chartRid, `${NS_REL}/chart`, `../charts/${chartPath.split('/').pop()}`
  ));
  zip.file(drawing.path, drawing.xml.replace(/<\/xdr:wsDr>\s*$/, anchorXml(full, chartRid, nextShapeId(drawing.xml)) + '</xdr:wsDr>'));
  if (!drawing.linked) linkSheetToDrawing(zip, sheet, drawing);

  addContentTypes(zip, [[drawing.path, CT_DRAWING], [chartPath, CT_CHART]]);
  return Buffer.from(zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

/** Test-only helper: verify a buffer holds injected chart parts. */
export function chartParts(buf) {
  let zip;
  try {
    zip = new PizZip(Buffer.from(buf));
  } catch {
    return [];
  }
  return zip.file(/xl\/(charts\/chart\d+\.xml|drawings\/drawing\d+\.xml)/).map((f) => f.name);
}
