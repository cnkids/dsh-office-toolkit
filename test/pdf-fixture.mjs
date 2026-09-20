// 测试用的最小合法 PDF(手写对象 + 正确的 xref 偏移):
// 只含标准 14 号字体 Helvetica 的文本,用来验证「抽文本」这条链而不依赖任何样例文件。
const ESCAPE = { '\\': String.raw`\\`, '(': String.raw`\(`, ')': String.raw`\)` };

function streamOf(lines) {
  return lines.map((line, index) => {
    const escaped = line.replaceAll(/[\\()]/g, (ch) => ESCAPE[ch]);
    return `BT /F1 18 Tf 72 ${720 - 30 * index} Td (${escaped}) Tj ET`;
  }).join('\n');
}

/**
 * @param {string[][]} pages 每页若干行文本
 * @returns {Buffer}
 */
export function pdfFixture(pages) {
  const objects = [];
  const pageIds = pages.map((_, index) => 4 + index * 2);
  objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  const kids = pageIds.map((id) => `${id} 0 R`).join(' ');
  objects[2] = `<< /Type /Pages /Kids [${kids}] /Count ${pages.length} >>`;
  objects[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>';
  pages.forEach((lines, index) => {
    const contentId = 5 + index * 2;
    objects[pageIds[index]] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${contentId} 0 R >>`;
    const body = streamOf(lines);
    objects[contentId] = `<< /Length ${body.length} >>\nstream\n${body}\nendstream`;
  });

  let pdf = '%PDF-1.4\n';
  const offsets = [];
  for (let id = 1; id < objects.length; id += 1) {
    if (!objects[id]) continue;
    offsets[id] = pdf.length;
    pdf += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xrefAt = pdf.length;
  const size = objects.length;
  pdf += `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let id = 1; id < size; id += 1) {
    pdf += offsets[id] === undefined
      ? '0000000000 65535 f \n'
      : `${String(offsets[id]).padStart(10, '0')} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${size} /Root 1 0 R >>\nstartxref\n${xrefAt}\n%%EOF\n`;
  return Buffer.from(pdf, 'latin1');
}
