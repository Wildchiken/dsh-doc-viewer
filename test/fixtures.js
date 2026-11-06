/** Builds the documents the engine tests run against, in a temporary folder. */
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from 'docx';
import { PDFDocument, StandardFonts } from 'pdf-lib';

/** Folders the tests made, removed when the test file finishes. */
const made = [];

/** A fresh temporary folder that `removeTempDirs` deletes. @param {string} prefix */
export function tempDir(prefix) {
  const dir = mkdtempSync(join(tmpdir(), prefix));
  made.push(dir);
  return dir;
}

/** Delete every folder `tempDir` made in this process. */
export function removeTempDirs() {
  for (const dir of made.splice(0)) rmSync(dir, { recursive: true, force: true });
}

export async function makeFixtures() {
  const dir = tempDir('doc-viewer-test-');

  const book = new ExcelJS.Workbook();
  const data = book.addWorksheet('数据');
  data.addRow(['区域', '产品', '销量', '单价', '日期', '编号']);
  const regions = ['华东', '华南', '华北', '西南'];
  for (let i = 0; i < 24; i += 1) {
    data.addRow([regions[i % 4], `产品${(i % 3) + 1}`, 100 + i * 7, i % 5 === 0 ? `${20 + i}` : 20 + i, new Date(Date.UTC(2026, i % 12, 15)), `00${i + 1}`]);
  }
  data.getColumn(5).numFmt = 'yyyy-mm-dd';
  data.views = [{ state: 'frozen', ySplit: 1 }];
  const summary = book.addWorksheet('汇总');
  summary.addRow(['指标', '值']);
  summary.addRow(['总销量', { formula: 'SUM(数据!C2:C25)' }]);
  summary.addRow(['华东销量', { formula: 'SUMIF(数据!A2:A25,"华东",数据!C2:C25)' }]);
  summary.mergeCells('A5:B5');
  summary.getCell('A5').value = '合并单元格示例';
  await book.xlsx.writeFile(join(dir, '销售.xlsx'));

  const big = new ExcelJS.Workbook();
  const sheet = big.addWorksheet('大表');
  sheet.addRow(['编号', 'A', 'B', 'C', 'D', 'E', 'F', 'G']);
  for (let i = 1; i <= 50_000; i += 1) sheet.addRow([i, i * 2, i * 3, `文本${i % 97}`, i % 11, i % 13, i / 7, `备注${i % 5}`]);
  await big.xlsx.writeFile(join(dir, '大表.xlsx'));

  writeFileSync(join(dir, '工资.csv'), '﻿姓名,工号,工资\n张三,00123,12000\n李四,00456,9800.5\n王五,13812345678901234,1\n');
  writeFileSync(join(dir, '工资.tsv'), '姓名\t工资\n张三\t100\n李四\t200\n');
  // "姓名,部门\n张三,研发\n李四,市场\n" in GB18030.
  writeFileSync(join(dir, 'gbk.csv'), Buffer.from(new Uint8Array([0xd0, 0xd5, 0xc3, 0xfb, 0x2c, 0xb2, 0xbf, 0xc3, 0xc5, 0x0a, 0xd5, 0xc5, 0xc8, 0xfd, 0x2c, 0xd1, 0xd0, 0xb7, 0xa2, 0x0a, 0xc0, 0xee, 0xcb, 0xc4, 0x2c, 0xca, 0xd0, 0xb3, 0xa1, 0x0a])));

  const word = new Document({
    sections: [{
      children: [
        new Paragraph({ text: '项目周报', heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ children: [new TextRun('本周完成 '), new TextRun({ text: '接口联调', bold: true }), new TextRun(' 与 A&B 测试。')] }),
      ],
    }],
  });
  writeFileSync(join(dir, '周报.docx'), await Packer.toBuffer(word));

  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  for (let p = 1; p <= 8; p += 1) {
    const page = pdf.addPage([595, 842]);
    page.drawText(`Report page ${p}`, { x: 60, y: 780, size: 20, font });
    page.drawText('Quarterly revenue grew in every region during this period.', { x: 60, y: 740, size: 12, font });
  }
  writeFileSync(join(dir, '报告.pdf'), await pdf.save());
  const other = await PDFDocument.create();
  const otherFont = await other.embedFont(StandardFonts.Helvetica);
  other.addPage([595, 842]).drawText('Appendix', { x: 60, y: 780, size: 20, font: otherFont });
  writeFileSync(join(dir, '附录.pdf'), await other.save());

  writeFileSync(join(dir, '笔记.txt'), '第一行\n第二行\n'.repeat(4000));
  writeFileSync(join(dir, '图片.png'), Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1]));
  writeFileSync(join(dir, '坏文件.pdf'), 'this is not a pdf');
  return dir;
}
