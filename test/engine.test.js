/**
 * Engine tests against real files. LibreOffice is not assumed: formula values
 * are only asserted where the engine computes them itself.
 */
import { after, before, describe, test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync, statSync } from 'node:fs';
import { dirname, join } from 'node:path';
import ExcelJS from 'exceljs';
import { makeFixtures, removeTempDirs, tempDir } from './fixtures.js';
import {
  convertMany,
  documentKind,
  editFile,
  openDocument,
  pdfOperation,
  readDocument,
  saveSpreadsheetCopy,
  statsForFile,
} from '../src/host/documents.js';
import { computeStats, fileFormula, frozenSplit, modelFormula, parseNumericText, withCelldata, workbookSheetNames } from '../src/host/spreadsheet.js';
import { ackOpenRequests, enqueueOpen, listOpenRequests } from '../src/host/requests.js';
import { listVersions, restoreVersion } from '../src/host/versions.js';
import { accessSync, chmodSync, existsSync, constants, copyFileSync, lstatSync, mkdtempSync, symlinkSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { calculateSheets, changeReport, createDocument, inPlaceBlocker, saveInPlace, saveTextInPlace } from '../src/host/documents.js';
import { displayValue } from '../src/host/spreadsheet.js';
import { ensureExecutable } from '../src/host/pdf.js';
import { textToDocx } from '../src/host/print.js';
import { describeChanges, registerTools } from '../src/host/tools.js';
import { liteEvaluator } from '../src/host/formula-lite.js';
import mammoth from 'mammoth';

let dir;
let out;
const noRecalc = { recalc: false };

before(async () => {
  // Versions go to a throwaway Harness home, never the real one.
  process.env.DSH_HOME = tempDir('doc-viewer-home-');
  dir = await makeFixtures();
  out = join(dir, '输出');
});
after(() => {
  rmSync(dir, { recursive: true, force: true });
  removeTempDirs();
});

async function readBack(path) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path);
  return workbook;
}

describe('PDF export inside the Harness', () => {
  test('the officeToPdf provider sees the same version it was promised', async () => {
    const { convertOfficeToPdf } = await import('../src/host/libreoffice.js');
    const provider = {
      async convert({ source }) {
        await new Promise((resolve) => setTimeout(resolve, 5));
        const read = await source.read(undefined, Number.MAX_SAFE_INTEGER);
        if (read.version !== source.version) throw new Error('The source changed while waiting for conversion.');
        return { pdf: new Uint8Array([37, 80, 68, 70]), missingFonts: [] };
      },
    };
    const output = join(tempDir('doc-viewer-provider-'), 'out.pdf');
    const result = await convertOfficeToPdf({ get: (name) => (name === 'officeToPdf' ? provider : undefined) }, join(dir, '销售.xlsx'), output, 'xlsx');
    assert.equal(result.backend, 'officeToPdf');
    assert.equal(readFileSync(output, 'latin1'), '%PDF');
  });
});

describe('number formats', () => {
  test('a cell shows its value in the format the file gives it', () => {
    assert.equal(displayValue({ v: 1990, ct: { fa: '#,##0.00', t: 'n' } }), '1,990.00');
    assert.equal(displayValue({ v: 0.256, ct: { fa: '0.0%', t: 'n' } }), '25.6%');
    assert.equal(displayValue({ v: 46000, ct: { fa: 'yyyy"年"m"月"d"日"', t: 'd' } }), '2025年12月9日');
  });

  test('values and saved formula results are formatted when a workbook opens', async () => {
    const path = join(tempDir('doc-viewer-format-'), '格式.xlsx');
    const book = new ExcelJS.Workbook();
    const sheet = book.addWorksheet('表');
    sheet.getCell('A1').value = 2000;
    sheet.getCell('A1').numFmt = '#,##0.00';
    sheet.getCell('B1').value = { formula: 'A1*2', result: 4000 };
    sheet.getCell('B1').numFmt = '#,##0.00';
    await book.xlsx.writeFile(path);
    const { sheets, stats } = await openDocument(path, noRecalc);
    assert.deepEqual(sheets[0].celldata.map((entry) => entry.v.m), ['2,000.00', '4,000.00']);
    assert.equal(stats.unresolved, false, 'saved results leave the grid nothing to compute');
    const read = await readDocument(path, { range: 'A1:B1' });
    assert.equal(read.rows[0][1].display, '4,000.00');
  });

  test('the page can open a workbook before LibreOffice works out missing results', async () => {
    const path = join(tempDir('doc-viewer-later-'), '没算过.xlsx');
    const book = new ExcelJS.Workbook();
    const sheet = book.addWorksheet('表');
    sheet.getCell('A1').value = 2;
    sheet.getCell('B1').value = { formula: 'A1*3' };
    await book.xlsx.writeFile(path);
    const { stats } = await openDocument(path, { recalc: 'later' });
    assert.equal(stats.recalculated, false, 'nothing waits for LibreOffice');
    assert.equal(stats.pendingCalculation, true, 'the page knows to ask again');
    assert.equal(stats.unresolved, true, 'the grid computes what it can meanwhile');
    const saved = join(tempDir('doc-viewer-later-'), '算过.xlsx');
    sheet.getCell('B1').value = { formula: 'A1*3', result: 6 };
    await book.xlsx.writeFile(saved);
    assert.equal((await openDocument(saved, { recalc: 'later' })).stats.pendingCalculation, false, 'saved results need no second look');
  });
});

describe('what an edit reports', () => {
  const cell = (v, f) => ({ v, m: String(v), ...(f ? { f } : {}) });
  const change = (address, before, after) => ({ sheet: '表', r: 0, c: 0, address, before, after });

  test('each written cell before and after, and the formulas that moved because of it', () => {
    const changes = changeReport(
      [change('A1', cell(2), cell(5)), change('B1', cell(20, '=A1*10'), cell(50, '=A1*10')), change('B3', cell(50, '=SUM(B1:B2)'), cell(80, '=SUM(B1:B2)'))],
      [{ sheet: '表', address: 'A1' }],
    );
    assert.deepEqual(changes.written.map(({ address, before, after }) => [address, before, after]), [['A1', '2', '5']]);
    assert.deepEqual(changes.dependents.map(({ address }) => address), ['B1', 'B3']);
    const lines = describeChanges(changes).join('\n');
    assert.match(lines, /表!A1：2 → 5/);
    assert.match(lines, /共 2 处/);
    assert.match(lines, /表!B3（=SUM\(B1:B2\)）：50 → 80/);
  });

  test('a formula LibreOffice could not work out is not reported as having become empty', () => {
    const changes = changeReport(
      [change('A1', cell(2), cell(5)), change('B1', cell(20, '=A1*10'), { f: '=A1*10' })],
      [{ sheet: '表', address: 'A1' }],
    );
    assert.equal(changes.dependents.length, 0);
    assert.equal(changes.uncalculated, 1);
    const lines = describeChanges(changes).join('\n');
    assert.doesNotMatch(lines, /→ 空/);
    assert.match(lines, /1 处公式这次没能算出结果/);
  });

  test('a real edit reports what it wrote, with or without LibreOffice', async () => {
    const path = join(tempDir('doc-viewer-report-'), '报告.xlsx');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('表');
    sheet.getCell('A1').value = 2;
    sheet.getCell('B1').value = { formula: 'A1*10', result: 20 };
    await workbook.xlsx.writeFile(path);
    const { changes } = await editFile(path, { edits: [{ address: 'A1', value: 5 }] }, dirname(path));
    assert.deepEqual(changes.written.map(({ address, before, after }) => [address, before, after]), [['A1', '2', '5']]);
    assert.equal(changes.dependents.length + changes.uncalculated, 1);
  });

  test('a long list keeps its start and its end, where totals usually are', () => {
    const dependents = Array.from({ length: 100 }, (_, index) => ({ sheet: '表', address: `E${index + 2}`, before: '1', after: '2' }));
    const lines = describeChanges({ written: [], dependents });
    assert.match(lines.join('\n'), /共 100 处/);
    assert.match(lines.join('\n'), /中间另有 70 处/);
    assert.match(lines.at(-1), /E101/);
  });
});

describe('saving a text document', () => {
  const file = (name, bytes) => {
    const path = join(tempDir('doc-viewer-text-'), name);
    writeFileSync(path, bytes);
    return path;
  };

  test('goes back to the file, with the old content kept as a version', async () => {
    const path = file('笔记.md', '# 旧\n');
    const { mtime, inPlace } = await saveTextInPlace({ path, text: '# 新\n', openedMtime: statSync(path).mtimeMs });
    assert.equal(inPlace, true);
    assert.equal(mtime, statSync(path).mtimeMs);
    assert.equal(readFileSync(path, 'utf8'), '# 新\n');
    assert.equal(listVersions(path).length, 1);
  });

  test('keeps the encoding, byte order mark and line endings the file had', async () => {
    const crlf = file('crlf.txt', '一\r\n二\r\n');
    await saveTextInPlace({ path: crlf, text: '一\n改\n' });
    assert.equal(readFileSync(crlf, 'utf8'), '一\r\n改\r\n');
    const bom = file('bom.txt', Buffer.from('\ufeff旧', 'utf8'));
    await saveTextInPlace({ path: bom, text: '新' });
    assert.equal(readFileSync(bom, 'utf8'), '\ufeff新');
    const gbk = file('gbk.txt', Buffer.from([0xd6, 0xd0, 0xce, 0xc4, 0xb2, 0xe2, 0xca, 0xd4, 0xce, 0xc4, 0xb1, 0xbe]));
    const { warnings } = await saveTextInPlace({ path: gbk, text: '中文改过了' });
    assert.deepEqual(warnings, []);
    assert.equal(readFileSync(gbk).length, 10, 'two bytes a character, as GB18030 has them');
  });

  test('a file changed since it was opened is left alone until the user says otherwise', async () => {
    const path = file('别处改.txt', '原来');
    const opened = statSync(path).mtimeMs;
    writeFileSync(path, '别处改的');
    utimesSync(path, new Date(), new Date(Date.now() + 5000));
    assert.deepEqual(await saveTextInPlace({ path, text: '我的', openedMtime: opened }), { conflict: true });
    assert.equal(readFileSync(path, 'utf8'), '别处改的');
    await saveTextInPlace({ path, text: '我的' });
    assert.equal(readFileSync(path, 'utf8'), '我的');
    assert.equal(listVersions(path).length, 1, 'what was overwritten can be restored');
  });

  test('a file too long to have been loaded whole is never cut short', async () => {
    const whole = 'x'.repeat(2 * 1024 * 1024 + 10);
    const path = file('很长.txt', whole);
    await assert.rejects(saveTextInPlace({ path, text: 'x' }), /太大/);
    assert.equal(readFileSync(path, 'utf8'), whole);
  });

  test('anything but a text file is refused', async () => {
    const path = join(tempDir('doc-viewer-text-'), '表.xlsx');
    const book = new ExcelJS.Workbook();
    book.addWorksheet('表').getCell('A1').value = 1;
    await book.xlsx.writeFile(path);
    await assert.rejects(saveTextInPlace({ path, text: '覆盖' }), /不是文本/);
  });
});

describe('formulas worked out without LibreOffice', () => {
  const book = (cells) => (sheet, address) => cells[`${sheet}!${address}`];
  const f = (formula) => ({ formula });
  const v = (value) => ({ value });

  test('arithmetic, references, ranges and the common functions', () => {
    const at = liteEvaluator(book({
      '表!A1': v(12), '表!B1': v(810), '表!C1': f('=A1*B1'),
      '表!A2': v(3), '表!B2': v('备注'), '表!C2': f('A2*10+2^3-(4/2)'),
      '表!C3': f('=SUM(C1:C2)'), '表!C4': f('=ROUND(AVERAGE(C1:C2)/7,1)'),
      '表!C5': f('=MAX(C1:C2)-MIN(A1:A2)+COUNT(A1:B2)'), '表!C6': f('=ABS(-A2)'),
      '汇总 表!A1': f("='表'!C3*2"), '表!C7': f('=SUM(B1:B2)'),
    }));
    assert.equal(at('表', 'C1'), 9720);
    assert.equal(at('表', 'C2'), 36);
    assert.equal(at('表', 'C3'), 9756);
    assert.equal(at('表', 'C4'), 696.9);
    assert.equal(at('表', 'C5'), 9720 - 3 + 3);
    assert.equal(at('表', 'C6'), 3);
    assert.equal(at('汇总 表', 'A1'), 19512);
    assert.equal(at('表', 'C7'), 810, 'text in a range is skipped, as in Excel');
  });

  test('operators bind as they do in Excel', () => {
    const at = liteEvaluator(book({
      '表!A1': v(1234), '表!B1': f('=-2^2'), '表!B2': f('=2^-1'), '表!B3': f('=-A1+10'), '表!B4': f('=ROUND(A1,-2)'),
      '表!B5': f('=ROUND(-2.5,0)'), '表!B6': f('=10%*A1'), '表!B7': f('=2*3^2'),
    }));
    assert.deepEqual(['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7'].map((address) => at('表', address)), [4, 0.5, -1224, 1200, -3, 123.4, 18]);
  });

  test('anything it does not know stays unknown instead of being guessed', () => {
    const at = liteEvaluator(book({
      '表!A1': v(1), '表!A2': v('文字'),
      '表!B1': f('=XLOOKUP(A1,A1:A2,A1:A2)'), '表!B2': f('=A2*2'), '表!B3': f('=A1/0'),
      '表!B4': f('=B5+1'), '表!B5': f('=B4+1'), '表!B6': f('=B1+1'), '表!B7': f('=IF(A1>0,1,2)'),
    }));
    for (const address of ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7']) assert.equal(at('表', address), undefined, address);
  });

  test('an edit LibreOffice cannot calculate still leaves results in the file', async () => {
    const path = join(tempDir('doc-viewer-lite-'), '没有 LibreOffice.xlsx');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('表');
    sheet.getCell('A1').value = 2;
    sheet.getCell('B1').value = { formula: 'A1*10', result: 20 };
    sheet.getCell('B2').value = { formula: 'SUM(B1:B1)+1', result: 21 };
    await workbook.xlsx.writeFile(path);
    const previous = process.env.DSH_LIBREOFFICE_CLI;
    process.env.DSH_LIBREOFFICE_CLI = 'none';
    try {
      const { changes } = await editFile(path, { edits: [{ address: 'A1', value: 5 }] }, dirname(path));
      assert.deepEqual(changes.dependents.map(({ address, after }) => [address, after]), [['B1', '50'], ['B2', '51']]);
      assert.equal(changes.uncalculated, 0);
      const saved = new ExcelJS.Workbook();
      await saved.xlsx.readFile(path);
      assert.equal(saved.worksheets[0].getCell('B2').value.result, 51);
    } finally {
      if (previous === undefined) delete process.env.DSH_LIBREOFFICE_CLI;
      else process.env.DSH_LIBREOFFICE_CLI = previous;
    }
  });
});

describe('package', () => {
  test('every dependency comes from the registry, or Harness refuses to install the plugin', () => {
    const { dependencies } = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
    for (const [name, version] of Object.entries(dependencies)) {
      assert.match(version, /^[\^~]?\d+\.\d+\.\d+/, `${name}: ${version}`);
    }
  });
});

describe('agent tool definitions', () => {
  const tools = new Map();
  registerTools({ effect: (fn) => fn(), tools: { register: (tool) => tools.set(tool.name, tool) } });

  test('tools that default to the open document do not require a path', () => {
    for (const name of ['doc_viewer_read', 'doc_viewer_edit', 'doc_viewer_stats']) {
      assert.ok(!(tools.get(name).parameters.required || []).includes('path'), name);
    }
  });
});

describe('the OCR helper', () => {
  test('is made runnable when its archive dropped the permission', () => {
    const path = join(tempDir('doc-viewer-exec-'), 'helper');
    writeFileSync(path, '#!/bin/sh\n');
    chmodSync(path, 0o644);
    ensureExecutable(path);
    accessSync(path, constants.X_OK);
  });
});

describe('Markdown export', () => {
  const textOf = async (markdown) => {
    const bytes = await textToDocx(markdown, { title: 't', markdown: true });
    const { value } = await mammoth.convertToHtml({ buffer: bytes });
    return value;
  };

  test('a table is a table, with its cells formatted', async () => {
    const html = await textOf('| 指标 | 本周 |\n| --- | ---: |\n| **用户** | `13,982` |\n');
    assert.match(html, /<table>/);
    assert.match(html, /<strong>用户<\/strong>/);
    assert.doesNotMatch(html, /\|/);
    assert.doesNotMatch(html, /---/);
  });

  test('links, task lists, nesting, numbers and rules leave no syntax behind', async () => {
    const html = await textOf('[看板](https://example.com/a)\n\n- [x] 完成\n- [ ] 没做\n  - 子项\n\n3. 三\n4. 四\n\n---\n\n![架构图](a.png)\n\na &amp; b <b>粗</b>\n');
    assert.match(html, /href="https:\/\/example\.com\/a"/);
    assert.match(html, /☑\s*完成/);
    assert.match(html, /☐\s*没做/);
    assert.match(html, /子项/);
    assert.match(html, /3\.\s*三/);
    assert.match(html, /\[图片：架构图\]/);
    assert.doesNotMatch(html, /\]\(|\[x\]|!\[|^---/);
  });

  test('a code block keeps its lines and its characters', async () => {
    const html = await textOf('```bash\necho "a<b"\n  **不是粗体**\n```\n');
    assert.match(html, /\*\*不是粗体\*\*/);
    assert.doesNotMatch(html, /```/);
  });
});

describe('recognizing files', () => {
  test('goes by the bytes, whatever the name', () => {
    assert.equal(documentKind(join(dir, '销售.xlsx')), 'spreadsheet');
    assert.equal(documentKind(join(dir, '周报.docx')), 'word');
    assert.equal(documentKind(join(dir, '报告.pdf')), 'pdf');
    assert.equal(documentKind(join(dir, '工资.csv')), 'spreadsheet');
    assert.equal(documentKind(join(dir, '笔记.txt')), 'text');
  });

  test('refuses binary files instead of showing garbage', async () => {
    await assert.rejects(openDocument(join(dir, '图片.png')), { code: 'unsupported' });
  });

  test('reports a damaged PDF as damaged', async () => {
    await assert.rejects(openDocument(join(dir, '坏文件.pdf')), { code: 'corrupt-pdf' });
  });
});

describe('opening a workbook', () => {
  test('gives every value the display text the grid paints', async () => {
    const doc = await openDocument(join(dir, '销售.xlsx'), noRecalc);
    const cell = (r, c) => doc.sheets[0].celldata.find((entry) => entry.r === r && entry.c === c)?.v;
    assert.equal(cell(1, 2).v, 100);
    assert.equal(cell(1, 2).m, '100');
    assert.equal(cell(1, 4).m, '2026-01-15', 'dates keep their calendar day in every timezone');
    assert.equal(doc.sheets[0].celldata.filter((entry) => entry.v.v !== undefined && entry.v.m === undefined).length, 0);
  });

  test('keeps formulas in the form the grid evaluates, and chains them for recalculation', async () => {
    const doc = await openDocument(join(dir, '销售.xlsx'), noRecalc);
    const summary = doc.sheets[1];
    const formula = summary.celldata.find((entry) => entry.r === 1 && entry.c === 1).v.f;
    assert.equal(formula, "=SUM('数据'!C2:C25)");
    assert.deepEqual(summary.calcChain.map(({ r, c, id }) => ({ r, c, id })), [
      { r: 1, c: 1, id: summary.id },
      { r: 2, c: 1, id: summary.id },
    ]);
  });

  test('reads freeze panes without shifting them by a row', async () => {
    const doc = await openDocument(join(dir, '销售.xlsx'), noRecalc);
    assert.deepEqual(frozenSplit(doc.sheets[0].frozen), { rows: 1, columns: 0 });
    assert.equal(doc.sheets[1].frozen, undefined, 'a sheet without freeze panes gets none');
  });

  test('reports the size of each sheet', async () => {
    const doc = await openDocument(join(dir, '销售.xlsx'), noRecalc);
    assert.deepEqual(doc.sheetSizes[0], { name: '数据', rows: 25, columns: 6 });
  });
});

describe('plain text tables', () => {
  test('real numbers calculate; identifiers stay text', async () => {
    const doc = await openDocument(join(dir, '工资.csv'), noRecalc);
    const cell = (r, c) => doc.sheets[0].celldata.find((entry) => entry.r === r && entry.c === c)?.v;
    assert.equal(cell(1, 2).v, 12000);
    assert.equal(cell(2, 2).v, 9800.5);
    assert.equal(cell(1, 1).v, '00123');
    assert.equal(cell(3, 1).v, '13812345678901234');
  });

  test('decodes legacy Chinese encodings', async () => {
    const doc = await openDocument(join(dir, 'gbk.csv'), noRecalc);
    assert.equal(doc.sheets[0].celldata.find((entry) => entry.r === 1 && entry.c === 0).v.v, '张三');
    assert.ok(doc.warnings.some((warning) => warning.code === 'encoding'));
  });

  test('stays quiet about the ordinary case', async () => {
    const doc = await openDocument(join(dir, '工资.tsv'), noRecalc);
    assert.deepEqual(doc.warnings, []);
  });
});

describe('saving', () => {
  test('a copy keeps every value, formula, merge and freeze', async () => {
    const doc = await openDocument(join(dir, '销售.xlsx'), noRecalc);
    const { output } = await saveSpreadsheetCopy(null, { sheets: doc.sheets, sourcePath: join(dir, '销售.xlsx'), outputDirectory: out, format: 'xlsx' });
    const saved = await readBack(output);
    assert.deepEqual(saved.worksheets.map((sheet) => sheet.name), ['数据', '汇总']);
    assert.equal(saved.worksheets[0].getCell('C2').value, 100);
    assert.equal(saved.worksheets[0].getCell('F2').value, '001');
    assert.ok(String(saved.worksheets[1].getCell('B2').formula).includes('C2:C25'));
    assert.ok(saved.worksheets[1].getCell('B5').isMerged);
    assert.equal(saved.worksheets[0].views[0].ySplit, 1);
  });

  test('accepts the matrix form the grid holds after loading', async () => {
    const data = [[{ v: '标题', m: '标题' }, null], [null, { v: 42, m: '42' }]];
    const { output } = await saveSpreadsheetCopy(null, { sheets: [{ name: '表', data }], outputDirectory: out, format: 'xlsx', title: '矩阵' });
    const saved = await readBack(output);
    assert.equal(saved.worksheets[0].getCell('A1').value, '标题');
    assert.equal(saved.worksheets[0].getCell('B2').value, 42);
  });

  test('typed numbers are saved as numbers, and formulas in their file form', async () => {
    const celldata = [
      { r: 0, c: 0, v: { v: '1001100', m: '1001100', ct: { fa: 'General', t: 'n' } } },
      { r: 0, c: 1, v: { v: '00123', m: '00123', ct: { fa: '@', t: 's' } } },
      { r: 0, c: 2, v: { f: '=IF(A1>1,TRUE(),"FALSE()")', v: true, ct: { fa: 'General', t: 'b' } } },
    ];
    const { output } = await saveSpreadsheetCopy(null, { sheets: [{ name: '表', celldata }], outputDirectory: out, format: 'xlsx', title: '类型' });
    const sheet = (await readBack(output)).worksheets[0];
    assert.equal(sheet.getCell('A1').value, 1001100);
    assert.equal(sheet.getCell('B1').value, '00123');
    assert.equal(sheet.getCell('C1').formula, 'IF(A1>1,TRUE,"FALSE()")');
  });

  test('refuses to write a file with no sheets', async () => {
    await assert.rejects(saveSpreadsheetCopy(null, { sheets: [], outputDirectory: out, format: 'xlsx' }), { code: 'no-sheet' });
    await assert.rejects(saveSpreadsheetCopy(null, { sheets: undefined, outputDirectory: out, format: 'xlsx' }), { code: 'no-sheet' });
  });

  test('never overwrites: a second save gets a new name', async () => {
    const doc = await openDocument(join(dir, '工资.tsv'), noRecalc);
    const first = await saveSpreadsheetCopy(null, { sheets: doc.sheets, outputDirectory: out, format: 'xlsx', title: '同名' });
    const second = await saveSpreadsheetCopy(null, { sheets: doc.sheets, outputDirectory: out, format: 'xlsx', title: '同名' });
    assert.notEqual(first.output, second.output);
  });

  test('duplicate and illegal sheet names become valid ones', () => {
    assert.deepEqual(workbookSheetNames([{ name: '统计' }, { name: '统计' }, { name: 'a/b' }]), ['统计', '统计 (2)', 'a_b']);
  });

  test('CSV says which sheet it kept', async () => {
    const doc = await openDocument(join(dir, '销售.xlsx'), noRecalc);
    const { output, warnings } = await saveSpreadsheetCopy(null, { sheets: doc.sheets, outputDirectory: out, format: 'csv', options: { sheetName: '汇总' } });
    assert.match(readFileSync(output, 'utf8'), /总销量/);
    assert.equal(warnings[0].code, 'csv-one-sheet');
  });

  test('legacy XLS round-trips, including Chinese text', async () => {
    const doc = await openDocument(join(dir, '工资.tsv'), noRecalc);
    const { output } = await saveSpreadsheetCopy(null, { sheets: doc.sheets, outputDirectory: out, format: 'xls', title: '旧版' });
    const back = await openDocument(output, noRecalc);
    assert.equal(back.sheets[0].celldata.find((entry) => entry.r === 1 && entry.c === 0).v.v, '张三');
    assert.equal(back.sheets[0].celldata.find((entry) => entry.r === 1 && entry.c === 1).v.v, 100);
  });

  test('an agent edit leaves no stale formula results behind', async () => {
    const { output } = await editFile(join(dir, '销售.xlsx'), { edits: [{ sheet: '数据', address: 'C2', value: 1100 }], inPlace: false }, out);
    const saved = await readBack(output);
    // The new total where LibreOffice could calculate it, else none: never the old 4332.
    assert.ok([5332, undefined].includes(saved.worksheets[1].getCell('B2').result), String(saved.worksheets[1].getCell('B2').result));
    assert.match(execFileSync('unzip', ['-p', output, 'xl/workbook.xml'], { encoding: 'utf8' }), /fullCalcOnLoad="1"/);
  });

  test('a source file is untouched', async () => {
    const source = join(dir, '销售.xlsx');
    const before_ = statSync(source).mtimeMs;
    await editFile(source, { edits: [{ address: 'A1', value: '改' }], inPlace: false }, out);
    assert.equal(statSync(source).mtimeMs, before_);
  });
});

describe('new documents', () => {
  test('a workbook starts with its rows, a bold frozen header, and calculated formulas', async () => {
    const path = join(tempDir('doc-viewer-new-'), '报销.xlsx');
    await createDocument(path, { rows: [['项目', '金额'], ['打车', 30], ['午餐', 45], ['合计', '=SUM(B2:B3)']] });
    const book = await readBack(path);
    const sheet = book.worksheets[0];
    assert.equal(sheet.name, '工作表1');
    assert.equal(sheet.getCell('A1').font?.bold, true);
    assert.equal(sheet.views[0]?.ySplit, 1);
    assert.equal(sheet.getCell('B4').formula, 'SUM(B2:B3)');
    const read = await readDocument(path, { range: 'B4:B4' });
    assert.ok(['75', ''].includes(read.rows[0][0].display), read.rows[0][0].display);
  });

  test('an empty workbook opens', async () => {
    const path = join(tempDir('doc-viewer-new-'), '空白.xlsx');
    await createDocument(path);
    assert.equal((await openDocument(path, noRecalc)).sheets.length, 1);
  });

  test('nothing is replaced, and a broken formula makes nothing', async () => {
    const dir = tempDir('doc-viewer-new-');
    const existing = join(dir, '已有.md');
    writeFileSync(existing, '原内容');
    await assert.rejects(createDocument(existing, { content: '新内容' }), { code: 'exists' });
    assert.equal(readFileSync(existing, 'utf8'), '原内容');
    const broken = join(dir, '坏.xlsx');
    await assert.rejects(createDocument(broken, { rows: [['=SUM(A1:A3']] }), { code: 'bad-formula' });
    assert.equal(existsSync(broken), false);
    await assert.rejects(createDocument(join(dir, '文稿.docx')), { code: 'bad-format' });
  });

  test('Markdown and text get their content', async () => {
    const path = join(tempDir('doc-viewer-new-'), '笔记.md');
    await createDocument(path, { content: '# 标题\n\n正文' });
    assert.equal(readFileSync(path, 'utf8'), '# 标题\n\n正文');
  });
});

describe('edits that must not lose or corrupt data', () => {
  const copyOf = (name) => {
    const target = join(tempDir('doc-viewer-safe-'), name);
    copyFileSync(join(dir, name), target);
    return target;
  };

  test('parallel edits to one file all land', async () => {
    const path = copyOf('销售.xlsx');
    await Promise.all(['D2', 'D3', 'D4', 'D5'].map((address) => editFile(path, { edits: [{ sheet: '数据', address, value: address }] }, out)));
    const sheet = (await readBack(path)).worksheets[0];
    assert.deepEqual(['D2', 'D3', 'D4', 'D5'].map((address) => sheet.getCell(address).value), ['D2', 'D3', 'D4', 'D5']);
  });

  for (const [formula, problem] of [['=SUM(C2:C5', /少了右括号/], ['=SUM(C2:C5))', /多了一个右括号/], ['=C2+', /没有写完/], ['=D2+1', /循环引用/], ['=SUM(D1:D9)', /循环引用/], ['="a', /引号没有闭合/]]) {
    test(`a malformed formula ${formula} changes nothing`, async () => {
      const path = copyOf('销售.xlsx');
      const before = readFileSync(path);
      await assert.rejects(editFile(path, { edits: [{ sheet: '数据', address: 'D2', value: 1 }, { sheet: '数据', address: 'D2', formula }] }, out), problem);
      assert.deepEqual(readFileSync(path), before);
      assert.equal(listVersions(path).length, 0);
    });
  }

  test('references elsewhere are not mistaken for a loop', async () => {
    const path = copyOf('销售.xlsx');
    const result = await editFile(path, { edits: [{ sheet: '数据', address: 'D2', formula: '=SUM(C2:C5)+汇总!D2+"D2"' }] }, out);
    assert.equal(result.inPlace, true);
  });

  test('writing inside a merged cell changes nothing and names the cell to write', async () => {
    const path = join(tempDir('doc-viewer-merge-'), '合并.xlsx');
    const book = new ExcelJS.Workbook();
    const sheet = book.addWorksheet('表');
    sheet.getCell('A1').value = '标题';
    sheet.mergeCells('A1:C1');
    await book.xlsx.writeFile(path);
    const before = readFileSync(path);
    await assert.rejects(editFile(path, { edits: [{ address: 'B1', value: 'x' }] }, out), /A1:C1.*写 A1/);
    assert.deepEqual(readFileSync(path), before);
    await editFile(path, { edits: [{ address: 'A1', value: '新标题' }] }, out);
    assert.equal((await readBack(path)).worksheets[0].getCell('A1').value, '新标题');
  });

  test('the page can have its formulas worked out on the Host while editing', async () => {
    const path = copyOf('销售.xlsx');
    const { sheets } = await openDocument(path, noRecalc);
    const edited = structuredClone(sheets);
    edited[0].celldata.find((entry) => entry.r === 1 && entry.c === 2).v.v = 1100;
    let results;
    try {
      results = await calculateSheets({ path, sheets: edited });
    } catch (error) {
      // Without LibreOffice there is nothing to calculate with; anything else is a bug.
      if (error?.code === 'office-engine-unavailable') return;
      throw error;
    }
    const total = results.find((item) => item.id === sheets[1].id && item.r === 1 && item.c === 1);
    assert.equal(total?.v, 5332);
  });

  test('an edit in place reports every changed cell, dependent results included', async () => {
    const path = copyOf('销售.xlsx');
    const result = await editFile(path, { edits: [{ sheet: '数据', address: 'C2', value: 1100 }] }, out);
    const cell = (sheet, address) => result.cells.find((change) => change.sheet === sheet && change.address === address);
    assert.equal(cell('数据', 'C2').before.v, 100);
    assert.equal(cell('数据', 'C2').after.v, 1100);
    assert.equal(result.mtime, statSync(path).mtimeMs);
    // Only where LibreOffice could calculate is the total known to have moved.
    const total = cell('汇总', 'B2');
    if (total) assert.equal(total.after.v, 5332);
  });

  test('edits for a page with unsaved edits leave the file alone and say what to put where', async () => {
    const path = copyOf('销售.xlsx');
    const before = readFileSync(path);
    const result = await editFile(path, { live: true, edits: [{ sheet: '数据', address: 'C2', value: 1100 }, { sheet: '数据', address: 'H2', formula: '=C2*2' }] }, out);
    assert.equal(result.live, true);
    assert.deepEqual(readFileSync(path), before);
    assert.equal(listVersions(path).length, 0);
    assert.deepEqual(result.cells.map((change) => change.address).sort(), ['C2', 'H2']);
    assert.equal(result.cells.find((change) => change.address === 'H2').after.f, '=C2*2');
  });

  test('edits a page could not show go to a copy instead of being lost', async () => {
    const path = copyOf('销售.xlsx');
    const result = await editFile(path, { live: true, edits: [{ sheet: '数据', address: 'C5000', value: 1 }] }, out);
    assert.equal(result.live, undefined);
    assert.equal(result.inPlace, false);
    assert.notEqual(result.output, path);
    assert.match(result.keptAside, /超出了表格当前的范围/);
  });

  test('newer functions carry the prefix a file needs, and the grid sees them without it', () => {
    assert.equal(fileFormula('=XLOOKUP(1,A1:A3,B1:B3)'), '_xlfn.XLOOKUP(1,A1:A3,B1:B3)');
    assert.equal(fileFormula('=filter(A1:A3,A1:A3>1)'), '_xlfn._xlws.FILTER(A1:A3,A1:A3>1)');
    assert.equal(fileFormula('="XLOOKUP(" & A1'), '"XLOOKUP(" & A1', 'text is left alone');
    assert.equal(fileFormula(modelFormula('_xlfn.XLOOKUP(1,A1:A3,B1:B3)', [])), '_xlfn.XLOOKUP(1,A1:A3,B1:B3)', 'a round trip keeps one prefix');
    assert.equal(modelFormula('_xlfn._xlws.SORT(A1:A3)', []), '=SORT(A1:A3)');
  });
});

describe('agent edits', () => {
  test('an unknown sheet is an error, not a write to the first sheet', async () => {
    await assert.rejects(editFile(join(dir, '销售.xlsx'), { edits: [{ sheet: '不存在', address: 'A1', value: 1 }] }, out), { code: 'bad-sheet' });
  });

  test('a bad address is an error', async () => {
    await assert.rejects(editFile(join(dir, '销售.xlsx'), { edits: [{ address: 'not-a-cell', value: 1 }] }, out), { code: 'bad-address' });
  });

  test('doing nothing produces no file', async () => {
    await assert.rejects(editFile(join(dir, '销售.xlsx'), {}, out), { code: 'nothing-to-do' });
  });

  test('editing does not leak into the cached workbook', async () => {
    const source = join(dir, '销售.xlsx');
    await editFile(source, { edits: [{ address: 'C2', value: 999 }], inPlace: false }, out);
    const again = await readDocument(source, { range: 'C2:C2' });
    assert.equal(again.rows[0][0].value, 100);
  });

  test('converting text numbers leaves identifiers alone', async () => {
    const { output, converted } = await editFile(join(dir, '工资.csv'), { edits: [{ address: 'A1', value: '名字' }], convertTextNumbers: true, inPlace: false }, out);
    assert.equal(converted, 0, 'numbers were already parsed on load');
    const saved = await readBack(output);
    assert.equal(saved.worksheets[0].getCell('B2').value, '00123');
    assert.equal(saved.worksheets[0].getCell('A1').value, '名字');
  });

  test('parseNumericText', () => {
    assert.equal(parseNumericText('1,234.5'), 1234.5);
    assert.equal(parseNumericText('-3'), -3);
    assert.equal(parseNumericText('0.5'), 0.5);
    assert.equal(parseNumericText('007'), undefined);
    assert.equal(parseNumericText('13812345678901234'), undefined);
    assert.equal(parseNumericText('abc'), undefined);
    assert.equal(parseNumericText(5), undefined);
  });

  test('modelFormula quotes Chinese sheet names and is idempotent about =', () => {
    assert.equal(modelFormula('=SUM(数据!A1:A2)', ['数据']), "=SUM('数据'!A1:A2)");
    assert.equal(modelFormula('SUM(数据!A1:A2)', ['数据']), "=SUM('数据'!A1:A2)");
  });
});

describe('saving where the file came from', () => {
  const copyOf = (name) => {
    const target = join(tempDir('doc-viewer-place-'), name);
    copyFileSync(join(dir, name), target);
    return target;
  };

  test('an agent edit changes the file itself, and the old content can be restored', async () => {
    const path = copyOf('销售.xlsx');
    const result = await editFile(path, { edits: [{ sheet: '数据', address: 'C2', value: 555 }] }, out);
    assert.equal(result.inPlace, true);
    assert.equal(result.output, path);
    assert.equal((await readBack(path)).worksheets[0].getCell('C2').value, 555);
    assert.equal(listVersions(path).length, 1);
    restoreVersion(path);
    assert.equal((await readBack(path)).worksheets[0].getCell('C2').value, 100);
    assert.equal(listVersions(path).length, 2, 'the undone state is kept too');
  });

  test('the page saves over the original CSV, with one sheet', async () => {
    const path = copyOf('工资.csv');
    const sheets = structuredClone((await openDocument(path, noRecalc)).sheets);
    Object.assign(sheets[0].celldata.find((entry) => entry.r === 1 && entry.c === 0).v, { v: '张三丰', m: '张三丰' });
    const result = await saveInPlace({ path, sheets });
    assert.equal(result.inPlace, true);
    assert.match(readFileSync(path, 'utf8'), /张三丰,00123,12000/);
  });

  test('never overwrites what it cannot write back', async () => {
    assert.equal(await inPlaceBlocker(copyOf('工资.tsv'), {}), undefined, 'a plain TSV is fine');
    const twoSheets = await saveInPlace({ path: copyOf('工资.csv'), sheets: [{ name: 'a', celldata: [] }, { name: 'b', celldata: [] }] });
    assert.equal(twoSheets.needsLocation, true);
    assert.match(twoSheets.reason, /只能保存一个工作表/);
  });

  test('a file changed by someone else since it was opened is not overwritten', async () => {
    const path = copyOf('销售.xlsx');
    const doc = await openDocument(path, noRecalc);
    const opened = statSync(path).mtimeMs;
    utimesSync(path, new Date(), new Date(Date.now() + 5000));
    const result = await saveInPlace({ path, sheets: doc.sheets, openedMtime: opened });
    assert.equal(result.needsLocation, true);
    assert.match(result.reason, /被其他程序修改过/);
  });

  test('a workbook with conditional formats is edited in place and keeps them', async () => {
    const path = join(tempDir('doc-viewer-place-'), '条件格式.xlsx');
    const book = new ExcelJS.Workbook();
    const sheet = book.addWorksheet('表');
    sheet.addRow([1]);
    sheet.addConditionalFormatting({ ref: 'A1:A5', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [0], style: { font: { bold: true } } }] });
    await book.xlsx.writeFile(path);
    const result = await editFile(path, { edits: [{ address: 'A2', value: 2 }] }, out);
    assert.equal(result.inPlace, true);
    const saved = await readBack(path);
    assert.equal(saved.worksheets[0].getCell('A2').value, 2);
    assert.equal(saved.worksheets[0].conditionalFormattings.length, 1);
  });

  test('a long run of edits can still get back to the file it started from', async () => {
    const path = copyOf('工资.csv');
    const original = readFileSync(path);
    const sheets = structuredClone((await openDocument(path, noRecalc)).sheets);
    const cell = sheets[0].celldata.find((entry) => entry.r === 1 && entry.c === 0).v;
    for (let step = 0; step < 25; step += 1) {
      Object.assign(cell, { v: `第${step}次`, m: `第${step}次` });
      await saveInPlace({ path, sheets });
    }
    const versions = listVersions(path);
    assert.equal(versions.length, 20);
    restoreVersion(path, versions.at(-1).id);
    assert.deepEqual(readFileSync(path), original);
  });

  test('undoing twice goes back twice, and a chosen version leaves the steps alone', async () => {
    const path = copyOf('工资.csv');
    const states = [readFileSync(path, 'utf8')];
    const sheets = structuredClone((await openDocument(path, noRecalc)).sheets);
    const cell = sheets[0].celldata.find((entry) => entry.r === 1 && entry.c === 0).v;
    for (const name of ['甲', '乙', '丙']) {
      Object.assign(cell, { v: name, m: name });
      await saveInPlace({ path, sheets });
      states.push(readFileSync(path, 'utf8'));
    }
    assert.equal(restoreVersion(path).remaining, 2);
    assert.equal(readFileSync(path, 'utf8'), states[2]);
    assert.equal(restoreVersion(path).remaining, 1);
    assert.equal(readFileSync(path, 'utf8'), states[1]);
    Object.assign(cell, { v: '丁', m: '丁' });
    await saveInPlace({ path, sheets });
    restoreVersion(path);
    assert.equal(readFileSync(path, 'utf8'), states[1], 'a new change is undone first');
    restoreVersion(path);
    assert.equal(readFileSync(path, 'utf8'), states[0], 'then the step before the earlier undo');
    assert.throws(() => restoreVersion(path), { code: 'no-version' });
    assert.ok(!listVersions(path).some((version) => /前前/.test(`${version.reason}前`)));
  });

  test('saving through a symbolic link changes the target and keeps the link', async () => {
    const real = copyOf('销售.xlsx');
    const link = join(tempDir('doc-viewer-link-'), '链接.xlsx');
    symlinkSync(real, link);
    const result = await editFile(link, { edits: [{ sheet: '数据', address: 'C2', value: 777 }] }, out);
    assert.equal(result.inPlace, true);
    assert.equal(lstatSync(link).isSymbolicLink(), true);
    assert.equal((await readBack(real)).worksheets[0].getCell('C2').value, 777);
  });

  test('saving keeps the permissions of the file', async () => {
    for (const mode of [0o600, 0o640, 0o755]) {
      const path = copyOf('销售.xlsx');
      chmodSync(path, mode);
      await editFile(path, { edits: [{ sheet: '数据', address: 'C2', value: 1 }] }, out);
      assert.equal(statSync(path).mode & 0o777, mode);
    }
  });

  test('unsaved edits in the page keep the agent off the file', async () => {
    const path = copyOf('销售.xlsx');
    const result = await editFile(path, { edits: [{ address: 'A1', value: 'x' }], blockedReason: '用户有未保存的修改' }, out);
    assert.equal(result.inPlace, false);
    assert.notEqual(result.output, path);
  });
});

describe('statistics', () => {
  test('a value column works without grouping', async () => {
    const result = await statsForFile(join(dir, '销售.xlsx'), { valueColumn: '销量' }, out);
    assert.equal(result.summary.count, 24);
    assert.equal(result.summary.sum, 24 * 100 + 7 * (23 * 24) / 2);
    assert.equal(result.summary.min, 100);
  });

  test('groups by header or by column letter', async () => {
    const byName = await statsForFile(join(dir, '销售.xlsx'), { groupBy: '区域', valueColumn: '销量', agg: 'count' }, out);
    const byLetter = await statsForFile(join(dir, '销售.xlsx'), { groupBy: 'A', valueColumn: 'C', agg: 'count' }, out);
    assert.deepEqual(byName.groups, byLetter.groups);
    assert.equal(byName.groups.length, 4);
    assert.equal(byName.groups[0].count, 6);
  });

  test('a column letter is read within the selected range', () => {
    const celldata = [
      { r: 0, c: 3, v: { v: '组' } }, { r: 0, c: 4, v: { v: '值' } },
      { r: 1, c: 3, v: { v: '甲' } }, { r: 1, c: 4, v: { v: 5 } },
    ];
    const result = computeStats([{ name: 'S', celldata }], { range: 'D1:E2', groupBy: 'D', valueColumn: 'E' });
    assert.deepEqual(result.groups, [{ key: '甲', sum: 5, count: 1 }]);
  });

  describe('the first row of a range', () => {
    const celldata = [
      ...['部门', '数量', '单价'].map((v, c) => ({ r: 0, c, v: { v } })),
      ...[['华东', 31, 2], ['华北', 38, 3], ['华南', 45, 4], ['西南', 52, 5]].flatMap((row, r) => row.map((v, c) => ({ r: r + 1, c, v: { v } }))),
      { r: 6, c: 0, v: { v: '小计' } }, { r: 6, c: 1, v: { v: '合计' } },
      { r: 7, c: 0, v: { v: '甲' } }, { r: 7, c: 1, v: { v: 9 } },
    ];
    const stats = (request) => computeStats([{ name: 'S', celldata }], request);

    test('a selection of numbers counts every cell', () => {
      const { summary, hasHeader } = stats({ range: 'B2:B5' });
      assert.deepEqual([summary.sum, summary.count, hasHeader], [166, 4, false]);
    });

    test('a selection of text in the middle of a sheet counts every cell', () => {
      assert.deepEqual([stats({ range: 'A2:A5' }).summary.count, stats({ range: 'A2:A5' }).hasHeader], [4, false]);
    });

    test('a row of data in the middle of a sheet is not a header', () => {
      assert.equal(stats({ range: 'A3:C5' }).summary.numericCount, 6);
    });

    test('a selection from the top of the sheet leaves its header out', () => {
      const { summary, hasHeader } = stats({ range: 'B1:B5' });
      assert.deepEqual([summary.sum, summary.count, hasHeader], [166, 4, true]);
    });

    test('labels over numbers further down are a header', () => {
      assert.equal(stats({ range: 'A7:B8' }).hasHeader, true);
    });

    test('what the caller says wins', () => {
      assert.equal(stats({ range: 'B2:B5', hasHeader: true }).summary.count, 3);
      assert.equal(stats({ range: 'B1:B5', hasHeader: false }).summary.count, 5);
    });
  });

  test('names what is wrong with a column', async () => {
    await assert.rejects(statsForFile(join(dir, '销售.xlsx'), { groupBy: '区域' }, out), { code: 'bad-column' });
    await assert.rejects(statsForFile(join(dir, '销售.xlsx'), { valueColumn: '没有这列' }, out), /可用的列/);
  });

  test('saving twice into one workbook keeps both', async () => {
    const { saved } = await statsForFile(join(dir, '销售.xlsx'), { save: true }, out);
    const again = await statsForFile(saved, { save: true }, out);
    const workbook = await readBack(again.saved);
    assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ['数据', '汇总', '统计结果', '统计结果 (2)']);
  });

  test('refuses a document that is not a table', async () => {
    await assert.rejects(statsForFile(join(dir, '报告.pdf'), {}, out), { code: 'wrong-kind' });
  });
});

describe('reading for agents', () => {
  test('stays inside its budget and says how to continue', async () => {
    const first = await readDocument(join(dir, '大表.xlsx'), { limit: 2000 });
    assert.equal(first.truncated, true);
    assert.ok(first.rows.length < 40);
    assert.match(first.nextRange, /^A\d+:H50001$/);
    const next = await readDocument(join(dir, '大表.xlsx'), { range: first.nextRange, limit: 2000 });
    assert.equal(next.rows[0][0].value, first.rows.length);
    assert.equal(next.firstRow, first.rows.length);
  });

  test('a 50,000 row workbook is read and summarized without failing', async () => {
    const started = Date.now();
    const read = await readDocument(join(dir, '大表.xlsx'), { range: 'A49990:C50001' });
    assert.equal(read.rows[read.rows.length - 1][0].value, 50_000);
    const stats = await statsForFile(join(dir, '大表.xlsx'), { valueColumn: '编号' }, out);
    assert.equal(stats.summary.count, 50_000);
    assert.equal(stats.summary.max, 50_000);
    assert.ok(Date.now() - started < 30_000);
  });

  test('an unknown sheet lists the real ones', async () => {
    await assert.rejects(readDocument(join(dir, '销售.xlsx'), { sheet: '不存在' }), /数据、汇总/);
  });

  test('a bad range is an error', async () => {
    await assert.rejects(readDocument(join(dir, '销售.xlsx'), { range: 'hello' }), { code: 'bad-range' });
  });

  test('PDF page ranges work, alone and in lists', async () => {
    const alone = await readDocument(join(dir, '报告.pdf'), { pages: '3-5' });
    assert.deepEqual(alone.pages.map((page) => page.page), [3, 4, 5]);
    const list = await readDocument(join(dir, '报告.pdf'), { pages: '1,3-4' });
    assert.deepEqual(list.pages.map((page) => page.page), [1, 3, 4]);
    assert.match(list.pages[0].text, /Report page 1/);
    assert.equal(list.pageCount, 8);
  });

  test('a page budget continues on the next page', async () => {
    const part = await readDocument(join(dir, '报告.pdf'), { pages: '1-8', limit: 200 });
    assert.equal(part.truncated, true);
    assert.ok(part.nextPages);
  });

  test('Word text is read as paragraphs with entities decoded', async () => {
    const { text } = await readDocument(join(dir, '周报.docx'));
    assert.match(text, /本周完成 接口联调 与 A&B 测试。/);
  });

  test('long text continues from an offset', async () => {
    const first = await readDocument(join(dir, '笔记.txt'), { limit: 1000 });
    assert.equal(first.text.length, 1000);
    const next = await readDocument(join(dir, '笔记.txt'), { offset: first.nextOffset, limit: 1000 });
    assert.equal(next.offset, 1000);
  });
});

describe('Word preview', () => {
  test('counts text, not markup', async () => {
    const doc = await openDocument(join(dir, '周报.docx'));
    assert.match(doc.html, /<h1>项目周报<\/h1>/);
    assert.ok(doc.stats.chars < 40);
  });
});

describe('conversion', () => {
  test('one bad file does not stop the batch', async () => {
    const { results } = await convertMany(null, [join(dir, '工资.csv'), join(dir, '图片.png')], 'xlsx', { outputDirectory: out });
    assert.equal(results[0].ok, true);
    assert.equal(results[1].ok, false);
    assert.equal(results[1].code, 'unsupported');
  });

  test('PDF to text, Word, Excel and CSV', async () => {
    for (const target of ['txt', 'docx', 'xlsx', 'csv']) {
      const { results } = await convertMany(null, [join(dir, '报告.pdf')], target, { outputDirectory: out });
      assert.equal(results[0].ok, true, `${target}: ${results[0].message}`);
      assert.ok(results[0].size > 0);
    }
  });

  test('Word to text', async () => {
    const { results } = await convertMany(null, [join(dir, '周报.docx')], 'txt', { outputDirectory: out });
    assert.match(readFileSync(results[0].output, 'utf8'), /项目周报/);
  });

  test('without a folder, outputs go beside the source', async () => {
    const { results } = await convertMany(null, [join(dir, '工资.csv')], 'xlsx', { workspaceRoot: dir });
    assert.equal(results[0].output, join(dir, '工资-转换.xlsx'));
    const split = await pdfOperation(null, 'split', { paths: [join(dir, '附录.pdf')], workspaceRoot: dir });
    assert.equal(split.outputs[0].output, join(dir, '附录-第1页.pdf'));
  });

  test('an impossible conversion is explained', async () => {
    const { results } = await convertMany(null, [join(dir, '周报.docx')], 'xlsx', { outputDirectory: out });
    assert.equal(results[0].code, 'unsupported-conversion');
  });
});

describe('PDF tools', () => {
  test('merges in order', async () => {
    const result = await pdfOperation(null, 'merge', { paths: [join(dir, '报告.pdf'), join(dir, '附录.pdf')], outputDirectory: out });
    assert.equal(result.pages, 9);
  });

  test('needs two files to merge', async () => {
    await assert.rejects(pdfOperation(null, 'merge', { paths: [join(dir, '报告.pdf')], outputDirectory: out }), { code: 'need-two-pdfs' });
  });

  test('splits by range, or into single pages', async () => {
    const ranges = await pdfOperation(null, 'split', { paths: [join(dir, '报告.pdf')], ranges: '1-3,4-8', outputDirectory: out });
    assert.equal(ranges.outputs.length, 2);
    const each = await pdfOperation(null, 'split', { paths: [join(dir, '附录.pdf')], outputDirectory: out });
    assert.equal(each.outputs.length, 1);
  });

  test('deletes pages, and insists on being told which', async () => {
    const result = await pdfOperation(null, 'delete', { paths: [join(dir, '报告.pdf')], pages: '2,4-6', outputDirectory: out });
    assert.equal(result.removed, 4);
    assert.equal(result.remaining, 4);
    await assert.rejects(pdfOperation(null, 'delete', { paths: [join(dir, '报告.pdf')], outputDirectory: out }), { code: 'bad-pages' });
  });

  test('page numbers out of range are explained', async () => {
    await assert.rejects(pdfOperation(null, 'delete', { paths: [join(dir, '报告.pdf')], pages: '99', outputDirectory: out }), /超出范围/);
  });
});

describe('sheet helpers', () => {
  test('withCelldata leaves complete sheets alone', () => {
    const sheet = { name: 'a', celldata: [{ r: 0, c: 0, v: { v: 1 } }] };
    assert.equal(withCelldata(sheet), sheet);
  });

  test('frozenSplit understands every grid freeze mode', () => {
    assert.deepEqual(frozenSplit({ type: 'rangeRow', range: { row_focus: 0, column_focus: 0 } }), { rows: 1, columns: 0 });
    assert.deepEqual(frozenSplit({ type: 'rangeBoth', range: { row_focus: 2, column_focus: 1 } }), { rows: 3, columns: 2 });
    assert.deepEqual(frozenSplit({ type: 'row' }), { rows: 1, columns: 0 });
    assert.deepEqual(frozenSplit(undefined), { rows: 0, columns: 0 });
  });
});

describe('open requests', () => {
  test('only the first claimant gets a request', () => {
    enqueueOpen('/a.pdf', { reveal: true });
    const [request] = listOpenRequests();
    assert.equal(request.reveal, true);
    assert.deepEqual(ackOpenRequests([request.id]), [request.id]);
    assert.deepEqual(ackOpenRequests([request.id]), []);
  });
});
