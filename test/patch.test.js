/**
 * Saving into the original: everything the grid does not show must survive.
 */
import { after, before, describe, test } from 'node:test';
import { removeTempDirs, tempDir } from './fixtures.js';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { openDocument, saveInPlace } from '../src/host/documents.js';
import { featuresLostOnSave } from '../src/host/xlsx-patch.js';
import { decodeText } from '../src/host/encoding.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
let dir;

after(removeTempDirs);

before(() => {
  process.env.DSH_HOME = tempDir('doc-viewer-home-');
  dir = tempDir('doc-viewer-patch-');
});

async function richWorkbook(name) {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet('数据', { pageSetup: { orientation: 'landscape', paperSize: 9 } });
  sheet.addRow(['名称', '数量', '链接']);
  sheet.addRow(['甲', 10, { text: '官网', hyperlink: 'https://example.com' }]);
  sheet.addRow(['乙', 20, '']);
  sheet.addRow(['丙', 30, '']);
  const thin = { style: 'thin', color: { argb: 'FF999999' } };
  for (let r = 1; r <= 4; r += 1) for (let c = 1; c <= 3; c += 1) sheet.getCell(r, c).border = { top: thin, left: thin, bottom: thin, right: thin };
  sheet.getCell('A1').font = { bold: true, color: { theme: 1 } };
  sheet.getCell('A2').note = '这是批注';
  sheet.addConditionalFormatting({ ref: 'B2:B4', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [15], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
  sheet.getCell('B2').dataValidation = { type: 'whole', operator: 'between', formulae: [0, 100], allowBlank: true };
  sheet.autoFilter = 'A1:C4';
  sheet.getCell('A3').alignment = { wrapText: true };
  book.definedNames.add('数据!$B$2:$B$4', '数量区');
  const image = book.addImage({ buffer: PNG, extension: 'png' });
  sheet.addImage(image, { tl: { col: 4, row: 1 }, ext: { width: 40, height: 40 } });
  const hidden = book.addWorksheet('隐藏');
  hidden.state = 'hidden';
  hidden.addRow(['秘密']);
  const path = join(dir, name);
  await book.xlsx.writeFile(path);
  return path;
}

async function load(path) {
  const book = new ExcelJS.Workbook();
  await book.xlsx.readFile(path);
  return book;
}

/** The grid's version of a file, as the page would send it. */
async function gridSheets(path) {
  return structuredClone((await openDocument(path, { recalc: false })).sheets);
}

const cell = (sheet, r, c) => sheet.celldata.find((entry) => entry.r === r && entry.c === c);

describe('saving a workbook into the original', () => {
  test('keeps what the grid does not show, and writes what changed', async () => {
    const path = await richWorkbook('丰富.xlsx');
    const sheets = await gridSheets(path);
    cell(sheets[0], 1, 1).v = { ...cell(sheets[0], 1, 1).v, v: 99, m: '99' };
    cell(sheets[0], 2, 0).v.bl = 1;
    const result = await saveInPlace({ path, sheets });
    assert.equal(result.inPlace, true);

    const book = await load(path);
    const sheet = book.getWorksheet('数据');
    assert.equal(sheet.getCell('B2').value, 99, 'the edit is written');
    assert.equal(sheet.getCell('A3').font?.bold, true, 'the new bold is written');
    assert.equal(sheet.getCell('C3').border?.left?.style, 'thin', 'borders stay');
    assert.equal(sheet.getCell('C2').value?.hyperlink, 'https://example.com', 'hyperlinks stay');
    assert.equal(sheet.getCell('A2').note?.texts?.[0]?.text ?? sheet.getCell('A2').note, '这是批注', 'notes stay');
    assert.equal(sheet.conditionalFormattings.length, 1, 'conditional formats stay');
    assert.equal(sheet.getCell('B2').dataValidation?.type, 'whole', 'data validation stays');
    assert.equal(sheet.getImages().length, 1, 'images stay');
    assert.equal(sheet.pageSetup.orientation, 'landscape', 'print settings stay');
    assert.ok(sheet.autoFilter, 'the filter stays');
    assert.equal(sheet.getCell('A3').alignment?.wrapText, true, 'wrapped text stays wrapped');
    assert.equal(sheet.getCell('A1').font?.color?.theme, 1, 'theme colours stay');
    assert.equal(book.getWorksheet('隐藏').state, 'hidden', 'hidden sheets stay hidden');
    assert.ok(book.definedNames.getRanges('数量区').ranges.length > 0, 'named ranges stay');
  });

  test('wrapped text reads as wrapped in the grid', async () => {
    const path = await richWorkbook('换行.xlsx');
    const sheets = await gridSheets(path);
    assert.equal(cell(sheets[0], 2, 0).v.tb, '2');
  });

  test('a row inserted in the grid moves the formatting with it', async () => {
    const path = await richWorkbook('插行.xlsx');
    const sheets = await gridSheets(path);
    const sheet = sheets[0];
    for (const entry of sheet.celldata) if (entry.r >= 1) entry.r += 1;
    sheet.celldata.push({ r: 1, c: 0, v: { v: '新行', m: '新行', ct: { fa: 'General', t: 's' } } });
    await saveInPlace({ path, sheets, structure: [{ op: 'insertRowCol', id: 'sheet-0', type: 'row', index: 1, count: 1, direction: 'lefttop' }] });
    const saved = (await load(path)).getWorksheet('数据');
    assert.equal(saved.getCell('A2').value, '新行');
    assert.equal(saved.getCell('A3').value, '甲');
    assert.equal(saved.getCell('C3').value?.hyperlink, 'https://example.com', 'the link moved down with its row');
    assert.equal(saved.getCell('C5').border?.bottom?.style, 'thin', 'the borders moved down with their rows');
  });

  test('borders drawn in the grid are saved; others are left alone', async () => {
    const path = await richWorkbook('边框.xlsx');
    const sheets = await gridSheets(path);
    const sheet = sheets[0];
    // What the page sends: every cell's borders as the grid draws them, plus a thick line under E1.
    const borders = {};
    for (const entry of sheet.config.borderInfo) borders[`${entry.value.row_index}_${entry.value.col_index}`] = { l: entry.value.l, r: entry.value.r, t: entry.value.t, b: entry.value.b };
    borders['0_4'] = { b: { style: 13, color: '#ff0000' } };
    sheet.borders = borders;
    await saveInPlace({ path, sheets });
    const saved = (await load(path)).getWorksheet('数据');
    assert.equal(saved.getCell('E1').border?.bottom?.style, 'thick');
    assert.equal(saved.getCell('E2').border?.top?.style, 'thick', 'the shared line is drawn on both cells');
    assert.equal(saved.getCell('B3').border?.right?.style, 'thin');
  });

  test('a removed border is removed', async () => {
    const path = await richWorkbook('去边框.xlsx');
    const sheets = await gridSheets(path);
    sheets[0].borders = {};
    await saveInPlace({ path, sheets });
    const saved = (await load(path)).getWorksheet('数据');
    assert.equal(saved.getCell('B3').border?.right, undefined);
  });

  test('a sheet added in the grid is added; a sheet deleted is deleted', async () => {
    const path = await richWorkbook('增删表.xlsx');
    const sheets = await gridSheets(path);
    const result = [sheets[0], { name: '新表', id: 'grid-new', celldata: [{ r: 0, c: 0, v: { v: '你好', ct: { fa: 'General', t: 's' } } }], config: {} }];
    await saveInPlace({ path, sheets: result });
    const book = await load(path);
    assert.deepEqual(book.worksheets.map((sheet) => sheet.name), ['数据', '新表']);
    assert.equal(book.getWorksheet('新表').getCell('A1').value, '你好');
  });
});

describe('what cannot be written back', () => {
  test('a workbook with a chart is never overwritten', async () => {
    const path = await richWorkbook('图表.xlsx');
    const unpacked = tempDir('doc-viewer-zip-');
    execFileSync('unzip', ['-q', path, '-d', unpacked]);
    execFileSync('mkdir', ['-p', join(unpacked, 'xl', 'charts')]);
    writeFileSync(join(unpacked, 'xl', 'charts', 'chart1.xml'), '<c:chartSpace/>');
    const withChart = join(dir, '有图表.xlsx');
    execFileSync('zip', ['-q', '-r', withChart, '.'], { cwd: unpacked });
    assert.deepEqual(featuresLostOnSave(readFileSync(withChart)), ['图表']);
    const before_ = readFileSync(withChart);
    const result = await saveInPlace({ path: withChart, sheets: await gridSheets(withChart) });
    assert.equal(result.needsLocation, true);
    assert.deepEqual(readFileSync(withChart), before_);
  });

  test('pictures alone are not a reason to refuse', async () => {
    const path = await richWorkbook('图片.xlsx');
    assert.deepEqual(featuresLostOnSave(readFileSync(path)), []);
  });
});

describe('saving a CSV into the original', () => {
  test('keeps a GBK file in GBK, without adding a byte order mark', async () => {
    const path = join(dir, '名单.csv');
    // "姓名,部门\r\n张三,研发\r\n" in GBK.
    writeFileSync(path, Buffer.from('d0d5c3fb2cb2bfc3c50d0ad5c5c8fd2cd1d0b7a20d0a', 'hex'));
    const sheets = await gridSheets(path);
    Object.assign(cell(sheets[0], 1, 1).v, { v: '市场', m: '市场' });
    const result = await saveInPlace({ path, sheets });
    assert.deepEqual(result.warnings, []);
    const bytes = readFileSync(path);
    assert.notEqual(bytes[0], 0xef, 'no UTF-8 byte order mark');
    const { text, encoding } = decodeText(bytes);
    assert.equal(encoding, 'gb18030');
    assert.equal(text, '姓名,部门\r\n张三,市场');
  });

  test('a character GBK cannot hold switches the file to UTF-8, and says so', async () => {
    const path = join(dir, '表情.csv');
    writeFileSync(path, Buffer.from('d0d5c3fb0a', 'hex'));
    const sheets = await gridSheets(path);
    sheets[0].celldata.push({ r: 1, c: 0, v: { v: '😀', m: '😀', ct: { fa: 'General', t: 's' } } });
    const result = await saveInPlace({ path, sheets });
    assert.match(result.warnings[0].message, /UTF-8/);
    assert.match(readFileSync(path, 'utf8'), /😀/);
  });
});
