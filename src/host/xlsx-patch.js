/**
 * Saving a workbook without losing what the grid does not know about.
 *
 * The grid holds values, formulas and plain formatting. A real workbook also
 * has hyperlinks, named ranges, hidden sheets, print settings, conditional
 * formats, data validation, theme colours and more. So a save does not rebuild
 * the file from the grid: it opens the original, replays the rows and columns
 * the user inserted or deleted, compares the grid with what the original holds,
 * and writes only the differences. Everything else stays as it was.
 */
import { inflateRawSync } from 'node:zlib';
import { t } from './i18n.js';
import ExcelJS from 'exceljs';
import {
  cellBorders,
  cellKey,
  cellLiteral,
  excelAlignment,
  excelBorderSide,
  excelSheetToFortune,
  fileFormula,
  frozenSplit,
  withCelldata,
  workbookSheetNames,
  writeFortuneSheet,
} from './spreadsheet.js';

/* ------------------------------------------------------------------ */
/* What a file holds that cannot be written back                       */
/* ------------------------------------------------------------------ */

/**
 * The entries of a zip archive, read from its central directory.
 * @param {Buffer} bytes
 * @returns {{ name: string, read: () => Buffer }[]}
 */
function zipEntries(bytes) {
  let end = -1;
  for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 65_557); at -= 1) {
    if (bytes.readUInt32LE(at) === 0x06054b50) {
      end = at;
      break;
    }
  }
  if (end < 0) return [];
  const count = bytes.readUInt16LE(end + 10);
  let at = bytes.readUInt32LE(end + 16);
  const entries = [];
  for (let index = 0; index < count && at + 46 <= bytes.length; index += 1) {
    if (bytes.readUInt32LE(at) !== 0x02014b50) break;
    const method = bytes.readUInt16LE(at + 10);
    const size = bytes.readUInt32LE(at + 20);
    const nameLength = bytes.readUInt16LE(at + 28);
    const extraLength = bytes.readUInt16LE(at + 30);
    const commentLength = bytes.readUInt16LE(at + 32);
    const local = bytes.readUInt32LE(at + 42);
    const name = bytes.toString('utf8', at + 46, at + 46 + nameLength);
    entries.push({
      name,
      read: () => {
        const start = local + 30 + bytes.readUInt16LE(local + 26) + bytes.readUInt16LE(local + 28);
        const data = bytes.subarray(start, start + size);
        return method === 8 ? inflateRawSync(data) : data;
      },
    });
    at += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/**
 * Parts and markup ExcelJS drops when it writes a workbook. A file holding any
 * of them is never overwritten.
 * @param {Buffer} bytes an .xlsx file
 * @returns {string[]} what would be lost, in words for the user
 */
export function featuresLostOnSave(bytes) {
  const entries = zipEntries(bytes);
  const names = entries.map((entry) => entry.name);
  const found = new Set();
  const has = (prefix) => names.some((name) => name.startsWith(prefix));
  if (has('xl/charts/') || has('xl/chartsheets/')) found.add(t('图表'));
  if (has('xl/pivotTables/') || has('xl/pivotCache/')) found.add(t('数据透视表'));
  if (has('xl/vbaProject.bin')) found.add(t('宏'));
  if (has('xl/slicers/') || has('xl/slicerCaches/') || has('xl/timelines/')) found.add(t('切片器'));
  if (has('xl/externalLinks/')) found.add(t('外部链接'));
  if (has('xl/activeX/') || has('xl/ctrlProps/')) found.add(t('表单控件'));
  if (has('xl/embeddings/')) found.add(t('嵌入的对象'));
  if (has('xl/threadedComments/')) found.add(t('批注'));
  for (const entry of entries) {
    if (/^xl\/drawings\/drawing\d+\.xml$/.test(entry.name)) {
      // Pictures survive; shapes and text boxes do not.
      const xml = entry.read().toString('utf8');
      if (/<xdr:sp[\s>]/.test(xml) || /<xdr:cxnSp[\s>]/.test(xml) || /<xdr:grpSp[\s>]/.test(xml)) found.add(t('形状或文本框'));
    } else if (/^xl\/worksheets\/sheet\d+\.xml$/.test(entry.name)) {
      const xml = entry.read().toString('utf8');
      if (xml.includes('sparklineGroup')) found.add(t('迷你图'));
      if (xml.includes('<x14:conditionalFormatting')) found.add(t('条件格式（新版样式）'));
      if (xml.includes('<x14:dataValidation')) found.add(t('引用其他工作表的数据验证'));
    }
  }
  return [...found];
}

/* ------------------------------------------------------------------ */
/* Writing only the differences                                        */
/* ------------------------------------------------------------------ */

const FONT_KEYS = ['bl', 'it', 'un', 'cl', 'fs', 'ff', 'fc'];
const ALIGN_KEYS = ['ht', 'vt', 'tb'];

/** A comparable form of one style property. */
function styleValue(cell, key) {
  const value = cell?.[key];
  if (value === undefined || value === null || value === '') return '';
  if (key === 'bl' || key === 'it' || key === 'un' || key === 'cl') return Number(value) ? '1' : '';
  if (key === 'tb') return String(value) === '2' ? '2' : '';
  if (key === 'fc' || key === 'bg') return String(value).toLowerCase();
  return String(value);
}

/** @param {any} cell */
function formulaOf(cell) {
  return typeof cell?.f === 'string' && cell.f.trim() !== '' ? fileFormula(cell.f) : undefined;
}

/** @param {unknown} a @param {unknown} b */
function sameLiteral(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a));
  return a === b;
}

/** Apply the changed font properties of a cell, keeping the rest of its font. */
function applyFont(target, cell) {
  const font = { ...(target.font || {}) };
  if (styleValue(cell, 'bl')) font.bold = true; else delete font.bold;
  if (styleValue(cell, 'it')) font.italic = true; else delete font.italic;
  if (styleValue(cell, 'un')) font.underline = true; else delete font.underline;
  if (styleValue(cell, 'cl')) font.strike = true; else delete font.strike;
  if (cell?.fs) font.size = Number(cell.fs);
  if (cell?.ff) font.name = String(cell.ff);
  if (/^#[0-9a-f]{6}$/i.test(cell?.fc || '')) font.color = { argb: `FF${cell.fc.slice(1).toUpperCase()}` };
  target.font = font;
}

/**
 * Every border line of a sheet once: the bottom of one cell is the top of the
 * cell below. Comparing lines rather than cells means a border set from either
 * side counts as the same border.
 * @param {Map<string, any>} borders
 * @returns {Map<string, any>} `h:r:c` is the line above (r, c), `v:r:c` the line left of it
 */
function borderLines(borders) {
  const lines = new Map();
  const put = (key, side) => { if (side && Number(side.style)) lines.set(key, side); };
  for (const [key, sides] of borders) {
    const [r, c] = key.split('_').map(Number);
    put(`h:${r}:${c}`, sides.t);
    put(`h:${r + 1}:${c}`, sides.b);
    put(`v:${r}:${c}`, sides.l);
    put(`v:${r}:${c + 1}`, sides.r);
  }
  return lines;
}

/** @param {any} a @param {any} b */
function sameSide(a, b) {
  if (!a || !b) return !a && !b;
  return Number(a.style) === Number(b.style) && String(a.color || '').toLowerCase() === String(b.color || '').toLowerCase();
}

/** Draw or clear one border line on both cells that share it. */
function setLine(worksheet, key, side) {
  const [kind, rText, cText] = key.split(':');
  const r = Number(rText);
  const c = Number(cText);
  const value = excelBorderSide(side);
  const edit = (row, column, edge) => {
    if (row < 0 || column < 0) return;
    const cell = worksheet.getCell(row + 1, column + 1);
    const border = { ...(cell.border || {}) };
    if (value) border[edge] = value; else delete border[edge];
    cell.border = border;
  };
  if (kind === 'h') {
    edit(r, c, 'top');
    edit(r - 1, c, 'bottom');
  } else {
    edit(r, c, 'left');
    edit(r, c - 1, 'right');
  }
}

/**
 * Bring one worksheet in line with the grid's version of it.
 * @param {import('exceljs').Worksheet} worksheet
 * @param {any} base the worksheet as the grid would have loaded it
 * @param {any} sheet the grid's sheet
 */
function applySheet(worksheet, base, sheet) {
  const before = new Map(base.celldata.map((entry) => [cellKey(entry.r, entry.c), entry]));
  const after = new Map(sheet.celldata.filter((entry) => entry?.v).map((entry) => [cellKey(Number(entry.r), Number(entry.c)), entry]));

  // Merges first: writing into a cell a merge covers would go to its master cell.
  const mergesBefore = new Map(Object.values(base.config?.merge || {}).map((m) => [`${m.r}_${m.c}_${m.rs}_${m.cs}`, m]));
  const mergesAfter = new Map(Object.values(sheet.config?.merge || {}).filter((m) => m && (m.rs > 1 || m.cs > 1)).map((m) => [`${m.r}_${m.c}_${m.rs}_${m.cs}`, m]));
  for (const [key, m] of mergesBefore) if (!mergesAfter.has(key)) worksheet.unMergeCells(m.r + 1, m.c + 1, m.r + m.rs, m.c + m.cs);
  for (const [key, m] of mergesAfter) {
    if (mergesBefore.has(key)) continue;
    try {
      worksheet.mergeCells(m.r + 1, m.c + 1, m.r + m.rs, m.c + m.cs);
    } catch { /* overlaps a merge the grid also dropped; the values are still written */ }
  }

  for (const key of new Set([...before.keys(), ...after.keys()])) {
    const entry = after.get(key) || before.get(key);
    const r = Number(entry.r);
    const c = Number(entry.c);
    const was = before.get(key)?.v;
    const now = after.get(key)?.v;
    if (now?.mc && (now.mc.r !== r || now.mc.c !== c)) continue;
    const target = worksheet.getCell(r + 1, c + 1);

    const formulaNow = formulaOf(now);
    const formulaWas = formulaOf(was);
    if (formulaNow || formulaWas) {
      if (formulaNow !== formulaWas) target.value = formulaNow ? { formula: formulaNow } : cellLiteral(now || {}) ?? null;
    } else {
      const literalNow = now ? cellLiteral(now) : undefined;
      const literalWas = was ? cellLiteral(was) : undefined;
      if (!sameLiteral(literalNow, literalWas)) target.value = literalNow ?? null;
    }

    if ((now?.ct?.fa || 'General') !== (was?.ct?.fa || 'General')) target.numFmt = now?.ct?.fa && now.ct.fa !== 'General' ? now.ct.fa : 'General';
    if (FONT_KEYS.some((name) => styleValue(now, name) !== styleValue(was, name))) applyFont(target, now);
    if (styleValue(now, 'bg') !== styleValue(was, 'bg')) {
      target.fill = /^#[0-9a-f]{6}$/i.test(now?.bg || '')
        ? { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${now.bg.slice(1).toUpperCase()}` } }
        : { type: 'pattern', pattern: 'none' };
    }
    if (ALIGN_KEYS.some((name) => styleValue(now, name) !== styleValue(was, name))) {
      target.alignment = { ...(target.alignment || {}), horizontal: undefined, vertical: undefined, wrapText: undefined, ...(excelAlignment(now || {}) || {}) };
    }
  }

  // Borders, line by line. Only when the page sent what the grid shows.
  if (sheet.borders) {
    const linesBefore = borderLines(cellBorders(base));
    const linesAfter = borderLines(cellBorders(sheet));
    for (const key of new Set([...linesBefore.keys(), ...linesAfter.keys()])) {
      if (!sameSide(linesBefore.get(key), linesAfter.get(key))) setLine(worksheet, key, linesAfter.get(key));
    }
  }

  // Column widths, row heights and hidden rows/columns that changed.
  const sizes = (config, key) => config?.[key] || {};
  for (const index of new Set([...Object.keys(sizes(base.config, 'columnlen')), ...Object.keys(sizes(sheet.config, 'columnlen'))])) {
    const was = sizes(base.config, 'columnlen')[index];
    const now = sizes(sheet.config, 'columnlen')[index];
    if (now !== undefined && Math.abs(Number(now) - Number(was ?? -1)) > 1) worksheet.getColumn(Number(index) + 1).width = Math.max(1, Math.round((Number(now) / 7) * 100) / 100);
  }
  for (const index of new Set([...Object.keys(sizes(base.config, 'rowlen')), ...Object.keys(sizes(sheet.config, 'rowlen'))])) {
    const was = sizes(base.config, 'rowlen')[index];
    const now = sizes(sheet.config, 'rowlen')[index];
    if (now !== undefined && Math.abs(Number(now) - Number(was ?? -1)) > 1) worksheet.getRow(Number(index) + 1).height = Math.round(((Number(now) * 72) / 96) * 100) / 100;
  }
  for (const [key, getter] of [['colhidden', (i) => worksheet.getColumn(i)], ['rowhidden', (i) => worksheet.getRow(i)]]) {
    const was = new Set(Object.keys(sizes(base.config, key)));
    const now = new Set(Object.keys(sizes(sheet.config, key)));
    for (const index of now) if (!was.has(index)) getter(Number(index) + 1).hidden = true;
    for (const index of was) if (!now.has(index)) getter(Number(index) + 1).hidden = false;
  }

  const splitBefore = frozenSplit(base.frozen);
  const splitAfter = frozenSplit(sheet.frozen);
  if (splitBefore.rows !== splitAfter.rows || splitBefore.columns !== splitAfter.columns) {
    const view = { ...(worksheet.views?.[0] || {}) };
    worksheet.views = [splitAfter.rows || splitAfter.columns
      ? { ...view, state: 'frozen', xSplit: splitAfter.columns, ySplit: splitAfter.rows, topLeftCell: undefined }
      : { ...view, state: 'normal', xSplit: undefined, ySplit: undefined, topLeftCell: undefined }];
  }
}

/**
 * Replay a row or column insert or delete the grid made.
 * @param {import('exceljs').Worksheet} worksheet
 * @param {{ op: string, type: string, index?: number, count?: number, direction?: string, start?: number, end?: number }} change
 */
function replayStructure(worksheet, change) {
  const rows = change.type === 'row';
  if (change.op === 'insertRowCol') {
    const at = (change.direction === 'rightbottom' ? Number(change.index) + 1 : Number(change.index)) + 1;
    const blanks = Array.from({ length: Math.max(1, Number(change.count) || 1) }, () => []);
    if (rows) worksheet.spliceRows(at, 0, ...blanks);
    else worksheet.spliceColumns(at, 0, ...blanks);
  } else if (change.op === 'deleteRowCol') {
    const count = Number(change.end) - Number(change.start) + 1;
    if (count < 1) return;
    if (rows) worksheet.spliceRows(Number(change.start) + 1, count);
    else worksheet.spliceColumns(Number(change.start) + 1, count);
  }
}

/** @param {Buffer} bytes an .xlsx @returns {boolean} */
export function hasFormulas(bytes) {
  return zipEntries(bytes).some((entry) => /^xl\/worksheets\/[^/]+\.xml$/.test(entry.name) && /<f[\s>/]/.test(entry.read().toString('latin1')));
}

/**
 * Replace the saved result of every formula. A save changes values that other
 * formulas depend on, and anything reading the file without calculating (this
 * plugin, previews, data tools) trusts these results. Where there is no answer the
 * result is left out, since no result is better than a wrong one.
 * @param {Buffer} bytes an .xlsx
 * @param {((sheet: string, address: string) => unknown)|undefined} resultOf
 * @returns {Promise<Buffer>}
 */
export async function withFormulaResults(bytes, resultOf) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  for (const worksheet of workbook.worksheets) {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell.type !== ExcelJS.ValueType.Formula) return;
        const { result: _stale, ...formula } = cell.value;
        const result = resultOf?.(worksheet.name, cell.address);
        cell.value = result === undefined ? formula : { ...formula, result };
      });
    });
  }
  // Excel still recalculates on opening; the results above are for readers that do not.
  workbook.calcProperties = { ...(workbook.calcProperties || {}), fullCalcOnLoad: true };
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

/**
 * The original workbook with the grid's changes written into it.
 * @param {Buffer} original the .xlsx the sheets were loaded from
 * @param {any[]} sheets the grid's sheets; ids `sheet-N` are the file's sheets
 * @param {{ structure?: any[] }} [options] row/column inserts and deletes, in order
 * @returns {Promise<Buffer>}
 */
export async function patchXlsx(original, sheets, options = {}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(original);
  const worksheets = workbook.worksheets.slice();
  const names = worksheets.map((worksheet) => worksheet.name);
  const byId = new Map(worksheets.map((worksheet, order) => [`sheet-${order}`, worksheet]));

  for (const change of options.structure || []) {
    const worksheet = byId.get(change.id);
    if (worksheet) replayStructure(worksheet, change);
  }

  const complete = sheets.map(withCelldata);
  const finalNames = workbookSheetNames(complete);
  // Match by file id; a sheet the page added and saved before has no file id yet, so its name finds it.
  const claimed = new Set(complete.map((sheet) => byId.get(sheet.id)).filter(Boolean));
  const kept = new Set();
  const matched = complete.map((sheet) => {
    let worksheet = byId.get(sheet.id);
    if (!worksheet || kept.has(worksheet)) worksheet = worksheets.find((candidate) => !kept.has(candidate) && !claimed.has(candidate) && candidate.name === sheet.name);
    if (worksheet) kept.add(worksheet);
    return worksheet;
  });
  for (const worksheet of worksheets) if (!kept.has(worksheet)) workbook.removeWorksheet(worksheet.id);

  complete.forEach((sheet, index) => {
    let worksheet = matched[index];
    if (!worksheet) {
      worksheet = writeFortuneSheet(workbook, sheet, finalNames[index]);
    } else {
      const order = worksheets.indexOf(worksheet);
      const base = excelSheetToFortune(worksheet, order, names);
      applySheet(worksheet, base, sheet);
      if (worksheet.name !== finalNames[index]) worksheet.name = finalNames[index];
    }
    worksheet.orderNo = index;
  });

  workbook.calcProperties = { ...(workbook.calcProperties || {}), fullCalcOnLoad: true };
  return Buffer.from(await workbook.xlsx.writeBuffer());
}
