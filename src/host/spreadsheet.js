/**
 * Spreadsheet parsing and export.
 *
 * XLSX goes through ExcelJS (style fidelity), XLS/BIFF goes through SheetJS.
 * CSV/TSV and HTML/plain-text files masquerading as `.xls` are detected from
 * the bytes, never trusted by extension. Formula cells keep their `f` text;
 * calculated values are filled from a LibreOffice recalculation when needed,
 * so the panel can show a correct value immediately.
 */
import ExcelJS from 'exceljs';
import { t } from './i18n.js';
import * as XLSX from '../../vendor/xlsx/xlsx.mjs';
import { execFile } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { decodeText, looksDelimited } from './encoding.js';
import { recalculateWorkbook } from './libreoffice.js';
import { layoutSheetForPrint } from './print.js';
import {
  ensureDir,
  escapeRe,
  friendlyMessage,
  extensionOf,
  fileExists,
  httpError,
  readBytes,
  sniffContainer,
  token,
} from './util.js';

const execFileAsync = promisify(execFile);
const HERE = dirname(fileURLToPath(import.meta.url));

/** Date serial for the 1900 date system, with the Excel leap-year bug. */
export function dateToSerial(date) {
  return (date.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000;
}

/** @param {number} serial */
export function serialToDate(serial) {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000);
}

/** Number formatting. Node's named exports of this CommonJS package do not include it. */
const SSF = XLSX.SSF ?? XLSX.default.SSF;
/**
 * Functions newer than the file format, and the prefix they carry in a file. Without
 * it Excel and LibreOffice both answer #NAME?, however correct the formula reads.
 */
const NEWER_FUNCTIONS = {
  XLOOKUP: '_xlfn.', XMATCH: '_xlfn.', MAXIFS: '_xlfn.', MINIFS: '_xlfn.', TEXTJOIN: '_xlfn.', IFS: '_xlfn.', SWITCH: '_xlfn.', CONCAT: '_xlfn.',
  UNIQUE: '_xlfn.', SORTBY: '_xlfn.', SEQUENCE: '_xlfn.', FILTER: '_xlfn._xlws.', SORT: '_xlfn._xlws.',
};
const NEWER_FUNCTION_NAMES = Object.keys(NEWER_FUNCTIONS).join('|');
const NEWER_FUNCTION_CALL = new RegExp(`(?<![\\w.])(?:${NEWER_FUNCTION_NAMES})(?=\\s*\\()`, 'gi');
const FILE_FUNCTION_PREFIX = new RegExp(`\\b_xlfn\\.(?:_xlws\\.)?(?=(?:${NEWER_FUNCTION_NAMES})\\s*\\()`, 'gi');
const DATE_FORMAT_RE = /(^|[^\\])(y{2,4}|m{1,4}|d{1,4}|h{1,2}|s{1,2}|AM\/PM)/i;

/** @param {string|undefined} format */
export function looksDateFormat(format) {
  if (!format || format === 'General') return false;
  const plain = format.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  return DATE_FORMAT_RE.test(plain) && !/[#0?]/.test(plain.replace(/s\.0+/gi, 's'));
}

/**
 * Make an Excel formula safe for FortuneSheet's parser while keeping it valid
 * for Excel:
 * - quote worksheet names before `!` (FortuneSheet requires quotes for
 *   non-ASCII names such as `数据!A1`);
 * - turn `TRUE`/`FALSE` constants into `TRUE()`/`FALSE()` (the parser treats
 *   bare boolean literals as unknown names in some functions).
 * String literals inside double quotes are never rewritten.
 * @param {string} formula
 * @param {string[]} sheetNames
 */
export function normalizeFormula(formula, sheetNames) {
  if (!formula) return formula;
  const names = [...new Set((sheetNames || []).filter(Boolean))].sort((a, b) => b.length - a.length);
  const escape = (value) => value.replaceAll("'", "''");
  const rewriteCode = (code) => {
    let result = code;
    for (const name of names) {
      const quoted = `'${escape(name)}'`;
      const pattern = new RegExp(`(?<![A-Za-z0-9_'])${escapeRe(name)}!`, 'g');
      result = result.replace(pattern, `${quoted}!`);
    }
    result = result.replace(/\b(FALSE|TRUE)\b(?!\s*\()/gi, (match) => `${match.toUpperCase()}()`);
    return result.replace(FILE_FUNCTION_PREFIX, '');
  };
  let output = '';
  let codeStart = 0;
  let index = 0;
  while (index < formula.length) {
    if (formula[index] === '"') {
      output += rewriteCode(formula.slice(codeStart, index));
      let end = index + 1;
      while (end < formula.length) {
        if (formula[end] === '"') {
          if (formula[end + 1] === '"') { end += 2; continue; }
          break;
        }
        end += 1;
      }
      output += formula.slice(index, Math.min(formula.length, end + 1));
      index = end + 1;
      codeStart = index;
    } else {
      index += 1;
    }
  }
  output += rewriteCode(formula.slice(codeStart));
  return output;
}

/**
 * Convert a formula from file form into the panel's model, which keeps the
 * leading `=` FortuneSheet needs in order to evaluate it.
 * @param {string} formula
 * @param {string[]} sheetNames
 */
export function modelFormula(formula, sheetNames) {
  return `=${normalizeFormula(String(formula).trim().replace(/^=/, ''), sheetNames)}`;
}

/**
 * Parse text that is unambiguously a number. Identifiers such as `00123`, or
 * phone and ID numbers with more digits than a double can hold, stay text.
 * @param {unknown} value
 * @returns {number|undefined}
 */
export function parseNumericText(value) {
  if (typeof value !== 'string') return undefined;
  const text = value.trim().replaceAll(',', '').replaceAll('，', '');
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) return undefined;
  if (/^[+-]?0\d/.test(text)) return undefined;
  if (text.replace(/[^\d]/g, '').length > 15) return undefined;
  const number = Number(text);
  return Number.isFinite(number) ? number : undefined;
}

/** @param {string|undefined} argb */
function colorHex(argb) {
  if (!argb) return undefined;
  const value = String(argb).replace(/^#/, '');
  return value.length === 8 ? `#${value.slice(2)}` : `#${value}`;
}

/** FortuneSheet's border style codes, by position. */
const BORDER_STYLES = ['none', 'thin', 'hair', 'dotted', 'dashed', 'dashDot', 'dashDotDot', 'double', 'medium', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot', 'thick'];

/** @param {any} side an ExcelJS border side */
function modelBorderSide(side) {
  if (!side?.style || side.style === 'none') return null;
  const style = BORDER_STYLES.indexOf(side.style);
  return { style: style > 0 ? style : 1, color: colorHex(side.color?.argb) || '#000000' };
}

/** @param {any} side a FortuneSheet border side */
export function excelBorderSide(side) {
  if (!side || !Number(side.style)) return undefined;
  const hex = /^#[0-9a-f]{6}$/i.test(side.color || '') ? side.color.slice(1).toUpperCase() : '000000';
  return { style: BORDER_STYLES[Number(side.style)] || 'thin', color: { argb: `FF${hex}` } };
}

/**
 * Every cell's borders as `"r_c" -> { l, r, t, b }`: what the page resolved from
 * the grid when it sent them, otherwise the cell entries read from the file.
 * @param {any} sheet
 * @returns {Map<string, any>}
 */
export function cellBorders(sheet) {
  const map = new Map();
  if (sheet.borders && typeof sheet.borders === 'object') {
    for (const [key, sides] of Object.entries(sheet.borders)) if (sides) map.set(key, sides);
    return map;
  }
  for (const entry of sheet.config?.borderInfo || []) {
    if (entry?.rangeType !== 'cell' || !entry.value) continue;
    const { row_index: r, col_index: c, l, r: right, t, b } = entry.value;
    map.set(`${r}_${c}`, { l, r: right, t, b });
  }
  return map;
}

/** @param {import('exceljs').Cell} cell */
function excelStyle(cell) {
  /** @type {Record<string, unknown>} */
  const style = {};
  const font = cell.font || {};
  if (font.bold) style.bl = 1;
  if (font.italic) style.it = 1;
  if (font.underline) style.un = 1;
  if (font.strike) style.cl = 1;
  if (font.size) style.fs = Number(font.size);
  if (font.name) style.ff = font.name;
  const fontColor = colorHex(font.color?.argb);
  if (fontColor) style.fc = fontColor;
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const bg = colorHex(fill.fgColor?.argb);
    if (bg) style.bg = bg;
  }
  const align = cell.alignment || {};
  if (align.horizontal === 'center') style.ht = 0;
  else if (align.horizontal === 'right') style.ht = 2;
  else if (align.horizontal === 'left') style.ht = 1;
  if (align.vertical === 'middle') style.vt = 0;
  else if (align.vertical === 'bottom') style.vt = 2;
  else if (align.vertical === 'top') style.vt = 1;
  if (align.wrapText) style.tb = '2';
  return style;
}

/** @param {import('exceljs').Cell} cell */
function excelCellValue(cell) {
  const raw = cell.value;
  if (raw === null || raw === undefined) return null;
  if (raw instanceof Date) {
    return { v: dateToSerial(raw), type: 'd' };
  }
  if (typeof raw === 'object') {
    if ('formula' in raw || 'sharedFormula' in raw) {
      const result = raw.result;
      if (result instanceof Date) return { v: dateToSerial(result), type: 'd', f: raw.formula || cell.formula };
      if (result && typeof result === 'object') {
        if ('error' in result) return { v: String(result.error), type: 'e', f: raw.formula || cell.formula };
        if ('richText' in result) return { v: result.richText.map((part) => part.text).join(''), type: 's', f: raw.formula || cell.formula };
        if ('text' in result) return { v: String(result.text), type: 's', f: raw.formula || cell.formula };
        return { v: String(result), type: 's', f: raw.formula || cell.formula };
      }
      return { v: result ?? undefined, type: typeof result === 'number' ? 'n' : typeof result === 'boolean' ? 'b' : 's', f: raw.formula || cell.formula };
    }
    if ('richText' in raw) return { v: raw.richText.map((part) => part.text).join(''), type: 's' };
    if ('text' in raw) return { v: String(raw.text), type: 's' };
    if ('error' in raw) return { v: String(raw.error), type: 'e' };
    if ('hyperlink' in raw) return { v: String(raw.text ?? raw.hyperlink), type: 's' };
  }
  if (typeof raw === 'number') return { v: raw, type: 'n' };
  if (typeof raw === 'boolean') return { v: raw, type: 'b' };
  return { v: String(raw), type: 's' };
}

/** @param {import('exceljs').Worksheet} worksheet */
export function excelSheetToFortune(worksheet, order, sheetNames) {
  const name = worksheet.name;
  const celldata = [];
  let maxRow = 0;
  let maxCol = 0;
  const borderInfo = [];
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const border = cell.border;
      const sides = border ? { l: modelBorderSide(border.left), r: modelBorderSide(border.right), t: modelBorderSide(border.top), b: modelBorderSide(border.bottom) } : undefined;
      if (sides && (sides.l || sides.r || sides.t || sides.b)) {
        borderInfo.push({ rangeType: 'cell', value: { row_index: rowNumber - 1, col_index: colNumber - 1, ...sides } });
      }
      let value = excelCellValue(cell);
      // An empty cell still shows its fill.
      if (!value && cell.fill?.type === 'pattern' && cell.fill.pattern === 'solid' && cell.fill.fgColor?.argb) value = { type: 's' };
      if (!value) return;
      maxRow = Math.max(maxRow, rowNumber - 1);
      maxCol = Math.max(maxCol, colNumber - 1);
      /** @type {any} */
      const cellData = {};
      if (value.v !== undefined) cellData.v = value.v;
      if (value.f) cellData.f = modelFormula(value.f, sheetNames);
      const numFmt = cell.numFmt || 'General';
      const dateType = value.type === 'd' || looksDateFormat(numFmt);
      cellData.ct = { fa: numFmt || 'General', t: dateType ? 'd' : value.type };
      const style = excelStyle(cell);
      Object.assign(cellData, style);
      if (value.type === 'e') cellData.m = String(value.v);
      celldata.push({ r: rowNumber - 1, c: colNumber - 1, v: cellData });
    });
  });
  const lookup = new Map(celldata.map((entry) => [cellKey(entry.r, entry.c), entry]));

  const merges = {};
  for (const range of worksheet.model?.merges || []) {
    const [start, end] = String(range).split(':');
    const s = XLSX.utils.decode_cell(start);
    const e = XLSX.utils.decode_cell(end);
    const rs = e.r - s.r + 1;
    const cs = e.c - s.c + 1;
    if (rs < 1 || cs < 1) continue;
    merges[`${s.r}_${s.c}`] = { r: s.r, c: s.c, rs, cs };
    const main = lookup.get(cellKey(s.r, s.c));
    if (main) main.v.mc = { r: s.r, c: s.c, rs, cs };
    else celldata.push({ r: s.r, c: s.c, v: { mc: { r: s.r, c: s.c, rs, cs }, ct: { fa: 'General', t: 's' } } });
    maxRow = Math.max(maxRow, s.r + rs - 1);
    maxCol = Math.max(maxCol, s.c + cs - 1);
  }

  const rowlen = {};
  const rowhidden = {};
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (row.height) rowlen[rowNumber - 1] = Math.round((row.height * 96) / 72);
    if (row.hidden) rowhidden[rowNumber - 1] = 0;
  });
  const columnlen = {};
  const colhidden = {};
  // An empty sheet has no columns at all: ExcelJS gives null rather than a list.
  (worksheet.columns || []).forEach((column, index) => {
    if (column.width) columnlen[index] = Math.round(column.width * 7);
    if (column.hidden) colhidden[index] = 0;
  });

  const view = worksheet.views?.find((entry) => entry.state?.startsWith('frozen'));
  const frozen = view ? frozenFromSplit(Number(view.ySplit || 0), Number(view.xSplit || 0)) : undefined;

  return finishSheet({
    name,
    id: sheetId(order),
    order,
    row: Math.max(maxRow + 120, 200),
    column: Math.max(maxCol + 26, 40),
    celldata,
    config: { merge: merges, rowlen, columnlen, rowhidden, colhidden, borderInfo },
    frozen,
    status: order === 0 ? 1 : 0,
    defaultRowHeight: 24,
    defaultColWidth: 86,
    showGridLines: true,
  });
}

/**
 * The id of the n-th sheet of a file. Stable across loads, so a sheet the page
 * sends back is matched to the same sheet of the file it came from.
 * @param {number} order
 */
export function sheetId(order) {
  return `sheet-${order}`;
}

/** A numeric key for a cell address; Excel has at most 16,384 columns. */
export function cellKey(r, c) {
  return r * 20_000 + c;
}

/** File freeze panes count rows/columns; FortuneSheet's focus is the last frozen index. */
function frozenFromSplit(ySplit, xSplit) {
  if (!ySplit && !xSplit) return undefined;
  return {
    type: xSplit && ySplit ? 'rangeBoth' : xSplit ? 'rangeColumn' : 'rangeRow',
    range: { row_focus: Math.max(0, ySplit - 1), column_focus: Math.max(0, xSplit - 1) },
  };
}

/** @returns {{ rows: number, columns: number }} rows and columns held frozen in a file. */
export function frozenSplit(frozen) {
  if (!frozen?.type) return { rows: 0, columns: 0 };
  const focusRow = Number(frozen.range?.row_focus ?? 0) + 1;
  const focusColumn = Number(frozen.range?.column_focus ?? 0) + 1;
  switch (frozen.type) {
    case 'row': return { rows: 1, columns: 0 };
    case 'column': return { rows: 0, columns: 1 };
    case 'both': return { rows: 1, columns: 1 };
    case 'rangeRow': return { rows: focusRow, columns: 0 };
    case 'rangeColumn': return { rows: 0, columns: focusColumn };
    case 'rangeBoth': return { rows: focusRow, columns: focusColumn };
    default: return { rows: 0, columns: 0 };
  }
}

/**
 * Prepare a sheet for the grid. FortuneSheet paints `m`, so a cell without it is
 * blank however correct `v` is. Its `calcChain` is what lets an edited cell
 * recalculate the formulas that depend on it.
 * @param {any} sheet
 */
function finishSheet(sheet) {
  const calcChain = [];
  for (const entry of sheet.celldata) {
    const cell = entry.v;
    if (!cell) continue;
    if (cell.f) calcChain.push({ r: entry.r, c: entry.c, id: sheet.id });
    if (cell.m === undefined && cell.v !== undefined && cell.v !== null) {
      cell.m = String(displayValue(cell));
    }
  }
  sheet.calcChain = calcChain;
  return sheet;
}

/** @param {import('exceljs').Workbook} workbook */
function excelWorkbookToSheets(workbook) {
  const names = workbook.worksheets.map((worksheet) => worksheet.name);
  return workbook.worksheets.map((worksheet, index) => excelSheetToFortune(worksheet, index, names));
}

/** Map a SheetJS worksheet to a FortuneSheet sheet. */
function sheetJsSheetToFortune(worksheet, name, order, sheetNames = [name], options = {}) {
  const celldata = [];
  const range = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const address = XLSX.utils.encode_cell({ r, c });
      const source = worksheet[address];
      if (!source) continue;
      /** @type {any} */
      const cell = {};
      if (source.f) cell.f = modelFormula(source.f, sheetNames);
      if (source.v !== undefined) cell.v = source.v;
      // Plain-text tables arrive as strings. Real numbers must calculate; identifiers stay text.
      const parsed = options.inferNumbers && !source.f ? parseNumericText(source.v) : undefined;
      if (parsed !== undefined) {
        cell.v = parsed;
        source.t = 'n';
      }
      const fmt = source.z || 'General';
      const dateType = source.t === 'd' || looksDateFormat(fmt);
      cell.ct = { fa: fmt, t: dateType ? 'd' : source.t === 'n' ? 'n' : source.t === 'b' ? 'b' : source.t === 'e' ? 'e' : 's' };
      if (source.t === 'e') cell.m = String(source.v);
      if (source.s) {
        if (source.s.patternType === 'solid' && source.s.fgColor?.rgb) cell.bg = colorHex(source.s.fgColor.rgb);
        if (source.s.font?.bold) cell.bl = 1;
        if (source.s.font?.italic) cell.it = 1;
      }
      celldata.push({ r, c, v: cell });
    }
  }
  const lookup = new Map(celldata.map((entry) => [cellKey(entry.r, entry.c), entry]));
  const merges = {};
  for (const merge of worksheet['!merges'] || []) {
    const rs = merge.e.r - merge.s.r + 1;
    const cs = merge.e.c - merge.s.c + 1;
    merges[`${merge.s.r}_${merge.s.c}`] = { r: merge.s.r, c: merge.s.c, rs, cs };
    const main = lookup.get(cellKey(merge.s.r, merge.s.c));
    if (main) main.v.mc = { r: merge.s.r, c: merge.s.c, rs, cs };
  }
  const columnlen = {};
  const colhidden = {};
  (worksheet['!cols'] || []).forEach((col, index) => {
    if (!col) return;
    const width = col.wpx || (col.wch ? col.wch * 7 : col.width ? col.width * 7 : 0);
    if (width) columnlen[index] = Math.round(width);
    if (col.hidden) colhidden[index] = 0;
  });
  const rowlen = {};
  const rowhidden = {};
  (worksheet['!rows'] || []).forEach((row, index) => {
    if (!row) return;
    if (row.hpx) rowlen[index] = Math.round(row.hpx);
    else if (row.hpt) rowlen[index] = Math.round((row.hpt * 96) / 72);
    if (row.hidden) rowhidden[index] = 0;
  });
  return finishSheet({
    name,
    id: sheetId(order),
    order,
    row: Math.max(range.e.r + 120, 200),
    column: Math.max(range.e.c + 26, 40),
    celldata,
    config: { merge: merges, rowlen, columnlen, rowhidden, colhidden },
    status: order === 0 ? 1 : 0,
    defaultRowHeight: 24,
    defaultColWidth: 86,
    showGridLines: true,
  });
}

/** @param {Uint8Array} bytes @param {string} label */
function parseXlsxBytes(bytes, label) {
  const workbook = new ExcelJS.Workbook();
  return workbook.xlsx.load(bytes).then(() => ({
    sheets: excelWorkbookToSheets(workbook),
    source: 'xlsx',
    label,
  })).catch((error) => {
    throw httpError(422, 'corrupt-file', friendlyMessage(error), error);
  });
}

/** @param {Uint8Array} bytes @param {string} label */
function parseXlsBytes(bytes, label) {
  try {
    const workbook = XLSX.read(bytes, {
      type: 'buffer',
      cellFormula: true,
      cellNF: true,
      cellStyles: true,
      cellDates: false,
      raw: true,
      dense: false,
    });
    return {
      sheets: workbook.SheetNames.map((name, index) => sheetJsSheetToFortune(workbook.Sheets[name], name, index, workbook.SheetNames)),
      source: 'xls',
      label,
    };
  } catch (error) {
    throw httpError(422, 'corrupt-file', friendlyMessage(error), error);
  }
}

/** @param {string} text @param {string} label */
function parseDelimitedText(text, label) {
  const workbook = XLSX.read(text, { type: 'string', raw: true, cellFormula: true, cellNF: true, cellStyles: false });
  const names = workbook.SheetNames.length ? workbook.SheetNames : ['Sheet1'];
  return {
    sheets: names.map((name, index) => sheetJsSheetToFortune(workbook.Sheets[name], index === 0 ? (name === 'Sheet1' ? t('数据') : name) : name, index, names, { inferNumbers: true })),
    source: 'text',
    label,
  };
}

/** @param {string} text @param {string} label */
function parseHtmlText(text, label) {
  const workbook = XLSX.read(text, { type: 'string', cellFormula: true, cellNF: true, cellStyles: false });
  return {
    sheets: workbook.SheetNames.map((name, index) => sheetJsSheetToFortune(workbook.Sheets[name], name, index, workbook.SheetNames)),
    source: 'html',
    label,
  };
}

/** @param {string} text */
function parsePlainLines(text) {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  const celldata = [];
  lines.slice(0, 200_000).forEach((line, r) => {
    if (line === '') return;
    celldata.push({ r, c: 0, v: { v: line, ct: { fa: 'General', t: 's' } } });
  });
  return {
    sheets: [finishSheet({
      name: t('文本'),
      id: sheetId(0),
      order: 0,
      row: Math.max(lines.length + 50, 200),
      column: 30,
      celldata,
      config: {},
      status: 1,
      showGridLines: true,
    })],
    source: 'text',
    label: t('纯文本'),
  };
}

/**
 * Count formula cells and choose whether recalculation is worthwhile.
 * @param {any[]} sheets
 */
export function formulaAddresses(sheets) {
  const result = [];
  for (const sheet of sheets) {
    for (const entry of sheet.celldata || []) {
      if (entry?.v?.f) result.push({ sheet: sheet.name, address: XLSX.utils.encode_cell({ r: entry.r, c: entry.c }) });
    }
  }
  return result;
}

/**
 * Parse any supported spreadsheet/table source.
 * @param {string} path
 * @param {{ recalc?: boolean|'later', maxCells?: number }} [options] `recalc` true forces a
 *   LibreOffice recalculation, false skips it, 'later' skips it and says so in
 *   `stats.pendingCalculation`; by default it runs only when some formula has no saved result.
 */
export async function loadSpreadsheet(path, options = {}) {
  if (!fileExists(path)) throw httpError(404, 'not-found', t('文件不存在'));
  const bytes = readBytes(path);
  const ext = extensionOf(path);
  const container = sniffContainer(bytes);
  const warnings = [];
  let parsed;

  if (container === 'pdf') throw httpError(415, 'wrong-kind', t('这是一个 PDF 文件，不是表格。'));
  if (container === 'zip') {
    parsed = await parseXlsxBytes(bytes, ext === '.et' ? 'WPS ET (OOXML)' : 'XLSX');
  } else if (container === 'ole') {
    parsed = parseXlsBytes(bytes, ext === '.et' ? 'WPS ET (BIFF)' : 'XLS (BIFF)');
  } else if (container === 'html') {
    const { text, encoding } = decodeText(bytes);
    parsed = parseHtmlText(text, t('网页/HTML 表格'));
    warnings.push({ code: 'html-table', message: t('这是网页或 HTML 表格伪装成的 {0} 文件，已按表格内容读取（{1}）。', ext.slice(1) || t('表格'), encoding.toUpperCase()) });
  } else if (container === 'text' || container === 'json') {
    const { text, encoding } = decodeText(bytes);
    const kind = looksDelimited(text);
    if (kind) {
      parsed = parseDelimitedText(text, kind === 'tsv' ? 'TSV' : 'CSV');
      // Nothing to say about the ordinary case; a legacy encoding is worth pointing out.
      if (encoding !== 'utf-8') warnings.push({ code: 'encoding', message: t('文本表格已按 {0} 解码，保存时统一使用 UTF-8。', encoding.toUpperCase()) });
    } else {
      parsed = parsePlainLines(text);
      warnings.push({ code: 'plain-text', message: t('该文件是纯文本（{0}），已放进单列表格。', encoding.toUpperCase()) });
    }
  } else {
    throw httpError(415, 'unsupported', t('无法识别 {0} 的内容格式。', ext || t('该文件')));
  }

  let sheets = parsed.sheets;
  const formulas = formulaAddresses(sheets);
  let recalculated = false;
  let missingFonts = [];
  const missingResults = formulas.length > 0 && sheets.some((sheet) => sheet.celldata.some((entry) => entry.v?.f && (entry.v.v === undefined || entry.v.v === null || entry.v.v === '')));
  // 'later': the page shows the workbook at once and asks again for the results.
  const later = options.recalc === 'later';
  if (formulas.length > 0 && (options.recalc === true || (options.recalc !== false && !later && missingResults))) {
    let recalcSource = path;
    let tempDir;
    try {
      if (ext === '.et') {
        tempDir = mkdtempSync(join(tmpdir(), 'dsh-doc-et-'));
        recalcSource = join(tempDir, `wps-${token().slice(0, 8)}${container === 'zip' ? '.xlsx' : '.xls'}`);
        writeFileSync(recalcSource, bytes);
      }
      const recalc = await recalculateWorkbook(recalcSource);
      try {
        const values = XLSX.read(readFileSync(recalc.path), {
          type: 'buffer', cellFormula: true, cellNF: true, cellDates: false, raw: true,
        });
        for (const sheet of sheets) {
          const valueSheet = values.Sheets[sheet.name];
          if (!valueSheet) continue;
          for (const entry of sheet.celldata) {
            if (!entry?.v?.f) continue;
            const source = valueSheet[XLSX.utils.encode_cell({ r: entry.r, c: entry.c })];
            if (source && source.v !== undefined) {
              entry.v.v = source.v;
              const isDate = source.t === 'd' || looksDateFormat(entry.v.ct?.fa);
              entry.v.ct = { ...(entry.v.ct || { fa: 'General' }), t: isDate ? 'd' : source.t === 'n' ? 'n' : source.t === 'b' ? 'b' : source.t === 'e' ? 'e' : 's' };
              delete entry.v.m;
              entry.v.m = source.t === 'e' ? String(source.w ?? source.v) : String(displayValue(entry.v));
            }
          }
        }
        recalculated = true;
      } finally {
        recalc.cleanup?.();
      }
      // Fonts matter when exporting a PDF, where they are reported; the grid paints with its own.
      missingFonts = recalc.missingFonts || [];
    } catch (error) {
      warnings.push({ code: 'recalc', message: t('LibreOffice 重算没有成功（{0}），表格里的公式改由内置引擎计算，个别函数可能不支持。', error instanceof Error ? error.message : String(error)) });
    } finally {
      if (tempDir) rmSync(tempDir, { recursive: true, force: true });
    }
  }

  const cellCount = sheets.reduce((sum, sheet) => sum + (sheet.celldata?.length || 0), 0);
  if (cellCount > (options.maxCells || 1_000_000)) {
    warnings.push({ code: 'large', message: t('表格较大（{0} 个非空单元格），操作可能会变慢。', cellCount.toLocaleString()) });
  }
  if (sheets.length === 0) throw httpError(422, 'empty-workbook', t('这个工作簿里没有工作表，文件可能已损坏。'));

  return {
    kind: 'spreadsheet',
    path,
    name: basename(path),
    ext,
    format: parsed.source,
    label: parsed.label,
    sheets,
    sheetNames: sheets.map((sheet) => sheet.name),
    sheetSizes: sheets.map((sheet) => {
      const bounds = sheetBounds(sheet);
      return { name: sheet.name, rows: bounds.e.r + 1, columns: bounds.e.c + 1 };
    }),
    stats: {
      sheets: sheets.length,
      cells: cellCount,
      formulas: formulas.length,
      recalculated,
      // Formulas with no result yet, which the grid's own engine has to compute.
      unresolved: missingResults && !recalculated,
      // LibreOffice was not asked yet; opening again without `later` gets its results.
      pendingCalculation: later && missingResults,
    },
    warnings,
    missingFonts,
  };
}

/**
 * Convert a FortuneSheet sheet into an ExcelJS worksheet.
 * @param {import('exceljs').Workbook} workbook
 * @param {any} sheet
 * @param {string} name a name valid for and unique in the workbook
 */
export function writeFortuneSheet(workbook, sheet, name) {
  const split = frozenSplit(sheet.frozen);
  const worksheet = workbook.addWorksheet(name, {
    views: split.rows || split.columns ? [{ state: 'frozen', xSplit: split.columns, ySplit: split.rows }] : undefined,
  });
  const merges = sheet.config?.merge || {};
  for (const merge of Object.values(merges)) {
    if (!merge || !merge.rs || !merge.cs) continue;
    if (merge.rs === 1 && merge.cs === 1) continue;
    worksheet.mergeCells(merge.r + 1, merge.c + 1, merge.r + merge.rs, merge.c + merge.cs);
  }
  for (const entry of sheet.celldata || []) {
    if (!entry?.v) continue;
    const row = Number(entry.r) + 1;
    const col = Number(entry.c) + 1;
    if (row < 1 || col < 1) continue;
    const data = entry.v;
    // Cells covered by a merge forward their writes to the master cell in ExcelJS.
    if (data.mc && (data.mc.r !== entry.r || data.mc.c !== entry.c)) continue;
    const cell = worksheet.getCell(row, col);
    const formula = typeof data.f === 'string' && data.f.trim() !== '' ? fileFormula(data.f) : undefined;
    const literal = cellLiteral(data);
    if (formula) {
      cell.value = literal === undefined ? { formula } : { formula, result: literal };
    } else if (literal !== undefined) {
      cell.value = literal;
    }
    if (data.ct?.fa && data.ct.fa !== 'General') cell.numFmt = data.ct.fa;
    const font = {};
    if (data.bl) font.bold = true;
    if (data.it) font.italic = true;
    if (data.un) font.underline = true;
    if (data.cl) font.strike = true;
    if (data.fs) font.size = Number(data.fs);
    if (data.ff) font.name = data.ff;
    if (data.fc && /^#[0-9a-f]{6}$/i.test(data.fc)) font.color = { argb: `FF${data.fc.slice(1).toUpperCase()}` };
    if (Object.keys(font).length) cell.font = font;
    if (data.bg && /^#[0-9a-f]{6}$/i.test(data.bg)) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${data.bg.slice(1).toUpperCase()}` } };
    }
    const alignment = excelAlignment(data);
    if (alignment) cell.alignment = alignment;
  }
  for (const [key, sides] of cellBorders(sheet)) {
    const [r, c] = key.split('_').map(Number);
    const border = {};
    if (excelBorderSide(sides.l)) border.left = excelBorderSide(sides.l);
    if (excelBorderSide(sides.r)) border.right = excelBorderSide(sides.r);
    if (excelBorderSide(sides.t)) border.top = excelBorderSide(sides.t);
    if (excelBorderSide(sides.b)) border.bottom = excelBorderSide(sides.b);
    if (Object.keys(border).length) worksheet.getCell(r + 1, c + 1).border = border;
  }
  const columnlen = sheet.config?.columnlen || {};
  const colhidden = sheet.config?.colhidden || {};
  for (const [index, width] of Object.entries(columnlen)) {
    const column = worksheet.getColumn(Number(index) + 1);
    column.width = Math.max(4, Math.round((Number(width) / 7) * 100) / 100);
  }
  for (const index of Object.keys(colhidden)) worksheet.getColumn(Number(index) + 1).hidden = true;
  const rowlen = sheet.config?.rowlen || {};
  const rowhidden = sheet.config?.rowhidden || {};
  for (const [index, height] of Object.entries(rowlen)) {
    worksheet.getRow(Number(index) + 1).height = Math.round(((Number(height) * 72) / 96) * 100) / 100;
  }
  for (const index of Object.keys(rowhidden)) worksheet.getRow(Number(index) + 1).hidden = true;
  if (Array.isArray(sheet.images) && sheet.images.length > 0) {
    for (const image of sheet.images) {
      if (!image?.src || !String(image.src).startsWith('data:image/')) continue;
      try {
        const base64 = String(image.src).split(',')[1] || '';
        const imageId = workbook.addImage({ base64, extension: String(image.src).includes('image/jpeg') ? 'jpeg' : 'png' });
        worksheet.addImage(imageId, {
          tl: { col: Number(image.left || 0) / (sheet.defaultColWidth || 86), row: Number(image.top || 0) / (sheet.defaultRowHeight || 24) },
          ext: { width: Number(image.width || 640), height: Number(image.height || 360) },
        });
      } catch { /* an unsupported image must not fail the workbook */ }
    }
  }
  return worksheet;
}

/**
 * The ExcelJS alignment for a model cell, or undefined when it has none.
 * FortuneSheet: ht 0 centre, 1 left, 2 right; vt 0 middle, 1 top, 2 bottom; tb "2" wraps.
 * @param {any} data
 */
export function excelAlignment(data) {
  const ht = data.ht === undefined || data.ht === null || data.ht === '' ? undefined : Number(data.ht);
  const vt = data.vt === undefined || data.vt === null || data.vt === '' ? undefined : Number(data.vt);
  const wrap = String(data.tb) === '2';
  if (ht === undefined && vt === undefined && !wrap) return undefined;
  return {
    horizontal: ht === 0 ? 'center' : ht === 2 ? 'right' : ht === 1 ? 'left' : undefined,
    vertical: vt === 0 ? 'middle' : vt === 2 ? 'bottom' : vt === 1 ? 'top' : undefined,
    wrapText: wrap || undefined,
  };
}

/**
 * Undo what `modelFormula` did for the grid: the leading `=`, `TRUE()`/`FALSE()`, and
 * the prefix newer functions must carry in a file. String literals are left as written.
 * @param {string} formula
 */
export function fileFormula(formula) {
  return formula.trim().replace(/^=/, '').split(/("(?:[^"]|"")*")/).map((part, index) => (index % 2 === 1
    ? part
    : part
      .replace(/\b(TRUE|FALSE)\(\)/gi, '$1')
      .replace(NEWER_FUNCTION_CALL, (name) => `${NEWER_FUNCTIONS[name.toUpperCase()]}${name.toUpperCase()}`))).join('');
}

/**
 * What is wrong with a formula that would still be written to a file, if anything. A
 * spreadsheet program may quietly repair `=SUM(A1:A3` or refuse the whole file, so
 * it is stopped here. Unknown functions are found later, by calculating.
 * @param {string} formula
 * @param {{ sheet: string, r: number, c: number }} [cell] where it is going
 * @returns {string|undefined}
 */
export function formulaProblem(formula, cell) {
  const text = String(formula).trim().replace(/^=/, '').trim();
  if (text === '') return t('公式是空的');
  let depth = 0;
  let quote;
  let code = '';
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (quote) {
      if (char !== quote) continue;
      if (text[index + 1] === quote) index += 1;
      else quote = undefined;
    } else if (char === '"' || char === "'") {
      quote = char;
      // A quoted sheet name stays, to tell which sheet a reference is on.
      code += char === "'" ? '§' : '0';
    } else {
      if (char === '(') depth += 1;
      if (char === ')') depth -= 1;
      if (depth < 0) return t('多了一个右括号');
      code += char;
    }
  }
  if (quote) return quote === '"' ? t('引号没有闭合') : t('工作表名的引号没有闭合');
  if (depth > 0) return t('少了右括号');
  if (/[+\-*/^&=<>,(]\s*$/.test(code)) return t('公式没有写完');
  if (!cell) return undefined;
  const reference = /(?<![\w.$§!])(?:([\p{L}\p{N}_§]+)!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![\w(])/giu;
  for (const match of code.matchAll(reference)) {
    if (match[1] && match[1] !== '§' && match[1] !== cell.sheet) continue;
    if (match[1] === '§') continue;
    const start = XLSX.utils.decode_cell(`${match[2]}${match[3]}`.toUpperCase());
    const end = match[4] ? XLSX.utils.decode_cell(`${match[4]}${match[5]}`.toUpperCase()) : start;
    if (cell.r >= Math.min(start.r, end.r) && cell.r <= Math.max(start.r, end.r) && cell.c >= Math.min(start.c, end.c) && cell.c <= Math.max(start.c, end.c)) {
      return t('公式引用了它自己所在的单元格（循环引用）');
    }
  }
  return undefined;
}

/**
 * The plain value of a model cell. Text edited as rich text keeps its runs in
 * `ct.s` and has no `v`.
 * @param {any} data
 * @returns {string|number|boolean|undefined}
 */
export function cellLiteral(data) {
  // The grid keeps typed input as text even when it has classified it as a number.
  if (typeof data.v === 'string' && data.ct?.t === 'n' && data.v.trim() !== '' && Number.isFinite(Number(data.v.replaceAll(',', '')))) {
    return Number(data.v.replaceAll(',', ''));
  }
  if (typeof data.v === 'string' && data.ct?.t === 'b' && /^(true|false)$/i.test(data.v)) return /^true$/i.test(data.v);
  if (data.v !== undefined && data.v !== null && data.v !== '') return data.v;
  if (Array.isArray(data.ct?.s)) return data.ct.s.map((run) => run?.v ?? '').join('');
  return undefined;
}

/**
 * Accept both shapes of a sheet: `celldata` as sent to the grid, or the `data`
 * matrix the grid keeps after loading.
 * @param {any} sheet
 */
export function withCelldata(sheet) {
  if (Array.isArray(sheet?.celldata) && sheet.celldata.length > 0) return sheet;
  if (!Array.isArray(sheet?.data)) return { ...sheet, celldata: sheet?.celldata || [] };
  const celldata = [];
  sheet.data.forEach((row, r) => row?.forEach((v, c) => { if (v !== null && v !== undefined) celldata.push({ r, c, v }); }));
  const { data: _matrix, ...rest } = sheet;
  return { ...rest, celldata };
}

/**
 * Sheet names as the file formats require them: at most 31 characters, none of
 * `[]:*?/\\`, unique ignoring case.
 * @param {any[]} sheets
 * @returns {string[]}
 */
export function workbookSheetNames(sheets) {
  const used = new Set();
  return sheets.map((sheet, index) => {
    const base = String(sheet.name || `Sheet${index + 1}`).replace(/[[\]:*?/\\]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31) || `Sheet${index + 1}`;
    let name = base;
    for (let suffix = 2; used.has(name.toLowerCase()); suffix += 1) {
      const tail = ` (${suffix})`;
      name = `${base.slice(0, 31 - tail.length)}${tail}`;
    }
    used.add(name.toLowerCase());
    return name;
  });
}

/** @param {any} sheets */
function requireSheets(sheets) {
  if (!Array.isArray(sheets) || sheets.length === 0) {
    throw httpError(400, 'no-sheet', t('没有收到任何工作表，已取消保存，原文件未改动。'));
  }
  return sheets.map(withCelldata);
}

/**
 * Build an XLSX buffer from FortuneSheet sheets.
 * @param {any[]} sheets
 * @param {{ creator?: string, print?: { title: string } }} [options] `print` lays the sheets out for a PDF
 */
export async function buildXlsx(sheets, options = {}) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = options.creator || t('DeepSeek Harness 文档查看器');
  workbook.created = new Date();
  // Excel, WPS and Numbers recalculate on open, so a stale saved result never shows.
  workbook.calcProperties = { fullCalcOnLoad: true };
  const complete = requireSheets(sheets);
  const names = workbookSheetNames(complete);
  complete.forEach((sheet, index) => {
    const worksheet = writeFortuneSheet(workbook, sheet, names[index]);
    // For a PDF, empty sheets are left out and the rest are laid out for paper.
    if (options.print && !layoutSheetForPrint(worksheet, sheet, { title: options.print.title, sheetCount: complete.length, display: displayValue })) {
      workbook.removeWorksheet(worksheet.id);
    }
  });
  if (options.print && workbook.worksheets.length === 0) throw httpError(400, 'empty-workbook', t('表格里没有内容，没有可导出的页面。'));
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

/**
 * Write a workbook in the requested format. All paths are new; the caller
 * chooses a unique output path.
 * @param {any[]} sheets
 * @param {string} output
 * @param {'xlsx'|'xls'|'csv'} format
 * @param {{ sheetName?: string, csvEncoding?: string }} [options]
 */
export async function writeSpreadsheet(sheets, output, format, options = {}) {
  ensureDir(dirname(output));
  if (format === 'xlsx') {
    writeFileSync(output, await buildXlsx(sheets));
    return;
  }
  if (format === 'xls') {
    await writeXlsWithPython(sheets, output);
    return;
  }
  if (format === 'csv') {
    const complete = requireSheets(sheets);
    const sheet = complete.find((entry) => entry.name === options.sheetName) || complete[0];
    writeFileSync(output, csvFromSheet(sheet, 'utf8'));
    return;
  }
  throw httpError(400, 'bad-format', t('不支持的表格输出格式：{0}', format));
}

/** @param {any} sheet @param {'utf8'|'gbk'} encoding */
export function csvFromSheet(sheet, encoding = 'utf8', delimiter = ',') {
  const rows = sheetToRows(sheet);
  const escape = (value) => {
    const text = value === undefined || value === null ? '' : String(value);
    return text.includes(delimiter) || /["\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  };
  const text = rows.map((row) => row.map(escape).join(delimiter)).join('\r\n');
  // 'plain' leaves the byte order mark to the caller, which knows the target encoding.
  return encoding === 'plain' || encoding === 'gbk' ? text : `\uFEFF${text}`;
}

/**
 * Materialize a FortuneSheet sheet as a dense row array (values, not formulas).
 * @param {any} sheet
 */
export function sheetToRows(sheet) {
  let maxRow = 0;
  let maxCol = 0;
  const cells = new Map();
  for (const entry of sheet.celldata || []) {
    if (!entry?.v) continue;
    if (entry.v.mc && (entry.v.mc.r !== entry.r || entry.v.mc.c !== entry.c)) continue;
    const value = displayValue(entry.v);
    if (value === undefined || value === '') continue;
    cells.set(`${entry.r}:${entry.c}`, value);
    maxRow = Math.max(maxRow, entry.r);
    maxCol = Math.max(maxCol, entry.c);
  }
  const rows = [];
  for (let r = 0; r <= maxRow; r += 1) {
    const row = [];
    for (let c = 0; c <= maxCol; c += 1) row.push(cells.get(`${r}:${c}`) ?? '');
    rows.push(row);
  }
  return rows;
}

/** Format a cell's display value using the number format where practical. */
export function displayValue(cell) {
  if (!cell) return '';
  if (cell.m !== undefined && cell.m !== null && cell.m !== '') return cell.m;
  if (cell.v === undefined || cell.v === null) return '';
  if (cell.ct?.t === 'd' || looksDateFormat(cell.ct?.fa)) {
    if (typeof cell.v === 'number') {
      const format = cell.ct?.fa && cell.ct.fa !== 'General' ? cell.ct.fa : 'yyyy-mm-dd';
      try { return SSF.format(format, cell.v); } catch { return serialToDate(cell.v).toISOString().slice(0, 10); }
    }
  }
  if (typeof cell.v === 'boolean') return cell.v ? 'TRUE' : 'FALSE';
  if (typeof cell.v === 'number' && cell.ct?.fa && cell.ct.fa !== 'General') {
    try { return SSF.format(cell.ct.fa, cell.v); } catch { return String(cell.v); }
  }
  return String(cell.v);
}

/**
 * Write BIFF8 XLS with formulas, merges, number formats, freeze panes and
 * basic styles. xlwt is bundled under vendor/python because SheetJS CE's BIFF
 * writer drops formulas.
 * @param {any[]} sheets
 * @param {string} output
 */
export async function writeXlsWithPython(sheets, output) {
  const work = mkdtempSync(join(tmpdir(), 'dsh-doc-xls-'));
  try {
    const payload = join(work, 'workbook.json');
    const script = join(HERE, 'xls-writer.py');
    if (!existsSync(script)) throw httpError(500, 'missing-writer', t('缺少 XLS 写入组件'));
    const complete = requireSheets(sheets);
    const names = workbookSheetNames(complete).map((name) => name);
    writeFileSync(payload, JSON.stringify({
      sheets: complete.map((sheet, index) => ({
        name: names[index],
        celldata: sheet.celldata,
        config: sheet.config,
        frozenRows: frozenSplit(sheet.frozen).rows,
        frozenColumns: frozenSplit(sheet.frozen).columns,
      })),
    }));
    const python = await findPython();
    await execFileAsync(python, [script, payload, output], {
      env: { ...process.env, PYTHONPATH: join(HERE, '..', '..', 'vendor', 'python') },
      timeout: 120_000,
      maxBuffer: 2 * 1024 * 1024,
    });
    if (!fileExists(output)) throw httpError(500, 'xls-failed', t('没有生成 XLS 文件'));
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

/** @returns {Promise<string>} */
async function findPython() {
  for (const candidate of [process.env.DSH_PYTHON, '/usr/bin/python3', 'python3']) {
    if (!candidate) continue;
    try {
      await execFileAsync(candidate, ['-c', 'import sys; print(sys.version)'], { timeout: 3000 });
      return candidate;
    } catch { /* next */ }
  }
  throw httpError(503, 'python-unavailable', t('未找到 Python 3，无法写出旧版 XLS。可改用 XLSX。'));
}

/**
 * Apply cell edits and text-number conversion to a loaded spreadsheet model.
 * @param {any[]} sheets
 * @param {{ sheet?: string, address: string, value?: unknown, formula?: string }[]} edits
 * @param {{ convertTextNumbers?: boolean, range?: string }} [options]
 */
export function applyEdits(sheets, edits, options = {}) {
  const names = sheets.map((sheet) => sheet.name);
  const findSheet = (name) => {
    if (!name) return sheets[0];
    const sheet = sheets.find((entry) => entry.name === name);
    if (!sheet) throw httpError(400, 'bad-sheet', t('没有名为“{0}”的工作表。现有工作表：{1}', name, names.join(t('、'))));
    return sheet;
  };
  const lookups = new Map();
  const lookupOf = (sheet) => {
    if (!lookups.has(sheet)) lookups.set(sheet, new Map(sheet.celldata.map((entry) => [cellKey(entry.r, entry.c), entry])));
    return lookups.get(sheet);
  };
  const applied = [];
  for (const edit of edits || []) {
    const sheet = findSheet(edit.sheet || options.sheet);
    const address = String(edit.address || '').toUpperCase();
    if (!/^[A-Z]{1,3}[1-9]\d*$/.test(address)) throw httpError(400, 'bad-address', t('单元格地址不正确：{0}', edit.address));
    const decoded = XLSX.utils.decode_cell(address);
    // Nothing is written unless every edit can be: these stop the whole call.
    const merge = Object.values(sheet.config?.merge || {}).find((m) => m && decoded.r >= m.r && decoded.r < m.r + m.rs && decoded.c >= m.c && decoded.c < m.c + m.cs);
    if (merge && (merge.r !== decoded.r || merge.c !== decoded.c)) {
      const master = XLSX.utils.encode_cell({ r: merge.r, c: merge.c });
      const last = XLSX.utils.encode_cell({ r: merge.r + merge.rs - 1, c: merge.c + merge.cs - 1 });
      throw httpError(400, 'inside-merge', t('{0} 在合并单元格 {1}:{2} 内部，写入会覆盖 {3} 的内容。要改这个合并单元格，请写 {4}。文件没有改动。', address, master, last, master, master));
    }
    if (typeof edit.formula === 'string' && edit.formula.trim() !== '') {
      const problem = formulaProblem(edit.formula, { sheet: sheet.name, r: decoded.r, c: decoded.c });
      if (problem) throw httpError(400, 'bad-formula', t('{0} 的公式 {1} 有问题：{2}。文件没有改动。', address, edit.formula, problem));
    }
    const lookup = lookupOf(sheet);
    let entry = lookup.get(cellKey(decoded.r, decoded.c));
    if (!entry) {
      entry = { r: decoded.r, c: decoded.c, v: {} };
      sheet.celldata.push(entry);
      lookup.set(cellKey(decoded.r, decoded.c), entry);
      sheet.row = Math.max(sheet.row || 0, decoded.r + 2);
      sheet.column = Math.max(sheet.column || 0, decoded.c + 2);
    }
    const cell = entry.v || (entry.v = {});
    delete cell.m;
    if (typeof edit.formula === 'string' && edit.formula.trim() !== '') {
      cell.f = modelFormula(edit.formula, names);
      delete cell.v;
      cell.ct = cell.ct || { fa: 'General', t: 'n' };
    } else {
      delete cell.f;
      if (edit.value === undefined || edit.value === null) {
        delete cell.v;
      } else {
        cell.v = edit.value;
        cell.m = String(displayValue({ ...cell, ct: { ...(cell.ct || { fa: 'General' }), t: typeof edit.value === 'number' ? 'n' : typeof edit.value === 'boolean' ? 'b' : 's' } }));
      }
      cell.ct = { ...(cell.ct || { fa: 'General' }), t: typeof edit.value === 'number' ? 'n' : typeof edit.value === 'boolean' ? 'b' : 's' };
    }
    applied.push({ sheet: sheet.name, address });
  }
  let converted = 0;
  if (options.convertTextNumbers) {
    for (const sheet of sheets) {
      for (const entry of sheet.celldata) {
        const cell = entry?.v;
        if (!cell || cell.f) continue;
        const number = parseNumericText(cell.v);
        if (number === undefined) continue;
        cell.v = number;
        cell.ct = { ...(cell.ct || { fa: 'General' }), t: 'n' };
        cell.m = String(displayValue({ ...cell, m: undefined }));
        converted += 1;
      }
    }
  }
  return { applied, converted };
}

/**
 * Whether the first row of a range names its columns rather than holding data, so
 * that "sum the selected cells" counts every selected cell. Mirrored by the page's
 * statistics in sheetUtils.js.
 * @param {any[][]} matrix the cells of the range
 * @param {number} startRow the sheet row the range starts on, from 0
 */
function firstRowIsHeader(matrix, startRow) {
  if (matrix.length < 2) return false;
  const isNumber = (cell) => typeof cell?.v === 'number' || parseNumericText(cell?.v) !== undefined;
  const isFilled = (cell) => cell && cell.v !== undefined && cell.v !== null && cell.v !== '';
  const first = matrix[0].filter(isFilled);
  if (first.length === 0 || first.every(isNumber)) return false;
  // The top of a sheet with some text in it reads as a header, as in "部门 | 2024 | 2025".
  if (startRow === 0) return true;
  // Further down, only labels sitting over numbers do; anything else is a row of data.
  if (first.some((cell) => isNumber(cell) || cell.f)) return false;
  return matrix.slice(1).some((row) => row.some((cell, index) => isFilled(matrix[0][index]) && isNumber(cell)));
}

/**
 * Compute summary statistics for a range, optionally grouped.
 * @param {any[]} sheets
 * @param {{ sheet?: string, range?: string, groupBy?: string, valueColumn?: string, agg?: string, hasHeader?: boolean }} request
 */
export function computeStats(sheets, request = {}) {
  const sheet = request.sheet ? sheets.find((entry) => entry.name === request.sheet) : sheets[0];
  if (!sheet) {
    throw httpError(400, 'bad-sheet', request.sheet
      ? t('没有名为“{0}”的工作表。现有工作表：{1}', request.sheet, sheets.map((entry) => entry.name).join(t('、')))
      : t('没有工作表'));
  }
  const bounds = request.range ? decodeRange(request.range) : sheetBounds(sheet);
  const lookup = new Map(sheet.celldata.map((entry) => [cellKey(entry.r, entry.c), entry.v]));
  const matrix = [];
  for (let r = bounds.s.r; r <= bounds.e.r; r += 1) {
    const row = [];
    for (let c = bounds.s.c; c <= bounds.e.c; c += 1) row.push(lookup.get(cellKey(r, c)) ?? null);
    matrix.push(row);
  }
  const hasHeader = request.hasHeader ?? firstRowIsHeader(matrix, bounds.s.r);
  const lettered = (index) => XLSX.utils.encode_col(bounds.s.c + index);
  const header = hasHeader && matrix.length > 0
    ? matrix[0].map((cell, index) => (cell ? String(displayValue(cell)) : '') || lettered(index))
    : (matrix[0] || []).map((_, index) => lettered(index));
  const body = hasHeader ? matrix.slice(1) : matrix;
  const groupIndex = columnIndex(request.groupBy, header, bounds);
  const valueIndex = columnIndex(request.valueColumn, header, bounds);
  if (request.groupBy && groupIndex === undefined) throw httpError(400, 'bad-column', t('分组列“{0}”不在所选范围内。可用的列：{1}', request.groupBy, header.join(t('、'))));
  if (request.valueColumn && valueIndex === undefined) throw httpError(400, 'bad-column', t('数值列“{0}”不在所选范围内。可用的列：{1}', request.valueColumn, header.join(t('、'))));
  if (groupIndex !== undefined && valueIndex === undefined) throw httpError(400, 'bad-column', t('按列分组时还需要指定数值列。'));
  const numberOf = (cell) => {
    if (typeof cell?.v === 'number') return Number.isFinite(cell.v) ? cell.v : undefined;
    return parseNumericText(cell?.v);
  };
  const numeric = [];
  const all = [];
  for (const row of body) {
    for (const cell of valueIndex !== undefined ? [row[valueIndex]] : row) {
      if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;
      const value = numberOf(cell);
      if (value !== undefined) numeric.push(value);
      all.push(String(displayValue(cell)));
    }
  }
  const total = numeric.reduce((sum, value) => sum + value, 0);
  const summary = {
    sheet: sheet.name,
    range: XLSX.utils.encode_range(bounds),
    count: all.length,
    numericCount: numeric.length,
    distinct: new Set(all).size,
    sum: round(total),
    average: numeric.length ? round(total / numeric.length) : null,
    min: numeric.length ? numeric.reduce((min, value) => Math.min(min, value)) : null,
    max: numeric.length ? numeric.reduce((max, value) => Math.max(max, value)) : null,
  };
  const result = { summary, header, hasHeader, groups: [] };
  if (groupIndex !== undefined) {
    const groups = new Map();
    for (const row of body) {
      const key = row[groupIndex] ? String(displayValue(row[groupIndex])) || t('(空)') : t('(空)');
      const value = numberOf(row[valueIndex]);
      if (value === undefined) continue;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(value);
    }
    const agg = request.agg || 'sum';
    result.groups = [...groups.entries()].map(([key, values]) => {
      const sum = values.reduce((acc, value) => acc + value, 0);
      return {
        key,
        [agg]: agg === 'average' ? round(sum / values.length)
          : agg === 'count' ? values.length
          : agg === 'max' ? values.reduce((max, value) => Math.max(max, value))
          : agg === 'min' ? values.reduce((min, value) => Math.min(min, value))
          : round(sum),
        count: values.length,
      };
    }).sort((a, b) => (Number(b[agg]) || 0) - (Number(a[agg]) || 0));
    result.groupBy = { column: header[groupIndex], valueColumn: header[valueIndex], agg };
  }
  return result;
}

/** @param {string} range @returns {import('xlsx').Range} */
export function decodeRange(range) {
  const text = String(range).trim().toUpperCase();
  if (!/^\$?[A-Z]{1,3}\$?[1-9]\d*(:\$?[A-Z]{1,3}\$?[1-9]\d*)?$/.test(text)) {
    throw httpError(400, 'bad-range', t('范围格式不正确：{0}（应为 A1:F40 这样的写法）', range));
  }
  return XLSX.utils.decode_range(text.replaceAll('$', ''));
}

/** @param {number} value */
function round(value) {
  return Math.round(value * 1e6) / 1e6;
}

/** @param {any} sheet */
export function sheetBounds(sheet) {
  let maxRow = 0;
  let maxCol = 0;
  for (const entry of sheet.celldata || []) {
    maxRow = Math.max(maxRow, entry.r);
    maxCol = Math.max(maxCol, entry.c);
  }
  return { s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } };
}

/**
 * Find a column by its header text or its letter. The index is relative to the selected range.
 * @param {string|undefined} name
 * @param {string[]} header
 * @param {{ s: { c: number }, e: { c: number } }} bounds
 */
function columnIndex(name, header, bounds) {
  if (!name) return undefined;
  const index = header.findIndex((value) => value === name);
  if (index >= 0) return index;
  if (/^[A-Za-z]{1,3}$/.test(name)) {
    const column = XLSX.utils.decode_col(name.toUpperCase());
    if (column >= bounds.s.c && column <= bounds.e.c) return column - bounds.s.c;
  }
  return undefined;
}

/**
 * Build a new workbook sheet containing stats and optional chart image.
 * @param {any} sheet
 * @param {any} stats
 * @param {{ chartSvg?: string, chartPng?: string, title?: string }} [options]
 */
export function statsSheet(sheet, stats, options = {}) {
  const rows = [
    [t('指标'), t('值')],
    [t('合计'), stats.summary.sum],
    [t('平均'), stats.summary.average],
    [t('最大'), stats.summary.max],
    [t('最小'), stats.summary.min],
    [t('计数'), stats.summary.count],
    [t('去重计数'), stats.summary.distinct],
  ];
  if (stats.groups?.length) {
    rows.push([], [t('分组'), stats.groupBy?.agg || t('汇总')]);
    for (const group of stats.groups) rows.push([group.key, group[stats.groupBy?.agg || 'sum'] ?? group.sum]);
  }
  const celldata = rows.flatMap((row, r) => row.map((value, c) => ({
    r,
    c,
    v: {
      v: value,
      ct: { fa: typeof value === 'number' ? '#,##0.00' : 'General', t: typeof value === 'number' ? 'n' : 's' },
      ...(r === 0 ? { bl: 1, bg: '#E8F5EE' } : {}),
    },
  })));
  return finishSheet({
    name: options.title || t('统计结果'),
    id: `stats-${token().slice(0, 8)}`,
    order: 999,
    row: Math.max(rows.length + 40, 120),
    column: 20,
    celldata,
    config: { columnlen: { 0: 180, 1: 180 } },
    status: 0,
    showGridLines: true,
  });
}

export { XLSX };
