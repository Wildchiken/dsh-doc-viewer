/**
 * Document-level operations shared by the HTTP panel and the agent tools.
 * A file is only ever replaced by a save the user or an agent asked for, after
 * its content is kept as a version; workbooks are saved by writing the changes
 * into the original, never by rebuilding it.
 */
import { copyFileSync, existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { t } from './i18n.js';
import { tmpdir } from 'node:os';
import { basename, dirname, extname, join } from 'node:path';
import { Document, Packer, PageBreak, Paragraph, TextRun } from 'docx';
import * as XLSX from '../../vendor/xlsx/xlsx.mjs';
import ExcelJS from 'exceljs';
import { convertOfficeToPdf, convertLegacyOffice, recalculateWorkbook } from './libreoffice.js';
import { textToDocx } from './print.js';
import { backupVersion, withFileLock, writeAtomically } from './versions.js';
import { featuresLostOnSave, hasFormulas, patchXlsx, withFormulaResults } from './xlsx-patch.js';
import { decodeText, encodeText, textFormatOf } from './encoding.js';
import { liteEvaluator } from './formula-lite.js';
import {
  applyEdits,
  buildXlsx,
  cellKey,
  computeStats,
  csvFromSheet,
  decodeRange,
  displayValue,
  loadSpreadsheet,
  sheetBounds,
  withCelldata,
  workbookSheetNames,
  statsSheet,
  writeSpreadsheet,
} from './spreadsheet.js';
import { formulaProblem } from './spreadsheet.js';
import {
  deletePdfPages,
  docxRawText,
  docxToHtml,
  docxToTxt,
  extractPdf,
  inspectPdf,
  mergePdfs,
  ocrDocument,
  parsePageList,
  pdfTables,
  pdfToDocx,
  pdfToTxt,
  pdfToXlsx,
  readTextDocument,
  splitPdf,
} from './pdf.js';
import {
  ensureDir,
  extensionOf,
  fileExists,
  httpError,
  MAX_TEXT_PREVIEW_CHARS,
  exportDirFor,
  sniffFile,
  uniquePath,
} from './util.js';

const SPREADSHEET_EXTENSIONS = new Set(['.xlsx', '.xls', '.csv', '.tsv', '.et']);
const WORD_EXTENSIONS = new Set(['.docx', '.doc', '.wps']);
const DEFAULT_SHEET_ROWS = 40;
const DEFAULT_SHEET_COLUMNS = 26;
const MAX_READ_ROWS = 500;
const MAX_READ_COLUMNS = 100;
const DEFAULT_TEXT_LIMIT = 20_000;

/** @param {string} path */
function assertExists(path) {
  if (!fileExists(path)) throw httpError(404, 'not-found', t('文件不存在：{0}', path));
}

/**
 * Detect the document family from bytes, not just the suffix. Files that are
 * none of these are refused rather than shown as garbage.
 * @param {string} path
 * @returns {'spreadsheet'|'pdf'|'word'|'text'}
 */
export function documentKind(path) {
  const ext = extensionOf(path);
  const { container, zipKind } = sniffFile(path);
  if (container === 'pdf') return 'pdf';
  if (container === 'zip') {
    if (zipKind === 'xlsx') return 'spreadsheet';
    if (zipKind === 'docx') return 'word';
    throw unsupported(path, t('压缩包或暂不支持的 Office 文档'));
  }
  if (container === 'ole') {
    if (WORD_EXTENSIONS.has(ext)) return 'word';
    if (ext === '.xls' || ext === '.et') return 'spreadsheet';
    throw unsupported(path, t('旧版 Office 二进制文档'));
  }
  if (container === 'binary') throw unsupported(path, t('二进制文件'));
  if (ext === '.pdf') throw httpError(422, 'corrupt-pdf', t('PDF 文件可能已损坏，已保留原文件。'));
  if (SPREADSHEET_EXTENSIONS.has(ext)) return 'spreadsheet';
  return 'text';
}

/** @param {string} path @param {string} what */
function unsupported(path, what) {
  return httpError(415, 'unsupported', t('暂不支持打开“{0}”（{1}）。可打开的类型：xlsx、xls、csv、tsv、et、pdf、docx、doc、wps 和文本文件。', basename(path), what));
}

/* ------------------------------------------------------------------ */
/* Parse cache                                                         */
/* ------------------------------------------------------------------ */

const CACHE_ENTRIES = 4;
const CACHE_MAX_CELLS = 1_500_000;
/** @type {Map<string, Promise<any>>} */
const workbookCache = new Map();

/**
 * Parsing and recalculating with LibreOffice takes seconds. An agent reading a
 * workbook in several steps must not pay that every time, so results are kept
 * until the file changes. Callers must treat the returned model as read-only.
 * @param {string} path
 * @param {{ recalc?: boolean }} [options]
 */
function loadWorkbook(path, options = {}) {
  assertExists(path);
  const info = statSync(path);
  const key = `${path}|${info.mtimeMs}|${info.size}|${String(options.recalc)}`;
  const cached = workbookCache.get(key);
  if (cached) {
    workbookCache.delete(key);
    workbookCache.set(key, cached);
    return cached;
  }
  const loading = loadSpreadsheet(path, options);
  workbookCache.set(key, loading);
  loading.then((doc) => {
    if (doc.stats.cells > CACHE_MAX_CELLS) workbookCache.delete(key);
  }, () => workbookCache.delete(key));
  while (workbookCache.size > CACHE_ENTRIES) workbookCache.delete(workbookCache.keys().next().value);
  return loading;
}

/* ------------------------------------------------------------------ */
/* Open                                                                */
/* ------------------------------------------------------------------ */

/**
 * Load a descriptor used by the panel. Spreadsheet sheets are returned in
 * FortuneSheet's model; PDFs return page metadata; Word returns HTML.
 * @param {string} path
 * @param {{ recalc?: boolean }} [options]
 */
export async function openDocument(path, options = {}) {
  assertExists(path);
  const kind = documentKind(path);
  if (kind === 'spreadsheet') return loadWorkbook(path, options);
  if (kind === 'pdf') {
    const info = await inspectPdf(path);
    return { ...info, warnings: [], stats: { pages: info.pageCount } };
  }
  if (kind === 'word') {
    const normalized = await normalizeWord(path);
    try {
      const [html, text] = await Promise.all([docxToHtml(normalized.path), docxRawText(normalized.path)]);
      const warnings = [];
      if (normalized.converted) warnings.push({ code: 'legacy-word', message: t('这是旧版 Word/WPS 文件，已在本机转换后预览，版式可能与原文件略有差异。') });
      return {
        kind: 'docx',
        path,
        name: basename(path),
        ext: extensionOf(path),
        format: normalized.converted ? 'converted-docx' : 'docx',
        html: html.html,
        messages: html.messages,
        warnings,
        stats: { chars: text.length },
      };
    } finally {
      normalized.cleanup();
    }
  }
  const { text, encoding } = readTextDocument(path);
  const warnings = encoding !== 'utf-8' ? [{ code: 'encoding', message: t('文本已按 {0} 解码。', encoding.toUpperCase()) }] : [];
  const truncated = text.length > MAX_TEXT_PREVIEW_CHARS;
  if (truncated) warnings.push({ code: 'preview-truncated', message: t('文件较大，这里只预览前 {0} MB 字符。', (MAX_TEXT_PREVIEW_CHARS / 1024 / 1024).toFixed(0)) });
  return {
    kind: 'text',
    path,
    name: basename(path),
    ext: extensionOf(path),
    format: 'text',
    text: truncated ? text.slice(0, MAX_TEXT_PREVIEW_CHARS) : text,
    truncated,
    encoding,
    warnings,
    stats: { chars: text.length, lines: countLines(text) },
  };
}

/** @param {string} text */
function countLines(text) {
  if (!text) return 0;
  let lines = 1;
  for (let index = text.indexOf('\n'); index !== -1; index = text.indexOf('\n', index + 1)) lines += 1;
  return lines;
}

/**
 * Normalize DOC/DOCX/WPS to a DOCX file usable by mammoth and LibreOffice. A
 * `.doc` that is really OOXML needs no conversion; the bytes decide.
 * @param {string} path
 * @returns {Promise<{ path: string, converted: boolean, cleanup: () => void }>}
 */
async function normalizeWord(path) {
  if (sniffFile(path).container === 'zip') return { path, converted: false, cleanup: () => {} };
  const dir = mkdtempSync(join(tmpdir(), 'dsh-doc-word-'));
  try {
    const result = await convertLegacyOffice(path, 'docx');
    const output = join(dir, `${basename(path, extname(path))}.docx`);
    writeFileSync(output, result.bytes);
    result.cleanup?.();
    return { path: output, converted: true, cleanup: () => rmSync(dir, { recursive: true, force: true }) };
  } catch (error) {
    rmSync(dir, { recursive: true, force: true });
    throw httpError(422, 'word-convert-failed', t('无法读取这个 Word/WPS 文件：{0}', error instanceof Error ? error.message : String(error)), error);
  }
}

/* ------------------------------------------------------------------ */
/* Save and convert                                                    */
/* ------------------------------------------------------------------ */

/**
 * Save a new copy of a spreadsheet model.
 * @param {any} ctx
 * @param {{ sheets: any[], sourcePath?: string, outputDirectory: string, format: string, title?: string, options?: any }} request
 */
export async function saveSpreadsheetCopy(ctx, request) {
  if (request.outputPath) return saveSheetsTo(ctx, request.sheets, request.outputPath, request);
  const format = String(request.format || 'xlsx').toLowerCase();
  const outputDirectory = ensureDir(request.outputDirectory ?? exportDirFor(request.sourcePath, request.workspaceRoot));
  const sourceBase = request.sourcePath ? basename(request.sourcePath, extname(request.sourcePath)) : t('工作簿');
  const base = request.title || t('{0}-已编辑', sourceBase);
  if (format === 'pdf') {
    const output = uniquePath(outputDirectory, base, '.pdf');
    const result = await sheetsToPdf(ctx, request.sheets, output, sourceBase);
    return { output, format, size: statSync(output).size, missingFonts: result.missingFonts || [], warnings: fontWarnings(result.missingFonts) };
  }
  if (!['xlsx', 'xls', 'csv'].includes(format)) throw httpError(400, 'bad-format', t('不支持的保存格式：{0}', format));
  const output = uniquePath(outputDirectory, base, `.${format}`);
  await writeSpreadsheet(request.sheets, output, format, { sheetName: request.options?.sheetName });
  const warnings = format === 'csv' && request.sheets.length > 1
    ? [{ code: 'csv-one-sheet', message: t('CSV 只能保存一个工作表，已保存“{0}”。', request.options?.sheetName || request.sheets[0]?.name) }]
    : [];
  if (format === 'xlsx') writeFileSync(output, await finishWorkbook(readFileSync(output), warnings));
  return { output, format, size: statSync(output).size, warnings };
}

/** Formats a sheet can be written to, by extension. */
const SHEET_FORMATS = { '.xlsx': 'xlsx', '.xls': 'xls', '.csv': 'csv', '.tsv': 'tsv', '.pdf': 'pdf' };

/**
 * Write sheets to a path the user chose. A file already there is kept as a
 * version first: the save panel asked before replacing it, but it can still be undone.
 * @param {any} ctx
 * @param {any[]} sheets
 * @param {string} target
 * @param {{ sourcePath?: string, options?: any }} request
 */
function saveSheetsTo(ctx, sheets, target, request = {}) {
  return withFileLock(target, () => saveSheetsToLocked(ctx, sheets, target, request));
}

/** @param {any} ctx @param {any[]} sheets @param {string} target @param {{ sourcePath?: string, options?: any, structure?: any[] }} request */
async function saveSheetsToLocked(ctx, sheets, target, request) {
  const format = SHEET_FORMATS[extname(target).toLowerCase()];
  if (!format) throw httpError(400, 'bad-format', t('不能存成 {0}。可用的格式：xlsx、xls、csv、tsv、pdf。', extname(target) || t('没有扩展名的文件')));
  const work = mkdtempSync(join(tmpdir(), 'dsh-doc-save-'));
  try {
    const temporary = join(work, `out${extname(target)}`);
    let warnings = [];
    if (format === 'pdf') {
      const result = await sheetsToPdf(ctx, sheets, temporary, basename(request.sourcePath || target, extname(request.sourcePath || target)));
      warnings = fontWarnings(result.missingFonts);
    } else if (format === 'tsv') {
      const complete = sheets.map(withCelldata);
      writeFileSync(temporary, csvFromSheet(complete.find((entry) => entry.name === request.options?.sheetName) || complete[0], 'utf8', '\t'));
    } else if (format === 'xlsx' && isWorkbookFile(request.sourcePath)) {
      writeFileSync(temporary, await patchXlsx(readFileSync(request.sourcePath), sheets, { structure: request.structure }));
    } else {
      await writeSpreadsheet(sheets, temporary, format, { sheetName: request.options?.sheetName });
    }
    if ((format === 'csv' || format === 'tsv') && sheets.length > 1) {
      warnings.push({ code: 'csv-one-sheet', message: t('{0} 只能保存一个工作表，已保存“{1}”。', format.toUpperCase(), request.options?.sheetName || sheets[0]?.name) });
    }
    if (format === 'xlsx') writeFileSync(temporary, await finishWorkbook(readFileSync(temporary), warnings));
    const versionId = backupVersion(target, '被另存为替换');
    writeAtomically(target, readFileSync(temporary));
    return { output: target, format, size: statSync(target).size, versionId, mtime: statSync(target).mtimeMs, warnings };
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

/** An existing .xlsx a save can be written into. @param {string|undefined} path */
function isWorkbookFile(path) {
  return Boolean(path) && extensionOf(path) === '.xlsx' && fileExists(path) && sniffFile(path).zipKind === 'xlsx';
}

/**
 * Why a workbook must not be overwritten with what this plugin writes, or
 * undefined when it can be.
 * @param {string} path
 * @param {{ sheetCount?: number, openedMtime?: number, format?: string }} options
 */
export async function inPlaceBlocker(path, options = {}) {
  const ext = extensionOf(path);
  if (options.format && SHEET_FORMATS[ext] !== options.format) return t('原文件是 {0}，要存成 {1} 需要另存', ext, options.format);
  if (ext === '.xls' || ext === '.et') return t('原文件是旧版 Excel/WPS 格式，直接写回会丢失部分格式');
  if (ext === '.xlsm') return t('原文件含有宏，直接写回会丢失宏');
  if (!['.xlsx', '.csv', '.tsv'].includes(ext)) return t('不能直接写回 {0} 文件', ext || t('这种'));
  if ((ext === '.csv' || ext === '.tsv') && (options.sheetCount || 1) > 1) return t('{0} 只能保存一个工作表，现在有 {1} 个', ext.slice(1).toUpperCase(), options.sheetCount);
  if (options.openedMtime && Math.abs(statSync(path).mtimeMs - options.openedMtime) > 1) return t('文件在打开之后被其他程序修改过，直接写回会覆盖那些修改');
  if (ext === '.xlsx') {
    const lost = featuresLostOnSave(readFileSync(path));
    if (lost.length) return t('原文件里有{0}，直接写回会丢失', lost.join(t('、')));
  }
  return undefined;
}

/** How long a save or a live calculation waits for LibreOffice before going without. */
const SAVE_CALCULATION_MS = 20_000;
const NEW_CALCULATION_MS = 6_000;

/**
 * An .xlsx with the right result in every formula, calculated by LibreOffice, the
 * same engine that answers when a file has none. Without it the results are left
 * out, for whatever opens the file next to calculate.
 * @param {Buffer} bytes
 * @param {any[]} [warnings] told when the results could not be worked out
 * @returns {Promise<Buffer>}
 */
async function finishWorkbook(bytes, warnings, timeoutMs = SAVE_CALCULATION_MS) {
  if (!hasFormulas(bytes)) return bytes;
  let resultOf;
  try {
    // A save must not wait long on the calculation; without it the file is still right.
    const calculated = await calculatedCells(bytes, timeoutMs);
    resultOf = (sheet, address) => {
      const cell = calculated.Sheets[sheet]?.[address];
      if (!cell || cell.v === undefined) return undefined;
      return cell.t === 'e' ? { error: String(cell.w ?? cell.v) } : cell.v;
    };
  } catch (error) {
    // Simple formulas are worked out here, so the file and whoever reads it still have their results.
    const lite = await liteResults(bytes);
    resultOf = lite.resultOf;
    const reason = error instanceof Error ? error.message : String(error);
    warnings?.push(lite.missing === 0
      ? { code: 'results-lite', message: t('LibreOffice 没能计算（{0}），公式结果已用内置计算补上。', reason) }
      : { code: 'results-missing', message: t('公式结果没能全部算出来（{0}）：{1} 个已用内置计算补上，{2} 个用到了内置计算不支持的写法，文件里只存了公式，用 Excel 打开时会重新计算。', reason, lite.filled, lite.missing) });
  }
  return withFormulaResults(bytes, resultOf);
}

/**
 * Results for the formulas the lite evaluator can work out.
 * @param {Buffer} bytes an xlsx
 */
async function liteResults(bytes) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  const sheets = new Map(workbook.worksheets.map((worksheet) => [worksheet.name, worksheet]));
  const cellAt = (sheet, address) => {
    const cell = sheets.get(sheet)?.findCell(address);
    if (!cell) return undefined;
    return cell.type === ExcelJS.ValueType.Formula ? { formula: cell.formula } : { value: cell.value };
  };
  const resultOf = liteEvaluator(cellAt);
  let filled = 0;
  let missing = 0;
  for (const worksheet of workbook.worksheets) {
    worksheet.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.type !== ExcelJS.ValueType.Formula) return;
      if (resultOf(worksheet.name, cell.address) === undefined) missing += 1;
      else filled += 1;
    }));
  }
  return { resultOf, filled, missing };
}

/**
 * Calculate a workbook held in memory.
 * @param {Buffer} bytes an .xlsx
 * @returns {Promise<any>} the calculated workbook, as SheetJS reads it
 */
async function calculatedCells(bytes, timeoutMs) {
  const work = mkdtempSync(join(tmpdir(), 'dsh-doc-calc-'));
  try {
    const source = join(work, 'book.xlsx');
    writeFileSync(source, bytes);
    const recalc = await recalculateWorkbook(source, { timeoutMs });
    try {
      return XLSX.read(readFileSync(recalc.path), { type: 'buffer', cellDates: false });
    } finally {
      recalc.cleanup?.();
    }
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

/**
 * The results of the formulas in what the grid holds now, for functions its own
 * engine does not have.
 * @param {{ path: string, sheets: any[], structure?: any[] }} request
 * @returns {Promise<{ id: string, r: number, c: number, v: unknown, m: string }[]>}
 */
export async function calculateSheets(request) {
  const complete = (request.sheets || []).map(withCelldata);
  const bytes = isWorkbookFile(request.path)
    ? await patchXlsx(readFileSync(request.path), complete, { structure: request.structure })
    : await buildXlsx(complete);
  const calculated = await calculatedCells(bytes, SAVE_CALCULATION_MS);
  const names = workbookSheetNames(complete);
  const results = [];
  complete.forEach((sheet, index) => {
    const values = calculated.Sheets[names[index]];
    if (!values) return;
    for (const entry of sheet.celldata) {
      if (!entry?.v?.f) continue;
      const cell = values[XLSX.utils.encode_cell({ r: entry.r, c: entry.c })];
      if (!cell || cell.v === undefined) continue;
      const isError = cell.t === 'e';
      const v = isError ? String(cell.w ?? cell.v) : cell.v;
      results.push({ id: sheet.id, r: entry.r, c: entry.c, v, m: isError ? v : String(displayValue({ v, ct: entry.v.ct })) });
    }
  });
  return results;
}

/**
 * What a sheet's new content is, in the file's own format.
 * @param {string} path
 * @param {any[]} complete
 * @param {any[]} [structure]
 * @returns {Promise<{ content: Buffer|Uint8Array, warnings: any[] }>}
 */
async function inPlaceContent(path, complete, structure) {
  const ext = extensionOf(path);
  const original = readFileSync(path);
  const warnings = [];
  if (ext === '.xlsx') {
    return { content: await finishWorkbook(await patchXlsx(original, complete, { structure }), warnings), warnings };
  }
  // A CSV keeps its encoding, byte order mark and line endings.
  const encoded = encodeText(csvFromSheet(complete[0], 'plain', ext === '.tsv' ? '\t' : ','), textFormatOf(original));
  if (encoded.changed) warnings.push({ code: 'encoding', message: encoded.changed });
  return { content: encoded.bytes, warnings };
}

/**
 * Save sheets over the file they came from, keeping what was there as a version.
 * @param {{ path: string, sheets: any[], openedMtime?: number, reason?: string, structure?: any[] }} request
 * @returns {Promise<{ output: string, inPlace: true, versionId?: string, mtime: number } | { needsLocation: true, reason: string, directory: string, suggestedName: string }>}
 */
export function saveInPlace(request) {
  return withFileLock(request.path, async () => {
    const { path, sheets } = request;
    assertExists(path);
    const blocker = await inPlaceBlocker(path, { sheetCount: sheets?.length, openedMtime: request.openedMtime });
    if (blocker) {
      const ext = extensionOf(path);
      return { needsLocation: true, reason: blocker, directory: dirname(path), suggestedName: `${basename(path, ext)}${ext === '.xlsx' ? t('-编辑') : ''}.xlsx` };
    }
    const complete = (sheets || []).map(withCelldata);
    if (complete.length === 0) throw httpError(400, 'no-sheet', t('没有收到任何工作表，已取消保存，原文件未改动。'));
    const { content, warnings } = await inPlaceContent(path, complete, request.structure);
    const versionId = backupVersion(path, request.reason || '保存');
    writeAtomically(path, content);
    return { output: path, inPlace: true, versionId, mtime: statSync(path).mtimeMs, warnings };
  });
}

/**
 * Write a text document back to its own file, in the encoding and line endings it already
 * has, with the previous content kept as a version. A file changed since it was opened is
 * left alone unless no `openedMtime` is given, which means the user chose to overwrite it.
 * @param {{ path: string, text: string, openedMtime?: number }} request
 */
export function saveTextInPlace({ path, text, openedMtime }) {
  return withFileLock(path, async () => {
    assertExists(path);
    if (documentKind(path) !== 'text') throw httpError(415, 'wrong-kind', t('“{0}”不是文本文件，不能这样保存。', basename(path)));
    if (typeof text !== 'string') throw httpError(400, 'bad-text', t('没有收到文本内容，原文件未改动。'));
    if (openedMtime !== undefined && statSync(path).mtimeMs !== openedMtime) return { conflict: true };
    const original = readFileSync(path);
    // The page only ever had the start of a file this long; saving would cut off the rest.
    if (decodeText(original).text.length > MAX_TEXT_PREVIEW_CHARS) throw httpError(413, 'too-long', t('文件太大，这里只显示了开头一部分，不能直接保存，原文件未改动。'));
    const { bytes, changed } = encodeText(text, textFormatOf(original));
    const versionId = backupVersion(path, '保存');
    writeAtomically(path, bytes);
    return { output: path, inPlace: true, versionId, mtime: statSync(path).mtimeMs, warnings: changed ? [{ code: 'encoding', message: changed }] : [] };
  });
}

/**
 * Move a finished file to where the user chose, keeping a file it replaces as a version.
 * @param {string} from
 * @param {string} target
 */
function deliver(from, target) {
  backupVersion(target, '被导出替换');
  copyFileSync(from, target);
  rmSync(from, { force: true });
}

/** @param {any} ctx @param {any[]} sheets @param {string} output */
async function sheetsToPdf(ctx, sheets, output, title) {
  const work = mkdtempSync(join(tmpdir(), 'dsh-doc-pdf-'));
  try {
    const xlsx = join(work, 'source.xlsx');
    writeFileSync(xlsx, await buildXlsx(sheets, { print: { title } }));
    return await convertOfficeToPdf(ctx, xlsx, output, 'xlsx');
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

/** @param {string[]|undefined} fonts */
function fontWarnings(fonts) {
  return fonts?.length
    ? [{ code: 'fonts', message: t('以下字体本机没有，PDF 中已用替代字体：{0}。', fonts.slice(0, 8).join(t('、'))) }]
    : [];
}

/** What a new document may be. */
const CREATABLE = new Set(['.xlsx', '.csv', '.md', '.markdown', '.txt']);

/**
 * Make a new document. A workbook can start empty or with rows, whose strings that
 * begin with = are formulas; Markdown and text start with `content`. Nothing that
 * exists already is replaced.
 * @param {string} path
 * @param {{ rows?: unknown[][], header?: boolean, sheet?: string, content?: string }} [options]
 */
export async function createDocument(path, options = {}) {
  const ext = extensionOf(path);
  if (!CREATABLE.has(ext)) throw httpError(400, 'bad-format', t('只能新建 xlsx、csv、md 或 txt 文件，不能新建 {0}。', ext || t('没有扩展名的文件')));
  if (existsSync(path)) throw httpError(409, 'exists', t('{0} 已经存在，没有覆盖它。换一个名字，或者用 doc_viewer_edit 修改它。', basename(path)));
  ensureDir(dirname(path));
  const rows = Array.isArray(options.rows) ? options.rows.filter(Array.isArray) : [];
  if (ext === '.xlsx') {
    const sheetName = String(options.sheet || t('工作表1')).replace(/[[\]:*?/\\]/g, '_').slice(0, 31) || t('工作表1');
    rows.forEach((row, r) => row.forEach((value, c) => {
      if (typeof value === 'string' && value.trim().startsWith('=')) {
        const problem = formulaProblem(value, { sheet: sheetName, r, c });
        if (problem) throw httpError(400, 'bad-formula', t('{0} 的公式 {1} 有问题：{2}。没有新建文件。', XLSX.utils.encode_cell({ r, c }), value, problem));
      }
    }));
    const book = new ExcelJS.Workbook();
    book.creator = t('DeepSeek Harness 文档');
    const sheet = book.addWorksheet(sheetName, rows.length && options.header !== false ? { views: [{ state: 'frozen', ySplit: 1 }] } : undefined);
    rows.forEach((row, r) => {
      row.forEach((value, c) => {
        const cell = sheet.getCell(r + 1, c + 1);
        cell.value = typeof value === 'string' && value.trim().startsWith('=') ? { formula: value.trim().slice(1) } : value ?? null;
        if (r === 0 && options.header !== false) cell.font = { bold: true };
      });
    });
    const warnings = [];
    // A new workbook is opened and calculated by the grid at once; the file's results matter less.
    writeFileSync(path, await finishWorkbook(Buffer.from(await book.xlsx.writeBuffer()), warnings, NEW_CALCULATION_MS));
    return { output: path, rows: rows.length, warnings };
  }
  if (ext === '.csv') {
    writeFileSync(path, rowsToCsv(rows), 'utf8');
    return { output: path, rows: rows.length, warnings: [] };
  }
  writeFileSync(path, String(options.content ?? ''), 'utf8');
  return { output: path, warnings: [] };
}

/** @param {any[][]} rows */
function rowsToCsv(rows) {
  const escape = (value) => {
    const text = value === undefined || value === null ? '' : String(value);
    return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  };
  return `﻿${rows.map((row) => row.map(escape).join(',')).join('\r\n')}`;
}

/**
 * Convert one document to a target format.
 * @param {object|null} ctx host context (used to reuse officeToPdf)
 * @param {string} path
 * @param {string} target
 * @param {{ outputDirectory?: string, workspaceRoot?: string, outputPath?: string }} request without a folder, outputs go
 *   beside each source; `outputPath` is a place the user chose in the save panel
 */
export async function convertOne(ctx, path, target, request) {
  assertExists(path);
  if (request.outputPath) {
    const work = mkdtempSync(join(tmpdir(), 'dsh-doc-export-'));
    try {
      const result = await convertOne(ctx, path, target, { ...request, outputPath: undefined, outputDirectory: work });
      deliver(result.output, request.outputPath);
      return { ...result, output: request.outputPath, size: statSync(request.outputPath).size };
    } finally {
      rmSync(work, { recursive: true, force: true });
    }
  }
  const outputDirectory = ensureDir(request.outputDirectory ?? exportDirFor(path, request.workspaceRoot));
  const format = String(target || '').toLowerCase();
  const base = basename(path, extname(path));
  const ext = extensionOf(path);
  const kind = documentKind(path);
  const warnings = [];
  let output;

  if (kind === 'spreadsheet') {
    if (!['xlsx', 'xls', 'csv', 'pdf'].includes(format)) throw httpError(400, 'unsupported-conversion', t('表格不能转换为 {0}', format));
    const doc = await loadWorkbook(path);
    output = uniquePath(outputDirectory, t('{0}-转换', base), `.${format}`);
    if (format === 'pdf') {
      const result = await sheetsToPdf(ctx, doc.sheets, output, base);
      warnings.push(...fontWarnings(result.missingFonts));
    } else if (format === 'csv') {
      writeFileSync(output, csvFromSheet(doc.sheets[0], 'utf8'), 'utf8');
      if (doc.sheets.length > 1) warnings.push({ code: 'csv-one-sheet', message: t('CSV 只能保存一个工作表，已转换“{0}”，其余 {1} 个工作表未包含。', doc.sheets[0].name, doc.sheets.length - 1) });
    } else {
      await writeSpreadsheet(doc.sheets, output, format);
    }
  } else if (kind === 'pdf') {
    if (format === 'txt') {
      output = uniquePath(outputDirectory, t('{0}-提取', base), '.txt');
      await pdfToTxt(path, output);
    } else if (format === 'docx') {
      output = uniquePath(outputDirectory, t('{0}-提取', base), '.docx');
      warnings.push({ code: 'lossy', message: t('PDF 转 Word 是有损转换：复杂排版、图片、表格和分栏可能丢失，只保留可提取文字。') });
      await pdfToDocx(path, output);
    } else if (format === 'xlsx') {
      output = uniquePath(outputDirectory, t('{0}-提取', base), '.xlsx');
      warnings.push({ code: 'lossy', message: t('PDF 转 XLSX 是有损转换：按文字坐标尽力还原表格，复杂排版会丢失。') });
      const result = await pdfToXlsx(path, output);
      if (result.warning) warnings.push({ code: 'ocr-needed', message: result.warning });
    } else if (format === 'csv') {
      output = uniquePath(outputDirectory, t('{0}-提取', base), '.csv');
      warnings.push({ code: 'lossy', message: t('PDF 表格按文字位置还原为 CSV，复杂排版会丢失。') });
      writeFileSync(output, rowsToCsv(pdfTables(await extractPdf(path))), 'utf8');
    } else if (format === 'pdf') {
      throw httpError(400, 'same-format', t('源文件已经是 PDF；如需合并、拆分或删页，请使用 PDF 工具。'));
    } else {
      throw httpError(400, 'unsupported-conversion', t('PDF 不能直接转换为 {0}', format));
    }
  } else if (kind === 'word') {
    if (format !== 'pdf' && format !== 'txt') throw httpError(400, 'unsupported-conversion', t('Word 不能直接转换为 {0}', format));
    const normalized = await normalizeWord(path);
    try {
      if (format === 'pdf') {
        output = uniquePath(outputDirectory, t('{0}-转换', base), '.pdf');
        const result = await convertOfficeToPdf(ctx, normalized.path, output, 'docx');
        warnings.push(...fontWarnings(result.missingFonts));
      } else {
        output = uniquePath(outputDirectory, t('{0}-提取', base), '.txt');
        await docxToTxt(normalized.path, output);
      }
    } finally {
      normalized.cleanup();
    }
  } else if (format === 'pdf') {
    output = uniquePath(outputDirectory, t('{0}-转换', base), '.pdf');
    const work = mkdtempSync(join(tmpdir(), 'dsh-doc-txt-'));
    try {
      const docx = join(work, 'text.docx');
      const { text } = readTextDocument(path);
      writeFileSync(docx, await textToDocx(text, {
        title: basename(path),
        markdown: ['.md', '.markdown'].includes(ext),
        monospace: ['.log', '.json', '.xml', '.yaml', '.yml', '.ini', '.html', '.htm'].includes(ext),
      }));
      const result = await convertOfficeToPdf(ctx, docx, output, 'docx');
      warnings.push(...fontWarnings(result.missingFonts));
    } finally {
      rmSync(work, { recursive: true, force: true });
    }
  } else if (format === 'txt') {
    output = uniquePath(outputDirectory, t('{0}-转换', base), '.txt');
    writeFileSync(output, `${readTextDocument(path).text}`, 'utf8');
  } else {
    throw httpError(400, 'unsupported-conversion', t('{0} 不能直接转换为 {1}', ext || t('文本'), format));
  }

  if (!existsSync(output)) throw httpError(500, 'conversion-failed', t('没有生成输出文件'));
  return {
    ok: true,
    source: path,
    output,
    target: format,
    size: statSync(output).size,
    warnings,
    lossy: warnings.some((warning) => warning.code === 'lossy'),
  };
}

/**
 * Batch conversion. Each file is independent; failures are returned per item.
 * @param {object|null} ctx
 * @param {string[]} paths
 * @param {string} target
 * @param {{ outputDirectory: string }} request
 */
export async function convertMany(ctx, paths, target, request) {
  const results = [];
  for (const path of paths || []) {
    try {
      results.push(await convertOne(ctx, path, target, request));
    } catch (error) {
      results.push({
        ok: false,
        source: path,
        target,
        code: typeof error?.code === 'string' ? error.code : 'failed',
        message: error instanceof Error ? error.message : String(error),
      });
    }
  }
  return { results, outputDirectory: request.outputDirectory };
}

/**
 * PDF toolbox: merge/split/delete/ocr.
 * @param {object|null} ctx
 * @param {string} action
 * @param {{ paths?: string[], pages?: string, ranges?: string, outputDirectory: string, format?: string }} request
 */
export async function pdfOperation(ctx, action, request) {
  const paths = request.paths || [];
  for (const path of paths) assertExists(path);
  if (paths[0] && !request.outputDirectory) request = { ...request, outputDirectory: exportDirFor(paths[0], request.workspaceRoot) };
  ensureDir(request.outputDirectory);
  if (action === 'merge') {
    if (request.outputPath) backupVersion(request.outputPath, '被合并结果替换');
    const output = request.outputPath || uniquePath(request.outputDirectory, t('{0}-合并', basename(paths[0] || t('合并文档'), extname(paths[0] || ''))), '.pdf');
    return mergePdfs(paths, output);
  }
  const source = paths[0];
  if (!source && ['split', 'delete', 'ocr'].includes(action)) throw httpError(400, 'bad-path', t('请选择一个文件'));
  if (action === 'split') {
    return splitPdf(source, request.outputDirectory, { ranges: request.ranges, title: basename(source, extname(source)) });
  }
  if (action === 'delete') {
    if (!String(request.pages || '').trim()) throw httpError(400, 'bad-pages', t('请指定要删除的页码，例如 2,4-6'));
    const output = uniquePath(request.outputDirectory, t('{0}-删页', basename(source, extname(source))), '.pdf');
    return deletePdfPages(source, output, request.pages);
  }
  if (action === 'ocr') {
    const result = await ocrDocument(source, { maxPages: 80, pages: request.pages });
    const format = request.format === 'docx' ? 'docx' : 'txt';
    const output = uniquePath(request.outputDirectory, `${basename(source, extname(source))}-OCR`, `.${format}`);
    if (format === 'docx') {
      const children = [];
      result.pages.forEach((page, index) => {
        if (index > 0) children.push(new Paragraph({ children: [new PageBreak()] }));
        for (const line of String(page.text || '').split(/\n+/).filter(Boolean)) {
          children.push(new Paragraph({ children: [new TextRun({ text: line, font: 'PingFang SC', size: 22 })] }));
        }
      });
      const doc = new Document({ title: basename(source), sections: [{ children }] });
      writeFileSync(output, await Packer.toBuffer(doc));
    } else {
      writeFileSync(output, `﻿${result.text}\n`, 'utf8');
    }
    return { output, pages: result.pages.length, engine: result.engine, chars: result.text.length };
  }
  throw httpError(400, 'bad-pdf-action', t('不支持的 PDF 操作：{0}', action));
}

/* ------------------------------------------------------------------ */
/* Read for agents                                                     */
/* ------------------------------------------------------------------ */

/**
 * Read a bounded part of a document for agent tools. `limit` is a budget in
 * characters; each kind stops at a natural boundary and says how to continue.
 * @param {string} path
 * @param {{ sheet?: string, range?: string, pages?: string, offset?: number, limit?: number }} request
 */
export async function readDocument(path, request = {}) {
  assertExists(path);
  const kind = documentKind(path);
  const limit = Math.max(200, Number(request.limit) || DEFAULT_TEXT_LIMIT);
  if (kind === 'spreadsheet') return readSheetRange(await loadWorkbook(path), request, limit);
  if (kind === 'pdf') return readPdfPages(path, request, limit);
  let text;
  let extra = {};
  if (kind === 'word') {
    const normalized = await normalizeWord(path);
    try {
      text = await docxRawText(normalized.path);
    } finally {
      normalized.cleanup();
    }
    extra = { kind: 'docx' };
  } else {
    const decoded = readTextDocument(path);
    text = decoded.text;
    extra = { kind: 'text', encoding: decoded.encoding };
  }
  const offset = Math.max(0, Math.floor(Number(request.offset) || 0));
  const end = Math.min(text.length, offset + limit);
  return {
    ...extra,
    text: text.slice(offset, end),
    offset,
    totalChars: text.length,
    nextOffset: end < text.length ? end : undefined,
    truncated: end < text.length,
  };
}

/** @param {any} doc @param {any} request @param {number} limit */
function readSheetRange(doc, request, limit) {
  const sheet = request.sheet ? doc.sheets.find((entry) => entry.name === request.sheet) : doc.sheets[0];
  if (!sheet) throw httpError(400, 'bad-sheet', t('没有名为“{0}”的工作表。现有工作表：{1}', request.sheet, doc.sheetNames.join(t('、'))));
  const used = sheetBounds(sheet);
  const bounds = request.range ? decodeRange(request.range) : {
    s: { r: 0, c: 0 },
    e: { r: Math.min(DEFAULT_SHEET_ROWS - 1, used.e.r), c: Math.min(DEFAULT_SHEET_COLUMNS - 1, used.e.c) },
  };
  const lookup = new Map(sheet.celldata.map((entry) => [cellKey(entry.r, entry.c), entry.v]));
  const lastColumn = Math.min(bounds.e.c, bounds.s.c + MAX_READ_COLUMNS - 1);
  const lastWanted = Math.min(bounds.e.r, used.e.r, bounds.s.r + MAX_READ_ROWS - 1);
  const rows = [];
  let chars = 0;
  let lastRow = bounds.s.r - 1;
  for (let r = bounds.s.r; r <= lastWanted; r += 1) {
    const row = [];
    for (let c = bounds.s.c; c <= lastColumn; c += 1) {
      const cell = lookup.get(cellKey(r, c));
      const hasValue = cell && (cell.v !== undefined || cell.f);
      const value = hasValue ? { value: cell.v, formula: cell.f, display: String(displayValue(cell)) } : null;
      row.push(value);
      chars += value ? value.display.length + (value.formula?.length || 0) + 3 : 3;
    }
    if (chars > limit && rows.length > 0) break;
    rows.push(row);
    lastRow = r;
  }
  const hasMoreRows = lastRow < Math.min(bounds.e.r, used.e.r);
  const shown = { s: bounds.s, e: { r: Math.max(lastRow, bounds.s.r), c: lastColumn } };
  return {
    kind: 'spreadsheet',
    sheets: doc.sheetNames,
    sheet: sheet.name,
    range: XLSX.utils.encode_range(shown),
    usedRange: XLSX.utils.encode_range(used),
    firstRow: bounds.s.r,
    firstColumn: bounds.s.c,
    rows,
    truncated: hasMoreRows || lastColumn < bounds.e.c,
    nextRange: hasMoreRows
      ? XLSX.utils.encode_range({ s: { r: lastRow + 1, c: bounds.s.c }, e: { r: request.range ? bounds.e.r : used.e.r, c: bounds.e.c } })
      : undefined,
    warnings: doc.warnings,
  };
}

/** @param {string} path @param {any} request @param {number} limit */
async function readPdfPages(path, request, limit) {
  const info = await inspectPdf(path);
  const wanted = request.pages ? parsePageList(request.pages, info.pageCount) : [1];
  const extraction = await extractPdf(path, { pages: wanted });
  const pages = [];
  let chars = 0;
  for (const page of extraction.pages) {
    const room = limit - chars;
    if (room <= 0) break;
    const text = page.text.slice(0, room);
    pages.push({ page: page.page, text, partial: text.length < page.text.length, empty: page.text.trim() === '' });
    chars += text.length;
  }
  const lastShown = pages[pages.length - 1];
  const remaining = wanted.filter((page) => page > (lastShown?.page ?? 0));
  return {
    kind: 'pdf',
    pageCount: info.pageCount,
    pages,
    truncated: remaining.length > 0 || Boolean(lastShown?.partial),
    nextPages: remaining.length ? remaining.join(',') : undefined,
    scanned: pages.length > 0 && pages.every((page) => page.empty),
  };
}

/* ------------------------------------------------------------------ */
/* Statistics and edits                                                */
/* ------------------------------------------------------------------ */

/**
 * Compute stats and optionally save the result as a new workbook.
 * @param {string} path
 * @param {any} request
 * @param {string} outputDirectory
 */
export async function statsForFile(path, request, outputDirectory) {
  const doc = await loadWorkbookForTable(path);
  const result = computeStats(doc.sheets, request);
  let saved;
  if (request.save) {
    const base = basename(path, extname(path));
    const merged = [...doc.sheets, statsSheet(doc.sheets[0], result, { title: t('统计结果') })];
    saved = uniquePath(outputDirectory, t('{0}-统计', base), '.xlsx');
    if (isWorkbookFile(path)) writeFileSync(saved, await patchXlsx(readFileSync(path), merged));
    else await writeSpreadsheet(merged, saved, 'xlsx');
  }
  return { ...result, sheetNames: doc.sheetNames, saved, warnings: doc.warnings };
}

/** @param {string} path */
async function loadWorkbookForTable(path) {
  assertExists(path);
  const kind = documentKind(path);
  if (kind !== 'spreadsheet') throw httpError(415, 'wrong-kind', t('“{0}”不是表格，不能做这个操作。', basename(path)));
  return loadWorkbook(path);
}

/**
 * Apply tool edits to a workbook and save a new copy.
 * @param {string} path
 * @param {{ sheet?: string, edits?: any[], convertTextNumbers?: boolean, saveAs?: string }} request
 * @param {string} outputDirectory
 */
export function editFile(path, request, outputDirectory) {
  return withFileLock(path, () => editFileLocked(path, request, outputDirectory));
}

/** @param {string} path @param {any} request @param {string} outputDirectory */
async function editFileLocked(path, request, outputDirectory) {
  const doc = await loadWorkbookForTable(path);
  // The parsed model may be shared through the cache; edits work on their own copy.
  const sheets = structuredClone(doc.sheets);
  const applied = applyEdits(sheets, request.edits || [], { convertTextNumbers: request.convertTextNumbers, sheet: request.sheet });
  if (applied.applied.length === 0 && applied.converted === 0) {
    throw httpError(400, 'nothing-to-do', request.convertTextNumbers ? t('没有可转换的文本数字，也没有要写入的单元格，未生成新文件。') : t('没有要写入的单元格，未生成新文件。'));
  }
  // What the edits change, for a page that has the file open with unsaved edits of its own.
  const liveChanges = request.live ? cellChanges(doc.sheets, sheets) : undefined;
  const live = Boolean(liveChanges && liveChanges.every((change) => insideGrid(doc.sheets, change)));
  // Results saved for the old values are wrong now.
  for (const sheet of sheets) for (const entry of sheet.celldata) if (entry.v?.f) { delete entry.v.v; delete entry.v.m; }
  // By default the file is changed where it is, with the old content kept as a version.
  const wanted = request.saveAs || SHEET_FORMATS[extensionOf(path)];
  const blocker = live ? undefined
    : request.inPlace === false ? t('要求另存')
    : request.live ? t('{0}，而这次改动{1}，没法直接放进页面', request.blockedReason || t('用户在文档页里对这个文件有未保存的修改'), liveChanges ? t('超出了表格当前的范围') : t('太多'))
    : request.blockedReason || await inPlaceBlocker(path, { sheetCount: sheets.length, format: wanted });
  const warnings = [...(doc.warnings || [])];
  const format = blocker ? (request.saveAs === 'xls' ? 'xls' : 'xlsx') : extensionOf(path).slice(1);
  const work = mkdtempSync(join(tmpdir(), 'dsh-doc-edit-'));
  try {
    // Everything is worked out and checked before anything the user owns changes.
    const trial = join(work, `trial.${format}`);
    let content;
    if (!blocker) {
      const built = await inPlaceContent(path, sheets);
      content = built.content;
      warnings.push(...built.warnings);
    } else if (format === 'xlsx') {
      content = await finishWorkbook(isWorkbookFile(path) ? await patchXlsx(readFileSync(path), sheets) : await buildXlsx(sheets), warnings);
    } else {
      await writeSpreadsheet(sheets, trial, format);
      content = readFileSync(trial);
    }
    writeFileSync(trial, content);
    let values;
    let written;
    try {
      // The results were just calculated into it; reading must not calculate again.
      written = (await loadSpreadsheet(trial, { recalc: false })).sheets;
      values = verifyFormulas(written, request.edits || [], request.sheet);
    } catch { /* value verification is optional */ }
    for (const item of (values || []).filter((value) => value.formula && /^#[A-Z/0!?]+/.test(value.display))) {
      if (item.display.startsWith('#NAME?')) {
        throw httpError(400, 'bad-formula', t('{0}!{1} 的公式 {2} 算出来是 #NAME?：里面有无法识别的函数或名称。文件没有改动。', item.sheet, item.address, item.formula));
      }
      warnings.push({ code: 'formula-error', message: t('{0}!{1} 的结果是 {2}，请检查公式和它引用的数据。', item.sheet, item.address, item.display) });
    }
    // What changed, results of dependent formulas included, so the agent need not read the file back to find out.
    const changes = written ? changeReport(cellChanges(doc.sheets, written), values) : undefined;
    if (live) {
      // Nothing is written: the page puts the edits among the user's own, and the next save writes them all.
      return { output: path, inPlace: false, live: true, cells: liveChanges, applied: applied.applied, converted: applied.converted, values, changes, warnings };
    }
    let output;
    let versionId;
    if (blocker) {
      output = uniquePath(outputDirectory, t('{0}-已编辑', basename(path, extname(path))), `.${format}`);
      writeFileSync(output, content);
    } else {
      output = path;
      versionId = backupVersion(path, 'Agent 修改');
      writeAtomically(path, content);
    }
    return {
      output,
      inPlace: !blocker,
      // Every cell whose content changed, results of dependent formulas included, so an open page can show the file as it now is.
      cells: blocker || !written ? undefined : cellChanges(doc.sheets, written),
      mtime: blocker ? undefined : statSync(path).mtimeMs,
      versionId,
      keptAside: blocker || undefined,
      format: extensionOf(output).slice(1),
      applied: applied.applied,
      converted: applied.converted,
      values,
      changes,
      warnings,
    };
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

/** Past this many changed cells a page reloads the file instead of being sent each one. */
const MAX_SYNCED_CELLS = 3000;

/** @param {any} cell @returns {string} what makes two cells show the same thing */
function cellIdentity(cell) {
  if (!cell) return '';
  return JSON.stringify([cell.v ?? null, cell.f ?? null, cell.m ?? null, Array.isArray(cell.ct?.s) ? cell.ct.s : null]);
}

/**
 * The cells that differ between two versions of a workbook, by sheet name.
 * @param {any[]} before
 * @param {any[]} after
 * @returns {{ sheet: string, r: number, c: number, address: string, before: any, after: any }[]|undefined} undefined past the limit
 */
export function cellChanges(before, after) {
  const changes = [];
  for (const sheet of after) {
    const old = before.find((entry) => entry.name === sheet.name);
    const was = new Map((old?.celldata || []).map((entry) => [`${entry.r},${entry.c}`, entry.v]));
    const now = new Map((sheet.celldata || []).map((entry) => [`${entry.r},${entry.c}`, entry.v]));
    for (const key of new Set([...was.keys(), ...now.keys()])) {
      const a = was.get(key);
      const b = now.get(key);
      if (cellIdentity(a) === cellIdentity(b)) continue;
      const [r, c] = key.split(',').map(Number);
      changes.push({ sheet: sheet.name, r, c, address: XLSX.utils.encode_cell({ r, c }), before: a ? structuredClone(a) : null, after: b ? structuredClone(b) : null });
      if (changes.length > MAX_SYNCED_CELLS) return undefined;
    }
  }
  return changes;
}

/**
 * Sort a workbook's changes into the cells an edit wrote and the formulas that moved because of it.
 * @param {any[]|undefined} changes from cellChanges; undefined when there were too many to list
 * @param {{ sheet: string, address: string }[]|undefined} values the cells the edit wrote
 */
export function changeReport(changes, values) {
  if (!changes) return undefined;
  const asked = new Set((values || []).map((item) => `${item.sheet}!${item.address}`));
  const shown = (cell) => (cell ? String(displayValue(cell)) : '');
  const lines = changes
    .slice()
    .sort((a, b) => a.sheet.localeCompare(b.sheet) || a.r - b.r || a.c - b.c)
    .map((change) => ({ sheet: change.sheet, address: change.address, before: shown(change.before), after: shown(change.after), formula: change.after?.f }));
  const others = lines.filter((line) => !asked.has(`${line.sheet}!${line.address}`) && line.before !== line.after);
  // LibreOffice timed out or is missing: these have no result yet, which is not the same as becoming empty.
  const unknown = (line) => line.formula && !line.after;
  return {
    written: lines.filter((line) => asked.has(`${line.sheet}!${line.address}`)),
    dependents: others.filter((line) => !unknown(line)),
    uncalculated: others.filter(unknown).length,
  };
}

/** Whether a cell lies within the grid the page built for its sheet. */
function insideGrid(sheets, change) {
  const sheet = sheets.find((entry) => entry.name === change.sheet);
  return Boolean(sheet) && change.r < (sheet.row || 0) && change.c < (sheet.column || 0);
}

/** @param {any[]} sheets @param {any[]} edits @param {string|undefined} defaultSheet */
function verifyFormulas(sheets, edits, defaultSheet) {
  const lookups = new Map();
  return edits.map((edit) => {
    const address = String(edit.address || '').toUpperCase();
    const decoded = XLSX.utils.decode_cell(address);
    const sheet = sheets.find((entry) => entry.name === (edit.sheet || defaultSheet)) || sheets[0];
    if (!lookups.has(sheet)) lookups.set(sheet, new Map(sheet.celldata.map((entry) => [cellKey(entry.r, entry.c), entry.v])));
    const cell = lookups.get(sheet).get(cellKey(decoded.r, decoded.c));
    return { sheet: sheet.name, address, value: cell?.v, display: cell ? String(displayValue(cell)) : '', formula: cell?.f };
  });
}
