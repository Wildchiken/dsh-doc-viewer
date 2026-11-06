/**
 * Local PDF, Word and OCR operations.
 *
 * PDF rendering in the panel is done by PDF.js in the browser. This module
 * owns the Host-side work: text extraction, table extraction, PDF page edits,
 * scanned-page OCR (macOS Vision), and DOCX -> HTML/text.
 */
import { spawn } from 'node:child_process';
import { t } from './i18n.js';
import { accessSync, chmodSync, constants, existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Document, Packer, PageBreak, Paragraph, TextRun } from 'docx';
import mammoth from 'mammoth';
import { PDFDocument } from 'pdf-lib';
import * as XLSX from '../../vendor/xlsx/xlsx.mjs';
import { decodeText } from './encoding.js';
import { httpError, readBytes, uniquePath } from './util.js';

const HERE = dirname(fileURLToPath(import.meta.url));
let pdfjsPromise;

/** Load PDF.js once per process. */
async function pdfjs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import('pdfjs-dist/legacy/build/pdf.mjs');
  }
  return pdfjsPromise;
}

/** @param {unknown} error */
function pdfFailure(error) {
  const name = error?.name || '';
  const message = String(error?.message || error);
  if (name === 'PasswordException' || /password|encrypted/i.test(message)) {
    return httpError(423, 'encrypted', t('这个 PDF 已加密或需要密码，无法读取。请先解除密码后再试。'), error);
  }
  if (name === 'InvalidPDFException' || /Invalid PDF|structure|corrupt|bad xref|xref|Failed to parse PDF/i.test(message)) {
    return httpError(422, 'corrupt-pdf', t('PDF 文件可能已损坏，已保留原文件。'), error);
  }
  return httpError(422, 'pdf-failed', t('PDF 处理失败：{0}', message), error);
}

/** @param {Uint8Array} data */
function hostPdfOptions(data) {
  return {
    data,
    disableFontFace: true,
    useSystemFonts: false,
    isEvalSupported: false,
    useWorkerFetch: false,
    verbosity: 0,
  };
}

/**
 * Inspect a PDF without rendering it.
 * @param {string} path
 */
export async function inspectPdf(path) {
  const data = new Uint8Array(readBytes(path));
  try {
    const lib = await pdfjs();
    const loading = lib.getDocument(hostPdfOptions(data));
    const doc = await loading.promise;
    const meta = await doc.getMetadata().catch(() => ({}));
    const firstName = meta?.info?.Title || basename(path);
    const pages = [];
    const max = doc.numPages;
    for (let number = 1; number <= Math.min(max, 4); number += 1) {
      const page = await doc.getPage(number);
      const viewport = page.getViewport({ scale: 1 });
      pages.push({ page: number, width: round(viewport.width), height: round(viewport.height) });
    }
    const title = String(firstName);
    const pageCount = doc.numPages;
    await loading.destroy();
    return {
      kind: 'pdf',
      path,
      name: basename(path),
      ext: extname(path).toLowerCase(),
      title,
      pageCount,
      pages,
      encrypted: false,
    };
  } catch (error) {
    throw pdfFailure(error);
  }
}

/**
 * Extract text and positioned text items from every page.
 * @param {string} path
 * @param {{ pages?: number[], onProgress?: (page: number, total: number) => void }} [options]
 */
export async function extractPdf(path, options = {}) {
  const data = new Uint8Array(readBytes(path));
  try {
    const lib = await pdfjs();
    const loading = lib.getDocument(hostPdfOptions(data));
    const doc = await loading.promise;
    const wanted = options.pages?.length ? new Set(options.pages.filter((page) => page >= 1 && page <= doc.numPages)) : undefined;
    const pages = [];
    for (let number = 1; number <= doc.numPages; number += 1) {
      if (wanted && !wanted.has(number)) continue;
      const page = await doc.getPage(number);
      const viewport = page.getViewport({ scale: 1 });
      const content = await page.getTextContent();
      const items = content.items.map((item) => ({
        text: item.str,
        x: round(item.transform?.[4] || 0),
        y: round(item.transform?.[5] || 0),
        width: round(item.width || 0),
        height: round(item.height || 0),
      })).filter((item) => item.text !== '');
      const text = readingOrder(items);
      pages.push({
        page: number,
        width: round(viewport.width),
        height: round(viewport.height),
        text,
        items,
      });
      options.onProgress?.(number, doc.numPages);
    }
    const pageCount = doc.numPages;
    await loading.destroy();
    return { path, pages, pageCount };
  } catch (error) {
    throw pdfFailure(error);
  }
}

/**
 * Join positioned PDF text in reading order (top-to-bottom, left-to-right).
 * PDF.js emits content-stream order, which is not always visual order.
 * @param {{text:string,x:number,y:number}[]} items
 */
function readingOrder(items) {
  const lines = [];
  for (const item of [...items].sort((a, b) => b.y - a.y)) {
    const last = lines[lines.length - 1];
    if (last && Math.abs(last.y - item.y) <= 3) last.items.push(item);
    else lines.push({ y: item.y, items: [item] });
  }
  return lines
    .map((line) => line.items.sort((a, b) => a.x - b.x).map((item) => item.text).join(' '))
    .join('\n');
}

/** @param {number} value */
function round(value) {
  return Math.round(value * 100) / 100;
}

/**
 * Extract text to TXT.
 * @param {string} path
 * @param {string} output
 */
export async function pdfToTxt(path, output) {
  const result = await extractPdf(path);
  const text = result.pages.map((page) => t('===== 第 {0} 页 =====\n{1}', page.page, page.text)).join('\n\n');
  writeFileSync(output, `\uFEFF${text}\n`, 'utf8');
  return { output, pageCount: result.pageCount, chars: text.length };
}

/**
 * Build table rows from PDF text positions. Tables are detected by rows with
 * several positioned items; prose falls back to one paragraph per row.
 * @param {Awaited<ReturnType<typeof extractPdf>>} extraction
 */
export function pdfTables(extraction) {
  const allRows = [];
  for (const page of extraction.pages) {
    if (page.items.length === 0) {
      allRows.push([t('（第 {0} 页没有可提取文字，可能是扫描版，请使用 OCR）', page.page)]);
      continue;
    }
    const lines = [];
    for (const item of [...page.items].filter((entry) => entry.text.trim()).sort((a, b) => b.y - a.y)) {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.y - item.y) <= 3) last.items.push(item);
      else lines.push({ y: item.y, items: [item] });
    }
    for (const { items } of lines) {
      items.sort((a, b) => a.x - b.x);
      const cells = [];
      let current = null;
      for (const item of items) {
        if (current && item.x - current.end <= 1.6) {
          current.text += item.text;
          current.end = Math.max(current.end, item.x + item.width);
        } else {
          if (current) cells.push(current.text.trim());
          current = { text: item.text, end: item.x + item.width };
        }
      }
      if (current) cells.push(current.text.trim());
      allRows.push(cells.filter((cell, index, list) => cell !== '' || index === list.length - 1));
    }
  }
  const width = Math.max(1, ...allRows.map((row) => row.length));
  return allRows.map((row) => [...row, ...new Array(width - row.length).fill('')]);
}

/**
 * PDF -> XLSX. A best-effort table reconstruction; complex layouts are not
 * preserved. The caller shows a lossy-conversion warning first.
 * @param {string} path
 * @param {string} output
 */
export async function pdfToXlsx(path, output) {
  const extraction = await extractPdf(path);
  const rows = pdfTables(extraction);
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols'] = new Array(Math.max(1, rows[0]?.length || 1)).fill({ wch: 18 });
  XLSX.utils.book_append_sheet(workbook, sheet, t('PDF 表格'));
  writeFileSync(output, XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer', compression: true }));
  const textChars = extraction.pages.reduce((sum, page) => sum + page.text.length, 0);
  return {
    output,
    pages: extraction.pageCount,
    rows: rows.length,
    columns: rows[0]?.length || 0,
    textChars,
    warning: textChars < 20 ? t('这份 PDF 几乎没有可提取文字，可能是扫描版；请使用 OCR。') : undefined,
  };
}

/**
 * PDF -> DOCX with page breaks and selectable text. Layout, images and tables
 * are not preserved; the UI marks this as a lossy conversion.
 * @param {string} path
 * @param {string} output
 */
export async function pdfToDocx(path, output) {
  const extraction = await extractPdf(path);
  const children = [];
  extraction.pages.forEach((page, index) => {
    if (index > 0) children.push(new Paragraph({ children: [new PageBreak()] }));
    const lines = page.text.split(/\n+/).map((line) => line.trim()).filter(Boolean);
    if (lines.length === 0) {
      children.push(new Paragraph({ children: [new TextRun({ text: t('（本页没有可提取文字）'), italics: true })] }));
    } else {
      for (const line of lines) {
        children.push(new Paragraph({
          spacing: { after: 80 },
          children: [new TextRun({ text: line, font: 'PingFang SC', size: 22 })],
        }));
      }
    }
  });
  const doc = new Document({
    creator: t('DeepSeek Harness 文档查看器'),
    title: basename(path, extname(path)),
    sections: [{ properties: {}, children }],
  });
  writeFileSync(output, await Packer.toBuffer(doc));
  return { output, pages: extraction.pageCount };
}

/**
 * Package archives keep executable bits only for declared `bin` entries, so the OCR
 * helper can arrive as a plain file; it is made runnable where it is installed.
 * @param {string} path
 */
export function ensureExecutable(path) {
  try {
    accessSync(path, constants.X_OK);
  } catch {
    try {
      chmodSync(path, statSync(path).mode | 0o111);
    } catch (error) {
      throw httpError(503, 'ocr-unavailable', t('OCR 组件没有执行权限，也无法自动修复。可以手动运行：chmod +x "{0}"', path), error);
    }
  }
}

/**
 * OCR a scanned PDF or image with macOS Vision. Pages are rasterized by
 * PDFKit and recognized locally; no network request is made.
 * @param {string} path
 * @param {{ maxPages?: number, pages?: string }} [options] `pages` limits a PDF to e.g. "1-20".
 */
export async function ocrDocument(path, options = {}) {
  if (process.platform !== 'darwin') {
    throw httpError(503, 'ocr-unavailable', t('OCR 使用 macOS Vision，当前系统不支持。'));
  }
  const helper = join(HERE, '..', '..', 'bin', 'vision-ocr');
  if (!existsSync(helper)) {
    throw httpError(503, 'ocr-unavailable', t('本机 OCR 组件未安装，无法识别扫描件。请重新安装文档查看器。'));
  }
  ensureExecutable(helper);
  const work = mkdtempSync(join(tmpdir(), 'dsh-doc-ocr-'));
  try {
    const output = join(work, 'ocr.json');
    let target = path;
    let pageNumbers;
    if (extname(path).toLowerCase() === '.pdf') {
      const info = await inspectPdf(path);
      pageNumbers = parsePageList(options.pages, info.pageCount);
      if (options.maxPages && pageNumbers.length > options.maxPages) {
        throw httpError(413, 'ocr-too-many-pages', t('要识别 {0} 页，超过单次上限 {1} 页。请用页码范围分批识别，例如 1-{2}。', pageNumbers.length, options.maxPages, options.maxPages));
      }
      if (pageNumbers.length < info.pageCount) {
        const source = await PDFDocument.load(readFileSync(path), { ignoreEncryption: false });
        const subset = await PDFDocument.create();
        for (const page of await subset.copyPages(source, pageNumbers.map((number) => number - 1))) subset.addPage(page);
        target = join(work, 'pages.pdf');
        writeFileSync(target, await subset.save());
      }
    }
    await new Promise((resolve, reject) => {
      const child = spawn(helper, [target, output], { stdio: ['ignore', 'pipe', 'pipe'] });
      let stderr = '';
      const timer = setTimeout(() => {
        child.kill('SIGTERM');
        reject(httpError(504, 'ocr-timeout', t('OCR 超时，请缩小页范围后重试。')));
      }, 180_000);
      child.stderr.on('data', (chunk) => { stderr += chunk; });
      child.on('error', (error) => {
        clearTimeout(timer);
        reject(httpError(503, 'ocr-failed', t('OCR 组件启动失败：{0}', error.message), error));
      });
      child.on('close', (code) => {
        clearTimeout(timer);
        if (code !== 0) reject(httpError(422, 'ocr-failed', stderr.trim() || t('OCR 失败（退出码 {0}）', code)));
        else resolve();
      });
    });
    const payload = JSON.parse(readFileSync(output, 'utf8'));
    // A subset was recognized as its own document; report the original page numbers.
    const pages = (Array.isArray(payload.pages) ? payload.pages : []).map((page, index) => ({ ...page, page: pageNumbers?.[index] ?? page.page }));
    return {
      pages,
      engine: payload.engine || 'macOS Vision',
      text: pages.map((page) => t('===== 第 {0} 页 =====\n{1}', page.page, page.text)).join('\n\n'),
    };
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

/**
 * Merge PDFs in the given order.
 * @param {string[]} paths
 * @param {string} output
 */
export async function mergePdfs(paths, output) {
  if (!Array.isArray(paths) || paths.length < 2) throw httpError(400, 'need-two-pdfs', t('合并至少需要两个 PDF 文件'));
  const merged = await PDFDocument.create();
  const sources = [];
  for (const path of paths) {
    try {
      const source = await PDFDocument.load(readFileSync(path), { ignoreEncryption: false });
      sources.push({ path, source });
    } catch (error) {
      throw pdfFailure(error);
    }
  }
  for (const { source } of sources) {
    const pages = await merged.copyPages(source, source.getPageIndices());
    for (const page of pages) merged.addPage(page);
  }
  writeFileSync(output, await merged.save());
  return { output, files: paths.length, pages: merged.getPageCount() };
}

/**
 * Split a PDF into one file per page range.
 * @param {string} path
 * @param {string} outputDirectory
 * @param {{ ranges?: string, title?: string }} [options]
 */
export async function splitPdf(path, outputDirectory, options = {}) {
  let source;
  try {
    source = await PDFDocument.load(readFileSync(path), { ignoreEncryption: false });
  } catch (error) {
    throw pdfFailure(error);
  }
  const ranges = parseRanges(options.ranges, source.getPageCount());
  const base = options.title || basename(path, extname(path));
  const outputs = [];
  for (const range of ranges) {
    const doc = await PDFDocument.create();
    const pages = await doc.copyPages(source, range.map((page) => page - 1));
    for (const page of pages) doc.addPage(page);
    const label = range.length === 1 ? t('第{0}页', range[0]) : t('{0}-{1}页', range[0], range[range.length - 1]);
    const output = uniquePath(outputDirectory, `${base}-${label}`, '.pdf');
    writeFileSync(output, await doc.save());
    outputs.push({ output, pages: range });
  }
  return { outputs, sourcePages: source.getPageCount() };
}

/**
 * Delete pages from a PDF.
 * @param {string} path
 * @param {string} output
 * @param {string} pages e.g. "2,4-6"
 */
export async function deletePdfPages(path, output, pages) {
  let source;
  try {
    source = await PDFDocument.load(readFileSync(path), { ignoreEncryption: false });
  } catch (error) {
    throw pdfFailure(error);
  }
  const remove = new Set(parseRanges(pages, source.getPageCount()).flat());
  const keep = source.getPageIndices().filter((index) => !remove.has(index + 1));
  if (keep.length === 0) throw httpError(400, 'all-pages-removed', t('不能删除全部页面，请保留至少一页。'));
  const doc = await PDFDocument.create();
  const copied = await doc.copyPages(source, keep);
  for (const page of copied) doc.addPage(page);
  writeFileSync(output, await doc.save());
  return { output, removed: remove.size, remaining: keep.length };
}

/**
 * Parse page expressions such as "1,3,5-8".
 * @param {string|undefined} expression
 * @param {number} pageCount
 * @returns {number[][]}
 */
export function parseRanges(expression, pageCount) {
  const text = String(expression || '').trim();
  const groups = [];
  if (!text || text.toLowerCase() === 'all') {
    for (let page = 1; page <= pageCount; page += 1) groups.push([page]);
    return groups;
  }
  for (const part of text.split(',')) {
    const token = part.trim();
    if (!token) continue;
    const match = /^(\d+)(?:\s*-\s*(\d+))?$/.exec(token);
    if (!match) throw httpError(400, 'bad-pages', t('页码格式不正确：{0}', token));
    const start = Number(match[1]);
    const end = match[2] ? Number(match[2]) : start;
    if (start < 1 || end > pageCount || start > end) {
      throw httpError(400, 'bad-pages', t('页码 {0} 超出范围（共 {1} 页）', token, pageCount));
    }
    const group = [];
    for (let page = start; page <= end; page += 1) group.push(page);
    groups.push(group);
  }
  if (groups.length === 0) throw httpError(400, 'bad-pages', t('请指定页码'));
  return groups;
}

/**
 * The pages named by an expression, each once and in reading order.
 * @param {string|undefined} expression
 * @param {number} pageCount
 * @returns {number[]}
 */
export function parsePageList(expression, pageCount) {
  return [...new Set(parseRanges(expression, pageCount).flat())].sort((a, b) => a - b);
}

/**
 * DOCX -> sanitized HTML for the panel.
 * @param {string} path
 */
export async function docxToHtml(path) {
  try {
    const result = await mammoth.convertToHtml(
      { path },
      {
        convertImage: mammoth.images.imgElement(async (image) => {
          const base64 = await image.read('base64');
          return { src: `data:${image.contentType};base64,${base64}`, alt: image.altText || '' };
        }),
      },
    );
    return {
      html: result.value,
      messages: result.messages.map((message) => message.message).slice(0, 20),
    };
  } catch (error) {
    throw httpError(422, 'docx-failed', t('无法读取 Word 文档：{0}', error instanceof Error ? error.message : String(error)), error);
  }
}

/**
 * The plain text of a DOCX, paragraph by paragraph.
 * @param {string} path
 * @returns {Promise<string>}
 */
export async function docxRawText(path) {
  try {
    const result = await mammoth.extractRawText({ path });
    return result.value.replace(/\n{3,}/g, '\n\n').trim();
  } catch (error) {
    throw httpError(422, 'docx-failed', t('无法读取 Word 文档：{0}', error instanceof Error ? error.message : String(error)), error);
  }
}

/**
 * DOCX -> plain text.
 * @param {string} path
 * @param {string} output
 */
export async function docxToTxt(path, output) {
  const text = await docxRawText(path);
  writeFileSync(output, `\uFEFF${text}\n`, 'utf8');
  return { output, chars: text.length };
}

/** Read a text file with encoding detection. */
export function readTextDocument(path) {
  const bytes = readBytes(path);
  return decodeText(bytes);
}
