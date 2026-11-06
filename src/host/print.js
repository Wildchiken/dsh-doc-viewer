/**
 * Print layout for PDF export. A workbook or a text file is laid out the way a
 * person would set it up before printing: A4, sensible margins, columns as wide
 * as their content, a table that reads as a table, and page numbers.
 */
import { AlignmentType, BorderStyle, Document, ExternalHyperlink, Footer, HeadingLevel, Packer, PageNumber, Paragraph, ShadingType, Table, TableCell, TableLayoutType, TableRow, TextRun, WidthType } from 'docx';
import { t } from './i18n.js';
import { marked } from 'marked';
import * as XLSX from '../../vendor/xlsx/xlsx.mjs';

const FONT = 'PingFang SC';
const MONO = 'Menlo';
const SECONDARY = '86868B';
const LINK = '4D6BFE';
const RULE = 'D2D2D7';
const CODE_FILL = 'F5F5F7';
/** A4 width minus the side margins, in twentieths of a point. */
const CONTENT_WIDTH = 11906 - 1250 * 2;
const MIN_COLUMN_WIDTH = 720;
const BULLETS = ['•', '◦', '▪'];
const HAIRLINE = 'FFE4E7EC';
const HEADER_FILL = 'FFF2F4F7';
const HEADER_RULE = 'FFC9CED6';
/** Characters that fill the printable width of portrait A4 at 10 pt. */
const PORTRAIT_CHARS = 95;
/** Beyond this many rows only the page is set up; drawing every cell would take too long. */
const MAX_STYLED_ROWS = 20_000;
const MIN_WIDTH = 6;
const MAX_WIDTH = 48;
/** Fonts files get by default; they are replaced so the page has one typeface. */
const DEFAULT_FONTS = new Set(['Calibri', 'Arial', 'Cambria', 'Times New Roman', '宋体', 'SimSun', '等线', 'DengXian', 'Microsoft YaHei', '微软雅黑']);

/** Width of text in Excel character units: CJK and full-width forms count double. */
function textWidth(text) {
  let width = 0;
  for (const char of String(text)) width += /[ᄀ-ᅟ⺀-꓏가-힣豈-﫿︰-﹏＀-｠￠-￦]/.test(char) ? 2 : 1;
  return width;
}

/** Excel header/footer codes treat & as a control character. */
function footerText(text) {
  return String(text).replaceAll('&', '&&');
}

/**
 * Lay out one written worksheet for printing.
 * @param {import('exceljs').Worksheet} worksheet
 * @param {{ celldata: any[], config?: any }} sheet the model it was written from
 * @param {{ title: string, sheetCount: number, display: (cell: any) => string }} options
 * @returns {boolean} whether the sheet has anything to print
 */
export function layoutSheetForPrint(worksheet, sheet, options) {
  let maxRow = -1;
  let maxColumn = -1;
  const widths = [];
  const rows = new Map();
  const numeric = [];
  const filled = [];
  for (const entry of sheet.celldata) {
    const text = entry.v ? String(options.display(entry.v) ?? '') : '';
    if (text === '' && !entry.v?.f) continue;
    maxRow = Math.max(maxRow, entry.r);
    maxColumn = Math.max(maxColumn, entry.c);
    // Text in a merged cell spreads over several columns and says nothing about one of them.
    if (!(entry.v?.mc && (entry.v.mc.cs > 1))) widths[entry.c] = Math.max(widths[entry.c] || 0, textWidth(text));
    if (entry.r > 0) {
      filled[entry.c] = (filled[entry.c] || 0) + 1;
      if (typeof entry.v?.v === 'number') numeric[entry.c] = (numeric[entry.c] || 0) + 1;
    }
    if (entry.r <= 1) {
      if (!rows.has(entry.r)) rows.set(entry.r, []);
      rows.get(entry.r).push(entry.v);
    }
  }
  if (maxRow < 0) return false;

  // Columns as wide as their content, within limits; longer text wraps.
  let total = 0;
  for (let c = 0; c <= maxColumn; c += 1) {
    const column = worksheet.getColumn(c + 1);
    const needed = Math.min(MAX_WIDTH, Math.max(MIN_WIDTH, (widths[c] || 0) * 1.1 + 3));
    column.width = Math.max(Math.min(column.width || 0, MAX_WIDTH), needed);
    if ((widths[c] || 0) * 1.1 + 3 > MAX_WIDTH) column.alignment = { ...(column.alignment || {}), wrapText: true, vertical: 'top' };
    total += column.width;
  }

  // A first row of short labels over more rows is a header.
  const first = rows.get(0) || [];
  const hasHeader = maxRow >= 1 && first.length >= 2
    && first.every((cell) => typeof cell?.v === 'string' && !cell.f && cell.v.length <= 40);
  // A sheet someone has styled keeps its look; a plain one gets a clean table.
  const styled = sheet.celldata.some((entry) => entry.v?.bg) || (sheet.config?.borderInfo?.length || 0) > 0;

  if (!styled && maxRow < MAX_STYLED_ROWS) {
    const border = { style: 'thin', color: { argb: HAIRLINE } };
    for (let r = 1; r <= maxRow + 1; r += 1) {
      const row = worksheet.getRow(r);
      row.height = Math.max(row.height || 0, 19);
      for (let c = 1; c <= maxColumn + 1; c += 1) {
        const cell = row.getCell(c);
        const header = hasHeader && r === 1;
        cell.border = header
          ? { top: border, left: border, right: border, bottom: { style: 'thin', color: { argb: HEADER_RULE } } }
          : { top: border, left: border, right: border, bottom: border };
        const own = cell.font || {};
        cell.font = {
          ...own,
          name: own.name && !DEFAULT_FONTS.has(own.name) ? own.name : FONT,
          size: own.size && own.size !== 11 ? own.size : 10,
          ...(header ? { bold: true } : {}),
        };
        if (header) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_FILL } };
        // A header sits over its column: right-aligned above numbers, as in Numbers.
        const numbers = (numeric[c - 1] || 0) > (filled[c - 1] || 0) / 2;
        cell.alignment = { vertical: 'middle', ...(cell.alignment || {}), ...(header ? { horizontal: cell.alignment?.horizontal || (numbers ? 'right' : 'left') } : {}) };
      }
    }
  }

  const range = `A1:${XLSX.utils.encode_col(maxColumn)}${maxRow + 1}`;
  const name = options.sheetCount > 1 ? `${options.title} · ${worksheet.name}` : options.title;
  worksheet.pageSetup = {
    ...worksheet.pageSetup,
    paperSize: 9,
    orientation: total > PORTRAIT_CHARS ? 'landscape' : 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: false,
    showGridLines: false,
    printArea: range,
    ...(hasHeader ? { printTitlesRow: '1:1' } : {}),
    margins: { left: 0.6, right: 0.6, top: 0.7, bottom: 0.75, header: 0.3, footer: 0.35 },
  };
  worksheet.headerFooter = {
    oddFooter: `&L&"${FONT}"&8&K${SECONDARY}${footerText(name)}&R&"${FONT}"&8&K${SECONDARY}&P / &N`,
  };
  return true;
}

/**
 * A text file as a document: paragraphs, or for Markdown its headings, lists and
 * code; logs and data files in a monospaced face.
 * @param {string} text
 * @param {{ title: string, markdown?: boolean, monospace?: boolean }} options
 * @returns {Promise<Buffer>} DOCX bytes
 */
export async function textToDocx(text, options) {
  const lines = text.replace(/\r\n?/g, '\n').replace(/\n+$/, '').split('\n');
  const children = options.markdown
    ? markdownBlocks(marked.lexer(text, { gfm: true }))
    : lines.map((line) => (options.monospace
      ? new Paragraph({ style: 'Code', children: [new TextRun({ text: line || ' ' })] })
      : new Paragraph({ children: [new TextRun({ text: line })] })));
  if (children.length === 0) children.push(new Paragraph({ children: [] }));

  const document = new Document({
    title: options.title,
    creator: t('DeepSeek Harness 文档'),
    styles: {
      default: {
        document: { run: { font: FONT, size: 21, color: '1D1D1F' }, paragraph: { spacing: { line: 330, after: options.markdown ? 120 : 0 } } },
        heading1: { run: { font: FONT, size: 36, bold: true, color: '1D1D1F' }, paragraph: { spacing: { before: 0, after: 200, line: 300 } } },
        heading2: { run: { font: FONT, size: 27, bold: true, color: '1D1D1F' }, paragraph: { spacing: { before: 280, after: 100, line: 300 } } },
        heading3: { run: { font: FONT, size: 23, bold: true, color: '1D1D1F' }, paragraph: { spacing: { before: 220, after: 80, line: 300 } } },
        heading4: { run: { font: FONT, size: 21, bold: true, color: '1D1D1F' }, paragraph: { spacing: { before: 180, after: 60, line: 300 } } },
      },
      paragraphStyles: [
        { id: 'Code', name: 'Code', run: { font: MONO, size: options.monospace ? 16 : 18, color: '1D1D1F' }, paragraph: { spacing: { line: 260, after: 0 } } },
        { id: 'Quote', name: 'Quote', run: { color: '515154' }, paragraph: { indent: { left: 240 }, spacing: { after: 120 } } },
        { id: 'PageFooter', name: 'Page Footer', run: { font: FONT, size: 16, color: SECONDARY }, paragraph: { spacing: { line: 240 } } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1250, right: 1250, footer: 600 } } },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            style: 'PageFooter',
            alignment: AlignmentType.RIGHT,
            children: [
              new TextRun({ text: `${options.title}    `, size: 16, color: SECONDARY }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: SECONDARY }),
              new TextRun({ text: ' / ', size: 16, color: SECONDARY }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: SECONDARY }),
            ],
          })],
        }),
      },
      children,
    }],
  });
  return Packer.toBuffer(document);
}

/**
 * Word paragraphs and tables for Markdown tokens, from the same parser the page's
 * preview uses, so what is previewed is what is exported.
 * @param {any[]} tokens
 * @param {{ quote?: boolean }} [context]
 */
function markdownBlocks(tokens, context = {}) {
  const quote = context.quote ? { style: 'Quote', border: { left: { style: BorderStyle.SINGLE, size: 12, color: RULE, space: 10 } } } : {};
  const blocks = [];
  for (const token of tokens) {
    if (token.type === 'heading') {
      const level = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4][Math.min(token.depth, 4) - 1];
      blocks.push(new Paragraph({ heading: level, children: inlineRuns(token.tokens) }));
    } else if (token.type === 'paragraph') {
      blocks.push(new Paragraph({ ...quote, children: inlineRuns(token.tokens) }));
    } else if (token.type === 'text') {
      blocks.push(new Paragraph({ ...quote, children: token.tokens ? inlineRuns(token.tokens) : [new TextRun(token.text)] }));
    } else if (token.type === 'code') {
      for (const line of token.text.split('\n')) {
        blocks.push(new Paragraph({ style: 'Code', shading: { type: ShadingType.CLEAR, fill: CODE_FILL, color: 'auto' }, children: [new TextRun({ text: line || ' ' })] }));
      }
      blocks.push(new Paragraph({ spacing: { after: 60, line: 120 }, children: [] }));
    } else if (token.type === 'blockquote') {
      blocks.push(...markdownBlocks(token.tokens, { quote: true }));
    } else if (token.type === 'list') {
      blocks.push(...listBlocks(token, 0));
    } else if (token.type === 'table') {
      blocks.push(markdownTable(token), new Paragraph({ spacing: { after: 60, line: 120 }, children: [] }));
    } else if (token.type === 'hr') {
      blocks.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 1 } }, spacing: { after: 160, line: 200 }, children: [] }));
    } else if (token.type === 'html') {
      const plain = token.text.replace(/<[^>]*>/g, '').trim();
      if (plain) blocks.push(new Paragraph({ ...quote, children: [new TextRun(plain)] }));
    }
  }
  return blocks;
}

/**
 * A list as paragraphs with their own markers: numbers keep the list's start, and
 * nesting is only indentation.
 * @param {any} list
 * @param {number} depth
 */
function listBlocks(list, depth) {
  const blocks = [];
  const left = 400 + depth * 360;
  const start = Number(list.start) || 1;
  list.items.forEach((item, index) => {
    const checkbox = item.tokens.find((token) => token.type === 'checkbox');
    const marker = item.task ? (checkbox?.checked ?? item.checked ? '☑' : '☐')
      : list.ordered ? `${start + index}.`
      : BULLETS[Math.min(depth, BULLETS.length - 1)];
    let marked = false;
    for (const token of item.tokens) {
      if (token.type === 'checkbox' || token.type === 'space') continue;
      if (token.type === 'list') {
        blocks.push(...listBlocks(token, depth + 1));
      } else if (token.type === 'text' || token.type === 'paragraph') {
        const runs = token.tokens ? inlineRuns(token.tokens) : [new TextRun(token.text)];
        blocks.push(new Paragraph({
          indent: marked ? { left } : { left, hanging: 300 },
          tabStops: [{ type: 'left', position: left }],
          spacing: { after: 60 },
          children: marked ? runs : [new TextRun(`${marker}\t`), ...runs],
        }));
        marked = true;
      } else {
        blocks.push(...markdownBlocks([token]));
      }
    }
  });
  return blocks;
}

/** @param {any} token a Markdown table */
function markdownTable(token) {
  const count = token.header.length;
  const widest = Array.from({ length: count }, (_, column) => Math.max(...[token.header, ...token.rows].map((row) => textWidth(row[column]?.text ?? '')), 2));
  const total = widest.reduce((sum, width) => sum + width, 0);
  // Content decides each column's share, and no column gets too narrow to read.
  const raw = widest.map((width) => Math.max(MIN_COLUMN_WIDTH, Math.round((width / total) * CONTENT_WIDTH)));
  const scale = CONTENT_WIDTH / raw.reduce((sum, width) => sum + width, 0);
  const columnWidths = raw.map((width) => Math.round(width * scale));
  const rule = { style: BorderStyle.SINGLE, size: 4, color: RULE };
  const alignment = { center: AlignmentType.CENTER, right: AlignmentType.RIGHT, left: AlignmentType.LEFT };
  const row = (cells, header) => new TableRow({
    tableHeader: header,
    cantSplit: true,
    children: cells.map((cell, column) => new TableCell({
      width: { size: columnWidths[column], type: WidthType.DXA },
      borders: { top: rule, bottom: rule, left: rule, right: rule },
      margins: { top: 50, bottom: 50, left: 100, right: 100 },
      shading: header ? { type: ShadingType.CLEAR, fill: CODE_FILL, color: 'auto' } : undefined,
      children: [new Paragraph({
        alignment: alignment[token.align[column]],
        spacing: { after: 0, line: 280 },
        children: inlineRuns(cell.tokens, header ? { bold: true } : {}),
      })],
    })),
  });
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths,
    layout: TableLayoutType.FIXED,
    rows: [row(token.header, true), ...token.rows.map((cells) => row(cells, false))],
  });
}

/**
 * Text runs for Markdown's inline tokens, carrying emphasis down through nesting.
 * @param {any[]} tokens
 * @param {any} [style]
 */
function inlineRuns(tokens, style = {}) {
  const runs = [];
  for (const token of tokens || []) {
    if (token.type === 'strong') runs.push(...inlineRuns(token.tokens, { ...style, bold: true }));
    else if (token.type === 'em') runs.push(...inlineRuns(token.tokens, { ...style, italics: true }));
    else if (token.type === 'del') runs.push(...inlineRuns(token.tokens, { ...style, strike: true }));
    else if (token.type === 'codespan') runs.push(new TextRun({ ...style, text: token.text, font: MONO, size: 19, shading: { type: ShadingType.CLEAR, fill: CODE_FILL, color: 'auto' } }));
    else if (token.type === 'br') runs.push(new TextRun({ ...style, break: 1 }));
    else if (token.type === 'link') {
      const children = inlineRuns(token.tokens, { ...style, color: LINK });
      if (/^(https?:|mailto:)/i.test(token.href)) runs.push(new ExternalHyperlink({ link: token.href, children }));
      else runs.push(...children);
    } else if (token.type === 'image') runs.push(new TextRun({ ...style, text: t('[图片：{0}]', token.text || token.href), color: SECONDARY }));
    else if (token.type === 'html') {
      if (/^<br\s*\/?>$/i.test(token.text)) runs.push(new TextRun({ ...style, break: 1 }));
      // Other inline tags only wrap text, which arrives as its own tokens.
    } else if (token.tokens) runs.push(...inlineRuns(token.tokens, style));
    else if (token.text) runs.push(new TextRun({ ...style, text: token.text }));
  }
  return runs;
}
