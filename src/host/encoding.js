import { createRequire } from 'node:module';

/**
 * Text decoding for CSV/TSV/TXT and legacy text-like files.
 *
 * utf-8 with a strict decoder is tried first. If the bytes are not valid UTF-8,
 * GB18030 is used (a superset of GBK/GB2312), which covers the CSV exports from
 * Chinese ERP, WPS and government systems. UTF-16 BOMs are honoured.
 */

/** @param {Uint8Array} bytes */
function decodeStrict(bytes, encoding) {
  return new TextDecoder(encoding, { fatal: true }).decode(bytes);
}

/** @param {Uint8Array} bytes */
function countReplacement(text) {
  return (text.match(/\uFFFD/g) || []).length;
}

/**
 * @param {Uint8Array} bytes
 * @returns {{ text: string, encoding: string, bom: boolean }}
 */
export function decodeText(bytes) {
  if (!bytes || bytes.length === 0) return { text: '', encoding: 'utf-8', bom: false };
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return { text: new TextDecoder('utf-8').decode(bytes.subarray(3)), encoding: 'utf-8', bom: true };
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return { text: new TextDecoder('utf-16le').decode(bytes.subarray(2)), encoding: 'utf-16le', bom: true };
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return { text: new TextDecoder('utf-16be').decode(bytes.subarray(2)), encoding: 'utf-16be', bom: true };
  }
  try {
    return { text: decodeStrict(bytes, 'utf-8'), encoding: 'utf-8', bom: false };
  } catch {
    // Not valid UTF-8. Try GB18030 before falling back to a lenient decoder.
  }
  try {
    const text = decodeStrict(bytes, 'gb18030');
    return { text, encoding: 'gb18030', bom: false };
  } catch {
    // Some legacy files contain a few invalid bytes; GB18030 is lenient enough
    // in non-fatal mode and still renders the overwhelming majority correctly.
  }
  try {
    const text = new TextDecoder('gb18030').decode(bytes);
    if (countReplacement(text) < text.length * 0.01) {
      return { text, encoding: 'gb18030', bom: false };
    }
  } catch {
    // fall through
  }
  try {
    return { text: new TextDecoder('big5').decode(bytes), encoding: 'big5', bom: false };
  } catch {
    return { text: new TextDecoder('utf-8').decode(bytes), encoding: 'utf-8', bom: false };
  }
}

/**
 * Decide whether text bytes look like a delimited table rather than prose.
 * @param {string} text
 */
export function looksDelimited(text) {
  const lines = text.split(/\r?\n/).filter((line) => line.trim() !== '').slice(0, 12);
  if (lines.length < 2) return false;
  for (const [delimiter, name] of [[',', 'csv'], ['\t', 'tsv'], [';', 'csv']]) {
    const counts = lines.map((line) => line.split(delimiter).length);
    const first = counts[0];
    if (first > 1 && counts.filter((n) => n === first).length >= Math.min(3, counts.length)) {
      return name;
    }
  }
  return false;
}

/** @param {string} text */
export function countLines(text) {
  if (!text) return 0;
  let lines = 1;
  for (let i = 0; i < text.length; i += 1) if (text.charCodeAt(i) === 10) lines += 1;
  return lines;
}

/** @param {string} text */
export function stripBom(text) {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/* Writing text back the way the file was written ----------------------- */


let codepages;
/** SheetJS's code page tables, loaded the first time a legacy encoding is written. */
function cptable() {
  codepages ??= createRequire(import.meta.url)('../../vendor/xlsx/cpexcel.cjs');
  return codepages;
}

/**
 * How a text file is written: its encoding, whether it starts with a byte order
 * mark, and its line endings.
 * @param {Uint8Array} bytes
 * @returns {{ encoding: string, bom: boolean, newline: string }}
 */
export function textFormatOf(bytes) {
  const { text, encoding, bom } = decodeText(bytes);
  return { encoding, bom, newline: text.includes('\r\n') ? '\r\n' : '\n' };
}

/**
 * Text as bytes in a given format. A character the legacy encoding cannot hold
 * would turn into "?", so then the text is written as UTF-8 with a BOM instead,
 * which every spreadsheet program reads, and `changed` says so.
 * @param {string} text with \n line endings
 * @param {{ encoding: string, bom: boolean, newline: string }} format
 * @returns {{ bytes: Buffer, changed?: string }}
 */
export function encodeText(text, format) {
  const body = text.replace(/\r\n?/g, '\n').replaceAll('\n', format.newline || '\n');
  if (format.encoding === 'gb18030' || format.encoding === 'big5') {
    const page = format.encoding === 'big5' ? 950 : 936;
    const bytes = Buffer.from(cptable().utils.encode(page, body));
    if (cptable().utils.decode(page, bytes) === body) return { bytes };
    return { bytes: Buffer.from(`﻿${body}`, 'utf8'), changed: `新内容里有 ${format.encoding.toUpperCase()} 编码表示不了的字符，已改存为 UTF-8` };
  }
  if (format.encoding === 'utf-16le' || format.encoding === 'utf-16be') {
    const le = Buffer.from(`﻿${body}`, 'utf16le');
    if (format.encoding === 'utf-16le') return { bytes: le };
    for (let index = 0; index + 1 < le.length; index += 2) [le[index], le[index + 1]] = [le[index + 1], le[index]];
    return { bytes: le };
  }
  return { bytes: Buffer.from(format.bom ? `﻿${body}` : body, 'utf8') };
}
