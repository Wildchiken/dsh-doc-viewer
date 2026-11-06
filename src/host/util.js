/**
 * Small, dependency-free helpers shared by the host document engine.
 * Everything here is local: no network, no external service.
 */
import { accessSync, closeSync, constants, existsSync, mkdirSync, openSync, readFileSync, readSync, statSync } from 'node:fs';
import { t } from './i18n.js';
import { basename, dirname, extname, join } from 'node:path';
import { randomUUID } from 'node:crypto';

export const MAX_OPEN_BYTES = 120 * 1024 * 1024;
/** Text larger than this is previewed partially; the panel stays responsive. */
export const MAX_TEXT_PREVIEW_CHARS = 2 * 1024 * 1024;

/** Extensions the panel offers to open. Other files are still sniffed if named explicitly. */
export const SUPPORTED_EXTENSIONS = new Set([
  '.xlsx', '.xls', '.csv', '.tsv', '.et',
  '.pdf',
  '.docx', '.doc', '.wps',
  '.txt', '.md', '.markdown', '.log', '.json', '.xml', '.yaml', '.yml', '.ini', '.html', '.htm',
]);

/** @returns {string} a URL-safe random token. */
export function token() {
  return randomUUID().replaceAll('-', '');
}

/**
 * Turn an arbitrary filename into a safe single path segment while keeping
 * Chinese characters readable.
 * @param {string} name
 * @param {string} fallback
 */
export function safeName(name, fallback = 'document') {
  const raw = basename(String(name || fallback));
  const cleaned = raw
    .replace(/[\u0000-\u001f\u007f/\\:*?"<>|]/g, '_')
    .replace(/^\.+$/, '_')
    .slice(0, 180);
  return cleaned || fallback;
}

/** @param {string} path */
export function ensureDir(path) {
  mkdirSync(path, { recursive: true });
  return path;
}

/** @param {string} path */
export function fileExists(path) {
  try {
    return statSync(path).isFile();
  } catch {
    return false;
  }
}

/**
 * Choose a path that does not exist yet. Never overwrites an input or output.
 * @param {string} directory
 * @param {string} baseName
 * @param {string} extension with or without a leading dot
 */
export function uniquePath(directory, baseName, extension) {
  ensureDir(directory);
  const ext = extension.startsWith('.') ? extension : `.${extension}`;
  const stem = safeName(baseName, 'document').replace(new RegExp(`${escapeRe(ext)}$`, 'i'), '');
  let candidate = join(directory, `${stem}${ext}`);
  let index = 2;
  while (existsSync(candidate)) {
    candidate = join(directory, `${stem} (${index})${ext}`);
    index += 1;
  }
  return candidate;
}

/** @param {string} value */
export function escapeRe(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** @param {string} path */
export function readBytes(path, maxBytes = MAX_OPEN_BYTES) {
  const info = statSync(path);
  if (!info.isFile()) throw httpError(400, 'not-a-file', t('目标不是文件'));
  if (info.size > maxBytes) {
    throw httpError(413, 'too-large', t('文件过大（{0}），上限为 {1}', formatBytes(info.size), formatBytes(maxBytes)));
  }
  return readFileSync(path);
}

/** @param {number} bytes */
export function formatBytes(bytes) {
  if (!Number.isFinite(bytes)) return t('未知大小');
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

/**
 * Create an error carrying a stable machine code and a user-facing message.
 * @param {number} status HTTP status for route callers.
 * @param {string} code
 * @param {string} message
 * @param {unknown} [cause]
 */
export function httpError(status, code, message, cause) {
  const error = new Error(message);
  error.status = status;
  error.code = code;
  if (cause !== undefined) error.cause = cause;
  return error;
}

/** @param {unknown} error */
export function errorPayload(error) {
  const err = error instanceof Error ? error : new Error(String(error));
  return {
    ok: false,
    code: typeof err.code === 'string' ? err.code : 'failed',
    // Errors raised on purpose already carry the right wording; only translate library errors.
    message: err.status ? err.message : friendlyMessage(err),
    detail: process.env.DSH_DOC_VIEWER_DEBUG === '1' ? String(err.stack || err.message) : undefined,
  };
}

/** @param {Error} error */
export function friendlyMessage(error) {
  const message = String(error?.message || error);
  if (/password|encrypt/i.test(message)) {
    return t('这个文件已加密或受密码保护，无法直接打开。请先解除密码后再试。');
  }
  if (/unsupported zip|corrupt|bad zip|unexpected end|central directory|invalid pdf|format error|end of data/i.test(message)) {
    return t('文件可能已损坏或格式不受支持，已保留原文件，请尝试其他副本。');
  }
  return message.length > 300 ? `${message.slice(0, 300)}…` : message;
}

/** @param {Uint8Array} bytes */
export function sniffContainer(bytes) {
  if (!bytes || bytes.length === 0) return 'empty';
  if (bytes.length >= 4) {
    if (bytes[0] === 0x50 && bytes[1] === 0x4b) return 'zip';
    if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) return 'ole';
    if (bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) return 'pdf';
  }
  // A UTF-16 BOM marks text even though the payload is full of NUL bytes.
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff)) return 'text';
  const head = Buffer.from(bytes.subarray(0, Math.min(bytes.length, 512))).toString('utf8').replace(/^﻿/, '').trimStart();
  if (/^<!doctype\s+html|^<html|^<table|^<\?xml[\s\S]*<Workbook|^<Workbook/i.test(head)) return 'html';
  if (/^[[{]/.test(head)) return 'json';
  const textish = bytes.subarray(0, Math.min(bytes.length, 4096));
  let control = 0;
  for (const value of textish) {
    if (value === 0) return 'binary';
    if (value < 8 || (value > 13 && value < 32 && value !== 27)) control += 1;
  }
  return control / Math.max(1, textish.length) > 0.02 ? 'binary' : 'text';
}

/**
 * Identify a file from its bytes without reading all of it. A zip's central
 * directory sits at the end and names its parts, which separates a workbook
 * from a Word document or a presentation.
 * @param {string} path
 * @returns {{ container: string, zipKind?: 'xlsx'|'docx'|'other' }}
 */
export function sniffFile(path) {
  const size = statSync(path).size;
  if (size === 0) return { container: 'empty' };
  const fd = openSync(path, 'r');
  try {
    const head = Buffer.alloc(Math.min(size, 8192));
    readSync(fd, head, 0, head.length, 0);
    const container = sniffContainer(head);
    if (container !== 'zip') return { container };
    const tailLength = Math.min(size, 256 * 1024);
    const tail = Buffer.alloc(tailLength);
    readSync(fd, tail, 0, tailLength, size - tailLength);
    const listing = tail.toString('latin1');
    if (listing.includes('xl/workbook.xml') || listing.includes('xl/workbook.bin')) return { container, zipKind: 'xlsx' };
    if (listing.includes('word/document.xml')) return { container, zipKind: 'docx' };
    return { container, zipKind: 'other' };
  } finally {
    closeSync(fd);
  }
}

/** @param {string} name */
export function extensionOf(name) {
  return extname(String(name || '')).toLowerCase();
}

/** @param {string} name */
export function mimeFor(name) {
  switch (extensionOf(name)) {
    case '.pdf': return 'application/pdf';
    case '.xlsx': return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    case '.xls':
    case '.et': return 'application/vnd.ms-excel';
    case '.csv': return 'text/csv; charset=utf-8';
    case '.docx': return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    case '.doc':
    case '.wps': return 'application/msword';
    case '.txt': return 'text/plain; charset=utf-8';
    case '.json': return 'application/json; charset=utf-8';
    case '.svg': return 'image/svg+xml';
    case '.png': return 'image/png';
    case '.jpg':
    case '.jpeg': return 'image/jpeg';
    case '.mjs':
    case '.js': return 'text/javascript; charset=utf-8';
    case '.css': return 'text/css; charset=utf-8';
    case '.wasm': return 'application/wasm';
    case '.woff': return 'font/woff';
    case '.woff2': return 'font/woff2';
    case '.ttf': return 'font/ttf';
    case '.otf': return 'font/otf';
    case '.ttc': return 'font/collection';
    default: return 'application/octet-stream';
  }
}

/**
 * Where a file made from `source` goes: beside it, like Preview or Excel do, so it
 * is found where the original is. A folder that cannot be written to falls back
 * to the workspace's output folder.
 * @param {string} source
 * @param {string} root the workspace
 */
export function exportDirFor(source, root) {
  const directory = dirname(source);
  try {
    accessSync(directory, constants.W_OK);
    return directory;
  } catch {
    return ensureDir(outputDirFor(root));
  }
}

/** @param {string} root */
export function outputDirFor(root) {
  return join(root, '文档查看器', '输出');
}

/** @param {string} root */
export function uploadDirFor(root) {
  return join(root, '文档查看器', '上传');
}
