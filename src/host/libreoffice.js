/**
 * LibreOffice Kit bridge.
 *
 * The Harness ships a native LibreOffice kit (and the `officeToPdf` provider
 * built on top of it). We reuse that provider for Office -> PDF and only fall
 * back to the CLI for operations it does not expose, such as workbook
 * recalculation, PDF page rasterisation for OCR, and legacy WPS conversions.
 */
import { spawn } from 'node:child_process';
import { copyFileSync, existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { ensureDir, httpError, token } from './util.js';

const HERE = dirname(fileURLToPath(import.meta.url));

/** @returns {string|undefined} */
export function resolveCli() {
  // 'none' hides LibreOffice, as on a machine without DeepSeek Harness (CI).
  if (process.env.DSH_LIBREOFFICE_CLI === 'none') return undefined;
  const candidates = [
    process.env.DSH_LIBREOFFICE_CLI,
    process.resourcesPath
      ? join(process.resourcesPath, 'app.asar.unpacked', 'dsh', 'node_modules', '@deepseek-ai', 'libreoffice-kit', 'lib', 'cli.js')
      : undefined,
    '/Applications/DeepSeek Harness.app/Contents/Resources/app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js',
    join(HERE, '..', '..', 'node_modules', '@deepseek-ai', 'libreoffice-kit', 'lib', 'cli.js'),
  ].filter(Boolean);
  return candidates.find((candidate) => existsSync(candidate));
}

/** @param {string} input */
export function workbookExtension(input) {
  const ext = extname(input).toLowerCase();
  return ext === '.xls' || ext === '.xlsx' || ext === '.ods' ? ext : '.xlsx';
}

/**
 * Run the kit CLI with a bounded deadline.
 * @param {string[]} args
 * @param {{ timeoutMs?: number, cwd?: string }} [options]
 */
export async function runKit(args, options = {}) {
  const cli = resolveCli();
  if (!cli) throw httpError(503, 'office-engine-unavailable', '本机未找到 LibreOffice 转换引擎，无法完成这个 Office/PDF 操作。');
  const timeoutMs = options.timeoutMs ?? 120_000;
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [cli, ...args], {
      cwd: options.cwd,
      env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    const timer = setTimeout(() => {
      child.kill('SIGTERM');
      // A kit stuck inside LibreOffice ignores SIGTERM; left alone it would stay forever.
      setTimeout(() => { if (child.exitCode === null && child.signalCode === null) child.kill('SIGKILL'); }, 2000).unref();
      reject(httpError(504, 'office-timeout', '转换超时，请尝试缩小文件或分批处理。'));
    }, timeoutMs);
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });
    child.on('error', (error) => {
      clearTimeout(timer);
      reject(httpError(503, 'office-engine-failed', `无法启动 LibreOffice 引擎：${error.message}`, error));
    });
    child.on('close', (code) => {
      clearTimeout(timer);
      if (code !== 0) {
        reject(httpError(502, 'office-failed', parseKitError(stderr, stdout, code)));
        return;
      }
      resolve({ stdout, stderr });
    });
  });
}

/** @param {string} stderr @param {string} stdout @param {number|null} code */
function parseKitError(stderr, stdout, code) {
  const payload = parseLastJson(stdout) || parseLastJson(stderr);
  if (payload?.error) return String(payload.error);
  const lines = `${stderr}\n${stdout}`.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  return lines.slice(-3).join(' ') || `LibreOffice 转换失败（退出码 ${code}）`;
}

/** @param {string} text */
function parseLastJson(text) {
  for (const line of String(text || '').trim().split(/\r?\n/).reverse()) {
    if (!line.startsWith('{')) continue;
    try { return JSON.parse(line); } catch { /* ignore */ }
  }
  return undefined;
}

/**
 * Recalculate a workbook and return a new XLSX path plus missing fonts.
 * The output path must not exist (the kit refuses to overwrite).
 * @param {string} input
 * @param {{ timeoutMs?: number }} [options]
 */
export async function recalculateWorkbook(input, options = {}) {
  const work = mkdtempSync(join(tmpdir(), 'dsh-doc-recalc-'));
  try {
    const output = join(work, `${basename(input, extname(input))}-${token().slice(0, 8)}.xlsx`);
    const { stdout } = await runKit(['recalculate', '--input', input, '--output', output], { timeoutMs: options.timeoutMs ?? 120_000 });
    if (!existsSync(output)) throw httpError(502, 'office-failed', 'LibreOffice 没有生成重算后的工作簿。');
    const result = parseLastJson(stdout) || {};
    return { path: output, missingFonts: result.missingFonts || [], cleanup: () => rmSync(work, { recursive: true, force: true }) };
  } catch (error) {
    rmSync(work, { recursive: true, force: true });
    throw error;
  }
}

/**
 * Convert a document with the CLI. Used when the officeToPdf service is not
 * mounted and for legacy `.wps`/`.et` content.
 * @param {string} input
 * @param {string} output
 * @param {{ timeoutMs?: number }} [options]
 */
export async function convertWithCli(input, output, options = {}) {
  ensureDir(dirname(output));
  const { stdout } = await runKit(['convert', '--input', input, '--output', output], { timeoutMs: options.timeoutMs ?? 120_000 });
  if (!existsSync(output)) throw httpError(502, 'office-failed', 'LibreOffice 没有生成输出文件。');
  return parseLastJson(stdout) || {};
}

/**
 * Render PDF pages to PNG with the kit's PDFium renderer.
 * @param {string} input
 * @param {number[]} pages one-based page numbers
 * @param {{ dpi?: number }} [options]
 * @returns {Promise<{ directory: string, files: string[], cleanup: () => void }>}
 */
export async function renderPdfPages(input, pages, options = {}) {
  const work = mkdtempSync(join(tmpdir(), 'dsh-doc-render-'));
  try {
    const args = ['render', '--input', input, '--output-dir', work, '--pages', pages.join(','), '--dpi', String(options.dpi || 180)];
    const { stdout } = await runKit(args, { timeoutMs: 180_000 });
    const result = parseLastJson(stdout) || {};
    const files = Array.isArray(result.files)
      ? result.files.map((file) => (typeof file === 'string' ? file : file.path)).filter(Boolean)
      : [];
    if (files.length === 0) {
      // Different kit generations may name the entries differently.
      for (const name of readdirSync(work)) if (/\.png$/i.test(name)) files.push(join(work, name));
    }
    return { directory: work, files, cleanup: () => rmSync(work, { recursive: true, force: true }) };
  } catch (error) {
    rmSync(work, { recursive: true, force: true });
    throw error;
  }
}

/**
 * Reuse the Host `officeToPdf` provider when it is mounted; fall back to the
 * kit CLI otherwise. The provider applies the profile's font configuration and
 * reports missing font families.
 * @param {object} ctx
 * @param {string} input
 * @param {string} output
 * @param {string} extension
 */
export async function convertOfficeToPdf(ctx, input, output, extension) {
  const provider = ctx?.get?.('officeToPdf');
  const normalized = String(extension || extname(input)).replace(/^\./, '').toLowerCase();
  const providerExt = normalized === 'xls' || normalized === 'xlsx' || normalized === 'doc' || normalized === 'docx' || normalized === 'ppt' || normalized === 'pptx'
    ? normalized
    : undefined;
  if (provider && providerExt) {
    const bytes = readFileSync(input);
    if (bytes.byteLength > 50 * 1024 * 1024) {
      throw httpError(413, 'too-large', '文件超过 50 MB，无法使用内置字体转换，请先瘦身。');
    }
    // One version for the whole conversion: the provider gives up when what it reads
    // claims to be a different version from the one it was promised.
    const version = `${bytes.byteLength}:${token()}`;
    const result = await provider.convert({
      extension: providerExt,
      priority: 'foreground',
      source: {
        key: `doc-viewer:${token()}`,
        version,
        bytes: bytes.byteLength,
        async read(_signal, maxBytes) {
          if (bytes.byteLength > maxBytes) throw new Error('source too large');
          return { bytes: new Uint8Array(bytes), version };
        },
      },
    });
    ensureDir(dirname(output));
    writeFileSync(output, Buffer.from(result.pdf));
    return { missingFonts: result.missingFonts || [], backend: 'officeToPdf' };
  }
  const result = await convertWithCli(input, output);
  return { missingFonts: result.missingFonts || [], backend: 'libreoffice-kit' };
}

/**
 * Convert a legacy DOC/WPS/ET file with LibreOffice. The source keeps its real
 * extension, which is what selects LibreOffice's import filter. Native
 * xls/xlsx/docx parsing happens first; this is only for what those cannot read.
 * @param {string} input
 * @param {'xlsx'|'docx'} outputExtension
 */
export async function convertLegacyOffice(input, outputExtension) {
  const work = mkdtempSync(join(tmpdir(), 'dsh-doc-legacy-'));
  try {
    const renamed = join(work, `legacy-${token().slice(0, 8)}${extname(input).toLowerCase()}`);
    copyFileSync(input, renamed);
    const output = join(work, `converted-${token().slice(0, 8)}.${outputExtension}`);
    await convertWithCli(renamed, output);
    const bytes = readFileSync(output);
    return { bytes, cleanup: () => rmSync(work, { recursive: true, force: true }) };
  } catch (error) {
    rmSync(work, { recursive: true, force: true });
    throw error;
  }
}
