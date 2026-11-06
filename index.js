/**
 * Host half of 文档查看器 / Document Viewer.
 *
 * Registers one loopback-only HTTP surface for the panel (import, open, save,
 * convert, PDF tools, pending open requests), plus six agent tools. All file
 * processing is local. The `officeToPdf` service is reused when available.
 */
import { spawn } from 'node:child_process';
import { t, withLang } from './src/host/i18n.js';
import { createReadStream, existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { basename, dirname, extname, isAbsolute, join, normalize, resolve, sep } from 'node:path';
import { promisify } from 'node:util';
import { gzip } from 'node:zlib';
import { errorPayload, ensureDir, exportDirFor, httpError, mimeFor, safeName, SUPPORTED_EXTENSIONS, uniquePath } from './src/host/util.js';
import { tmpdir } from 'node:os';
import { calculateSheets, createDocument, openDocument, saveInPlace, saveSpreadsheetCopy, saveTextInPlace, statsForFile, convertMany, pdfOperation } from './src/host/documents.js';
import { pickFiles, pickSaveLocation } from './src/host/native.js';
import { listVersions, restoreVersion, withFileLock } from './src/host/versions.js';
import { discardRecovery, listRecovery, readRecovery, saveRecovery } from './src/host/recovery.js';
import { listOpenRequests, ackOpenRequests, setPanelState } from './src/host/requests.js';
import { registerTools } from './src/host/tools.js';

const PLUGIN_NAME = 'dsh-doc-viewer';
const ROUTE_PREFIX = '/dsh-doc-viewer';
const BODY_JSON_MAX = 220 * 1024 * 1024;
const BODY_UPLOAD_MAX = 150 * 1024 * 1024;
const MAX_LISTED_ENTRIES = 500;
const require = createRequire(import.meta.url);
const VERSION = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8')).version;
const gzipAsync = promisify(gzip);
const allowedFiles = new Set();
/** Places the user picked in the save panel. */
const chosenTargets = new Set();
/** Files dragged in from Finder arrive without their path; their copies live here, not in the user's folders. */
const DROPPED_DIR = join(tmpdir(), 'dsh-doc-viewer-dropped');

export const inject = ['webServer', 'tools', 'workingDirectory'];

/**
 * Register routes and agent tools for the lifetime of the plugin.
 * @param {any} ctx
 */
export function apply(ctx) {
  registerTools(ctx);
  ctx.effect(() => {
    const dispose = ctx.webServer.register({
      kind: 'prefix',
      path: ROUTE_PREFIX,
      handler: (req, res) => {
        void handle(ctx, req, res).catch((error) => {
          if (res.headersSent) {
            res.destroy();
            return;
          }
          return sendJson(res, Number(error?.status) || 500, errorPayload(error));
        });
      },
    });
    return typeof dispose === 'function' ? dispose : () => {};
  }, `${PLUGIN_NAME}: http surface`);
}

/**
 * @param {any} ctx
 * @param {import('node:http').IncomingMessage} req
 * @param {import('node:http').ServerResponse} res
 */
async function handle(ctx, req, res) {
  // Messages follow the language the page is shown in; the agent's tools then use it too.
  const language = req.headers['x-dvd-lang'];
  return withLang(Array.isArray(language) ? language[0] : language, () => handleRequest(ctx, req, res));
}

/** @param {any} ctx @param {import('node:http').IncomingMessage} req @param {import('node:http').ServerResponse} res */
async function handleRequest(ctx, req, res) {
  const refused = refusalOf(ctx, req);
  if (refused !== undefined) {
    await sendJson(res, refused, { ok: false, code: refused === 401 ? 'unauthorized' : 'forbidden', message: t('仅允许本机页面访问文档查看器。') });
    return;
  }
  const url = new URL(req.url || '/', 'http://127.0.0.1');
  let pathname;
  try {
    pathname = decodeURIComponent(url.pathname);
  } catch {
    throw httpError(400, 'bad-url', t('请求地址格式不正确。'));
  }
  if (pathname.startsWith(`${ROUTE_PREFIX}/assets/`)) {
    await serveAsset(pathname.slice(`${ROUTE_PREFIX}/assets/`.length), res);
    return;
  }
  const route = pathname.slice(ROUTE_PREFIX.length) || '/';
  const key = `${req.method} ${route}`;
  switch (key) {
    case 'GET /api/ping':
    case 'GET /ping':
      await sendJson(res, 200, { ok: true, name: PLUGIN_NAME, version: VERSION });
      return;
    case 'GET /api/requests':
      await sendJson(res, 200, { ok: true, requests: listOpenRequests() });
      return;
    case 'POST /api/requests/ack': {
      const body = await readJson(req, 1024 * 1024);
      await sendJson(res, 200, { ok: true, claimed: ackOpenRequests(Array.isArray(body.ids) ? body.ids : []) });
      return;
    }
    case 'POST /api/state': {
      setPanelState(await readJson(req, 256 * 1024));
      await sendJson(res, 200, { ok: true });
      return;
    }
    case 'GET /api/recovery':
      await sendJson(res, 200, { ok: true, entries: listRecovery() });
      return;
    case 'GET /api/recovery/item': {
      const path = resolve(String(url.searchParams.get('path') || ''));
      const record = readRecovery(path);
      // Opening a recovered document makes it this page's again, as if opened.
      if (record) allowedFiles.add(path);
      await sendJson(res, 200, { ok: true, record: record || null });
      return;
    }
    case 'POST /api/recovery': {
      const body = await readJson(req, BODY_JSON_MAX);
      const path = resolve(String(body.path || ''));
      requireOpened(path);
      if (typeof body.text !== 'string' && (!Array.isArray(body.sheets) || body.sheets.length === 0)) throw httpError(400, 'no-sheet', t('没有可保留的内容'));
      const at = saveRecovery({ ...body, path });
      await sendJson(res, 200, { ok: true, at });
      return;
    }
    case 'POST /api/recovery/discard': {
      const body = await readJson(req, 64 * 1024);
      discardRecovery(resolve(String(body.path || '')));
      await sendJson(res, 200, { ok: true });
      return;
    }
    case 'POST /api/pick-open': {
      const body = await readJson(req, 64 * 1024);
      const paths = await pickFiles({ directory: resolveRoot(ctx, body.directory) });
      await sendJson(res, 200, { ok: true, paths });
      return;
    }
    case 'POST /api/create': {
      const body = await readJson(req, 64 * 1024);
      const path = resolve(requireChosen(body.path));
      const result = await createDocument(path, {});
      allowedFiles.add(path);
      await sendJson(res, 200, { ok: true, ...result });
      return;
    }
    case 'POST /api/pick-save': {
      const body = await readJson(req, 64 * 1024);
      const directory = body.directory && isAbsolute(body.directory) && existsSync(body.directory) ? body.directory : resolveRoot(ctx, body.workspaceRoot);
      let name = safeName(body.name || t('未命名'));
      // A new file's suggested name is one not taken yet, so saving it never asks to replace anything.
      if (body.fresh) name = basename(uniquePath(directory, basename(name, extname(name)), extname(name) || '.xlsx'));
      const path = await pickSaveLocation({ directory, name, message: body.message });
      if (path) {
        allowedFiles.add(resolve(path));
        chosenTargets.add(resolve(path));
      }
      await sendJson(res, 200, { ok: true, path });
      return;
    }
    case 'GET /api/versions': {
      const path = resolve(String(url.searchParams.get('path') || ''));
      if (!allowedFiles.has(path)) throw httpError(403, 'not-authorized', t('文件尚未在文档页打开。'));
      await sendJson(res, 200, { ok: true, versions: listVersions(path) });
      return;
    }
    case 'POST /api/versions/restore': {
      const body = await readJson(req, 64 * 1024);
      const path = resolve(String(body.path || ''));
      if (!allowedFiles.has(path)) throw httpError(403, 'not-authorized', t('文件尚未在文档页打开。'));
      const version = await withFileLock(path, () => restoreVersion(path, body.id));
      await sendJson(res, 200, { ok: true, version, mtime: statSync(path).mtimeMs });
      return;
    }
    case 'POST /api/stat': {
      // When each open document last changed, to notice edits made elsewhere.
      const body = await readJson(req, 256 * 1024);
      const stats = {};
      for (const raw of Array.isArray(body.paths) ? body.paths.slice(0, 100) : []) {
        const path = resolve(String(raw));
        if (!allowedFiles.has(path)) continue;
        try {
          stats[raw] = statSync(path).mtimeMs;
        } catch {
          stats[raw] = null;
        }
      }
      await sendJson(res, 200, { ok: true, stats });
      return;
    }
    case 'POST /api/reveal': {
      const body = await readJson(req, 64 * 1024);
      await reveal(String(body.path || ''));
      await sendJson(res, 200, { ok: true });
      return;
    }
    case 'GET /api/file':
      await serveFile(url, res);
      return;
    case 'GET /api/list': {
      const root = resolveRoot(ctx, url.searchParams.get('root'));
      const target = safeInside(root, url.searchParams.get('path') || '.');
      await sendJson(res, 200, { ok: true, ...listDirectory(target, root) });
      return;
    }
    case 'POST /api/import':
      await importFile(ctx, req, res, url);
      return;
    case 'POST /api/open': {
      const body = await readJson(req, BODY_JSON_MAX);
      const root = resolveRoot(ctx, body.workspaceRoot);
      const path = resolveInputPath(root, body.path);
      assertReadable(path);
      const doc = await openDocument(path, { recalc: [true, false, 'later'].includes(body.recalc) ? body.recalc : undefined });
      allowedFiles.add(path);
      await sendJson(res, 200, { ok: true, doc: { ...doc, mtime: statSync(path).mtimeMs, transient: path.startsWith(DROPPED_DIR + sep) }, workspaceRoot: root });
      return;
    }
    case 'POST /api/save-text': {
      const body = await readJson(req, BODY_JSON_MAX);
      const root = resolveRoot(ctx, body.workspaceRoot);
      const path = resolveInputPath(root, String(body.path || ''));
      requireOpened(path);
      const result = await saveTextInPlace({ path, text: body.text, openedMtime: typeof body.openedMtime === 'number' ? body.openedMtime : undefined });
      if (result.inPlace) discardRecovery(path);
      await sendJson(res, 200, { ok: true, ...result });
      return;
    }
    case 'POST /api/save': {
      const body = await readJson(req, BODY_JSON_MAX);
      const root = resolveRoot(ctx, body.workspaceRoot);
      const sourcePath = body.sourcePath ? resolveInputPath(root, body.sourcePath) : undefined;
      if (body.mode === 'in-place') {
        if (!sourcePath) throw httpError(400, 'bad-path', t('缺少原文件路径'));
        requireOpened(sourcePath);
        const result = await saveInPlace({ path: sourcePath, sheets: body.sheets, openedMtime: body.openedMtime, structure: body.structure });
        if (result.inPlace) discardRecovery(sourcePath);
        await sendJson(res, 200, { ok: true, ...result });
        return;
      }
      if (body.outputPath) requireChosen(body.outputPath);
      const result = await saveSpreadsheetCopy(ctx, {
        outputPath: body.outputPath,
        structure: body.structure,
        sheets: body.sheets,
        sourcePath,
        outputDirectory: sourcePath ? exportDirFor(sourcePath, root) : root,
        format: body.format,
        title: body.title,
        options: body.options,
      });
      allowedFiles.add(resolve(result.output));
      await sendJson(res, 200, { ok: true, ...result, workspaceRoot: root });
      return;
    }
    case 'POST /api/calculate': {
      const body = await readJson(req, BODY_JSON_MAX);
      const path = resolve(String(body.path || ''));
      requireOpened(path);
      await sendJson(res, 200, { ok: true, results: await calculateSheets({ path, sheets: body.sheets, structure: body.structure }) });
      return;
    }
    case 'POST /api/stats': {
      const body = await readJson(req, BODY_JSON_MAX);
      const root = resolveRoot(ctx, body.workspaceRoot);
      const path = resolveInputPath(root, body.path);
      assertReadable(path);
      const result = await statsForFile(path, body, exportDirFor(path, root));
      if (result.saved) allowedFiles.add(resolve(result.saved));
      await sendJson(res, 200, { ok: true, ...result, workspaceRoot: root });
      return;
    }
    case 'POST /api/convert': {
      const body = await readJson(req, BODY_JSON_MAX);
      const root = resolveRoot(ctx, body.workspaceRoot);
      const paths = (body.paths || []).map((path) => resolveInputPath(root, path));
      if (['docx', 'xlsx', 'csv'].includes(body.target) && body.acknowledgeLossy !== true && paths.some((path) => /\.pdf$/i.test(path))) {
        await sendJson(res, 200, {
          ok: true,
          requiresConfirmation: true,
          warnings: [{ code: 'lossy', message: t('PDF 转 Word/Excel/CSV 是有损转换：复杂排版、图片、表格和分栏可能丢失。确认后才会开始转换。') }],
        });
        return;
      }
      if (body.outputPath) {
        if (paths.length !== 1) throw httpError(400, 'bad-path', t('指定保存位置时只能转换一个文件'));
        requireChosen(body.outputPath);
      }
      const result = await convertMany(ctx, paths, body.target, { workspaceRoot: root, outputPath: body.outputPath });
      for (const item of result.results) if (item.output) allowedFiles.add(resolve(item.output));
      await sendJson(res, 200, { ok: true, ...result, workspaceRoot: root });
      return;
    }
    case 'POST /api/pdf': {
      const body = await readJson(req, BODY_JSON_MAX);
      const root = resolveRoot(ctx, body.workspaceRoot);
      const result = await pdfOperation(ctx, body.action, {
        paths: (body.paths || []).map((path) => resolveInputPath(root, path)),
        pages: body.pages,
        ranges: body.ranges,
        format: body.format,
        workspaceRoot: root,
        outputPath: body.outputPath ? requireChosen(body.outputPath) : undefined,
      });
      if (result.output) allowedFiles.add(resolve(result.output));
      for (const item of result.outputs || []) allowedFiles.add(resolve(item.output));
      await sendJson(res, 200, { ok: true, ...result, workspaceRoot: root });
      return;
    }
    default:
      await sendJson(res, 404, { ok: false, code: 'not-found', message: t('接口不存在') });
  }
}

/** @param {any} ctx @param {any} req */
function refusalOf(ctx, req) {
  const connection = ctx.get('connection');
  if (typeof connection?.requestRejection === 'function') return connection.requestRejection(req);
  if (!isLoopbackPeer(req.socket.remoteAddress)) return 403;
  const host = req.headers.host;
  if (host !== undefined) {
    if (!URL.canParse(`http://${host}`)) return 403;
    const name = new URL(`http://${host}`).hostname;
    if (name !== 'localhost' && name !== '[::1]' && !/^127\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(name)) return 403;
  }
  const site = req.headers['sec-fetch-site'];
  if (site !== undefined && site !== 'same-origin' && site !== 'none') return 403;
  return undefined;
}

/** @param {string|undefined} address */
function isLoopbackPeer(address) {
  if (address === undefined) return false;
  const plain = address.startsWith('::ffff:') ? address.slice('::ffff:'.length) : address;
  return plain === '::1' || /^127\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(plain);
}

/**
 * Resolve the workspace root from the request, falling back to the Host's
 * default directory. A requested root must exist and be absolute.
 * @param {any} ctx
 * @param {string|null|undefined} requested
 */
function resolveRoot(ctx, requested) {
  if (requested && typeof requested === 'string' && isAbsolute(requested) && existsSync(requested) && statSync(requested).isDirectory()) {
    return resolve(requested);
  }
  return resolve(ctx.workingDirectory?.defaultDirectory || process.cwd());
}

/** @param {string} root @param {string} path */
function safeInside(root, path) {
  const target = resolve(root, path || '.');
  if (target !== root && !target.startsWith(root + sep)) {
    throw httpError(403, 'outside-workspace', t('只能浏览当前工作区内的文件。'));
  }
  return target;
}

/** @param {string} root @param {unknown} path */
function resolveInputPath(root, path) {
  if (!path || typeof path !== 'string') throw httpError(400, 'bad-path', t('缺少文件路径'));
  return resolve(root, path);
}

/**
 * Writes go only where the user said: over a file this page opened, or to a
 * place picked in the save panel. Another page on this origin cannot direct a
 * save anywhere else.
 * @param {string} path
 */
function requireOpened(path) {
  if (!allowedFiles.has(resolve(path))) throw httpError(403, 'not-authorized', t('只能保存在文档页打开的文件。'));
  return path;
}

/** @param {unknown} path a place picked in the save panel */
function requireChosen(path) {
  if (typeof path !== 'string' || !isAbsolute(path) || !chosenTargets.has(resolve(path))) {
    throw httpError(403, 'not-authorized', t('保存位置需要在存储面板里选择。'));
  }
  return path;
}

/** @param {string} path */
function assertReadable(path) {
  if (!existsSync(path)) throw httpError(404, 'not-found', t('文件不存在：{0}', path));
  if (!statSync(path).isFile()) throw httpError(400, 'not-file', t('目标不是文件。'));
}

/**
 * @param {import('node:http').IncomingMessage} req
 * @param {number} limit
 */
function readBody(req, limit) {
  return new Promise((resolveBody, reject) => {
    const chunks = [];
    let size = 0;
    let settled = false;
    const settle = (value, error) => {
      if (settled) return;
      settled = true;
      if (error) reject(error);
      else resolveBody(value);
    };
    req.on('data', (chunk) => {
      if (settled) return;
      size += chunk.length;
      if (size > limit) {
        chunks.length = 0;
        settle(undefined, httpError(413, 'too-large', t('请求体过大，请分批处理。')));
        return;
      }
      chunks.push(Buffer.from(chunk));
    });
    req.on('end', () => settle(Buffer.concat(chunks)));
    req.on('error', (error) => settle(undefined, error));
  });
}

/** @param {import('node:http').IncomingMessage} req @param {number} limit */
async function readJson(req, limit) {
  const body = await readBody(req, limit);
  try {
    return JSON.parse(body.toString('utf8'));
  } catch {
    throw httpError(400, 'bad-json', t('请求 JSON 格式不正确。'));
  }
}

/**
 * Take a file dragged in from Finder. The page cannot learn where it came from, so it
 * is kept in the system's temporary folder and saving it asks where to put it.
 * @param {any} ctx
 * @param {import('node:http').IncomingMessage} req
 * @param {import('node:http').ServerResponse} res
 * @param {URL} url
 */
async function importFile(ctx, req, res, url) {
  const name = safeName(url.searchParams.get('name') || t('上传文件'));
  const extension = extname(name) || '.bin';
  const body = await readBody(req, BODY_UPLOAD_MAX);
  // Each dropped file gets its own folder so it keeps its own name.
  const output = uniquePath(ensureDir(join(DROPPED_DIR, Date.now().toString(36))), basename(name, extension), extension);
  await writeFile(output, body);
  allowedFiles.add(resolve(output));
  await sendJson(res, 200, { ok: true, path: output, name, size: body.byteLength, transient: true });
}

/** @param {URL} url @param {import('node:http').ServerResponse} res */
async function serveFile(url, res) {
  const path = url.searchParams.get('path');
  if (!path) throw httpError(400, 'bad-path', t('缺少 path'));
  const absolute = resolve(path);
  if (!allowedFiles.has(absolute)) {
    throw httpError(403, 'not-authorized', t('文件尚未通过文档查看器打开。'));
  }
  assertReadable(absolute);
  res.writeHead(200, {
    'content-type': mimeFor(absolute),
    'content-length': String(statSync(absolute).size),
    'cache-control': 'no-store',
    'content-disposition': `inline; filename*=UTF-8''${encodeURIComponent(basename(absolute))}`,
    // Documents are data: nothing served here may run as a page on this origin.
    'x-content-type-options': 'nosniff',
    'content-security-policy': "sandbox; default-src 'none'",
  });
  pipeFile(absolute, res);
}

/**
 * Stream a file. A read error (the file deleted meanwhile) must end this
 * response, not reach the process as an unhandled error event.
 * @param {string} path
 * @param {import('node:http').ServerResponse} res
 */
function pipeFile(path, res) {
  const stream = createReadStream(path);
  stream.on('error', () => res.destroy());
  stream.pipe(res);
}

/**
 * Show a file in Finder. Only files the viewer opened or made can be revealed.
 * @param {string} path
 */
function reveal(path) {
  const absolute = resolve(path);
  if (!allowedFiles.has(absolute) || !existsSync(absolute)) throw httpError(403, 'not-authorized', t('只能在访达中显示文档页打开或生成的文件。'));
  const [command, args] = process.platform === 'darwin' ? ['open', ['-R', absolute]]
    : process.platform === 'win32' ? ['explorer.exe', [`/select,${absolute}`]]
    : ['xdg-open', [dirname(absolute)]];
  return new Promise((resolveReveal, reject) => {
    const child = spawn(command, args, { stdio: 'ignore', detached: true });
    child.once('error', (error) => reject(httpError(500, 'reveal-failed', t('无法打开文件夹：{0}', error.message))));
    child.once('spawn', () => {
      child.unref();
      resolveReveal();
    });
  });
}

/** Where pnpm placed PDF.js: hoisted to the profile, or nested under this package. */
let pdfjsRoot;
function pdfjsDirectory() {
  pdfjsRoot ??= dirname(require.resolve('pdfjs-dist/package.json'));
  return pdfjsRoot;
}

const ASSET_DIRECTORIES = [
  { directory: 'build', prefix: '' },
  { directory: 'standard_fonts', prefix: 'standard_fonts/' },
  { directory: 'cmaps', prefix: 'cmaps/' },
  { directory: 'wasm', prefix: 'wasm/' },
  { directory: 'iccs', prefix: 'iccs/' },
  { directory: 'image_decoders', prefix: 'image_decoders/' },
];

/**
 * Serve the PDF.js worker and standard fonts from the installed package.
 * Only known subpaths are exposed.
 * @param {string} relative
 * @param {import('node:http').ServerResponse} res
 */
async function serveAsset(relative, res) {
  const safe = normalize(relative).replace(/^([.]{2}[\\/])+/, '');
  let found;
  for (const { directory, prefix } of ASSET_DIRECTORIES) {
    if (prefix && !safe.startsWith(prefix)) continue;
    const root = join(pdfjsDirectory(), directory);
    const candidate = resolve(root, safe.slice(prefix.length));
    if (candidate.startsWith(root + sep) && existsSync(candidate) && statSync(candidate).isFile()) {
      found = candidate;
      break;
    }
  }
  if (!found) {
    await sendJson(res, 404, { ok: false, code: 'asset-not-found', message: t('资源不存在') });
    return;
  }
  res.writeHead(200, {
    'content-type': mimeFor(found),
    'content-length': String(statSync(found).size),
    'cache-control': 'public, max-age=3600',
  });
  pipeFile(found, res);
}

/**
 * @param {import('node:http').ServerResponse} res
 * @param {number} status
 * @param {unknown} payload
 */
async function sendJson(res, status, payload) {
  const raw = Buffer.from(JSON.stringify(payload));
  const accept = String(res.req?.headers?.['accept-encoding'] || '');
  const useGzip = raw.byteLength > 1024 && accept.includes('gzip');
  // A large workbook compresses for a while; the Host's other sessions must keep running.
  const body = useGzip ? await gzipAsync(raw, { level: 1 }) : raw;
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'content-length': String(body.byteLength),
    'cache-control': 'no-store',
    ...(useGzip ? { 'content-encoding': 'gzip', vary: 'Accept-Encoding' } : {}),
  });
  res.end(body);
}

/** @param {string} directory @param {string} root */
function listDirectory(directory, root) {
  if (!existsSync(directory) || !statSync(directory).isDirectory()) {
    throw httpError(404, 'not-found', t('目录不存在'));
  }
  const entries = [];
  try {
    for (const entry of readdirSync(directory, { withFileTypes: true })) {
      if (entry.name.startsWith('.')) continue;
      const absolute = join(directory, entry.name);
      if (entry.isDirectory()) {
        if (entry.name !== 'node_modules') entries.push({ name: entry.name, type: 'directory', path: absolute });
      } else if (entry.isFile() && SUPPORTED_EXTENSIONS.has(extname(entry.name).toLowerCase())) {
        entries.push({ name: entry.name, type: 'file', path: absolute, size: statSync(absolute).size, ext: extname(entry.name).toLowerCase() });
      }
    }
  } catch (error) {
    throw httpError(500, 'list-failed', t('无法列出目录：{0}', error instanceof Error ? error.message : String(error)));
  }
  entries.sort((a, b) => (a.type === b.type ? a.name.localeCompare(b.name, 'zh-CN') : a.type === 'directory' ? -1 : 1));
  return {
    directory,
    root,
    atRoot: directory === root,
    parent: directory === root ? undefined : dirname(directory),
    entries: entries.slice(0, MAX_LISTED_ENTRIES),
    truncated: entries.length > MAX_LISTED_ENTRIES,
    relative: directory === root ? '.' : directory.slice(root.length + 1),
  };
}
