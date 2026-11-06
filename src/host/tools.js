/**
 * Agent-facing tools. Each tool calls exactly the same document engine as the
 * panel; there is no second implementation for the model.
 */
import { basename, resolve } from 'node:path';
import { t } from './i18n.js';
import { XLSX } from './spreadsheet.js';
import { exportDirFor } from './util.js';
import { createDocument, openDocument, readDocument, editFile, statsForFile, convertMany, pdfOperation } from './documents.js';
import { enqueueOpen, getPanelState, pageListening } from './requests.js';

const FALLBACK_READ_CHARS = 12_000;
const MIN_READ_CHARS = 1_200;
const MAX_READ_CHARS = 60_000;
/** One read may fill this share of what is left of the model's window after its reply. */
const READ_WINDOW_SHARE = 0.15;
/** A conservative figure for mixed Chinese and tabular text. */
const CHARS_PER_TOKEN = 2;
const MAX_GROUPS_SHOWN = 30;

/** @param {string} path @param {string} root */
function resolvePath(path, root) {
  if (!path) throw new Error(t('缺少文件路径'));
  return resolve(root, path);
}

/**
 * The file a tool works on: the given path, or the document the user has in front
 * of them in the Documents (文档) page, so "求和选中的这一列" needs no file name.
 * @param {string|undefined} path
 * @param {string} root
 */
function targetPath(path, root) {
  if (path) return resolve(root, path);
  const active = getPanelState()?.active?.path;
  if (!active) throw new Error(t('用户现在没有在“文档”页打开文件，请提供 path。'));
  return active;
}

/** A warning when the user has edits in the page that are not in the file yet. @param {string} path */
function unsavedNote(path) {
  const tab = getPanelState()?.tabs?.find((item) => item.path === path);
  return tab?.dirty ? t('\n注意：用户在“文档”页里对这个文件有未保存的修改，上面的结果不包含这些修改；需要的话先请用户按 ⌘S 保存。') : '';
}

/** @param {any} active */
function describeActive(active) {
  if (!active) return t('“文档”页里没有打开的文件。');
  const parts = [t('当前文档：{0}（{1}）', active.path, active.kind)];
  if (active.sheet) parts.push(active.sheets?.length ? t('工作表：{0}（全部：{1}）', active.sheet, active.sheets.join(t('、'))) : t('工作表：{0}', active.sheet));
  if (active.selection) parts.push(t('选中：{0}', active.selection));
  if (active.page) parts.push(t('正在看第 {0} 页，共 {1} 页', active.page, active.pageCount));
  if (active.dirty) parts.push(t('有未保存的修改'));
  return parts.join('\n');
}

/** @param {string} text */
function textContent(text) {
  return [{ type: 'text', text: String(text) }];
}

/** Strip undefined and non-finite values so the Host accepts the value as lossless JSON. */
function jsonSafe(value) {
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return { ok: false, code: 'serialization', message: t('结果无法序列化') };
  }
}

/**
 * The model sees `message`, never a JSON dump of the structured result: small
 * models cannot afford the punctuation, and the whole text must survive intact.
 */
function resultOf() {
  return {
    schema: {
      type: 'object',
      properties: {
        ok: { type: 'boolean' },
        message: { type: 'string' },
        output: { type: 'string' },
        warnings: { type: 'array', items: { type: 'object' } },
      },
      additionalProperties: true,
    },
    render(_args, result) {
      const value = result?.value || result || {};
      return textContent(value.message ?? JSON.stringify(value));
    },
  };
}

/** @param {unknown} warnings */
function warningText(warnings) {
  if (!Array.isArray(warnings) || warnings.length === 0) return '';
  return `\n${t('提示：')}\n${warnings.map((warning) => `- ${warning.message || warning.code}`).join('\n')}`;
}

/**
 * How much text one read may return, from the running model's window rather
 * than any particular model's numbers. Without model information it falls back
 * to a size that is safe for small windows.
 * @param {any} ctx
 * @param {any} agent
 * @param {AbortSignal} [signal]
 */
async function readBudgetChars(ctx, agent, signal) {
  try {
    const llm = ctx.get?.('llm');
    const routed = agent?.session?.requestHeader?.()?.config;
    const provider = routed?.provider || agent?.options?.provider;
    const model = routed?.model || agent?.options?.model;
    if (!llm?.resolveModelInfo || !provider || !model) return FALLBACK_READ_CHARS;
    const info = await llm.resolveModelInfo(provider, model, signal);
    const window = Number(info?.context?.contextWindow);
    if (!Number.isFinite(window) || window <= 0) return FALLBACK_READ_CHARS;
    const usable = Math.max(window - (routed?.maxTokens ?? info?.defaultMaxTokens ?? 0), window * 0.25);
    return Math.round(Math.min(MAX_READ_CHARS, Math.max(MIN_READ_CHARS, usable * READ_WINDOW_SHARE * CHARS_PER_TOKEN)));
  } catch {
    return FALLBACK_READ_CHARS;
  }
}

/** Render a read result as text with a clear way to continue. */
function readText(value) {
  if (value.kind === 'spreadsheet') {
    const header = ['', ...value.rows[0]?.map((_, index) => XLSX.utils.encode_col(value.firstColumn + index)) ?? []].join(' | ');
    const lines = value.rows.map((row, index) => [
      value.firstRow + index + 1,
      ...row.map((cell) => {
        if (!cell) return '';
        return cell.formula ? `${cell.formula} → ${cell.display}` : cell.display;
      }),
    ].join(' | '));
    const parts = [
      t('工作表「{0}」，本次范围 {1}，已用范围 {2}。全部工作表：{3}', value.sheet, value.range, value.usedRange, value.sheets.join(t('、'))),
      header,
      ...lines,
    ];
    if (value.nextRange) parts.push(t('（还有更多行，继续读取请传 range="{0}"）', value.nextRange));
    else if (value.truncated) parts.push(t('（列数超出单次上限，请缩小 range）'));
    return parts.join('\n') + warningText(value.warnings);
  }
  if (value.kind === 'pdf') {
    const body = value.pages.map((page) => t('--- 第 {0} 页 ---\n{1}{2}', page.page, page.empty ? t('（本页没有可提取文字）') : page.text, page.partial ? t('\n…（本页已截断）') : '')).join('\n\n');
    const notes = [t('共 {0} 页。', value.pageCount)];
    if (value.nextPages) notes.push(t('其余页继续读取请传 pages="{0}"。', value.nextPages));
    if (value.scanned) notes.push(t('这些页没有文字层，可能是扫描件，可用 doc_viewer_pdf 的 ocr 操作识别。'));
    return `${body}\n\n${notes.join('')}`;
  }
  const more = value.nextOffset !== undefined
    ? t('\n\n（共 {0} 字符，已读到 {1}，继续读取请传 offset={2}）', value.totalChars, value.nextOffset, value.nextOffset)
    : '';
  return `${value.text}${more}`;
}

/** @param {{ name: string, rows: number, columns: number }[]|undefined} sizes */
function sheetSummary(sizes) {
  return (sizes || []).map((sheet) => t('{0}（{1} 行 × {2} 列）', sheet.name, sheet.rows, sheet.columns)).join(t('、'));
}

/**
 * Register the persistent tools on the Host tool registry.
 * @param {any} ctx
 */
/** Listed one by one up to here; beyond it, the start and the end, where totals usually are. */
const LISTED_DEPENDENTS = { head: 20, tail: 10 };
const LISTED_WRITTEN = 30;

/**
 * Lines telling the agent what an edit changed.
 * @param {{ written: any[], dependents: any[], uncalculated?: number }} changes
 */
export function describeChanges({ written, dependents, uncalculated = 0 }) {
  const line = (item) => (item.formula
    ? t('- {0}!{1}（{2}）：{3} → {4}', item.sheet, item.address, item.formula, item.before || t('空'), item.after || t('空'))
    : t('- {0}!{1}：{2} → {3}', item.sheet, item.address, item.before || t('空'), item.after || t('空')));
  const lines = written.slice(0, LISTED_WRITTEN).map(line);
  if (written.length > LISTED_WRITTEN) lines.push(t('- 另有 {0} 个写入的单元格', written.length - LISTED_WRITTEN));
  if (dependents.length) {
    const { head, tail } = LISTED_DEPENDENTS;
    lines.push(t('公式结果随之变化（共 {0} 处）：', dependents.length));
    if (dependents.length <= head + tail) lines.push(...dependents.map(line));
    else lines.push(...dependents.slice(0, head).map(line), t('- ……中间另有 {0} 处……', dependents.length - head - tail), ...dependents.slice(-tail).map(line));
  }
  if (uncalculated) lines.push(t('另有 {0} 处公式这次没能算出结果（LibreOffice 超时或不可用）。公式本身还在，用户打开文件时会重新计算；现在不能据此回答它们的数值。', uncalculated));
  return lines;
}

export function registerTools(ctx) {
  ctx.effect(() => ctx.tools.register({
    name: 'doc_viewer_context',
    description: 'What the user has open in the Documents (文档) page right now: the document in front of them, its sheet and selected range or PDF page, other open documents, and unsaved edits. Call this first when the user says "这个表", "选中的", "这一页" or otherwise refers to a document without naming it. The other doc_viewer tools use this document when path is omitted.',
    parameters: { type: 'object', properties: {}, additionalProperties: false },
    output: resultOf(),
    async execute() {
      const state = getPanelState();
      if (!state) return jsonSafe({ ok: true, open: false, message: t('“文档”页还没有打开过文件。') });
      const others = state.tabs.filter((tab) => tab.path !== state.active?.path).map((tab) => `- ${tab.path}${tab.dirty ? t('（未保存）') : ''}`);
      const message = [
        describeActive(state.active),
        state.visible ? '' : t('（用户此刻不在“文档”页，在看对话）'),
        others.length ? t('其他打开的文档：\n{0}', others.join('\n')) : '',
      ].filter(Boolean).join('\n');
      return jsonSafe({ ok: true, ...state, message });
    },
    presentResult(_args, result) {
      return { card: 'generic', title: t('当前文档'), content: textContent(result?.value?.message || '') };
    },
  }), 'doc-viewer: tool context');

  ctx.effect(() => ctx.tools.register({
    name: 'doc_viewer_create',
    description: 'Make a new document and show it to the user beside the conversation: an .xlsx workbook (empty, or with rows; a string starting with = is a formula, and the first row is a bold, frozen header unless header is false), a .csv from rows, or a .md/.txt file with content. Relative paths resolve against the current workspace. An existing file is never replaced; to change one, use doc_viewer_edit.',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'New file path ending in .xlsx, .csv, .md or .txt.' },
        rows: { type: 'array', items: { type: 'array', items: {} }, description: 'xlsx/csv: rows of cell values, first row first.' },
        header: { type: 'boolean', description: 'xlsx: treat the first row as a header (default true).' },
        sheet: { type: 'string', description: 'xlsx: worksheet name.' },
        content: { type: 'string', description: 'md/txt: the text.' },
        show: { type: 'boolean', description: 'Show it to the user (default true).' },
      },
      required: ['path'],
      additionalProperties: false,
    },
    output: resultOf(),
    async execute(args, exec) {
      const root = await ctx.workingDirectory.ensure(exec.agent);
      const path = resolvePath(args.path, root);
      const value = await createDocument(path, args);
      enqueueOpen(path, { sessionId: exec.agent?.id, reveal: args.show !== false });
      return jsonSafe({ ok: true, ...value, message: (value.rows ? t('已新建 {0}，{1} 行。', path, value.rows) : t('已新建 {0}。', path)) + warningText(value.warnings) });
    },
    presentResult(_args, result) {
      return { card: 'generic', title: t('已新建文件'), content: textContent(result?.value?.message || '') };
    },
  }), 'doc-viewer: tool create');

  ctx.effect(() => ctx.tools.register({
    name: 'doc_viewer_open',
    description: 'Show a local xlsx/xls/csv/pdf/docx/txt/et/wps file to the user in the Documents (文档) page (it switches the window to it) and return its structure. Relative paths resolve against the current workspace. Use this only when the user wants to see a document; use doc_viewer_read to look at contents yourself.',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'File path, relative to the current workspace or absolute.' },
      },
      required: ['path'],
      additionalProperties: false,
    },
    output: resultOf(),
    async execute(args, exec) {
      const root = await ctx.workingDirectory.ensure(exec.agent);
      const path = resolvePath(args.path, root);
      const doc = await openDocument(path);
      enqueueOpen(path, { sessionId: exec.agent?.id, reveal: true });
      if (doc.kind === 'spreadsheet') {
        return jsonSafe({
          ok: true,
          kind: doc.kind,
          path,
          name: doc.name,
          sheets: doc.sheetSizes,
          stats: doc.stats,
          warnings: doc.warnings,
          message: t('已在“文档”页打开「{0}」。\n工作表：{1}\n非空单元格 {2} 个，公式 {3} 个{4}。', doc.name, sheetSummary(doc.sheetSizes), doc.stats.cells, doc.stats.formulas,
            doc.stats.formulas ? t('（{0}）', doc.stats.recalculated ? t('已重新计算') : t('使用文件里保存的计算结果')) : '') + warningText(doc.warnings),
        });
      }
      if (doc.kind === 'pdf') {
        return jsonSafe({
          ok: true,
          kind: doc.kind,
          path,
          name: doc.name,
          pageCount: doc.pageCount,
          message: t('已在“文档”页打开「{0}」，共 {1} 页。', doc.name, doc.pageCount),
        });
      }
      return jsonSafe({
        ok: true,
        kind: doc.kind,
        path,
        name: doc.name,
        chars: doc.stats.chars,
        encoding: doc.encoding,
        warnings: doc.warnings,
        message: t('已在“文档”页打开「{0}」，约 {1} 字符。{2}', doc.name, doc.stats.chars, warningText(doc.warnings)),
      });
    },
    presentResult(_args, result) {
      return { card: 'generic', title: t('文档查看器'), content: textContent(result?.value?.message || t('已打开文档')) };
    },
  }), 'doc-viewer: tool open');

  ctx.effect(() => ctx.tools.register({
    name: 'doc_viewer_read',
    description: 'Read part of a local document without showing it to the user: a sheet range with values and formulas, PDF pages, or Word/text by character offset. Without path it reads the document the user has open in the Documents (文档) page. Long results are cut at a natural boundary and tell you what to pass next.',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'File path; omit for the document the user has open.' },
        sheet: { type: 'string', description: 'Worksheet name; defaults to the first sheet.' },
        range: { type: 'string', description: 'A1 range such as A1:F40; defaults to the top of the sheet.' },
        pages: { type: 'string', description: 'PDF pages such as 1,3-5; defaults to page 1.' },
        offset: { type: 'number', description: 'Word/text: character offset to continue from.' },
      },
      additionalProperties: false,
    },
    output: resultOf(),
    async execute(args, exec) {
      const root = await ctx.workingDirectory.ensure(exec.agent);
      const limit = await readBudgetChars(ctx, exec.agent, exec.signal);
      const path = targetPath(args.path, root);
      const value = await readDocument(path, { ...args, limit });
      const message = `${args.path ? '' : `${t('（{0}）', path).trim()}\n`}${readText(value)}${unsavedNote(path)}`;
      // The rows are in the text already; keeping them too would double the session log.
      return jsonSafe({ ok: true, ...value, rows: undefined, text: undefined, pages: value.kind === 'pdf' ? value.pages.map((page) => page.page) : undefined, message });
    },
    presentResult(_args, result) {
      const message = String(result?.value?.message || '');
      return { card: 'generic', title: t('文档内容'), content: textContent(message.length > 6000 ? `${message.slice(0, 6000)}\n…` : message) };
    },
  }), 'doc-viewer: tool read');

  ctx.effect(() => ctx.tools.register({
    name: 'doc_viewer_edit',
    description: 'Write cells or formulas in an xlsx/csv workbook and return the recalculated values of what was written. The file is changed where it is, and its previous content is kept as a version the user can restore in one click. When the user has the file open in the Documents (文档) page with unsaved edits of their own, the edits go into that page instead, among those, and reach the file with the next save the user makes. When writing in place would lose something (charts, pivot tables, macros, conditional formats, old .xls), a copy is written beside it and the result says why. Without path it edits the document the user has open; an open page updates where it is, without taking the user away from the conversation, and lists the change so it can be undone. convertTextNumbers turns numeric-looking text into real numbers so formulas work (identifiers such as 00123 stay text). Formulas use Excel syntax and may reference other sheets, e.g. =SUM(数据!B2:B13). A call with a malformed formula, an unknown function or a write inside a merged cell changes nothing and says what to fix. The result lists each written cell before and after, and the formula results that changed because of it (totals included); those are what the file now holds, so there is no need to check them another way.',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Source workbook path; omit for the document the user has open.' },
        sheet: { type: 'string', description: 'Default worksheet for edits that do not name one.' },
        edits: {
          type: 'array',
          description: 'Cells to write. A formula replaces the value; value sets a literal; null clears.',
          items: {
            type: 'object',
            properties: {
              sheet: { type: 'string', description: 'Worksheet name; defaults to the first sheet.' },
              address: { type: 'string', description: 'A1 address.' },
              value: { description: 'Literal value (string, number or boolean).' },
              formula: { type: 'string', description: 'Formula with or without a leading =.' },
            },
            required: ['address'],
            additionalProperties: false,
          },
        },
        convertTextNumbers: { type: 'boolean', description: 'Convert numeric-looking text cells to real numbers.' },
        inPlace: { type: 'boolean', description: 'Set false to always write a separate copy beside the file.' },
        saveAs: { type: 'string', enum: ['xlsx', 'xls'], description: 'Format of a separate copy. xls is BIFF8 for old systems.' },
        show: { type: 'boolean', description: 'Also switch the window to the Documents (文档) page. Only when the user asked to see it; default false.' },
      },
      additionalProperties: false,
    },
    output: resultOf(),
    async execute(args, exec) {
      const root = await ctx.workingDirectory.ensure(exec.agent);
      const source = targetPath(args.path, root);
      const dirty = Boolean(getPanelState()?.tabs?.find((tab) => tab.path === source)?.dirty);
      // Edits join the user's unsaved ones only while a page is there to take them.
      const live = dirty && pageListening();
      const value = await editFile(source, {
        ...args,
        live,
        blockedReason: dirty && !live ? t('用户在文档页里对这个文件有未保存的修改，而文档页现在没有打开') : undefined,
      }, exportDirFor(source, root));
      const note = value.converted ? t('已把 {0} 个文本数字转换为数字', value.converted) : undefined;
      if (value.live || value.inPlace) {
        enqueueOpen(source, { kind: 'agent-edit', sessionId: exec.agent?.id, reveal: args.show === true, live: value.live, cells: value.cells, edited: value.applied, mtime: value.mtime, note });
      } else {
        enqueueOpen(value.output, { sessionId: exec.agent?.id, reveal: args.show === true });
      }
      const written = value.changes ? describeChanges(value.changes) : (value.values || []).slice(0, 30).map((item) => `- ${item.sheet}!${item.address} = ${item.display}${item.formula ? t('（{0}）', item.formula) : ''}`);
      return jsonSafe({
        ok: true,
        ...value,
        // The page has them; the session log does not need them twice.
        cells: undefined,
        // Spelled out in the message.
        changes: undefined,
        message: [
          value.live
            ? t('用户在文档页里对 {0} 有未保存的修改，所以改动已放进用户正在编辑的表格，和用户的修改一起在下次保存时写入文件；在那之前读文件看不到这些改动。', value.output)
            : value.inPlace
              ? t('已直接修改 {0}。修改前的版本已备份，用户可在文档页撤销。', value.output)
              : t('没有覆盖原文件（{0}），改动保存在：{1}', value.keptAside, value.output),
          value.converted ? t('已把 {0} 个文本数字转换为数字。', value.converted) : '',
          written.length ? `${value.changes ? t('改动（改前 → 改后）：') : t('写入结果：')}\n${written.join('\n')}` : '',
        ].filter(Boolean).join('\n') + warningText(value.warnings),
      });
    },
    presentResult(_args, result) {
      return { card: 'generic', title: result?.value?.live ? t('已放进文档页') : result?.value?.inPlace ? t('已修改文件') : t('已保存副本'), content: textContent(result?.value?.message || '') };
    },
  }), 'doc-viewer: tool edit');

  ctx.effect(() => ctx.tools.register({
    name: 'doc_viewer_stats',
    description: 'Compute sum/average/min/max/count/distinct count for a sheet range or one column, or group by one column and aggregate another. Without path it uses the document the user has open; pass the user\'s selection from doc_viewer_context as range when they mean "the selected cells". Can save the result as a new xlsx with a 统计结果 sheet.',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Workbook path; omit for the document the user has open.' },
        sheet: { type: 'string', description: 'Worksheet name.' },
        range: { type: 'string', description: 'A1 range such as A1:D100; defaults to the used range.' },
        groupBy: { type: 'string', description: 'Header or column letter to group by. Needs valueColumn.' },
        valueColumn: { type: 'string', description: 'Header or column letter to aggregate.' },
        agg: { type: 'string', enum: ['sum', 'average', 'count', 'max', 'min'], description: 'Aggregation for grouped results.' },
        hasHeader: { type: 'boolean', description: 'Whether the first selected row is headers and left out of the numbers. Decided from its content when omitted: a selection of data counts every cell.' },
        save: { type: 'boolean', description: 'Save a new xlsx containing the statistics sheet.' },
      },
      additionalProperties: false,
    },
    output: resultOf(),
    async execute(args, exec) {
      const root = await ctx.workingDirectory.ensure(exec.agent);
      const source = targetPath(args.path, root);
      const value = await statsForFile(source, args, exportDirFor(source, root));
      if (value.saved) enqueueOpen(value.saved, { sessionId: exec.agent?.id, reveal: true });
      const summary = value.summary;
      const lines = [
        t('统计范围：{0}!{1}', summary.sheet, summary.range),
        t('合计：{0}｜平均：{1}｜最大：{2}｜最小：{3}', summary.sum, summary.average ?? '—', summary.max ?? '—', summary.min ?? '—'),
        t('计数：{0}（其中数字 {1}）｜去重计数：{2}', summary.count, summary.numericCount, summary.distinct),
        value.hasHeader ? t('首行作为表头，没有计入；要计入请传 hasHeader:false。') : '',
      ].filter(Boolean);
      if (value.groups?.length) {
        const agg = value.groupBy?.agg;
        lines.push(t('分组（{0} → {1}，{2}）：', value.groupBy?.column, value.groupBy?.valueColumn, agg));
        for (const group of value.groups.slice(0, MAX_GROUPS_SHOWN)) lines.push(t('- {0}: {1}（{2} 条）', group.key, group[agg], group.count));
        if (value.groups.length > MAX_GROUPS_SHOWN) lines.push(t('…另有 {0} 组未列出。', value.groups.length - MAX_GROUPS_SHOWN));
      }
      if (value.saved) lines.push(t('已另存为：{0}', value.saved));
      return jsonSafe({ ok: true, ...value, groups: value.groups?.slice(0, MAX_GROUPS_SHOWN), message: lines.join('\n') + warningText(value.warnings) + unsavedNote(source) });
    },
    presentResult(_args, result) {
      return { card: 'generic', title: t('统计结果'), content: textContent(result?.value?.message || '') };
    },
  }), 'doc-viewer: tool stats');

  ctx.effect(() => ctx.tools.register({
    name: 'doc_viewer_convert',
    description: 'Batch-convert local documents; each output is saved beside its source. Targets: pdf, xlsx, xls, csv, docx, txt. Office→PDF embeds Chinese fonts through the Harness LibreOffice provider. PDF→Word/XLSX/CSV is lossy; call without acknowledgeLossy first to receive the warning, then call again with acknowledgeLossy:true after telling the user.',
    parameters: {
      type: 'object',
      properties: {
        paths: { type: 'array', items: { type: 'string' }, description: 'Source files (one or more).' },
        target: { type: 'string', enum: ['pdf', 'xlsx', 'xls', 'csv', 'docx', 'txt'], description: 'Output format.' },
        acknowledgeLossy: { type: 'boolean', description: 'Set true only after the user accepted PDF→Word/XLSX/CSV layout loss.' },
      },
      required: ['paths', 'target'],
      additionalProperties: false,
    },
    output: resultOf(),
    async execute(args, exec) {
      const root = await ctx.workingDirectory.ensure(exec.agent);
      const paths = (args.paths || []).map((path) => resolvePath(path, root));
      if (['docx', 'xlsx', 'csv'].includes(args.target) && args.acknowledgeLossy !== true && paths.some((path) => /\.pdf$/i.test(path))) {
        return jsonSafe({
          ok: true,
          requiresConfirmation: true,
          message: t('PDF 转 Word/Excel/CSV 是有损转换：复杂排版、图片、表格和分栏可能丢失。告知用户并确认后，再次调用并传 acknowledgeLossy:true。'),
        });
      }
      const value = await convertMany(ctx, paths, args.target, { workspaceRoot: root });
      for (const result of value.results) if (result.ok && result.output) enqueueOpen(result.output, { sessionId: exec.agent?.id });
      const lines = value.results.map((result) => (result.ok
        ? `✓ ${basename(result.source)} → ${result.output}${result.warnings?.length ? t('（{0}）', result.warnings.map((warning) => warning.message).join(t('；'))) : ''}`
        : t('✗ {0}：{1}', basename(result.source), result.message)));
      return jsonSafe({ ok: value.results.some((result) => result.ok), ...value, message: lines.join('\n') });
    },
    presentResult(_args, result) {
      return { card: 'generic', title: t('转换结果'), content: textContent(result?.value?.message || '') };
    },
  }), 'doc-viewer: tool convert');

  ctx.effect(() => ctx.tools.register({
    name: 'doc_viewer_pdf',
    description: 'PDF tools: merge several PDFs, split by page ranges, delete pages, or OCR a scanned PDF/image with macOS Vision (Apple silicon). Outputs are new files beside the first source.',
    parameters: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['merge', 'split', 'delete', 'ocr'], description: 'PDF operation.' },
        paths: { type: 'array', items: { type: 'string' }, description: 'Source PDF paths (one for split/delete/ocr, two or more for merge).' },
        pages: { type: 'string', description: 'delete: pages to remove, e.g. 2,4-6. ocr: pages to recognize; defaults to all, at most 80 per call.' },
        ranges: { type: 'string', description: 'split: ranges, e.g. 1-3,4-6. Default splits every page.' },
        format: { type: 'string', enum: ['txt', 'docx'], description: 'OCR output format.' },
      },
      required: ['action', 'paths'],
      additionalProperties: false,
    },
    output: resultOf(),
    async execute(args, exec) {
      const root = await ctx.workingDirectory.ensure(exec.agent);
      const value = await pdfOperation(ctx, args.action, {
        paths: (args.paths || []).map((path) => resolvePath(path, root)),
        pages: args.pages,
        ranges: args.ranges,
        format: args.format,
        workspaceRoot: root,
      });
      for (const item of value.outputs || []) enqueueOpen(item.output, { sessionId: exec.agent?.id });
      if (value.output) enqueueOpen(value.output, { sessionId: exec.agent?.id });
      const lines = [];
      if (value.outputs) for (const item of value.outputs) lines.push(t('✓ {0}（第 {1} 页）', item.output, item.pages.join(',')));
      else lines.push(`✓ ${value.output}`);
      if (args.action === 'merge') lines.push(t('共 {0} 页', value.pages));
      if (args.action === 'delete') lines.push(t('删除 {0} 页，剩余 {1} 页', value.removed, value.remaining));
      if (args.action === 'ocr') lines.push(t('识别 {0} 页，共 {1} 字符', value.pages, value.chars));
      return jsonSafe({ ok: true, ...value, message: lines.join('\n') });
    },
    presentResult(_args, result) {
      return { card: 'generic', title: t('PDF 处理结果'), content: textContent(result?.value?.message || '') };
    },
  }), 'doc-viewer: tool pdf');
}
