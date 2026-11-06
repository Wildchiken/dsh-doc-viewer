import React from 'react';
import { lang, t } from './i18n.js';
import { Workbook } from '@fortune-sheet/react';
import { calculate, open as openOnHost, restoreVersion } from './api.js';
import { Icon } from './icons.jsx';
import { FindBar, Menu, Segmented, Sheet, TabActions } from './ui.jsx';
import {
  aggLabels,
  buildStatsSheet,
  colName,
  computeSheetStats,
  displayOf,
  firstRowIsHeader,
  formatNumber,
  parseNumericText,
  renderChartSvg,
  sheetsForSaving,
  svgToPngDataUrl,
  uniqueSheetName,
  usedBounds,
} from './sheetUtils.js';

/** What the grid's own toolbar offers: the everyday set, without the rarely used half. */
const TOOLBAR_ITEMS = [
  'undo', 'redo', 'format-painter', 'clear-format', '|',
  'percentage-format', 'number-decrease', 'number-increase', 'format', '|',
  'font-size', 'bold', 'italic', 'strike-through', 'underline', 'font-color', 'background', '|',
  'border', 'merge-cell', 'horizontal-align', 'vertical-align', 'text-wrap', '|',
  'filter', 'quick-formula',
];
const MAX_MATCHES = 1000;
/** Right-click menus, without what a save could not keep (pictures, links, charts, validation) or keep in place (deleting single cells). */
const CELL_MENU = ['copy', 'paste', '|', 'insert-row', 'insert-column', 'delete-row', 'delete-column', 'hide-row', 'hide-column', 'set-row-height', 'set-column-width', '|', 'clear', 'sort', 'orderAZ', 'orderZA', 'filter', 'cell-format'];
const HEADER_MENU = ['copy', 'paste', '|', 'insert-row', 'insert-column', 'delete-row', 'delete-column', 'hide-row', 'hide-column', 'set-row-height', 'set-column-width', '|', 'clear', 'sort', 'orderAZ', 'orderZA'];
const SHEET_TAB_MENU = ['delete', 'copy', 'rename', '|', 'move'];
/** What the grid's engine answers for a function it does not have. */
const UNKNOWN_FUNCTION = /^#(ERROR!|NAME\?)/;
/** Past this many cells a round trip for every edit would get in the way of typing. */
const MAX_CALCULATED_CELLS = 200_000;
const CALCULATE_DELAY_MS = 600;
/**
 * Agent marks stay until the user clicks into the sheet or types, so they are still there
 * after reading the agent's reply; at most this long while the sheet is in view.
 */
const MARK_MAX_MS = 60_000;
const MARK_FADE_MS = 1200;
/** A click already on its way when the marks appear is not the user dismissing them. */
const MARK_SETTLE_MS = 600;
const MAX_MARKS = 200;
const MARK_COLOR = '#4d6bfe';

/**
 * One editable spreadsheet. FortuneSheet provides the grid, the formula engine,
 * multiple sheets, merged cells, frozen panes, filters, sorting and undo.
 */
export function SpreadsheetTab({ tab, actionsEl, pageActive, selected, setDirty, patchTab, saveSheets, saveSnapshot, registerCommands, onContextChange, onVersions, toast, toolItems }) {
  const workbook = React.useRef(null);
  const [findOpen, setFindOpen] = React.useState(false);
  const [statsOpen, setStatsOpen] = React.useState(false);
  // The grid takes its data once; later changes to the tab must not reset it.
  const [initialSheets] = React.useState(tab.sheets);
  const tabRef = React.useRef(tab);
  tabRef.current = tab;

  // Rows and columns inserted or deleted since the file was last written; a save replays them on the original.
  // A document restored from unsaved edits starts with the ones made before.
  const structure = React.useRef(tab.structure ? [...tab.structure] : []);
  const saving = React.useRef(null);
  const settling = React.useRef(false);
  /** What agents changed in this document, newest last, until the user clears the list. */
  const [agentChanges, setAgentChanges] = React.useState([]);
  const agentChangesRef = React.useRef(agentChanges);
  agentChangesRef.current = agentChanges;

  // Unsaved edits are snapshotted a moment after the user stops, so quitting or a crash loses little.
  const snapshotTimer = React.useRef(null);
  const snapshotNow = React.useCallback(() => {
    clearTimeout(snapshotTimer.current);
    snapshotTimer.current = null;
    const api = workbook.current;
    if (!api || !tabRef.current.dirty || saving.current) return;
    saveSnapshot?.(tabRef.current, { sheets: sheetsForSaving(api.getAllSheets()), structure: structure.current.slice() });
  }, [saveSnapshot]);
  const scheduleSnapshot = React.useCallback(() => {
    clearTimeout(snapshotTimer.current);
    // A large workbook takes a while to serialize; snapshot it less eagerly.
    snapshotTimer.current = setTimeout(snapshotNow, (tabRef.current.doc?.stats?.cells || 0) > 200_000 ? 10_000 : 2_000);
  }, [snapshotNow]);
  // Leaving the page or this tab writes a pending snapshot at once; so does closing the window.
  React.useEffect(() => {
    if ((!pageActive || !selected) && snapshotTimer.current) snapshotNow();
  }, [pageActive, selected, snapshotNow]);
  React.useEffect(() => {
    const flush = () => { if (snapshotTimer.current) snapshotNow(); };
    window.addEventListener('pagehide', flush);
    return () => {
      window.removeEventListener('pagehide', flush);
      clearTimeout(snapshotTimer.current);
    };
  }, [snapshotNow]);

  const save = React.useCallback((format) => {
    const api = workbook.current;
    if (!api) return undefined;
    // ⌘S held down or pressed twice: one save at a time.
    if (saving.current) return saving.current;
    const sent = structure.current.slice();
    clearTimeout(snapshotTimer.current);
    snapshotTimer.current = null;
    // No format: back to the original file. A format: Save As through the save panel.
    saving.current = saveSheets(tabRef.current, format, sheetsForSaving(api.getAllSheets()), { sheetName: api.getSheet()?.name, structure: sent })
      .then((result) => {
        if (result?.rebased) structure.current = structure.current.slice(sent.length);
        return result;
      })
      .finally(() => { saving.current = null; });
    return saving.current;
  }, [saveSheets]);

  React.useEffect(() => {
    registerCommands(tab.id, {
      save: () => save(),
      agentEdit: (request) => takeAgentEdit(request),
      find: () => setFindOpen(true),
      // What the agent sees through doc_viewer_context.
      describe: () => {
        const api = workbook.current;
        const sheet = api?.getSheet?.();
        const selection = validSelection(api?.getSelection?.()?.[0]);
        return {
          sheet: sheet?.name,
          sheets: api?.getAllSheets?.().map((entry) => entry.name),
          selection: selection ? rangeA1(selection) : undefined,
        };
      },
    });
    onContextChange?.();
  // takeAgentEdit is declared below and only called later, through the command.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [onContextChange, registerCommands, save, tab.id]);

  const hooks = React.useMemo(() => ({
    afterSelectionChange: () => onContextChange?.(),
    afterActivateSheet: () => onContextChange?.(),
  }), [onContextChange]);

  const markDirty = React.useCallback(() => {
    setDirty(tab.id, true);
    scheduleSnapshot();
  }, [scheduleSnapshot, setDirty, tab.id]);

  // The grid's engine lacks newer functions such as XLOOKUP. Once a workbook shows that
  // it uses one, every edit also gets the answers from the Host, which has them all.
  const needsHost = React.useRef(false);
  const calculation = React.useRef({ timer: null, latest: 0, warned: false });
  const calculateOnHost = React.useCallback(async () => {
    const grid = workbook.current;
    if (!grid) return;
    const sheets = grid.getAllSheets();
    needsHost.current ||= sheets.some((sheet) => sheet.data?.some((row) => row?.some((cell) => cell?.f && UNKNOWN_FUNCTION.test(String(cell.m ?? cell.v ?? '')))));
    if (!needsHost.current || (tabRef.current.doc?.stats?.cells || 0) > MAX_CALCULATED_CELLS) return;
    calculation.current.latest += 1;
    const asked = calculation.current.latest;
    try {
      const { results } = await calculate({ path: tabRef.current.path, sheets: sheetsForSaving(sheets), structure: structure.current.slice() });
      // Someone has typed since; a newer answer is on its way.
      if (asked !== calculation.current.latest || !workbook.current) return;
      const shown = new Map(workbook.current.getAllSheets().map((sheet) => [sheet.id, sheet]));
      const changes = [];
      for (const { id, r, c, v, m } of results) {
        const cell = shown.get(id)?.data?.[r]?.[c];
        if (!cell?.f || (cell.v === v && cell.m === m)) continue;
        changes.push({ op: 'replace', id, path: ['data', r, c, 'v'], value: v }, { op: 'replace', id, path: ['data', r, c, 'm'], value: m });
      }
      // Applied as a remote change: this neither joins the undo history nor counts as an edit.
      if (changes.length) workbook.current.applyOp(changes);
    } catch (error) {
      if (!calculation.current.warned) {
        calculation.current.warned = true;
        toast(t('有些公式没能计算（{0}），保存后重新打开会显示正确结果', error.message), true);
      }
    }
  }, [toast]);
  const scheduleCalculation = React.useCallback(() => {
    clearTimeout(calculation.current.timer);
    calculation.current.timer = setTimeout(calculateOnHost, CALCULATE_DELAY_MS);
  }, [calculateOnHost]);
  React.useEffect(() => () => clearTimeout(calculation.current.timer), []);

  const onOps = React.useCallback((ops) => {
    if (settling.current) return;
    for (const op of ops || []) {
      if (op?.op === 'insertRowCol' || op?.op === 'deleteRowCol') structure.current.push({ ...op.value, op: op.op, id: op.id ?? op.value?.id });
    }
    markDirty();
    scheduleCalculation();
  }, [markDirty, scheduleCalculation]);

  /**
   * The file has formulas without saved results. The grid showed what its own engine
   * could work out; LibreOffice's answers (every function, the file's own settings)
   * replace them when they come. Like the Host's other answers, they are not an edit.
   */
  const fillHostResults = React.useCallback(async () => {
    const opened = tabRef.current;
    let doc;
    try {
      ({ doc } = await openOnHost(opened.path));
    } catch {
      return; // The grid's own results stay.
    }
    const grid = workbook.current;
    const now = tabRef.current;
    if (!grid || now.path !== opened.path || (doc.mtime && now.mtime && doc.mtime !== now.mtime)) return;
    const failed = doc.warnings?.find((warning) => warning.code === 'recalc');
    if (failed) patchTab(now.id, { doc: { ...now.doc, warnings: [...(now.doc?.warnings || []), failed] } });
    if (!doc.stats?.recalculated) return;
    const fromHost = new Map(doc.sheets.map((sheet) => [sheet.name, new Map((sheet.celldata || []).map((entry) => [`${entry.r}:${entry.c}`, entry.v]))]));
    const changes = [];
    for (const sheet of grid.getAllSheets()) {
      const cells = fromHost.get(sheet.name);
      if (!cells) continue;
      sheet.data?.forEach((row, r) => row?.forEach((cell, c) => {
        const answer = cell?.f && cells.get(`${r}:${c}`);
        // Only formulas still as the file has them; anything retyped since keeps the grid's result.
        if (!answer || answer.f !== cell.f || (answer.v === cell.v && answer.m === cell.m)) return;
        changes.push({ op: 'replace', id: sheet.id, path: ['data', r, c, 'v'], value: answer.v }, { op: 'replace', id: sheet.id, path: ['data', r, c, 'm'], value: answer.m });
      }));
    }
    if (changes.length) grid.applyOp(changes);
  }, [patchTab]);

  // On first show: select what an agent changed, or A1 (FortuneSheet otherwise starts
  // with a selection ending in NaN, shown as "A1:NaN"). Formulas with no result, which
  // LibreOffice could not compute either, go to the grid's own engine; the grid reports
  // that as edits, but nobody has changed anything.
  React.useEffect(() => {
    const timer = setTimeout(() => {
      const api = workbook.current;
      if (!api) return;
      try {
        if (tab.doc?.stats?.unresolved) {
          settling.current = true;
          setTimeout(() => { settling.current = false; }, 500);
          api.calculateFormula();
        }
      } catch { /* the saved results stay */ }
      if (tab.doc?.stats?.pendingCalculation) fillHostResults();
      try {
        const target = highlightTarget(api, tab.highlight);
        if (target) {
          if (api.getSheet()?.id !== target.id) api.activateSheet({ id: target.id });
          api.setSelection(target.ranges, { id: target.id });
          api.scroll({ targetRow: target.ranges[0].row[0], targetColumn: target.ranges[0].column[0] });
        } else {
          api.setSelection([{ row: [0, 0], column: [0, 0] }]);
        }
      } catch { /* selection is cosmetic */ }
      onContextChange?.();
    }, 60);
    return () => clearTimeout(timer);
  // Runs once per tab; a replaced document is a new tab.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /**
   * Put cells into the grid. Recorded, they are an edit like any other: undoable with
   * ⌘Z and saved with the next save. Otherwise they only make the grid show what the
   * file already holds. False when a cell is outside the grid, which then stays as it was.
   * @param {{ sheet: string, r: number, c: number, value: any }[]} cells
   * @param {boolean} recorded
   */
  const putCells = React.useCallback((cells, recorded) => {
    const grid = workbook.current;
    if (!grid) return false;
    const sheets = new Map(grid.getAllSheets().map((sheet) => [sheet.name, sheet]));
    const fits = cells.every(({ sheet, r, c }) => {
      const data = sheets.get(sheet)?.data;
      return data && r < data.length && c < (data[0]?.length || 0);
    });
    if (!fits) return false;
    if (recorded) {
      for (const { sheet, r, c, value } of cells) {
        const id = sheets.get(sheet).id;
        // Writing nothing leaves a formula behind; clearing removes it.
        if (value) grid.setCellValue(r, c, gridInput(value), { id });
        else grid.clearCell(r, c, { id });
      }
      return true;
    }
    const ops = [];
    for (const { sheet, r, c, value } of cells) {
      const target = sheets.get(sheet);
      ops.push({ op: 'replace', id: target.id, path: ['data', r, c], value: value ? structuredClone(value) : null });
      // A new formula must be known to the grid's engine, or it would not follow later edits.
      if (value?.f && !target.calcChain?.some((item) => item.r === r && item.c === c)) {
        ops.push({ op: 'add', id: target.id, path: ['calcChain', (target.calcChain?.length || 0) + ops.filter((op) => op.path[0] === 'calcChain' && op.id === target.id).length], value: { r, c, id: target.id } });
      }
    }
    // A change applied from outside: it neither joins the undo history nor counts as an edit.
    grid.applyOp(ops);
    return true;
  }, []);

  /** What the grid shows in each of these cells now. */
  const cellsNow = React.useCallback((cells) => {
    const sheets = new Map((workbook.current?.getAllSheets() || []).map((sheet) => [sheet.name, sheet]));
    return cells.map(({ sheet, r, c }) => ({ sheet, r, c, value: sheets.get(sheet)?.data?.[r]?.[c] ? structuredClone(sheets.get(sheet).data[r][c]) : null }));
  }, []);

  /**
   * An agent's edit to this document. Written to the file already, the grid is brought
   * level with it; meant for the user's unsaved edits, it joins them. Either way the
   * view stays where the user left it, and the change goes on the list.
   * @returns {boolean} whether it was taken
   */
  const takeAgentEdit = (request) => {
    const cells = request.cells || [];
    const shownBefore = cellsNow(cells);
    const taken = putCells(cells.map(({ sheet, r, c, after }) => ({ sheet, r, c, value: after })), Boolean(request.live));
    if (!taken) return false;
    if (!request.live) patchTab?.(tab.id, { mtime: request.mtime });
    // LibreOffice may not have worked out the formulas this touched (busy, or not there):
    // the grid fills in what it can now, and LibreOffice's answers follow when they come.
    if (!request.live && cells.some(({ after }) => after?.f && (after.v === undefined || after.v === null || after.v === ''))) {
      settling.current = true;
      setTimeout(() => { settling.current = false; }, 500);
      try { workbook.current?.calculateFormula(); } catch { /* the cells stay as the file has them */ }
      setTimeout(() => fillHostResults(), 0);
    }
    const edited = request.edited?.length ? request.edited : cells.map(({ sheet, address }) => ({ sheet, address }));
    setAgentChanges((list) => [...list, { id: request.id, at: Date.now(), live: Boolean(request.live), edited, cells, shownBefore, mtime: request.mtime, note: request.note }]);
    markCells(edited);
    toast(request.live
      ? t('Agent 改了 {0} 个单元格，和你没保存的修改放在一起，保存时一起写入', edited.length)
      : request.note ? t('Agent 改了 {0} 个单元格，{1}', edited.length, request.note) : t('Agent 改了 {0} 个单元格', edited.length), false, { label: t('查看'), run: () => jumpTo(edited[0]) });
    return true;
  };

  // Cells an agent wrote are outlined in place, through the grid's presence boxes,
  // which neither touch the data nor join the undo history. They wait for the user,
  // and the time limit only runs while the user can see the sheet.
  const [marks, setMarks] = React.useState([]);
  const [fading, setFading] = React.useState(false);
  const shown = pageActive && selected;
  React.useEffect(() => {
    const grid = workbook.current;
    if (!grid) return undefined;
    const sheets = new Map(grid.getAllSheets().map((sheet) => [sheet.name, sheet.id]));
    const presences = marks.map((mark, index) => ({
      userId: `dvd-agent-${mark.sheet}-${mark.r}-${mark.c}`,
      username: index === 0 ? 'Agent' : '',
      color: MARK_COLOR,
      sheetId: sheets.get(mark.sheet),
      selection: { r: mark.r, c: mark.c },
    })).filter((presence) => presence.sheetId);
    try {
      grid.addPresences(presences);
    } catch { /* marks are a courtesy */ }
    return () => {
      try { workbook.current?.removePresences(presences); } catch { /* gone with the grid */ }
    };
  }, [marks]);
  const sheetEl = React.useRef(null);
  React.useEffect(() => {
    if (!marks.length || !shown) return undefined;
    let clear;
    const fadeOut = () => {
      setFading(true);
      clear ??= setTimeout(() => { setMarks([]); setFading(false); }, MARK_FADE_MS);
    };
    const limit = setTimeout(fadeOut, MARK_MAX_MS);
    // Clicking into the sheet or typing means the user has moved on; scrolling to look around does not.
    const since = Date.now();
    const onTouch = () => { if (Date.now() - since > MARK_SETTLE_MS) fadeOut(); };
    const el = sheetEl.current;
    el?.addEventListener('pointerdown', onTouch, true);
    el?.addEventListener('keydown', onTouch, true);
    return () => {
      clearTimeout(limit);
      clearTimeout(clear);
      el?.removeEventListener('pointerdown', onTouch, true);
      el?.removeEventListener('keydown', onTouch, true);
    };
  }, [marks, shown]);
  /** Outline cells, replacing what is outlined now. */
  const markCells = React.useCallback((cells) => {
    setFading(false);
    setMarks(cells.slice(0, MAX_MARKS).map(({ sheet, address }) => {
      const decoded = decodeAddress(address);
      return decoded && { sheet, ...decoded };
    }).filter(Boolean));
  }, []);

  /** Show one cell an agent changed. */
  const jumpTo = React.useCallback(({ sheet, address }) => {
    markCells([{ sheet, address }]);
    const grid = workbook.current;
    const target = grid?.getAllSheets().find((entry) => entry.name === sheet);
    const match = /^([A-Z]+)(\d+)$/.exec(String(address).toUpperCase());
    if (!target || !match) return;
    const column = [...match[1]].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
    const row = Number(match[2]) - 1;
    try {
      if (grid.getSheet()?.id !== target.id) grid.activateSheet({ id: target.id });
      grid.setSelection([{ row: [row, row], column: [column, column] }], { id: target.id });
      // Some room above and to the left, so a frozen header does not cover the cell.
      grid.scroll({ targetRow: Math.max(0, row - 2), targetColumn: Math.max(0, column - 1) });
    } catch { /* showing is best effort */ }
  }, [markCells]);

  /**
   * Take back one agent change. The newest change written to a file the user has not
   * touched since is undone in the file too; any other comes out of the grid, as an
   * edit the next save writes.
   */
  const undoAgentChange = React.useCallback(async (group) => {
    const current = tabRef.current;
    const newestWritten = agentChangesRef.current.filter((item) => !item.live).at(-1) === group;
    if (!group.live && newestWritten && !current.dirty && current.mtime === group.mtime) {
      try {
        const { mtime } = await restoreVersion(current.path);
        putCells(group.cells.map(({ sheet, r, c, before }) => ({ sheet, r, c, value: before })), false);
        patchTab?.(tab.id, { mtime });
        toast(t('已撤销 Agent 的这次修改'));
      } catch (error) {
        toast(t('没能撤销：{0}', error.message), true);
        return;
      }
    } else if (putCells(group.shownBefore, true)) {
      toast(group.live ? t('已撤销 Agent 的这次修改') : t('已在表格里撤销，保存后写入文件'));
    } else {
      toast(t('没能撤销这次修改，可以用「恢复到之前的版本」'), true);
      return;
    }
    setAgentChanges((list) => list.filter((item) => item !== group));
  }, [patchTab, putCells, tab.id, toast]);

  const freezeFirstRow = () => {
    const api = workbook.current;
    const frozen = api?.getSheet()?.frozen;
    api?.freeze(frozen ? 'cancel' : 'row', { row: 0, column: 0 });
    markDirty();
    toast(frozen ? t('已取消冻结') : t('已冻结首行'));
  };

  const convertTextNumbers = () => {
    const api = workbook.current;
    if (!api) return;
    const sheet = api.getSheet();
    const selection = validSelection(api.getSelection()?.[0]);
    const single = !selection || (selection.row[0] === selection.row[1] && selection.column[0] === selection.column[1]);
    const calls = [];
    (sheet.data || []).forEach((row, r) => row?.forEach((cell, c) => {
      if (!cell || cell.f) return;
      if (!single && (r < selection.row[0] || r > selection.row[1] || c < selection.column[0] || c > selection.column[1])) return;
      const number = parseNumericText(cell.v);
      if (number !== undefined) calls.push({ name: 'setCellValue', args: [r, c, number, { id: sheet.id }] });
    }));
    if (calls.length === 0) {
      toast(single ? t('这个工作表里没有可转换的文本数字') : t('选中范围里没有可转换的文本数字'));
      return;
    }
    api.batchCallApis(calls);
    markDirty();
    toast(t('已把 {0} 个文本数字转换为数字', calls.length));
  };

  const insertStats = async (result, chartSvg) => {
    const api = workbook.current;
    if (!api) return;
    const sheets = api.getAllSheets();
    const stats = buildStatsSheet(result, uniqueSheetName(sheets, t('统计结果')));
    stats.order = sheets.length;
    if (chartSvg) {
      const src = await svgToPngDataUrl(chartSvg).catch(() => undefined);
      if (src) stats.images = [{ id: `chart-${Date.now()}`, src, left: 360, top: 12, width: 620, height: 260 }];
    }
    api.updateSheet([...sheets, stats]);
    markDirty();
    setStatsOpen(false);
    toast(t('已插入工作表“{0}”', stats.name));
  };

  const saveStats = async (result) => {
    const api = workbook.current;
    if (!api) return;
    const sheets = api.getAllSheets();
    const stats = buildStatsSheet(result, uniqueSheetName(sheets, t('统计结果')));
    const saved = await saveSheets(tabRef.current, 'xlsx', [...sheetsForSaving(sheets), { ...stats, order: sheets.length }], {
      suggestedName: t('{0}-统计.xlsx', tabRef.current.name.replace(/\.[^.]+$/, '')),
      detached: true,
    });
    if (saved) setStatsOpen(false);
  };

  const hasFrozen = Boolean(workbook.current?.getSheet?.()?.frozen);

  return (
    <div className="dvd-sheet" ref={sheetEl} data-marks-fading={fading ? '1' : undefined}>
      <TabActions target={actionsEl}>
        {agentChanges.length > 0 && <AgentChanges groups={agentChanges} onJump={jumpTo} onUndo={undoAgentChange} onClear={() => setAgentChanges([])} />}
        <button type="button" className="dvd-iconbtn" title={t('查找替换 ⌘F')} aria-label={t('查找替换')} aria-pressed={findOpen} onClick={() => setFindOpen((open) => !open)}><Icon name="search" /></button>
        <button type="button" className="dvd-iconbtn" title={t('统计')} aria-label={t('统计')} onClick={() => setStatsOpen(true)}><Icon name="stats" /></button>
        <Menu
          label={t('表格操作')}
          items={[
            { label: hasFrozen ? t('取消冻结') : t('冻结首行'), icon: 'freeze', onClick: freezeFirstRow },
            { label: t('文本转数字'), icon: 'number', onClick: convertTextNumbers },
            '-',
            { label: t('另存为 Excel…'), icon: 'sheet', onClick: () => save('xlsx') },
            { label: t('另存为旧版 Excel (.xls)…'), icon: 'sheet', onClick: () => save('xls') },
            { label: t('另存为 CSV…'), icon: 'text', onClick: () => save('csv') },
            { label: t('导出 PDF…'), icon: 'pdf', onClick: () => save('pdf') },
            { label: t('恢复到之前的版本…'), icon: 'undo', onClick: () => onVersions(tabRef.current.path), disabled: tab.transient },
            '-',
            ...toolItems,
          ]}
        />
        <button type="button" className={`dvd-btn${tab.dirty ? ' primary' : ''}`} style={{ margin: '0 4px' }} title={tab.transient ? t('这是拖进来的文件，保存时选择存放位置 ⌘S') : t('存回原文件，修改前的版本会自动保留 ⌘S')} onClick={() => save()}>{t('保存')}</button>
      </TabActions>
      <Workbook
        ref={workbook}
        data={initialSheets}
        lang={lang()}
        allowEdit
        showToolbar
        showFormulaBar
        showSheetTabs
        toolbarItems={TOOLBAR_ITEMS}
        cellContextMenu={CELL_MENU}
        headerContextMenu={HEADER_MENU}
        sheetTabContextMenu={SHEET_TAB_MENU}
        onOp={onOps}
        hooks={hooks}
      />
      {findOpen && <SheetFind workbook={workbook} onClose={() => setFindOpen(false)} onChanged={markDirty} toast={toast} />}
      {statsOpen && <StatsSheet workbook={workbook} onClose={() => setStatsOpen(false)} onInsert={insertStats} onSave={saveStats} />}
    </div>
  );
}

/** @param {string} address A1 @returns {{ r: number, c: number }|undefined} */
function decodeAddress(address) {
  const match = /^([A-Z]+)(\d+)$/.exec(String(address).toUpperCase());
  if (!match) return undefined;
  return { r: Number(match[2]) - 1, c: [...match[1]].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1 };
}

/** A cell as the grid takes it through setCellValue: a formula is calculated again. */
function gridInput(cell) {
  if (cell.f) return { f: cell.f, ct: cell.ct || { fa: 'General', t: 'n' } };
  const { mc: _merge, ...rest } = cell;
  return rest;
}

/** The changes agents made here, each to look at or take back. */
function AgentChanges({ groups, onJump, onUndo, onClear }) {
  const [open, setOpen] = React.useState(false);
  const wrap = React.useRef(null);
  React.useEffect(() => {
    if (!open) return undefined;
    const close = (event) => { if (!wrap.current?.contains(event.target)) setOpen(false); };
    const onKey = (event) => { if (event.key === 'Escape') setOpen(false); };
    document.addEventListener('mousedown', close);
    window.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', close);
      window.removeEventListener('keydown', onKey, true);
    };
  }, [open]);
  const total = groups.reduce((sum, group) => sum + group.edited.length, 0);
  return (
    <span className="dvd-menuwrap" ref={wrap}>
      <button type="button" className="dvd-btn plain dvd-changes-chip" title={t('Agent 改了 {0} 处', total)} aria-label={t('Agent 改了 {0} 处', total)} aria-expanded={open} onClick={() => setOpen((value) => !value)}>
        <Icon name="sparkle" size={14} /><span className="dvd-wide">{t('Agent 改了')} </span><span className="dvd-count" key={total}>{total}</span><span className="dvd-wide"> {t('处')}</span>
      </button>
      {open && (
        <div className="dvd-menu dvd-changes" role="dialog" aria-label={t('Agent 的修改')}>
          {groups.slice().reverse().map((group) => (
            <div className="dvd-change" key={group.id}>
              <div className="dvd-change-head">
                <span>{new Date(group.at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                <span className="dvd-change-tag">{group.live ? t('在你未保存的修改里') : t('已写入文件')}</span>
                <button type="button" className="dvd-change-undo" onClick={() => onUndo(group)}>{t('撤销')}</button>
              </div>
              <div className="dvd-change-cells">
                {group.edited.slice(0, 24).map((cell) => (
                  <button type="button" key={`${cell.sheet}!${cell.address}`} className="dvd-change-cell" title={`${cell.sheet}!${cell.address}`} onClick={() => onJump(cell)}>{cell.address}</button>
                ))}
                {group.edited.length > 24 && <span className="dvd-change-more">{t('另有 {0} 处', group.edited.length - 24)}</span>}
              </div>
              {group.note && <div className="dvd-change-note">{group.note}</div>}
            </div>
          ))}
          <div className="dvd-sep" role="separator" />
          <button type="button" onClick={() => { setOpen(false); onClear(); }}>{t('清空这个列表')}</button>
        </div>
      )}
    </span>
  );
}

function SheetFind({ workbook, onClose, onChanged, toast }) {
  const [query, setQuery] = React.useState('');
  const [replacement, setReplacement] = React.useState('');
  const [replacing, setReplacing] = React.useState(false);
  const [matches, setMatches] = React.useState([]);
  const [index, setIndex] = React.useState(0);

  const collect = React.useCallback((text) => {
    const needle = text.trim().toLocaleLowerCase();
    const found = [];
    if (needle === '') return found;
    for (const sheet of workbook.current?.getAllSheets() || []) {
      (sheet.data || []).some((row, r) => row?.some((cell, c) => {
        if (cell && (displayOf(cell).toLocaleLowerCase().includes(needle) || String(cell.f || '').toLocaleLowerCase().includes(needle))) {
          found.push({ sheetId: sheet.id, r, c });
        }
        return found.length >= MAX_MATCHES;
      }));
      if (found.length >= MAX_MATCHES) break;
    }
    return found;
  }, [workbook]);

  const goTo = React.useCallback((match) => {
    const api = workbook.current;
    if (!api || !match) return;
    try {
      if (api.getSheet()?.id !== match.sheetId) api.activateSheet({ id: match.sheetId });
      api.setSelection([{ row: [match.r, match.r], column: [match.c, match.c] }], { id: match.sheetId });
      api.scroll({ targetRow: match.r, targetColumn: match.c });
    } catch { /* navigation is best effort */ }
  }, [workbook]);

  React.useEffect(() => {
    const timer = setTimeout(() => {
      const found = collect(query);
      setMatches(found);
      setIndex(0);
      if (found.length) goTo(found[0]);
    }, 160);
    return () => clearTimeout(timer);
  }, [collect, goTo, query]);

  const move = (delta) => {
    if (matches.length === 0) return;
    const next = (index + delta + matches.length) % matches.length;
    setIndex(next);
    goTo(matches[next]);
  };

  const replaceAll = () => {
    const api = workbook.current;
    if (!api || query === '') return;
    const needle = query.toLocaleLowerCase();
    const calls = [];
    for (const match of collect(query)) {
      const sheet = api.getAllSheets().find((entry) => entry.id === match.sheetId);
      const cell = sheet?.data?.[match.r]?.[match.c];
      if (!cell || cell.f) continue;
      const before = displayOf(cell);
      const position = before.toLocaleLowerCase().indexOf(needle);
      if (position < 0) continue;
      // Replace every occurrence, whatever its case, as the search found them.
      const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      calls.push({ name: 'setCellValue', args: [match.r, match.c, before.replace(pattern, () => replacement), { id: match.sheetId }] });
    }
    if (calls.length === 0) {
      toast(t('没有可替换的内容（公式单元格不会被替换）'));
      return;
    }
    api.batchCallApis(calls);
    onChanged();
    toast(t('已替换 {0} 个单元格', calls.length));
    setMatches(collect(query));
    setIndex(0);
  };

  return (
    <FindBar query={query} onQuery={setQuery} count={matches.length} index={index} onNext={() => move(1)} onPrev={() => move(-1)} onClose={onClose} placeholder={t('在所有工作表中查找')}>
      {replacing ? (
        <div className="dvd-find-row" style={{ gap: 6 }}>
          <input className="dvd-input" value={replacement} placeholder={t('替换为')} aria-label={t('替换为')} onChange={(event) => setReplacement(event.target.value)} />
          <button type="button" className="dvd-btn" disabled={matches.length === 0} onClick={replaceAll}>{t('全部替换')}</button>
        </div>
      ) : (
        <div className="dvd-find-row">
          <button type="button" className="dvd-btn plain" style={{ height: 24, padding: '0 8px' }} onClick={() => setReplacing(true)}>{t('替换…')}</button>
        </div>
      )}
    </FindBar>
  );
}

/** @param {{ row: number[], column: number[] }} range */
function rangeA1({ row, column }) {
  const start = `${colName(Math.min(...column))}${Math.min(...row) + 1}`;
  const end = `${colName(Math.max(...column))}${Math.max(...row) + 1}`;
  return start === end ? start : `${start}:${end}`;
}

/**
 * The cells an agent wrote, on the first sheet it wrote to, as grid ranges.
 * @param {any} api
 * @param {{ sheet: string, address: string }[]|undefined} highlight
 */
function highlightTarget(api, highlight) {
  if (!highlight?.length) return undefined;
  const sheets = api.getAllSheets();
  const name = highlight[0].sheet;
  const sheet = sheets.find((entry) => entry.name === name) || sheets[0];
  const ranges = highlight.filter((item) => (item.sheet || sheet.name) === sheet.name).slice(0, 50).map((item) => {
    const match = /^([A-Z]+)(\d+)$/.exec(String(item.address).toUpperCase());
    if (!match) return undefined;
    const column = [...match[1]].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
    const row = Number(match[2]) - 1;
    return { row: [row, row], column: [column, column] };
  }).filter(Boolean);
  return ranges.length ? { id: sheet.id, ranges } : undefined;
}

/** A selection with real bounds, or undefined. */
function validSelection(selection) {
  return selection && [...(selection.row || []), ...(selection.column || [])].every(Number.isFinite) ? selection : undefined;
}

/** The cells the statistics apply to: the selection if it is a range, else the whole sheet. */
function statsScope(api) {
  const sheet = api?.getSheet?.();
  if (!sheet?.data) return null;
  const { maxRow, maxColumn } = usedBounds(sheet.data);
  if (maxRow < 0) return null;
  const selection = validSelection(api.getSelection()?.[0]);
  const isRange = selection && (selection.row[0] !== selection.row[1] || selection.column[0] !== selection.column[1]);
  const rows = isRange ? [Math.min(...selection.row), Math.min(Math.max(...selection.row), maxRow)] : [0, maxRow];
  const columns = isRange ? [Math.min(...selection.column), Math.min(Math.max(...selection.column), maxColumn)] : [0, maxColumn];
  const cells = [];
  for (let r = rows[0]; r <= rows[1]; r += 1) cells.push((sheet.data[r] || []).slice(columns[0], columns[1] + 1));
  return {
    cells,
    firstRow: rows[0],
    firstColumn: columns[0],
    label: t('{0}，{1}', sheet.name, isRange ? `${colName(columns[0])}${rows[0] + 1}:${colName(columns[1])}${rows[1] + 1}` : t('全部数据')),
  };
}

function StatsSheet({ workbook, onClose, onInsert, onSave }) {
  const [scope] = React.useState(() => statsScope(workbook.current));
  const [hasHeader, setHasHeader] = React.useState(() => (scope ? firstRowIsHeader(scope.cells, scope.firstRow) : true));
  const [groupIndex, setGroupIndex] = React.useState(-1);
  const [valueIndex, setValueIndex] = React.useState(-1);
  const [agg, setAgg] = React.useState('sum');
  const [chartType, setChartType] = React.useState('bar');

  const columns = React.useMemo(() => {
    const first = scope?.cells[0] || [];
    return first.map((cell, index) => ({ index, name: (hasHeader ? displayOf(cell) : '') || colName(scope.firstColumn + index) }));
  }, [hasHeader, scope]);

  const result = React.useMemo(() => {
    if (!scope) return null;
    const { summary, groups } = computeSheetStats(scope.cells, { hasHeader, groupIndex, valueIndex, agg });
    return {
      summary,
      groups,
      agg,
      groupName: columns[groupIndex]?.name,
      valueName: columns[valueIndex]?.name,
    };
  }, [agg, columns, groupIndex, hasHeader, scope, valueIndex]);

  const chartTitle = result && `${result.groupName} · ${result.valueName}${aggLabels()[agg]}`;
  const chartSvg = React.useMemo(() => (result?.groups.length ? renderChartSvg(result.groups, chartTitle, chartType) : ''), [chartTitle, chartType, result]);

  if (!scope) {
    return (
      <Sheet title={t('统计')} onClose={onClose} footer={<button type="button" className="dvd-btn" onClick={onClose}>{t('好')}</button>}>
        <p>{t('这个工作表还没有数据。')}</p>
      </Sheet>
    );
  }
  const { summary, groups } = result;
  return (
    <Sheet
      title={t('统计')}
      wide
      onClose={onClose}
      footer={(
        <>
          <button type="button" className="dvd-btn" onClick={() => onSave(result)}>{t('另存为…')}</button>
          <button type="button" className="dvd-btn primary" onClick={() => onInsert(result, chartSvg)}>{t('插入工作表')}</button>
        </>
      )}
    >
      <div className="dvd-crumb" style={{ justifyContent: 'space-between', gap: 12 }}>
        <span className="dvd-label">{scope.label}</span>
        <label className="dvd-check"><input type="checkbox" checked={hasHeader} onChange={(event) => setHasHeader(event.target.checked)} />{t('首行是表头')}</label>
      </div>
      <div className="dvd-fieldrow">
        <label className="dvd-field">
          <span className="dvd-label">{t('分组')}</span>
          <select className="dvd-select" value={groupIndex} onChange={(event) => setGroupIndex(Number(event.target.value))}>
            <option value={-1}>{t('不分组')}</option>
            {columns.map((column) => <option key={column.index} value={column.index}>{column.name}</option>)}
          </select>
        </label>
        <label className="dvd-field">
          <span className="dvd-label">{t('数值')}</span>
          <select className="dvd-select" value={valueIndex} onChange={(event) => setValueIndex(Number(event.target.value))}>
            <option value={-1}>{t('全部列')}</option>
            {columns.map((column) => <option key={column.index} value={column.index}>{column.name}</option>)}
          </select>
        </label>
        <label className="dvd-field">
          <span className="dvd-label">{t('汇总方式')}</span>
          <select className="dvd-select" value={agg} onChange={(event) => setAgg(event.target.value)}>
            {Object.entries(aggLabels()).map(([value, label]) => <option key={value} value={value}>{label}</option>)}
          </select>
        </label>
      </div>
      <div className="dvd-tiles">
        {[[t('合计'), summary.sum], [t('平均'), summary.average], [t('计数'), summary.count], [t('最大'), summary.max], [t('最小'), summary.min], [t('去重'), summary.distinct]].map(([label, value]) => (
          <div className="dvd-tile" key={label}><span>{label}</span><strong title={String(value ?? '')}>{formatNumber(value)}</strong></div>
        ))}
      </div>
      {groups.length > 0 ? (
        <>
          <Segmented value={chartType} onChange={setChartType} options={[{ value: 'bar', label: t('柱状图') }, { value: 'line', label: t('折线图') }]} />
          {/* The markup is built by renderChartSvg, which escapes every value from the sheet. */}
          <div className="dvd-chart" dangerouslySetInnerHTML={{ __html: chartSvg }} />
        </>
      ) : (
        <div className="dvd-callout">{t('选择“分组”和“数值”两列，可按类别汇总并生成图表。')}</div>
      )}
    </Sheet>
  );
}
