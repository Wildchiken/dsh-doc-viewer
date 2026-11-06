/**
 * The document page, loaded on first use. It mounts once and stays mounted for
 * the session; `active` says whether the page is the one on screen.
 */
import React from 'react';
import fortuneCss from '@fortune-sheet/react/dist/index.css';
import { CSS } from './styles.js';
import { Icon, iconForKind } from './icons.jsx';
import * as api from './api.js';
import { forgetRecent, loadRecent, rememberRecent } from './recent.js';
import { Menu, Sheet, Skeleton, skeletonKind } from './ui.jsx';
import { SpreadsheetTab } from './SpreadsheetTab.jsx';
import { PdfTab } from './PdfTab.jsx';
import { DocxTab, TextTab } from './DocumentTabs.jsx';
import { ConvertDialog, EmptyState, OpenDialog, PdfToolsDialog } from './Dialogs.jsx';

let tabCounter = 0;
/** The documents open in the page, so they come back after a restart. */
const TABS_KEY = 'dsh-doc-viewer:tabs';
/** Set once the person has seen that a document can sit beside the conversation. */
const HINT_KEY = 'dsh-doc-viewer:hint-beside';
/** How often open documents are checked for changes made elsewhere. */
const STAT_INTERVAL_MS = 3000;
/** A change is only someone else's once the page's own update for it has had time to arrive. */
const STAT_GRACE_MS = 6000;

function loadSavedTabs() {
  try {
    const saved = JSON.parse(localStorage.getItem(TABS_KEY) || '{}');
    return { paths: Array.isArray(saved.paths) ? saved.paths.filter((path) => typeof path === 'string') : [], active: saved.active };
  } catch {
    return { paths: [] };
  }
}

function ensureStyles() {
  if (document.querySelector('style[data-dvd-viewer]')) return;
  const tag = document.createElement('style');
  tag.dataset.plugin = 'dsh-doc-viewer';
  tag.dataset.dvdViewer = '';
  tag.textContent = `${fortuneCss}\n${CSS}`;
  document.head.appendChild(tag);
}

/** @param {any} error */
function friendly(error) {
  if (error?.code === 'office-engine-unavailable') return '本机没有找到 Office 转换引擎（Harness 自带的 LibreOffice）。';
  if (error instanceof TypeError && /fetch/i.test(error.message)) return '和 Harness 的连接断开了，稍后再试。';
  return error?.message || '操作失败，请重试。';
}

/** @param {string} path */
const baseName = (path) => path.split('/').pop();
/** @param {string} path */
const folderName = (path) => path.split('/').slice(-2, -1)[0] || '/';
/** @param {string} path */
const dirOf = (path) => path.split('/').slice(0, -1).join('/') || '/';
/** @param {string} name */
const stemOf = (name) => name.replace(/\.[^.]+$/, '');

/** How long things take to leave: matches --dvd-fast in the styles. */
const EXIT_MS = 140;

/** Put tabs that are folding away back where they were, for the strip only. */
function withClosing(tabs, closing) {
  if (!closing.length) return tabs;
  const list = [...tabs];
  for (const { tab, index } of closing) {
    if (!list.some((item) => item.id === tab.id)) list.splice(Math.min(index, list.length), 0, { ...tab, closing: true });
  }
  return list;
}

/** Describe a range of the open sheet the way Excel writes it. */
function rangeLabel(sheet, selection) {
  return selection ? `${sheet}!${selection}` : sheet;
}

export function Viewer({ active, shell, placement }) {
  ensureStyles();
  React.useState(() => api.configure({ root: () => shell.workspaceRoot?.() }));
  const [tabs, setTabs] = React.useState([]);
  const [activeId, setActiveId] = React.useState(null);
  const [dialog, setDialog] = React.useState(null);
  const [toast, setToast] = React.useState(null);
  const [busy, setBusy] = React.useState(0);
  const [dragging, setDragging] = React.useState(false);
  const [recent, setRecent] = React.useState(loadRecent);
  const [actionsEl, setActionsEl] = React.useState(null);
  const tabsRef = React.useRef(tabs);
  tabsRef.current = tabs;
  const activeIdRef = React.useRef(activeId);
  activeIdRef.current = activeId;
  const activeRef = React.useRef(active);
  activeRef.current = active;
  /** Per-tab commands for the keyboard: { save, find }. */
  const commands = React.useRef(new Map());
  const opening = React.useRef(new Map());
  const fileInput = React.useRef(null);
  const toastTimer = React.useRef(null);
  const [closing, setClosing] = React.useState([]);
  const [besideHint, setBesideHint] = React.useState(false);
  const dismissBesideHint = React.useCallback(() => {
    setBesideHint(false);
    try { localStorage.setItem(HINT_KEY, '1'); } catch { /* the hint simply comes back next time */ }
  }, []);
  const dragDepth = React.useRef(0);

  /** `actions` is one { label, run } or a list of them. */
  const showToast = React.useCallback((message, isError = false, actions) => {
    const list = actions ? [].concat(actions) : [];
    setToast({ id: Date.now(), message, error: isError, actions: list });
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => hideToast(), isError || list.length ? 8000 : 3200);
  // hideToast is a stable callback declared just below.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /** Let the toast slide away before it is gone. */
  const hideToast = React.useCallback(() => {
    clearTimeout(toastTimer.current);
    setToast((current) => (current && !current.leaving ? { ...current, leaving: true } : current));
    toastTimer.current = setTimeout(() => setToast(null), EXIT_MS);
  }, []);

  /** Hand a question to the agent in the conversation, with what it needs to know. */
  const askAgent = React.useCallback((prompt) => {
    if (shell.askAgent?.(prompt)) return;
    navigator.clipboard?.writeText(prompt).then(
      () => showToast('没能自动填进对话框，已复制，粘贴给 Agent 即可'),
      () => showToast('没能打开对话，请切到会话里描述问题', true),
    );
  }, [shell, showToast]);

  /** An error, with a way to ask the agent about it. */
  const showError = React.useCallback((message, context) => {
    showToast(message, true, { label: '问 Agent', run: () => askAgent(`文档页出了问题：${message}${context ? `\n文件：${context}` : ''}\n请帮我看看原因，能处理的话直接处理。`) });
  }, [askAgent, showToast]);

  /** Where a new file went, and ways to get to it. */
  const announceFile = React.useCallback((output, message) => {
    showToast(message || `已保存到“${folderName(output)}”文件夹：${baseName(output)}`, false, [
      { label: '打开', run: () => openPath(output) },
      { label: '在访达中显示', run: () => api.reveal(output).catch((error) => showToast(friendly(error), true)) },
    ]);
  // openPath is declared below; the reference is resolved when a toast button is pressed.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [showToast]);

  /** Run something slow, with the thin progress line showing meanwhile. */
  const track = React.useCallback(async (work) => {
    setBusy((count) => count + 1);
    try {
      return await work();
    } finally {
      setBusy((count) => count - 1);
    }
  }, []);

  const remember = React.useCallback((path, name, kind) => setRecent(rememberRecent({ path, name, kind })), []);

  /**
   * Open a file in a tab. `replaces` makes it take the place of the tab showing an
   * earlier version (an agent's edit), with `highlight` selected there.
   */
  const openPath = React.useCallback((path, options = {}) => {
    if (!path) return undefined;
    const existing = tabsRef.current.find((tab) => tab.path === path);
    if (existing && !options.force && options.replaces !== path) {
      if (!options.background) setActiveId(existing.id);
      return Promise.resolve(existing.id);
    }
    // Two quick requests for one file must not make two tabs.
    if (opening.current.has(path)) return opening.current.get(path);
    const previous = options.replaces ? tabsRef.current.find((tab) => tab.path === options.replaces) : undefined;
    // A file changed in place is shown again in its own tab, unless the user has edits there
    // (or has just chosen to throw them away).
    if (existing && options.replaces === path && previous?.dirty && !options.fresh) return Promise.resolve(existing.id);
    const replacing = previous && (!previous.dirty || options.fresh) ? previous : undefined;
    tabCounter += 1;
    const id = `tab-${tabCounter}`;
    const name = baseName(path) || '文档';
    const placeholder = { id, path, name, kind: 'loading', loading: true };
    setTabs((list) => (replacing ? list.map((tab) => (tab.id === replacing.id ? placeholder : tab)) : [...list, placeholder]));
    if (replacing) commands.current.delete(replacing.id);
    if (!options.background || (replacing && activeIdRef.current === replacing.id)) setActiveId(id);
    const loading = track(async () => {
      try {
        // Formula results LibreOffice has to work out arrive after the grid is on screen.
        const { doc } = await api.open(path, { recalc: 'later' });
        // Edits left unsaved last time come back on top of the file.
        const recovered = (doc.kind === 'spreadsheet' || doc.kind === 'text') && !options.agent && !options.fresh ? (await api.recoveryItem(path).catch(() => ({}))).record : undefined;
        setTabs((list) => list.map((tab) => (tab.id === id
          ? {
            id, path, name: doc.name || name, kind: doc.kind, ext: doc.ext, doc,
            sheets: recovered?.sheets || doc.sheets,
            recoveredText: recovered?.text,
            structure: recovered?.structure,
            dirty: Boolean(recovered),
            mtime: recovered?.openedMtime ?? doc.mtime,
            transient: doc.transient || recovered?.transient,
            highlight: options.highlight,
          }
          : tab)));
        if (!doc.transient) remember(path, doc.name || name, doc.kind);
        if (recovered) {
          showToast(`已恢复「${doc.name || name}」上次没有保存的修改（${new Date(recovered.at).toLocaleString()}）`, false, [
            { label: '放弃这些修改', run: () => discardEdits(path) },
          ]);
        }
        if (options.agent) {
          const count = options.highlight?.length || 0;
          const what = [count ? `改了 ${count} 个单元格` : '', options.note || ''].filter(Boolean).join('，') || '修改了文件';
          showToast(options.inPlace
            ? `Agent ${what}`
            : `Agent ${what}，另存为 ${name}${previous?.dirty ? '（原文件有你未保存的修改，所以没有覆盖它）' : ''}`, false, options.inPlace
            ? [{ label: '撤销', run: () => undoLastSave(path) }]
            : [{ label: '看原文件', run: () => openPath(options.replaces) }, { label: '在访达中显示', run: () => revealFile(path) }]);
        } else if (options.message) {
          showToast(options.message);
        }
        return id;
      } catch (error) {
        setTabs((list) => (replacing
          ? list.map((tab) => (tab.id === id ? replacing : tab))
          : list.filter((tab) => tab.id !== id)));
        setActiveId((current) => (current === id ? (replacing?.id ?? tabsRef.current.find((tab) => tab.id !== id)?.id ?? null) : current));
        if (error?.code === 'not-found') {
          setRecent(forgetRecent(path));
          // The file is gone, but unsaved edits to it are not: open those, to be saved somewhere.
          const recovered = (await api.recoveryItem(path).catch(() => ({}))).record;
          if (recovered?.sheets) {
            setTabs((list) => [...list.filter((tab) => tab.id !== id), {
              id, path, name: recovered.name || name, kind: 'spreadsheet', ext: '.xlsx', doc: { kind: 'spreadsheet', stats: {}, warnings: [] },
              sheets: recovered.sheets, dirty: true, transient: true,
            }]);
            setActiveId(id);
            showToast(`原文件已经不在了，已恢复「${recovered.name || name}」没保存的内容，保存时请选个位置`);
            return id;
          }
        }
        if (!options.restoring) showError(friendly(error), path);
        return undefined;
      } finally {
        opening.current.delete(path);
      }
    });
    opening.current.set(path, loading);
    return loading;
  }, [remember, showError, showToast, track]);

  const openFiles = React.useCallback(async (files) => {
    const list = [...(files || [])];
    await track(async () => {
      for (const file of list) {
        try {
          const uploaded = await api.upload(file);
          await openPath(uploaded.path);
        } catch (error) {
          showError(`${file.name}：${friendly(error)}`);
        }
      }
    });
  }, [openPath, showError, track]);

  // Files asked for from outside: the Host's preview and agent tools.
  React.useEffect(() => {
    const drain = () => {
      while (shell.queue.length > 0) {
        const { path, announce, ...options } = shell.queue.shift();
        if (options.kind === 'agent-edit') {
          applyAgentEdit({ path, ...options });
          continue;
        }
        if (options.replaces) options.agent = true;
        if (announce) {
          remember(path, baseName(path));
          announceFile(path, `Agent 生成了 ${baseName(path)}，在“${folderName(path)}”文件夹`);
        } else {
          openPath(path, options);
        }
      }
    };
    const onRecent = () => setRecent(loadRecent());
    drain();
    window.addEventListener('dsh-doc-viewer:queued', drain);
    window.addEventListener('dsh-doc-viewer:recent', onRecent);
    return () => {
      window.removeEventListener('dsh-doc-viewer:queued', drain);
      window.removeEventListener('dsh-doc-viewer:recent', onRecent);
    };
  // applyAgentEdit is declared below; the listener only calls it later.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [announceFile, openPath, remember, shell]);

  /** Change some of a tab's state, such as the file time after an update it made itself. */
  const patchTab = React.useCallback((tabId, partial) => {
    setTabs((previous) => previous.map((tab) => (tab.id === tabId ? { ...tab, ...partial } : tab)));
  }, []);

  /**
   * An agent changed a document. A workbook open here takes the change where it is;
   * anything else open here is shown again from the file; a document not open only
   * goes to the recent list.
   */
  const applyAgentEdit = React.useCallback((request) => {
    const tab = tabsRef.current.find((item) => item.path === request.path && !item.loading);
    if (!tab) {
      remember(request.path, baseName(request.path));
      if (request.reveal) openPath(request.path);
      return;
    }
    const take = commands.current.get(tab.id)?.agentEdit;
    if (take && request.cells && take(request)) {
      if (request.reveal) setActiveId(tab.id);
      return;
    }
    // Nothing here could take it piece by piece: show the file as it now is, if that loses nothing.
    if (!tab.dirty) openPath(request.path, { replaces: request.path, background: !request.reveal, message: `Agent 修改了「${tab.name}」` });
    else showToast(`Agent 的改动没能放进「${tab.name}」，请保存后重新打开`, true);
  }, [openPath, remember, showToast]);

  const setDirty = React.useCallback((tabId, dirty = true) => {
    setTabs((previous) => {
      const current = previous.find((tab) => tab.id === tabId);
      if (!current || Boolean(current.dirty) === dirty) return previous;
      return previous.map((tab) => (tab.id === tabId ? { ...tab, dirty } : tab));
    });
  }, []);

  /** Keep a snapshot of a document's unsaved edits, to restore after quitting or a crash. */
  const saveSnapshot = React.useCallback((tab, { sheets, structure, text }) => {
    api.saveRecovery({ path: tab.path, name: tab.name, sheets, structure, text, openedMtime: tab.mtime, transient: tab.transient })
      .catch((error) => console.warn('[doc-viewer] could not keep unsaved edits', error));
  }, []);

  /** Throw away restored edits and show the file as it is. */
  const discardEdits = React.useCallback(async (path) => {
    await api.discardRecovery(path);
    await openPath(path, { replaces: path, fresh: true, message: '已放弃没保存的修改' });
  }, [openPath]);

  const revealFile = React.useCallback((path) => {
    api.reveal(path).catch((error) => showToast(friendly(error), true));
  }, [showToast]);

  /**
   * Ask where a new file goes, in the system's save panel, starting beside the
   * document. Null when the user cancels; undefined when there is no panel, in
   * which case the file goes beside the document.
   */
  const chooseLocation = React.useCallback(async (tab, name) => {
    try {
      const { path } = await api.pickSave({ name, directory: tab && !tab.transient ? dirOf(tab.path) : undefined });
      return path || null;
    } catch (error) {
      if (error?.code === 'native-unavailable') return undefined;
      throw error;
    }
  }, []);

  /** Step back to what the file was before its last change, and show it. */
  const undoLastSave = React.useCallback(async function undo(path) {
    try {
      const { version } = await api.restoreVersion(path);
      await openPath(path, { replaces: path });
      showToast('已撤销这次修改', false, version.remaining > 0 ? { label: '继续撤销', run: () => undo(path) } : undefined);
    } catch (error) {
      showError(friendly(error), path);
    }
  }, [openPath, showError, showToast]);

  /**
   * Save a text document back to its file. `onSaved` also runs if the user only gets
   * there by choosing to overwrite a file changed elsewhere.
   * @returns the Host's answer, or undefined when nothing was saved
   */
  const saveText = React.useCallback(async function save(tab, text, { force, onSaved } = {}) {
    try {
      const result = await track(() => api.saveText({ path: tab.path, text, openedMtime: force ? undefined : (tab.acceptedMtime ?? tab.mtime) }));
      if (result.conflict) {
        showToast(`「${tab.name}」在别处被修改过，这次没有保存`, true, [
          { label: '用我的覆盖', run: () => save(tab, text, { force: true, onSaved }) },
          { label: '载入新内容', run: () => discardEdits(tab.path) },
        ]);
        return undefined;
      }
      setTabs((list) => list.map((item) => (item.id === tab.id ? { ...item, dirty: false, mtime: result.mtime, acceptedMtime: undefined } : item)));
      onSaved?.();
      showToast(result.warnings?.length ? `已保存。${result.warnings[0].message}` : '已保存', false, { label: '撤销保存', run: () => undoLastSave(tab.path) });
      return result;
    } catch (error) {
      showError(friendly(error), tab.path);
      return undefined;
    }
  }, [discardEdits, showError, showToast, track, undoLastSave]);

  const convertTab = React.useCallback(async (tab, target, acknowledgeLossy = false, chosen) => {
    try {
      let outputPath = chosen;
      if (outputPath === undefined) {
        outputPath = await chooseLocation(tab, `${stemOf(tab.name)}.${target}`);
        if (outputPath === null) return;
      }
      const response = await track(() => api.convert({ paths: [tab.path], target, acknowledgeLossy, outputPath: outputPath || undefined }));
      if (response.requiresConfirmation) {
        setDialog({ type: 'lossy', tab, target, outputPath: outputPath || false, message: response.warnings?.[0]?.message });
        return;
      }
      const result = response.results?.[0];
      if (!result?.ok) throw new Error(result?.message || '转换失败');
      remember(result.output, baseName(result.output));
      const warning = result.warnings?.find((item) => item.code !== 'lossy');
      announceFile(result.output, `${target === 'pdf' ? '已导出' : '已另存为'} ${baseName(result.output)}${warning ? `。${warning.message}` : ''}`);
    } catch (error) {
      showError(friendly(error), tab.path);
    }
  }, [announceFile, chooseLocation, remember, showError, track]);

  /**
   * Save a sheet. Without a format it goes back to the file it came from, with the
   * previous content kept as a version; with one, the save panel decides where.
   * @returns the Host's answer, or undefined when nothing was saved
   */
  const saveSheets = React.useCallback(async (tab, format, sheets, options = {}) => {
    try {
      if (!format && !tab.transient) {
        const result = await track(() => api.save({ mode: 'in-place', sourcePath: tab.path, sheets, openedMtime: tab.mtime, structure: options.structure }));
        if (result.needsLocation) {
          // Overwriting would lose something; say what, and offer to save elsewhere.
          return await new Promise((resolve) => setDialog({
            type: 'needsLocation',
            reason: result.reason,
            onCancel: () => resolve(undefined),
            onSaveAs: async () => resolve(await saveSheets(tab, 'xlsx', sheets, { ...options, suggestedName: result.suggestedName })),
          }));
        }
        setTabs((list) => list.map((item) => (item.id === tab.id ? { ...item, dirty: false, mtime: result.mtime } : item)));
        showToast(result.warnings?.length ? `已保存。${result.warnings[0].message}` : '已保存', false, { label: '撤销保存', run: () => undoLastSave(tab.path) });
        return { ...result, rebased: true };
      }
      const target = format || 'xlsx';
      const name = options.suggestedName || `${stemOf(tab.name)}.${target}`;
      const outputPath = await chooseLocation(tab, name);
      if (outputPath === null) return undefined;
      const result = await track(() => api.save({
        sheets,
        sourcePath: tab.path,
        format: target,
        outputPath: outputPath || undefined,
        title: outputPath ? undefined : stemOf(name),
        structure: options.structure,
        options: { sheetName: options.sheetName },
      }));
      const fileName = baseName(result.output);
      const warning = result.warnings?.[0]?.message;
      remember(result.output, fileName, target === 'pdf' ? 'pdf' : 'spreadsheet');
      if (target === 'pdf' || options.detached) {
        announceFile(result.output, `${target === 'pdf' ? '已导出' : '已另存为'} ${fileName}${warning ? `。${warning}` : ''}`);
        return result;
      }
      // Like Save As anywhere else: this tab is now the new file, and nothing of the old one is pending.
      if (result.output !== tab.path) api.discardRecovery(tab.path);
      setTabs((list) => list.map((item) => (item.id === tab.id
        ? { ...item, path: result.output, name: fileName, ext: `.${target}`, dirty: false, mtime: result.mtime, transient: false }
        : item)));
      announceFile(result.output, `已另存为 ${fileName}${warning ? `。${warning}` : ''}`);
      return { ...result, rebased: true };
    } catch (error) {
      showError(friendly(error), tab.path);
      return undefined;
    }
  }, [announceFile, chooseLocation, remember, showError, showToast, track, undoLastSave]);

  /** A new, empty workbook, named and placed in the save panel, then opened. */
  const newSpreadsheet = React.useCallback(async () => {
    try {
      const { path } = await api.pickSave({ name: '未命名表格.xlsx', message: '新建表格', fresh: true });
      if (!path) return;
      const { output } = await track(() => api.create(path));
      await openPath(output, { message: `已新建「${baseName(output)}」` });
    } catch (error) {
      if (error?.code === 'native-unavailable') showToast('这个系统上没有保存面板，可以在对话里让 Agent 新建表格', true);
      else showError(friendly(error));
    }
  }, [openPath, showError, showToast, track]);

  /** Open files where they are, through the system's open panel. */
  const openWithPanel = React.useCallback(async () => {
    try {
      const { paths } = await api.pickOpen();
      for (const path of paths) await openPath(path);
    } catch (error) {
      if (error?.code === 'native-unavailable') setDialog({ type: 'open' });
      else showError(friendly(error));
    }
  }, [openPath, showError]);

  const closeNow = React.useCallback((tab) => {
    commands.current.delete(tab.id);
    // Closed on purpose: its unsaved edits, if any, were saved or thrown away.
    api.discardRecovery(tab.path);
    const list = tabsRef.current;
    const index = list.findIndex((item) => item.id === tab.id);
    const next = list.filter((item) => item.id !== tab.id);
    setTabs(next);
    if (activeIdRef.current === tab.id) setActiveId(next[Math.max(0, index - 1)]?.id ?? null);
    // The tab is gone at once; only its label lingers in the strip while it folds away.
    setClosing((list) => [...list, { tab, index }]);
    setTimeout(() => setClosing((list) => list.filter((item) => item.tab.id !== tab.id)), EXIT_MS);
  }, []);

  const closeTab = (tab) => {
    if (tab.dirty) setDialog({ type: 'close', tab });
    else closeNow(tab);
  };

  // What was open last time opens again, and documents left with unsaved edits come back with them.
  const restored = React.useRef(false);
  React.useEffect(() => {
    (async () => {
      const saved = loadSavedTabs();
      const { entries = [] } = await api.recoveryList().catch(() => ({}));
      const paths = [...saved.paths, ...entries.map((entry) => entry.path).filter((path) => !saved.paths.includes(path))];
      let activeTab;
      let shownWhileLoading;
      for (const path of paths) {
        if (tabsRef.current.some((tab) => tab.path === path)) continue;
        // With nothing on screen yet, the first file shows its outline while it loads.
        const showNow = !activeIdRef.current;
        const id = await openPath(path, { background: !showNow, restoring: true });
        if (showNow) shownWhileLoading = id;
        if (path === saved.active || (!activeTab && !saved.active)) activeTab = id ?? activeTab;
      }
      const current = activeIdRef.current;
      if (activeTab && (!current || current === shownWhileLoading || !tabsRef.current.some((tab) => tab.id === current))) setActiveId(activeTab);
      restored.current = true;
    })();
  // Once, when the page first loads.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  React.useEffect(() => {
    if (!restored.current) return;
    try {
      const kept = tabs.filter((tab) => !tab.transient && tab.path);
      localStorage.setItem(TABS_KEY, JSON.stringify({ paths: kept.map((tab) => tab.path), active: kept.find((tab) => tab.id === activeId)?.path }));
    } catch { /* storage may be unavailable */ }
  }, [tabs, activeId]);

  // Documents changed elsewhere: shown again when nothing here would be lost, else the user decides.
  React.useEffect(() => {
    const seen = new Map();
    const tick = async () => {
      if (document.visibilityState !== 'visible') return;
      const watched = tabsRef.current.filter((tab) => !tab.loading && tab.path && !tab.transient && tab.mtime);
      if (watched.length === 0) return;
      const { stats = {} } = await api.stat(watched.map((tab) => tab.path)).catch(() => ({}));
      for (const { id } of watched) {
        const tab = tabsRef.current.find((item) => item.id === id);
        const now = tab ? stats[tab.path] : undefined;
        if (!tab || now === undefined || now === tab.mtime || now === tab.acceptedMtime) {
          seen.delete(id);
          continue;
        }
        const entry = seen.get(id);
        if (!entry || entry.mtime !== now) {
          seen.set(id, { mtime: now, since: Date.now() });
          continue;
        }
        if (entry.handled || Date.now() - entry.since < STAT_GRACE_MS) continue;
        entry.handled = true;
        if (now === null) {
          showToast(`「${tab.name}」已被删除或移走，这里显示的是之前的内容`, true);
        } else if (!tab.dirty) {
          openPath(tab.path, { replaces: tab.path, background: true, message: `「${tab.name}」在别处被修改，已显示最新内容` });
        } else {
          showToast(`「${tab.name}」在别处被修改了，而你这里有没保存的修改`, true, [
            { label: '载入新内容', run: () => discardEdits(tab.path) },
            { label: '保留我的', run: () => patchTab(tab.id, { acceptedMtime: now }) },
          ]);
        }
      }
    };
    const timer = setInterval(tick, STAT_INTERVAL_MS);
    return () => clearInterval(timer);
  // discardEdits and openPath are stable enough; the interval reads tabs through the ref.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // A browser asks before leaving a page with unsaved edits. Electron shows no
  // such prompt and would silently refuse to close or reload the window, so
  // there it is left alone.
  React.useEffect(() => {
    if (/Electron/i.test(navigator.userAgent)) return undefined;
    const warn = (event) => {
      if (!tabsRef.current.some((tab) => tab.dirty)) return;
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, []);

  // The page is only a keyboard target while it is on screen.
  React.useEffect(() => {
    const onKey = (event) => {
      if (!activeRef.current || !(event.metaKey || event.ctrlKey)) return;
      const key = event.key.toLowerCase();
      if (key === 'o') {
        event.preventDefault();
        openWithPanel();
      } else if (key === 'n') {
        event.preventDefault();
        newSpreadsheet();
      } else if (key === 's' || key === 'f') {
        const command = commands.current.get(activeIdRef.current)?.[key === 's' ? 'save' : 'find'];
        if (command) {
          event.preventDefault();
          command();
        }
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  React.useEffect(() => () => clearTimeout(toastTimer.current), []);

  // The first time a document is open full page, point out once that it can sit beside the chat.
  const hasDocument = tabs.some((tab) => !tab.loading);
  React.useEffect(() => {
    if (!hasDocument || !active || (placement && placement !== 'page') || !shell.canOpenBeside?.()) return undefined;
    try { if (localStorage.getItem(HINT_KEY)) return undefined; } catch { return undefined; }
    const timer = setTimeout(() => setBesideHint(true), 900);
    const hide = setTimeout(() => setBesideHint(false), 900 + 9000);
    return () => { clearTimeout(timer); clearTimeout(hide); };
  }, [active, hasDocument, placement, shell]);

  // A file preview frames one file. Opening or making another one there would show it
  // under the first file's name, so the page moves to its own tab beside the conversation.
  React.useEffect(() => {
    if (!placement?.startsWith('preview')) return;
    const current = tabs.find((tab) => tab.id === activeId);
    if (current && !current.loading && shell.previewPath && current.path !== shell.previewPath) shell.openBeside?.();
  }, [activeId, placement, shell, tabs]);

  // Tell the Host what is open, so the agent knows what "this table" means.
  const reportTimer = React.useRef(null);
  const reportState = React.useCallback(() => {
    clearTimeout(reportTimer.current);
    reportTimer.current = setTimeout(() => {
      const list = tabsRef.current.filter((tab) => !tab.loading);
      const current = list.find((tab) => tab.id === activeIdRef.current);
      const detail = current ? commands.current.get(current.id)?.describe?.() : undefined;
      api.reportState({
        visible: activeRef.current,
        active: current ? { path: current.path, name: current.name, kind: current.kind, dirty: Boolean(current.dirty), ...detail } : undefined,
        tabs: list.map((tab) => ({ path: tab.path, name: tab.name, kind: tab.kind, dirty: Boolean(tab.dirty) })),
      });
    }, 300);
  }, []);
  React.useEffect(reportState, [tabs, activeId, active, reportState]);

  const registerCommands = React.useCallback((tabId, value) => commands.current.set(tabId, value), []);

  /** "问 Agent" from the top bar: the current document and selection, ready for one sentence. */
  const askAboutActive = () => {
    const tab = tabsRef.current.find((item) => item.id === activeIdRef.current);
    if (!tab) {
      askAgent('关于“文档”页：');
      return;
    }
    const detail = commands.current.get(tab.id)?.describe?.() || {};
    const where = detail.sheet ? `，${rangeLabel(detail.sheet, detail.selection)}` : detail.page ? `，第 ${detail.page} 页` : '';
    askAgent(`关于「${tab.name}」${where}：`);
  };
  const closeDialog = React.useCallback(() => setDialog(null), []);
  const pdfTools = React.useCallback((action, path) => setDialog({ type: 'pdf', action, path }), []);
  const activeTab = tabs.find((tab) => tab.id === activeId);
  // Workspace-wide tools, offered at the end of every document's menu.
  const toolItems = React.useMemo(() => [
    { label: '批量转换…', icon: 'convert', onClick: () => setDialog({ type: 'convert' }) },
    { label: 'PDF 工具…', icon: 'pdf', onClick: () => setDialog({ type: 'pdf', action: 'merge' }) },
  ], []);

  return (
    <div
      className="dvd-root"
      onDragEnter={(event) => {
        if (!event.dataTransfer?.types?.includes('Files')) return;
        event.preventDefault();
        dragDepth.current += 1;
        setDragging(true);
      }}
      onDragOver={(event) => { if (event.dataTransfer?.types?.includes('Files')) event.preventDefault(); }}
      onDragLeave={() => {
        dragDepth.current = Math.max(0, dragDepth.current - 1);
        if (dragDepth.current === 0) setDragging(false);
      }}
      onDrop={(event) => {
        if (!event.dataTransfer?.files?.length) return;
        event.preventDefault();
        dragDepth.current = 0;
        setDragging(false);
        openFiles(event.dataTransfer.files);
      }}
    >
      <div className="dvd-bar" data-window-drag="true">
        <div className="dvd-tabs" role="tablist" aria-label="打开的文档">
          {withClosing(tabs, closing).map((tab) => tab.closing ? (
            <div key={tab.id} className="dvd-tab closing" aria-hidden="true">
              <Icon name={iconForKind(tab.kind)} size={14} />
              <span className="dvd-tab-name">{stemOf(tab.name)}</span>
            </div>
          ) : (
            <div
              key={tab.id}
              role="tab"
              tabIndex={0}
              aria-selected={tab.id === activeId}
              className="dvd-tab"
              title={tab.path}
              onClick={() => setActiveId(tab.id)}
              onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') setActiveId(tab.id); }}
              onAuxClick={(event) => { if (event.button === 1) closeTab(tab); }}
            >
              <Icon name={iconForKind(tab.kind)} size={14} />
              <span className="dvd-tab-name">{stemOf(tab.name)}<span className="dvd-tab-ext">{tab.name.slice(stemOf(tab.name).length)}</span></span>
              <button
                type="button"
                className="dvd-tab-close"
                aria-label={`关闭 ${tab.name}`}
                onClick={(event) => { event.stopPropagation(); closeTab(tab); }}
              >
                <Icon name="close" size={11} strokeWidth={2} />
              </button>
              {tab.dirty && tab.id !== activeId && <span className="dvd-dirty" style={{ right: 12 }} title="有未保存的修改" />}
            </div>
          ))}
        </div>
        <Menu
          label="打开或新建"
        align="left"
          icon="plus"
          items={[
            { label: '打开…  ⌘O', icon: 'folder', onClick: openWithPanel },
            { label: '新建表格  ⌘N', icon: 'sheet', onClick: newSpreadsheet },
          ]}
        />
        <div className="dvd-actions">
          {shell.canOpenBeside?.() && !placement?.startsWith('preview') && (
            <span className="dvd-menuwrap">
              <button
                type="button"
                className="dvd-iconbtn"
                title={placement && placement !== 'page' ? '整页查看' : '在对话旁打开'}
                aria-label={placement && placement !== 'page' ? '整页查看' : '在对话旁打开'}
                onClick={() => { dismissBesideHint(); if (placement && placement !== 'page') shell.openAsPage(); else shell.openBeside(); }}
              >
                <Icon name={placement && placement !== 'page' ? 'expand' : 'sidebar'} size={16} />
              </button>
              {besideHint && (!placement || placement === 'page') && (
                <span className="dvd-hint" role="note">
                  也可以开在对话旁，边聊边看
                  <button type="button" onClick={dismissBesideHint}>知道了</button>
                </span>
              )}
            </span>
          )}
          {(!placement || placement === 'page') && <button type="button" className="dvd-btn plain dvd-ask" title="让 Agent 处理这个文档" aria-label="问 Agent" onClick={askAboutActive}>
            <Icon name="sparkle" size={15} /><span className="dvd-wide">问 Agent</span>
          </button>}
          <span className="dvd-actions-slot" ref={setActionsEl} />
          {(!activeTab || activeTab.loading) && <Menu label="更多工具" items={toolItems} />}
        </div>
      </div>
      {busy > 0 && <div className="dvd-progress" role="progressbar" aria-label="处理中" />}

      <div className="dvd-body">
        {tabs.length === 0 && (
          <div className="dvd-pane active">
            <EmptyState
              recent={recent}
              onOpen={openWithPanel}
              onBrowse={() => setDialog({ type: 'open' })}
              onOpenPath={openPath}
              onNew={newSpreadsheet}
            />
          </div>
        )}
        {tabs.map((tab) => (
          <div key={tab.id} className={`dvd-pane${tab.id === activeId ? ' active' : ''}`} aria-hidden={tab.id !== activeId}>
            <TabContent
              tab={tab}
              selected={tab.id === activeId}
              pageActive={active}
              actionsEl={tab.id === activeId ? actionsEl : null}
              setDirty={setDirty}
              patchTab={patchTab}
              saveSheets={saveSheets}
              saveText={saveText}
              onConvert={convertTab}
              onPdfTools={pdfTools}
              registerCommands={registerCommands}
              onContextChange={reportState}
              saveSnapshot={saveSnapshot}
              onVersions={(path) => setDialog({ type: 'versions', path })}
              askAgent={askAgent}
              toolItems={toolItems}
              toast={showToast}
              openPath={openPath}
              track={track}
            />
          </div>
        ))}
        {dragging && <div className="dvd-drop">松开以打开</div>}
        {toast && (
          <div key={toast.id} className={`dvd-toast${toast.error ? ' error' : ''}${toast.leaving ? ' leaving' : ''}`} role="status">
            <span>{toast.message}</span>
            {toast.actions.map((action) => (
              <button type="button" key={action.label} onClick={() => { hideToast(); action.run(); }}>{action.label}</button>
            ))}
          </div>
        )}
      </div>

      {dialog?.type === 'open' && (
        <OpenDialog
          onClose={closeDialog}
          onUpload={() => { closeDialog(); fileInput.current?.click(); }}
          onOpenPath={(path) => { closeDialog(); openPath(path); }}
        />
      )}
      {dialog?.type === 'convert' && (
        <ConvertDialog tabs={tabs} onClose={closeDialog} onOpenPath={(path) => { closeDialog(); openPath(path); }} toast={showToast} track={track} />
      )}
      {dialog?.type === 'pdf' && (
        <PdfToolsDialog
          tabs={tabs}
          activeTab={activeTab}
          initialAction={dialog.action}
          initialPath={dialog.path}
          onClose={closeDialog}
          onOpenPath={(path) => { closeDialog(); openPath(path); }}
          toast={showToast}
          track={track}
        />
      )}
      {dialog?.type === 'lossy' && (
        <Sheet
          title="这是有损转换"
          onClose={closeDialog}
          footer={(
            <>
              <button type="button" className="dvd-btn" onClick={closeDialog}>取消</button>
              <button type="button" className="dvd-btn primary" onClick={() => { const { tab, target, outputPath } = dialog; closeDialog(); convertTab(tab, target, true, outputPath); }}>继续转换</button>
            </>
          )}
        >
          <p>{dialog.message || '复杂排版、图片和分栏可能丢失。'}</p>
        </Sheet>
      )}
      {dialog?.type === 'needsLocation' && (
        <Sheet
          title="不能直接存回原文件"
          onClose={() => { dialog.onCancel(); closeDialog(); }}
          footer={(
            <>
              <button type="button" className="dvd-btn" onClick={() => { dialog.onCancel(); closeDialog(); }}>取消</button>
              <button type="button" className="dvd-btn primary" onClick={() => { closeDialog(); dialog.onSaveAs(); }}>另存为…</button>
            </>
          )}
        >
          <p>{dialog.reason}。原文件保持不变；另存的新文件里不会有这些内容。</p>
        </Sheet>
      )}
      {dialog?.type === 'versions' && (
        <VersionsSheet path={dialog.path} onClose={closeDialog} onRestore={async (id) => {
          closeDialog();
          try {
            await api.restoreVersion(dialog.path, id);
            await openPath(dialog.path, { replaces: dialog.path, message: '已恢复到所选版本，恢复前的内容也保留了' });
          } catch (error) {
            showError(friendly(error), dialog.path);
          }
        }}
        />
      )}
      {dialog?.type === 'close' && (
        <Sheet
          title={`“${dialog.tab.name}”有未保存的修改`}
          onClose={closeDialog}
          footer={(
            <>
              <button type="button" className="dvd-btn" onClick={closeDialog}>取消</button>
              <button type="button" className="dvd-btn" onClick={() => { const { tab } = dialog; closeDialog(); closeNow(tab); }}>不保存</button>
              <button type="button" className="dvd-btn primary" onClick={async () => {
                const { tab } = dialog;
                closeDialog();
                const saved = await commands.current.get(tab.id)?.save?.();
                if (saved) closeNow(tab);
              }}
              >保存</button>
            </>
          )}
        >
          <p>不保存的话，这些修改会丢失。</p>
        </Sheet>
      )}
      <input
        ref={fileInput}
        type="file"
        multiple
        hidden
        onChange={(event) => {
          const files = [...(event.target.files || [])];
          event.target.value = '';
          if (files.length) openFiles(files);
        }}
      />
    </div>
  );
}

function TabContent(props) {
  const { tab } = props;
  if (tab.loading) {
    return <Skeleton kind={skeletonKind(tab.name)} />;
  }
  const warnings = tab.doc?.warnings || [];
  return (
    <>
      <TabErrorBoundary tab={tab} askAgent={props.askAgent}>
        {tab.kind === 'spreadsheet' && <SpreadsheetTab {...props} />}
        {tab.kind === 'pdf' && <PdfTab {...props} />}
        {tab.kind === 'docx' && <DocxTab {...props} />}
        {tab.kind === 'text' && <TextTab {...props} />}
      </TabErrorBoundary>
      {warnings.length > 0 && (
        <div className="dvd-notice">
          <Icon name="warning" size={13} />
          <span>{warnings.map((warning) => warning.message).join(' ')}</span>
        </div>
      )}
    </>
  );
}

class TabErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  componentDidCatch(error, info) {
    console.error('[doc-viewer] tab render failed', error, info);
  }

  render() {
    if (this.state.error) {
      return (
        <div className="dvd-center">
          <h3>这个文档无法显示</h3>
          <p>{String(this.state.error?.message || this.state.error)}</p>
          <div className="dvd-empty-actions" style={{ marginTop: 4 }}>
            <button type="button" className="dvd-btn" onClick={() => this.setState({ error: null })}>重试</button>
            <button type="button" className="dvd-btn" onClick={() => this.props.askAgent(`文档页显示「${this.props.tab.name}」时出错：${String(this.state.error?.message || this.state.error)}\n文件：${this.props.tab.path}\n请帮我看看原因。`)}>问 Agent</button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

/** Earlier versions of a file saved in place, newest first. */
function VersionsSheet({ path, onClose, onRestore }) {
  const [versions, setVersions] = React.useState(null);
  const [error, setError] = React.useState(null);
  React.useEffect(() => {
    api.versions(path).then((result) => setVersions(result.versions), (reason) => setError(reason.message));
  }, [path]);
  return (
    <Sheet title="恢复到之前的版本" onClose={onClose} footer={<button type="button" className="dvd-btn" onClick={onClose}>完成</button>}>
      <p>每次保存或 Agent 修改前，原来的内容都会自动保留（最近的版本和最早的一份，共 20 个，保留 30 天），不占用你的文件夹。</p>
      <div className="dvd-list">
        {versions?.map((version, index) => (
          <div className="dvd-result" key={version.id}>
            <span className="dvd-row-name">{new Date(version.at).toLocaleString()}</span>
            <span className="dvd-row-meta">{version.reason.endsWith('前') ? version.reason : `${version.reason}前`}{index > 0 && index === versions.length - 1 ? ' · 最早' : ''}</span>
            <button type="button" className="dvd-btn plain" onClick={() => onRestore(version.id)}>恢复</button>
          </div>
        ))}
        {versions?.length === 0 && <div className="dvd-empty-note">还没有保存过的版本</div>}
        {!versions && !error && <div className="dvd-empty-note">正在读取…</div>}
        {error && <div className="dvd-empty-note dvd-bad">{error}</div>}
      </div>
    </Sheet>
  );
}
