/**
 * The always-loaded half of the plugin: a sidebar entry beside Automation tasks,
 * the main-area page behind it, and the bridge that shows files agents ask for.
 *
 * The Host unmounts a page when another one is selected. A workbook being edited
 * must survive that, with its undo history and scroll position, so the viewer is
 * rendered through a portal into one long-lived element. The page only decides
 * where that element sits and whether it is visible; switching away hides it.
 */
import React from 'react';
import { t, useLang } from './i18n.js';
import { createPortal } from 'react-dom';
import { ack, recoveryList, requests } from './api.js';
import { rememberRecent } from './recent.js';

const PLUGIN_ID = 'dsh-doc-viewer';
const PANEL_ID = 'doc-viewer';
/** The same page as a tab of the Harness's right sidebar, beside the conversation. */
const SIDE_KIND = 'doc-viewer';
const SIDE_ID = 'dsh-doc-viewer';
/** As the Harness's own file preview: the default for workbooks, a choice for the rest. */
const EDITOR_ID = 'dsh-doc-viewer/editor';
const READER_ID = 'dsh-doc-viewer/reader';
const EDITABLE = ['xlsx', 'xls', 'csv', 'tsv', 'et', 'wps'];
const READABLE = ['pdf', 'docx', 'doc', 'md', 'markdown', 'txt'];
const BINARY = ['xlsx', 'xls', 'et', 'wps', 'pdf', 'docx', 'doc'];
const POLL_MS = 2500;
const POLL_MAX_MS = 30_000;

export const inject = ['slots', 'locale', 'layout', 'sidebarRight', 'sidebarRightTabs', 'documentPreviews'];

/** Shared between the sidebar entry, the page and the persistent viewer. */
function createShell() {
  const listeners = new Set();
  const host = document.createElement('div');
  host.className = 'dvd-host';
  const state = { active: false, unseen: 0, Viewer: null, loadError: null, placement: undefined };
  let snapshot = { ...state };
  const emit = () => {
    snapshot = { ...state };
    for (const listener of listeners) listener();
  };
  return {
    host,
    /** Paths waiting for the viewer, consumed when it mounts or by a live viewer. */
    queue: [],
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    get: () => snapshot,
    set(partial) {
      Object.assign(state, partial);
      emit();
    },
  };
}

const shell = createShell();

// The grid sizes itself on window resizes only. Moving between the column and the
// sidebar, or dragging the sidebar's edge, changes its room without one, so a change
// of the page's own size is passed on as one.
if (typeof ResizeObserver === 'function') {
  let pending = 0;
  new ResizeObserver(() => {
    cancelAnimationFrame(pending);
    pending = requestAnimationFrame(() => window.dispatchEvent(new Event('resize')));
  }).observe(shell.host);
}

/** @param {any} require the loader's require, which can fetch package-local chunks */
function viewerLoader(require) {
  let loading;
  return () => {
    loading ??= require.async('./client.viewer.js').then(
      (module) => { shell.set({ Viewer: module.Viewer, loadError: null }); },
      (error) => {
        loading = undefined;
        console.error('[doc-viewer] failed to load the viewer', error);
        shell.set({ loadError: String(error?.message || error) });
      },
    );
    return loading;
  };
}

function useShell() {
  return React.useSyncExternalStore(shell.subscribe, shell.get);
}

/** The sidebar glyph. A dot marks files an agent produced since the page was last open. */
function NavIcon({ size = 16 }) {
  const { unseen } = useShell();
  return (
    <span style={{ position: 'relative', display: 'inline-flex' }}>
      <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
        <path d="M6.5 2.75h7.25L19 8v12a1.25 1.25 0 0 1-1.25 1.25H6.5A1.25 1.25 0 0 1 5.25 20V4A1.25 1.25 0 0 1 6.5 2.75Z" />
        <path d="M13.5 3v5.25H19" />
        <path d="M8.75 13h6.5M8.75 16.75h6.5" />
      </svg>
      {unseen > 0 && (
        <span style={{ position: 'absolute', top: -1, right: -2, width: 6, height: 6, borderRadius: '50%', background: 'var(--dsw-alias-state-error-primary)' }} />
      )}
    </span>
  );
}

/**
 * Show the one viewer at a place: the main column ('page') or a right-sidebar tab
 * ('side'). It moves rather than mounting again, so its documents, unsaved edits and
 * undo history go with it.
 * @param {HTMLElement} container
 * @param {'page'|'side'} placement
 */
function claimHost(container, placement) {
  if (shell.host.parentElement !== container) {
    if (getComputedStyle(container).position === 'static') container.style.position = 'relative';
    container.appendChild(shell.host);
  }
  shell.set({ active: true, unseen: 0, placement });
}

/** @param {'page'|'side'} placement */
function releaseHost(placement) {
  if (shell.get().placement === placement) shell.set({ active: false, placement: undefined });
}

/** Marks where the viewer belongs. The viewer stays in place when this unmounts. */
function createPage(load) {
  return function Page() {
    const anchor = React.useRef(null);
    React.useLayoutEffect(() => {
      // Slot wrappers use display: contents and cannot position anything; use the column itself.
      let parent = anchor.current?.parentElement;
      while (parent && getComputedStyle(parent).display === 'contents') parent = parent.parentElement;
      if (parent) claimHost(parent, 'page');
      void load();
      return () => releaseHost('page');
    }, []);
    return <div ref={anchor} style={{ flex: 1, minWidth: 0, minHeight: 0 }} />;
  };
}

/**
 * A place in the right sidebar that shows the one viewer while it is on screen. When
 * another place has the viewer at the same time (two panes), it offers to bring it here.
 * @param {{ placement: string, visible: boolean, load: () => void, onShown?: () => void }} props
 */
function ViewerAnchor({ placement: mine, visible, load, onShown }) {
  useLang();
  const { placement } = useShell();
  const anchor = React.useRef(null);
  const [claimed, setClaimed] = React.useState(0);
  React.useLayoutEffect(() => {
    if (!visible || placement === 'page' || !anchor.current) return undefined;
    claimHost(anchor.current, mine);
    void load();
    onShown?.();
    return () => releaseHost(mine);
  // Claimed again when this place comes into view, or when asked to.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [visible, placement === 'page', claimed]);
  const elsewhere = visible && placement !== undefined && placement !== mine;
  return (
    <div ref={anchor} style={{ position: 'relative', display: 'flex', width: '100%', height: '100%', minWidth: 0, minHeight: 0 }}>
      {elsewhere && (
        <div style={{ margin: 'auto', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10, fontSize: 13, color: 'var(--dsw-alias-label-secondary)' }}>
          <span>{t('文档正显示在另一处')}</span>
          <button type="button" onClick={() => setClaimed((value) => value + 1)} style={{ height: 30, padding: '0 12px', border: 0, borderRadius: 'var(--dsw-radius-sm)', background: 'var(--dsw-alias-interactive-bg-hover)', color: 'var(--dsw-alias-label-primary)', font: 'inherit', cursor: 'pointer' }}>{t('在这里显示')}</button>
        </div>
      )}
    </div>
  );
}

/** The right-sidebar tab: takes the viewer while it is the tab on screen. */
function createSideTab(load) {
  return function SideTab({ useTabInfo }) {
    const visible = useTabInfo?.()?.tab?.visible !== false;
    return <ViewerAnchor placement="side" visible={visible} load={load} />;
  };
}

/**
 * The Harness's file preview, answered by this plugin: the file opens here, editable,
 * in the same viewer as everywhere else.
 */
function createPreviewBody(load) {
  return function PreviewBody(props) {
    const { tab } = props.useTabInfo();
    const meta = props.useResource(tab.contentId);
    const path = meta?.value?.absolutePath;
    const version = meta?.value?.version;
    const request = props.content?.kind === 'renderer' ? props.content : undefined;
    // The preview shows its own loading until told the file is shown.
    React.useEffect(() => {
      if (request && path) request.loaded(version);
    }, [request?.revision, path, version]);
    const show = React.useCallback(() => {
      if (!path) return;
      shell.previewPath = path;
      shell.queue.push({ path, focus: true });
      window.dispatchEvent(new CustomEvent('dsh-doc-viewer:queued'));
    }, [path]);
    if (!path) return null;
    return <ViewerAnchor placement={`preview:${tab.id}`} visible={tab.visible !== false} load={load} onShown={show} />;
  };
}

/** The right sidebar's controller, when this Harness has one. */
function sidebarRight(ctx) {
  try {
    return ctx.get?.('sidebarRight') ?? ctx.sidebarRight;
  } catch {
    return undefined;
  }
}

/** Whether a conversation fills the main column, so a sidebar tab can open beside it. */
function conversationOnScreen(ctx) {
  try {
    return ctx.layout.panelInfo?.getSnapshot?.()?.activePanelId === null && Boolean(currentSessionId(ctx));
  } catch {
    return false;
  }
}

/**
 * Show the page beside the conversation. False when there is no conversation to
 * stand beside or the Harness has no right sidebar.
 */
function openBeside(ctx, load) {
  const right = sidebarRight(ctx);
  if (!right?.openTab || !currentSessionId(ctx)) return false;
  try {
    if (!conversationOnScreen(ctx)) ctx.layout.selectPanel(null);
    // The conversation's sidebar exists once the conversation is back on screen.
    setTimeout(() => {
      try {
        right.openTab(SIDE_KIND);
      } catch (error) {
        console.warn('[doc-viewer] could not open beside the conversation', error);
        ctx.layout.selectPanel(PANEL_ID);
      }
    }, conversationOnScreen(ctx) ? 0 : 80);
    void load();
    return true;
  } catch {
    return false;
  }
}

/** Lives for the whole session in a slot that is never unmounted. */
function PersistentViewer() {
  useLang();
  const { Viewer, active, loadError, placement } = useShell();
  if (loadError && !Viewer) {
    return createPortal(
      <div style={{ padding: 32, fontSize: 13, color: 'var(--dsw-alias-label-secondary)' }}>{t('文档查看器加载失败：')}{loadError}</div>,
      shell.host,
    );
  }
  if (!Viewer) return null;
  shell.host.dataset.active = active ? '1' : '0';
  shell.host.dataset.placement = placement || '';
  return createPortal(<Viewer active={active} shell={shell} placement={placement} />, shell.host);
}

/**
 * Queue something for the viewer and, unless it only announces a file, show the page.
 * @param {any} ctx
 * @param {string|{ path: string, replaces?: string, highlight?: any[], note?: string, announce?: boolean }} [request]
 */
function openFile(ctx, load, request) {
  if (request) {
    shell.queue.push(typeof request === 'string' ? { path: request } : request);
    window.dispatchEvent(new CustomEvent('dsh-doc-viewer:queued'));
  }
  // Already on screen, in the column or beside the conversation: nothing to switch.
  if (!request?.announce && !shell.get().active) {
    // While the user is in a conversation, the page opens beside it rather than taking its place.
    if (!(conversationOnScreen(ctx) && openBeside(ctx, load))) ctx.layout.selectPanel(PANEL_ID);
  }
  void load();
}

/** The session the user last had open, which the conversation view returns to. */
function currentSessionId(ctx) {
  return ctx.get?.('uiWorkspace')?.mainReference?.sessionId;
}

/** The workspace folder of that session; files are browsed and uploaded there. */
function workspaceRoot(ctx) {
  const id = currentSessionId(ctx);
  if (!id) return undefined;
  try {
    return ctx.get?.('sessions')?.list?.getSnapshot()?.byId?.[id]?.cwd;
  } catch {
    return undefined;
  }
}

/**
 * Put a request into the conversation's input box and show the conversation, so the
 * user only has to finish the sentence. Returns false when that is not possible.
 * @param {any} ctx
 * @param {string} prompt
 */
function askAgent(ctx, prompt) {
  try {
    const workspace = ctx.get?.('uiWorkspace');
    const conversation = ctx.get?.('conversation');
    const sessionId = currentSessionId(ctx);
    if (sessionId && workspace?.sessions?.binding && conversation?.input?.requestDraftInitialization) {
      const binding = workspace.sessions.binding(sessionId);
      if (binding) {
        ctx.layout.selectPanel(null);
        if (conversation.input.requestDraftInitialization(binding, { prompt, whenDraftExists: 'append' }) !== 'blocked') {
          // The conversation's input, not an editable element of the hidden document page.
          setTimeout(() => [...document.querySelectorAll('[contenteditable="true"]')].find((element) => !element.closest('.dvd-host'))?.focus(), 120);
          return true;
        }
      }
    }
    if (workspace?.startSession) {
      workspace.startSession(undefined, { prompt, clearPreviousDraft: false });
      return true;
    }
  } catch (error) {
    console.warn('[doc-viewer] could not hand over to the conversation', error);
  }
  return false;
}

/** The button the Host's own file preview shows for a file. */
function createOpenAction(ctx, load) {
  return function OpenInViewer(props) {
    const path = props?.absolutePath || props?.path;
    if (!path) return null;
    return (
      <button
        type="button"
        title={t('在文档中打开')}
        aria-label={t('在文档中打开')}
        onClick={() => openFile(ctx, load, path)}
        style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 28, height: 28, padding: 0, border: 'none', borderRadius: 'var(--dsw-radius-sm)', background: 'transparent', color: 'var(--dsw-alias-label-secondary)', cursor: 'pointer' }}
      >
        <NavIcon size={15} />
      </button>
    );
  };
}

/**
 * Pick up what agent tools asked to show. Only the window that claims a request
 * acts on it, and a hidden window leaves it for a visible one.
 * @param {any} ctx
 */
function startRequestPolling(ctx, load) {
  let timer;
  let stopped = false;
  let delay = POLL_MS;
  const tick = async () => {
    if (document.visibilityState === 'visible') {
      try {
        const { requests: pending = [] } = await requests();
        if (pending.length > 0) {
          const { claimed = [] } = await ack(pending.map((request) => request.id));
          for (const request of pending) {
            if (!claimed.includes(request.id)) continue;
            const { path, replaces, highlight, note, inPlace } = request;
            if (request.kind === 'agent-edit') {
              // An edit never takes the user away from the conversation. An open page
              // updates in place; a dot says there is something new to look at.
              if (shell.get().Viewer) openFile(ctx, load, { ...request, announce: !request.reveal });
              else rememberRecent({ path, name: path.split('/').pop() });
              if (request.reveal && !shell.get().Viewer) openFile(ctx, load);
              if (!shell.get().active && !request.reveal) shell.set({ unseen: shell.get().unseen + 1 });
            } else if (request.reveal) {
              openFile(ctx, load, { path, replaces, highlight, note, inPlace });
            } else if (shell.get().active) {
              openFile(ctx, load, { path, announce: true });
            } else {
              rememberRecent({ path, name: path.split('/').pop() });
              window.dispatchEvent(new CustomEvent('dsh-doc-viewer:recent'));
              shell.set({ unseen: shell.get().unseen + 1 });
            }
          }
        }
        delay = POLL_MS;
      } catch {
        delay = Math.min(POLL_MAX_MS, delay * 2);
      }
    }
    if (!stopped) timer = setTimeout(tick, delay);
  };
  timer = setTimeout(tick, POLL_MS);
  return () => {
    stopped = true;
    clearTimeout(timer);
  };
}

/**
 * Register the sidebar entry, its page and the open-in-viewer actions.
 * @param {any} ctx
 */
export function apply(ctx) {
  const load = viewerLoader(require);
  shell.workspaceRoot = () => workspaceRoot(ctx);
  shell.askAgent = (prompt) => askAgent(ctx, prompt);
  shell.showConversation = () => ctx.layout.selectPanel(null);
  shell.openBeside = () => openBeside(ctx, load);
  shell.openAsPage = () => ctx.layout.selectPanel(PANEL_ID);
  shell.canOpenBeside = () => Boolean(sidebarRight(ctx)?.openTab) && Boolean(currentSessionId(ctx));
  // Harness's own dictionaries; the page's language is only known once the Harness has started,
  // so labels below are translated each time they are shown.
  const zh = { panel: '文档' };
  const en = { panel: 'Documents' };
  const panelLabel = () => t('文档');
  ctx.effect(() => ctx.locale.register('doc-viewer', { zh, en }), `${PLUGIN_ID}: dictionaries`);
  ctx.effect(() => {
    const tag = document.createElement('style');
    tag.dataset.plugin = PLUGIN_ID;
    tag.textContent = '.dvd-host{position:absolute;inset:0;z-index:1;display:flex}.dvd-host:not([data-active="1"]){visibility:hidden;pointer-events:none}';
    document.head.appendChild(tag);
    return () => tag.remove();
  }, `${PLUGIN_ID}: host styles`);
  ctx.effect(() => startRequestPolling(ctx, load), `${PLUGIN_ID}: open requests`);
  // Unsaved edits from last time: a dot on the sidebar entry until the page is opened.
  ctx.effect(() => {
    const timer = setTimeout(() => {
      recoveryList().then(({ entries = [] }) => {
        if (entries.length && !shell.get().Viewer) shell.set({ unseen: shell.get().unseen + entries.length });
      }, () => {});
    }, 1500);
    return () => clearTimeout(timer);
  }, `${PLUGIN_ID}: unsaved edits`);
  ctx.effect(() => () => shell.host.remove(), `${PLUGIN_ID}: viewer element`);

  ctx.slots.inject('sidebar.panellist', () => ctx.slots.register({
    name: 'sidebar.panellist',
    id: PANEL_ID,
    order: 20,
    label: panelLabel,
  }, NavIcon));

  ctx.slots.inject('main', () => ctx.slots.register({ name: 'main', key: PANEL_ID }, createPage(load)));

  // Beside the conversation, as a tab of the right sidebar, where the Harness has one.
  ctx.effect(() => {
    let tabs;
    try {
      tabs = ctx.get?.('sidebarRightTabs') ?? ctx.sidebarRightTabs;
    } catch { /* an older Harness */ }
    if (!tabs?.register) return () => {};
    return tabs.register({
      id: SIDE_ID,
      kind: SIDE_KIND,
      keepMounted: true,
      title: panelLabel,
      guide: [{ id: 'documents', order: 40, title: panelLabel, description: () => t('在对话旁查看和编辑表格、PDF、Word') }],
    });
  }, `${PLUGIN_ID}: sidebar tab type`);
  ctx.slots.inject('sidebar.right.pane.tab', () => ctx.slots.register({ name: 'sidebar.right.pane.tab', key: SIDE_ID }, createSideTab(load)));

  // Files opened anywhere in the Harness (links in the conversation, the file tree) come here.
  ctx.effect(() => {
    let previews;
    try {
      previews = ctx.get?.('documentPreviews') ?? ctx.documentPreviews;
    } catch { /* an older Harness */ }
    if (!previews?.register) return () => {};
    const editable = previews.register({ id: EDITOR_ID, extensions: EDITABLE, binaryExtensions: EDITABLE.filter((ext) => BINARY.includes(ext)), priority: 'extension', title: () => t('文档（可编辑）'), loading: 'renderer', wrap: false });
    const readable = previews.register({ id: READER_ID, extensions: READABLE, binaryExtensions: READABLE.filter((ext) => BINARY.includes(ext)), priority: 'fallback', title: () => t('文档插件'), loading: 'renderer', wrap: false });
    return () => {
      editable?.();
      readable?.();
    };
  }, `${PLUGIN_ID}: file previews`);
  const PreviewBody = createPreviewBody(load);
  for (const key of [EDITOR_ID, READER_ID]) {
    ctx.slots.inject('sidebar.right.tab.document', () => ctx.slots.register({ name: 'sidebar.right.tab.document', key }, PreviewBody));
  }

  ctx.slots.inject('shell.overlay', () => ctx.slots.register({
    name: 'shell.overlay',
    id: 'doc-viewer-persistent',
    order: 60,
  }, PersistentViewer));

  const OpenAction = createOpenAction(ctx, load);
  for (const name of ['sidebar.right.tab.document.actions', 'sidebar.right.tab.document.unpreviewable']) {
    ctx.slots.inject(name, () => ctx.slots.register({ name, id: 'doc-viewer-open', order: 60, get label() { return t('在文档中打开'); } }, OpenAction));
  }
}
