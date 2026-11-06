import React from 'react';
import { Marked } from 'marked';
import { Icon } from './icons.jsx';
import { FindBar, Menu, Segmented, TabActions } from './ui.jsx';

const MAX_MARKS = 2000;
const MARKDOWN_EXTENSIONS = ['.md', '.markdown'];

/** @param {string} text */
const escapeHtml = (text) => String(text).replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char]);

/** The parser the PDF export uses too; pictures read the same in both. */
const markdown = new Marked({
  gfm: true,
  renderer: {
    image({ text, href }) {
      return `<span style="color:#86868b">[图片：${escapeHtml(text || href)}]</span>`;
    },
  },
});

export function DocxTab({ tab, actionsEl, onConvert, toolItems }) {
  const srcDoc = React.useMemo(() => paperDocument(tab.doc?.html || ''), [tab.doc?.html]);
  return (
    <>
      <TabActions target={actionsEl}>
        <Menu label="Word 操作" items={[{ label: '提取文字', icon: 'text', onClick: () => onConvert(tab, 'txt') }, '-', ...toolItems]} />
        <button type="button" className="dvd-btn" style={{ margin: '0 4px' }} onClick={() => onConvert(tab, 'pdf')}>导出 PDF…</button>
      </TabActions>
      <div className="dvd-paper">
        <iframe title={tab.name} srcDoc={srcDoc} sandbox="" />
      </div>
    </>
  );
}

/**
 * Word and Markdown content is untrusted: it renders in a sandboxed frame that can run
 * and load nothing, and whose links go nowhere.
 */
function paperDocument(html) {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="color-scheme" content="light"><base target="_blank"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: blob:; style-src 'unsafe-inline'; font-src data:;"><style>
  html,body{margin:0;padding:0;background:#fff;color:#1d1d1f}
  body{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Hiragino Sans GB","Microsoft YaHei",sans-serif;font-size:15px;line-height:1.75;padding:64px 72px 96px;box-sizing:border-box;-webkit-font-smoothing:antialiased}
  h1,h2,h3,h4{line-height:1.3;font-weight:600;letter-spacing:-.01em}
  h1{font-size:26px;margin:0 0 18px}h2{font-size:21px;margin:30px 0 12px}h3{font-size:17px;margin:24px 0 8px}
  p{margin:0 0 12px}table{border-collapse:collapse;width:100%;margin:16px 0;font-size:14px}td,th{border:.5px solid #d2d2d7;padding:7px 9px;vertical-align:top}th{background:#f5f5f7;font-weight:600}
  img{max-width:100%;height:auto}blockquote{margin:14px 0;padding-left:14px;border-left:2px solid #d2d2d7;color:#515154}
  ul,ol{padding-left:26px;margin:8px 0 12px}pre{background:#f5f5f7;padding:12px;border-radius:8px;overflow:auto;font-size:12px}
  code{font-family:ui-monospace,"SF Mono",SFMono-Regular,Menlo,monospace;font-size:.9em}:not(pre)>code{background:#f5f5f7;padding:1px 4px;border-radius:3px}
  hr{border:0;border-top:1px solid #d2d2d7;margin:20px 0}li>input{margin:0 6px 0 -20px}li:has(>input){list-style:none}
  a{color:#4d6bfe;text-decoration:none}
  </style></head><body>${html || '<p style="color:#86868b">这个文档没有可显示的正文。</p>'}</body></html>`;
}

export function TextTab({ tab, actionsEl, pageActive, selected, setDirty, saveText, saveSnapshot, onConvert, registerCommands, toast, toolItems }) {
  // A file only partly loaded, or a dropped copy with no place of its own, can't be saved back.
  const editable = !tab.doc?.truncated && !tab.transient;
  const savedText = React.useRef(tab.doc?.text || '');
  const [text, setText] = React.useState(tab.recoveredText ?? savedText.current);
  const textRef = React.useRef(text);
  const tabRef = React.useRef(tab);
  tabRef.current = tab;
  const [findOpen, setFindOpen] = React.useState(false);
  const [query, setQuery] = React.useState('');
  const [index, setIndex] = React.useState(0);
  const isMarkdown = MARKDOWN_EXTENSIONS.includes(tab.ext);
  const [view, setView] = React.useState(isMarkdown && tab.recoveredText === undefined ? 'preview' : 'source');
  // Find works on the text as written, so it shows the source while it is open.
  const previewing = view === 'preview' && !findOpen;
  const previewDocument = React.useMemo(() => (previewing ? paperDocument(markdown.parse(text)) : ''), [previewing, text]);
  const body = React.useRef(null);

  // Unsaved edits are kept a moment after typing stops, so they survive quitting or a crash.
  const snapshotTimer = React.useRef(null);
  const snapshotNow = React.useCallback(() => {
    clearTimeout(snapshotTimer.current);
    snapshotTimer.current = null;
    if (tabRef.current.dirty) saveSnapshot?.(tabRef.current, { text: textRef.current });
  }, [saveSnapshot]);
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

  const edit = (value) => {
    textRef.current = value;
    setText(value);
    setDirty(tab.id, value !== savedText.current);
    clearTimeout(snapshotTimer.current);
    snapshotTimer.current = setTimeout(snapshotNow, 2000);
  };

  const save = React.useCallback(async () => {
    clearTimeout(snapshotTimer.current);
    snapshotTimer.current = null;
    const saving = textRef.current;
    return saveText(tabRef.current, saving, { onSaved: () => { savedText.current = saving; } });
  }, [saveText]);

  React.useEffect(() => {
    registerCommands(tab.id, { find: () => setFindOpen(true), save: editable ? save : undefined });
  }, [editable, registerCommands, save, tab.id]);

  const exportPdf = async () => {
    // The export reads the file, so what is on screen has to be there first.
    if (tabRef.current.dirty && !(await save())) return;
    onConvert(tabRef.current, 'pdf');
  };

  const positions = React.useMemo(() => {
    const needle = query.toLocaleLowerCase();
    const found = [];
    if (!findOpen || needle.trim() === '') return found;
    const lower = text.toLocaleLowerCase();
    for (let at = lower.indexOf(needle); at >= 0 && found.length < MAX_MARKS; at = lower.indexOf(needle, at + needle.length)) found.push(at);
    return found;
  }, [findOpen, query, text]);

  React.useEffect(() => setIndex(0), [positions]);
  React.useEffect(() => {
    body.current?.querySelector('mark.current')?.scrollIntoView({ block: 'center' });
  }, [index, positions]);

  const content = React.useMemo(() => {
    if (positions.length === 0) return text;
    const parts = [];
    let from = 0;
    positions.forEach((at, order) => {
      parts.push(text.slice(from, at));
      parts.push(<mark key={at} className={order === index ? 'current' : undefined}>{text.slice(at, at + query.length)}</mark>);
      from = at + query.length;
    });
    parts.push(text.slice(from));
    return parts;
  }, [index, positions, query.length, text]);

  const move = (delta) => {
    if (positions.length) setIndex((index + delta + positions.length) % positions.length);
  };

  const copyAll = async () => {
    try {
      await navigator.clipboard.writeText(text);
      toast('已复制全文');
    } catch {
      toast('无法写入剪贴板', true);
    }
  };

  return (
    <>
      <TabActions target={actionsEl}>
        {isMarkdown && <Segmented value={view} onChange={setView} options={[{ value: 'preview', label: '预览' }, { value: 'source', label: editable ? '编辑' : '源码' }]} />}
        <button type="button" className="dvd-iconbtn" title="查找 ⌘F" aria-label="查找" aria-pressed={findOpen} onClick={() => setFindOpen((open) => !open)}><Icon name="search" /></button>
        <Menu label="文本操作" items={[{ label: '复制全文', icon: 'copy', onClick: copyAll }, '-', ...toolItems]} />
        <button type="button" className="dvd-btn" style={{ margin: '0 4px' }} onClick={exportPdf}>导出 PDF…</button>
        {editable && <button type="button" className={`dvd-btn${tab.dirty ? ' primary' : ''}`} style={{ marginRight: 4 }} title="存回原文件，修改前的版本会自动保留 ⌘S" onClick={save}>保存</button>}
      </TabActions>
      {findOpen && (
        <FindBar query={query} onQuery={setQuery} count={positions.length} index={index} onNext={() => move(1)} onPrev={() => move(-1)} onClose={() => { setFindOpen(false); setQuery(''); }} />
      )}
      {previewing
        ? <div className="dvd-paper"><iframe title={tab.name} srcDoc={previewDocument} sandbox="" /></div>
        : editable && !findOpen
          ? <div className="dvd-text editing"><textarea aria-label={`编辑 ${tab.name}`} value={text} spellCheck={false} onChange={(event) => edit(event.target.value)} /></div>
          : <div className="dvd-text" ref={body}><pre>{content}</pre></div>}
    </>
  );
}
