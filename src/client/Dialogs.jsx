import React from 'react';
import { t } from './i18n.js';
import * as api from './api.js';
import { Icon, iconForKind, iconForName } from './icons.jsx';
import { formatSize, relativeTime, Segmented, Sheet } from './ui.jsx';

const PDF_EXTENSION = /\.pdf$/i;

export function EmptyState({ recent, onOpen, onBrowse, onOpenPath, onNew }) {
  return (
    <div className="dvd-empty">
      <div className="dvd-empty-inner">
        <h1>{t('文档')}</h1>
        <p>{t('查看和编辑表格、PDF、Word 与文本，全部在本机处理。')}</p>
        <div className="dvd-empty-actions">
          <button type="button" className="dvd-btn primary" title="⌘O" onClick={onOpen}>{t('打开…')}</button>
          <button type="button" className="dvd-btn" title="⌘N" onClick={onNew}>{t('新建表格')}</button>
          <button type="button" className="dvd-btn plain" onClick={onBrowse}>{t('从工作区选择')}</button>
        </div>
        <ul className="dvd-empty-tips">
          <li><Icon name="save" size={14} /><span><b>{t('拖进来')}</b>{t('把文件拖到这里就能打开')}</span></li>
          <li><Icon name="sidebar" size={14} /><span><b>{t('开在对话旁')}</b>{t('打开后点右上角的按钮，边聊边看')}</span></li>
          <li><Icon name="sparkle" size={14} /><span><b>{t('交给 Agent')}</b>{t('说要改哪里，改动会在表格里亮出来')}</span></li>
        </ul>
        {recent.length > 0 && (
          <>
            <div className="dvd-section-title">{t('最近')}</div>
            {recent.slice(0, 6).map((item) => (
              <button type="button" key={item.path} className="dvd-row-btn" title={item.path} onClick={() => onOpenPath(item.path)}>
                <Icon name={item.kind ? iconForKind(item.kind) : iconForName(item.name)} />
                <span className="dvd-row-name">{item.name}</span>
                <span className="dvd-row-meta">{item.path.split('/').slice(-2, -1)[0]} · {relativeTime(item.at)}</span>
              </button>
            ))}
          </>
        )}
      </div>
    </div>
  );
}

/**
 * The workspace's folders and documents. Choosing a file calls `onPick`;
 * `selected` marks files already chosen when picking several.
 */
function FileBrowser({ onPick, filter, selected }) {
  const [listing, setListing] = React.useState(null);
  const [error, setError] = React.useState(null);
  const [loading, setLoading] = React.useState(true);

  const load = React.useCallback(async (path) => {
    setLoading(true);
    setError(null);
    try {
      setListing(await api.list(path));
    } catch (reason) {
      setError(reason.message);
    } finally {
      setLoading(false);
    }
  }, []);

  React.useEffect(() => { load('.'); }, [load]);

  const entries = (listing?.entries || []).filter((entry) => entry.type === 'directory' || !filter || filter(entry));
  return (
    <>
      <div className="dvd-crumb">
        <button type="button" className="dvd-iconbtn small" disabled={!listing || listing.atRoot} onClick={() => load(listing.parent)} aria-label={t('上一级')}><Icon name="chevronLeft" size={15} /></button>
        <span className="dvd-crumb-path" title={listing?.directory}>{!listing || listing.atRoot ? t('工作区') : listing.relative}</span>
      </div>
      <div className="dvd-list">
        {entries.map((entry) => (
          <button
            type="button"
            key={entry.path}
            className={`dvd-row-btn${selected?.includes(entry.path) ? ' selected' : ''}`}
            onClick={() => (entry.type === 'directory' ? load(entry.path) : onPick(entry.path, entry))}
          >
            <Icon name={entry.type === 'directory' ? 'folder' : iconForName(entry.name)} />
            <span className="dvd-row-name">{entry.name}</span>
            <span className="dvd-row-meta">{entry.type === 'directory' ? '' : formatSize(entry.size)}</span>
            {selected?.includes(entry.path) && <Icon name="check" size={14} />}
          </button>
        ))}
        {!loading && !error && entries.length === 0 && <div className="dvd-empty-note">{t('这里没有可打开的文档')}</div>}
        {loading && !listing && <div className="dvd-empty-note">{t('正在读取…')}</div>}
        {listing?.truncated && <div className="dvd-empty-note">{t('文件太多，只显示了前 500 个')}</div>}
      </div>
      {error && <div className="dvd-callout dvd-bad">{error}</div>}
    </>
  );
}

export function OpenDialog({ onClose, onUpload, onOpenPath }) {
  const [path, setPath] = React.useState('');
  return (
    <Sheet
      title={t('打开文档')}
      onClose={onClose}
      footer={<button type="button" className="dvd-btn" onClick={onUpload}>{t('从电脑选择…')}</button>}
    >
      <FileBrowser onPick={onOpenPath} />
      <input
        className="dvd-input"
        placeholder={t('或输入文件路径，按回车打开')}
        aria-label={t('文件路径')}
        value={path}
        onChange={(event) => setPath(event.target.value)}
        onKeyDown={(event) => { if (event.key === 'Enter' && path.trim()) onOpenPath(path.trim()); }}
      />
    </Sheet>
  );
}

/** The chosen files of a tool, each removable, plus a way to add more from the workspace. */
function FileChooser({ files, onChange, filter, allowMany }) {
  const [browsing, setBrowsing] = React.useState(files.length === 0);
  const toggle = (path) => {
    if (!allowMany) {
      onChange([path]);
      setBrowsing(false);
      return;
    }
    onChange(files.includes(path) ? files.filter((item) => item !== path) : [...files, path]);
  };
  if (browsing) {
    return (
      <>
        <FileBrowser onPick={toggle} filter={filter} selected={files} />
        {files.length > 0 && (
          <button type="button" className="dvd-btn plain" style={{ alignSelf: 'flex-start' }} onClick={() => setBrowsing(false)}>{t('完成选择（{0}）', files.length)}</button>
        )}
      </>
    );
  }
  return (
    <>
      <div className="dvd-list">
        {files.map((path, index) => (
          <div className="dvd-row-btn" key={path} title={path} style={{ cursor: 'default' }}>
            <Icon name={iconForName(path)} />
            <span className="dvd-row-name">{path.split('/').pop()}</span>
            {allowMany && files.length > 1 && (
              <button
                type="button"
                className="dvd-iconbtn small"
                disabled={index === 0}
                aria-label={t('上移')}
                onClick={() => {
                  const next = files.slice();
                  [next[index - 1], next[index]] = [next[index], next[index - 1]];
                  onChange(next);
                }}
              >
                <Icon name="chevronUp" size={14} />
              </button>
            )}
            <button type="button" className="dvd-iconbtn small" aria-label={t('移除 {0}', path.split('/').pop())} onClick={() => onChange(files.filter((item) => item !== path))}><Icon name="close" size={13} /></button>
          </div>
        ))}
      </div>
      <button type="button" className="dvd-btn plain" style={{ alignSelf: 'flex-start' }} onClick={() => setBrowsing(true)}>
        <Icon name="plus" size={14} />{allowMany ? t('添加文件') : t('换一个文件')}
      </button>
    </>
  );
}

const targets = () => [
  { value: 'pdf', label: 'PDF' },
  { value: 'xlsx', label: 'Excel' },
  { value: 'xls', label: t('旧版 Excel') },
  { value: 'csv', label: 'CSV' },
  { value: 'docx', label: 'Word' },
  { value: 'txt', label: t('文本') },
];

/** Which targets a file can become, by its name. */
function targetsFor(path) {
  if (PDF_EXTENSION.test(path)) return ['txt', 'docx', 'xlsx', 'csv'];
  if (/\.(xlsx|xls|csv|tsv|et)$/i.test(path)) return ['pdf', 'xlsx', 'xls', 'csv'];
  if (/\.(docx|doc|wps)$/i.test(path)) return ['pdf', 'txt'];
  return ['pdf', 'txt'];
}

export function ConvertDialog({ tabs, onClose, onOpenPath, toast, track }) {
  const [files, setFiles] = React.useState(() => tabs.filter((tab) => !tab.loading && tab.path).map((tab) => tab.path));
  const [target, setTarget] = React.useState('pdf');
  const [results, setResults] = React.useState(null);
  const [busy, setBusy] = React.useState(false);
  const [lossy, setLossy] = React.useState(null);

  const available = files.length
    ? targets().filter((item) => files.some((path) => targetsFor(path).includes(item.value)))
    : targets();
  const chosen = available.some((item) => item.value === target) ? target : available[0]?.value;

  const start = async (acknowledgeLossy) => {
    setBusy(true);
    setLossy(null);
    try {
      const response = await track(() => api.convert({ paths: files, target: chosen, acknowledgeLossy }));
      if (response.requiresConfirmation) {
        setLossy(response.warnings?.[0]?.message || t('这是有损转换。'));
        return;
      }
      setResults(response.results || []);
    } catch (error) {
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  };

  const copy = async (text) => {
    try {
      await navigator.clipboard.writeText(text);
      toast(t('已复制路径'));
    } catch {
      toast(t('无法写入剪贴板'), true);
    }
  };

  if (results) {
    const succeeded = results.filter((item) => item.ok).length;
    return (
      <Sheet title={t('完成 {0} / {1}', succeeded, results.length)} wide onClose={onClose} footer={<button type="button" className="dvd-btn primary" onClick={onClose}>{t('完成')}</button>}>
        <div className="dvd-list">
          {results.map((result) => (
            <div className="dvd-result" key={result.source}>
              <Icon name={result.ok ? 'check' : 'warning'} size={15} className={result.ok ? 'dvd-ok' : 'dvd-bad'} />
              <span className="dvd-row-name" title={result.ok ? result.output : result.message}>
                {result.ok ? result.output.split('/').pop() : `${result.source.split('/').pop()}：${result.message}`}
              </span>
              {result.ok && <button type="button" className="dvd-btn plain" onClick={() => onOpenPath(result.output)}>{t('打开')}</button>}
              {result.ok && <button type="button" className="dvd-iconbtn small" aria-label={t('复制路径')} title={t('复制路径')} onClick={() => copy(result.output)}><Icon name="copy" size={14} /></button>}
            </div>
          ))}
        </div>
      </Sheet>
    );
  }

  return (
    <Sheet
      title={t('批量转换')}
      wide
      onClose={onClose}
      footer={lossy ? (
        <>
          <button type="button" className="dvd-btn" onClick={() => setLossy(null)}>{t('取消')}</button>
          <button type="button" className="dvd-btn primary" disabled={busy} onClick={() => start(true)}>{t('继续转换')}</button>
        </>
      ) : (
        <button type="button" className="dvd-btn primary" disabled={busy || files.length === 0} onClick={() => start(false)}>{busy ? t('正在转换…') : t('转换 {0} 个文件', files.length || '')}</button>
      )}
    >
      <div className="dvd-field">
        <span className="dvd-label">{t('转换为')}</span>
        <Segmented value={chosen} onChange={setTarget} options={available} />
      </div>
      <div className="dvd-field">
        <span className="dvd-label">{t('文件')}</span>
        <FileChooser files={files} onChange={setFiles} allowMany />
      </div>
      {lossy && <div className="dvd-callout">{lossy}</div>}
    </Sheet>
  );
}

const pdfActions = () => [
  { value: 'merge', label: t('合并') },
  { value: 'split', label: t('拆分') },
  { value: 'delete', label: t('删除页面') },
  { value: 'ocr', label: t('文字识别') },
];

export function PdfToolsDialog({ tabs, activeTab, initialAction = 'merge', initialPath, onClose, onOpenPath, toast, track }) {
  const firstPdf = initialPath || (activeTab?.kind === 'pdf' ? activeTab.path : tabs.find((tab) => tab.kind === 'pdf')?.path);
  const [action, setAction] = React.useState(initialAction);
  const [files, setFiles] = React.useState(firstPdf ? [firstPdf] : []);
  const [pages, setPages] = React.useState('');
  const [ranges, setRanges] = React.useState('');
  const [format, setFormat] = React.useState('txt');
  const [results, setResults] = React.useState(null);
  const [busy, setBusy] = React.useState(false);

  const many = action === 'merge';
  const inputs = many ? files : files.slice(0, 1);
  const ready = many ? inputs.length >= 2 : inputs.length === 1 && (action !== 'delete' || pages.trim() !== '');

  const start = async () => {
    setBusy(true);
    try {
      setResults(await track(() => api.pdf({ action, paths: inputs, pages: pages.trim() || undefined, ranges: ranges.trim() || undefined, format })));
    } catch (error) {
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  };

  if (results) {
    const outputs = results.outputs ? results.outputs.map((item) => item.output) : [results.output];
    return (
      <Sheet title={t('已完成')} onClose={onClose} footer={<button type="button" className="dvd-btn primary" onClick={onClose}>{t('完成')}</button>}>
        <div className="dvd-list">
          {outputs.filter(Boolean).map((output) => (
            <div className="dvd-result" key={output}>
              <Icon name="check" size={15} className="dvd-ok" />
              <span className="dvd-row-name" title={output}>{output.split('/').pop()}</span>
              <button type="button" className="dvd-btn plain" onClick={() => onOpenPath(output)}>{t('打开')}</button>
            </div>
          ))}
        </div>
      </Sheet>
    );
  }

  return (
    <Sheet
      title={t('PDF 工具')}
      onClose={onClose}
      footer={<button type="button" className="dvd-btn primary" disabled={busy || !ready} onClick={start}>{busy ? t('正在处理…') : t('开始')}</button>}
    >
      <Segmented value={action} onChange={setAction} options={pdfActions()} />
      <div className="dvd-field">
        <span className="dvd-label">{many ? t('按顺序合并这些文件') : t('文件')}</span>
        <FileChooser
          key={many ? 'many' : 'one'}
          files={inputs}
          onChange={setFiles}
          allowMany={many}
          filter={action === 'ocr' ? undefined : (entry) => PDF_EXTENSION.test(entry.name)}
        />
      </div>
      {action === 'delete' && (
        <label className="dvd-field">
          <span className="dvd-label">{t('要删除的页码')}</span>
          <input className="dvd-input" value={pages} placeholder={t('例如 2, 4-6')} onChange={(event) => setPages(event.target.value)} />
        </label>
      )}
      {action === 'split' && (
        <label className="dvd-field">
          <span className="dvd-label">{t('拆分范围（留空则每页一个文件）')}</span>
          <input className="dvd-input" value={ranges} placeholder={t('例如 1-3, 4-6')} onChange={(event) => setRanges(event.target.value)} />
        </label>
      )}
      {action === 'ocr' && (
        <>
          <div className="dvd-fieldrow">
            <label className="dvd-field">
              <span className="dvd-label">{t('页码（留空则全部，单次最多 80 页）')}</span>
              <input className="dvd-input" value={pages} placeholder={t('例如 1-20')} onChange={(event) => setPages(event.target.value)} />
            </label>
            <label className="dvd-field">
              <span className="dvd-label">{t('输出为')}</span>
              <select className="dvd-select" value={format} onChange={(event) => setFormat(event.target.value)}>
                <option value="txt">{t('纯文本')}</option>
                <option value="docx">Word</option>
              </select>
            </label>
          </div>
          <p>{t('使用 macOS 的文字识别在本机逐页处理，页数多时需要一些时间。')}</p>
        </>
      )}
    </Sheet>
  );
}
