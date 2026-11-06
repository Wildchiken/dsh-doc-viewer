/** Small controls shared by the document tabs and sheets. */
import React from 'react';
import { t } from './i18n.js';
import { createPortal } from 'react-dom';
import { Icon } from './icons.jsx';

/** A modal card inside the page. Escape closes it unless something inside handled the key. */
export function Sheet({ title, wide, onClose, footer, children }) {
  React.useEffect(() => {
    const onKey = (event) => {
      if (event.key === 'Escape' && !event.defaultPrevented) {
        event.preventDefault();
        onClose();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);
  return (
    <div className="dvd-scrim" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
      <div className={`dvd-card${wide ? ' wide' : ''}`} role="dialog" aria-modal="true" aria-label={title}>
        <div className="dvd-card-head">
          <h3>{title}</h3>
          <button type="button" className="dvd-iconbtn small" onClick={onClose} aria-label={t('关闭')}><Icon name="close" size={14} /></button>
        </div>
        <div className="dvd-card-body">{children}</div>
        {footer && <div className="dvd-card-foot">{footer}</div>}
      </div>
    </div>
  );
}

/** A button that opens a small list of actions. `items` may contain `'-'` separators. */
export function Menu({ label, items, icon = 'more', align = 'right' }) {
  const [open, setOpen] = React.useState(false);
  const wrap = React.useRef(null);
  React.useEffect(() => {
    if (!open) return undefined;
    const close = (event) => { if (!wrap.current?.contains(event.target)) setOpen(false); };
    const onKey = (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', close);
    window.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', close);
      window.removeEventListener('keydown', onKey, true);
    };
  }, [open]);
  return (
    <span className="dvd-menuwrap" ref={wrap}>
      <button type="button" className="dvd-iconbtn" title={label} aria-label={label} aria-haspopup="menu" aria-expanded={open} aria-pressed={open} onClick={() => setOpen((value) => !value)}>
        {icon === 'more' ? <Icon name="more" size={18} strokeWidth={2.4} /> : <Icon name={icon} />}
      </button>
      {open && (
        <div className={`dvd-menu${align === 'left' ? ' left' : ''}`} role="menu">
          {items.filter(Boolean).map((item, index) => (item === '-'
            ? <div key={index} className="dvd-sep" role="separator" />
            : (
              <button type="button" role="menuitem" key={item.label} disabled={item.disabled} onClick={() => { setOpen(false); item.onClick(); }}>
                <Icon name={item.icon} size={15} />
                {item.label}
              </button>
            )))}
        </div>
      )}
    </span>
  );
}

/** A segmented control. @param {{ value: string, options: { value: string, label: string, disabled?: boolean }[], onChange: (value: string) => void }} props */
export function Segmented({ value, options, onChange }) {
  return (
    <div className="dvd-seg">
      {options.map((option) => (
        <button type="button" key={option.value} disabled={option.disabled} aria-pressed={option.value === value} onClick={() => onChange(option.value)}>{option.label}</button>
      ))}
    </div>
  );
}

/**
 * The floating find field used by every document kind. `children` holds extra
 * rows such as replace.
 */
export function FindBar({ query, onQuery, count, index, busy, onNext, onPrev, onClose, placeholder = t('查找'), children }) {
  const input = React.useRef(null);
  React.useEffect(() => {
    input.current?.focus();
    input.current?.select();
  }, []);
  const status = busy ? '…' : query.trim() === '' ? '' : count === 0 ? t('无结果') : `${index + 1} / ${count}${count >= 1000 ? '+' : ''}`;
  return (
    <div className="dvd-find" role="search">
      <div className="dvd-find-row">
        <input
          ref={input}
          className="dvd-input"
          value={query}
          placeholder={placeholder}
          aria-label={placeholder}
          onChange={(event) => onQuery(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              if (event.shiftKey) onPrev();
              else onNext();
            }
            if (event.key === 'Escape') {
              event.preventDefault();
              event.stopPropagation();
              onClose();
            }
          }}
        />
        <span className="dvd-find-count">{status}</span>
        <button type="button" className="dvd-iconbtn small" onClick={onPrev} disabled={!count} aria-label={t('上一个')}><Icon name="chevronUp" size={14} /></button>
        <button type="button" className="dvd-iconbtn small" onClick={onNext} disabled={!count} aria-label={t('下一个')}><Icon name="chevronDown" size={14} /></button>
        <button type="button" className="dvd-iconbtn small" onClick={onClose} aria-label={t('关闭查找')}><Icon name="close" size={14} /></button>
      </div>
      {children}
    </div>
  );
}

/** Put a tab's actions into the top bar while it is the selected one. */
export function TabActions({ target, children }) {
  return target ? createPortal(children, target) : null;
}

/** @param {number} bytes */
export function formatSize(bytes) {
  if (!Number.isFinite(bytes)) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

/** @param {number|undefined} at */
export function relativeTime(at) {
  if (!at) return '';
  const minutes = Math.floor((Date.now() - at) / 60_000);
  if (minutes < 1) return t('刚刚');
  if (minutes < 60) return t('{0} 分钟前', minutes);
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return t('{0} 小时前', hours);
  const days = Math.floor(hours / 24);
  return days < 30 ? t('{0} 天前', days) : new Date(at).toLocaleDateString();
}

/** A quiet outline of what is about to appear, instead of a spinner. */
export function Skeleton({ kind = 'document', label = t('正在打开…') }) {
  if (kind === 'spreadsheet') {
    return (
      <div className="dvd-skeleton grid" role="status" aria-label={label}>
        {Array.from({ length: 40 }, (_, row) => (
          <div className="dvd-skel-row" key={row}>
            {Array.from({ length: 7 }, (_, column) => <span key={column} className={row === 0 || column === 0 ? 'head' : ''} />)}
          </div>
        ))}
      </div>
    );
  }
  return (
    <div className={`dvd-skeleton ${kind === 'pdf' ? 'page' : 'text'}`} role="status" aria-label={label}>
      <div className="dvd-skel-sheet">
        <span className="title" />
        {[92, 100, 84, 96, 70, 100, 88, 60].map((width, index) => <span key={index} style={{ width: `${width}%` }} />)}
      </div>
    </div>
  );
}

/** Which outline a file should show while it opens, from its name alone. */
export function skeletonKind(name = '') {
  const ext = name.slice(name.lastIndexOf('.')).toLowerCase();
  if (['.xlsx', '.xlsm', '.xls', '.csv', '.tsv', '.ods', '.et'].includes(ext)) return 'spreadsheet';
  if (ext === '.pdf') return 'pdf';
  return 'document';
}
