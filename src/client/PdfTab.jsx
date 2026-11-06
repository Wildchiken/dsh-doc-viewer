import React from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import { fileUrl, open as authorize } from './api.js';
import { Icon } from './icons.jsx';
import { FindBar, Menu, Skeleton, TabActions } from './ui.jsx';

const ASSETS = `${window.location.origin}/dsh-doc-viewer/assets`;
pdfjsLib.GlobalWorkerOptions.workerSrc = `${ASSETS}/pdf.worker.min.mjs`;

const PAGE_GAP = 14;
const SCROLL_PADDING_TOP = 20;
const MAX_MATCHES = 1000;
const FIT_MAX_WIDTH = 920;

/**
 * A virtualized PDF viewer: pages are placeholders until they approach the
 * viewport, so a long file does not create hundreds of canvases. Every page has
 * a text layer, so text can be selected and copied, and matches show in place.
 */
export function PdfTab({ tab, actionsEl, onConvert, onPdfTools, registerCommands, onContextChange, toast, toolItems }) {
  const [doc, setDoc] = React.useState(null);
  const [sizes, setSizes] = React.useState([]);
  const [error, setError] = React.useState(null);
  const [attempt, setAttempt] = React.useState(0);
  const [current, setCurrent] = React.useState(1);
  const [zoomMode, setZoomMode] = React.useState('fit');
  const [zoom, setZoom] = React.useState(1);
  const [viewWidth, setViewWidth] = React.useState(0);
  const [thumbsOpen, setThumbsOpen] = React.useState(false);
  const [findOpen, setFindOpen] = React.useState(false);
  const [highlight, setHighlight] = React.useState('');
  const [scrollEl, setScrollEl] = React.useState(null);
  const pageEls = React.useRef([]);

  React.useEffect(() => {
    let cancelled = false;
    let task;
    let loaded;
    setError(null);
    setDoc(null);
    const load = () => fetch(fileUrl(tab.path), { credentials: 'same-origin' });
    // After the Host restarts it no longer knows this file was opened; open it again.
    load()
      .then(async (response) => (response.status === 403 ? authorize(tab.path).then(load) : response))
      .then((response) => {
        if (!response.ok) throw new Error(`无法读取文件（HTTP ${response.status}）`);
        return response.arrayBuffer();
      })
      .then((data) => {
        if (cancelled) return undefined;
        task = pdfjsLib.getDocument({
          data: new Uint8Array(data),
          standardFontDataUrl: `${ASSETS}/standard_fonts/`,
          cMapUrl: `${ASSETS}/cmaps/`,
          cMapPacked: true,
          wasmUrl: `${ASSETS}/wasm/`,
          iccUrl: `${ASSETS}/iccs/`,
          isEvalSupported: false,
        });
        return task.promise;
      })
      .then(async (pdf) => {
        if (cancelled || !pdf) return;
        loaded = pdf;
        // Assume every page matches the first; each page corrects itself when it renders.
        const viewport = (await pdf.getPage(1)).getViewport({ scale: 1 });
        if (cancelled) return;
        setSizes(Array.from({ length: pdf.numPages }, () => ({ width: viewport.width, height: viewport.height })));
        setDoc(pdf);
      })
      .catch((reason) => {
        if (cancelled) return;
        setError(reason?.name === 'PasswordException' || /password/i.test(String(reason?.message))
          ? '这个 PDF 已加密。请先解除密码后再打开。'
          : `无法打开：${reason?.message || String(reason)}`);
      });
    return () => {
      cancelled = true;
      try { task?.destroy?.(); } catch { /* ignore */ }
      try { loaded?.destroy?.(); } catch { /* ignore */ }
    };
  }, [attempt, tab.path]);

  React.useEffect(() => {
    if (!scrollEl) return undefined;
    const observer = new ResizeObserver(([entry]) => {
      if (entry.contentRect.width > 0) setViewWidth(entry.contentRect.width);
    });
    observer.observe(scrollEl);
    return () => observer.disconnect();
  }, [scrollEl]);

  const widest = sizes.reduce((max, size) => Math.max(max, size.width), 0) || 595;
  // Like Preview: fill the width, but not beyond a comfortable reading size.
  const fitScale = viewWidth ? Math.max(0.2, Math.min(viewWidth, FIT_MAX_WIDTH) / widest) : 1;
  const scale = zoomMode === 'fit' ? fitScale : zoom;

  // Keep the same place in the document when the scale changes.
  const previousScale = React.useRef(scale);
  React.useLayoutEffect(() => {
    if (scrollEl && previousScale.current !== scale) {
      scrollEl.scrollTop *= scale / previousScale.current;
      previousScale.current = scale;
    }
  }, [scale, scrollEl]);

  const goPage = React.useCallback((page) => {
    const safe = Math.max(1, Math.min(doc?.numPages || 1, Number(page) || 1));
    setCurrent(safe);
    pageEls.current[safe - 1]?.scrollIntoView({ block: 'start', behavior: 'auto' });
  }, [doc]);

  const onScroll = () => {
    if (!scrollEl || sizes.length === 0) return;
    const middle = scrollEl.scrollTop + scrollEl.clientHeight / 3 - SCROLL_PADDING_TOP;
    let bottom = 0;
    for (let index = 0; index < sizes.length; index += 1) {
      bottom += sizes[index].height * scale + PAGE_GAP;
      if (middle < bottom) {
        if (current !== index + 1) setCurrent(index + 1);
        return;
      }
    }
  };

  const zoomBy = (delta) => {
    setZoom(Math.max(0.25, Math.min(4, Math.round((scale + delta) * 20) / 20)));
    setZoomMode('custom');
  };

  React.useEffect(() => {
    registerCommands(tab.id, {
      find: () => setFindOpen(true),
      describe: () => ({ page: currentRef.current, pageCount: doc?.numPages }),
    });
    onContextChange?.();
  }, [doc, onContextChange, registerCommands, tab.id]);

  const currentRef = React.useRef(current);
  currentRef.current = current;
  React.useEffect(() => { onContextChange?.(); }, [current, onContextChange]);

  const onPageSize = React.useCallback((index, size) => {
    setSizes((previous) => {
      const known = previous[index];
      if (known && Math.abs(known.width - size.width) < 0.5 && Math.abs(known.height - size.height) < 0.5) return previous;
      const next = previous.slice();
      next[index] = size;
      return next;
    });
  }, []);

  return (
    <div className="dvd-pdf">
      <TabActions target={actionsEl}>
        <button type="button" className="dvd-iconbtn" title="查找 ⌘F" aria-label="查找" aria-pressed={findOpen} onClick={() => setFindOpen((open) => !open)}><Icon name="search" /></button>
        <button type="button" className="dvd-iconbtn" title="缩略图" aria-label="缩略图" aria-pressed={thumbsOpen} onClick={() => setThumbsOpen((open) => !open)}><Icon name="thumbs" /></button>
        <Menu
          label="PDF 操作"
          items={[
            { label: '提取文字', icon: 'text', onClick: () => onConvert(tab, 'txt') },
            { label: '转为 Word', icon: 'word', onClick: () => onConvert(tab, 'docx') },
            { label: '转为 Excel', icon: 'sheet', onClick: () => onConvert(tab, 'xlsx') },
            '-',
            { label: '合并…', icon: 'merge', onClick: () => onPdfTools('merge', tab.path) },
            { label: '拆分…', icon: 'convert', onClick: () => onPdfTools('split', tab.path) },
            { label: '删除页面…', icon: 'minus', onClick: () => onPdfTools('delete', tab.path) },
            { label: '文字识别 (OCR)…', icon: 'eye', onClick: () => onPdfTools('ocr', tab.path) },
            '-',
            { label: '批量转换…', icon: 'convert', onClick: toolItems[0].onClick },
          ]}
        />
      </TabActions>
      {thumbsOpen && doc && (
        <div className="dvd-pdf-thumbs">
          {Array.from({ length: doc.numPages }, (_, index) => (
            <PdfThumb key={index} doc={doc} pageNumber={index + 1} active={current === index + 1} onClick={() => goPage(index + 1)} />
          ))}
        </div>
      )}
      <div className="dvd-pdf-main">
        {findOpen && doc && (
          <PdfFind doc={doc} onClose={() => { setFindOpen(false); setHighlight(''); }} goPage={goPage} onQuery={setHighlight} toast={toast} />
        )}
        {error && (
          <div className="dvd-center">
            <h3>这个 PDF 打不开</h3>
            <p>{error}</p>
            <button type="button" className="dvd-btn" onClick={() => setAttempt((value) => value + 1)}>重试</button>
          </div>
        )}
        {!error && !doc && <Skeleton kind="pdf" />}
        {!error && doc && (
          <>
            <div className="dvd-pdf-scroll" ref={setScrollEl} onScroll={onScroll}>
              <div className="dvd-pdf-pages">
                {scrollEl && viewWidth > 0 && sizes.map((size, index) => (
                  <PdfPage
                    key={index}
                    doc={doc}
                    index={index}
                    size={size}
                    scale={scale}
                    scrollRoot={scrollEl}
                    highlight={highlight}
                    register={pageEls}
                    onSize={onPageSize}
                  />
                ))}
              </div>
            </div>
            <div className="dvd-hud">
              <button type="button" className="dvd-iconbtn" onClick={() => goPage(current - 1)} disabled={current <= 1} aria-label="上一页"><Icon name="chevronLeft" size={15} /></button>
              <PageInput current={current} total={doc.numPages} onGo={goPage} />
              <button type="button" className="dvd-iconbtn" onClick={() => goPage(current + 1)} disabled={current >= doc.numPages} aria-label="下一页"><Icon name="chevronRight" size={15} /></button>
              <span className="dvd-hud-sep" />
              <button type="button" className="dvd-iconbtn" onClick={() => zoomBy(-0.1)} aria-label="缩小"><Icon name="minus" size={15} /></button>
              <span className="dvd-hud-text">{Math.round(scale * 100)}%</span>
              <button type="button" className="dvd-iconbtn" onClick={() => zoomBy(0.1)} aria-label="放大"><Icon name="plus" size={15} /></button>
              <button type="button" className="dvd-iconbtn" aria-pressed={zoomMode === 'fit'} onClick={() => setZoomMode('fit')} title="适应宽度" aria-label="适应宽度"><Icon name="fit" size={15} /></button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/** Typing is free; the jump happens on Enter or when leaving the field. */
function PageInput({ current, total, onGo }) {
  const [draft, setDraft] = React.useState(null);
  const commit = () => {
    if (draft !== null && draft.trim() !== '') onGo(draft);
    setDraft(null);
  };
  return (
    <span className="dvd-hud-text" style={{ display: 'inline-flex', alignItems: 'center', gap: 3 }}>
      <input
        className="dvd-pageinput"
        inputMode="numeric"
        aria-label="页码"
        value={draft ?? current}
        onChange={(event) => setDraft(event.target.value.replace(/\D/g, ''))}
        onBlur={commit}
        onKeyDown={(event) => {
          if (event.key === 'Enter') commit();
          if (event.key === 'Escape') {
            event.stopPropagation();
            setDraft(null);
          }
        }}
      />
      / {total}
    </span>
  );
}

function PdfPage({ doc, index, size, scale, scrollRoot, highlight, register, onSize }) {
  const pageNumber = index + 1;
  const hostRef = React.useRef(null);
  const canvasRef = React.useRef(null);
  const textRef = React.useRef(null);
  const [near, setNear] = React.useState(false);
  const [textReady, setTextReady] = React.useState(0);

  React.useEffect(() => {
    const element = hostRef.current;
    register.current[index] = element;
    const observer = new IntersectionObserver(([entry]) => setNear(entry.isIntersecting), { root: scrollRoot, rootMargin: '900px 0px' });
    observer.observe(element);
    return () => observer.disconnect();
  }, [index, register, scrollRoot]);

  React.useEffect(() => {
    if (!near) return undefined;
    let cancelled = false;
    let renderTask;
    let textLayer;
    (async () => {
      const page = await doc.getPage(pageNumber);
      if (cancelled) return;
      const base = page.getViewport({ scale: 1 });
      onSize(index, { width: base.width, height: base.height });
      const viewport = page.getViewport({ scale });
      const ratio = Math.min(window.devicePixelRatio || 1, 2);
      const canvas = canvasRef.current;
      if (!canvas) return;
      canvas.width = Math.max(1, Math.floor(viewport.width * ratio));
      canvas.height = Math.max(1, Math.floor(viewport.height * ratio));
      canvas.style.width = `${viewport.width}px`;
      canvas.style.height = `${viewport.height}px`;
      renderTask = page.render({
        canvas,
        viewport,
        transform: ratio === 1 ? undefined : [ratio, 0, 0, ratio, 0, 0],
        background: 'rgb(255,255,255)',
      });
      await renderTask.promise;
      if (cancelled || !textRef.current) return;
      const container = textRef.current;
      container.replaceChildren();
      container.style.setProperty('--total-scale-factor', String(scale));
      textLayer = new pdfjsLib.TextLayer({ textContentSource: page.streamTextContent(), container, viewport });
      await textLayer.render();
      if (!cancelled) setTextReady((value) => value + 1);
    })().catch((reason) => {
      if (reason?.name !== 'RenderingCancelledException' && !cancelled) console.warn('[doc-viewer] page render failed', reason);
    });
    return () => {
      cancelled = true;
      try { renderTask?.cancel?.(); } catch { /* ignore */ }
      try { textLayer?.cancel?.(); } catch { /* ignore */ }
    };
  }, [doc, index, near, onSize, pageNumber, scale]);

  React.useEffect(() => {
    if (textRef.current) markMatches(textRef.current, highlight);
  }, [highlight, textReady]);

  return (
    <div className="dvd-pdf-page" ref={hostRef} style={{ width: Math.round(size.width * scale), height: Math.round(size.height * scale) }} data-page={pageNumber}>
      {near && <canvas ref={canvasRef} />}
      {near && <div className="textLayer" ref={textRef} />}
    </div>
  );
}

/** Wrap matches in a page's text spans; the spans are transparent over the canvas. */
function markMatches(container, query) {
  const needle = query.trim().toLocaleLowerCase();
  for (const span of container.querySelectorAll('span[role="presentation"], span:not([class])')) {
    const text = span.textContent;
    if (span.querySelector('mark')) span.textContent = text;
    if (!needle || !text.toLocaleLowerCase().includes(needle)) continue;
    const fragment = document.createDocumentFragment();
    const lower = text.toLocaleLowerCase();
    let from = 0;
    for (let at = lower.indexOf(needle); at >= 0; at = lower.indexOf(needle, from)) {
      fragment.append(text.slice(from, at));
      const mark = document.createElement('mark');
      mark.textContent = text.slice(at, at + needle.length);
      fragment.append(mark);
      from = at + needle.length;
    }
    fragment.append(text.slice(from));
    span.replaceChildren(fragment);
  }
}

function PdfThumb({ doc, pageNumber, active, onClick }) {
  const hostRef = React.useRef(null);
  const canvasRef = React.useRef(null);
  const [visible, setVisible] = React.useState(false);
  React.useEffect(() => {
    const element = hostRef.current;
    const observer = new IntersectionObserver(([entry]) => { if (entry.isIntersecting) setVisible(true); }, { root: element.closest('.dvd-pdf-thumbs'), rootMargin: '300px' });
    observer.observe(element);
    return () => observer.disconnect();
  }, []);
  React.useEffect(() => {
    if (active) hostRef.current?.scrollIntoView({ block: 'nearest' });
  }, [active]);
  React.useEffect(() => {
    if (!visible) return undefined;
    let cancelled = false;
    let task;
    (async () => {
      const page = await doc.getPage(pageNumber);
      if (cancelled || !canvasRef.current) return;
      const viewport = page.getViewport({ scale: 0.32 });
      const canvas = canvasRef.current;
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      task = page.render({ canvas, viewport });
      await task.promise;
    })().catch(() => {});
    return () => {
      cancelled = true;
      try { task?.cancel?.(); } catch { /* ignore */ }
    };
  }, [doc, pageNumber, visible]);
  return (
    <button type="button" ref={hostRef} className={`dvd-thumb${active ? ' active' : ''}`} onClick={onClick} aria-label={`第 ${pageNumber} 页`}>
      <canvas ref={canvasRef} />
      {pageNumber}
    </button>
  );
}

function PdfFind({ doc, onClose, goPage, onQuery, toast }) {
  const [query, setQuery] = React.useState('');
  const [matches, setMatches] = React.useState([]);
  const [index, setIndex] = React.useState(0);
  const [busy, setBusy] = React.useState(false);

  React.useEffect(() => {
    const needle = query.trim().toLocaleLowerCase();
    onQuery(query);
    setMatches([]);
    setIndex(0);
    if (!needle) {
      setBusy(false);
      return undefined;
    }
    let cancelled = false;
    setBusy(true);
    const timer = setTimeout(async () => {
      const found = [];
      try {
        for (let pageNumber = 1; pageNumber <= doc.numPages && found.length < MAX_MATCHES; pageNumber += 1) {
          const page = await doc.getPage(pageNumber);
          const content = await page.getTextContent();
          if (cancelled) return;
          const lower = content.items.map((item) => item.str ?? '').join(' ').toLocaleLowerCase();
          for (let at = lower.indexOf(needle); at >= 0 && found.length < MAX_MATCHES; at = lower.indexOf(needle, at + needle.length)) {
            found.push(pageNumber);
          }
          // Show results as they arrive, and jump to the first one straight away.
          if (found.length > 0 && pageNumber % 8 === 0) setMatches(found.slice());
          if (found.length > 0 && found[0] === pageNumber && found.length === 1) goPage(pageNumber);
        }
        if (!cancelled) {
          setMatches(found);
          if (found.length) goPage(found[0]);
        }
      } catch (error) {
        if (!cancelled) toast(`查找失败：${error.message}`, true);
      } finally {
        if (!cancelled) setBusy(false);
      }
    }, 200);
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [doc, goPage, onQuery, query, toast]);

  const move = (delta) => {
    if (matches.length === 0) return;
    const next = (index + delta + matches.length) % matches.length;
    setIndex(next);
    goPage(matches[next]);
  };

  return <FindBar query={query} onQuery={setQuery} count={matches.length} index={index} busy={busy} onNext={() => move(1)} onPrev={() => move(-1)} onClose={onClose} />;
}
