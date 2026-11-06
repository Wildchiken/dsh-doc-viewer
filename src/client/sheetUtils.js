/** Client-side sheet helpers: reading the grid's matrix, statistics and the chart. */
import { getBorderInfoCompute } from '@fortune-sheet/core';
import { lang, t } from './i18n.js';

/** @param {number} index */
export function colName(index) {
  let value = index + 1;
  let name = '';
  while (value > 0) {
    const rem = (value - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

/** @param {any} cell */
export function displayOf(cell) {
  if (!cell) return '';
  if (cell.m !== undefined && cell.m !== null && cell.m !== '') return String(cell.m);
  if (Array.isArray(cell.ct?.s)) return cell.ct.s.map((run) => run?.v ?? '').join('');
  if (cell.v === undefined || cell.v === null) return '';
  return String(cell.v);
}

/**
 * Text that is unambiguously a number. Identifiers such as `00123`, and phone
 * or ID numbers with more digits than a double holds, stay text. Mirrors the
 * rule the Host uses.
 * @param {unknown} value
 * @returns {number|undefined}
 */
export function parseNumericText(value) {
  if (typeof value !== 'string') return undefined;
  const text = value.trim().replaceAll(',', '').replaceAll('，', '');
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) return undefined;
  if (/^[+-]?0\d/.test(text)) return undefined;
  if (text.replace(/[^\d]/g, '').length > 15) return undefined;
  const number = Number(text);
  return Number.isFinite(number) ? number : undefined;
}

/** @param {any} cell */
export function numberOf(cell) {
  if (!cell) return undefined;
  if (typeof cell.v === 'number') return Number.isFinite(cell.v) ? cell.v : undefined;
  return parseNumericText(cell.v);
}

/**
 * Whether the first row of a rectangle names its columns rather than holding data.
 * Mirrors the rule the Host uses.
 * @param {any[][]} cells
 * @param {number} startRow the sheet row it starts on, from 0
 */
export function firstRowIsHeader(cells, startRow) {
  if (cells.length < 2) return false;
  const isNumber = (cell) => numberOf(cell) !== undefined;
  const isFilled = (cell) => cell && cell.v !== undefined && cell.v !== null && cell.v !== '';
  const first = cells[0].filter(isFilled);
  if (first.length === 0 || first.every(isNumber)) return false;
  if (startRow === 0) return true;
  if (first.some((cell) => isNumber(cell) || cell.f)) return false;
  return cells.slice(1).some((row) => row.some((cell, index) => isFilled(cells[0][index]) && isNumber(cell)));
}

/**
 * The grid keeps a matrix once loaded. Saving wants the sparse form, without the
 * grid's transient state.
 * @param {any[]} sheets from `workbook.getAllSheets()`
 */
export function sheetsForSaving(sheets) {
  return sheets.map((sheet) => {
    const celldata = [];
    (sheet.data || []).forEach((row, r) => row?.forEach((v, c) => {
      if (v !== null && v !== undefined) celldata.push({ r, c, v });
    }));
    return {
      name: sheet.name,
      id: sheet.id,
      order: sheet.order,
      config: withoutFilteredRows(sheet),
      frozen: sheet.frozen,
      images: sheet.images,
      celldata: sheet.data ? celldata : sheet.celldata || [],
      borders: resolvedBorders(sheets, sheet),
      // Not needed in a file, but a snapshot reopened in the grid needs them to calculate and size.
      calcChain: sheet.calcChain,
      row: sheet.data?.length,
      column: sheet.data?.[0]?.length,
    };
  });
}

/**
 * A filter in the grid hides rows to show matches. That is a view, not a change
 * to the file: saving must not leave those rows hidden.
 */
function withoutFilteredRows(sheet) {
  const filtered = new Set(Object.values(sheet.filter || {}).flatMap((column) => Object.keys(column?.rowhidden || {})));
  if (filtered.size === 0 || !sheet.config?.rowhidden) return sheet.config;
  const rowhidden = Object.fromEntries(Object.entries(sheet.config.rowhidden).filter(([row]) => !filtered.has(row)));
  return { ...sheet.config, rowhidden };
}

/**
 * Every cell's borders as the grid draws them. The grid stores borders as the
 * commands that drew them (outline a range, all inner lines, …); the file wants
 * them per cell.
 * @returns {Record<string, any>|undefined}
 */
function resolvedBorders(sheets, sheet) {
  if (!sheet.data) return undefined;
  try {
    return getBorderInfoCompute({ luckysheetfile: sheets, currentSheetId: sheet.id, config: sheet.config || {} }, sheet.id) || {};
  } catch (error) {
    console.warn('[doc-viewer] could not read borders; they are left as they are in the file', error);
    return undefined;
  }
}

/** The bottom-right corner holding any value. @param {any[][]|undefined} data */
export function usedBounds(data) {
  let maxRow = -1;
  let maxColumn = -1;
  (data || []).forEach((row, r) => row?.forEach((cell, c) => {
    if (cell && (cell.v !== undefined && cell.v !== null && cell.v !== '' || cell.f || cell.ct?.s)) {
      if (r > maxRow) maxRow = r;
      if (c > maxColumn) maxColumn = c;
    }
  }));
  return { maxRow, maxColumn };
}

/**
 * @param {any[][]} cells a rectangle of cells
 * @param {{ hasHeader: boolean, groupIndex: number, valueIndex: number, agg: string }} options
 */
export function computeSheetStats(cells, { hasHeader, groupIndex, valueIndex, agg }) {
  const body = hasHeader ? cells.slice(1) : cells;
  const numbers = [];
  const values = [];
  for (const row of body) {
    for (const cell of valueIndex >= 0 ? [row[valueIndex]] : row) {
      if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;
      const number = numberOf(cell);
      if (number !== undefined) numbers.push(number);
      values.push(displayOf(cell));
    }
  }
  const sum = numbers.reduce((total, value) => total + value, 0);
  const summary = {
    count: values.length,
    numericCount: numbers.length,
    distinct: new Set(values).size,
    sum: round(sum),
    average: numbers.length ? round(sum / numbers.length) : null,
    min: numbers.length ? numbers.reduce((min, value) => Math.min(min, value)) : null,
    max: numbers.length ? numbers.reduce((max, value) => Math.max(max, value)) : null,
  };
  let groups = [];
  if (groupIndex >= 0 && valueIndex >= 0) {
    const buckets = new Map();
    for (const row of body) {
      const key = displayOf(row[groupIndex]) || t('(空)');
      const value = numberOf(row[valueIndex]);
      if (value === undefined) continue;
      if (!buckets.has(key)) buckets.set(key, []);
      buckets.get(key).push(value);
    }
    groups = [...buckets.entries()].map(([key, list]) => {
      const total = list.reduce((acc, value) => acc + value, 0);
      const value = agg === 'count' ? list.length
        : agg === 'average' ? round(total / list.length)
        : agg === 'max' ? list.reduce((max, item) => Math.max(max, item))
        : agg === 'min' ? list.reduce((min, item) => Math.min(min, item))
        : round(total);
      return { key, value, count: list.length };
    }).sort((a, b) => b.value - a.value);
  }
  return { summary, groups };
}

/** Names of the aggregations, in the interface language. */
export const aggLabels = () => ({ sum: t('合计'), average: t('平均'), count: t('计数'), max: t('最大'), min: t('最小') });

/** @param {number} value */
export function formatNumber(value) {
  if (value === null || value === undefined) return '—';
  return typeof value === 'number' ? value.toLocaleString(undefined, { maximumFractionDigits: 4 }) : String(value);
}

/** Build a sheet holding the statistics table, ready for the grid or for saving. */
export function buildStatsSheet({ summary, groups, agg, groupName, valueName }, name) {
  const rows = [
    [t('指标'), t('值')],
    [t('合计'), summary.sum],
    [t('平均'), summary.average],
    [t('最大'), summary.max],
    [t('最小'), summary.min],
    [t('计数'), summary.count],
    [t('去重计数'), summary.distinct],
  ];
  if (groups.length) {
    rows.push([], [groupName || t('分组'), t('{0}（{1}）', valueName || t('数值'), aggLabels()[agg])]);
    for (const group of groups) rows.push([group.key, group.value]);
  }
  const celldata = rows.flatMap((row, r) => row.map((value, c) => {
    const numeric = typeof value === 'number';
    return {
      r,
      c,
      v: {
        v: value ?? '',
        m: numeric ? formatNumber(value) : String(value ?? ''),
        ct: { fa: 'General', t: numeric ? 'n' : 's' },
        ...(r === 0 ? { bl: 1 } : {}),
      },
    };
  }));
  return {
    name,
    id: `stats-${Date.now()}-${Math.floor(Math.random() * 1e6)}`,
    row: Math.max(rows.length + 40, 120),
    column: 24,
    celldata,
    config: { columnlen: { 0: 160, 1: 160 } },
    status: 0,
    showGridLines: true,
  };
}

/** @param {any[]} sheets @param {string} base */
export function uniqueSheetName(sheets, base) {
  const used = new Set(sheets.map((sheet) => sheet.name));
  if (!used.has(base)) return base;
  for (let index = 2; ; index += 1) {
    if (!used.has(`${base} ${index}`)) return `${base} ${index}`;
  }
}

const CHART_FONT = '-apple-system,BlinkMacSystemFont,PingFang SC,sans-serif';

/**
 * A chart in the page's one accent colour on white paper, which also serves as
 * the image inserted into the workbook.
 */
export function renderChartSvg(groups, title, type = 'bar', width = 620, height = 260) {
  const data = (groups || []).slice(0, 16);
  if (data.length === 0) return '';
  const accent = '#4d6bfe';
  const padding = { left: 52, right: 20, top: 40, bottom: 38 };
  const innerW = width - padding.left - padding.right;
  const innerH = height - padding.top - padding.bottom;
  const max = Math.max(...data.map((item) => item.value), 0);
  const min = Math.min(0, ...data.map((item) => item.value));
  const span = max - min || 1;
  const y = (value) => padding.top + innerH - ((value - min) / span) * innerH;
  const parts = [`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" font-family="${CHART_FONT}">`];
  parts.push(`<rect width="${width}" height="${height}" fill="#ffffff"/>`);
  parts.push(`<text x="20" y="24" fill="#0f1115" font-size="13" font-weight="600">${escapeXml(title)}</text>`);
  for (let tick = 0; tick <= 3; tick += 1) {
    const value = min + (span * tick) / 3;
    const yy = y(value);
    parts.push(`<line x1="${padding.left}" y1="${yy}" x2="${width - padding.right}" y2="${yy}" stroke="#0f1115" stroke-opacity=".07"/>`);
    parts.push(`<text x="${padding.left - 8}" y="${yy + 3.5}" text-anchor="end" fill="#81858c" font-size="10">${formatTick(value)}</text>`);
  }
  const step = innerW / data.length;
  if (type === 'line') {
    const at = (index) => padding.left + step * index + step / 2;
    parts.push(`<polyline points="${data.map((item, index) => `${at(index)},${y(item.value)}`).join(' ')}" fill="none" stroke="${accent}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`);
    data.forEach((item, index) => parts.push(`<circle cx="${at(index)}" cy="${y(item.value)}" r="3" fill="#ffffff" stroke="${accent}" stroke-width="2"/>`));
  } else {
    const bar = Math.min(36, step * 0.56);
    data.forEach((item, index) => {
      const top = Math.min(y(item.value), y(0));
      const barHeight = Math.max(1.5, Math.abs(y(item.value) - y(0)));
      parts.push(`<rect x="${padding.left + step * index + (step - bar) / 2}" y="${top}" width="${bar}" height="${barHeight}" rx="3" fill="${accent}"/>`);
    });
  }
  data.forEach((item, index) => {
    parts.push(`<text x="${padding.left + step * index + step / 2}" y="${height - 16}" text-anchor="middle" fill="#61666b" font-size="10.5">${escapeXml(String(item.key).slice(0, 8))}</text>`);
  });
  parts.push('</svg>');
  return parts.join('');
}

/** @param {number} value */
function formatTick(value) {
  // Chinese counts in 万 (10,000); English in thousands.
  if (Math.abs(value) >= 10000) return lang() === 'zh' ? `${(value / 10000).toFixed(1)}万` : `${Math.round(value / 1000)}k`;
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

/** @param {number} value */
function round(value) {
  return Math.round(value * 1e6) / 1e6;
}

/** @param {string} value */
function escapeXml(value) {
  return String(value).replace(/[<>&"']/g, (char) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' }[char]));
}

/** SVG -> PNG data URL for inserting a chart image into the sheet. */
export async function svgToPngDataUrl(svg, width = 620, height = 260) {
  if (!svg) return undefined;
  const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
  try {
    const image = new Image();
    image.src = url;
    await new Promise((resolve, reject) => {
      image.onload = resolve;
      image.onerror = reject;
    });
    const canvas = document.createElement('canvas');
    const ratio = Math.min(window.devicePixelRatio || 1, 2);
    canvas.width = Math.round(width * ratio);
    canvas.height = Math.round(height * ratio);
    const context = canvas.getContext('2d');
    context.scale(ratio, ratio);
    context.drawImage(image, 0, 0, width, height);
    return canvas.toDataURL('image/png');
  } finally {
    URL.revokeObjectURL(url);
  }
}
