/**
 * Draw the README illustration (docs/hero.svg and docs/hero.zh-CN.svg): a conversation
 * beside the open workbook, with the cells the agent changed outlined the way the plugin
 * outlines them. The numbers are those of a real run on the same table.
 *   node build/hero.mjs
 */
import { writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const ACCENT = '#4d6bfe';
const FONT = "-apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', 'Noto Sans CJK SC', 'Noto Sans SC', Helvetica, Arial, sans-serif";

const TEXT = {
  en: {
    title: 'Raise East prices',
    ask: ['East prices are too low. Raise them 10%,', 'rounded to tens. How much does the total go up?'],
    step: 'Edited 9 cells in place',
    answer: ['Done. All 9 East prices are up 10%,', 'rounded to tens. Nothing else changed.'],
    before: 'Before', after: 'After', change: 'Change',
    file: 'Q3 sales.xlsx', chip: '9', save: 'Save',
    head: ['Region', 'Product', 'Qty', 'Price', 'Amount'],
    regions: { e: 'East', n: 'North', s: 'South', w: 'Southwest' },
    products: ['Smartwatch Pro', 'ANC Earbuds', 'Portable Display', 'Keyboard'],
    toast: 'Agent changed 9 cells', view: 'View',
    input: 'Ask anything, @ files',
  },
  zh: {
    title: '华东单价上调 10%',
    ask: ['华东的单价定低了，都涨 10%，', '取整到十位。总额多了多少？'],
    step: '原地修改了 9 个单元格',
    answer: ['改好了：华东 9 行单价都上调 10%，', '取整到十位，其他区域没动。'],
    before: '改前', after: '改后', change: '变化',
    file: '季度销售.xlsx', chip: '9', save: '保存',
    head: ['区域', '产品', '数量', '单价', '金额'],
    regions: { e: '华东', n: '华北', s: '华南', w: '西南' },
    products: ['智能手表 Pro', '降噪耳机', '便携显示器', '机械键盘'],
    toast: 'Agent 改了 9 个单元格', view: '查看',
    input: '描述任务，@ 引用文件',
  },
};

// Rows 2-11 of the demo workbook after the agent's edit: region, product, quantity, old price, new price.
const ROWS = [
  ['e', 0, 12, 899, 990], ['n', 1, 29, 459], ['s', 2, 46, 1299], ['e', 3, 63, 349, 380], ['w', 0, 19, 899],
  ['n', 1, 36, 459], ['e', 2, 53, 1299, 1430], ['s', 3, 70, 349], ['e', 0, 26, 899, 990], ['n', 1, 43, 459],
];

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const num = (n, digits = 0) => n.toLocaleString('en-US', { minimumFractionDigits: digits, maximumFractionDigits: digits });
const text = (x, y, s, attrs = '') => `<text x="${x}" y="${y}" ${attrs}>${esc(s)}</text>`;

function hero(lang) {
  const t = TEXT[lang];
  const W = 1200, H = 680;
  const win = { x: 24, y: 20, w: W - 48, h: H - 48 };
  const split = win.x + 470;
  const out = [];
  out.push(`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT}" font-size="15" fill="#1f2329">`);
  out.push(`<defs>
  <filter id="shadow" x="-5%" y="-5%" width="110%" height="115%"><feDropShadow dx="0" dy="10" stdDeviation="14" flood-color="#1f2a44" flood-opacity="0.12"/></filter>
  <clipPath id="win"><rect x="${win.x}" y="${win.y}" width="${win.w}" height="${win.h}" rx="16"/></clipPath>
</defs>`);
  // Window
  out.push(`<rect x="${win.x}" y="${win.y}" width="${win.w}" height="${win.h}" rx="16" fill="#fff" stroke="#e4e7ec" filter="url(#shadow)"/>`);
  out.push(`<g clip-path="url(#win)">`);
  out.push(`<rect x="${win.x}" y="${win.y}" width="${split - win.x}" height="${win.h}" fill="#fafbfc"/>`);
  out.push(`<line x1="${split}" y1="${win.y}" x2="${split}" y2="${win.y + win.h}" stroke="#e8eaee"/>`);

  // --- Conversation ---
  const cx = win.x + 32;
  out.push(text(cx, win.y + 42, t.title, 'font-size="16" font-weight="600"'));
  out.push(`<line x1="${win.x}" y1="${win.y + 64}" x2="${split}" y2="${win.y + 64}" stroke="#eceef2"/>`);
  // User bubble, right-aligned
  const bw = 390, bx = split - 32 - bw, by = win.y + 92;
  out.push(`<rect x="${bx}" y="${by}" width="${bw}" height="74" rx="14" fill="#edf1ff"/>`);
  t.ask.forEach((line, i) => out.push(text(bx + 18, by + 30 + i * 24, line, 'font-size="15"')));
  // Agent step
  const sy = by + 112;
  out.push(`<circle cx="${cx + 8}" cy="${sy - 5}" r="8" fill="#e8f6ee"/><path d="M${cx + 4} ${sy - 5} l3 3 l5 -6" stroke="#1f9d55" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`);
  out.push(text(cx + 24, sy, t.step, 'font-size="13.5" fill="#6b7280"'));
  out.push(text(cx + 24 + (lang === 'en' ? 150 : 152), sy, 'doc_viewer_edit', 'font-size="12.5" fill="#9aa1ad" font-family="ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"'));
  // Answer
  t.answer.forEach((line, i) => out.push(text(cx, sy + 44 + i * 26, line, 'font-size="15.5"')));
  // Totals card
  const ty = sy + 104, tw = split - cx - 32;
  out.push(`<rect x="${cx}" y="${ty}" width="${tw}" height="112" rx="12" fill="#fff" stroke="#e6e8ec"/>`);
  const col = [cx + 22, cx + 22 + tw / 3, cx + 22 + (2 * tw) / 3];
  [[t.before, '707,036'], [t.after, '735,853'], [t.change, '+28,817']].forEach(([label, value], i) => {
    out.push(text(col[i], ty + 40, label, 'font-size="13" fill="#8a919e"'));
    out.push(text(col[i], ty + 78, value, `font-size="24" font-weight="600" ${i === 2 ? `fill="${ACCENT}"` : ''} style="font-variant-numeric:tabular-nums"`));
  });
  // Input box
  const iy = win.y + win.h - 92;
  out.push(`<rect x="${cx}" y="${iy}" width="${tw}" height="60" rx="14" fill="#fff" stroke="#e1e4ea"/>`);
  out.push(text(cx + 18, iy + 36, t.input, 'font-size="14.5" fill="#a3a9b4"'));
  out.push(`<circle cx="${cx + tw - 30}" cy="${iy + 30}" r="15" fill="${ACCENT}" opacity="0.35"/><path d="M${cx + tw - 30} ${iy + 37} v-13 m-5 5 l5 -5 l5 5" stroke="#fff" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`);

  // --- Document pane ---
  const dx = split, dw = win.x + win.w - split;
  // Tab bar
  out.push(`<rect x="${dx + 18}" y="${win.y + 20}" width="${lang === 'en' ? 150 : 140}" height="32" rx="8" fill="#f1f3f6"/>`);
  out.push(`<rect x="${dx + 31}" y="${win.y + 30}" width="12" height="12" rx="2.5" fill="none" stroke="#6b7280" stroke-width="1.4"/><line x1="${dx + 31}" y1="${win.y + 36}" x2="${dx + 43}" y2="${win.y + 36}" stroke="#6b7280" stroke-width="1.2"/><line x1="${dx + 37}" y1="${win.y + 30}" x2="${dx + 37}" y2="${win.y + 42}" stroke="#6b7280" stroke-width="1.2"/>`);
  out.push(text(dx + 52, win.y + 41, t.file, 'font-size="14" font-weight="500"'));
  // Toolbar right: changes chip + save
  const rx = win.x + win.w - 20;
  out.push(`<rect x="${rx - 64}" y="${win.y + 20}" width="64" height="32" rx="8" fill="#f1f3f6"/>`);
  out.push(text(rx - 32, win.y + 41, t.save, 'font-size="14" text-anchor="middle" font-weight="500"'));
  out.push(`<path d="M${rx - 112} ${win.y + 29} l2.2 5.6 5.6 2.2 -5.6 2.2 -2.2 5.6 -2.2 -5.6 -5.6 -2.2 5.6 -2.2z" fill="${ACCENT}"/>`);
  out.push(text(rx - 98, win.y + 42, t.chip, `font-size="15" font-weight="600" fill="${ACCENT}"`));
  out.push(`<line x1="${dx}" y1="${win.y + 64}" x2="${win.x + win.w}" y2="${win.y + 64}" stroke="#eceef2"/>`);

  // Grid
  const gx = dx + 0, gy = win.y + 64;
  const colW = [44, 92, lang === 'en' ? 150 : 132, 70, 106, 128];
  const rowH = 38;
  const xs = [gx];
  colW.forEach((w) => xs.push(xs[xs.length - 1] + w));
  const gridRight = win.x + win.w;
  const letters = ['', 'A', 'B', 'C', 'D', 'E'];
  // Column letters row
  out.push(`<rect x="${gx}" y="${gy}" width="${gridRight - gx}" height="30" fill="#f7f8fa"/>`);
  letters.forEach((l, i) => l && out.push(text(xs[i] + colW[i] / 2, gy + 20, l, 'font-size="12.5" fill="#8a919e" text-anchor="middle"')));
  const top = gy + 30;
  const rows = [t.head, ...ROWS.map(([r, p, q, old, now]) => [t.regions[r], t.products[p], String(q), now ?? old, (now ?? old) * q])];
  rows.forEach((cells, ri) => {
    const y = top + ri * rowH;
    if (ri === 0) out.push(`<rect x="${xs[1]}" y="${y}" width="${gridRight - xs[1]}" height="${rowH}" fill="#f4f6fb"/>`);
    out.push(`<rect x="${gx}" y="${y}" width="${colW[0]}" height="${rowH}" fill="#f7f8fa"/>`);
    out.push(text(gx + colW[0] / 2, y + 24, String(ri + 1), 'font-size="12.5" fill="#8a919e" text-anchor="middle"'));
    out.push(`<line x1="${gx}" y1="${y + rowH}" x2="${gridRight}" y2="${y + rowH}" stroke="#eef0f3"/>`);
    cells.forEach((value, ci) => {
      const x = xs[ci + 1];
      const numeric = ri > 0 && ci >= 2;
      const shown = ri === 0 ? value : ci === 3 ? num(value) : ci === 4 ? num(value, 2) : value;
      const attrs = `font-size="14.5" ${ri === 0 ? 'font-weight="600"' : ''} ${numeric ? 'style="font-variant-numeric:tabular-nums"' : ''}`;
      out.push(text(x + 12, y + 25, shown, attrs));
    });
  });
  xs.slice(1, -1).forEach((x) => out.push(`<line x1="${x}" y1="${gy}" x2="${x}" y2="${top + rows.length * rowH}" stroke="#eef0f3"/>`));
  out.push(`<line x1="${xs[1]}" y1="${gy}" x2="${xs[1]}" y2="${win.y + win.h}" stroke="#e6e8ec"/>`);
  // Agent marks on the changed prices, as the plugin draws them
  let first = true;
  ROWS.forEach(([, , , , now], i) => {
    if (now === undefined) return;
    const y = top + (i + 1) * rowH;
    const x = xs[4];
    out.push(`<rect x="${x + 1.5}" y="${y + 1.5}" width="${colW[4] - 3}" height="${rowH - 3}" rx="2" fill="${ACCENT}" fill-opacity="0.12" stroke="${ACCENT}" stroke-width="2"/>`);
    if (first) {
      first = false;
      // The label hangs from the cell's top-right corner, clear of the header text.
      const lx = x + colW[4] - 1.5 - 52;
      out.push(`<path d="M${lx} ${y - 16} a4 4 0 0 1 4 -4 h44 a4 4 0 0 1 4 4 v18 h-48 a4 4 0 0 1 -4 -4 z" fill="${ACCENT}"/>`);
      out.push(text(lx + 26, y - 1, 'Agent', 'font-size="11.5" font-weight="600" fill="#fff" text-anchor="middle"'));
    }
  });
  // Toast
  const tw2 = lang === 'en' ? 252 : 236, tx = dx + (dw - tw2) / 2, tyy = win.y + win.h - 62;
  out.push(`<rect x="${tx}" y="${tyy}" width="${tw2}" height="40" rx="20" fill="#fff" stroke="#e4e7ec" filter="url(#shadow)"/>`);
  out.push(text(tx + 20, tyy + 25, t.toast, 'font-size="14"'));
  out.push(text(tx + tw2 - 20, tyy + 25, t.view, `font-size="14" font-weight="600" fill="${ACCENT}" text-anchor="end"`));
  out.push('</g></svg>');
  return out.join('\n') + '\n';
}

writeFileSync(join(ROOT, 'docs', 'hero.svg'), hero('en'));
writeFileSync(join(ROOT, 'docs', 'hero.zh-CN.svg'), hero('zh'));
console.log('wrote docs/hero.svg and docs/hero.zh-CN.svg');
