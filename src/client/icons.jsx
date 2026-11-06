import React from 'react';

/** Thin line icons on a 24 grid, drawn to sit alongside the Host's own glyphs. */
const paths = {
  plus: ['M12 5.5v13', 'M5.5 12h13'],
  close: ['M7 7l10 10', 'M17 7L7 17'],
  chevronLeft: ['M14.5 6l-6 6 6 6'],
  chevronRight: ['M9.5 6l6 6-6 6'],
  chevronUp: ['M6 14.5l6-6 6 6'],
  chevronDown: ['M6 9.5l6 6 6-6'],
  more: ['M5.5 12h.01', 'M12 12h.01', 'M18.5 12h.01'],
  save: ['M12 4v11', 'M7.5 10.5L12 15l4.5-4.5', 'M5 19.5h14'],
  convert: ['M4.5 8h13', 'M14 4.5L17.5 8 14 11.5', 'M19.5 16h-13', 'M10 12.5L6.5 16l3.5 3.5'],
  stats: ['M6 19.5v-7', 'M12 19.5v-15', 'M18 19.5v-10'],
  search: ['M10.75 17.5a6.75 6.75 0 1 1 0-13.5 6.75 6.75 0 0 1 0 13.5Z', 'M15.75 15.75L20 20'],
  pdf: ['M6.5 2.75h7.25L19 8v12a1.25 1.25 0 0 1-1.25 1.25H6.5A1.25 1.25 0 0 1 5.25 20V4A1.25 1.25 0 0 1 6.5 2.75Z', 'M13.5 3v5.25H19', 'M9 16.5h6'],
  sheet: ['M5.5 4h13A1.5 1.5 0 0 1 20 5.5v13a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 4 18.5v-13A1.5 1.5 0 0 1 5.5 4Z', 'M4 9.5h16', 'M4 14.75h16', 'M10 9.5V20'],
  word: ['M6.5 2.75h7.25L19 8v12a1.25 1.25 0 0 1-1.25 1.25H6.5A1.25 1.25 0 0 1 5.25 20V4A1.25 1.25 0 0 1 6.5 2.75Z', 'M13.5 3v5.25H19', 'M8.75 12.5h6.5', 'M8.75 16h6.5'],
  text: ['M5 6.5h14', 'M5 11h14', 'M5 15.5h9'],
  freeze: ['M4.5 8.5h15', 'M6 4h12a1.5 1.5 0 0 1 1.5 1.5v13A1.5 1.5 0 0 1 18 20H6a1.5 1.5 0 0 1-1.5-1.5v-13A1.5 1.5 0 0 1 6 4Z'],
  number: ['M9.5 4L8 20', 'M16 4l-1.5 16', 'M4.5 9h15', 'M4 15h15'],
  copy: ['M9.5 8.5h9A1.5 1.5 0 0 1 20 10v9a1.5 1.5 0 0 1-1.5 1.5h-9A1.5 1.5 0 0 1 8 19v-9a1.5 1.5 0 0 1 1.5-1.5Z', 'M5 15.5H4.5A.5.5 0 0 1 4 15V5a1.5 1.5 0 0 1 1.5-1.5H15a.5.5 0 0 1 .5.5v.5'],
  thumbs: ['M4 5.5A1.5 1.5 0 0 1 5.5 4h13A1.5 1.5 0 0 1 20 5.5v13a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 4 18.5Z', 'M9.5 4v16'],
  eye: ['M2.75 12S6.25 5.75 12 5.75 21.25 12 21.25 12 17.75 18.25 12 18.25 2.75 12 2.75 12Z', 'M12 14.75a2.75 2.75 0 1 0 0-5.5 2.75 2.75 0 0 0 0 5.5Z'],
  warning: ['M12 4.5L21 19.5H3Z', 'M12 10v4.5', 'M12 17h.01'],
  folder: ['M3.75 7A1.5 1.5 0 0 1 5.25 5.5h4l2 2.25h7.5A1.5 1.5 0 0 1 20.25 9.25v8.25A1.5 1.5 0 0 1 18.75 19H5.25a1.5 1.5 0 0 1-1.5-1.5Z'],
  fit: ['M4.5 9V4.5H9', 'M19.5 9V4.5H15', 'M4.5 15v4.5H9', 'M19.5 15v4.5H15'],
  minus: ['M5.5 12h13'],
  sidebar: ['M4.5 5.5h15a1 1 0 0 1 1 1v11a1 1 0 0 1-1 1h-15a1 1 0 0 1-1-1v-11a1 1 0 0 1 1-1Z', 'M14.5 5.5v13'],
  expand: ['M14 4.5h5.5V10', 'M19.5 4.5L13 11', 'M10 19.5H4.5V14', 'M4.5 19.5L11 13'],
  check: ['M5.5 12.5l4.5 4.5 8.5-9.5'],
  merge: ['M7 4.5v5a4 4 0 0 0 4 4h6', 'M14 10l3.5 3.5L14 17'],
  undo: ['M9 14L4.5 9.5 9 5', 'M4.5 9.5h9.75a5.25 5.25 0 0 1 0 10.5H11'],
  sparkle: ['M12 3.5l1.9 5.1 5.1 1.9-5.1 1.9-1.9 5.1-1.9-5.1-5.1-1.9 5.1-1.9Z', 'M18.5 15.5l.8 2.2 2.2.8-2.2.8-.8 2.2-.8-2.2-2.2-.8 2.2-.8Z'],
};

export function Icon({ name, size = 16, className = '', strokeWidth = 1.6 }) {
  return (
    <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="currentColor" strokeWidth={strokeWidth} strokeLinecap="round" strokeLinejoin="round" className={`dvd-icon ${className}`} aria-hidden="true">
      {(paths[name] || paths.text).map((d, index) => <path key={index} d={d} />)}
    </svg>
  );
}

/** @param {string|undefined} kind */
export function iconForKind(kind) {
  return { spreadsheet: 'sheet', pdf: 'pdf', docx: 'word', text: 'text' }[kind] || 'word';
}

/** @param {string|undefined} name a file name or extension */
export function iconForName(name) {
  const ext = String(name || '').toLowerCase().match(/\.[a-z0-9]+$/)?.[0];
  if (ext === '.pdf') return 'pdf';
  if (['.xlsx', '.xls', '.csv', '.tsv', '.et'].includes(ext)) return 'sheet';
  if (['.docx', '.doc', '.wps'].includes(ext)) return 'word';
  return 'text';
}
