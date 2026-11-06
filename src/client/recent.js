/** The list of recently seen files, kept per browser profile. */
const KEY = 'dsh-doc-viewer:recent';
const MAX_ITEMS = 12;

/** @returns {{ path: string, name: string, kind?: string, at?: number }[]} */
export function loadRecent() {
  try {
    const parsed = JSON.parse(localStorage.getItem(KEY) || '[]');
    return Array.isArray(parsed) ? parsed.filter((item) => item && typeof item.path === 'string') : [];
  } catch {
    return [];
  }
}

/** @param {{ path: string, name?: string, kind?: string }} item */
export function rememberRecent(item) {
  const entry = { name: item.path.split('/').pop(), ...item, at: Date.now() };
  const next = [entry, ...loadRecent().filter((other) => other.path !== item.path)].slice(0, MAX_ITEMS);
  try {
    localStorage.setItem(KEY, JSON.stringify(next));
  } catch { /* localStorage may be disabled */ }
  return next;
}

/** @param {string} path */
export function forgetRecent(path) {
  const next = loadRecent().filter((item) => item.path !== path);
  try {
    localStorage.setItem(KEY, JSON.stringify(next));
  } catch { /* localStorage may be disabled */ }
  return next;
}
