/**
 * Unsaved edits, kept so they survive quitting or a crash. The page sends a
 * snapshot of each document with unsaved changes a moment after the user stops
 * editing; the next time the document page opens, they are put back.
 * Snapshots live in the Harness home, never in the user's folders.
 */
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { writeAtomically } from './versions.js';

const KEEP_DAYS = 30;

function storeRoot() {
  return join(process.env.DSH_HOME || join(homedir(), '.dsh'), 'doc-viewer', 'recovery');
}

/** @param {string} path */
function fileFor(path) {
  return join(storeRoot(), `${createHash('sha256').update(path).digest('hex').slice(0, 24)}.json`);
}

/**
 * @param {{ path: string, name?: string, sheets?: any[], text?: string, structure?: any[], openedMtime?: number, transient?: boolean }} snapshot
 */
export function saveRecovery(snapshot) {
  mkdirSync(storeRoot(), { recursive: true });
  const record = {
    path: snapshot.path,
    name: snapshot.name,
    sheets: snapshot.sheets,
    text: typeof snapshot.text === 'string' ? snapshot.text : undefined,
    structure: Array.isArray(snapshot.structure) ? snapshot.structure : [],
    openedMtime: snapshot.openedMtime,
    transient: snapshot.transient === true,
    at: Date.now(),
  };
  writeAtomically(fileFor(snapshot.path), JSON.stringify(record));
  return record.at;
}

/** @param {string} path @returns {any|undefined} the snapshot with its sheets */
export function readRecovery(path) {
  const file = fileFor(path);
  if (!existsSync(file)) return undefined;
  try {
    const record = JSON.parse(readFileSync(file, 'utf8'));
    return record?.path === path ? record : undefined;
  } catch {
    return undefined;
  }
}

/** @param {string} path */
export function discardRecovery(path) {
  rmSync(fileFor(path), { force: true });
}

/**
 * Documents with unsaved edits, newest first, without their content. Snapshots
 * older than a month are dropped.
 * @returns {{ path: string, name?: string, at: number, sourceExists: boolean }[]}
 */
export function listRecovery() {
  const root = storeRoot();
  if (!existsSync(root)) return [];
  const cutoff = Date.now() - KEEP_DAYS * 86_400_000;
  const entries = [];
  for (const name of readdirSync(root)) {
    if (!name.endsWith('.json')) continue;
    const file = join(root, name);
    try {
      if (statSync(file).mtimeMs < cutoff) {
        rmSync(file, { force: true });
        continue;
      }
      const record = JSON.parse(readFileSync(file, 'utf8'));
      if (typeof record?.path !== 'string') continue;
      entries.push({ path: record.path, name: record.name, at: record.at, transient: record.transient, sourceExists: existsSync(record.path) });
    } catch { /* a half-written or foreign file is skipped */ }
  }
  return entries.sort((a, b) => b.at - a.at);
}
