/**
 * Earlier versions of files saved in place. Before a file is overwritten, what was
 * there is copied into the Harness home, out of the user's folders, so a save or an
 * agent's edit can always be taken back.
 */
import { createHash } from 'node:crypto';
import { t } from './i18n.js';
import { chmodSync, copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, realpathSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { basename, dirname, join } from 'node:path';
import { httpError, token } from './util.js';

const KEEP_VERSIONS = 20;
const KEEP_DAYS = 30;

function storeRoot() {
  return join(process.env.DSH_HOME || join(homedir(), '.dsh'), 'doc-viewer', 'versions');
}

/** @param {string} path */
function folderFor(path) {
  return join(storeRoot(), createHash('sha256').update(path).digest('hex').slice(0, 20));
}

/**
 * Keep a copy of the file as it is now.
 * @param {string} path
 * @param {string} reason who replaced it: 保存, Agent 修改, 恢复
 * @param {{ beforeRestore?: boolean }} [extra]
 * @returns {string|undefined} the version id, or undefined when there was no file
 */
export function backupVersion(path, reason, extra = {}) {
  if (!existsSync(path)) return undefined;
  const folder = folderFor(path);
  mkdirSync(folder, { recursive: true });
  writeFileSync(join(folder, 'source.json'), JSON.stringify({ path }));
  // Strictly later than the last one: undo and pruning go by this order, and two saves can share a millisecond.
  const at = Math.max(Date.now(), (readVersions(folder)[0]?.at ?? 0) + 1);
  const id = `${at}-${token().slice(0, 6)}`;
  copyFileSync(path, join(folder, `${id}.bin`));
  writeFileSync(join(folder, `${id}.json`), JSON.stringify({ id, reason, at, size: statSync(path).size, name: basename(path), ...extra }));
  prune(folder);
  return id;
}

/**
 * Past the limit the oldest versions go, except the earliest one still in date: an
 * agent editing in a loop must not push the file it started from out of reach.
 * @param {string} folder
 */
function prune(folder) {
  const cutoff = Date.now() - KEEP_DAYS * 86_400_000;
  const versions = readVersions(folder);
  const earliest = versions.filter((version) => version.at >= cutoff).at(-1);
  versions.forEach((version, index) => {
    if (version.at < cutoff || (index >= KEEP_VERSIONS - 1 && version !== earliest)) {
      rmSync(join(folder, `${version.id}.bin`), { force: true });
      rmSync(join(folder, `${version.id}.json`), { force: true });
    }
  });
}

/** @param {string} folder newest first */
function readVersions(folder) {
  if (!existsSync(folder)) return [];
  return readdirSync(folder)
    .filter((name) => name.endsWith('.json') && name !== 'source.json')
    .map((name) => {
      try {
        return JSON.parse(readFileSync(join(folder, name), 'utf8'));
      } catch {
        return undefined;
      }
    })
    .filter((version) => version && existsSync(join(folder, `${version.id}.bin`)))
    .sort((a, b) => b.at - a.at);
}

/** @param {string} path newest first */
export function listVersions(path) {
  return readVersions(folderFor(path));
}

/**
 * What an undo can still go back to: not what it has used already, and not the
 * copies kept only so that a restore can itself be reversed.
 * @param {any} version
 */
function canStepBackTo(version) {
  return !version.undone && !version.beforeRestore && version.reason !== '恢复前';
}

/**
 * Put an earlier version back. What is there now is kept as a version too, so a
 * restore can itself be reversed from the list of versions.
 * @param {string} path
 * @param {string} [id] the version the user picked. Without one this is an undo:
 *   one step back from where the last undo stopped, so undoing twice goes back
 *   twice rather than returning to where it began.
 * @returns {any} the version restored, with `remaining` undo steps left after it
 */
export function restoreVersion(path, id) {
  const folder = folderFor(path);
  const versions = readVersions(folder);
  const version = id ? versions.find((item) => item.id === id) : versions.find(canStepBackTo);
  if (!version) throw httpError(404, 'no-version', t('没有找到可以恢复的版本。'));
  const content = readFileSync(join(folder, `${version.id}.bin`));
  if (!id) writeFileSync(join(folder, `${version.id}.json`), JSON.stringify({ ...version, undone: true }));
  backupVersion(path, '恢复', { beforeRestore: true });
  writeAtomically(path, content);
  return { ...version, remaining: readVersions(folder).filter(canStepBackTo).length };
}

/** The last piece of work queued on each file. @type {Map<string, Promise<unknown>>} */
const fileQueues = new Map();

/**
 * Run read-change-write work on a file one at a time. Agents issue tool calls in
 * parallel: without this each would start from the same content and the last one
 * to finish would erase the others.
 * @template T
 * @param {string} path
 * @param {() => Promise<T>|T} work
 * @returns {Promise<T>}
 */
export function withFileLock(path, work) {
  let key = path;
  try {
    key = realpathSync(path);
  } catch { /* a file not there yet */ }
  const previous = fileQueues.get(key) ?? Promise.resolve();
  const run = previous.then(work, work);
  const tail = run.catch(() => {}).finally(() => {
    if (fileQueues.get(key) === tail) fileQueues.delete(key);
  });
  fileQueues.set(key, tail);
  return run;
}

/**
 * Replace a file's content without a moment where it is half written. The file stays
 * the same file to everything else: through a symbolic link the target is replaced
 * and the link stays, and the permissions are kept.
 * @param {string} path
 * @param {Buffer|Uint8Array|string} content
 */
export function writeAtomically(path, content) {
  let target = path;
  let mode;
  try {
    target = realpathSync(path);
    mode = statSync(target).mode & 0o7777;
  } catch { /* a new file */ }
  const temporary = join(dirname(target), `.${basename(target)}.${token().slice(0, 8)}.tmp`);
  // Created with the final mode, so a private file is never readable for a moment.
  writeFileSync(temporary, content, mode === undefined ? undefined : { mode });
  try {
    // The umask may have narrowed it.
    if (mode !== undefined) chmodSync(temporary, mode);
    renameSync(temporary, target);
  } catch (error) {
    rmSync(temporary, { force: true });
    throw error;
  }
}
