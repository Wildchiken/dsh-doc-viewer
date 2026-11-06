/** Same-origin API used by the viewer page. */
import { lang, t } from './i18n.js';

const PREFIX = '/dsh-doc-viewer/api';

/** @param {Response} response */
async function parse(response) {
  let payload;
  try {
    payload = await response.json();
  } catch {
    throw new Error(t('请求失败（HTTP {0}）', response.status));
  }
  if (!response.ok || payload?.ok === false) {
    const error = new Error(payload?.message || t('请求失败（HTTP {0}）', response.status));
    error.code = payload?.code;
    error.status = response.status;
    throw error;
  }
  return payload;
}

/** @param {string} path @param {unknown} body */
async function post(path, body) {
  const response = await fetch(`${PREFIX}${path}`, {
    method: 'POST',
    credentials: 'same-origin',
    headers: { 'content-type': 'application/json', 'x-dvd-lang': lang() },
    body: JSON.stringify(body),
  });
  return parse(response);
}

/** @param {string} query */
async function get(query) {
  return parse(await fetch(`${PREFIX}${query}`, { credentials: 'same-origin', headers: { 'x-dvd-lang': lang() } }));
}

/** @param {string} path */
export function fileUrl(path) {
  return `${PREFIX}/file?path=${encodeURIComponent(path)}`;
}

/** Upload one browser File into the workspace upload folder. @param {File} file */
export async function upload(file) {
  const workspace = rootOf();
  const response = await fetch(`${PREFIX}/import?name=${encodeURIComponent(file.name)}${workspace ? `&root=${encodeURIComponent(workspace)}` : ''}`, {
    method: 'POST',
    credentials: 'same-origin',
    headers: { 'content-type': 'application/octet-stream', 'x-dvd-lang': lang() },
    body: file,
  });
  return parse(response);
}

/**
 * @param {string} path
 * @param {{ recalc?: boolean|'later' }} [options] 'later' returns at once, without waiting for
 *   LibreOffice to fill in formula results the file does not hold.
 */
export function open(path, options = {}) {
  return post('/open', { path, workspaceRoot: rootOf(), recalc: options.recalc });
}

/** @param {string} [path] */
export function list(path = '.') {
  const workspace = rootOf();
  return get(`/list?path=${encodeURIComponent(path)}${workspace ? `&root=${encodeURIComponent(workspace)}` : ''}`);
}

export const save = (payload) => post('/save', { workspaceRoot: rootOf(), ...payload });
export const convert = (payload) => post('/convert', { workspaceRoot: rootOf(), ...payload });
export const pdf = (payload) => post('/pdf', { workspaceRoot: rootOf(), ...payload });
export const requests = () => get('/requests');
export const ack = (ids) => post('/requests/ack', { ids });

/* Workspace-aware calls --------------------------------------------------- */

let rootOf = () => undefined;

/** @param {{ root: () => string|undefined }} options */
export function configure(options) {
  rootOf = options.root;
}

/** The workspace of the conversation the user is in, if known. */
export function root() {
  return rootOf();
}

/** @param {string} path */
export const reveal = (path) => post('/reveal', { path });

/** Tell the Host what is open, for the agent's doc_viewer_context. @param {any} state */
export function reportState(state) {
  return post('/state', state).catch(() => {});
}

/** The system open panel; the files chosen, opened where they are. */
export const pickOpen = () => post('/pick-open', { directory: rootOf() });
/** The system save panel. @param {{ name: string, directory?: string }} options */
export const pickSave = (options) => post('/pick-save', { ...options, workspaceRoot: rootOf() });
/** @param {string} path */
/** The results of the formulas in what the grid holds now, worked out on the Host. */
export const calculate = (payload) => post('/calculate', payload);
/** When each document last changed on disk; null for one that is gone. */
export const stat = (paths) => post('/stat', { paths });
/** Make an empty document at a place picked in the save panel. */
export const create = (path) => post('/create', { path });
export const versions = (path) => get(`/versions?path=${encodeURIComponent(path)}`);
/** @param {string} path @param {string} [id] newest when omitted */
export const restoreVersion = (path, id) => post('/versions/restore', { path, id });

/* Unsaved edits ----------------------------------------------------------- */

/** @param {{ path: string, text: string, openedMtime?: number }} request */
export const saveText = (request) => post('/save-text', { ...request, workspaceRoot: rootOf() });

export const recoveryList = () => get('/recovery');
/** @param {string} path */
export const recoveryItem = (path) => get(`/recovery/item?path=${encodeURIComponent(path)}`);
export const saveRecovery = (snapshot) => post('/recovery', snapshot);
/** @param {string} path */
export const discardRecovery = (path) => post('/recovery/discard', { path }).catch(() => {});
