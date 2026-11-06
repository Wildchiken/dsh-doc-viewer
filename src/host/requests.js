/**
 * Pending "show this in the panel" requests raised by agent tools.
 *
 * `reveal` requests (doc_viewer_open) bring the panel up. The others only announce
 * a new output file, so an agent editing a workbook in a loop does not keep
 * covering the screen. Requests expire: a panel connecting tomorrow must not
 * replay what an agent did today.
 */
const TTL_MS = 2 * 60 * 1000;
const MAX_PENDING = 20;
const pending = [];
let sequence = 0;

function dropExpired(now = Date.now()) {
  while (pending.length > 0 && now - pending[0].at > TTL_MS) pending.shift();
}

/**
 * @param {string} path
 * @param {{ sessionId?: string, reveal?: boolean, replaces?: string, inPlace?: boolean, highlight?: { sheet: string, address: string }[], note?: string }} [options]
 *   `replaces` names the document this file is a new version of: an open tab for
 *   it shows the new file in its place, with `highlight` selected.
 */
export function enqueueOpen(path, options = {}) {
  dropExpired();
  sequence += 1;
  pending.push({
    id: `open-${Date.now()}-${sequence}`,
    path,
    sessionId: options.sessionId,
    reveal: options.reveal === true,
    replaces: options.replaces,
    inPlace: options.inPlace === true,
    highlight: options.highlight,
    note: options.note,
    kind: options.kind,
    live: options.live === true,
    cells: options.cells,
    edited: options.edited,
    mtime: options.mtime,
    at: Date.now(),
  });
  if (pending.length > MAX_PENDING) pending.splice(0, pending.length - MAX_PENDING);
}

/** When a page last asked for requests. */
let lastPoll = 0;

/** Whether a document page is open and picking up requests right now. */
export function pageListening(withinMs = 10_000) {
  return Date.now() - lastPoll < withinMs;
}

/** @returns {readonly any[]} */
export function listOpenRequests() {
  lastPoll = Date.now();
  dropExpired();
  return pending.slice();
}

/**
 * Remove requests and report which ones this caller removed, so that of several
 * open windows only the first one acts on a request.
 * @param {string[]} ids
 * @returns {string[]} the ids that were still pending
 */
export function ackOpenRequests(ids) {
  const remove = new Set(ids || []);
  const claimed = [];
  for (let index = pending.length - 1; index >= 0; index -= 1) {
    if (remove.has(pending[index].id)) {
      claimed.push(pending[index].id);
      pending.splice(index, 1);
    }
  }
  return claimed;
}

/**
 * What the user has open in the document page, as the page last reported it.
 * @type {{ at: number, active?: any, tabs: any[] } | undefined}
 */
let panelState;

/** @param {any} state */
export function setPanelState(state) {
  const tabs = Array.isArray(state?.tabs) ? state.tabs.slice(0, 50) : [];
  panelState = { at: Date.now(), visible: state?.visible === true, active: state?.active, tabs };
}

export function getPanelState() {
  return panelState;
}
