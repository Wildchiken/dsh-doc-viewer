/**
 * The system's own open and save panels, so files are opened where they are and
 * saved where the user says. macOS uses NSOpenPanel/NSSavePanel through the JXA
 * host, the way the Harness's own folder picker does; Linux uses zenity.
 */
import { execFile } from 'node:child_process';
import { t } from './i18n.js';
import { httpError } from './util.js';

/** An Edit menu makes ⌘C/⌘V work in the panel's fields (Go to Folder, the name field). */
const MAC_PRELUDE = `
ObjC.import('AppKit');
function prepare() {
  var app = $.NSApplication.sharedApplication;
  app.setActivationPolicy($.NSApplicationActivationPolicyAccessory);
  var mainMenu = $.NSMenu.alloc.init;
  mainMenu.addItem($.NSMenuItem.alloc.init);
  var editMenu = $.NSMenu.alloc.initWithTitle('Edit');
  var editItem = $.NSMenuItem.alloc.init;
  editItem.submenu = editMenu;
  mainMenu.addItem(editItem);
  editMenu.addItemWithTitleActionKeyEquivalent('Undo', 'undo:', 'z');
  editMenu.addItemWithTitleActionKeyEquivalent('Cut', 'cut:', 'x');
  editMenu.addItemWithTitleActionKeyEquivalent('Copy', 'copy:', 'c');
  editMenu.addItemWithTitleActionKeyEquivalent('Paste', 'paste:', 'v');
  editMenu.addItemWithTitleActionKeyEquivalent('Select All', 'selectAll:', 'a');
  app.mainMenu = mainMenu;
  app.activateIgnoringOtherApps(true);
}`;

/** argv: message, directory. Prints one chosen path per line; nothing when cancelled. */
const MAC_OPEN = `${MAC_PRELUDE}
function run(argv) {
  prepare();
  var panel = $.NSOpenPanel.openPanel;
  panel.canChooseFiles = true;
  panel.canChooseDirectories = false;
  panel.allowsMultipleSelection = true;
  panel.message = argv[0];
  if (argv[1]) panel.directoryURL = $.NSURL.fileURLWithPath(argv[1]);
  if (panel.runModal !== $.NSModalResponseOK) return '';
  var urls = panel.URLs;
  var paths = [];
  for (var i = 0; i < urls.count; i++) paths.push(ObjC.unwrap(urls.objectAtIndex(i).path));
  return paths.join('\\n');
}`;

/** argv: message, directory, file name. The panel itself asks before replacing a file. */
const MAC_SAVE = `${MAC_PRELUDE}
function run(argv) {
  prepare();
  var panel = $.NSSavePanel.savePanel;
  panel.canCreateDirectories = true;
  panel.message = argv[0];
  if (argv[1]) panel.directoryURL = $.NSURL.fileURLWithPath(argv[1]);
  panel.nameFieldStringValue = argv[2];
  if (panel.runModal !== $.NSModalResponseOK) return '';
  return ObjC.unwrap(panel.URL.path);
}`;

/** @param {string} command @param {string[]} args */
function run(command, args) {
  return new Promise((resolve, reject) => {
    execFile(command, args, { timeout: 30 * 60 * 1000, maxBuffer: 1024 * 1024 }, (error, stdout) => {
      if (error) reject(error);
      else resolve(String(stdout).replace(/[\r\n]+$/, ''));
    });
  });
}

function unavailable() {
  return httpError(501, 'native-unavailable', t('这个系统上没有可用的文件面板。'));
}

/**
 * @param {{ message?: string, directory?: string }} options
 * @returns {Promise<string[]>} the chosen files; empty when cancelled
 */
export async function pickFiles(options = {}) {
  const message = options.message || t('选择要打开的文档');
  if (process.platform === 'darwin') {
    const out = await run('osascript', ['-l', 'JavaScript', '-e', MAC_OPEN, message, options.directory || '']);
    return out ? out.split('\n').filter(Boolean) : [];
  }
  if (process.platform === 'linux') {
    try {
      const out = await run('zenity', ['--file-selection', '--multiple', '--separator=\n', `--title=${message}`, ...(options.directory ? [`--filename=${options.directory}/`] : [])]);
      return out ? out.split('\n').filter(Boolean) : [];
    } catch (error) {
      if (error?.code === 1) return [];
      throw unavailable();
    }
  }
  throw unavailable();
}

/**
 * @param {{ message?: string, directory?: string, name: string }} options
 * @returns {Promise<string|null>} the chosen path; null when cancelled
 */
export async function pickSaveLocation(options) {
  const message = options.message || t('存储到');
  if (process.platform === 'darwin') {
    return (await run('osascript', ['-l', 'JavaScript', '-e', MAC_SAVE, message, options.directory || '', options.name])) || null;
  }
  if (process.platform === 'linux') {
    try {
      return (await run('zenity', ['--file-selection', '--save', '--confirm-overwrite', `--title=${message}`, `--filename=${options.directory ? `${options.directory}/` : ''}${options.name}`])) || null;
    } catch (error) {
      if (error?.code === 1) return null;
      throw unavailable();
    }
  }
  throw unavailable();
}
