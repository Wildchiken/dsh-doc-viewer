/** Unsaved edits kept across a quit. */
import { after, before, test } from 'node:test';
import { removeTempDirs, tempDir } from './fixtures.js';
import assert from 'node:assert/strict';
import { mkdtempSync, readdirSync, utimesSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { discardRecovery, listRecovery, readRecovery, saveRecovery } from '../src/host/recovery.js';

after(removeTempDirs);

before(() => {
  process.env.DSH_HOME = tempDir('doc-viewer-home-');
});

const sheets = [{ name: '表', id: 'sheet-0', celldata: [{ r: 0, c: 0, v: { v: '改过', m: '改过' } }] }];

test('a snapshot is kept, listed without its content, and read back whole', () => {
  saveRecovery({ path: '/工作/报价.xlsx', name: '报价.xlsx', sheets, structure: [{ op: 'insertRowCol', id: 'sheet-0', type: 'row', index: 1, count: 1 }], openedMtime: 123 });
  const [entry] = listRecovery();
  assert.equal(entry.path, '/工作/报价.xlsx');
  assert.equal(entry.sheets, undefined);
  assert.equal(entry.sourceExists, false);
  const record = readRecovery('/工作/报价.xlsx');
  assert.deepEqual(record.sheets, sheets);
  assert.equal(record.openedMtime, 123);
  assert.equal(record.structure.length, 1);
});

test('a newer snapshot replaces the older one', () => {
  saveRecovery({ path: '/工作/a.xlsx', sheets });
  saveRecovery({ path: '/工作/a.xlsx', sheets: [{ ...sheets[0], name: '新' }] });
  assert.equal(readRecovery('/工作/a.xlsx').sheets[0].name, '新');
  assert.equal(listRecovery().filter((entry) => entry.path === '/工作/a.xlsx').length, 1);
});

test('discarding removes it', () => {
  saveRecovery({ path: '/工作/b.xlsx', sheets });
  discardRecovery('/工作/b.xlsx');
  assert.equal(readRecovery('/工作/b.xlsx'), undefined);
});

test('snapshots older than a month are dropped', () => {
  saveRecovery({ path: '/工作/旧.xlsx', sheets });
  const folder = join(process.env.DSH_HOME, 'doc-viewer', 'recovery');
  const old = new Date(Date.now() - 40 * 86_400_000);
  for (const name of readdirSync(folder)) utimesSync(join(folder, name), old, old);
  assert.deepEqual(listRecovery(), []);
});
