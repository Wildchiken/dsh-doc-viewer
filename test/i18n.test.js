import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import EN from '../src/client/locale-en.js';

const DIR = new URL('../src/client/', import.meta.url);
const KEY = /(?<![\w$.])t\('((?:[^'\\]|\\.)*)'/g;
/** The phrase as JavaScript reads the literal: escapes such as \n become characters. */
const unescape = (literal) => JSON.parse(`"${literal.replace(/\\'/g, "'").replace(/"/g, '\\"')}"`);

const used = new Map();
for (const file of readdirSync(DIR).filter((name) => /\.(js|jsx)$/.test(name) && !['i18n.js', 'locale-en.js'].includes(name))) {
  for (const [, literal] of readFileSync(new URL(file, DIR), 'utf8').matchAll(KEY)) used.set(unescape(literal), file);
}

test('every phrase in the interface has an English translation', () => {
  const missing = [...used].filter(([phrase]) => !(phrase in EN)).map(([phrase, file]) => `${file}: ${phrase}`);
  assert.deepEqual(missing, []);
});

test('translations keep their placeholders', () => {
  const placeholders = (text) => [...text.matchAll(/\{\d+\}/g)].map((m) => m[0]).sort().join(',');
  const broken = Object.entries(EN).filter(([zh, en]) => placeholders(zh) !== placeholders(en)).map(([zh]) => zh);
  assert.deepEqual(broken, []);
});

test('the dictionary has no phrases the interface no longer uses', () => {
  const unused = Object.keys(EN).filter((phrase) => !used.has(phrase));
  assert.deepEqual(unused, []);
});
