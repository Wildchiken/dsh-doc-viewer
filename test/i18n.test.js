import { describe, test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import CLIENT_EN from '../src/client/locale-en.js';
import HOST_EN from '../src/host/locale-en.js';
import { lang, t, withLang } from '../src/host/i18n.js';

const KEY = /(?<![\w$.])t\('((?:[^'\\]|\\.)*)'/g;
/** The phrase as JavaScript reads the literal: escapes such as \n become characters. */
const unescape = (literal) => JSON.parse(`"${literal.replace(/\\'/g, "'").replace(/"/g, '\\"')}"`);
const placeholders = (text) => [...text.matchAll(/\{(\d+)(?::[^}]*)?\}/g)].map((m) => m[1]).sort().join(',');

/** @param {URL[]} files */
function phrasesIn(files) {
  const used = new Map();
  for (const file of files) {
    for (const [, literal] of readFileSync(file, 'utf8').matchAll(KEY)) used.set(unescape(literal), file.pathname.split('/').pop());
  }
  return used;
}

const sources = (dir) => readdirSync(dir).filter((name) => /\.(js|jsx)$/.test(name) && !['i18n.js', 'locale-en.js'].includes(name)).map((name) => new URL(name, dir));

for (const [side, dictionary, files] of [
  ['interface', CLIENT_EN, sources(new URL('../src/client/', import.meta.url))],
  ['Host', HOST_EN, [...sources(new URL('../src/host/', import.meta.url)), new URL('../index.js', import.meta.url)]],
]) {
  describe(`${side} translations`, () => {
    const used = phrasesIn(files);

    test('every phrase has an English translation', () => {
      assert.deepEqual([...used].filter(([phrase]) => !(phrase in dictionary)).map(([phrase, file]) => `${file}: ${phrase}`), []);
    });

    test('translations keep their placeholders', () => {
      assert.deepEqual(Object.entries(dictionary).filter(([zh, en]) => placeholders(zh) !== placeholders(en)).map(([zh]) => zh), []);
    });

    test('the dictionary holds nothing the code no longer uses', () => {
      assert.deepEqual(Object.keys(dictionary).filter((phrase) => !used.has(phrase)), []);
    });
  });
}

describe('translation helpers are imported where they are used', () => {
  for (const file of [...sources(new URL('../src/client/', import.meta.url)), ...sources(new URL('../src/host/', import.meta.url)), new URL('../index.js', import.meta.url)]) {
    const text = readFileSync(file, 'utf8');
    const name = file.pathname.split('/').pop();
    for (const helper of ['t', 'lang', 'useLang', 'withLang']) {
      if (!new RegExp(`(?<![\\w$.])${helper}\\(`).test(text)) continue;
      test(`${name} imports ${helper}`, () => {
        assert.match(text, new RegExp(`import \\{[^}]*\\b${helper}\\b[^}]*\\} from '\\./(src/host/)?i18n\\.js'`));
      });
    }
  }
});

describe('Host message language', () => {
  test('follows the page for a request, and stays with it for the agent', () => {
    assert.equal(withLang('en-GB', () => t('文件不存在：{0}', 'a.xlsx')), 'File not found: a.xlsx');
    assert.equal(withLang('zh-CN', () => t('文件不存在：{0}', 'a.xlsx')), '文件不存在：a.xlsx');
    withLang('en', () => {});
    assert.equal(lang(), 'en', 'a tool call outside any request uses the language last seen');
    withLang('zh', () => {});
    assert.equal(lang(), 'zh');
  });

  test('a count picks the word that agrees with it', () => {
    withLang('en', () => {
      assert.equal(t('已新建 {0}，{1} 行。', 'a.xlsx', 1), 'Created a.xlsx with 1 row.');
      assert.equal(t('已新建 {0}，{1} 行。', 'a.xlsx', 3), 'Created a.xlsx with 3 rows.');
    });
    assert.equal(withLang('zh', () => t('已新建 {0}，{1} 行。', 'a.xlsx', 1)), '已新建 a.xlsx，1 行。');
  });

  test('a request without a language changes nothing', () => {
    withLang('zh', () => {});
    assert.equal(withLang(undefined, () => lang()), 'zh');
  });
});
