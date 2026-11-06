/**
 * Language of the messages the Host sends: errors and notes the page shows, and the
 * results the agent's tools report. The page says which language it shows with every
 * request; the agent's tools follow the language the page last used. Chinese is the
 * default and the text as written; English comes from locale-en.js.
 */
import { AsyncLocalStorage } from 'node:async_hooks';
import EN from './locale-en.js';

const request = new AsyncLocalStorage();
let lastUsed = 'zh';

/** @param {string|undefined} value a language tag such as zh-CN or en */
export function normalizeLang(value) {
  if (typeof value !== 'string' || value === '') return undefined;
  return value.toLowerCase().startsWith('zh') ? 'zh' : 'en';
}

/**
 * Run `fn` with messages in `language`, which also becomes the default for tool results.
 * @template T @param {string|undefined} language @param {() => T} fn @returns {T}
 */
export function withLang(language, fn) {
  const lang = normalizeLang(language);
  if (!lang) return fn();
  lastUsed = lang;
  return request.run(lang, fn);
}

/** 'zh' or 'en' */
export function lang() {
  return request.getStore() ?? lastUsed;
}

/**
 * The message in the current language, with {0}, {1}… replaced by `values`.
 * @param {string} zh @param {...any} values
 */
export function t(zh, ...values) {
  const phrase = lang() === 'zh' ? zh : (EN[zh] ?? zh);
  return values.length ? phrase.replace(/\{(\d+)\}/g, (match, index) => (values[index] === undefined ? match : String(values[index]))) : phrase;
}
