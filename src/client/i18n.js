/**
 * Interface language. The text is written in Chinese and looked up in an English
 * dictionary when Harness shows any other language; a phrase missing there stays Chinese
 * rather than showing a key. Harness keeps the current language on <html lang>.
 */
import React from 'react';
import EN from './locale-en.js';

/** 'zh' or 'en': the language the interface is shown in. */
export function lang() {
  return (document.documentElement.lang || navigator.language || 'zh').toLowerCase().startsWith('zh') ? 'zh' : 'en';
}

/**
 * The phrase in the current language, with {0}, {1}… replaced by `values`.
 * @param {string} zh the phrase as written, in Chinese
 * @param {...any} values
 */
export function t(zh, ...values) {
  const phrase = lang() === 'zh' ? zh : (EN[zh] ?? zh);
  return values.length ? phrase.replace(/\{(\d+)\}/g, (match, index) => (values[index] === undefined ? match : String(values[index]))) : phrase;
}

/** Re-render when Harness switches language. */
export function useLang() {
  const [value, setValue] = React.useState(lang);
  React.useEffect(() => {
    const observer = new MutationObserver(() => setValue(lang()));
    observer.observe(document.documentElement, { attributes: true, attributeFilter: ['lang'] });
    return () => observer.disconnect();
  }, []);
  return value;
}
