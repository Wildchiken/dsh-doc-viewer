/**
 * Build the browser artifacts in the format the Harness module loader expects.
 *
 * `client.js` is merged into the bundle every window loads at startup, so it holds
 * only the sidebar entry and the wiring. The spreadsheet grid and PDF.js live in
 * `client.viewer.js`, which the loader fetches the first time the page opens.
 * React and ReactDOM come from the loader and are never duplicated.
 */
import { build } from 'esbuild';
import { readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const PLUGIN_ID = 'dsh-doc-viewer';

/** FortuneSheet paints unstyled cells in Times New Roman; use the system's UI font instead. */
const systemFont = {
  name: 'system-font',
  setup(context) {
    context.onLoad({ filter: /@fortune-sheet[\\/]core[\\/]dist[\\/]index\.esm\.js$/ }, async ({ path }) => {
      const source = await readFile(path, 'utf8');
      return { contents: source.replaceAll('fontarray: ["Times New Roman"', 'fontarray: ["PingFang SC"'), loader: 'js' };
    });
  },
};

/**
 * @param {string} entry
 * @param {string} output
 * @param {string|undefined} chunk
 */
async function bundle(entry, output, chunk) {
  const temporary = resolve(ROOT, 'build', `${output}.tmp.cjs`);
  await build({
    entryPoints: [resolve(ROOT, entry)],
    outfile: temporary,
    bundle: true,
    format: 'cjs',
    platform: 'browser',
    target: ['chrome120', 'safari17'],
    jsx: 'transform',
    external: ['react', 'react-dom'],
    loader: { '.css': 'text' },
    plugins: [systemFont],
    define: { 'process.env.NODE_ENV': '"production"' },
    legalComments: 'eof',
    minify: true,
    logLevel: 'warning',
  });
  const code = await readFile(temporary, 'utf8');
  const wrapped = `window.__ModuleLoader__.load({
  id: ${JSON.stringify(PLUGIN_ID)},${chunk ? `\n  chunk: ${JSON.stringify(chunk)},` : ''}
  factory(require) {
    const module = { exports: {} };
    const exports = module.exports;
    (function (module, exports, require) {
${code}
    })(module, exports, require);
    return module.exports;
  }
});
`;
  await writeFile(resolve(ROOT, output), wrapped);
  await rm(temporary, { force: true });
  console.log(`wrote ${output} (${(Buffer.byteLength(wrapped) / 1024).toFixed(0)} KiB)`);
}

await bundle('src/client/main.jsx', 'client.js');
await bundle('src/client/viewer.jsx', 'client.viewer.js', 'client.viewer.js');
