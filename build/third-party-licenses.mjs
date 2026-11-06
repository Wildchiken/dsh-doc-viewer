/**
 * Write THIRD_PARTY_LICENSES.md: the full license text of every package bundled into
 * client.js / client.viewer.js, and of the code shipped under vendor/. Packages the
 * plugin depends on at runtime are installed from npm with their own license files.
 * Run by prepack, so the list follows the dependencies.
 */
import { build } from 'esbuild';
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** License texts kept here for packages that publish none to npm (taken from their repositories). */
const FALLBACK = {
  '@fortune-sheet/core': 'build/licenses/fortune-sheet.txt',
  '@fortune-sheet/react': 'build/licenses/fortune-sheet.txt',
};

const { metafile } = await build({
  entryPoints: ['src/client/main.jsx', 'src/client/viewer.jsx'].map((entry) => join(ROOT, entry)),
  bundle: true,
  write: false,
  outdir: join(ROOT, 'build', '.licenses-tmp'),
  metafile: true,
  format: 'esm',
  platform: 'browser',
  external: ['react', 'react-dom'],
  loader: { '.css': 'text' },
  jsx: 'transform',
  logLevel: 'silent',
});

/** @type {Map<string, string>} package name -> its directory */
const packages = new Map();
for (const input of Object.keys(metafile.inputs)) {
  const match = /^(.*node_modules\/)((?:@[^/]+\/)?[^/]+)\//.exec(input);
  if (match && !packages.has(match[2])) packages.set(match[2], resolve(ROOT, match[1], match[2]));
}

const sections = [];
for (const [name, dir] of [...packages].sort(([a], [b]) => a.localeCompare(b))) {
  const pkg = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'));
  const license = pkg.license || pkg.licenses?.map((item) => item.type).join(' OR ') || 'see below';
  const file = readdirSync(dir).find((entry) => /^(licen[cs]e|copying)(\.(md|txt))?$/i.test(entry));
  const notice = readdirSync(dir).find((entry) => /^notice(\.(md|txt))?$/i.test(entry));
  const textPath = file ? join(dir, file) : FALLBACK[name] && join(ROOT, FALLBACK[name]);
  if (!textPath) throw new Error(`No license text for ${name}; add one to FALLBACK`);
  sections.push(`## ${name} ${pkg.version} (${license})\n\n\`\`\`\n${readFileSync(textPath, 'utf8').trim()}\n\`\`\`${notice ? `\n\nNOTICE:\n\n\`\`\`\n${readFileSync(join(dir, notice), 'utf8').trim()}\n\`\`\`` : ''}`);
}

const vendored = [
  ['SheetJS Community Edition 0.20.3 (Apache-2.0), in vendor/xlsx', 'vendor/xlsx/LICENSE'],
  ['xlwt 1.3.0 (BSD), in vendor/python', 'vendor/python/xlwt-1.3.0.dist-info/LICENSE.rst'],
];
for (const [title, path] of vendored) {
  if (!existsSync(join(ROOT, path))) throw new Error(`Missing ${path}`);
  sections.push(`## ${title}\n\n\`\`\`\n${readFileSync(join(ROOT, path), 'utf8').trim()}\n\`\`\``);
}

writeFileSync(join(ROOT, 'THIRD_PARTY_LICENSES.md'), `# Third-party licenses

dsh-doc-viewer bundles the following code into \`client.js\` and \`client.viewer.js\`, or ships it under \`vendor/\`. Their licenses follow in full.

${sections.join('\n\n')}
`);
console.log(`wrote THIRD_PARTY_LICENSES.md (${sections.length} entries)`);
