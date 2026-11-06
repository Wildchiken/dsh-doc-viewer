/**
 * Build the macOS Vision OCR helper (bin/vision-ocr) from src/host/vision-ocr.m.
 * A universal binary (Apple silicon and Intel), macOS 13 or later. Elsewhere, or
 * without Xcode's command line tools, it says so and leaves OCR unavailable.
 */
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const source = join(root, 'src', 'host', 'vision-ocr.m');
const output = join(root, 'bin', 'vision-ocr');

if (process.platform !== 'darwin') {
  console.log('vision-ocr: skipped (OCR needs macOS)');
  process.exit(0);
}
mkdirSync(dirname(output), { recursive: true });
try {
  execFileSync('xcrun', [
    'clang', '-fobjc-arc', '-O2',
    '-arch', 'arm64', '-arch', 'x86_64',
    '-mmacosx-version-min=13.0',
    '-framework', 'Foundation', '-framework', 'Vision', '-framework', 'PDFKit',
    '-framework', 'ImageIO', '-framework', 'CoreGraphics',
    source, '-o', output,
  ], { stdio: 'inherit' });
  console.log('wrote bin/vision-ocr');
} catch (error) {
  if (existsSync(output)) {
    console.warn(`vision-ocr: build failed (${error.message}); keeping the existing binary`);
  } else {
    console.error('vision-ocr: build failed; install Xcode command line tools (xcode-select --install)');
    process.exit(1);
  }
}
