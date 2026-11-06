# dsh-doc-viewer: third-party notices

This plugin bundles the following open-source packages. They are used locally;
no document content leaves the machine.

- FortuneSheet (`@fortune-sheet/core`, `@fortune-sheet/react`) — MIT
- SheetJS Community Edition 0.20.3 (bundled under `vendor/xlsx`) — Apache-2.0
- PDF.js (`pdfjs-dist`) — Apache-2.0
- ExcelJS — MIT
- mammoth — BSD-2-Clause
- pdf-lib — MIT
- docx — MIT
- marked — MIT
- xlwt (bundled under `vendor/python`) — BSD-3-Clause

The macOS Vision OCR helper is compiled from the Objective-C source in
`src/host/vision-ocr.m` and uses only system frameworks (Vision, PDFKit,
Foundation, CoreGraphics, ImageIO).

The full license texts of the code bundled into `client.js` / `client.viewer.js` or shipped under `vendor/` are in [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md). Packages installed from npm at runtime carry their own license files.
