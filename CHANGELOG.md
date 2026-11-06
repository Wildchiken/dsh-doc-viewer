# Changelog

## 1.7.0

First public release.

- Renamed from `@local/doc-viewer` to `dsh-doc-viewer`. Remove the old package before installing.
- Text and Markdown files can be edited and saved in place, keeping their encoding and line endings.
- Cells the agent changed stay outlined until you click into the table or type (a minute at most).
- `doc_viewer_edit` reports each cell before and after, and the formula results that moved because of it.
- If LibreOffice can't recalculate, common formulas are worked out without it instead of being left blank.
- SheetJS upgraded to 0.20.3, fixing the known issues in 0.18.5.
- The OCR helper is built from source and runs on Apple silicon and Intel, macOS 13 or later.

## 1.6.1

- Workbooks open without waiting for LibreOffice; formula results are filled in afterwards.

## 1.6.0

- Consistent animations, skeleton loading, clearer wording, and an empty state that explains how to start.

Earlier versions were not published.
