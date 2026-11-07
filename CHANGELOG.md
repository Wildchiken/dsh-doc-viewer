# Changelog

## 1.9.0

- Large tables whose formulas have no saved results open in about a second instead of freezing the page: 20,000 rows with 40,000 formulas took 40 seconds before.
- Numbers show at most 15 significant digits, as in Excel: 27 × 117.29 reads 3166.83, not 3166.8300000000004.
- Without LibreOffice, the built-in calculator also handles comparisons, `&`, IF, IFERROR, AND, OR, NOT, COUNTA, ROUNDUP, ROUNDDOWN and error values such as `#DIV/0!`.

## 1.8.2

- The Documents entry in the sidebar follows the Harness language again; in 1.8.0 and 1.8.1 it always read "Documents".

## 1.8.1

- `doc_viewer_stats` leaves out total rows (`=SUM` / `=SUBTOTAL` over the same column above), so nothing is added twice; `includeTotals: true` counts them.
- English messages agree with their counts ("1 row", "3 rows").
- Without LibreOffice, notes say so once and plainly, and a formula with no result yet reads "not calculated yet" instead of "empty".
- English guide.

## 1.8.0

- The interface and the agent's tool results follow the Harness language: English or Chinese.
- Messages read naturally in English, including punctuation and number formats.

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
