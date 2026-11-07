# Guide

English · [简体中文](guide.zh-CN.md)

## Using it

- **Documents** sits in the sidebar next to Automation tasks. Open documents, edits and undo history survive switching to a conversation and back; after a restart, the documents you had open open again.
- **Default viewer in Harness.** Clicking an xlsx / xls / csv / tsv / et / wps file in a conversation link or in the right panel's file tree opens it there as an editable table. The renderer menu at the top of the preview switches back to Harness's own preview. PDF, Word, Markdown and text still open in Harness's preview by default; choose "Documents plugin" in that menu to use this one.
- **Beside the conversation.** Documents can be a tab in Harness's right panel, next to the chat (Documents in the panel's "+" menu, or the side-by-side button in the Documents page; the same button goes back to the full page). While you are in a conversation, documents the agent opens appear beside it instead of taking you away. The full page and the side panel are the same page, so open documents, unsaved edits and undo history go with it.
- **Working with the agent.** The agent never pulls you out of the conversation; it puts a dot on the Documents entry. An open table updates in place, keeping your view, selection and undo history. "The agent changed N cells" in the toolbar lists each change, to jump to one by one or undo together. If you have unsaved edits, the agent's changes go into the table you are editing and are written with your next ⌘S, without a separate copy. Cells the agent wrote are outlined in blue with an "Agent" label until you click into the table or start typing (a minute at most), so there is time to read its answer first. Scrolling doesn't clear them, and nothing counts down while you are on another page.
- **Open.** "Open…" / ⌘O uses the system's open panel; the file opens where it is, not copied. Files dragged in from Finder reach the page without their path, so they are kept in the system's temporary folder and Save asks where to put them.
- **New table.** The "+" menu, "New table" in the empty page, or ⌘N. Name it in the system's save panel and it opens right away; the suggested name avoids existing files.
- **Save.** ⌘S writes back to the original file. Before overwriting, the old content is kept as a version (`~/.dsh/doc-viewer/versions`: the latest versions plus the oldest, 20 per file, kept 30 days, never in your folders). After saving you can "Undo save" step by step, or pick one from "Restore a previous version…".
- **No waiting to open.** Tables made by scripts or agents often hold formulas without results. The table shows first, with the common formulas already worked out; LibreOffice fills in the rest in the background (including functions such as XLOOKUP), and this doesn't count as an edit. LibreOffice takes ten seconds or so after Harness starts, and the table can be read and edited meanwhile.
- **Quitting without saving.** When you stop editing, unsaved edits are kept as a snapshot (`~/.dsh/doc-viewer/recovery`) and restored the next time the Documents page opens. You can also discard them.
- **No risky overwrites.** If the original holds something this plugin can't write back (charts, pivot tables, macros, conditional formats, data validation, comments, images), is an old `.xls/.et`, or was changed by another program since it was opened, it isn't overwritten: you are told why and asked where to save a copy.
- **Save / Save As… / Export PDF….** Save writes the original file. Save As… writes a new file elsewhere or in another format, and the tab follows the new file. Export PDF… makes a PDF and leaves the original alone. Save As and Export open the system's save panel, next to the original with the same name. Batch conversions and files the agent makes go next to their originals.
- **Ask Agent.** "Ask Agent" in the toolbar switches to the conversation with "About ‹file›, the selected range:" already in the input box. Error messages have the same button.
- **Shortcuts.** ⌘O open, ⌘N new table, ⌘S save, ⌘F find.
- **Language.** The interface and the agent's tool results follow the Harness language (Settings → General → Language): English or Chinese.
- With the system's Reduce Motion on, nothing animates.

| Type | What you can do |
| --- | --- |
| Tables: xlsx / xls / csv / tsv / et | Edit, live recalculation, references across sheets, merged cells, filters, find and replace, text to numbers, grouped statistics and charts, export a laid-out PDF |
| PDF | Continuous reading, thumbnails, selectable text, find with highlights; merge, split, delete pages, OCR (macOS Vision), convert to text / Word / Excel |
| Word: docx / doc / wps | Preview (sandboxed), export PDF, extract text |
| Text / Markdown | Edit and save in place, keeping the original encoding (UTF-8 / GB18030 / UTF-16) and line endings; find with highlights. Markdown shows a formatted preview by default (switch to edit); the preview and the exported PDF use the same parser |

## Exported PDFs

Tables: A4, columns sized to their content; wide tables turn landscape and shrink to one page wide. The header row is bold on a light fill and repeats on every page, grid lines are light, and the footer has the file name and page number. Colors and borders set in the file are kept. Text is laid out as a document, not as table cells.

## Agent tools

`doc_viewer_create` (a new xlsx / csv / md / txt; a table can come with header, rows and formulas, and opens beside the conversation; never replaces an existing file), `doc_viewer_context` (which document, sheet, selection or page the user is looking at), `doc_viewer_open` (show it to the user), `doc_viewer_read`, `doc_viewer_edit`, `doc_viewer_stats`, `doc_viewer_convert`, `doc_viewer_pdf`.

- Without `path`, read, edit and stats act on the document the user has open, so "sum the selected column" is enough.
- `doc_viewer_edit` calls on one file run in turn, so parallel calls don't overwrite each other. Edits are checked before anything is written: a formula with a syntax error, an unknown function, a reference to itself or a write inside a merged cell writes nothing and says what to fix. Other formula errors (such as `#DIV/0!`) are written and reported. The result lists each written cell before and after, and the formula results that moved because of it (totals included). The original file is changed by the same rules as ⌘S and a version is kept; if the user has unsaved edits in the page, the changes go into the page instead of the file. An open page updates in place without switching, and the changes join the "The agent changed N cells" list. Only when the page isn't running, or the edit reaches past the table's current size, is a copy written beside the file.
- `doc_viewer_stats` leaves out total rows (rows that `SUM` or `SUBTOTAL` the same column above them) so nothing is added twice, and says which rows it skipped; pass `includeTotals: true` to count them.
- How much `doc_viewer_read` returns at once is worked out from the running model's `contextWindow` (read through `llm.resolveModelInfo`), not written for any one model. Longer content is cut at a row, page or character boundary, with the parameters for the next read.

## Known limits

- Saving an xlsx writes each formula's result, calculated by LibreOffice, so other programs read correct values without calculating. Without LibreOffice, common formulas are worked out here: arithmetic, comparisons, `&`, references, SUM / AVERAGE / MIN / MAX / COUNT / COUNTA / ROUND / ROUNDUP / ROUNDDOWN / ABS / IF / IFERROR / AND / OR / NOT, and error values such as `#DIV/0!`. The rest are saved without a result and Excel recalculates them on opening, and function names can't be checked before writing. Functions Excel has and LibreOffice doesn't (such as `LAMBDA`) can't be written by the agent.
- The table's own engine lacks newer functions such as `XLOOKUP`, `FILTER` and `LET`. LibreOffice gives their results about a second after an edit, and `#NAME?` may show briefly meanwhile. Beyond 200,000 cells there is no live calculation; reopen after saving for correct results.
- Saving a table opens the original and writes only the changed cells (values, formulas, fonts, fills, alignment, number formats, borders, merges, row heights and column widths, frozen panes); inserted and deleted rows and columns are replayed on the original first. Hyperlinks, named ranges, hidden sheets, print settings, conditional formats, data validation, images, comments and theme colors are kept as they are.
- Files with charts, pivot tables, macros, slicers, form controls, sparklines or shapes / text boxes are never overwritten, only saved as a copy (without those parts).
- Sorting moves values only; cell formats stay where they were (unlike Excel). Rich text in a cell becomes plain text once the cell is edited.
- Images in Markdown show as "[Image: alt text]" in the preview and the PDF; image files aren't loaded.
- Saving in place through a symbolic link writes the file it points to and keeps its permissions. Hard links and extended attributes such as Finder tags aren't kept.
- Open and save panels: the system's own on macOS, zenity on Linux (untested). Without a panel, exports go next to the original.
- In dark mode the grid is shown inverted, so cell colors change on screen (not in the file).
- PDF to Word / Excel / CSV is a best effort; complex layout, images and columns are lost, and you are asked first.
- Old `.doc` / `.wps` / `.et` need the LibreOffice Kit that comes with Harness.
- OCR uses macOS Vision and needs macOS 13 or later (Apple silicon or Intel).
- Text files over 2 MB and files dragged in from Finder are read-only. Word and PDF can't be edited.
- A table of 20,000 rows and 40,000 formulas opens in about a second.
- Tested only on macOS. Linux and Windows are untried; reports are welcome.

## Development

```bash
pnpm install
pnpm build      # client.js (loaded at startup) and client.viewer.js (loaded when the page first opens)
pnpm test       # engine tests on real files; DSH_LIBREOFFICE_CLI=none simulates a machine without LibreOffice, as CI does
pnpm pack       # prepack builds the page and the OCR helper first (needs the Xcode command line tools)
```

```
index.js                 Host routes and plugin registration
src/host/
  documents.js           open / read / edit / stats / convert / PDF, saving in place, parse cache
  versions.js            versions kept before saving in place, and restoring them
  recovery.js            snapshots of unsaved edits, restored after a quit or crash
  native.js              system open and save panels
  print.js               PDF export layout (table page setup, text to document)
  xlsx-patch.js          writes only the changed cells back; finds what can't be written back
  spreadsheet.js         ExcelJS / SheetJS reading and writing, formula normalization, statistics
  formula-lite.js        common formulas when LibreOffice is unavailable
  encoding.js            text encoding detection and writing back in the same encoding
  pdf.js                 PDF.js text, pdf-lib page operations, Word preview, OCR
  libreoffice.js         officeToPdf, falling back to the LibreOffice Kit CLI
  tools.js               the doc_viewer_* agent tools
  i18n.js / locale-en.js language of Host messages, English dictionary
src/client/
  main.jsx               sidebar entry, main-area page, persistent mount, agent request polling
  viewer.jsx             the Documents page (lazy-loaded)
  SpreadsheetTab.jsx / PdfTab.jsx / DocumentTabs.jsx / Dialogs.jsx
  styles.js              Harness theme variables only, so it follows light/dark and skins
  i18n.js / locale-en.js interface language, English dictionary (the Chinese text is the key)
test/                    engine tests
```
