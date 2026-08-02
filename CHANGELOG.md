# Changelog

## 1.1.1

- added `Data > Clean Data > Amount commas`
- supported adding thousands separators to strict numeric strings
- supported removing thousands separators from valid comma-formatted numeric
  strings
- skipped invalid numbers, leading-zero IDs, currency symbols, and text so
  non-numeric content is not modified
- added explicit `Global` and `Current selection` scope choices to
  `Find / Replace`
- defaulted `Find / Replace` to the current selection when a multi-cell range
  is selected before opening the panel
- kept amount-comma cleanup undoable and available in saved standalone CWS
  HTML files after rebuilding the standalone runtime
- updated public release metadata and CWS HTML generator version to `1.1.1`

## 1.1.0

- added `Data > Clean Data > Remove Duplicates`
- supported duplicate-row removal by one selected column or by a multi-column
  comparison key
- added `Data > Clean Data > Remove Blank Rows`
- removed fully blank rows from the current selection row range or from the
  whole data area below the header row
- added `Data > Clean Data > Trim Spaces`
- supported configurable leading, trailing, and inner-space cleanup
- added inner-space modes for collapsing consecutive spaces to one space or
  removing all inner spaces
- added `Data > Clean Data > Format Conversion`
- supported lowercase, uppercase, capitalized-word, full-width, and half-width
  text conversion
- expanded full-width and half-width conversion to Japanese katakana, including
  voiced half-width katakana combinations; hiragana remains unchanged because
  there is no standard half-width hiragana form
- added `Data > Clean Data > Find Encoding Issues` to locate suspicious
  mojibake, replacement characters, and control characters without modifying
  workbook data
- documented the data-cleaning tools in the in-app Help page
- improved capitalized-word detection across punctuation and non-ASCII letters
- kept data-cleaning actions undoable and available in saved standalone CWS
  HTML files after rebuilding the standalone runtime
- updated public release metadata and CWS HTML generator version to `1.1.0`

## 1.0.0

- added `Option > Language` with English, Simplified Chinese, and Japanese UI
  choices
- persisted the selected UI language in browser `localStorage` and applied it
  after reload
- localized application menus, toolbar labels, dialogs, status messages,
  validation feedback, and in-app help/version text
- kept workbook data, CWS HTML schema fields, file format values, encodings,
  MIME types, and import/export payloads language-neutral
- embedded standalone vendor script payloads so saved local CWS HTML files can
  import Excel without a sibling `vendor` folder
- fixed `Delete` key handling so it clears the entire selected cell range
  instead of only starting an edit on the active cell
- added `Help > Site` below `Version` to open the product documentation site
- updated public release metadata and CWS HTML generator version to `1.0.0`
- regenerated the saved standalone runtime payload from the source app shell

## 0.1.3

- added a top icon toolbar for `Open`, `Save`, `Print`, `Undo`, `Redo`,
  `Copy`, `Paste`, `Find / Replace`, `Advanced Filter`, `Clear Filters`, and
  `Grid Limits`
- added `Help` and `Version` pages and aligned in-app product naming around
  `CWS Light Table`
- added direct keyboard shortcuts for open, save, and print
- kept saved standalone files functionally aligned with the localhost editor
  for the shipped feature set
- disabled `Save CWS HTML` when a saved standalone file is opened directly
  through local `file://`, avoiding misleading save behavior in unsupported
  browser contexts
- replaced print-through-grid behavior with a dedicated browser print surface
  so the printed table no longer depends on the current editing viewport
- applied browser-print guardrails that block oversized print jobs and advise
  exporting to Excel for large-range printing
- prepared the public `0.1.3` release snapshot without benchmark datasets,
  internal comparison tooling, or standalone save test fixtures

## 0.1.2

- completed the `v0.1.2` large-grid optimization and verification cycle
- kept the stable accepted implementation baseline after later experimental
  changes failed to provide a reliable end-to-end performance win
- retained the shipped row-and-column virtualization path and the optimized
  large-write behavior from the accepted optimization work
- added `View` menu row-delete and column-delete actions, including multi-row
  and multi-column deletion based on the current selection span
- fixed blank-cell copy/paste so pasting an empty copied cell now correctly
  clears an existing target cell value
- added repeatable regression coverage for anchor import, large search/filter,
  structural insert actions, export, and save/reopen roundtrip
- added benchmark generation, benchmark-suite execution, and cross-baseline
  comparison tooling for large datasets
- documented the measured release conclusion so future optimization work starts
  from benchmark evidence instead of broad renderer changes

## 0.1.1

- first public open-source repository release for `CWS Light Table`
- standardized the public product name and repository-facing project metadata
- published the current lightweight editor feature set:
  - single-sheet `CWS HTML` editing
  - Excel and text-like import
  - `CSV` / `TSV` / `TXT` export
  - filtering, printing, grid limits, and structural insert actions
  - standalone `CWS HTML` save
- removed development-only tests, fixtures, logs, and temporary verification
  files from the public repository snapshot
- replaced remaining phase-style source naming with feature-based naming in
  retained public files
- added the initial public repository documentation, MIT license, third-party
  notices, contribution guide, and security policy
