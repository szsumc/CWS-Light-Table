# Feature Overview

`CWS Light Table` is a lightweight browser-based editor for a simplified
single-sheet `CWS HTML` workflow.

## Primary Features

- open existing `CWS HTML` files
- save the current sheet as standalone `CWS HTML`
- switch the application UI between English, Simplified Chinese, and Japanese
- open the product documentation website from `Help > Site`
- import one selected sheet from `.xlsx`, `.xlsm`, or `.xls`
- import `csv`, `tsv`, `txt`, `json`, and `xml` into the current sheet
- quick-action icon toolbar for common commands
- edit cells directly in the grid or formula bar
- use row, column, range, and whole-sheet selection
- copy, paste, fill-repeat, clear the current selection with `Delete`, undo,
  and redo
- set a formal header row
- apply quick filters and advanced filters
- remove duplicate rows by one selected column or a multi-column key
- remove fully blank rows from the current selection row range or the whole
  data area
- clean leading, trailing, and inner spaces with configurable inner-space
  handling
- convert text to lowercase, uppercase, capitalized words, full-width, or
  half-width, including Japanese katakana width conversion; hiragana remains
  unchanged because there is no standard half-width hiragana form
- add or remove thousands separators for strict numeric strings while skipping
  invalid numbers, leading-zero IDs, currency symbols, and text
- find suspicious encoding issues such as mojibake, replacement characters,
  and control characters without modifying workbook data
- choose global or current-selection scope in find and replace
- print with lightweight page setup and a dedicated print rendering surface
- export `CSV`, `TSV`, and `TXT` with encoding and line-ending controls
- delete selected rows or columns from the `View` menu
- control the current editable grid size through `Grid Limits`
- reopen saved standalone `CWS HTML` files with embedded runtime dependencies
  for supported import and export flows

## Design Boundaries

This project intentionally does not aim to be a full workbook suite.

Excluded areas include:

- multi-sheet authoring in a persistent workbook session
- formula recalculation
- cell styling and formatting design tools
- comments, notes, hyperlinks, drawings, shapes, and images
- charting, pivot tools, and workbook-analysis features
