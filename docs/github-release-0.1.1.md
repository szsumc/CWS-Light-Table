# GitHub Release Draft: 0.1.1

## Title

`CWS Light Table 0.1.1`

## Short Release Summary

This is the first open-source repository release of `CWS Light Table`, a
lightweight browser-based editor for simplified single-sheet `CWS HTML`
documents.

## Release Notes

`CWS Light Table 0.1.1` publishes the current lightweight editor as a clean
open-source repository snapshot under the MIT license.

This release focuses on practical single-sheet table work in the browser:

- open and save standalone `CWS HTML`
- import Excel workbooks (`.xlsx`, `.xlsm`, `.xls`)
- import text-like formats (`csv`, `tsv`, `txt`, `json`, `xml`)
- edit data in a spreadsheet-style grid with formula-bar support
- use range, row, column, and whole-sheet selection
- copy, paste, fill-repeat, undo, and redo
- apply find/replace, header-row control, filtering, and sorting
- configure page setup and print the current sheet
- export `CSV`, `TSV`, and `TXT`
- control editable sheet size through `Grid Limits`
- save a standalone HTML output that can open directly in a browser

This open-source packaging release also cleans up the repository for public
distribution:

- standardized the public project name to `CWS Light Table`
- removed internal development-only tests, fixtures, logs, and temporary
  verification files from the published repository snapshot
- replaced remaining phase-style source naming with feature-based naming where
  retained in the source tree
- added public repository documentation, licensing, third-party notices,
  contribution guidance, and security policy files

## Scope Notes

This release is intentionally lightweight and does not attempt to provide the
full feature surface of a large workbook suite.

Out of scope in `0.1.1`:

- multi-sheet authoring in one editing session
- formula recalculation
- formatting and styling authoring
- comments, drawings, shapes, and images
- advanced workbook-analysis tooling

## Homepage

```text
https://szsumc.github.io/CWS-Light-Table-Docs/
```

## Suggested GitHub Release Tag

```text
v0.1.1
```
