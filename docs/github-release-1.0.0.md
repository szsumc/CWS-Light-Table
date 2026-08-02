# GitHub Release Draft: 1.0.0

## Title

`CWS Light Table 1.0.0`

## Short Release Summary

This release promotes `CWS Light Table` to `1.0.0` with a multilingual
application UI, stronger saved-standalone behavior, and small editing workflow
fixes for daily table work.

## Release Notes

`CWS Light Table 1.0.0` keeps the lightweight single-sheet `CWS HTML` editing
scope while making the application easier to use across English, Simplified
Chinese, and Japanese environments.

What this release adds:

- `Option > Language` with English, Simplified Chinese, and Japanese choices
- persistent UI language preference through browser `localStorage`
- localized menus, toolbar labels, dialogs, status messages, validation
  feedback, and in-app help/version text
- `Help > Site` below `Version`, opening the product documentation website
- embedded standalone vendor script payloads so saved local `CWS HTML` files
  can import Excel without a sibling `vendor` folder

What this release fixes:

- pressing `Delete` now clears the entire selected cell range
- saved standalone `CWS HTML` files keep Excel import support when opened from
  a local folder outside the project tree

What remains language-neutral:

- workbook cell data
- `CWS HTML` schema fields and internal model keys
- import/export format values
- encodings, MIME types, file extensions, and generated payload bytes
- advanced filter logic syntax, which continues to use `AND` / `OR`

## Scope Notes

`1.0.0` is still focused on a simplified single-sheet workflow. It does not add
multi-sheet authoring, formula recalculation, styling authoring, comments,
drawings, charts, or pivot-table features.

## Verification

Release verification included:

- standalone runtime rebuild with `npm.cmd run build`
- JavaScript syntax checks for source and generated standalone runtime
- browser smoke tests for language switching
- browser smoke test for `Delete` clearing a selected `2x2` range and undoing
  the change
- local `file://` saved-standalone smoke test importing an Excel workbook
- browser smoke test confirming `Help > Help`, `Help > Version`, `Help > Site`
  order and the documentation site open target

## Homepage

```text
http://cws-light-table-docs.macsocloud.com/
```

## Suggested GitHub Release Tag

```text
v1.0.0
```
