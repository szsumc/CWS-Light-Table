# GitHub Release Draft: 1.1.1

## Title

`CWS Light Table 1.1.1`

## Short Release Summary

This release adds an `Amount commas` cleanup tool for numeric strings. Users can
add thousands separators or remove existing thousands separators from values
that match a strict numeric format. It also adds explicit `Global` and
`Current selection` scope choices to `Find / Replace`.

## Release Notes

`CWS Light Table 1.1.1` continues the data-cleaning work from `1.1.0` with a
focused number-format cleanup command.

What this release adds:

- `Data > Clean Data > Amount commas`
- add commas every three digits for strict numeric strings
- remove commas from valid comma-formatted numeric strings
- preview counts for checked cells, numeric cells, and cells that will change
- English, Simplified Chinese, and Japanese UI text for the new panel
- in-app Help page documentation for the new cleanup tool
- explicit `Global` and `Current selection` scope choices in `Find / Replace`
- automatic `Current selection` default when a multi-cell range is selected
  before opening `Find / Replace`

Behavior notes:

- only strict numeric strings are changed
- invalid comma placement such as `12,34` is skipped
- leading-zero IDs such as `001234` are skipped
- currency symbols, percent signs, scientific notation, and text are skipped
- positive and negative signs are preserved
- decimal portions are preserved
- leading and trailing spaces around valid numeric strings are preserved
- the operation follows the current selection or whole-data scope
- the operation is undoable
- saved standalone `CWS HTML` files include the same runtime after rebuilding
  the standalone template
- `Find / Replace` keeps the captured selection scope while stepping through
  matches, even though the active cell moves to each match

## Scope Notes

This release does not add currency parsing, locale-specific decimal separators,
scientific notation support, automatic date/time conversion, or regular
expression search.

## Verification

Release verification should include:

- `npm.cmd run build`
- JavaScript syntax checks for source and generated standalone runtime
- add commas to values such as `1234567` and `-1234567.89`
- remove commas from values such as `1,234,567` and `-1,234,567.89`
- confirm invalid values such as `12,34`, `001234`, `JPY1000`, and text are
  not modified
- confirm undo and redo after applying the amount-comma cleanup
- confirm the tool is available in saved standalone `CWS HTML`
- confirm `Find / Replace` can search and replace globally
- confirm `Find / Replace` can search and replace only inside the selected
  range
- confirm a multi-cell selection defaults the `Find / Replace` scope to
  `Current selection`
- confirm Help includes the amount-comma cleanup description
- confirm the version panel and generated metadata report `1.1.1`

## Homepage

```text
http://cws-light-table-docs.macsocloud.com/
```

## Suggested GitHub Release Tag

```text
v1.1.1
```
