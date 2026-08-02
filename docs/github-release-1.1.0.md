# GitHub Release Draft: 1.1.0

## Title

`CWS Light Table 1.1.0`

## Short Release Summary

This release adds the first data-cleaning tools to `CWS Light Table`: removing
duplicate rows by selected columns, removing blank rows, and cleaning leading,
trailing, or inner spaces from table text, plus format conversion and suspicious
encoding issue lookup.

## Release Notes

`CWS Light Table 1.1.0` focuses on practical cleanup operations that fit the
single-sheet `CWS HTML` editing workflow.

What this release adds:

- `Data > Clean Data > Remove Duplicates`
- single-column and multi-column duplicate-row comparison
- duplicate preview counts before applying the cleanup
- `Data > Clean Data > Remove Blank Rows`
- blank-row preview counts before applying the cleanup
- `Data > Clean Data > Trim Spaces`
- `Data > Clean Data > Format Conversion`
- scope choices for current selection or the whole data area below the header
  row
- column selection for duplicate removal, while space cleanup follows the
  selected scope automatically
- configurable cleanup for leading spaces, trailing spaces, and inner spaces
- inner-space modes for collapsing consecutive spaces to one space or removing
  all inner spaces
- lowercase, uppercase, capitalized-word, full-width, and half-width conversion
- Japanese katakana support for full-width and half-width conversion, including
  voiced half-width katakana combinations; hiragana remains unchanged because
  there is no standard half-width hiragana form
- suspicious encoding issue lookup for mojibake, replacement characters, and
  control characters without modifying workbook data
- in-app Help page documentation for the data-cleaning tools

Behavior notes:

- duplicate removal keeps the first matching row and removes later duplicates
- the formal header row is not removed by duplicate cleanup
- blank-row cleanup removes only rows where every cell in the row is empty
- the formal header row is not removed by blank-row cleanup
- space cleanup treats half-width spaces, full-width spaces, and tabs as spaces
- case and width conversion follows the selected scope automatically
- all data-cleaning operations are undoable
- saved standalone `CWS HTML` files include the same data-cleaning runtime after
  rebuilding the standalone template

## Scope Notes

This release does not add fuzzy matching, automatic data-quality scoring,
formula recalculation, styling tools, or multi-sheet authoring.

## Verification

Release verification should include:

- `npm.cmd run build`
- JavaScript syntax checks for source and generated standalone runtime
- duplicate removal by one column
- duplicate removal by multiple columns
- blank-row removal in the current selection row range
- blank-row removal across the whole data area
- trim leading and trailing spaces with the default settings
- collapse consecutive inner spaces
- remove all inner spaces
- convert text to lowercase
- convert text to uppercase
- convert text to capitalized words
- convert text to full-width
- convert text to half-width
- convert Japanese katakana between full-width and half-width forms
- confirm hiragana remains unchanged during width conversion
- find suspicious encoding issues and jump to the flagged cells
- confirm encoding issue lookup does not create an undo history entry or modify
  workbook data
- confirm Help includes the data-cleaning section in all three UI languages
- undo and redo after each cleanup action
- saved standalone `file://` smoke testing for the same cleanup flows

## Homepage

```text
http://cws-light-table-docs.macsocloud.com/
```

## Suggested GitHub Release Tag

```text
v1.1.0
```
