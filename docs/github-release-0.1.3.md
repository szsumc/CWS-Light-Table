# GitHub Release Draft: 0.1.3

## Title

`CWS Light Table 0.1.3`

## Short Release Summary

This release hardens the daily editing workflow for `CWS Light Table` with a
quick-action toolbar, clearer standalone save behavior, and a dedicated
browser print rendering path.

## Release Notes

`CWS Light Table 0.1.3` focuses on practical usability improvements and on
making browser printing safer and more predictable.

What this release adds:

- a top icon toolbar for common actions such as open, save, print, undo,
  redo, copy, paste, find, filtering, and grid limits
- `Help` and `Version` pages inside the application shell
- dedicated print rendering that no longer reuses the live editing grid
- browser-print guardrails for oversized print ranges
- aligned standalone saved-file behavior so shipped features stay available
  after reopening the saved HTML file directly in a browser

What this release changes:

- local `file://` reopen sessions now disable `Save CWS HTML` when the browser
  environment cannot provide a reliable save-back path
- the public release snapshot removes internal benchmark data, comparison
  helpers, and standalone save test fixtures from the shipped package

## Scope Notes

`0.1.3` improves user flow and print correctness without changing the core
product scope.

Out of scope in `0.1.3`:

- multi-sheet authoring in one editing session
- formula recalculation
- styling and formatting authoring
- comments, drawings, shapes, and images
- full Excel-style manual page-break editing

## Homepage

```text
https://szsumc.github.io/CWS-Light-Table-Docs/
```

## Suggested GitHub Release Tag

```text
v0.1.3
```
