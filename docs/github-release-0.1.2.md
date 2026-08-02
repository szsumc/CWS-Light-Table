# GitHub Release Draft: 0.1.2

## Title

`CWS Light Table 0.1.2`

## Short Release Summary

This release closes the first large-grid optimization cycle for
`CWS Light Table` with a measured, conservative shipping decision.

## Release Notes

`CWS Light Table 0.1.2` focuses on large dataset practicality and on making
performance work evidence-driven.

What this release adds:

- accepted row-and-column virtualization improvements from the `v0.1.2`
  optimization work
- accepted bulk-write improvements for large paste and CSV import flows
- repeatable regression validation for anchor import, large search/filter,
  structural insert actions, export, and save/reopen behavior
- reusable benchmark dataset generation and benchmark-suite tooling for large
  row-and-column cases
- documented cross-baseline comparison against the stable accepted
  implementation baseline

Shipping decision in this release:

- keep the stable accepted implementation baseline
- do not retain later experimental renderer changes that failed to show a
  reliable end-to-end performance win
- preserve correctness and interaction stability over speculative complexity

This means `0.1.2` is a hardening and measurement release as much as a
performance release: it keeps the proven gains, adds repeatable diagnostics,
and records a clear stop point for any future optimization work.

## Scope Notes

`0.1.2` does not expand the functional product scope beyond the existing
single-sheet editor model.

Out of scope in `0.1.2`:

- multi-sheet authoring in one editing session
- formula recalculation
- formatting and styling authoring
- comments, drawings, shapes, and images
- risky renderer experiments without measured payoff

## Homepage

```text
https://szsumc.github.io/CWS-Light-Table-Docs/
```

## Suggested GitHub Release Tag

```text
v0.1.2
```
