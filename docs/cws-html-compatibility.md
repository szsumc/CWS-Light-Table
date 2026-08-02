# CWS HTML Compatibility

`CWS Light Table` uses `CWS HTML` as its native save format, but it does not
implement the full workbook surface of larger `CWS` applications.

## Supported Profile

Saved files from this project are designed around:

- one active visible sheet
- plain stored cell values
- embedded standalone HTML runtime
- lightweight page setup metadata
- lightweight grid-limit metadata
- formal header-row metadata

## Flattened Or Discarded On Save

The lightweight save path does not preserve:

- formulas as executable formulas
- style objects
- merges
- images
- shapes
- comments
- workbook-level view state
- validation, drawing, or formatting metadata

When importing workbook formats with richer features, this project stores the
plain visible value that best fits the simplified sheet model.

## Embedded Workbook Payload

Workbook data lives in:

```html
<script id="websheet-model" type="application/json">
```

Saved files also include:

- `meta name="cws:format" content="CWS_HTML"`
- `meta name="cws:schema-version" content="1"`
- an AI guidance payload
- the embedded standalone runtime script

## Purpose Of The Public Schema And Guide

This repository publishes:

- a machine-readable JSON guide
- a JSON Schema for the simplified workbook model

These documents describe the lightweight profile implemented by this project,
not the complete feature surface of larger `CWS` variants.

