# Third-Party Notices

This repository depends on third-party packages. Their own license terms
continue to apply when the packages are installed, bundled, or redistributed.

## Runtime Dependency

| Package | Version | License |
| --- | ---: | --- |
| `xlsx` | 0.18.5 | Apache-2.0 |

## Standalone Build Inputs

The standalone runtime build currently reads vendor assets from the installed
dependency tree produced by `xlsx` and its transitive packages.

| Package | Version | License |
| --- | ---: | --- |
| `codepage` | 1.15.0 | Apache-2.0 |
| `cfb` | 1.2.2 | Apache-2.0 |
| `ssf` | 0.11.2 | Apache-2.0 |
| `crc-32` | 1.2.2 | Apache-2.0 |
| `adler-32` | 1.3.1 | Apache-2.0 |
| `frac` | 1.1.2 | Apache-2.0 |
| `wmf` | 1.0.2 | Apache-2.0 |
| `word` | 0.3.0 | Apache-2.0 |

## Notes

- Package versions and licenses listed here are taken from `package-lock.json`
  in this release.
- If dependencies change, review this file before publishing a new release.

