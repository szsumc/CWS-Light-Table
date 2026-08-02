# Security Policy

## Reporting

If you discover a security issue in CWS Light Table, please avoid posting full
exploit details in a public issue before the problem is understood and a fix
path is available.

Use the project homepage for contact and project updates:

```text
https://szsumc.github.io/CWS-Light-Table-Docs/
```

## Scope Notes

This project runs as a local browser-based editor and can open external files.
Useful reports include:

- malformed-file handling that leads to unsafe execution
- unsafe HTML or script injection through imported workbook content
- file-save or import flows that bypass expected browser security boundaries
- standalone saved-file behavior that introduces script or data-integrity risks

General feature requests and usability bugs should go through the normal issue
tracker rather than the security process.
