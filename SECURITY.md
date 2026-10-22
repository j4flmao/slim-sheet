# Security policy

slim-sheet reads and writes files on behalf of AI agents, so sandbox escapes, code execution,
data exfiltration and destructive writes are in scope.

## Reporting

Report vulnerabilities privately through
[GitHub security advisories](https://github.com/j4flmao/slim-sheet/security/advisories/new).
Please do not open a public issue. You will get an acknowledgement within 7 days.

## Design guarantees we treat as security bugs if broken

- Paths outside `SLIM_SHEET_ROOTS` are never read or written (symlinks included).
- `query` cannot write data or read other files (DuckDB external access is locked).
- No model-supplied code is executed; formulas that reach the network, DDE or other workbooks are refused on write.
- Originals are never modified without `overwrite=true`; every replaced file is journaled.
- Secrets (`SLIM_SHEET_TOKEN`, `SLIM_SHEET_PASSWORD`) come only from the environment and never appear in output.
- The HTTP transport refuses non-loopback binds without a bearer token.

## Supported versions

Only the latest release receives fixes.
