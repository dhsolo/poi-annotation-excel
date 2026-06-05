# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- `@ExcelColumnParent` is now implemented: renders a two-row grouped header (parent label merged
  over its child columns). A new `value()` attribute supplies the parent label; child columns read
  their values via `sourceField`/`sourcePath`.
- `@ExcelInfoChild` is now implemented for both export and import: a nested object's
  `@ExcelColumn` columns are flattened into the parent sheet (ordered with the parent columns by
  `index()`); on import the nested object is rebuilt and populated via path. Flattened columns are
  keyed by their distinct `sourcePath`, so same-named child fields of different parents
  (e.g. `customer.name` vs `supplier.name`) no longer collide on import. Nesting is recursive
  (multi-level, e.g. `addr.geo.zip`) with type-cycle detection.
- Low-memory streaming `.xlsx` import via `StreamingExcelReader` (SAX): `read`,
  `readAsMaps`, and `readAsBeans` (bind by `@ExcelColumn` position or by header name),
  with leading-row skipping and numeric/boolean/`BigDecimal`/date-time conversion.
- `ImageDownLoadTask` raw byte passthrough (no decode) when no resize is needed, preserving
  the original image format (PNG/JPEG/GIF) including transparency.
- Parallel image download overlapped with data population; downloaded media injected into
  the workbook ZIP as `STORED` entries (no wasteful re-deflate).
- Disabled-by-default import benchmark (`ImportBenchmarkTest`) comparing DOM vs SAX import.

### Changed
- `@ExcelInfo` attribute `orderMergeIndex` renamed to `orderColumnSpan` for clarity.
- Cell value reads compile getters via `LambdaMetafactory` (cached), and the per-column
  resolver is selected once per column instead of per cell.
- `autoSizeColumns(true)` now estimates widths from a bounded row sample instead of POI's
  per-cell font metrics (faster; no longer throws on streaming sheets).

### Fixed
- Large integers (`Long`/`BigInteger` beyond 2^53) and high-precision `BigDecimal` are written as
  text on export and parsed exactly on streaming import, instead of being silently rounded through
  a `double` (e.g. 19-digit ID numbers).
- `@ExcelColumnParent`: a non-grouped column whose header repeats an adjacent name no longer
  triggers an overlapping-merge exception (its vertical merge is skipped); a group with
  non-contiguous child indices now fails fast with a clear `ExcelAnnotationException` instead of
  rendering a broken header.
- Multi-sheet export from independently built `ExcelCreatorBuilder`/`ExcelCreator` no longer
  fails with "Style does not belong to the supplied Workbook" — child sheets rebuild their
  styles against the shared workbook.
- `@ExcelCustomValidateMethod` is now actually applied on import (it was previously collected
  but never wired to the column); the annotated method must return an `ExcelCustomValidate`.
- Numeric and boolean cell values are now written with their native cell type instead of as
  text (previously every value was stored as a string, breaking sorting/aggregation).
- Date/time values (`Date`, `Calendar`, `LocalDate`, `LocalDateTime`) are written as typed,
  date-formatted cells; `@ExcelDateFormat(pattern = ...)` is now honored on export (it was
  previously ignored, leaving dates as plain text).
- Image-heavy top-level sheets now use the memory-safe disk-staging path (previously a
  gating bug forced in-memory image embedding).
- Windows file-lock failure when injecting images into the temporary workbook ZIP.
- Javadoc generation on project paths containing non-ASCII characters.

### Security
- Removed unused TLS-bypass constants (trust-all certificates / hostname verifier).
- Remote image downloads are restricted to http/https and capped in size.
- Opt-in SSRF guard (`ImageDownloadPolicy.setBlockPrivateNetworks(true)`) rejects image URLs
  resolving to loopback/link-local/site-local/private addresses (off by default).
