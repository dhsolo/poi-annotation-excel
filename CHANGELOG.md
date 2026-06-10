# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Security
- Apache POI upgraded 5.2.5 → 5.5.1 (covers CVE-2025-31672 in OOXML parsing — the import path
  parses user-supplied xlsx). Picture pre-registration was ported to POI ≥5.4's unmodifiable
  `getAllPictures()` by appending through the workbook's internal picture list, with a guard
  that fails loudly if POI internals change again.

### Changed
- Charset detection for CSV import now uses the maintained `com.github.albfernandez:juniversalchardet`
  fork (2.5.0) instead of the unmaintained 2011 `com.googlecode` artifact (same package and API).
- Tests route POI's internal log4j-api logging into slf4j (`log4j-to-slf4j`, test scope),
  removing the spurious `StatusLogger` ERROR line from every build.

### Fixed
- The CSV adapter now follows the POI contracts: `getLastCellNum()` returns last index **plus
  one** (so the trailing required-column check works for CSV exactly as for real sheets),
  out-of-range `getRow`/`getCell` return null instead of throwing, the row/cell iterators work
  (for-each over a CSV sheet/row no longer NPEs), `getPhysicalNumberOfRows` reports the real
  count, and the CSV reader is closed. CSV import errors now name the sheet as `CSV` instead of
  an empty string.
- A per-sheet start row (`addSheetStartRow`) no longer leaks into subsequent sheets that have
  no entry of their own; they fall back to the global `setStartRow` value.
- An empty `BusinessSXSSFSheet` reports `getLastRowNum() == -1` (POI contract) instead of 0.

## [1.1.1] - 2026-06-10

### Fixed
- Importing from `sheetIndex > 0` via `ExcelUtil.importExcel`/`importExcelToMap` works: column
  models used to be registered for sheet 0 regardless, so the Map variant returned one empty
  map per row (silent data loss) and the typed variant threw `IndexOutOfBoundsException`.
  `ExcelImportor` gained `addColumnName(int sheetNum, models)`; sheets without registered
  columns are now skipped during parsing instead of producing empty rows.
- `ImageUtils.resizeImage` no longer crashes on `TYPE_CUSTOM` images (a frequent ImageIO decode
  result): the resize target falls back to `TYPE_INT_ARGB`/`TYPE_INT_RGB`, so resize-enabled
  exports stop replacing such images with placeholders.
- The column clones for `mergeCellIndex > 1` are now a field-by-field shallow copy instead of
  Java-serialisation deep cloning, which threw `NotSerializableException` when the column
  carried a non-serialisable member (e.g. an `@ExcelTranslateMethod` lambda).
- The multi-sheet assembly helpers in `ExcelUtil` (`export`/`exportLocal`/`toInputStream`/
  `toBytes` over multiple creators or builders) now close every creator, not just the first,
  keeping the shared download-pool instance accounting balanced.
- `ExcelTemplateFiller`: pictures anchored below the list region now shift together with the
  rows inserted by list expansion (POI's `shiftRows` moves cells but not drawing anchors, so a
  template's footer stamp/QR picture used to end up overlapping the inserted rows).

## [1.1.0] - 2026-06-10

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
- Configurable validation/formula row coverage: dropdown-list validations and `@ExcelFormula`
  pre-fill cover `@ExcelInfo(validateRowCount = ...)` data rows (default `1000`, previously
  hardcoded); also settable via `ExcelCreator.setValidationRowCount(int)` and the builder's
  `validationRowCount(int)`.

### Changed
- `@ExcelInfo` attribute `orderMergeIndex` renamed to `orderColumnSpan` for clarity.
- Cell value reads compile getters via `LambdaMetafactory` (cached), and the per-column
  resolver is selected once per column instead of per cell.
- `autoSizeColumns(true)` now estimates widths from a bounded row sample instead of POI's
  per-cell font metrics (faster; no longer throws on streaming sheets). Formula columns are
  measured by their cached computed result rather than the formula expression text, so a long
  formula (e.g. `SUM(A1:A100)`) no longer over-widens its column.
- `ExcelCreator.close()` now documents (and relies on) the close-after-export contract: in
  big-data mode it disposes the streaming workbook's temp files, so exporting after `close()`
  is not supported.

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
  styles against the shared workbook. A child's **custom** title/header/data styles (set before
  it was attached) are now carried over by cloning them into the shared workbook, instead of
  being silently reset to the defaults.
- `@ExcelCustomValidateMethod` is now actually applied on import (it was previously collected
  but never wired to the column); the annotated method must return an `ExcelCustomValidate`.
- Numeric and boolean cell values are now written with their native cell type instead of as
  text (previously every value was stored as a string, breaking sorting/aggregation).
- Date/time values (`Date`, `Calendar`, `LocalDate`, `LocalDateTime`, `LocalTime`) are written as
  typed, date-formatted cells; `@ExcelDateFormat(pattern = ...)` is now honored on export (it was
  previously ignored, leaving dates as plain text). `LocalTime` is stored as a fraction-of-day time
  cell (default format `HH:mm:ss`) instead of plain text.
- Image-heavy top-level sheets now use the memory-safe disk-staging path (previously a
  gating bug forced in-memory image embedding).
- Windows file-lock failure when injecting images into the temporary workbook ZIP.
- Javadoc generation on project paths containing non-ASCII characters.
- **Import precision loss**: numeric cell values were silently rounded to 3 fraction digits by
  `NumberFormat.getInstance()` during DOM import (`0.123456789` became `"0.123"`, and date-time
  serial values lost up to ~43 seconds). Values are now converted via `BigDecimal` and keep their
  full precision; integer-valued cells still render without a trailing `.0`.
- **Custom validation skipped rows beyond the first**: `@ExcelCustomValidateMethod` snapshots
  were captured only for the first data row and re-validated against later rows, so an invalid
  value in row 2+ passed silently. Snapshots are now rebuilt per row; additionally, an error in an
  earlier row no longer disables custom validation for all subsequent rows when
  `firstErrorBreak = false`.
- **`@ExcelColumnParent` + `needMergeCell` crash**: the vertical-merge base row did not account
  for the extra grouped-header row, landing the data merge one row too high and failing with an
  overlapping-region `IllegalStateException` (superseded by the captured-first-data-row fix
  below, which handles this and the related layouts uniformly).
- **Multi-image column expansion desynced later columns**: the physical-to-data column offset
  created by multi-picture expansion was overwritten per column instead of accumulated, causing a
  `NullPointerException` (or shifted values) once two or more ordinary columns followed a
  multi-image column. The offset now accumulates; `CellResolveContext` gained a `dataColIndex`
  component (the old constructor signature is preserved), and `PictureCellResolver` looks up the
  per-column image count by data column index, fixing the same desync when an order column
  (`needOrder = true`) shifts the layout.
- Vertical merges (`needMergeCell`) now use the **actual first data row captured when data
  population starts**, instead of re-deriving it from the title/custom-row/grouped-header layout
  rules. This also fixes the merge landing at the top of the sheet for **complex multi-section
  sheets** and the off-by-one when a before-title custom row exists on a title-less sheet.
- The `needMergeCell` target column now shifts by the whole order block (`orderColumnSpan`) and
  by any multi-picture expansion to its left, instead of always shifting by one.
- `orderColumnSpan > 1` no longer breaks the header row with an
  `ArrayIndexOutOfBoundsException` (the order-block merge consumed a column and the header
  label index only skipped one column).
- `formart` passes `NaN`/`Infinity` through as text instead of failing in `BigDecimal`.
- Cascade dropdowns: option values that cannot form a legal Excel defined name (illegal
  characters such as spaces or hyphens, length over 255, cell-reference-like values such as
  `A1`, or empty values) now fail fast with an `ExcelException` naming the offending option
  chain, instead of surfacing a bare POI `IllegalArgumentException` or a
  `StringIndexOutOfBoundsException`.
- Cascade dropdowns: defined-name deduplication is now case-insensitive (Excel names are);
  parent values differing only by case (`ABC` vs `abc`) no longer crash with
  "The workbook already contains this name".
- Cascade dropdowns: option values containing characters Excel forbids in defined names
  (spaces, hyphens, slashes, parentheses, ...) now work transparently — the registered name
  substitutes them with `_` and the `INDIRECT` formula mirrors the substitution via
  `SUBSTITUTE(...)`. Overlong (&gt;255) chains, cell-reference-like values, and empty values
  still fail fast with an `ExcelException`.
- Dropdown/formula validation columns (`realIndex`) shift by the whole order block: with
  `orderColumnSpan > 1` the validation landed on the wrong physical column (the span was also
  assigned after the index computation consuming it, so it never took effect).
- Streaming (`isBigData`) workbooks: `close()` now disposes the SXSSF backing temp files (they
  used to accumulate until JVM exit); a child sheet stitched into a parent workbook disposes its
  own abandoned workbook as well.
- Shared image-download pool accounting is ownership-based: a wrapper creator built via
  `ExcelCreator(Workbook)` no longer drives the instance count negative on `close()` (which shut
  the pool down underneath other active exports), double `close()` is idempotent, and a re-init
  via `createWorkBook` no longer double-counts.
- The synchronous picture fallback path (taken when the disk-staging directory is unavailable)
  now goes through the same guarded stream as the async downloader: protocol whitelist,
  `ImageDownloadPolicy` SSRF checks, timeouts, and the 64 MB size cap (it used to bypass all of
  them).
- A sheet with a title and a single column no longer crashes on a one-cell title merge.
- Cell resolvers are matched in `getOrder()` priority (picture > handler > translate > plain) as
  documented; a custom handler used to hijack `@ExcelImage` columns and desync multi-picture
  layouts.
- Importing a sheet containing non-picture shapes (text boxes etc.) no longer throws
  `ClassCastException` during picture extraction.
- `DefaultExcelExporter.getWorkBook()` caches the workbook re-read from the picture-injected
  temp file, so a second call no longer returns the original picture-less workbook.
- `@ExcelColumn(columnWidth)` now applies to the column's physical position: with
  `needOrder = true` every custom width used to land one column to the left, on the order
  column. Data-to-physical column translation (order block + multi-picture expansion) is now
  centralised in one helper shared by widths, vertical merges, and validation columns.
- Complex multi-section sheets: each section's layout (header expansion, multi-image alignment,
  vertical merges) now uses only its own multi-picture expansion entries; the picture handler's
  cross-section mapping mixed colliding data-column keys, letting a parent section's multi-image
  column shift a child section's columns. `PictureHandler` gained `getSectionColumnMaxMapping()`
  and a mapping-explicit `expandHeaderForPictures` overload (defaulted for compatibility).
- `export`/`upload`/`getInputStream`/`exportLocal` serialise the picture-injected workbook when
  `getWorkBook()` was called first (the original in-memory workbook lacks the injected
  pictures).
- Import value translation no longer fails on comparator-based translate maps (e.g.
  `TreeMap`) combined with non-string cell values; the typed lookup falls back to the
  stringified scan.
- Appended child sheets (`getChild().add(...)`): pictures and dropdown validations used to
  land in the child's discarded original workbook and silently vanish from the export. The
  child now shares the parent's picture handler (rebound to the child's sheet via the new
  `PictureHandler.bindSheet`; download directory, image numbering, and ZIP staging are
  coordinated workbook-wide) and rebuilds its validator against the shared workbook.
- Complex multi-section sheets: child sections never built a data validator, so
  `@ExcelListBox`/cascade/`@ExcelFormula` on a child section were silently skipped; the
  validator is now created against the shared sheet, and `currentListNum`/name counters are
  carried back to the parent.
- Dropdown/formula validation columns now account for multi-picture expansion: `realIndex` is
  refreshed from the physical layout after picture analysis, so a dropdown to the right of a
  multi-image column lands on its real column.
- `ExcelExporter` gained a `close()` lifecycle hook (default no-op); the workbook re-read from
  the picture-injected temp file is now closed by `ExcelCreator.close()` instead of leaking its
  `OPCPackage`.

### Deprecated
- `ExcelImportor(Object)`: the constructor reads nothing and yields an unusable instance; use
  `ExcelImportor(InputStream)`.

### Security
- Removed unused TLS-bypass constants (trust-all certificates / hostname verifier).
- Remote image downloads are restricted to http/https and capped in size.
- Opt-in SSRF guard (`ImageDownloadPolicy.setBlockPrivateNetworks(true)`) rejects image URLs
  resolving to loopback/link-local/site-local/private addresses (off by default).
