# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.1.0] - 2026-02-28

### Added

- Number formatting (`numfmt` package): full cell value rendering including date/time,
  scientific notation, fractions, sub-second decimals, currency, text sections, and the
  `General` format (up to 10 significant digits).
- Style handling (`styles` package): parse `styles.bin` to resolve per-cell number
  format IDs; built-in format IDs 5–8, 14, 20–22, 27–36, 37–44, 47, 50–58 are now
  correctly mapped, including CJK locale fallback date formats.
- Sheet visibility: `Workbook.SheetVisible(name)` and `Workbook.SheetVisibility(name)` methods.
- Workbook date-system detection: `date1904` flag propagated through all date/time
  rendering paths so the 1904 date system is handled correctly.
- Cell error string mapping: Excel error codes (`#NULL!`, `#DIV/0!`, etc.) are now
  returned as human-readable strings instead of raw byte values.
- Chinese AM/PM token support (`上午`/`下午`) in `renderDateTime`.
- Fixed-denominator fraction support (e.g. `# ?/4`, `# ?/8`) via
  `TokenTypeDenominator` in `renderFraction`.
- `isDateFormat` detection extended to recognise formats containing `s`/`S` (e.g.
  `:ss`) across `xlsb.go`, `styles/styles.go`, and `workbook/workbook.go`.

### Fixed

- `renderDateTime`: midnight roll-over — when a fractional serial rounds up to 86400
  seconds (e.g. one microsecond before midnight), the day now advances instead of
  clamping to 23:59:59, matching Excel and Excelize behaviour.
- `renderDateTime` / `convertSerial`: minute/hour rounding now mirrors Excelize's
  half-second rounding logic (`roundEpsilon` + nearest-second rounding), eliminating
  off-by-one minute/hour errors near time boundaries.
- `excelDay` helper: serial < 1 returns day 0 (pure time); Lotus 1-2-3 off-by-one
  corrected for serials 1–59; fake Feb-29-1900 (serial 60) handled; 1904 date system
  bypasses 1900-only corrections.
- Built-in format IDs 20–22 now use zero-padded hours (`hh`) to match Excelize
  rendering.
- Built-in format ID 14 corrected from `MM-DD-YY` to `mm-dd-yy`.
- `renderNumber`: sign wrapper detection for two-section negative formats (e.g. `0;0`
  applied to `-5` now produces `-5` instead of `5`).
- `renderElapsed`: `[mm]`/`[ss]` tokens now return total elapsed minutes/seconds
  without modulo 60, using `int64` to avoid overflow on large serials.
- `renderNumber`: `@` (text placeholder) on numeric cells now renders as `General`
  instead of being silently dropped.
- Guards added in `renderDateTime` and `renderNumber` to prevent silent dropping of
  numeric values on unexpected token sequences.
- `isMinuteIndex` pre-scan in `renderDateTime` disambiguates `M`/`MM` as minutes
  using both look-behind (`H`/`HH`) and look-ahead (`S`/`SS`).

## [1.0.2] - 2026-02-28

### Fixed

- `record.Reader.readLen`: changed LEB-128 accumulator and byte variable from
  `int` to `uint32`, matching the existing `readID` implementation.  On 32-bit
  platforms the behaviour was already correct for all valid input, but using a
  signed architecture-dependent type was inconsistent and silently relied on the
  `maxRecordLen` guard keeping values in range.
- `worksheet.parseColRecord`: added a `maxStyleIndex` (`0x7FFFFFFF`) bounds
  check on the raw `uint32` style index before casting to `int`.  A corrupt
  value above `math.MaxInt32` previously produced a different `int` result on
  32-bit vs 64-bit platforms; it now returns a descriptive error on all
  platforms.

## [1.0.1] - 2026-02-28

### Added

- Updated godoc comments

## [1.0.0] - 2026-02-28

### Added

- Merge-cell value propagation in `Rows()`: non-anchor cells in a vertical merge
  now carry the anchor cell's value, matching the visual appearance in Excel.
  Previously, only the top-left anchor cell of a merged region held a value;
  all other rows in the merge returned `nil`.

## [0.2.0] - 2026-02-28

### Added

- Initial public release of the `go-xlsb` library.
- Pure-Go reader for Microsoft Excel Binary Workbook (`.xlsb`) files with no cgo dependency.
- `Open` function to open `.xlsb` files by path.
- `Workbook` type with `Sheets()` and `Sheet(n)` methods for workbook navigation.
- Row and cell iteration via `sheet.Rows()`.
- `ConvertDate` helper to convert Excel serial date floats to `time.Time`.
- Exported `Version` constant (`"0.2.0"`).
- Internal packages: `biff12`, `record`, `stringtable`, `workbook`, `worksheet`.
- Parsing and exposure of `MergeCells` (`[]MergeArea`) on `Worksheet`, including
  `MergeArea` type with `R`, `C`, `H`, `W` fields for the merged range.
- Integration test suite (`xlsb_integration_test.go`) covering real `.xlsb` fixtures.
- Developer probe tool (`_probe.go`, build-ignored).

### Removed

- `docs/official_microsoft_xlsb_format_documentation.pdf` removed from the
  repository; added to `.gitignore`.

[Unreleased]: https://github.com/TsubasaBE/go-xlsb/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/TsubasaBE/go-xlsb/compare/v1.0.2...v1.1.0
[1.0.2]: https://github.com/TsubasaBE/go-xlsb/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/TsubasaBE/go-xlsb/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/TsubasaBE/go-xlsb/compare/v0.2.0...v1.0.0
[0.2.0]: https://github.com/TsubasaBE/go-xlsb/releases/tag/v0.2.0
